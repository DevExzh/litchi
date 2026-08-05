//! Typed ODF paragraph border, padding, shadow, and background properties.
//!
//! Models the box attribute group of `style:paragraph-properties`
//! (`fo:border*`, `style:border-line-width*`, `fo:padding*`, `style:shadow`,
//! `fo:background-color`, `style:background-transparency`) plus the optional
//! `style:background-image` child. Attributes and child elements owned by
//! sibling paragraph modules (`style:tab-stops`, `style:drop-cap`) are ignored;
//! duplicates and malformed owned values are rejected.

use crate::{
    FlatDocument, Package,
    line_numbering::NonNegativeLength,
    style::paragraph::margin::rewrite_start_tag,
    style::table::row::{
        BackgroundColor, BackgroundImage, BackgroundPosition, BackgroundSource,
        HorizontalBackgroundPosition, Opacity, Repeat, VerticalBackgroundPosition,
    },
    style::table::table::Shadow,
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4_096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_BINARY: usize = 8 * 1024 * 1024;
const MAX_ATTRIBUTES: usize = 64;

/// Whether this module owns the attribute with the given expanded name.
fn owned_attribute(namespace: Ns, local: &[u8]) -> bool {
    match (namespace, local) {
        (Ns::Fo, local) => matches!(
            local,
            b"border"
                | b"border-top"
                | b"border-bottom"
                | b"border-left"
                | b"border-right"
                | b"padding"
                | b"padding-top"
                | b"padding-bottom"
                | b"padding-left"
                | b"padding-right"
                | b"background-color"
        ),
        (Ns::Style, local) => matches!(
            local,
            b"border-line-width"
                | b"border-line-width-top"
                | b"border-line-width-bottom"
                | b"border-line-width-left"
                | b"border-line-width-right"
                | b"shadow"
                | b"background-transparency"
        ),
        _ => false,
    }
}

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty())
        || value.len() > MAX_VALUE
        || value.chars().any(char::is_control)
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn name_ok(value: &str, field: &str) -> Result<()> {
    safe(value, field, false)
}
fn decimal(value: &str) -> bool {
    let mut parts = value.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some() {
        return false;
    }
    let digits = |part: &str| part.bytes().all(|byte| byte.is_ascii_digit());
    match fraction {
        None => !whole.is_empty() && digits(whole),
        Some(fraction) => {
            digits(whole) && digits(fraction) && (!whole.is_empty() || !fraction.is_empty())
        },
    }
}
fn positive_length(value: &str) -> bool {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return false;
    };
    decimal(number)
        && number
            .bytes()
            .any(|byte| byte.is_ascii_digit() && byte != b'0')
}

/// An `fo:border*` border description.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Border(String);
impl Border {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        safe(&value, "paragraph border", false)?;
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// One positive border line width component.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Width(String);
impl Width {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_VALUE || !positive_length(&value) {
            return Err(bad(
                "border line width must be a positive ODF physical length",
            ));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A three-part border line width: inner line, space between lines, outer line.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Widths {
    pub inner_width: Width,
    pub space: Width,
    pub outer_width: Width,
}
impl Widths {
    fn parse(value: &str) -> Result<Self> {
        let words: Vec<_> = value.split_ascii_whitespace().collect();
        let [inner, space, outer] = words.as_slice() else {
            return Err(bad("border line width needs exactly three lengths"));
        };
        Ok(Self {
            inner_width: Width::new(*inner)?,
            space: Width::new(*space)?,
            outer_width: Width::new(*outer)?,
        })
    }
    fn xml(&self) -> String {
        format!(
            "{} {} {}",
            self.inner_width.as_str(),
            self.space.as_str(),
            self.outer_width.as_str()
        )
    }
}

/// The `style:background-transparency` value: a percent between 0 and 100.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BackgroundTransparency(String);
impl BackgroundTransparency {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        let Some(number) = value.strip_suffix('%') else {
            return Err(bad("style:background-transparency must be a percentage"));
        };
        if number.starts_with(['+', '-']) || !decimal(number) {
            return Err(bad("invalid style:background-transparency"));
        }
        let parsed = number
            .parse::<f64>()
            .map_err(|_| bad("invalid style:background-transparency"))?;
        if !(0.0..=100.0).contains(&parsed) {
            return Err(bad("style:background-transparency is out of range"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// The border, padding, shadow, and background group of one
/// `style:paragraph-properties` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub border: Option<Border>,
    pub border_top: Option<Border>,
    pub border_bottom: Option<Border>,
    pub border_left: Option<Border>,
    pub border_right: Option<Border>,
    pub border_line_width: Option<Widths>,
    pub border_line_width_top: Option<Widths>,
    pub border_line_width_bottom: Option<Widths>,
    pub border_line_width_left: Option<Widths>,
    pub border_line_width_right: Option<Widths>,
    pub padding: Option<NonNegativeLength>,
    pub padding_top: Option<NonNegativeLength>,
    pub padding_bottom: Option<NonNegativeLength>,
    pub padding_left: Option<NonNegativeLength>,
    pub padding_right: Option<NonNegativeLength>,
    pub shadow: Option<Shadow>,
    pub background_color: Option<BackgroundColor>,
    pub background_transparency: Option<BackgroundTransparency>,
    pub background_image: Option<BackgroundImage>,
}
impl Properties {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn validate(&self) -> Result<()> {
        if let Some(image) = &self.background_image {
            image.validate()?;
        }
        Ok(())
    }
    /// Serialized owned attributes, each prefixed with one space.
    fn attributes_xml(&self) -> String {
        let mut xml = String::new();
        for (name, value) in [
            ("fo:border", &self.border),
            ("fo:border-top", &self.border_top),
            ("fo:border-bottom", &self.border_bottom),
            ("fo:border-left", &self.border_left),
            ("fo:border-right", &self.border_right),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(r#" {name}="{}""#, escape_xml(value.as_str())));
            }
        }
        for (name, value) in [
            ("style:border-line-width", &self.border_line_width),
            ("style:border-line-width-top", &self.border_line_width_top),
            (
                "style:border-line-width-bottom",
                &self.border_line_width_bottom,
            ),
            ("style:border-line-width-left", &self.border_line_width_left),
            (
                "style:border-line-width-right",
                &self.border_line_width_right,
            ),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(r#" {name}="{}""#, value.xml()));
            }
        }
        for (name, value) in [
            ("fo:padding", &self.padding),
            ("fo:padding-top", &self.padding_top),
            ("fo:padding-bottom", &self.padding_bottom),
            ("fo:padding-left", &self.padding_left),
            ("fo:padding-right", &self.padding_right),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(r#" {name}="{}""#, escape_xml(value.as_str())));
            }
        }
        if let Some(value) = &self.shadow {
            xml.push_str(&format!(
                r#" style:shadow="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = &self.background_color {
            xml.push_str(&format!(r#" fo:background-color="{}""#, value.as_str()));
        }
        if let Some(value) = &self.background_transparency {
            xml.push_str(&format!(
                r#" style:background-transparency="{}""#,
                value.as_str()
            ));
        }
        xml
    }
    /// Emit the properties as a `style:paragraph-properties` fragment.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:paragraph-properties xmlns:style="{STYLE_NS}" xmlns:office="{OFFICE_NS}" xmlns:fo="{FO_NS}" xmlns:draw="{DRAW_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        xml.push_str(&self.attributes_xml());
        if let Some(image) = &self.background_image {
            xml.push('>');
            xml.push_str(&image.to_xml_fragment()?);
            xml.push_str("</style:paragraph-properties>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

/// A named or default paragraph style and its border and background properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<Properties>,
}
impl Style {
    pub fn named(name: impl Into<String>, properties: Option<Properties>) -> Result<Self> {
        let result = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        result.validate()?;
        Ok(result)
    }
    pub fn default_style(properties: Option<Properties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(name), false) => name_ok(name, "paragraph style name")?,
            (None, true) => {},
            _ => return Err(bad("paragraph style identity is inconsistent")),
        }
        if let Some(parent) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default paragraph style cannot have a parent"));
            }
            name_ok(parent, "parent paragraph style name")?;
        }
        if let Some(properties) = &self.properties {
            properties.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="paragraph""#);
        if let Some(name) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(name)));
        }
        if let Some(parent) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(parent)
            ));
        }
        if let Some(properties) = &self.properties {
            xml.push('>');
            xml.push_str(&properties.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"));
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

/// All paragraph styles of a styles part that carry border or background properties.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}
impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&Style> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Fo,
    Draw,
    Xlink,
    Other,
}
fn known(resolve: ResolveResult<'_>) -> Ns {
    match resolve {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == FO => Ns::Fo,
        ResolveResult::Bound(value) if value.as_ref() == DRAW => Ns::Draw,
        ResolveResult::Bound(value) if value.as_ref() == XLINK => Ns::Xlink,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (known(namespace), local.as_ref().to_vec())
}
fn attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Ns, Vec<u8>, String)>> {
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid paragraph attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if result.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many paragraph-properties attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (known(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate paragraph-properties attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid paragraph value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("paragraph-properties attribute is too large"));
        }
        result.push((key.0, key.1, value));
    }
    Ok(result)
}
fn take(attrs: &mut Vec<(Ns, Vec<u8>, String)>, namespace: Ns, local: &[u8]) -> Option<String> {
    attrs
        .iter()
        .position(|x| x.0 == namespace && x.1 == local)
        .map(|at| attrs.remove(at).2)
}

fn style_header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<Style>> {
    let mut attrs = attributes(reader, version, start)?;
    if take(&mut attrs, Ns::Style, b"family").as_deref() != Some("paragraph") {
        return Ok(None);
    }
    let value = Style {
        name: take(&mut attrs, Ns::Style, b"name"),
        parent_style_name: take(&mut attrs, Ns::Style, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    value.validate()?;
    Ok(Some(value))
}

fn border_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Properties> {
    let mut attrs = attributes(reader, version, start)?;
    let border = |attrs: &mut Vec<(Ns, Vec<u8>, String)>, local: &[u8]| {
        take(attrs, Ns::Fo, local).map(Border::new).transpose()
    };
    let widths = |attrs: &mut Vec<(Ns, Vec<u8>, String)>, local: &[u8]| {
        take(attrs, Ns::Style, local)
            .map(|x| Widths::parse(&x))
            .transpose()
    };
    let padding = |attrs: &mut Vec<(Ns, Vec<u8>, String)>, local: &[u8]| {
        take(attrs, Ns::Fo, local)
            .map(NonNegativeLength::new)
            .transpose()
    };
    let value = Properties {
        border: border(&mut attrs, b"border")?,
        border_top: border(&mut attrs, b"border-top")?,
        border_bottom: border(&mut attrs, b"border-bottom")?,
        border_left: border(&mut attrs, b"border-left")?,
        border_right: border(&mut attrs, b"border-right")?,
        border_line_width: widths(&mut attrs, b"border-line-width")?,
        border_line_width_top: widths(&mut attrs, b"border-line-width-top")?,
        border_line_width_bottom: widths(&mut attrs, b"border-line-width-bottom")?,
        border_line_width_left: widths(&mut attrs, b"border-line-width-left")?,
        border_line_width_right: widths(&mut attrs, b"border-line-width-right")?,
        padding: padding(&mut attrs, b"padding")?,
        padding_top: padding(&mut attrs, b"padding-top")?,
        padding_bottom: padding(&mut attrs, b"padding-bottom")?,
        padding_left: padding(&mut attrs, b"padding-left")?,
        padding_right: padding(&mut attrs, b"padding-right")?,
        shadow: take(&mut attrs, Ns::Style, b"shadow")
            .map(Shadow::new)
            .transpose()?,
        background_color: take(&mut attrs, Ns::Fo, b"background-color")
            .map(BackgroundColor::new)
            .transpose()?,
        background_transparency: take(&mut attrs, Ns::Style, b"background-transparency")
            .map(BackgroundTransparency::new)
            .transpose()?,
        background_image: None,
    };
    // Remaining attributes are owned by sibling paragraph modules.
    value.validate()?;
    Ok(value)
}

fn position(value: &str) -> Result<BackgroundPosition> {
    let words: Vec<_> = value.split_ascii_whitespace().collect();
    let horizontal = |word| match word {
        "left" => Some(HorizontalBackgroundPosition::Left),
        "center" => Some(HorizontalBackgroundPosition::Center),
        "right" => Some(HorizontalBackgroundPosition::Right),
        _ => None,
    };
    let vertical = |word| match word {
        "top" => Some(VerticalBackgroundPosition::Top),
        "center" => Some(VerticalBackgroundPosition::Center),
        "bottom" => Some(VerticalBackgroundPosition::Bottom),
        _ => None,
    };
    match words.as_slice() {
        ["left"] => Ok(BackgroundPosition::Left),
        ["center"] => Ok(BackgroundPosition::Center),
        ["right"] => Ok(BackgroundPosition::Right),
        ["top"] => Ok(BackgroundPosition::Top),
        ["bottom"] => Ok(BackgroundPosition::Bottom),
        [a, b] => horizontal(a)
            .zip(vertical(b))
            .or_else(|| horizontal(b).zip(vertical(a)))
            .map(|(h, v)| BackgroundPosition::Pair(h, v))
            .ok_or_else(|| bad("invalid background position")),
        _ => Err(bad("invalid background position")),
    }
}

struct ParsedImage {
    image: BackgroundImage,
    linked: bool,
}
fn image_attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ParsedImage> {
    let mut attrs = attributes(reader, version, start)?;
    let repeat = take(&mut attrs, Ns::Style, b"repeat")
        .map(|x| match x.as_str() {
            "no-repeat" => Ok(Repeat::NoRepeat),
            "repeat" => Ok(Repeat::Repeat),
            "stretch" => Ok(Repeat::Stretch),
            _ => Err(bad("invalid background repeat")),
        })
        .transpose()?;
    let pos = take(&mut attrs, Ns::Style, b"position")
        .map(|x| position(&x))
        .transpose()?;
    let filter_name = take(&mut attrs, Ns::Style, b"filter-name");
    if let Some(x) = &filter_name {
        safe(x, "style:filter-name", true)?;
    }
    let opacity = take(&mut attrs, Ns::Draw, b"opacity")
        .map(Opacity::new)
        .transpose()?;
    let kind = take(&mut attrs, Ns::Xlink, b"type");
    let href = take(&mut attrs, Ns::Xlink, b"href");
    let show = take(&mut attrs, Ns::Xlink, b"show");
    let actuate = take(&mut attrs, Ns::Xlink, b"actuate");
    if !attrs.is_empty() {
        return Err(bad("unknown style:background-image attribute"));
    }
    let linked = kind.is_some() || href.is_some() || show.is_some() || actuate.is_some();
    let source = if linked {
        if kind.as_deref() != Some("simple")
            || href.is_none()
            || show.as_deref().is_some_and(|x| x != "embed")
            || actuate.as_deref().is_some_and(|x| x != "onLoad")
        {
            return Err(bad("invalid background-image xlink group"));
        }
        BackgroundSource::Link {
            href: href.unwrap(),
            show_embed: show.is_some(),
            actuate_on_load: actuate.is_some(),
        }
    } else {
        BackgroundSource::Empty
    };
    let image = BackgroundImage {
        repeat,
        position: pos,
        filter_name,
        opacity,
        source,
    };
    image.validate()?;
    Ok(ParsedImage { image, linked })
}

fn push_style(styles: &mut Vec<Style>, style: Style, total: &mut usize) -> Result<()> {
    if styles.len() >= MAX_STYLES {
        return Err(bad("too many paragraph styles"));
    }
    if styles
        .iter()
        .any(|old| old.name == style.name && old.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate paragraph style identity"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("paragraph border data is too large"));
    }
    styles.push(style);
    Ok(())
}

fn is_paragraph_style(current: &(Ns, Vec<u8>), parent: Option<&(Ns, Vec<u8>)>) -> bool {
    parent.is_some_and(|(namespace, local)| {
        *namespace == Ns::Office && matches!(local.as_slice(), b"styles" | b"automatic-styles")
    }) && current.0 == Ns::Style
        && matches!(current.1.as_slice(), b"style" | b"default-style")
}

struct Active {
    depth: usize,
    style: Style,
    seen_properties: bool,
    properties_depth: Option<usize>,
    image_depth: Option<usize>,
    binary_depth: Option<usize>,
    binary: String,
    image_linked: bool,
}
impl Active {
    fn properties_mut(&mut self) -> &mut Properties {
        self.style
            .properties
            .as_mut()
            .expect("properties depth implies properties")
    }
}

/// Parse paragraph styles and their border and background properties from a
/// styles part.
pub fn parse(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    if !xml.contains("paragraph-properties") {
        return Ok(Styles::default());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut styles = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        active = Some(Active {
                            depth,
                            style,
                            seen_properties: false,
                            properties_depth: None,
                            image_depth: None,
                            binary_depth: None,
                            binary: String::new(),
                            image_linked: false,
                        });
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"paragraph-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(border_properties(&reader, version, &start)?);
                        state.properties_depth = Some(depth);
                    } else if state.properties_depth.is_some()
                        && depth == state.properties_depth.unwrap() + 1
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        if state.image_depth.is_some()
                            || state.properties_mut().background_image.is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        let parsed = image_attributes(&reader, version, &start)?;
                        state.properties_mut().background_image = Some(parsed.image);
                        state.image_linked = parsed.linked;
                        state.image_depth = Some(depth);
                    } else if state.image_depth.is_some()
                        && depth == state.image_depth.unwrap() + 1
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        if state.image_linked || state.binary_depth.is_some() {
                            return Err(bad("invalid office:binary-data in background image"));
                        }
                        state.binary_depth = Some(depth);
                        state.binary.clear();
                    }
                    // Other children belong to sibling paragraph modules.
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                let depth = stack.len() + 1;
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push_style(&mut styles, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"paragraph-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(border_properties(&reader, version, &start)?);
                    } else if state.properties_depth.is_some()
                        && depth == state.properties_depth.unwrap() + 1
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        if state.properties_mut().background_image.is_some() {
                            return Err(bad("duplicate style:background-image"));
                        }
                        state.properties_mut().background_image =
                            Some(image_attributes(&reader, version, &start)?.image);
                    } else if state.image_depth.is_some()
                        && depth == state.image_depth.unwrap() + 1
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        if state.image_linked {
                            return Err(bad("linked background image cannot contain binary data"));
                        }
                        state
                            .properties_mut()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = BackgroundSource::Embedded(Vec::new());
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut()
                    && state.binary_depth.is_some()
                {
                    if state.binary.len() + bytes.len() > MAX_BINARY * 2 {
                        return Err(bad("encoded office:binary-data is too large"));
                    }
                    state.binary.push_str(&String::from_utf8_lossy(bytes));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut()
                    && state.binary_depth.is_some()
                {
                    if state.binary.len() + bytes.len() > MAX_BINARY * 2 {
                        return Err(bad("encoded office:binary-data is too large"));
                    }
                    state.binary.push_str(&String::from_utf8_lossy(bytes));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth == Some(depth) {
                        let data = base64_decode(&state.binary)?;
                        state
                            .properties_mut()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = BackgroundSource::Embedded(data);
                        state.binary_depth = None;
                    }
                    if state.image_depth == Some(depth) {
                        state.image_depth = None;
                    }
                    if state.properties_depth == Some(depth) {
                        state.properties_depth = None;
                    }
                }
                if active.as_ref().is_some_and(|state| state.depth == depth) {
                    push_style(&mut styles, active.take().unwrap().style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(Styles { styles })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
    owned: Vec<String>,
    missing_ns: (bool, bool),
}
#[derive(Default)]
struct TargetSpans {
    style: Span,
    properties: Option<Span>,
    image: Option<Span>,
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
fn replace_span(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand_span(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| bad("invalid empty element"))?;
    Ok(replace_span(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}

/// Qualified names of the owned attributes as they literally appear on one
/// start tag, so lossless mutation works under arbitrary prefix aliases.
fn owned_qnames(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Vec<String>> {
    let mut owned = Vec::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid paragraph attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if owned_attribute(known(namespace), local.as_ref()) {
            owned.push(String::from_utf8_lossy(attribute.key.as_ref()).into_owned());
        }
    }
    Ok(owned)
}

/// Whether the `fo:` and `style:` prefixes are not bound to their ODF
/// namespaces in the reader's current scope, so inserted attributes need
/// local namespace declarations.
fn missing_ns_decls(reader: &NsReader<&[u8]>) -> (bool, bool) {
    let unbound = |probe: &[u8], uri: &[u8]| {
        !matches!(
            reader.resolver().resolve_attribute(QName(probe)),
            (ResolveResult::Bound(namespace), _) if namespace.as_ref() == uri
        )
    };
    (unbound(b"fo:x", FO), unbound(b"style:x", STYLE))
}

/// Prepend local namespace declarations to the serialized attributes for any
/// prefix that is unbound in the target scope.
fn qualify_insert(insert: &str, missing: (bool, bool)) -> String {
    let mut qualified = String::new();
    if missing.0 && insert.contains(" fo:") {
        qualified.push_str(&format!(r#" xmlns:fo="{FO_NS}""#));
    }
    if missing.1 && insert.contains(" style:") {
        qualified.push_str(&format!(r#" xmlns:style="{STYLE_NS}""#));
    }
    qualified.push_str(insert);
    qualified
}

/// Losslessly replace, insert, or remove this module's border and background
/// attributes (and the `style:background-image` child) on one existing
/// paragraph style's `style:paragraph-properties` element. Attributes owned by
/// sibling modules and other child elements are preserved.
pub fn set_xml(xml: &str, requested: &Style) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut depth_target = None;
    let mut active: Option<TargetSpans> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target paragraph style"));
                        }
                        depth_target = Some(depth);
                        active = Some(TargetSpans {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        });
                    }
                } else if let Some(target) = depth_target {
                    if depth == target + 1
                        && current.0 == Ns::Style
                        && current.1 == b"paragraph-properties"
                    {
                        let span = Span {
                            start: begin,
                            end,
                            qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                            owned: owned_qnames(&reader, &start)?,
                            missing_ns: missing_ns_decls(&reader),
                            ..Default::default()
                        };
                        if active.as_mut().unwrap().properties.replace(span).is_some() {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                    } else if depth == target + 2
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        let span = Span {
                            start: begin,
                            qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                            ..Default::default()
                        };
                        if active.as_mut().unwrap().image.replace(span).is_some() {
                            return Err(bad("duplicate style:background-image"));
                        }
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                let depth = stack.len() + 1;
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                    ..Default::default()
                };
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target paragraph style"));
                        }
                        found = Some(TargetSpans {
                            style: span,
                            ..Default::default()
                        });
                    }
                } else if let Some(target) = depth_target {
                    if depth == target + 1
                        && current.0 == Ns::Style
                        && current.1 == b"paragraph-properties"
                    {
                        let span = Span {
                            owned: owned_qnames(&reader, &start)?,
                            missing_ns: missing_ns_decls(&reader),
                            ..span
                        };
                        if active.as_mut().unwrap().properties.replace(span).is_some() {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                    } else if depth == target + 2
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                        && active.as_mut().unwrap().image.replace(span).is_some()
                    {
                        return Err(bad("duplicate style:background-image"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.image.as_ref().is_some_and(|span| span.end == 0)
                        && depth_target.is_some_and(|target| depth == target + 2)
                    {
                        let span = spans.image.as_mut().unwrap();
                        span.end_start = begin;
                        span.end = end;
                    }
                    if spans
                        .properties
                        .as_ref()
                        .is_some_and(|span| span.end_start == 0)
                        && depth_target.is_some_and(|target| depth == target + 1)
                    {
                        spans.properties.as_mut().unwrap().end_start = begin;
                    }
                    if depth_target == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        depth_target = None;
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target paragraph style does not exist"))?;
    let insert = requested
        .properties
        .as_ref()
        .map(Properties::attributes_xml)
        .unwrap_or_default();
    let image = requested
        .properties
        .as_ref()
        .and_then(|properties| properties.background_image.as_ref())
        .map(BackgroundImage::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        // Edit the deeper spans first so earlier offsets stay valid.
        let mut out = xml.to_owned();
        if let Some(existing) = &spans.image {
            out = replace_span(&out, existing, image.as_deref().unwrap_or(""));
        } else if let Some(image) = &image {
            // The empty-element case is handled by the start-tag rewrite below.
            if !properties.empty {
                out.insert_str(properties.end_start, image);
            }
        }
        let raw = &xml[properties.start..properties.end];
        let insert = qualify_insert(&insert, properties.missing_ns);
        let rewritten = rewrite_start_tag(raw, &properties.owned, &insert)?;
        let rewritten = if properties.empty && image.is_some() && spans.image.is_none() {
            let slash = rewritten
                .rfind("/>")
                .ok_or_else(|| bad("invalid empty element"))?;
            format!(
                "{}>{}</{}>",
                &rewritten[..slash],
                image.as_deref().unwrap_or_default(),
                properties.qname
            )
        } else {
            rewritten
        };
        out.replace_range(properties.start..properties.end, &rewritten);
        return Ok(out);
    }
    if requested.properties.is_none() {
        return Ok(xml.to_owned());
    }
    let fragment = requested.properties.as_ref().unwrap().to_xml_fragment()?;
    if spans.style.empty {
        return expand_span(xml, &spans.style, &fragment);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &fragment);
    Ok(out)
}

fn base64_decode(value: &str) -> Result<Vec<u8>> {
    let clean: Vec<u8> = value.bytes().filter(|x| !x.is_ascii_whitespace()).collect();
    if !clean.len().is_multiple_of(4) {
        return Err(bad("invalid office:binary-data base64"));
    }
    let val = |x| match x {
        b'A'..=b'Z' => Some(x - b'A'),
        b'a'..=b'z' => Some(x - b'a' + 26),
        b'0'..=b'9' => Some(x - b'0' + 52),
        b'+' => Some(62),
        b'/' => Some(63),
        _ => None,
    };
    let mut out = Vec::with_capacity(clean.len() / 4 * 3);
    for (index, chunk) in clean.chunks(4).enumerate() {
        let last = index + 1 == clean.len() / 4;
        let pad = (chunk[2] == b'=') as usize + (chunk[3] == b'=') as usize;
        if !last && pad != 0 || chunk[2] == b'=' && chunk[3] != b'=' || pad > 2 {
            return Err(bad("invalid office:binary-data padding"));
        }
        let a = val(chunk[0]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32;
        let b = val(chunk[1]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32;
        let c = if chunk[2] == b'=' {
            0
        } else {
            val(chunk[2]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32
        };
        let d = if chunk[3] == b'=' {
            0
        } else {
            val(chunk[3]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32
        };
        let n = a << 18 | b << 12 | c << 6 | d;
        out.push((n >> 16) as u8);
        if pad < 2 {
            out.push((n >> 8) as u8)
        }
        if pad < 1 {
            out.push(n as u8)
        }
        if out.len() > MAX_BINARY {
            return Err(bad("office:binary-data is too large"));
        }
    }
    Ok(out)
}

impl Package {
    pub fn paragraph_style_borders(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Styles::default()), |xml| parse(&xml))
    }
}
impl FlatDocument {
    pub fn paragraph_style_borders(&self) -> Result<Styles> {
        parse(self.xml())
    }
}
