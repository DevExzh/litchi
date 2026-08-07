//! Typed ODF table-row style properties.

use crate::{FlatDocument, Package};
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
const MAX_ATTRIBUTES: usize = 32;

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn safe(value: &str, field: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty())
        || value.len() > MAX_VALUE
        || value.chars().any(char::is_control)
    {
        return Err(bad(format!("invalid {field}")));
    }
    Ok(())
}

/// A positive or non-negative ODF physical length.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Length(String);
impl Length {
    pub fn positive(value: impl Into<String>) -> Result<Self> {
        Self::new(value.into(), false)
    }
    pub fn non_negative(value: impl Into<String>) -> Result<Self> {
        Self::new(value.into(), true)
    }
    fn new(value: String, zero: bool) -> Result<Self> {
        if value.len() > MAX_VALUE {
            return Err(bad("table-row length is too large"));
        }
        let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
            .iter()
            .find_map(|unit| value.strip_suffix(unit))
        else {
            return Err(bad("table-row length must use an ODF physical unit"));
        };
        if number.starts_with(['+', '-']) {
            return Err(bad("table-row length cannot be signed"));
        }
        let mut parts = number.split('.');
        let whole = parts.next().unwrap_or_default();
        let fraction = parts.next();
        if parts.next().is_some()
            || !whole.bytes().all(|c| c.is_ascii_digit())
            || fraction
                .is_some_and(|part| part.is_empty() || !part.bytes().all(|c| c.is_ascii_digit()))
            || whole.is_empty()
        {
            return Err(bad("invalid table-row length"));
        }
        let nonzero = number.bytes().any(|c| c.is_ascii_digit() && c != b'0');
        if !zero && !nonzero {
            return Err(bad("style:row-height must be positive"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Break {
    Auto,
    Column,
    Page,
}
impl Break {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "auto" => Ok(Self::Auto),
            "column" => Ok(Self::Column),
            "page" => Ok(Self::Page),
            _ => Err(bad("invalid table-row break value")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Column => "column",
            Self::Page => "page",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum KeepTogether {
    Auto,
    Always,
}
impl KeepTogether {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "auto" => Ok(Self::Auto),
            "always" => Ok(Self::Always),
            _ => Err(bad("invalid fo:keep-together")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Always => "always",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BackgroundColor(String);
impl BackgroundColor {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        let color = value == "transparent"
            || value.len() == 7
                && value.starts_with('#')
                && value[1..].bytes().all(|c| c.is_ascii_hexdigit());
        if !color {
            return Err(bad("fo:background-color must be transparent or #RRGGBB"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Repeat {
    NoRepeat,
    Repeat,
    Stretch,
}
impl Repeat {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "no-repeat" => Ok(Self::NoRepeat),
            "repeat" => Ok(Self::Repeat),
            "stretch" => Ok(Self::Stretch),
            _ => Err(bad("invalid style:repeat")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::NoRepeat => "no-repeat",
            Self::Repeat => "repeat",
            Self::Stretch => "stretch",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalBackgroundPosition {
    Left,
    Center,
    Right,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalBackgroundPosition {
    Top,
    Center,
    Bottom,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BackgroundPosition {
    Left,
    Center,
    Right,
    Top,
    Bottom,
    Pair(HorizontalBackgroundPosition, VerticalBackgroundPosition),
}
impl BackgroundPosition {
    fn horizontal(value: &str) -> Option<HorizontalBackgroundPosition> {
        match value {
            "left" => Some(HorizontalBackgroundPosition::Left),
            "center" => Some(HorizontalBackgroundPosition::Center),
            "right" => Some(HorizontalBackgroundPosition::Right),
            _ => None,
        }
    }
    fn vertical(value: &str) -> Option<VerticalBackgroundPosition> {
        match value {
            "top" => Some(VerticalBackgroundPosition::Top),
            "center" => Some(VerticalBackgroundPosition::Center),
            "bottom" => Some(VerticalBackgroundPosition::Bottom),
            _ => None,
        }
    }
    fn parse(value: &str) -> Result<Self> {
        let words: Vec<_> = value.split_ascii_whitespace().collect();
        match words.as_slice() {
            ["left"] => Ok(Self::Left),
            ["center"] => Ok(Self::Center),
            ["right"] => Ok(Self::Right),
            ["top"] => Ok(Self::Top),
            ["bottom"] => Ok(Self::Bottom),
            [a, b] => Self::horizontal(a)
                .zip(Self::vertical(b))
                .or_else(|| Self::horizontal(b).zip(Self::vertical(a)))
                .map(|(h, v)| Self::Pair(h, v))
                .ok_or_else(|| bad("invalid style:position")),
            _ => Err(bad("invalid style:position")),
        }
    }
    fn xml(self) -> String {
        let h = |value| match value {
            HorizontalBackgroundPosition::Left => "left",
            HorizontalBackgroundPosition::Center => "center",
            HorizontalBackgroundPosition::Right => "right",
        };
        let v = |value| match value {
            VerticalBackgroundPosition::Top => "top",
            VerticalBackgroundPosition::Center => "center",
            VerticalBackgroundPosition::Bottom => "bottom",
        };
        match self {
            Self::Left => "left".into(),
            Self::Center => "center".into(),
            Self::Right => "right".into(),
            Self::Top => "top".into(),
            Self::Bottom => "bottom".into(),
            Self::Pair(x, y) => format!("{} {}", h(x), v(y)),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opacity(String);
impl Opacity {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        let Some(number) = value.strip_suffix('%') else {
            return Err(bad("draw:opacity must be a percentage"));
        };
        if number.starts_with(['+', '-']) || number.is_empty() {
            return Err(bad("invalid draw:opacity"));
        }
        let parsed = number
            .parse::<f64>()
            .map_err(|_| bad("invalid draw:opacity"))?;
        if !parsed.is_finite() || !(0.0..=100.0).contains(&parsed) {
            return Err(bad("draw:opacity is out of range"));
        }
        if number.contains(['e', 'E']) || number.split('.').count() > 2 {
            return Err(bad("invalid draw:opacity"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BackgroundSource {
    Empty,
    Link {
        href: String,
        show_embed: bool,
        actuate_on_load: bool,
    },
    Embedded(Vec<u8>),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BackgroundImage {
    pub repeat: Option<Repeat>,
    pub position: Option<BackgroundPosition>,
    pub filter_name: Option<String>,
    pub opacity: Option<Opacity>,
    pub source: BackgroundSource,
}
impl Default for BackgroundImage {
    fn default() -> Self {
        Self {
            repeat: None,
            position: None,
            filter_name: None,
            opacity: None,
            source: BackgroundSource::Empty,
        }
    }
}
impl BackgroundImage {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = &self.filter_name {
            safe(value, "style:filter-name", true)?;
        }
        match &self.source {
            BackgroundSource::Empty => {},
            BackgroundSource::Link { href, .. } => safe(href, "xlink:href", true)?,
            BackgroundSource::Embedded(data) if data.len() <= MAX_BINARY => {},
            BackgroundSource::Embedded(_) => {
                return Err(bad("office:binary-data is too large"));
            },
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:background-image xmlns:style="{STYLE_NS}" xmlns:office="{OFFICE_NS}" xmlns:draw="{DRAW_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        if let Some(value) = self.repeat {
            xml.push_str(&format!(r#" style:repeat="{}""#, value.xml()));
        }
        if let Some(value) = self.position {
            xml.push_str(&format!(r#" style:position="{}""#, value.xml()));
        }
        if let Some(value) = &self.filter_name {
            xml.push_str(&format!(r#" style:filter-name="{}""#, escape_xml(value)));
        }
        if let Some(value) = &self.opacity {
            xml.push_str(&format!(r#" draw:opacity="{}""#, value.as_str()));
        }
        match &self.source {
            BackgroundSource::Empty => xml.push_str("/>"),
            BackgroundSource::Link {
                href,
                show_embed,
                actuate_on_load,
            } => {
                xml.push_str(&format!(
                    r#" xlink:type="simple" xlink:href="{}""#,
                    escape_xml(href)
                ));
                if *show_embed {
                    xml.push_str(r#" xlink:show="embed""#);
                }
                if *actuate_on_load {
                    xml.push_str(r#" xlink:actuate="onLoad""#);
                }
                xml.push_str("/>");
            },
            BackgroundSource::Embedded(data) => {
                xml.push_str("><office:binary-data>");
                xml.push_str(&base64_encode(data));
                xml.push_str("</office:binary-data></style:background-image>");
            },
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub row_height: Option<Length>,
    pub min_row_height: Option<Length>,
    pub use_optimal_row_height: Option<bool>,
    pub background_color: Option<BackgroundColor>,
    pub break_before: Option<Break>,
    pub break_after: Option<Break>,
    pub keep_together: Option<KeepTogether>,
    pub background_image: Option<BackgroundImage>,
}
impl Properties {
    pub fn validate(&self) -> Result<()> {
        if let Some(image) = &self.background_image {
            image.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:table-row-properties xmlns:style="{STYLE_NS}" xmlns:office="{OFFICE_NS}" xmlns:fo="{FO_NS}" xmlns:draw="{DRAW_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        if let Some(value) = &self.row_height {
            xml.push_str(&format!(r#" style:row-height="{}""#, value.as_str()));
        }
        if let Some(value) = &self.min_row_height {
            xml.push_str(&format!(r#" style:min-row-height="{}""#, value.as_str()));
        }
        if let Some(value) = self.use_optimal_row_height {
            xml.push_str(&format!(r#" style:use-optimal-row-height="{value}""#));
        }
        if let Some(value) = &self.background_color {
            xml.push_str(&format!(r#" fo:background-color="{}""#, value.as_str()));
        }
        if let Some(value) = self.break_before {
            xml.push_str(&format!(r#" fo:break-before="{}""#, value.xml()));
        }
        if let Some(value) = self.break_after {
            xml.push_str(&format!(r#" fo:break-after="{}""#, value.xml()));
        }
        if let Some(value) = self.keep_together {
            xml.push_str(&format!(r#" fo:keep-together="{}""#, value.xml()));
        }
        if let Some(image) = &self.background_image {
            xml.push('>');
            xml.push_str(&image.to_xml_fragment()?);
            xml.push_str("</style:table-row-properties>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<Properties>,
}
impl Style {
    pub fn named(name: impl Into<String>, properties: Option<Properties>) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
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
            (Some(value), false) => safe(value, "table-row style name", false)?,
            (None, true) => {},
            _ => return Err(bad("invalid table-row style identity")),
        }
        if let Some(value) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default table-row style cannot have a parent"));
            }
            safe(value, "parent table-row style name", false)?;
        }
        if let Some(value) = &self.properties {
            value.validate()?;
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
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="table-row""#);
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)));
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ));
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"));
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

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

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Fo,
    Draw,
    Xlink,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(x) if x.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(x) if x.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(x) if x.as_ref() == FO => Ns::Fo,
        ResolveResult::Bound(x) if x.as_ref() == DRAW => Ns::Draw,
        ResolveResult::Bound(x) if x.as_ref() == XLINK => Ns::Xlink,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ns(namespace), local.as_ref().to_vec())
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
            attribute.map_err(|error| bad(format!("invalid table-row attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if result.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many table-row attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate table-row attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid table-row value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("table-row attribute value is too large"));
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
fn bool_value(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(bad("invalid style:use-optimal-row-height")),
    }
}
fn style_header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<Style>> {
    let mut attrs = attributes(reader, version, start)?;
    if take(&mut attrs, Ns::Style, b"family").as_deref() != Some("table-row") {
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
fn row_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Properties> {
    let mut attrs = attributes(reader, version, start)?;
    let value = Properties {
        row_height: take(&mut attrs, Ns::Style, b"row-height")
            .map(Length::positive)
            .transpose()?,
        min_row_height: take(&mut attrs, Ns::Style, b"min-row-height")
            .map(Length::non_negative)
            .transpose()?,
        use_optimal_row_height: take(&mut attrs, Ns::Style, b"use-optimal-row-height")
            .map(|x| bool_value(&x))
            .transpose()?,
        background_color: take(&mut attrs, Ns::Fo, b"background-color")
            .map(BackgroundColor::new)
            .transpose()?,
        break_before: take(&mut attrs, Ns::Fo, b"break-before")
            .map(|x| Break::parse(&x))
            .transpose()?,
        break_after: take(&mut attrs, Ns::Fo, b"break-after")
            .map(|x| Break::parse(&x))
            .transpose()?,
        keep_together: take(&mut attrs, Ns::Fo, b"keep-together")
            .map(|x| KeepTogether::parse(&x))
            .transpose()?,
        background_image: None,
    };
    if !attrs.is_empty() {
        return Err(bad("unknown style:table-row-properties attribute"));
    }
    Ok(value)
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
        .map(|x| Repeat::parse(&x))
        .transpose()?;
    let position = take(&mut attrs, Ns::Style, b"position")
        .map(|x| BackgroundPosition::parse(&x))
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
        position,
        filter_name,
        opacity,
        source,
    };
    image.validate()?;
    Ok(ParsedImage { image, linked })
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
fn push_style(out: &mut Vec<Style>, style: Style, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out
            .iter()
            .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate or excessive table-row style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("table-row style data is too large"));
    }
    out.push(style);
    Ok(())
}

/// Parse direct table-row styles in `office:styles` and `office:automatic-styles`.
pub fn parse(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
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
                        && current.1 == b"table-row-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:table-row-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(row_properties(&reader, version, &start)?);
                        state.properties_depth = Some(depth);
                    } else if current.1 == b"table-row-properties" {
                        return Err(bad(
                            "style:table-row-properties has invalid namespace or parent",
                        ));
                    } else if state.properties_depth.is_some()
                        && depth == state.properties_depth.unwrap() + 1
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        if state.image_depth.is_some()
                            || state
                                .style
                                .properties
                                .as_ref()
                                .unwrap()
                                .background_image
                                .is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        let parsed = image_attributes(&reader, version, &start)?;
                        state.style.properties.as_mut().unwrap().background_image =
                            Some(parsed.image);
                        state.image_linked = parsed.linked;
                        state.image_depth = Some(depth);
                    } else if current.1 == b"background-image" {
                        return Err(bad(
                            "style:background-image has invalid namespace or parent",
                        ));
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
                    } else if state.properties_depth.is_some()
                        && depth > state.properties_depth.unwrap()
                    {
                        return Err(bad("unexpected table-row property child"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push_style(&mut out, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"table-row-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:table-row-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(row_properties(&reader, version, &start)?);
                    } else if current.1 == b"table-row-properties" {
                        return Err(bad(
                            "style:table-row-properties has invalid namespace or parent",
                        ));
                    } else if state.properties_depth.is_some()
                        && depth == state.properties_depth.unwrap() + 1
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        if state
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .background_image
                            .is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        state.style.properties.as_mut().unwrap().background_image =
                            Some(image_attributes(&reader, version, &start)?.image);
                    } else if current.1 == b"background-image" {
                        return Err(bad(
                            "style:background-image has invalid namespace or parent",
                        ));
                    } else if state.image_depth.is_some()
                        && depth == state.image_depth.unwrap() + 1
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        if state.image_linked {
                            return Err(bad("linked background image cannot contain binary data"));
                        }
                        state
                            .style
                            .properties
                            .as_mut()
                            .unwrap()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = BackgroundSource::Embedded(Vec::new());
                    } else if state.properties_depth.is_some()
                        && depth > state.properties_depth.unwrap()
                    {
                        return Err(bad("unexpected table-row property child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth.is_some() {
                        if state.binary.len() + bytes.len() > MAX_BINARY * 2 {
                            return Err(bad("encoded office:binary-data is too large"));
                        }
                        state.binary.push_str(&String::from_utf8_lossy(bytes));
                    } else if state.properties_depth.is_some()
                        && !bytes.iter().all(u8::is_ascii_whitespace)
                    {
                        return Err(bad("unexpected text in table-row properties"));
                    }
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth.is_some() {
                        if state.binary.len() + bytes.len() > MAX_BINARY * 2 {
                            return Err(bad("encoded office:binary-data is too large"));
                        }
                        state.binary.push_str(&String::from_utf8_lossy(bytes));
                    } else if state.properties_depth.is_some()
                        && !bytes.iter().all(u8::is_ascii_whitespace)
                    {
                        return Err(bad("unexpected text in table-row properties"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth == Some(depth) {
                        let data = base64_decode(&state.binary)?;
                        state
                            .style
                            .properties
                            .as_mut()
                            .unwrap()
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
                if active.as_ref().is_some_and(|x| x.depth == depth) {
                    push_style(&mut out, active.take().unwrap().style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
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
    Ok(Styles { styles: out })
}

fn base64_encode(data: &[u8]) -> String {
    const C: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(data.len().div_ceil(3) * 4);
    for chunk in data.chunks(3) {
        let n = u32::from(chunk[0]) << 16
            | u32::from(chunk.get(1).copied().unwrap_or(0)) << 8
            | u32::from(chunk.get(2).copied().unwrap_or(0));
        out.push(C[((n >> 18) & 63) as usize] as char);
        out.push(C[((n >> 12) & 63) as usize] as char);
        out.push(if chunk.len() > 1 {
            C[((n >> 6) & 63) as usize] as char
        } else {
            '='
        });
        out.push(if chunk.len() > 2 {
            C[(n & 63) as usize] as char
        } else {
            '='
        });
    }
    out
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
        let pad = usize::from(chunk[2] == b'=') + usize::from(chunk[3] == b'=');
        if !last && pad != 0 || chunk[2] == b'=' && chunk[3] != b'=' || pad > 2 {
            return Err(bad("invalid office:binary-data padding"));
        }
        let a = u32::from(val(chunk[0]).ok_or_else(|| bad("invalid office:binary-data base64"))?);
        let b = u32::from(val(chunk[1]).ok_or_else(|| bad("invalid office:binary-data base64"))?);
        let c = if chunk[2] == b'=' {
            0
        } else {
            u32::from(val(chunk[2]).ok_or_else(|| bad("invalid office:binary-data base64"))?)
        };
        let d = if chunk[3] == b'=' {
            0
        } else {
            u32::from(val(chunk[3]).ok_or_else(|| bad("invalid office:binary-data base64"))?)
        };
        let n = a << 18 | b << 12 | c << 6 | d;
        out.push((n >> 16) as u8);
        if pad < 2 {
            out.push((n >> 8) as u8);
        }
        if pad < 1 {
            out.push(n as u8);
        }
        if out.len() > MAX_BINARY {
            return Err(bad("office:binary-data is too large"));
        }
    }
    Ok(out)
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct TargetSpans {
    style: Span,
    properties: Option<Span>,
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

/// Losslessly replace, insert, or remove one existing row style's property element.
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
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target table-row style"));
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
                } else if depth_target.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"table-row-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:table-row-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target table-row style"));
                        }
                        found = Some(TargetSpans {
                            style: span,
                            ..Default::default()
                        });
                    }
                } else if depth_target.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"table-row-properties"
                    && active.as_mut().unwrap().properties.replace(span).is_some()
                {
                    return Err(bad("duplicate style:table-row-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|s| s.end == 0)
                        && depth_target.is_some_and(|d| depth == d + 1)
                    {
                        let s = spans.properties.as_mut().unwrap();
                        s.end_start = begin;
                        s.end = end;
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
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid styles XML: {e}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target table-row style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(Properties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace_span(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand_span(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}

impl Package {
    pub fn row_style_properties(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Default::default()), |xml| parse(&xml))
    }
}
impl FlatDocument {
    pub fn row_style_properties(&self) -> Result<Styles> {
        parse(self.xml())
    }
}
