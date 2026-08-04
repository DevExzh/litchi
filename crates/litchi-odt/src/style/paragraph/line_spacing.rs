//! Typed ODF paragraph line-spacing and justification properties.
//!
//! Models the line-spacing attribute group of `style:paragraph-properties`
//! (`fo:line-height`, `style:line-spacing`, `style:line-height-at-least`,
//! `style:font-independent-line-spacing`, `style:text-autospace`,
//! `style:justify-single-word`, `style:auto-text-indent`,
//! `style:snap-to-layout-grid`, `style:tab-stop-distance`, `fo:text-align-last`).

use crate::{FlatOpenDocument, OdfNonNegativeLength, OpenDocumentPackage};
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
const STYLE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4_096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn name_ok(value: &str, field: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
        return Err(bad(format!("invalid {field}")));
    }
    Ok(())
}
fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(bad(format!("{field} must be an XML Schema boolean"))),
    }
}

/// Valid ODF (possibly negative) physical length for `style:line-spacing`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LineSpacingLength(String);
impl LineSpacingLength {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_VALUE || !physical_length(&value) {
            return Err(bad("style:line-spacing must be an ODF physical length"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
fn physical_length(value: &str) -> bool {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return false;
    };
    let number = number.strip_prefix('-').unwrap_or(number);
    let mut split = number.split('.');
    let whole = split.next().unwrap_or_default();
    let fraction = split.next();
    if split.next().is_some() {
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
fn percent(value: &str) -> bool {
    if value.len() > MAX_VALUE {
        return false;
    }
    let Some(number) = value.strip_suffix('%') else {
        return false;
    };
    let mut split = number.split('.');
    let whole = split.next().unwrap_or_default();
    let fraction = split.next();
    if split.next().is_some() {
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

/// Valid ODF non-negative `percent` lexical value for `fo:line-height`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LineHeightPercent(String);
impl LineHeightPercent {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if !percent(&value) {
            return Err(bad("fo:line-height percent must be an ODF percent"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// The `fo:line-height` value: `normal`, a length, or a percentage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LineHeight {
    Normal,
    Length(OdfNonNegativeLength),
    Percent(LineHeightPercent),
}
impl LineHeight {
    fn parse(value: &str) -> Result<Self> {
        if value == "normal" {
            return Ok(Self::Normal);
        }
        if value.ends_with('%') {
            return Ok(Self::Percent(LineHeightPercent::new(value)?));
        }
        Ok(Self::Length(OdfNonNegativeLength::new(value)?))
    }
    fn xml(&self) -> &str {
        match self {
            Self::Normal => "normal",
            Self::Length(value) => value.as_str(),
            Self::Percent(value) => value.as_str(),
        }
    }
}

/// The `style:text-autospace` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAutospace {
    None,
    IdeographAlpha,
}
impl TextAutospace {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "ideograph-alpha" => Ok(Self::IdeographAlpha),
            _ => Err(bad("invalid style:text-autospace value")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::IdeographAlpha => "ideograph-alpha",
        }
    }
}

/// The `fo:text-align-last` value: alignment of a justified paragraph's last line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignLast {
    Start,
    Center,
    Justify,
}
impl TextAlignLast {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "start" => Ok(Self::Start),
            "center" => Ok(Self::Center),
            "justify" => Ok(Self::Justify),
            _ => Err(bad("invalid fo:text-align-last value")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Start => "start",
            Self::Center => "center",
            Self::Justify => "justify",
        }
    }
}

/// The line-spacing attribute group of one `style:paragraph-properties` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphLineSpacing {
    pub line_height: Option<LineHeight>,
    pub line_spacing: Option<LineSpacingLength>,
    pub line_height_at_least: Option<OdfNonNegativeLength>,
    pub font_independent_line_spacing: Option<bool>,
    pub text_autospace: Option<TextAutospace>,
    pub justify_single_word: Option<bool>,
    pub auto_text_indent: Option<bool>,
    pub snap_to_layout_grid: Option<bool>,
    pub tab_stop_distance: Option<OdfNonNegativeLength>,
    pub text_align_last: Option<TextAlignLast>,
}
impl ParagraphLineSpacing {
    pub fn new() -> Self {
        Self::default()
    }
    fn is_empty(&self) -> bool {
        *self == Self::default()
    }
    pub fn validate(&self) -> Result<()> {
        Ok(())
    }
    /// Emit the properties as a `style:paragraph-properties` fragment.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml =
            format!(r#"<style:paragraph-properties xmlns:style="{STYLE_STR}" xmlns:fo="{FO_STR}""#);
        if let Some(value) = &self.line_height {
            xml.push_str(&format!(r#" fo:line-height="{}""#, escape_xml(value.xml())));
        }
        if let Some(value) = &self.line_spacing {
            xml.push_str(&format!(
                r#" style:line-spacing="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = &self.line_height_at_least {
            xml.push_str(&format!(
                r#" style:line-height-at-least="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = self.font_independent_line_spacing {
            xml.push_str(&format!(
                r#" style:font-independent-line-spacing="{value}""#
            ));
        }
        if let Some(value) = self.text_autospace {
            xml.push_str(&format!(r#" style:text-autospace="{}""#, value.xml()));
        }
        if let Some(value) = self.justify_single_word {
            xml.push_str(&format!(r#" style:justify-single-word="{value}""#));
        }
        if let Some(value) = self.auto_text_indent {
            xml.push_str(&format!(r#" style:auto-text-indent="{value}""#));
        }
        if let Some(value) = self.snap_to_layout_grid {
            xml.push_str(&format!(r#" style:snap-to-layout-grid="{value}""#));
        }
        if let Some(value) = &self.tab_stop_distance {
            xml.push_str(&format!(
                r#" style:tab-stop-distance="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = self.text_align_last {
            xml.push_str(&format!(r#" fo:text-align-last="{}""#, value.xml()));
        }
        xml.push_str("/>");
        Ok(xml)
    }
}

/// A named or default paragraph style and its line-spacing properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphStyleLineSpacing {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<ParagraphLineSpacing>,
}
impl ParagraphStyleLineSpacing {
    pub fn named(
        name: impl Into<String>,
        properties: Option<ParagraphLineSpacing>,
    ) -> Result<Self> {
        let result = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        result.validate()?;
        Ok(result)
    }
    pub fn default_style(properties: Option<ParagraphLineSpacing>) -> Self {
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
}

/// All paragraph styles of a styles part that carry line-spacing properties.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphStyleLineSpacingSet {
    pub styles: Vec<ParagraphStyleLineSpacing>,
}
impl ParagraphStyleLineSpacingSet {
    pub fn get(&self, name: &str) -> Option<&ParagraphStyleLineSpacing> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&ParagraphStyleLineSpacing> {
        self.styles.iter().find(|style| style.is_default_style)
    }
    /// Resolve a style's line-spacing properties through its parent chain,
    /// falling back to the default paragraph style.
    pub fn resolved(&self, name: &str) -> Result<Option<&ParagraphLineSpacing>> {
        let mut current = self.get(name);
        let mut seen = HashSet::new();
        while let Some(style) = current {
            let identity = style.name.as_deref().unwrap_or("<default>");
            if !seen.insert(identity) {
                return Err(bad("paragraph style inheritance cycle"));
            }
            if let Some(properties) = &style.properties {
                return Ok(Some(properties));
            }
            current = style
                .parent_style_name
                .as_deref()
                .and_then(|parent| self.get(parent));
            if style.parent_style_name.is_none() {
                break;
            }
        }
        Ok(self
            .default_style()
            .and_then(|style| style.properties.as_ref()))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Ns {
    Office,
    Style,
    Fo,
    Other,
}
fn known(resolve: ResolveResult<'_>) -> Ns {
    match resolve {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == FO => Ns::Fo,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (known(namespace), local.as_ref().to_vec())
}
fn value(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    attr: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attr.decoded_and_normalized_value(version, reader.decoder())
        .map(|value| value.into_owned())
        .map_err(|error| bad(format!("invalid attribute value: {error}")))
}

fn style_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Option<ParagraphStyleLineSpacing>> {
    let mut name = None;
    let mut parent = None;
    let mut family = None;
    let mut seen = HashSet::new();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|error| bad(format!("invalid style attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        if known(namespace) == Ns::Style {
            if !seen.insert(local.as_ref().to_vec()) {
                return Err(bad("duplicate paragraph style attribute"));
            }
            match local.as_ref() {
                b"name" => name = Some(value(reader, version, &attr)?),
                b"parent-style-name" => parent = Some(value(reader, version, &attr)?),
                b"family" => family = Some(value(reader, version, &attr)?),
                _ => {},
            }
        }
    }
    if family.as_deref() != Some("paragraph") {
        return Ok(None);
    }
    let result = ParagraphStyleLineSpacing {
        name,
        parent_style_name: parent,
        is_default_style: start.local_name().as_ref() == b"default-style",
        properties: None,
    };
    result.validate()?;
    Ok(Some(result))
}

fn property_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ParagraphLineSpacing> {
    let mut properties = ParagraphLineSpacing::new();
    let mut seen = HashSet::new();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|error| bad(format!("invalid property attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        let namespace = known(namespace);
        if !seen.insert(local.as_ref().to_vec()) {
            return Err(bad("duplicate paragraph-properties attribute"));
        }
        let value = value(reader, version, &attr)?;
        if value.len() > MAX_VALUE {
            return Err(bad("paragraph-properties attribute is too large"));
        }
        match (namespace, local.as_ref()) {
            (Ns::Fo, b"line-height") => properties.line_height = Some(LineHeight::parse(&value)?),
            (Ns::Fo, b"text-align-last") => {
                properties.text_align_last = Some(TextAlignLast::parse(&value)?);
            },
            (Ns::Style, b"line-spacing") => {
                properties.line_spacing = Some(LineSpacingLength::new(value)?);
            },
            (Ns::Style, b"line-height-at-least") => {
                properties.line_height_at_least = Some(OdfNonNegativeLength::new(value)?);
            },
            (Ns::Style, b"font-independent-line-spacing") => {
                properties.font_independent_line_spacing =
                    Some(parse_bool(&value, "style:font-independent-line-spacing")?);
            },
            (Ns::Style, b"text-autospace") => {
                properties.text_autospace = Some(TextAutospace::parse(&value)?);
            },
            (Ns::Style, b"justify-single-word") => {
                properties.justify_single_word =
                    Some(parse_bool(&value, "style:justify-single-word")?);
            },
            (Ns::Style, b"auto-text-indent") => {
                properties.auto_text_indent = Some(parse_bool(&value, "style:auto-text-indent")?);
            },
            (Ns::Style, b"snap-to-layout-grid") => {
                properties.snap_to_layout_grid =
                    Some(parse_bool(&value, "style:snap-to-layout-grid")?);
            },
            (Ns::Style, b"tab-stop-distance") => {
                properties.tab_stop_distance = Some(OdfNonNegativeLength::new(value)?);
            },
            // Other paragraph-properties attributes are owned by sibling modules.
            _ => {},
        }
    }
    properties.validate()?;
    Ok(properties)
}

struct Active {
    depth: usize,
    style: ParagraphStyleLineSpacing,
    properties: bool,
}

fn push_style(
    styles: &mut Vec<ParagraphStyleLineSpacing>,
    style: ParagraphStyleLineSpacing,
    total: &mut usize,
) -> Result<()> {
    if styles.len() >= MAX_STYLES {
        return Err(bad("too many paragraph styles"));
    }
    if styles
        .iter()
        .any(|old| old.name == style.name && old.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate paragraph style identity"));
    }
    *total += style.name.as_deref().map_or(0, str::len)
        + style.parent_style_name.as_deref().map_or(0, str::len);
    if *total > MAX_TOTAL {
        return Err(bad("paragraph line-spacing data is too large"));
    }
    styles.push(style);
    Ok(())
}

fn is_paragraph_style(current: &(Ns, Vec<u8>), parent: Option<&(Ns, Vec<u8>)>) -> bool {
    parent.is_some_and(|(n, l)| {
        *n == Ns::Office && matches!(l.as_slice(), b"styles" | b"automatic-styles")
    }) && current.0 == Ns::Style
        && matches!(current.1.as_slice(), b"style" | b"default-style")
}

/// Parse paragraph styles and their line-spacing properties from a styles part.
pub fn parse_paragraph_style_line_spacings(xml: &str) -> Result<ParagraphStyleLineSpacingSet> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    if !xml.contains("paragraph-properties") {
        return Ok(ParagraphStyleLineSpacingSet::default());
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
                    if let Some(style) = style_attrs(&reader, version, &start)? {
                        active = Some(Active {
                            depth,
                            style,
                            properties: false,
                        });
                    }
                    continue;
                }
                if let Some(state) = active.as_mut()
                    && depth == state.depth + 1
                    && current.0 == Ns::Style
                    && current.1 == b"paragraph-properties"
                {
                    if state.properties {
                        return Err(bad("duplicate style:paragraph-properties"));
                    }
                    state.properties = true;
                    let properties = property_attrs(&reader, version, &start)?;
                    if !properties.is_empty() {
                        state.style.properties = Some(properties);
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                if direct {
                    if let Some(style) = style_attrs(&reader, version, &start)? {
                        push_style(&mut styles, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    let depth = stack.len() + 1;
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"paragraph-properties"
                    {
                        if state.properties {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        state.properties = true;
                        let properties = property_attrs(&reader, version, &start)?;
                        if !properties.is_empty() {
                            state.style.properties = Some(properties);
                        }
                    }
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if active.as_ref().is_some_and(|state| state.depth == depth) {
                    push_style(&mut styles, active.take().unwrap().style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
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
    Ok(ParagraphStyleLineSpacingSet { styles })
}

impl OpenDocumentPackage {
    /// Parse the paragraph line-spacing properties of the package `styles.xml`.
    pub fn paragraph_style_line_spacings(&self) -> Result<ParagraphStyleLineSpacingSet> {
        self.styles_xml()?.map_or_else(
            || Ok(ParagraphStyleLineSpacingSet::default()),
            |xml| parse_paragraph_style_line_spacings(&xml),
        )
    }
}
impl FlatOpenDocument {
    /// Parse the paragraph line-spacing properties of a flat XML document.
    pub fn paragraph_style_line_spacings(&self) -> Result<ParagraphStyleLineSpacingSet> {
        parse_paragraph_style_line_spacings(self.xml())
    }
}
