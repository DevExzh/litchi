//! Typed residual `style:section-properties` support.
//!
//! `style:columns` and `text:notes-configuration` remain owned by their
//! dedicated modules. This module validates their grammar placement without
//! duplicating their data models.

use litchi_core::{Error, Result};
use quick_xml::{
    NsReader, XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_VALUE: usize = 4096;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn bounded(value: &str, label: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
        return Err(invalid(format!("invalid {label}")));
    }
    Ok(())
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('>', "&gt;")
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SectionLength(String);
impl SectionLength {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        bounded(&value, "section length")?;
        let unit = ["cm", "mm", "in", "pt", "pc", "px"]
            .into_iter()
            .find(|unit| value.ends_with(unit))
            .ok_or_else(|| invalid("section length requires a physical unit"))?;
        let number = &value[..value.len() - unit.len()];
        if number.is_empty()
            || matches!(number, "+" | "-")
            || number.contains(['e', 'E'])
            || number
                .parse::<f64>()
                .map_or(true, |number| !number.is_finite())
        {
            return Err(invalid("invalid section length number"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SectionBackgroundColor {
    Transparent,
    Rgb(u8, u8, u8),
}
impl SectionBackgroundColor {
    fn parse(value: &str) -> Result<Self> {
        if value == "transparent" {
            return Ok(Self::Transparent);
        }
        if value.len() != 7 || !value.starts_with('#') {
            return Err(invalid(
                "section background color must be transparent or #RRGGBB",
            ));
        }
        let rgb = u32::from_str_radix(&value[1..], 16)
            .map_err(|_| invalid("invalid section background color"))?;
        Ok(Self::Rgb((rgb >> 16) as u8, (rgb >> 8) as u8, rgb as u8))
    }
    fn lexical(self) -> String {
        match self {
            Self::Transparent => "transparent".into(),
            Self::Rgb(r, g, b) => format!("#{r:02X}{g:02X}{b:02X}"),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SectionWritingMode {
    LeftToRightTopToBottom,
    RightToLeftTopToBottom,
    TopToBottomRightToLeft,
    TopToBottomLeftToRight,
    Page,
}
impl SectionWritingMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "lr-tb" => Ok(Self::LeftToRightTopToBottom),
            "rl-tb" => Ok(Self::RightToLeftTopToBottom),
            "tb-rl" => Ok(Self::TopToBottomRightToLeft),
            "tb-lr" => Ok(Self::TopToBottomLeftToRight),
            "page" => Ok(Self::Page),
            _ => Err(invalid("invalid section writing mode")),
        }
    }
    fn lexical(self) -> &'static str {
        match self {
            Self::LeftToRightTopToBottom => "lr-tb",
            Self::RightToLeftTopToBottom => "rl-tb",
            Self::TopToBottomRightToLeft => "tb-rl",
            Self::TopToBottomLeftToRight => "tb-lr",
            Self::Page => "page",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BackgroundRepeat {
    Repeat,
    Stretch,
    NoRepeat,
}
impl BackgroundRepeat {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "repeat" => Ok(Self::Repeat),
            "stretch" => Ok(Self::Stretch),
            "no-repeat" => Ok(Self::NoRepeat),
            _ => Err(invalid("invalid section background repeat")),
        }
    }
    fn lexical(self) -> &'static str {
        match self {
            Self::Repeat => "repeat",
            Self::Stretch => "stretch",
            Self::NoRepeat => "no-repeat",
        }
    }
}

/// Background image references remain inert; no target is opened or fetched.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SectionBackgroundImage {
    pub href: Option<String>,
    pub repeat: Option<BackgroundRepeat>,
    pub position: Option<String>,
    pub filter_name: Option<String>,
    pub opacity_percent: Option<u8>,
    pub xlink_type: Option<String>,
    pub show: Option<String>,
    pub actuate: Option<String>,
}
impl SectionBackgroundImage {
    pub fn validate(&self) -> Result<()> {
        for (value, label) in [
            (&self.href, "background href"),
            (&self.position, "background position"),
            (&self.filter_name, "background filter"),
            (&self.xlink_type, "xlink type"),
            (&self.show, "xlink show"),
            (&self.actuate, "xlink actuate"),
        ] {
            if let Some(value) = value {
                bounded(value, label)?;
            }
        }
        if self.xlink_type.as_deref().is_some_and(|v| v != "simple") {
            return Err(invalid("xlink:type must be simple"));
        }
        if self
            .show
            .as_deref()
            .is_some_and(|v| !matches!(v, "embed" | "new" | "replace"))
        {
            return Err(invalid("invalid xlink:show"));
        }
        if self.actuate.as_deref().is_some_and(|v| v != "onLoad") {
            return Err(invalid("xlink:actuate must be onLoad"));
        }
        if self.opacity_percent.is_some_and(|v| v > 100) {
            return Err(invalid("draw:opacity exceeds 100%"));
        }
        Ok(())
    }
    fn write_xml(&self, xml: &mut String) -> Result<()> {
        self.validate()?;
        xml.push_str("<style:background-image");
        if let Some(v) = &self.href {
            xml.push_str(" xlink:href=\"");
            xml.push_str(&escape(v));
            xml.push('"');
        }
        if let Some(v) = self.repeat {
            xml.push_str(" style:repeat=\"");
            xml.push_str(v.lexical());
            xml.push('"');
        }
        if let Some(v) = &self.position {
            xml.push_str(" style:position=\"");
            xml.push_str(&escape(v));
            xml.push('"');
        }
        if let Some(v) = &self.filter_name {
            xml.push_str(" style:filter-name=\"");
            xml.push_str(&escape(v));
            xml.push('"');
        }
        if let Some(v) = self.opacity_percent {
            xml.push_str(&format!(" draw:opacity=\"{v}%\""));
        }
        if let Some(v) = &self.xlink_type {
            xml.push_str(" xlink:type=\"");
            xml.push_str(&escape(v));
            xml.push('"');
        }
        if let Some(v) = &self.show {
            xml.push_str(" xlink:show=\"");
            xml.push_str(&escape(v));
            xml.push('"');
        }
        if let Some(v) = &self.actuate {
            xml.push_str(" xlink:actuate=\"");
            xml.push_str(&escape(v));
            xml.push('"');
        }
        xml.push_str("/>");
        Ok(())
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SectionProperties {
    pub background_color: Option<SectionBackgroundColor>,
    pub margin_left: Option<SectionLength>,
    pub margin_right: Option<SectionLength>,
    pub editable: Option<bool>,
    pub protect: Option<bool>,
    pub writing_mode: Option<SectionWritingMode>,
    pub dont_balance_text_columns: Option<bool>,
    pub background_image: Option<SectionBackgroundImage>,
}
impl SectionProperties {
    pub fn validate(&self) -> Result<()> {
        if let Some(image) = &self.background_image {
            image.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            "<style:section-properties xmlns:style=\"{STYLE}\" xmlns:text=\"{TEXT}\" xmlns:fo=\"{FO}\" xmlns:draw=\"{DRAW}\" xmlns:xlink=\"{XLINK}\""
        );
        if let Some(v) = self.background_color {
            xml.push_str(" fo:background-color=\"");
            xml.push_str(&v.lexical());
            xml.push('"');
        }
        if let Some(v) = &self.margin_left {
            xml.push_str(" fo:margin-left=\"");
            xml.push_str(v.as_str());
            xml.push('"');
        }
        if let Some(v) = &self.margin_right {
            xml.push_str(" fo:margin-right=\"");
            xml.push_str(v.as_str());
            xml.push('"');
        }
        if let Some(v) = self.editable {
            xml.push_str(if v {
                " style:editable=\"true\""
            } else {
                " style:editable=\"false\""
            });
        }
        if let Some(v) = self.protect {
            xml.push_str(if v {
                " style:protect=\"true\""
            } else {
                " style:protect=\"false\""
            });
        }
        if let Some(v) = self.writing_mode {
            xml.push_str(" style:writing-mode=\"");
            xml.push_str(v.lexical());
            xml.push('"');
        }
        if let Some(v) = self.dont_balance_text_columns {
            xml.push_str(if v {
                " text:dont-balance-text-columns=\"true\""
            } else {
                " text:dont-balance-text-columns=\"false\""
            });
        }
        if let Some(image) = &self.background_image {
            xml.push('>');
            image.write_xml(&mut xml)?;
            xml.push_str("</style:section-properties>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SectionStyleProperties {
    pub name: String,
    pub properties: SectionProperties,
}
impl SectionStyleProperties {
    pub fn new(name: impl Into<String>, properties: SectionProperties) -> Result<Self> {
        let result = Self {
            name: name.into(),
            properties,
        };
        result.validate()?;
        Ok(result)
    }
    pub fn validate(&self) -> Result<()> {
        bounded(&self.name, "section style name")?;
        self.properties.validate()
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        Ok(format!(
            "<style:style xmlns:style=\"{STYLE}\" style:name=\"{}\" style:family=\"section\">{}</style:style>",
            escape(&self.name),
            self.properties.to_xml_fragment()?
        ))
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SectionStylePropertiesSet {
    pub styles: Vec<SectionStyleProperties>,
}
impl SectionStylePropertiesSet {
    pub fn get(&self, name: &str) -> Option<&SectionStyleProperties> {
        self.styles.iter().find(|style| style.name == name)
    }
}

fn is_ns(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace) if namespace.as_ref() == expected.as_bytes())
}
fn attr_value(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    attr: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attr.decoded_and_normalized_value(version, reader.decoder())
        .map(|v| v.into_owned())
        .map_err(|e| invalid(format!("invalid XML attribute: {e}")))
}
fn boolean(value: &str, label: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid(format!("{label} must be true or false"))),
    }
}
fn opacity(value: &str) -> Result<u8> {
    let value = value
        .strip_suffix('%')
        .ok_or_else(|| invalid("draw:opacity must be a percentage"))?
        .parse::<u8>()
        .map_err(|_| invalid("invalid draw:opacity"))?;
    if value > 100 {
        return Err(invalid("draw:opacity exceeds 100%"));
    }
    Ok(value)
}

fn parse_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<SectionProperties> {
    let mut p = SectionProperties::default();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|e| invalid(format!("invalid section attribute: {e}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (ns, local) = reader.resolver().resolve_attribute(attr.key);
        let value = attr_value(reader, version, &attr)?;
        bounded(&value, "section property value")?;
        match (ns, local.as_ref()) {
            (ResolveResult::Bound(n), b"background-color") if n.as_ref() == FO.as_bytes() => {
                if p.background_color
                    .replace(SectionBackgroundColor::parse(&value)?)
                    .is_some()
                {
                    return Err(invalid("duplicate fo:background-color"));
                }
            },
            (ResolveResult::Bound(n), b"margin-left") if n.as_ref() == FO.as_bytes() => {
                if p.margin_left.replace(SectionLength::new(value)?).is_some() {
                    return Err(invalid("duplicate fo:margin-left"));
                }
            },
            (ResolveResult::Bound(n), b"margin-right") if n.as_ref() == FO.as_bytes() => {
                if p.margin_right.replace(SectionLength::new(value)?).is_some() {
                    return Err(invalid("duplicate fo:margin-right"));
                }
            },
            (ResolveResult::Bound(n), b"editable") if n.as_ref() == STYLE.as_bytes() => {
                if p.editable
                    .replace(boolean(&value, "style:editable")?)
                    .is_some()
                {
                    return Err(invalid("duplicate style:editable"));
                }
            },
            (ResolveResult::Bound(n), b"protect") if n.as_ref() == STYLE.as_bytes() => {
                if p.protect
                    .replace(boolean(&value, "style:protect")?)
                    .is_some()
                {
                    return Err(invalid("duplicate style:protect"));
                }
            },
            (ResolveResult::Bound(n), b"writing-mode") if n.as_ref() == STYLE.as_bytes() => {
                if p.writing_mode
                    .replace(SectionWritingMode::parse(&value)?)
                    .is_some()
                {
                    return Err(invalid("duplicate style:writing-mode"));
                }
            },
            (ResolveResult::Bound(n), b"dont-balance-text-columns")
                if n.as_ref() == TEXT.as_bytes() =>
            {
                if p.dont_balance_text_columns
                    .replace(boolean(&value, "text:dont-balance-text-columns")?)
                    .is_some()
                {
                    return Err(invalid("duplicate text:dont-balance-text-columns"));
                }
            },
            _ => {
                return Err(invalid(
                    "unsupported or wrongly namespaced section-properties attribute",
                ));
            },
        }
    }
    Ok(p)
}

fn parse_background(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<SectionBackgroundImage> {
    let mut image = SectionBackgroundImage::default();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|e| invalid(format!("invalid background attribute: {e}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (ns, local) = reader.resolver().resolve_attribute(attr.key);
        let value = attr_value(reader, version, &attr)?;
        bounded(&value, "background value")?;
        match (ns, local.as_ref()) {
            (ResolveResult::Bound(n), b"href") if n.as_ref() == XLINK.as_bytes() => {
                if image.href.replace(value).is_some() {
                    return Err(invalid("duplicate xlink:href"));
                }
            },
            (ResolveResult::Bound(n), b"repeat") if n.as_ref() == STYLE.as_bytes() => {
                if image
                    .repeat
                    .replace(BackgroundRepeat::parse(&value)?)
                    .is_some()
                {
                    return Err(invalid("duplicate style:repeat"));
                }
            },
            (ResolveResult::Bound(n), b"position") if n.as_ref() == STYLE.as_bytes() => {
                if image.position.replace(value).is_some() {
                    return Err(invalid("duplicate style:position"));
                }
            },
            (ResolveResult::Bound(n), b"filter-name") if n.as_ref() == STYLE.as_bytes() => {
                if image.filter_name.replace(value).is_some() {
                    return Err(invalid("duplicate style:filter-name"));
                }
            },
            (ResolveResult::Bound(n), b"opacity") if n.as_ref() == DRAW.as_bytes() => {
                if image.opacity_percent.replace(opacity(&value)?).is_some() {
                    return Err(invalid("duplicate draw:opacity"));
                }
            },
            (ResolveResult::Bound(n), b"type") if n.as_ref() == XLINK.as_bytes() => {
                if image.xlink_type.replace(value).is_some() {
                    return Err(invalid("duplicate xlink:type"));
                }
            },
            (ResolveResult::Bound(n), b"show") if n.as_ref() == XLINK.as_bytes() => {
                if image.show.replace(value).is_some() {
                    return Err(invalid("duplicate xlink:show"));
                }
            },
            (ResolveResult::Bound(n), b"actuate") if n.as_ref() == XLINK.as_bytes() => {
                if image.actuate.replace(value).is_some() {
                    return Err(invalid("duplicate xlink:actuate"));
                }
            },
            _ => {
                return Err(invalid(
                    "unsupported or wrongly namespaced background-image attribute",
                ));
            },
        }
    }
    image.validate()?;
    Ok(image)
}

fn style_name(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Option<String>> {
    let mut family = None;
    let mut name = None;
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|e| invalid(format!("invalid style attribute: {e}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (ns, local) = reader.resolver().resolve_attribute(attr.key);
        if !is_ns(&ns, STYLE) {
            continue;
        }
        let value = attr_value(reader, version, &attr)?;
        match local.as_ref() {
            b"family" => family = Some(value),
            b"name" => name = Some(value),
            _ => {},
        }
    }
    if family.as_deref() != Some("section") {
        return Ok(None);
    }
    let name = name.ok_or_else(|| invalid("section style lacks style:name"))?;
    bounded(&name, "section style name")?;
    Ok(Some(name))
}

/// Parse section styles from `styles.xml` or a flat ODF document.
pub fn parse_section_style_properties(xml: &[u8]) -> Result<SectionStylePropertiesSet> {
    if xml.len() > MAX_XML {
        return Err(invalid("section-properties XML exceeds size cap"));
    }
    let xml_text =
        std::str::from_utf8(xml).map_err(|_| invalid("section-properties XML is not UTF-8"))?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut container_depth = None;
    let mut style_depth = None;
    let mut name = None;
    let mut properties = None;
    let mut properties_depth = None;
    let mut background_depth = None;
    let mut last_rank = 0u8;
    let mut seen = [false; 3];
    let mut aggregate = 0usize;
    let mut result = SectionStylePropertiesSet::default();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| invalid(format!("invalid section-properties XML: {e}")))?;
        match event {
            Event::Start(start) => {
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(invalid("XML depth exceeds cap"));
                }
                let local = start.local_name();
                if background_depth.is_some() {
                    return Err(invalid("style:background-image must be empty"));
                }
                if container_depth.is_none()
                    && is_ns(&namespace, OFFICE)
                    && matches!(local.as_ref(), b"styles" | b"automatic-styles")
                {
                    container_depth = Some(depth);
                } else if container_depth.is_some_and(|d| depth == d + 1)
                    && is_ns(&namespace, STYLE)
                    && local.as_ref() == b"style"
                {
                    if let Some(value) = style_name(&reader, version, &start)? {
                        style_depth = Some(depth);
                        name = Some(value);
                        properties = None;
                    }
                } else if style_depth.is_some_and(|d| depth == d + 1)
                    && is_ns(&namespace, STYLE)
                    && local.as_ref() == b"section-properties"
                {
                    if properties.is_some() || properties_depth.is_some() {
                        return Err(invalid("duplicate style:section-properties"));
                    }
                    properties = Some(parse_properties(&reader, version, &start)?);
                    properties_depth = Some(depth);
                    last_rank = 0;
                    seen = [false; 3];
                } else if properties_depth.is_some_and(|d| depth == d + 1) {
                    let rank = child_rank(&namespace, local.as_ref())?;
                    check_child(rank, &mut last_rank, &mut seen)?;
                    if rank == 1 {
                        properties.as_mut().expect("state").background_image =
                            Some(parse_background(&reader, version, &start)?);
                        background_depth = Some(depth);
                    }
                }
            },
            Event::Empty(start) => {
                if background_depth.is_some() {
                    return Err(invalid("style:background-image must be empty"));
                }
                let local = start.local_name();
                let event_depth = depth + 1;
                if style_depth.is_some_and(|d| event_depth == d + 1)
                    && is_ns(&namespace, STYLE)
                    && local.as_ref() == b"section-properties"
                {
                    if properties.is_some() {
                        return Err(invalid("duplicate style:section-properties"));
                    }
                    properties = Some(parse_properties(&reader, version, &start)?);
                } else if properties_depth.is_some_and(|d| event_depth == d + 1) {
                    let rank = child_rank(&namespace, local.as_ref())?;
                    check_child(rank, &mut last_rank, &mut seen)?;
                    if rank == 1 {
                        properties.as_mut().expect("state").background_image =
                            Some(parse_background(&reader, version, &start)?);
                    }
                }
            },
            Event::End(end) => {
                let local = end.local_name();
                if background_depth == Some(depth) {
                    if !is_ns(&namespace, STYLE) || local.as_ref() != b"background-image" {
                        return Err(invalid("malformed background-image"));
                    }
                    background_depth = None;
                }
                if properties_depth == Some(depth) {
                    properties_depth = None;
                }
                if style_depth == Some(depth) {
                    if let (Some(name), Some(properties)) = (name.take(), properties.take()) {
                        aggregate = aggregate
                            .checked_add(name.len())
                            .ok_or_else(|| invalid("aggregate overflow"))?;
                        if aggregate > MAX_AGGREGATE || result.styles.len() >= MAX_STYLES {
                            return Err(invalid("section style resource cap exceeded"));
                        }
                        if result.styles.iter().any(|item| item.name == name) {
                            return Err(invalid("duplicate section style name"));
                        }
                        result
                            .styles
                            .push(SectionStyleProperties::new(name, properties)?);
                    }
                    style_depth = None;
                }
                if container_depth == Some(depth) {
                    container_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced XML depth"))?;
            },
            Event::Text(text) => {
                if background_depth.is_some() || properties_depth == Some(depth) {
                    let value = text
                        .decode()
                        .map_err(|e| invalid(format!("invalid XML text: {e}")))?;
                    if !value.trim().is_empty() {
                        return Err(invalid("unexpected section-properties text"));
                    }
                }
            },
            Event::CData(data) => {
                if (background_depth.is_some() || properties_depth == Some(depth))
                    && !data.is_empty()
                {
                    return Err(invalid("section-properties contains CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are forbidden"));
            },
            Event::Decl(declaration) => {
                version = declaration
                    .xml_version()
                    .map_err(|e| invalid(format!("unsupported XML version: {e}")))?
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    crate::notes_configuration::parse_notes_configurations(xml_text)?;
    Ok(result)
}

fn child_rank(namespace: &ResolveResult<'_>, local: &[u8]) -> Result<u8> {
    if is_ns(namespace, STYLE) && local == b"background-image" {
        Ok(1)
    } else if is_ns(namespace, STYLE) && local == b"columns" {
        Ok(2)
    } else if is_ns(namespace, TEXT) && local == b"notes-configuration" {
        Ok(3)
    } else {
        Err(invalid(
            "unsupported or wrongly namespaced section-properties child",
        ))
    }
}
fn check_child(rank: u8, last: &mut u8, seen: &mut [bool; 3]) -> Result<()> {
    if rank < *last || seen[(rank - 1) as usize] {
        return Err(invalid(
            "section-properties child order or cardinality violation",
        ));
    }
    *last = rank;
    seen[(rank - 1) as usize] = true;
    Ok(())
}

impl crate::OpenDocumentPackage {
    pub fn section_style_properties(&self) -> Result<SectionStylePropertiesSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_section_style_properties(xml.as_bytes()),
        )
    }
}
impl crate::FlatOpenDocument {
    pub fn section_style_properties(&self) -> Result<SectionStylePropertiesSet> {
        parse_section_style_properties(self.xml().as_bytes())
    }
}
