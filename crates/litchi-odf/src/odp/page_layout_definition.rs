//! Typed ODF presentation page-layout definitions.

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const PRESENTATION_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_LAYOUTS: usize = 65_536;
const MAX_PLACEHOLDERS: usize = 4_096;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// Standard placeholder role from the ODF `presentation-classes` vocabulary.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PresentationPlaceholderClass {
    Title,
    Outline,
    Subtitle,
    Text,
    Graphic,
    Object,
    Chart,
    Table,
    OrganizationChart,
    Page,
    Notes,
    Handout,
    Header,
    Footer,
    DateTime,
    PageNumber,
}

impl PresentationPlaceholderClass {
    fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "title" => Self::Title,
            "outline" => Self::Outline,
            "subtitle" => Self::Subtitle,
            "text" => Self::Text,
            "graphic" => Self::Graphic,
            "object" => Self::Object,
            "chart" => Self::Chart,
            "table" => Self::Table,
            "orgchart" => Self::OrganizationChart,
            "page" => Self::Page,
            "notes" => Self::Notes,
            "handout" => Self::Handout,
            "header" => Self::Header,
            "footer" => Self::Footer,
            "date-time" => Self::DateTime,
            "page-number" => Self::PageNumber,
            _ => {
                return invalid(format!(
                    "unsupported presentation placeholder class '{value}'"
                ));
            },
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Outline => "outline",
            Self::Subtitle => "subtitle",
            Self::Text => "text",
            Self::Graphic => "graphic",
            Self::Object => "object",
            Self::Chart => "chart",
            Self::Table => "table",
            Self::OrganizationChart => "orgchart",
            Self::Page => "page",
            Self::Notes => "notes",
            Self::Handout => "handout",
            Self::Header => "header",
            Self::Footer => "footer",
            Self::DateTime => "date-time",
            Self::PageNumber => "page-number",
        }
    }
}

/// Unit accepted by ODF presentation placeholder coordinates and extents.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PresentationMeasureUnit {
    Centimeter,
    Millimeter,
    Inch,
    Point,
    Pica,
    Pixel,
    Percent,
}

impl PresentationMeasureUnit {
    const fn suffix(self) -> &'static str {
        match self {
            Self::Centimeter => "cm",
            Self::Millimeter => "mm",
            Self::Inch => "in",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::Pixel => "px",
            Self::Percent => "%",
        }
    }
}

/// A finite ODF length or percentage with a typed unit.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresentationMeasure {
    value: f64,
    unit: PresentationMeasureUnit,
}

impl PresentationMeasure {
    pub fn new(value: f64, unit: PresentationMeasureUnit) -> Result<Self> {
        if !value.is_finite() {
            return invalid("presentation measure must be finite");
        }
        Ok(Self { value, unit })
    }

    pub const fn value(self) -> f64 {
        self.value
    }

    pub const fn unit(self) -> PresentationMeasureUnit {
        self.unit
    }
}

impl FromStr for PresentationMeasure {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let (number, unit) = if let Some(number) = value.strip_suffix('%') {
            (number, PresentationMeasureUnit::Percent)
        } else if value.len() >= 2 {
            let (number, suffix) = value.split_at(value.len() - 2);
            let unit = match suffix {
                "cm" => PresentationMeasureUnit::Centimeter,
                "mm" => PresentationMeasureUnit::Millimeter,
                "in" => PresentationMeasureUnit::Inch,
                "pt" => PresentationMeasureUnit::Point,
                "pc" => PresentationMeasureUnit::Pica,
                "px" => PresentationMeasureUnit::Pixel,
                _ => return invalid(format!("invalid presentation measure '{value}'")),
            };
            (number, unit)
        } else {
            return invalid(format!("invalid presentation measure '{value}'"));
        };
        validate_decimal(number, value)?;
        let number = number
            .parse::<f64>()
            .map_err(|_| make_error(format!("invalid presentation measure '{value}'")))?;
        Self::new(number, unit)
    }
}

impl fmt::Display for PresentationMeasure {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let value = if self.value == 0.0 { 0.0 } else { self.value };
        write!(formatter, "{}{}", value, self.unit.suffix())
    }
}

/// One required placeholder rectangle in a named presentation page layout.
#[derive(Clone, Debug, PartialEq)]
pub struct PresentationPlaceholder {
    pub class: PresentationPlaceholderClass,
    pub x: PresentationMeasure,
    pub y: PresentationMeasure,
    pub width: PresentationMeasure,
    pub height: PresentationMeasure,
}

impl PresentationPlaceholder {
    pub fn new(
        class: PresentationPlaceholderClass,
        x: PresentationMeasure,
        y: PresentationMeasure,
        width: PresentationMeasure,
        height: PresentationMeasure,
    ) -> Self {
        Self {
            class,
            x,
            y,
            width,
            height,
        }
    }
}

/// A named custom presentation layout and its ordered placeholders.
#[derive(Clone, Debug, PartialEq)]
pub struct PresentationPageLayout {
    pub name: String,
    pub display_name: Option<String>,
    pub placeholders: Vec<PresentationPlaceholder>,
}

impl PresentationPageLayout {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            display_name: None,
            placeholders: Vec::new(),
        };
        value.validate()?;
        Ok(value)
    }

    pub fn validate(&self) -> Result<()> {
        validate_ncname(&self.name, "presentation page-layout name")?;
        if let Some(value) = &self.display_name {
            validate_text(value, "presentation page-layout display name", true)?;
        }
        if self.placeholders.len() > MAX_PLACEHOLDERS {
            return invalid(format!(
                "presentation page layout exceeds {MAX_PLACEHOLDERS} placeholders"
            ));
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(160 + self.placeholders.len() * 128);
        write_layout(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered presentation page-layout definitions from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PresentationPageLayouts {
    pub layouts: Vec<PresentationPageLayout>,
}

impl PresentationPageLayouts {
    pub fn get(&self, name: &str) -> Option<&PresentationPageLayout> {
        self.layouts.iter().find(|layout| layout.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.layouts.len() > MAX_LAYOUTS {
            return invalid(format!(
                "presentation styles exceed {MAX_LAYOUTS} page layouts"
            ));
        }
        let mut names = HashSet::with_capacity(self.layouts.len());
        let mut aggregate = 0usize;
        for layout in &self.layouts {
            layout.validate()?;
            if !names.insert(layout.name.as_str()) {
                return invalid(format!(
                    "duplicate presentation page-layout name '{}'",
                    layout.name
                ));
            }
            aggregate = aggregate
                .checked_add(layout.name.len())
                .and_then(|value| {
                    value.checked_add(layout.display_name.as_deref().map_or(0, str::len))
                })
                .ok_or_else(|| make_error("presentation page-layout size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("presentation page-layout values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize a standalone schema-valid `office:styles` fragment.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(256 + self.layouts.len() * 192);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">"#,
        );
        for layout in &self.layouts {
            write_layout(&mut output, layout, false);
        }
        output.push_str("</office:styles>");
        Ok(output)
    }
}

#[derive(Clone)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
}

struct ActiveLayout {
    depth: usize,
    value: PresentationPageLayout,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind {
    None,
    Office,
    Style,
    Presentation,
    Svg,
    Other,
}

type Attributes = HashMap<(NamespaceKind, String), String>;

/// Parse page-layout definitions from an ODF styles or flat-document XML part.
pub fn parse_presentation_page_layouts(xml: &str) -> Result<PresentationPageLayouts> {
    if !xml.contains("presentation-page-layout") {
        return Ok(PresentationPageLayouts::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("presentation page-layout XML exceeds 64 MiB");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveLayout> = None;
    let mut layouts = PresentationPageLayouts::default();

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                make_error(format!("invalid presentation page-layout XML: {error}"))
            })?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode_name(element.local_name().as_ref(), "element")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    if namespace == NamespaceKind::Presentation && local == "placeholder" {
                        return invalid("presentation:placeholder must be empty");
                    }
                    return invalid("presentation page layout may contain only empty placeholders");
                }
                if namespace == NamespaceKind::Style && local == "presentation-page-layout" {
                    ensure_location(&stack)?;
                    if layouts.layouts.len() >= MAX_LAYOUTS {
                        return invalid(format!(
                            "presentation styles exceed {MAX_LAYOUTS} page layouts"
                        ));
                    }
                    active = Some(ActiveLayout {
                        depth: stack.len(),
                        value: parse_layout(&reader, element)?,
                    });
                } else if namespace == NamespaceKind::Presentation && local == "placeholder" {
                    return invalid("presentation:placeholder must be inside a page layout");
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!(
                        "presentation page-layout XML exceeds {MAX_DEPTH} levels"
                    ));
                }
            },
            Event::Empty(ref element) => {
                let local = decode_name(element.local_name().as_ref(), "element")?;
                reject_spoofed_name(namespace, &local)?;
                if namespace == NamespaceKind::Style && local == "presentation-page-layout" {
                    if active.is_some() {
                        return invalid("presentation page layout cannot contain another layout");
                    }
                    ensure_location(&stack)?;
                    layouts.layouts.push(parse_layout(&reader, element)?);
                } else if namespace == NamespaceKind::Presentation && local == "placeholder" {
                    let Some(layout) = active.as_mut() else {
                        return invalid("presentation:placeholder must be inside a page layout");
                    };
                    if stack.len() != layout.depth + 1 {
                        return invalid("presentation:placeholder must be a direct layout child");
                    }
                    if layout.value.placeholders.len() >= MAX_PLACEHOLDERS {
                        return invalid(format!(
                            "presentation page layout exceeds {MAX_PLACEHOLDERS} placeholders"
                        ));
                    }
                    layout
                        .value
                        .placeholders
                        .push(parse_placeholder(&reader, element)?);
                } else if active.is_some() {
                    return invalid("presentation page layout contains an unsupported child");
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("presentation page-layout XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|layout| layout.depth == stack.len())
                {
                    if frame.namespace != NamespaceKind::Style
                        || frame.local != "presentation-page-layout"
                    {
                        return invalid("unexpected presentation page-layout end element");
                    }
                    layouts
                        .layouts
                        .push(active.take().expect("active layout checked").value);
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid page-layout text: {error}")))?;
                if !value.chars().all(char::is_whitespace) {
                    return invalid("presentation page layout cannot contain text");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("presentation page layout cannot contain character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in page layouts");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated presentation page-layout XML");
    }
    layouts.validate()?;
    Ok(layouts)
}

impl OpenDocumentPackage {
    /// Inspect named presentation page layouts in packaged `styles.xml`.
    pub fn presentation_page_layouts(&self) -> Result<PresentationPageLayouts> {
        self.styles_xml()?.map_or_else(
            || Ok(PresentationPageLayouts::default()),
            |xml| parse_presentation_page_layouts(&xml),
        )
    }
}

impl FlatOpenDocument {
    /// Inspect named presentation page layouts in a flat presentation.
    pub fn presentation_page_layouts(&self) -> Result<PresentationPageLayouts> {
        parse_presentation_page_layouts(self.xml())
    }
}

fn ensure_location(stack: &[Frame]) -> Result<()> {
    if !matches!(stack.last(), Some(Frame { namespace: NamespaceKind::Office, local }) if local == "styles")
    {
        return invalid("style:presentation-page-layout must be a direct office:styles child");
    }
    Ok(())
}

fn parse_layout(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<PresentationPageLayout> {
    let mut attributes = attributes(reader, element)?;
    let name = take_required(&mut attributes, NamespaceKind::Style, "name", "style:name")?;
    let display_name = attributes.remove(&(NamespaceKind::Style, "display-name".to_string()));
    reject_attributes(&attributes, "style:presentation-page-layout")?;
    let value = PresentationPageLayout {
        name,
        display_name,
        placeholders: Vec::new(),
    };
    value.validate()?;
    Ok(value)
}

fn parse_placeholder(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<PresentationPlaceholder> {
    let mut attributes = attributes(reader, element)?;
    let class = PresentationPlaceholderClass::parse(&take_required(
        &mut attributes,
        NamespaceKind::Presentation,
        "object",
        "presentation:object",
    )?)?;
    let x = take_measure(&mut attributes, "x", "svg:x")?;
    let y = take_measure(&mut attributes, "y", "svg:y")?;
    let width = take_measure(&mut attributes, "width", "svg:width")?;
    let height = take_measure(&mut attributes, "height", "svg:height")?;
    reject_attributes(&attributes, "presentation:placeholder")?;
    Ok(PresentationPlaceholder::new(class, x, y, width, height))
}

fn take_measure(
    attributes: &mut Attributes,
    local: &str,
    context: &str,
) -> Result<PresentationMeasure> {
    take_required(attributes, NamespaceKind::Svg, local, context)?.parse()
}

fn take_required(
    attributes: &mut Attributes,
    namespace: NamespaceKind,
    local: &str,
    context: &str,
) -> Result<String> {
    attributes
        .remove(&(namespace, local.to_string()))
        .ok_or_else(|| make_error(format!("missing required {context}")))
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid page-layout attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode_name(local.as_ref(), "attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid page-layout attribute value: {error}")))?
            .into_owned();
        validate_text(&value, "presentation page-layout attribute", true)?;
        if result.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded presentation page-layout attribute");
        }
    }
    Ok(result)
}

fn reject_attributes(attributes: &Attributes, context: &str) -> Result<()> {
    if let Some(((namespace, local), _)) = attributes.iter().next() {
        return invalid(format!(
            "unsupported {context} attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn namespace_kind(value: &ResolveResult<'_>) -> Result<NamespaceKind> {
    Ok(match value {
        ResolveResult::Unbound => NamespaceKind::None,
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == STYLE_NS => NamespaceKind::Style,
        ResolveResult::Bound(Namespace(value)) if *value == PRESENTATION_NS => {
            NamespaceKind::Presentation
        },
        ResolveResult::Bound(Namespace(value)) if *value == SVG_NS => NamespaceKind::Svg,
        ResolveResult::Bound(_) => NamespaceKind::Other,
        ResolveResult::Unknown(prefix) => {
            return invalid(format!(
                "undeclared presentation page-layout prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ));
        },
    })
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "presentation-page-layout" && namespace != NamespaceKind::Style {
        return invalid("presentation-page-layout uses an invalid namespace");
    }
    if local == "placeholder" && namespace != NamespaceKind::Presentation {
        return invalid("presentation placeholder uses an invalid namespace");
    }
    Ok(())
}

#[derive(Clone, Debug)]
struct XmlSpan {
    start: usize,
    end: usize,
}

enum StylesSite {
    Content { insertion: usize },
    Empty { span: XmlSpan, qname: String },
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| make_error("invalid page-layout XML event boundary"))
}

fn mutation_sites(xml: &str, name: &str) -> Result<(Option<XmlSpan>, StylesSite)> {
    parse_presentation_page_layouts(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut styles_site = None;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                make_error(format!("invalid presentation page-layout XML: {error}"))
            })?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                let depth = stack.len() + 1;
                if namespace == NamespaceKind::Style
                    && local == "presentation-page-layout"
                    && matches!(stack.last(), Some(Frame { namespace: NamespaceKind::Office, local }) if local == "styles")
                    && parse_layout(&reader, element)?.name == name
                    && (target.is_some() || open_target.replace((depth, start)).is_some())
                {
                    return invalid("duplicate target presentation page layout");
                }
                stack.push(Frame { namespace, local });
            },
            Event::Empty(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                if namespace == NamespaceKind::Style
                    && local == "presentation-page-layout"
                    && matches!(stack.last(), Some(Frame { namespace: NamespaceKind::Office, local }) if local == "styles")
                    && parse_layout(&reader, element)?.name == name
                    && (target.replace(XmlSpan { start, end }).is_some() || open_target.is_some())
                {
                    return invalid("duplicate target presentation page layout");
                }
                if namespace == NamespaceKind::Office && local == "styles" {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Empty {
                        span: XmlSpan { start, end },
                        qname: decode_name(element.name().as_ref(), "qualified element")?,
                    });
                }
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let depth = stack.len();
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("presentation page-layout XML depth underflow"))?;
                if open_target.is_some_and(|(target_depth, _)| target_depth == depth) {
                    let (_, target_start) = open_target.take().expect("target depth checked");
                    if target
                        .replace(XmlSpan {
                            start: target_start,
                            end,
                        })
                        .is_some()
                    {
                        return invalid("duplicate target presentation page layout");
                    }
                }
                if frame.namespace == NamespaceKind::Office && frame.local == "styles" {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Content { insertion: start });
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in page layouts");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || open_target.is_some() {
        return invalid("unterminated presentation page-layout XML");
    }
    Ok((
        target,
        styles_site.ok_or_else(|| make_error("document has no office:styles element"))?,
    ))
}

/// Insert or replace one page-layout definition while preserving unrelated XML bytes.
pub fn set_presentation_page_layout_xml(
    xml: &str,
    layout: &PresentationPageLayout,
) -> Result<String> {
    layout.validate()?;
    let (target, styles_site) = mutation_sites(xml, &layout.name)?;
    let fragment = layout.to_xml_fragment()?;
    if let Some(span) = target {
        return Ok(format!(
            "{}{}{}",
            &xml[..span.start],
            fragment,
            &xml[span.end..]
        ));
    }
    Ok(match styles_site {
        StylesSite::Content { insertion } => {
            format!("{}{}{}", &xml[..insertion], fragment, &xml[insertion..])
        },
        StylesSite::Empty { span, qname } => {
            let raw = &xml[span.start..span.end];
            let slash = raw
                .rfind("/>")
                .ok_or_else(|| make_error("invalid empty office:styles element"))?;
            format!(
                "{}{}>{}</{}>{}",
                &xml[..span.start],
                &raw[..slash],
                fragment,
                qname,
                &xml[span.end..]
            )
        },
    })
}

/// Remove one page-layout definition while preserving unrelated XML bytes.
pub fn remove_presentation_page_layout_xml(xml: &str, name: &str) -> Result<String> {
    validate_ncname(name, "presentation page-layout name")?;
    let (target, _) = mutation_sites(xml, name)?;
    let Some(span) = target else {
        return Ok(xml.to_owned());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}

fn write_layout(output: &mut String, layout: &PresentationPageLayout, standalone: bool) {
    output.push_str("<style:presentation-page-layout");
    if standalone {
        output.push_str(
            r#" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0""#,
        );
    }
    push_attribute(output, "style:name", &layout.name);
    if let Some(value) = &layout.display_name {
        push_attribute(output, "style:display-name", value);
    }
    if layout.placeholders.is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    for placeholder in &layout.placeholders {
        output.push_str("<presentation:placeholder");
        push_attribute(output, "presentation:object", placeholder.class.as_str());
        push_attribute(output, "svg:x", &placeholder.x.to_string());
        push_attribute(output, "svg:y", &placeholder.y.to_string());
        push_attribute(output, "svg:width", &placeholder.width.to_string());
        push_attribute(output, "svg:height", &placeholder.height.to_string());
        output.push_str("/>");
    }
    output.push_str("</style:presentation-page-layout>");
}

fn push_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
    output.push('"');
}

fn validate_decimal(value: &str, complete: &str) -> Result<()> {
    let value = value.strip_prefix('-').unwrap_or(value);
    if value.is_empty() {
        return invalid(format!("invalid presentation measure '{complete}'"));
    }
    let mut parts = value.split('.');
    let integer = parts.next().expect("split always yields one value");
    let fraction = parts.next();
    if parts.next().is_some()
        || !integer.bytes().all(|value| value.is_ascii_digit())
        || fraction.is_some_and(|value| !value.bytes().all(|byte| byte.is_ascii_digit()))
        || integer.is_empty() && fraction.is_none_or(str::is_empty)
    {
        return invalid(format!("invalid presentation measure '{complete}'"));
    }
    Ok(())
}

fn ncname_start(character: char) -> bool {
    matches!(character,
        'A'..='Z' | '_' | 'a'..='z'
        | '\u{c0}'..='\u{d6}' | '\u{d8}'..='\u{f6}' | '\u{f8}'..='\u{2ff}'
        | '\u{370}'..='\u{37d}' | '\u{37f}'..='\u{1fff}' | '\u{200c}'..='\u{200d}'
        | '\u{2070}'..='\u{218f}' | '\u{2c00}'..='\u{2fef}' | '\u{3001}'..='\u{d7ff}'
        | '\u{f900}'..='\u{fdcf}' | '\u{fdf0}'..='\u{fffd}' | '\u{10000}'..='\u{effff}'
    )
}

fn ncname_continue(character: char) -> bool {
    ncname_start(character)
        || matches!(character, '-' | '.' | '0'..='9' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}')
}

fn validate_ncname(value: &str, context: &str) -> Result<()> {
    validate_text(value, context, false)?;
    let mut characters = value.chars();
    if !characters.next().is_some_and(ncname_start) || !characters.all(ncname_continue) {
        return invalid(format!("{context} is not an XML NCName"));
    }
    Ok(())
}

fn validate_text(value: &str, context: &str, empty_allowed: bool) -> Result<()> {
    if !empty_allowed && value.is_empty() {
        return invalid(format!("{context} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{context} exceeds 64 KiB"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{context} contains an XML-prohibited character"));
    }
    Ok(())
}

fn decode_name(value: &[u8], context: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| make_error(format!("invalid UTF-8 in page-layout {context} name")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const PREFIX: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:styles>"#;
    const SUFFIX: &str = "</office:styles></office:document-styles>";

    #[test]
    fn parses_and_round_trips_page_layouts() {
        let xml = format!(
            r#"{PREFIX}<style:presentation-page-layout style:name="TitleBody" style:display-name="Title &amp; body"><presentation:placeholder presentation:object="title" svg:x="5%" svg:y="1.25cm" svg:width="90%" svg:height="3cm"/><presentation:placeholder presentation:object="outline" svg:x="2cm" svg:y="-0.5cm" svg:width="20cm" svg:height="12cm"/></style:presentation-page-layout>{SUFFIX}"#
        );
        let layouts = parse_presentation_page_layouts(&xml).unwrap();
        assert_eq!(layouts.layouts.len(), 1);
        assert_eq!(
            layouts.layouts[0].display_name.as_deref(),
            Some("Title & body")
        );
        assert_eq!(
            layouts.layouts[0].placeholders[0].x.unit(),
            PresentationMeasureUnit::Percent
        );
        assert_eq!(layouts.layouts[0].placeholders[1].y.value(), -0.5);

        let serialized = layouts.to_xml().unwrap();
        assert_eq!(
            parse_presentation_page_layouts(&serialized).unwrap(),
            layouts
        );
    }

    #[test]
    fn rejects_malformed_layouts() {
        for body in [
            r#"<style:presentation-page-layout style:name="x"><presentation:placeholder presentation:object="title" svg:x="1cm" svg:y="2cm" svg:width="3cm"/></style:presentation-page-layout>"#,
            r#"<style:presentation-page-layout style:name="x"><presentation:placeholder presentation:object="unknown" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"/></style:presentation-page-layout>"#,
            r#"<style:presentation-page-layout style:name="x"><presentation:placeholder presentation:object="title" svg:x="1e2cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"/></style:presentation-page-layout>"#,
            r#"<style:presentation-page-layout style:name="x"><presentation:placeholder presentation:object="title" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"></presentation:placeholder></style:presentation-page-layout>"#,
            r#"<style:presentation-page-layout style:name="x"/><style:presentation-page-layout style:name="x"/>"#,
        ] {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                parse_presentation_page_layouts(&xml).is_err(),
                "accepted {body}"
            );
        }
        let misplaced = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:automatic-styles><style:presentation-page-layout style:name="x"/></office:automatic-styles></office:document-styles>"#.to_string();
        assert!(parse_presentation_page_layouts(&misplaced).is_err());
    }

    #[test]
    fn parses_libreoffice_page_layout_fixture() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/xmloff/qa/unit/data/theme.fodp");
        let Ok(xml) = std::fs::read_to_string(path) else {
            return;
        };
        let layouts = parse_presentation_page_layouts(&xml).unwrap();
        assert_eq!(layouts.layouts.len(), 2);
        assert_eq!(layouts.layouts[0].placeholders.len(), 6);
        assert_eq!(layouts.layouts[1].placeholders.len(), 2);
        assert_eq!(
            parse_presentation_page_layouts(&layouts.to_xml().unwrap())
                .unwrap()
                .layouts
                .len(),
            2
        );
    }

    #[test]
    fn exhausts_classes_and_geometry_lexicals() {
        let classes = [
            PresentationPlaceholderClass::Title,
            PresentationPlaceholderClass::Outline,
            PresentationPlaceholderClass::Subtitle,
            PresentationPlaceholderClass::Text,
            PresentationPlaceholderClass::Graphic,
            PresentationPlaceholderClass::Object,
            PresentationPlaceholderClass::Chart,
            PresentationPlaceholderClass::Table,
            PresentationPlaceholderClass::OrganizationChart,
            PresentationPlaceholderClass::Page,
            PresentationPlaceholderClass::Notes,
            PresentationPlaceholderClass::Handout,
            PresentationPlaceholderClass::Header,
            PresentationPlaceholderClass::Footer,
            PresentationPlaceholderClass::DateTime,
            PresentationPlaceholderClass::PageNumber,
        ];
        let mut layout = PresentationPageLayout::new("_all.classes").unwrap();
        for class in classes {
            layout.placeholders.push(PresentationPlaceholder::new(
                class,
                "-.5cm".parse().unwrap(),
                "1.cm".parse().unwrap(),
                "-0.25%".parse().unwrap(),
                "-2px".parse().unwrap(),
            ));
        }
        let parsed = parse_presentation_page_layouts(&format!(
            "{PREFIX}{}{SUFFIX}",
            layout.to_xml_fragment().unwrap()
        ))
        .unwrap();
        assert_eq!(parsed.layouts[0].placeholders.len(), 16);
        assert_eq!(
            parsed.layouts[0].placeholders[15].class,
            PresentationPlaceholderClass::PageNumber
        );
        for value in [".5cm", "1.cm", "-.5%", "-0px", "01.00pt"] {
            assert!(
                value.parse::<PresentationMeasure>().is_ok(),
                "rejected {value}"
            );
        }
        for value in [".", ".cm", "+1cm", "1e2cm", "1 cm", "NaNcm"] {
            assert!(
                value.parse::<PresentationMeasure>().is_err(),
                "accepted {value}"
            );
        }
    }

    #[test]
    fn rejects_identity_duplicates_and_caps() {
        for name in ["", "1layout", "bad:name", "two words"] {
            assert!(
                PresentationPageLayout::new(name).is_err(),
                "accepted {name}"
            );
        }
        let aliased_duplicate = format!(
            r#"{PREFIX}<style:presentation-page-layout xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="a" s:name="b"/>{SUFFIX}"#
        );
        assert!(parse_presentation_page_layouts(&aliased_duplicate).is_err());
        let mut capped = PresentationPageLayout::new("cap").unwrap();
        let placeholder = PresentationPlaceholder::new(
            PresentationPlaceholderClass::Text,
            "0cm".parse().unwrap(),
            "0cm".parse().unwrap(),
            "1cm".parse().unwrap(),
            "1cm".parse().unwrap(),
        );
        capped.placeholders = vec![placeholder; MAX_PLACEHOLDERS + 1];
        assert!(capped.validate().is_err());
        assert!(PresentationPageLayout::new("x".repeat(MAX_VALUE_BYTES + 1)).is_err());
    }

    #[test]
    fn losslessly_inserts_replaces_and_removes() {
        let original = format!(r#"{PREFIX}<!--keep--><style:style style:name="other"/>{SUFFIX}"#);
        let mut layout = PresentationPageLayout::new("layout1").unwrap();
        layout.display_name = Some("First".to_string());
        let inserted = set_presentation_page_layout_xml(&original, &layout).unwrap();
        assert!(inserted.contains(
            "<!--keep--><style:style style:name=\"other\"/><style:presentation-page-layout"
        ));
        layout.display_name = Some("Replacement".to_string());
        let replaced = set_presentation_page_layout_xml(&inserted, &layout).unwrap();
        assert!(replaced.contains("style:display-name=\"Replacement\""));
        assert!(!replaced.contains("style:display-name=\"First\""));
        assert_eq!(
            remove_presentation_page_layout_xml(&replaced, "layout1").unwrap(),
            original
        );
        assert_eq!(
            remove_presentation_page_layout_xml(&original, "missing").unwrap(),
            original
        );
    }

    #[test]
    fn builder_writes_page_layouts() {
        let mut builder = crate::PresentationBuilder::new();
        let mut layout = PresentationPageLayout::new("builder_layout").unwrap();
        layout.placeholders.push(PresentationPlaceholder::new(
            PresentationPlaceholderClass::Title,
            "1cm".parse().unwrap(),
            "2cm".parse().unwrap(),
            "20cm".parse().unwrap(),
            "3cm".parse().unwrap(),
        ));
        builder.add_page_layout(layout).unwrap();
        let presentation = crate::Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            presentation.page_layouts().unwrap().layouts[0].name,
            "builder_layout"
        );
    }
}
