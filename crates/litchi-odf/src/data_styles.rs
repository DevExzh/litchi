//! Typed ODF 1.2/1.3 data styles and number-format tokens.

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 2_000_000;
const MAX_STYLES: usize = 65_536;
const MAX_PARTS: usize = 4_096;
const MAX_MAPS: usize = 1_024;
const MAX_ATTRIBUTES: usize = 128;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 32 * 1_048_576;

/// XML part containing a data style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfDataStylePart {
    Content,
    Styles,
    Flat,
}

/// Direct style container containing a data style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfDataStyleSection {
    Styles,
    AutomaticStyles,
}

/// Core schema version used to validate or serialize a style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfDataStyleVersion {
    V1_2,
    V1_3,
}

/// One of the seven standard data-style containers.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfDataStyleKind {
    Number,
    Currency,
    Percentage,
    Date,
    Time,
    Boolean,
    Text,
}

impl OdfDataStyleKind {
    fn local(self) -> &'static str {
        match self {
            Self::Number => "number-style",
            Self::Currency => "currency-style",
            Self::Percentage => "percentage-style",
            Self::Date => "date-style",
            Self::Time => "time-style",
            Self::Boolean => "boolean-style",
            Self::Text => "text-style",
        }
    }

    fn parse(local: &str) -> Option<Self> {
        Some(match local {
            "number-style" => Self::Number,
            "currency-style" => Self::Currency,
            "percentage-style" => Self::Percentage,
            "date-style" => Self::Date,
            "time-style" => Self::Time,
            "boolean-style" => Self::Boolean,
            "text-style" => Self::Text,
            _ => return None,
        })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfShortLong {
    Short,
    Long,
}

impl OdfShortLong {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "short" => Ok(Self::Short),
            "long" => Ok(Self::Long),
            _ => invalid(format!("invalid number:style '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Short => "short",
            Self::Long => "long",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfTransliterationStyle {
    Short,
    Medium,
    Long,
}

impl OdfTransliterationStyle {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "short" => Ok(Self::Short),
            "medium" => Ok(Self::Medium),
            "long" => Ok(Self::Long),
            _ => invalid(format!("invalid number:transliteration-style '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Short => "short",
            Self::Medium => "medium",
            Self::Long => "long",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfFormatSource {
    Fixed,
    Language,
}

impl OdfFormatSource {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "fixed" => Ok(Self::Fixed),
            "language" => Ok(Self::Language),
            _ => invalid(format!("invalid number:format-source '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Fixed => "fixed",
            Self::Language => "language",
        }
    }
}

/// Locale metadata shared by a data style or currency symbol.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfNumberLocale {
    pub language: Option<String>,
    pub country: Option<String>,
    pub script: Option<String>,
    pub rfc_language_tag: Option<String>,
}

/// A validated, opaque `style:text-properties` fragment.
///
/// Its independent typed property API remains the authority for the 84
/// text-property attributes; the data-style parser preserves the exact fragment.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfDataStyleTextProperties {
    xml: String,
}

impl OdfDataStyleTextProperties {
    pub fn as_xml(&self) -> &str {
        &self.xml
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct OdfNumberToken {
    pub decimal_replacement: Option<String>,
    pub display_factor: Option<f64>,
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
    pub embedded_text: Vec<OdfEmbeddedText>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfEmbeddedText {
    pub position: i64,
    pub text: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfScientificNumberToken {
    pub min_exponent_digits: Option<i64>,
    pub exponent_interval: Option<u64>,
    pub forced_exponent_sign: Option<bool>,
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfFractionToken {
    pub min_numerator_digits: Option<i64>,
    pub min_denominator_digits: Option<i64>,
    pub denominator_value: Option<i64>,
    pub max_denominator_value: Option<u64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfCurrencySymbolToken {
    pub locale: OdfNumberLocale,
    pub text: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfCalendarToken {
    pub style: Option<OdfShortLong>,
    pub calendar: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfMonthToken {
    pub style: Option<OdfShortLong>,
    pub textual: Option<bool>,
    pub possessive_form: Option<bool>,
    pub calendar: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfWeekOfYearToken {
    pub calendar: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfClockToken {
    pub style: Option<OdfShortLong>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfSecondsToken {
    pub style: Option<OdfShortLong>,
    pub decimal_places: Option<i64>,
}

/// One ordered formatting component within a data style.
#[derive(Clone, Debug, PartialEq)]
pub enum OdfDataStylePartToken {
    Text(String),
    FillCharacter(String),
    Number(OdfNumberToken),
    ScientificNumber(OdfScientificNumberToken),
    Fraction(OdfFractionToken),
    CurrencySymbol(OdfCurrencySymbolToken),
    Day(OdfCalendarToken),
    Month(OdfMonthToken),
    Year(OdfCalendarToken),
    Era(OdfCalendarToken),
    DayOfWeek(OdfCalendarToken),
    WeekOfYear(OdfWeekOfYearToken),
    Quarter(OdfCalendarToken),
    Hours(OdfClockToken),
    Minutes(OdfClockToken),
    Seconds(OdfSecondsToken),
    AmPm,
    Boolean,
    TextContent,
}

/// A trailing conditional style map. Conditions remain opaque strings.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfDataStyleMap {
    pub condition: String,
    pub apply_style_name: String,
    pub base_cell_address: Option<String>,
}

/// One complete standard data style.
#[derive(Clone, Debug, PartialEq)]
pub struct OdfDataStyle {
    pub source_part: OdfDataStylePart,
    pub section: OdfDataStyleSection,
    pub source_version: OdfDataStyleVersion,
    pub kind: OdfDataStyleKind,
    pub name: String,
    pub display_name: Option<String>,
    pub locale: OdfNumberLocale,
    pub title: Option<String>,
    pub volatile: Option<bool>,
    pub transliteration_format: Option<String>,
    pub transliteration_language: Option<String>,
    pub transliteration_country: Option<String>,
    pub transliteration_style: Option<OdfTransliterationStyle>,
    pub automatic_order: Option<bool>,
    pub format_source: Option<OdfFormatSource>,
    pub truncate_on_overflow: Option<bool>,
    pub text_properties: Option<OdfDataStyleTextProperties>,
    pub parts: Vec<OdfDataStylePartToken>,
    pub maps: Vec<OdfDataStyleMap>,
}

impl OdfDataStyle {
    pub fn new(
        name: impl Into<String>,
        kind: OdfDataStyleKind,
        section: OdfDataStyleSection,
    ) -> Result<Self> {
        let value = Self {
            source_part: OdfDataStylePart::Flat,
            section,
            source_version: OdfDataStyleVersion::V1_3,
            kind,
            name: name.into(),
            display_name: None,
            locale: OdfNumberLocale::default(),
            title: None,
            volatile: None,
            transliteration_format: None,
            transliteration_language: None,
            transliteration_country: None,
            transliteration_style: None,
            automatic_order: None,
            format_source: None,
            truncate_on_overflow: None,
            text_properties: None,
            parts: Vec::new(),
            maps: Vec::new(),
        };
        value.validate(OdfDataStyleVersion::V1_3)?;
        Ok(value)
    }

    pub fn validate(&self, version: OdfDataStyleVersion) -> Result<()> {
        self.validate_inner(version, false)
    }

    fn validate_inner(&self, version: OdfDataStyleVersion, allow_lo_aliases: bool) -> Result<()> {
        validate_name(&self.name, "style:name")?;
        validate_optional_string(self.display_name.as_deref(), "style:display-name")?;
        validate_locale(&self.locale)?;
        for (value, name) in [
            (self.title.as_deref(), "number:title"),
            (
                self.transliteration_format.as_deref(),
                "number:transliteration-format",
            ),
            (
                self.transliteration_language.as_deref(),
                "number:transliteration-language",
            ),
            (
                self.transliteration_country.as_deref(),
                "number:transliteration-country",
            ),
        ] {
            validate_optional_string(value, name)?;
        }
        if self.parts.len() > MAX_PARTS || self.maps.len() > MAX_MAPS {
            return invalid("data style exceeds component limits");
        }
        if self.automatic_order.is_some()
            && !matches!(self.kind, OdfDataStyleKind::Currency | OdfDataStyleKind::Date)
        {
            return invalid("number:automatic-order is invalid for this data style");
        }
        if self.format_source.is_some()
            && !matches!(self.kind, OdfDataStyleKind::Date | OdfDataStyleKind::Time)
        {
            return invalid("number:format-source is invalid for this data style");
        }
        if self.truncate_on_overflow.is_some() && self.kind != OdfDataStyleKind::Time {
            return invalid("number:truncate-on-overflow is valid only for time styles");
        }
        validate_sequence(self.kind, &self.parts, version, allow_lo_aliases)?;
        for part in &self.parts {
            validate_part(part, version, allow_lo_aliases)?;
        }
        for map in &self.maps {
            validate_text(&map.condition, "style:condition")?;
            validate_name(&map.apply_style_name, "style:apply-style-name")?;
            if let Some(address) = &map.base_cell_address {
                validate_cell_address(address)?;
            }
        }
        Ok(())
    }

    /// Serialize a normative ODF fragment for the requested core version.
    pub fn to_xml_fragment(&self, version: OdfDataStyleVersion) -> Result<String> {
        self.validate(version)?;
        let mut out = format!("<number:{} style:name=\"{}\"", self.kind.local(), esc(&self.name));
        attr(&mut out, "style:display-name", self.display_name.as_deref());
        locale_attrs(&mut out, &self.locale);
        attr(&mut out, "number:title", self.title.as_deref());
        bool_attr(&mut out, "style:volatile", self.volatile);
        attr(
            &mut out,
            "number:transliteration-format",
            self.transliteration_format.as_deref(),
        );
        attr(
            &mut out,
            "number:transliteration-language",
            self.transliteration_language.as_deref(),
        );
        attr(
            &mut out,
            "number:transliteration-country",
            self.transliteration_country.as_deref(),
        );
        if let Some(value) = self.transliteration_style {
            attr(
                &mut out,
                "number:transliteration-style",
                Some(value.as_str()),
            );
        }
        bool_attr(&mut out, "number:automatic-order", self.automatic_order);
        if let Some(value) = self.format_source {
            attr(&mut out, "number:format-source", Some(value.as_str()));
        }
        bool_attr(
            &mut out,
            "number:truncate-on-overflow",
            self.truncate_on_overflow,
        );
        out.push('>');
        if let Some(properties) = &self.text_properties {
            out.push_str(properties.as_xml());
        }
        for part in &self.parts {
            write_part(&mut out, part, version)?;
        }
        for map in &self.maps {
            out.push_str("<style:map");
            attr(&mut out, "style:condition", Some(&map.condition));
            attr(
                &mut out,
                "style:apply-style-name",
                Some(&map.apply_style_name),
            );
            attr(
                &mut out,
                "style:base-cell-address",
                map.base_cell_address.as_deref(),
            );
            out.push_str("/>");
        }
        out.push_str("</number:");
        out.push_str(self.kind.local());
        out.push('>');
        Ok(out)
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct OdfDataStyles {
    pub styles: Vec<OdfDataStyle>,
}

impl OdfDataStyles {
    pub fn get(
        &self,
        part: OdfDataStylePart,
        section: OdfDataStyleSection,
        name: &str,
    ) -> Option<&OdfDataStyle> {
        self.styles.iter().find(|style| {
            style.source_part == part && style.section == section && style.name == name
        })
    }

    pub fn in_section(
        &self,
        section: OdfDataStyleSection,
    ) -> impl Iterator<Item = &OdfDataStyle> {
        self.styles.iter().filter(move |style| style.section == section)
    }

    pub fn named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a OdfDataStyle> {
        self.styles.iter().filter(move |style| style.name == name)
    }

    fn append(&mut self, mut other: Self) -> Result<()> {
        for style in other.styles.drain(..) {
            if self.get(style.source_part, style.section, &style.name).is_some() {
                return invalid("duplicate data style identity");
            }
            if self.styles.len() >= MAX_STYLES {
                return invalid("too many data styles");
            }
            self.styles.push(style);
        }
        Ok(())
    }
}

#[derive(Clone)]
struct Attribute {
    namespace: Option<String>,
    local: String,
    value: String,
}

struct NodeBuilder {
    namespace: Option<String>,
    local: String,
    attributes: Vec<Attribute>,
    text: String,
    children: Vec<Node>,
    start: usize,
}

struct Node {
    namespace: Option<String>,
    local: String,
    attributes: Vec<Attribute>,
    text: String,
    children: Vec<Node>,
    raw: String,
}

#[derive(Clone)]
struct Frame {
    namespace: Option<String>,
    local: String,
}

/// Parse direct data styles from both standard style containers in one XML part.
pub fn parse_data_styles_xml(xml: &str, part: OdfDataStylePart) -> Result<OdfDataStyles> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("data-style XML exceeds 64 MiB");
    }
    if !xml.contains("-style") {
        return Ok(OdfDataStyles::default());
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut frames = Vec::<Frame>::new();
    let mut nodes = Vec::<NodeBuilder>::new();
    let mut output = OdfDataStyles::default();
    let mut version = OdfDataStyleVersion::V1_2;
    let mut xml_version = XmlVersion::Implicit1_0;
    let mut aggregate = 0usize;
    let mut events = 0usize;

    loop {
        events += 1;
        if events > MAX_EVENTS {
            return invalid("data-style XML has too many events");
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid data-style XML: {error}")))?;
        match event {
            Event::Decl(ref declaration) => {
                xml_version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            }
            Event::Start(ref element) => {
                if frames.len() >= MAX_DEPTH {
                    return invalid("data-style XML is too deep");
                }
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                if frames.is_empty() {
                    version = read_document_version(&reader, element, xml_version)?;
                }
                let direct = direct_style_section(&frames);
                reject_spoofed_container(namespace.as_deref(), &local, direct.is_some())?;
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                if !nodes.is_empty()
                    || direct.is_some()
                        && namespace.as_deref() == Some(NUMBER)
                        && OdfDataStyleKind::parse(&local).is_some()
                {
                    let attributes = collect_attributes(
                        &reader,
                        element,
                        xml_version,
                        &mut aggregate,
                    )?;
                    nodes.push(NodeBuilder {
                        namespace: namespace.clone(),
                        local: local.clone(),
                        attributes,
                        text: String::new(),
                        children: Vec::new(),
                        start,
                    });
                }
                frames.push(Frame { namespace, local });
            }
            Event::Empty(ref element) => {
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                let direct = direct_style_section(&frames);
                reject_spoofed_container(namespace.as_deref(), &local, direct.is_some())?;
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                if !nodes.is_empty() {
                    let node = Node {
                        namespace,
                        local,
                        attributes: collect_attributes(
                            &reader,
                            element,
                            xml_version,
                            &mut aggregate,
                        )?,
                        text: String::new(),
                        children: Vec::new(),
                        raw: xml[start..end].to_string(),
                    };
                    nodes.last_mut().expect("active data style").children.push(node);
                } else if let Some(section) = direct
                    && namespace.as_deref() == Some(NUMBER)
                    && OdfDataStyleKind::parse(&local).is_some()
                {
                    let node = Node {
                        namespace,
                        local,
                        attributes: collect_attributes(
                            &reader,
                            element,
                            xml_version,
                            &mut aggregate,
                        )?,
                        text: String::new(),
                        children: Vec::new(),
                        raw: xml[start..end].to_string(),
                    };
                    push_style(
                        &mut output,
                        parse_style_node(node, part, section, version)?,
                    )?;
                }
            }
            Event::Text(ref text) if !nodes.is_empty() => {
                let decoded = text
                    .decode()
                    .map_err(|error| bad(format!("invalid data-style text: {error}")))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| bad(format!("invalid data-style entity: {error}")))?;
                add_text(&mut nodes.last_mut().expect("active node").text, &unescaped, &mut aggregate)?;
            }
            Event::CData(ref text) if !nodes.is_empty() => {
                let decoded = reader
                    .decoder()
                    .decode(text.as_ref())
                    .map_err(|error| bad(format!("invalid data-style CDATA: {error}")))?;
                add_text(&mut nodes.last_mut().expect("active node").text, &decoded, &mut aggregate)?;
            }
            Event::GeneralRef(_) if !nodes.is_empty() => {
                return invalid("entity references are prohibited in data styles");
            }
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                if !nodes.is_empty() && nodes.len() == frames.len() - active_base_depth(&frames, &nodes) {
                    let builder = nodes.pop().expect("active data-style node");
                    let node = Node {
                        namespace: builder.namespace,
                        local: builder.local,
                        attributes: builder.attributes,
                        text: builder.text,
                        children: builder.children,
                        raw: xml[builder.start..end].to_string(),
                    };
                    if let Some(parent) = nodes.last_mut() {
                        parent.children.push(node);
                    } else {
                        let section = direct_parent_section(&frames)
                            .ok_or_else(|| bad("misplaced data style"))?;
                        push_style(
                            &mut output,
                            parse_style_node(node, part, section, version)?,
                        )?;
                    }
                }
                frames
                    .pop()
                    .ok_or_else(|| bad("data-style element stack underflow"))?;
            }
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited");
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !frames.is_empty() || !nodes.is_empty() {
        return invalid("truncated data-style XML");
    }
    Ok(output)
}

// The active node stack is always a suffix of the document frame stack. This
// helper avoids storing a second absolute depth in every node.
fn active_base_depth(frames: &[Frame], nodes: &[NodeBuilder]) -> usize {
    frames.len().saturating_sub(nodes.len())
}

fn direct_parent_section(frames: &[Frame]) -> Option<OdfDataStyleSection> {
    frames.get(frames.len().checked_sub(2)?).and_then(frame_section)
}

fn direct_style_section(frames: &[Frame]) -> Option<OdfDataStyleSection> {
    frames.last().and_then(frame_section)
}

fn frame_section(frame: &Frame) -> Option<OdfDataStyleSection> {
    if frame.namespace.as_deref() != Some(OFFICE) {
        return None;
    }
    match frame.local.as_str() {
        "styles" => Some(OdfDataStyleSection::Styles),
        "automatic-styles" => Some(OdfDataStyleSection::AutomaticStyles),
        _ => None,
    }
}

fn parse_style_node(
    mut node: Node,
    part: OdfDataStylePart,
    section: OdfDataStyleSection,
    version: OdfDataStyleVersion,
) -> Result<OdfDataStyle> {
    if node.namespace.as_deref() != Some(NUMBER) {
        return invalid("data style uses the wrong namespace");
    }
    let kind = OdfDataStyleKind::parse(&node.local)
        .ok_or_else(|| bad("unknown data-style container"))?;
    ensure_whitespace(&node.text, "data-style container")?;
    let name = required(&mut node.attributes, STYLE, "name")?;
    let display_name = take(&mut node.attributes, STYLE, "display-name");
    let locale = parse_locale(&mut node.attributes)?;
    let title = take(&mut node.attributes, NUMBER, "title");
    let volatile = take_bool(&mut node.attributes, STYLE, "volatile")?;
    let transliteration_format = take(
        &mut node.attributes,
        NUMBER,
        "transliteration-format",
    );
    let transliteration_language = take(
        &mut node.attributes,
        NUMBER,
        "transliteration-language",
    );
    let transliteration_country = take(
        &mut node.attributes,
        NUMBER,
        "transliteration-country",
    );
    let transliteration_style = take(
        &mut node.attributes,
        NUMBER,
        "transliteration-style",
    )
    .map(|value| OdfTransliterationStyle::parse(&value))
    .transpose()?;
    let automatic_order = take_bool(&mut node.attributes, NUMBER, "automatic-order")?;
    let format_source = take(&mut node.attributes, NUMBER, "format-source")
        .map(|value| OdfFormatSource::parse(&value))
        .transpose()?;
    let truncate_on_overflow = take_bool(
        &mut node.attributes,
        NUMBER,
        "truncate-on-overflow",
    )?;
    reject_remaining(&node.attributes, "data-style container")?;

    let mut text_properties = None;
    let mut parts = Vec::new();
    let mut maps = Vec::new();
    let mut compatibility_alias = false;
    let mut phase = 0u8;
    for child in node.children {
        if child.namespace.as_deref() == Some(STYLE) && child.local == "text-properties" {
            if phase != 0 || text_properties.is_some() {
                return invalid("style:text-properties must be the first and only such child");
            }
            ensure_whitespace(&child.text, "style:text-properties")?;
            if !child.children.is_empty() {
                return invalid("style:text-properties cannot contain elements");
            }
            text_properties = Some(OdfDataStyleTextProperties { xml: child.raw });
            continue;
        }
        phase = phase.max(1);
        if child.namespace.as_deref() == Some(STYLE) && child.local == "map" {
            phase = 2;
            if maps.len() >= MAX_MAPS {
                return invalid("too many style:map elements");
            }
            maps.push(parse_map(child)?);
        } else {
            if phase == 2 {
                return invalid("formatting tokens cannot follow style:map");
            }
            if parts.len() >= MAX_PARTS {
                return invalid("too many data-style tokens");
            }
            let (token, used_alias) = parse_part_node(child, version)?;
            compatibility_alias |= used_alias;
            parts.push(token);
        }
    }
    let style = OdfDataStyle {
        source_part: part,
        section,
        source_version: version,
        kind,
        name,
        display_name,
        locale,
        title,
        volatile,
        transliteration_format,
        transliteration_language,
        transliteration_country,
        transliteration_style,
        automatic_order,
        format_source,
        truncate_on_overflow,
        text_properties,
        parts,
        maps,
    };
    style.validate_inner(version, compatibility_alias)?;
    Ok(style)
}

fn parse_part_node(
    mut node: Node,
    version: OdfDataStyleVersion,
) -> Result<(OdfDataStylePartToken, bool)> {
    let standard = node.namespace.as_deref() == Some(NUMBER);
    let lo_fill = node.namespace.as_deref() == Some(LOEXT) && node.local == "fill-character";
    if !standard && !lo_fill {
        return invalid(format!(
            "unexpected data-style child {}:{}",
            node.namespace.as_deref().unwrap_or(""),
            node.local
        ));
    }
    if lo_fill {
        reject_remaining(&node.attributes, "loext:fill-character")?;
        ensure_no_children(&node, "loext:fill-character")?;
        validate_text(&node.text, "loext:fill-character")?;
        return Ok((OdfDataStylePartToken::FillCharacter(node.text), true));
    }
    let mut alias = false;
    let token = match node.local.as_str() {
        "text" => {
            reject_remaining(&node.attributes, "number:text")?;
            ensure_no_children(&node, "number:text")?;
            OdfDataStylePartToken::Text(node.text)
        }
        "fill-character" => {
            if version == OdfDataStyleVersion::V1_2 {
                return invalid("number:fill-character requires ODF 1.3");
            }
            reject_remaining(&node.attributes, "number:fill-character")?;
            ensure_no_children(&node, "number:fill-character")?;
            OdfDataStylePartToken::FillCharacter(node.text)
        }
        "number" => {
            ensure_whitespace(&node.text, "number:number")?;
            let mut embedded_text = Vec::new();
            for mut child in node.children.drain(..) {
                if child.namespace.as_deref() != Some(NUMBER) || child.local != "embedded-text" {
                    return invalid("number:number may contain only number:embedded-text");
                }
                ensure_no_children(&child, "number:embedded-text")?;
                let position = required_i64(&mut child.attributes, NUMBER, "position")?;
                reject_remaining(&child.attributes, "number:embedded-text")?;
                embedded_text.push(OdfEmbeddedText {
                    position,
                    text: child.text,
                });
            }
            let min_decimal_places = take_versioned_i64(
                &mut node.attributes,
                "min-decimal-places",
                version,
                &mut alias,
            )?;
            let value = OdfNumberToken {
                decimal_replacement: take(
                    &mut node.attributes,
                    NUMBER,
                    "decimal-replacement",
                ),
                display_factor: take_f64(&mut node.attributes, NUMBER, "display-factor")?,
                decimal_places: take_i64(&mut node.attributes, NUMBER, "decimal-places")?,
                min_decimal_places,
                min_integer_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-integer-digits",
                )?,
                grouping: take_bool(&mut node.attributes, NUMBER, "grouping")?,
                embedded_text,
            };
            reject_remaining(&node.attributes, "number:number")?;
            OdfDataStylePartToken::Number(value)
        }
        "scientific-number" => {
            ensure_empty_node(&node, "number:scientific-number")?;
            let exponent_interval = take_versioned_u64(
                &mut node.attributes,
                "exponent-interval",
                version,
                &mut alias,
            )?;
            let forced_exponent_sign = take_versioned_bool(
                &mut node.attributes,
                "forced-exponent-sign",
                version,
                &mut alias,
            )?;
            let min_decimal_places = take_versioned_i64(
                &mut node.attributes,
                "min-decimal-places",
                version,
                &mut alias,
            )?;
            let value = OdfScientificNumberToken {
                min_exponent_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-exponent-digits",
                )?,
                exponent_interval,
                forced_exponent_sign,
                decimal_places: take_i64(&mut node.attributes, NUMBER, "decimal-places")?,
                min_decimal_places,
                min_integer_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-integer-digits",
                )?,
                grouping: take_bool(&mut node.attributes, NUMBER, "grouping")?,
            };
            reject_remaining(&node.attributes, "number:scientific-number")?;
            OdfDataStylePartToken::ScientificNumber(value)
        }
        "fraction" => {
            ensure_empty_node(&node, "number:fraction")?;
            let max_denominator_value = take_versioned_u64(
                &mut node.attributes,
                "max-denominator-value",
                version,
                &mut alias,
            )?;
            let value = OdfFractionToken {
                min_numerator_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-numerator-digits",
                )?,
                min_denominator_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-denominator-digits",
                )?,
                denominator_value: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "denominator-value",
                )?,
                max_denominator_value,
                min_integer_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-integer-digits",
                )?,
                grouping: take_bool(&mut node.attributes, NUMBER, "grouping")?,
            };
            reject_remaining(&node.attributes, "number:fraction")?;
            OdfDataStylePartToken::Fraction(value)
        }
        "currency-symbol" => {
            ensure_no_children(&node, "number:currency-symbol")?;
            let locale = parse_locale(&mut node.attributes)?;
            reject_remaining(&node.attributes, "number:currency-symbol")?;
            OdfDataStylePartToken::CurrencySymbol(OdfCurrencySymbolToken {
                locale,
                text: node.text,
            })
        }
        "day" => OdfDataStylePartToken::Day(parse_calendar(&mut node)?),
        "month" => {
            ensure_empty_node(&node, "number:month")?;
            let value = OdfMonthToken {
                style: take_style(&mut node.attributes)?,
                textual: take_bool(&mut node.attributes, NUMBER, "textual")?,
                possessive_form: take_bool(
                    &mut node.attributes,
                    NUMBER,
                    "possessive-form",
                )?,
                calendar: take(&mut node.attributes, NUMBER, "calendar"),
            };
            reject_remaining(&node.attributes, "number:month")?;
            OdfDataStylePartToken::Month(value)
        }
        "year" => OdfDataStylePartToken::Year(parse_calendar(&mut node)?),
        "era" => OdfDataStylePartToken::Era(parse_calendar(&mut node)?),
        "day-of-week" => OdfDataStylePartToken::DayOfWeek(parse_calendar(&mut node)?),
        "week-of-year" => {
            ensure_empty_node(&node, "number:week-of-year")?;
            let value = OdfWeekOfYearToken {
                calendar: take(&mut node.attributes, NUMBER, "calendar"),
            };
            reject_remaining(&node.attributes, "number:week-of-year")?;
            OdfDataStylePartToken::WeekOfYear(value)
        }
        "quarter" => OdfDataStylePartToken::Quarter(parse_calendar(&mut node)?),
        "hours" => OdfDataStylePartToken::Hours(parse_clock(&mut node)?),
        "minutes" => OdfDataStylePartToken::Minutes(parse_clock(&mut node)?),
        "seconds" => {
            ensure_empty_node(&node, "number:seconds")?;
            let value = OdfSecondsToken {
                style: take_style(&mut node.attributes)?,
                decimal_places: take_i64(&mut node.attributes, NUMBER, "decimal-places")?,
            };
            reject_remaining(&node.attributes, "number:seconds")?;
            OdfDataStylePartToken::Seconds(value)
        }
        "am-pm" => {
            ensure_empty_node(&node, "number:am-pm")?;
            reject_remaining(&node.attributes, "number:am-pm")?;
            OdfDataStylePartToken::AmPm
        }
        "boolean" => {
            ensure_empty_node(&node, "number:boolean")?;
            reject_remaining(&node.attributes, "number:boolean")?;
            OdfDataStylePartToken::Boolean
        }
        "text-content" => {
            ensure_empty_node(&node, "number:text-content")?;
            reject_remaining(&node.attributes, "number:text-content")?;
            OdfDataStylePartToken::TextContent
        }
        _ => return invalid(format!("unexpected number:{} token", node.local)),
    };
    Ok((token, alias))
}

fn parse_map(mut node: Node) -> Result<OdfDataStyleMap> {
    ensure_empty_node(&node, "style:map")?;
    let value = OdfDataStyleMap {
        condition: required(&mut node.attributes, STYLE, "condition")?,
        apply_style_name: required(
            &mut node.attributes,
            STYLE,
            "apply-style-name",
        )?,
        base_cell_address: take(&mut node.attributes, STYLE, "base-cell-address"),
    };
    reject_remaining(&node.attributes, "style:map")?;
    Ok(value)
}

fn parse_calendar(node: &mut Node) -> Result<OdfCalendarToken> {
    ensure_empty_node(node, "calendar token")?;
    let value = OdfCalendarToken {
        style: take_style(&mut node.attributes)?,
        calendar: take(&mut node.attributes, NUMBER, "calendar"),
    };
    reject_remaining(&node.attributes, "calendar token")?;
    Ok(value)
}

fn parse_clock(node: &mut Node) -> Result<OdfClockToken> {
    ensure_empty_node(node, "clock token")?;
    let value = OdfClockToken {
        style: take_style(&mut node.attributes)?,
    };
    reject_remaining(&node.attributes, "clock token")?;
    Ok(value)
}

fn take_style(attributes: &mut Vec<Attribute>) -> Result<Option<OdfShortLong>> {
    take(attributes, NUMBER, "style")
        .map(|value| OdfShortLong::parse(&value))
        .transpose()
}

fn validate_sequence(
    kind: OdfDataStyleKind,
    parts: &[OdfDataStylePartToken],
    version: OdfDataStyleVersion,
    allow_lo: bool,
) -> Result<()> {
    let allow_fill = version == OdfDataStyleVersion::V1_3 || allow_lo;
    let mut index = 0usize;
    match kind {
        OdfDataStyleKind::Boolean => {
            consume_plain_text(parts, &mut index);
            if matches!(parts.get(index), Some(OdfDataStylePartToken::Boolean)) {
                index += 1;
                consume_plain_text(parts, &mut index);
            }
        }
        OdfDataStyleKind::Number => {
            consume_separator(parts, &mut index, allow_fill);
            if matches!(
                parts.get(index),
                Some(
                    OdfDataStylePartToken::Number(_)
                        | OdfDataStylePartToken::ScientificNumber(_)
                        | OdfDataStylePartToken::Fraction(_)
                )
            ) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
        }
        OdfDataStyleKind::Percentage => {
            consume_separator(parts, &mut index, allow_fill);
            if matches!(parts.get(index), Some(OdfDataStylePartToken::Number(_))) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
        }
        OdfDataStyleKind::Currency => {
            consume_separator(parts, &mut index, allow_fill);
            if matches!(parts.get(index), Some(OdfDataStylePartToken::Number(_))) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
                if matches!(
                    parts.get(index),
                    Some(OdfDataStylePartToken::CurrencySymbol(_))
                ) {
                    index += 1;
                    consume_separator(parts, &mut index, allow_fill);
                }
            } else if matches!(
                parts.get(index),
                Some(OdfDataStylePartToken::CurrencySymbol(_))
            ) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
                if matches!(parts.get(index), Some(OdfDataStylePartToken::Number(_))) {
                    index += 1;
                    consume_separator(parts, &mut index, allow_fill);
                }
            }
        }
        OdfDataStyleKind::Date => {
            consume_separator(parts, &mut index, allow_fill);
            let mut count = 0usize;
            while parts.get(index).is_some_and(is_date_token) {
                count += 1;
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
            if count == 0 {
                return invalid("number:date-style requires at least one date token");
            }
        }
        OdfDataStyleKind::Time => {
            consume_separator(parts, &mut index, allow_fill);
            let mut count = 0usize;
            while parts.get(index).is_some_and(is_time_token) {
                count += 1;
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
            if count == 0 {
                return invalid("number:time-style requires at least one time token");
            }
        }
        OdfDataStyleKind::Text => {
            consume_separator(parts, &mut index, allow_fill);
            while matches!(parts.get(index), Some(OdfDataStylePartToken::TextContent)) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
        }
    }
    if index != parts.len() {
        return invalid(format!(
            "invalid ordered token sequence for number:{}",
            kind.local()
        ));
    }
    Ok(())
}

fn consume_plain_text(parts: &[OdfDataStylePartToken], index: &mut usize) {
    if matches!(parts.get(*index), Some(OdfDataStylePartToken::Text(_))) {
        *index += 1;
    }
}

fn consume_separator(parts: &[OdfDataStylePartToken], index: &mut usize, allow_fill: bool) {
    consume_plain_text(parts, index);
    if allow_fill
        && matches!(
            parts.get(*index),
            Some(OdfDataStylePartToken::FillCharacter(_))
        )
    {
        *index += 1;
        consume_plain_text(parts, index);
    }
}

fn is_date_token(part: &OdfDataStylePartToken) -> bool {
    matches!(
        part,
        OdfDataStylePartToken::Day(_)
            | OdfDataStylePartToken::Month(_)
            | OdfDataStylePartToken::Year(_)
            | OdfDataStylePartToken::Era(_)
            | OdfDataStylePartToken::DayOfWeek(_)
            | OdfDataStylePartToken::WeekOfYear(_)
            | OdfDataStylePartToken::Quarter(_)
            | OdfDataStylePartToken::Hours(_)
            | OdfDataStylePartToken::Minutes(_)
            | OdfDataStylePartToken::Seconds(_)
            | OdfDataStylePartToken::AmPm
    )
}

fn is_time_token(part: &OdfDataStylePartToken) -> bool {
    matches!(
        part,
        OdfDataStylePartToken::Hours(_)
            | OdfDataStylePartToken::Minutes(_)
            | OdfDataStylePartToken::Seconds(_)
            | OdfDataStylePartToken::AmPm
    )
}

fn validate_part(
    part: &OdfDataStylePartToken,
    version: OdfDataStyleVersion,
    allow_lo: bool,
) -> Result<()> {
    match part {
        OdfDataStylePartToken::Text(value) => validate_text(value, "number:text")?,
        OdfDataStylePartToken::FillCharacter(value) => {
            validate_text(value, "number:fill-character")?;
            require_1_3(true, version, allow_lo)?;
        }
        OdfDataStylePartToken::Number(value) => {
            validate_optional_string(
                value.decimal_replacement.as_deref(),
                "number:decimal-replacement",
            )?;
            if value.embedded_text.len() > MAX_PARTS {
                return invalid("too many number:embedded-text elements");
            }
            for embedded in &value.embedded_text {
                validate_text(&embedded.text, "number:embedded-text")?;
            }
            require_1_3(value.min_decimal_places.is_some(), version, allow_lo)?;
        }
        OdfDataStylePartToken::ScientificNumber(value) => {
            require_1_3(
                value.min_decimal_places.is_some()
                    || value.exponent_interval.is_some()
                    || value.forced_exponent_sign.is_some(),
                version,
                allow_lo,
            )?;
            if value.exponent_interval == Some(0) {
                return invalid("number:exponent-interval must be positive");
            }
        }
        OdfDataStylePartToken::Fraction(value) => {
            require_1_3(value.max_denominator_value.is_some(), version, allow_lo)?;
            if value.max_denominator_value == Some(0) {
                return invalid("number:max-denominator-value must be positive");
            }
        }
        OdfDataStylePartToken::CurrencySymbol(value) => {
            validate_locale(&value.locale)?;
            validate_text(&value.text, "number:currency-symbol")?;
        }
        OdfDataStylePartToken::Day(value)
        | OdfDataStylePartToken::Year(value)
        | OdfDataStylePartToken::Era(value)
        | OdfDataStylePartToken::DayOfWeek(value)
        | OdfDataStylePartToken::Quarter(value) => {
            validate_optional_string(value.calendar.as_deref(), "number:calendar")?;
        }
        OdfDataStylePartToken::Month(value) => {
            validate_optional_string(value.calendar.as_deref(), "number:calendar")?;
        }
        OdfDataStylePartToken::WeekOfYear(value) => {
            validate_optional_string(value.calendar.as_deref(), "number:calendar")?;
        }
        _ => {}
    }
    Ok(())
}

fn require_1_3(
    present: bool,
    version: OdfDataStyleVersion,
    allow_lo: bool,
) -> Result<()> {
    if present && version == OdfDataStyleVersion::V1_2 && !allow_lo {
        return invalid("ODF 1.3 number-format feature used in ODF 1.2");
    }
    Ok(())
}

fn write_part(
    out: &mut String,
    part: &OdfDataStylePartToken,
    version: OdfDataStyleVersion,
) -> Result<()> {
    match part {
        OdfDataStylePartToken::Text(value) => element_text(out, "number:text", value),
        OdfDataStylePartToken::FillCharacter(value) => {
            if version != OdfDataStyleVersion::V1_3 {
                return invalid("number:fill-character requires ODF 1.3 output");
            }
            element_text(out, "number:fill-character", value);
        }
        OdfDataStylePartToken::Number(value) => {
            out.push_str("<number:number");
            attr(
                out,
                "number:decimal-replacement",
                value.decimal_replacement.as_deref(),
            );
            f64_attr(out, "number:display-factor", value.display_factor);
            i64_attr(out, "number:decimal-places", value.decimal_places);
            i64_attr(
                out,
                "number:min-decimal-places",
                value.min_decimal_places,
            );
            i64_attr(
                out,
                "number:min-integer-digits",
                value.min_integer_digits,
            );
            bool_attr(out, "number:grouping", value.grouping);
            if value.embedded_text.is_empty() {
                out.push_str("/>");
            } else {
                out.push('>');
                for embedded in &value.embedded_text {
                    out.push_str("<number:embedded-text");
                    i64_attr(out, "number:position", Some(embedded.position));
                    out.push('>');
                    out.push_str(&esc(&embedded.text));
                    out.push_str("</number:embedded-text>");
                }
                out.push_str("</number:number>");
            }
        }
        OdfDataStylePartToken::ScientificNumber(value) => {
            out.push_str("<number:scientific-number");
            i64_attr(
                out,
                "number:min-exponent-digits",
                value.min_exponent_digits,
            );
            u64_attr(out, "number:exponent-interval", value.exponent_interval);
            bool_attr(
                out,
                "number:forced-exponent-sign",
                value.forced_exponent_sign,
            );
            i64_attr(out, "number:decimal-places", value.decimal_places);
            i64_attr(
                out,
                "number:min-decimal-places",
                value.min_decimal_places,
            );
            i64_attr(
                out,
                "number:min-integer-digits",
                value.min_integer_digits,
            );
            bool_attr(out, "number:grouping", value.grouping);
            out.push_str("/>");
        }
        OdfDataStylePartToken::Fraction(value) => {
            out.push_str("<number:fraction");
            i64_attr(
                out,
                "number:min-numerator-digits",
                value.min_numerator_digits,
            );
            i64_attr(
                out,
                "number:min-denominator-digits",
                value.min_denominator_digits,
            );
            i64_attr(
                out,
                "number:denominator-value",
                value.denominator_value,
            );
            u64_attr(
                out,
                "number:max-denominator-value",
                value.max_denominator_value,
            );
            i64_attr(
                out,
                "number:min-integer-digits",
                value.min_integer_digits,
            );
            bool_attr(out, "number:grouping", value.grouping);
            out.push_str("/>");
        }
        OdfDataStylePartToken::CurrencySymbol(value) => {
            out.push_str("<number:currency-symbol");
            locale_attrs(out, &value.locale);
            out.push('>');
            out.push_str(&esc(&value.text));
            out.push_str("</number:currency-symbol>");
        }
        OdfDataStylePartToken::Day(value) => write_calendar(out, "day", value),
        OdfDataStylePartToken::Year(value) => write_calendar(out, "year", value),
        OdfDataStylePartToken::Era(value) => write_calendar(out, "era", value),
        OdfDataStylePartToken::DayOfWeek(value) => {
            write_calendar(out, "day-of-week", value)
        }
        OdfDataStylePartToken::Quarter(value) => write_calendar(out, "quarter", value),
        OdfDataStylePartToken::Month(value) => {
            out.push_str("<number:month");
            short_long_attr(out, value.style);
            bool_attr(out, "number:textual", value.textual);
            bool_attr(out, "number:possessive-form", value.possessive_form);
            attr(out, "number:calendar", value.calendar.as_deref());
            out.push_str("/>");
        }
        OdfDataStylePartToken::WeekOfYear(value) => {
            out.push_str("<number:week-of-year");
            attr(out, "number:calendar", value.calendar.as_deref());
            out.push_str("/>");
        }
        OdfDataStylePartToken::Hours(value) => write_clock(out, "hours", value),
        OdfDataStylePartToken::Minutes(value) => write_clock(out, "minutes", value),
        OdfDataStylePartToken::Seconds(value) => {
            out.push_str("<number:seconds");
            short_long_attr(out, value.style);
            i64_attr(out, "number:decimal-places", value.decimal_places);
            out.push_str("/>");
        }
        OdfDataStylePartToken::AmPm => out.push_str("<number:am-pm/>"),
        OdfDataStylePartToken::Boolean => out.push_str("<number:boolean/>"),
        OdfDataStylePartToken::TextContent => out.push_str("<number:text-content/>"),
    }
    Ok(())
}

fn write_calendar(out: &mut String, local: &str, value: &OdfCalendarToken) {
    out.push_str("<number:");
    out.push_str(local);
    short_long_attr(out, value.style);
    attr(out, "number:calendar", value.calendar.as_deref());
    out.push_str("/>");
}

fn write_clock(out: &mut String, local: &str, value: &OdfClockToken) {
    out.push_str("<number:");
    out.push_str(local);
    short_long_attr(out, value.style);
    out.push_str("/>");
}

fn element_text(out: &mut String, qname: &str, value: &str) {
    out.push('<');
    out.push_str(qname);
    out.push('>');
    out.push_str(&esc(value));
    out.push_str("</");
    out.push_str(qname);
    out.push('>');
}

fn locale_attrs(out: &mut String, locale: &OdfNumberLocale) {
    attr(out, "number:language", locale.language.as_deref());
    attr(out, "number:country", locale.country.as_deref());
    attr(out, "number:script", locale.script.as_deref());
    attr(
        out,
        "number:rfc-language-tag",
        locale.rfc_language_tag.as_deref(),
    );
}

fn short_long_attr(out: &mut String, value: Option<OdfShortLong>) {
    if let Some(value) = value {
        attr(out, "number:style", Some(value.as_str()));
    }
}

fn attr(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&esc(value));
        out.push('"');
    }
}

fn bool_attr(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        attr(out, name, Some(if value { "true" } else { "false" }));
    }
}

fn i64_attr(out: &mut String, name: &str, value: Option<i64>) {
    if let Some(value) = value {
        attr(out, name, Some(&value.to_string()));
    }
}

fn u64_attr(out: &mut String, name: &str, value: Option<u64>) {
    if let Some(value) = value {
        attr(out, name, Some(&value.to_string()));
    }
}

fn f64_attr(out: &mut String, name: &str, value: Option<f64>) {
    if let Some(value) = value {
        let lexical = if value.is_nan() {
            "NaN".to_string()
        } else if value == f64::INFINITY {
            "INF".to_string()
        } else if value == f64::NEG_INFINITY {
            "-INF".to_string()
        } else {
            value.to_string()
        };
        attr(out, name, Some(&lexical));
    }
}

fn esc(value: &str) -> String {
    litchi_core::xml::escape_xml(value)
}

#[derive(Clone)]
struct XmlSpan {
    start: usize,
    end: usize,
}

#[derive(Clone)]
struct ContainerSpan {
    section: OdfDataStyleSection,
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}

/// Losslessly insert or replace a data style in an existing style container.
pub fn set_data_style_xml(xml: &str, style: &OdfDataStyle) -> Result<String> {
    let version = document_version(xml)?;
    let fragment = style.to_xml_fragment(version)?;
    let (target, container) = find_style_span(xml, style.section, &style.name)?;
    if let Some(span) = target {
        return Ok(format!("{}{}{}", &xml[..span.start], fragment, &xml[span.end..]));
    }
    let container = container.ok_or_else(|| bad("target ODF style container does not exist"))?;
    if container.empty {
        let raw = &xml[container.start..container.end];
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| bad("invalid empty ODF style container"))?;
        let expanded = format!(
            "{}>{}</{}>",
            &raw[..slash], fragment, container.qname
        );
        return Ok(format!(
            "{}{}{}",
            &xml[..container.start],
            expanded,
            &xml[container.end..]
        ));
    }
    Ok(format!(
        "{}{}{}",
        &xml[..container.end_start],
        fragment,
        &xml[container.end_start..]
    ))
}

/// Losslessly remove one named data style from the requested section.
pub fn remove_data_style_xml(
    xml: &str,
    section: OdfDataStyleSection,
    name: &str,
) -> Result<String> {
    validate_name(name, "style:name")?;
    let (target, _) = find_style_span(xml, section, name)?;
    let target = target.ok_or_else(|| bad("target data style does not exist"))?;
    Ok(format!("{}{}", &xml[..target.start], &xml[target.end..]))
}

fn document_version(xml: &str) -> Result<OdfDataStyleVersion> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("data-style XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut xml_version = XmlVersion::Implicit1_0;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid data-style XML: {error}")))?
        {
            Event::Decl(ref declaration) => {
                xml_version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            }
            Event::Start(ref element) | Event::Empty(ref element) => {
                return read_document_version(&reader, element, xml_version);
            }
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited");
            }
            Event::Eof => return invalid("missing ODF document root"),
            _ => {}
        }
        buffer.clear();
    }
}

fn find_style_span(
    xml: &str,
    wanted_section: OdfDataStyleSection,
    wanted_name: &str,
) -> Result<(Option<XmlSpan>, Option<ContainerSpan>)> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("data-style XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut frames = Vec::<Frame>::new();
    let mut active: Option<(usize, usize)> = None;
    let mut target = None;
    let mut container = None;
    let mut container_depth = None;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid data-style XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let direct = direct_style_section(&frames);
                frames.push(Frame {
                    namespace: namespace.clone(),
                    local: local.clone(),
                });
                let depth = frames.len();
                if namespace.as_deref() == Some(OFFICE)
                    && frame_section(frames.last().expect("pushed frame")) == Some(wanted_section)
                {
                    if container.is_some() || container_depth.is_some() {
                        return invalid("duplicate target ODF style container");
                    }
                    container_depth = Some(depth);
                    container = Some(ContainerSpan {
                        section: wanted_section,
                        start,
                        end: 0,
                        end_start: 0,
                        qname: decode(element.name().as_ref(), "container QName")?,
                        empty: false,
                    });
                }
                if direct == Some(wanted_section)
                    && namespace.as_deref() == Some(NUMBER)
                    && OdfDataStyleKind::parse(&local).is_some()
                {
                    let mut aggregate = 0;
                    let attrs = collect_attributes(
                        &reader,
                        element,
                        XmlVersion::Implicit1_0,
                        &mut aggregate,
                    )?;
                    if attrs.iter().any(|attribute| {
                        attribute.namespace.as_deref() == Some(STYLE)
                            && attribute.local == "name"
                            && attribute.value == wanted_name
                    }) {
                        if active.is_some() || target.is_some() {
                            return invalid("duplicate target data style");
                        }
                        active = Some((depth, start));
                    }
                }
            }
            Event::Empty(ref element) => {
                let namespace = namespace_uri(&resolved)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let direct = direct_style_section(&frames);
                let current = Frame {
                    namespace: namespace.clone(),
                    local: local.clone(),
                };
                if frame_section(&current) == Some(wanted_section) {
                    if container.is_some() || container_depth.is_some() {
                        return invalid("duplicate target ODF style container");
                    }
                    container = Some(ContainerSpan {
                        section: wanted_section,
                        start,
                        end,
                        end_start: start,
                        qname: decode(element.name().as_ref(), "container QName")?,
                        empty: true,
                    });
                }
                if direct == Some(wanted_section)
                    && namespace.as_deref() == Some(NUMBER)
                    && OdfDataStyleKind::parse(&local).is_some()
                {
                    let mut aggregate = 0;
                    let attrs = collect_attributes(
                        &reader,
                        element,
                        XmlVersion::Implicit1_0,
                        &mut aggregate,
                    )?;
                    if attrs.iter().any(|attribute| {
                        attribute.namespace.as_deref() == Some(STYLE)
                            && attribute.local == "name"
                            && attribute.value == wanted_name
                    }) {
                        if target.is_some() {
                            return invalid("duplicate target data style");
                        }
                        target = Some(XmlSpan { start, end });
                    }
                }
            }
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let end_start = event_start(xml, end)?;
                let depth = frames.len();
                if active.is_some_and(|(active_depth, _)| active_depth == depth) {
                    let (_, start) = active.take().expect("active target");
                    if target.replace(XmlSpan { start, end }).is_some() {
                        return invalid("duplicate target data style");
                    }
                }
                if container_depth == Some(depth) {
                    let value = container.as_mut().expect("active target container");
                    value.end = end;
                    value.end_start = end_start;
                    container_depth = None;
                }
                frames
                    .pop()
                    .ok_or_else(|| bad("data-style element stack underflow"))?;
            }
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited");
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !frames.is_empty() || active.is_some() || container_depth.is_some() {
        return invalid("truncated data-style XML");
    }
    if container.as_ref().is_some_and(|value| value.section != wanted_section) {
        return invalid("internal data-style section mismatch");
    }
    Ok((target, container))
}

impl OpenDocumentPackage {
    /// Inspect data styles from `styles.xml` and `content.xml`.
    pub fn data_styles(&self) -> Result<OdfDataStyles> {
        let mut output = OdfDataStyles::default();
        if let Some(styles) = self.styles_xml()? {
            output.append(parse_data_styles_xml(
                &styles,
                OdfDataStylePart::Styles,
            )?)?;
        }
        let content = self.content_xml()?;
        output.append(parse_data_styles_xml(
            &content,
            OdfDataStylePart::Content,
        )?)?;
        Ok(output)
    }
}

impl FlatOpenDocument {
    /// Inspect data styles from both containers in a flat document.
    pub fn data_styles(&self) -> Result<OdfDataStyles> {
        parse_data_styles_xml(self.xml(), OdfDataStylePart::Flat)
    }
}

fn push_style(output: &mut OdfDataStyles, style: OdfDataStyle) -> Result<()> {
    if output.styles.len() >= MAX_STYLES {
        return invalid("too many data styles");
    }
    if output.get(style.source_part, style.section, &style.name).is_some() {
        return invalid("duplicate data style identity");
    }
    output.styles.push(style);
    Ok(())
}

fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
    aggregate: &mut usize,
) -> Result<Vec<Attribute>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| bad(format!("invalid data-style attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if output.len() >= MAX_ATTRIBUTES {
            return invalid("data-style element has too many attributes");
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_uri(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        if !seen.insert((namespace.clone(), local.clone())) {
            return invalid("duplicate expanded data-style attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid data-style attribute value: {error}")))?
            .into_owned();
        add_size(value.len(), aggregate)?;
        output.push(Attribute {
            namespace,
            local,
            value,
        });
    }
    Ok(output)
}

fn read_document_version(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
) -> Result<OdfDataStyleVersion> {
    let mut aggregate = 0usize;
    let attributes = collect_attributes(reader, element, version, &mut aggregate)?;
    match attributes
        .iter()
        .find(|attribute| {
            attribute.namespace.as_deref() == Some(OFFICE) && attribute.local == "version"
        })
        .map(|attribute| attribute.value.as_str())
    {
        None | Some("1.2") => Ok(OdfDataStyleVersion::V1_2),
        Some("1.3") => Ok(OdfDataStyleVersion::V1_3),
        Some(value) => invalid(format!("unsupported ODF data-style version '{value}'")),
    }
}

fn parse_locale(attributes: &mut Vec<Attribute>) -> Result<OdfNumberLocale> {
    Ok(OdfNumberLocale {
        language: take(attributes, NUMBER, "language"),
        country: take(attributes, NUMBER, "country"),
        script: take(attributes, NUMBER, "script"),
        rfc_language_tag: take(attributes, NUMBER, "rfc-language-tag"),
    })
}

fn take(attributes: &mut Vec<Attribute>, namespace: &str, local: &str) -> Option<String> {
    attributes
        .iter()
        .position(|attribute| {
            attribute.namespace.as_deref() == Some(namespace) && attribute.local == local
        })
        .map(|index| attributes.remove(index).value)
}

fn required(attributes: &mut Vec<Attribute>, namespace: &str, local: &str) -> Result<String> {
    take(attributes, namespace, local)
        .ok_or_else(|| bad(format!("missing required {namespace}:{local} attribute")))
}

fn take_bool(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<Option<bool>> {
    take(attributes, namespace, local)
        .map(|value| parse_bool(&value))
        .transpose()
}

fn take_i64(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<Option<i64>> {
    take(attributes, namespace, local)
        .map(|value| parse_i64(&value, local))
        .transpose()
}

fn required_i64(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<i64> {
    let value = required(attributes, namespace, local)?;
    parse_i64(&value, local)
}

fn take_f64(
    attributes: &mut Vec<Attribute>,
    namespace: &str,
    local: &str,
) -> Result<Option<f64>> {
    take(attributes, namespace, local)
        .map(|value| {
            match value.as_str() {
                "INF" => Ok(f64::INFINITY),
                "-INF" => Ok(f64::NEG_INFINITY),
                "NaN" => Ok(f64::NAN),
                _ => {
                    let parsed: f64 = value
                        .parse()
                        .map_err(|_| bad(format!("invalid {local} double '{value}'")))?;
                    if !parsed.is_finite() {
                        return invalid(format!("invalid {local} double '{value}'"));
                    }
                    Ok(parsed)
                }
            }
        })
        .transpose()
}

fn take_versioned_i64(
    attributes: &mut Vec<Attribute>,
    local: &str,
    version: OdfDataStyleVersion,
    alias: &mut bool,
) -> Result<Option<i64>> {
    let standard = take(attributes, NUMBER, local);
    let extension = take(attributes, LOEXT, local);
    if standard.is_some() && extension.is_some() {
        return invalid(format!("duplicate standard/LO alias for {local}"));
    }
    if standard.is_some() && version == OdfDataStyleVersion::V1_2 {
        return invalid(format!("number:{local} requires ODF 1.3"));
    }
    if extension.is_some() {
        *alias = true;
    }
    standard
        .or(extension)
        .map(|value| parse_i64(&value, local))
        .transpose()
}

fn take_versioned_u64(
    attributes: &mut Vec<Attribute>,
    local: &str,
    version: OdfDataStyleVersion,
    alias: &mut bool,
) -> Result<Option<u64>> {
    let standard = take(attributes, NUMBER, local);
    let extension = take(attributes, LOEXT, local);
    if standard.is_some() && extension.is_some() {
        return invalid(format!("duplicate standard/LO alias for {local}"));
    }
    if standard.is_some() && version == OdfDataStyleVersion::V1_2 {
        return invalid(format!("number:{local} requires ODF 1.3"));
    }
    if extension.is_some() {
        *alias = true;
    }
    standard
        .or(extension)
        .map(|value| parse_u64(&value, local))
        .transpose()
}

fn take_versioned_bool(
    attributes: &mut Vec<Attribute>,
    local: &str,
    version: OdfDataStyleVersion,
    alias: &mut bool,
) -> Result<Option<bool>> {
    let standard = take(attributes, NUMBER, local);
    let extension = take(attributes, LOEXT, local);
    if standard.is_some() && extension.is_some() {
        return invalid(format!("duplicate standard/LO alias for {local}"));
    }
    if standard.is_some() && version == OdfDataStyleVersion::V1_2 {
        return invalid(format!("number:{local} requires ODF 1.3"));
    }
    if extension.is_some() {
        *alias = true;
    }
    standard
        .or(extension)
        .map(|value| parse_bool(&value))
        .transpose()
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid ODF boolean '{value}'")),
    }
}

fn parse_i64(value: &str, name: &str) -> Result<i64> {
    value
        .parse()
        .map_err(|_| bad(format!("invalid {name} integer '{value}'")))
}

fn parse_u64(value: &str, name: &str) -> Result<u64> {
    let parsed: u64 = value
        .parse()
        .map_err(|_| bad(format!("invalid {name} positive integer '{value}'")))?;
    if parsed == 0 {
        return invalid(format!("invalid {name} positive integer '{value}'"));
    }
    Ok(parsed)
}

fn reject_remaining(attributes: &[Attribute], element: &str) -> Result<()> {
    if let Some(attribute) = attributes.first() {
        return invalid(format!(
            "unexpected {element} attribute {}:{}",
            attribute.namespace.as_deref().unwrap_or(""),
            attribute.local
        ));
    }
    Ok(())
}

fn ensure_no_children(node: &Node, element: &str) -> Result<()> {
    if !node.children.is_empty() {
        return invalid(format!("{element} cannot contain elements"));
    }
    Ok(())
}

fn ensure_empty_node(node: &Node, element: &str) -> Result<()> {
    ensure_no_children(node, element)?;
    ensure_whitespace(&node.text, element)
}

fn ensure_whitespace(value: &str, element: &str) -> Result<()> {
    if value.chars().any(|character| !character.is_whitespace()) {
        return invalid(format!("{element} cannot contain text"));
    }
    Ok(())
}

fn validate_name(value: &str, name: &str) -> Result<()> {
    validate_text(value, name)?;
    if value.is_empty() || value.chars().any(char::is_control) {
        return invalid(format!("invalid {name}"));
    }
    Ok(())
}

fn validate_cell_address(value: &str) -> Result<()> {
    validate_text(value, "style:base-cell-address")?;
    if value.is_empty() || value.chars().any(char::is_whitespace) || !value.contains('.') {
        return invalid("invalid style:base-cell-address");
    }
    Ok(())
}

fn validate_locale(locale: &OdfNumberLocale) -> Result<()> {
    for (value, name) in [
        (locale.language.as_deref(), "number:language"),
        (locale.country.as_deref(), "number:country"),
        (locale.script.as_deref(), "number:script"),
        (
            locale.rfc_language_tag.as_deref(),
            "number:rfc-language-tag",
        ),
    ] {
        if let Some(value) = value {
            validate_text(value, name)?;
            if value.is_empty()
                || value
                    .bytes()
                    .any(|byte| !(byte.is_ascii_alphanumeric() || byte == b'-'))
            {
                return invalid(format!("invalid {name} '{value}'"));
            }
        }
    }
    Ok(())
}

fn validate_optional_string(value: Option<&str>, name: &str) -> Result<()> {
    if let Some(value) = value {
        validate_text(value, name)?;
    }
    Ok(())
}

fn validate_text(value: &str, name: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds 64 KiB"));
    }
    if value.chars().any(|character| {
        matches!(character, '\u{0}'..='\u{8}' | '\u{B}' | '\u{C}' | '\u{E}'..='\u{1F}')
    }) {
        return invalid(format!("{name} contains an invalid XML character"));
    }
    Ok(())
}

fn reject_spoofed_container(
    namespace: Option<&str>,
    local: &str,
    direct: bool,
) -> Result<()> {
    if direct && OdfDataStyleKind::parse(local).is_some() && namespace != Some(NUMBER) {
        return invalid("data-style container uses the wrong namespace");
    }
    Ok(())
}

fn add_text(target: &mut String, value: &str, aggregate: &mut usize) -> Result<()> {
    add_size(value.len(), aggregate)?;
    if target.len() + value.len() > MAX_VALUE_BYTES {
        return invalid("data-style text exceeds 64 KiB");
    }
    target.push_str(value);
    Ok(())
}

fn add_size(size: usize, aggregate: &mut usize) -> Result<()> {
    if size > MAX_VALUE_BYTES {
        return invalid("data-style value exceeds 64 KiB");
    }
    *aggregate = aggregate
        .checked_add(size)
        .ok_or_else(|| bad("data-style aggregate size overflow"))?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return invalid("data-style metadata exceeds 32 MiB");
    }
    Ok(())
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid data-style XML event boundary"))
}

fn namespace_uri(result: &ResolveResult<'_>) -> Result<Option<String>> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(decode(value, "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        )),
    }
}

fn decode(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| bad(format!("invalid UTF-8 {description}")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(bad(message))
}

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const HEAD_12: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" office:version="1.2"><office:styles>"#;
    const HEAD_13: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3"><office:styles>"#;
    const TAIL: &str = "</office:styles><office:automatic-styles/></office:document>";

    fn doc13(body: &str) -> String {
        format!("{HEAD_13}{body}{TAIL}")
    }

    fn doc12(body: &str) -> String {
        format!("{HEAD_12}{body}{TAIL}")
    }

    #[test]
    fn parses_all_seven_containers_and_all_standard_tokens() {
        let xml = doc13(r##"
            <number:number-style style:name="n" style:display-name="N" number:language="en" number:country="US" number:script="Latn" number:rfc-language-tag="en-US" number:title="title" style:volatile="true" number:transliteration-format="一" number:transliteration-language="zh" number:transliteration-country="CN" number:transliteration-style="medium">
              <style:text-properties fo:color="#ff0000"/><number:text>[</number:text><number:fill-character> </number:fill-character><number:number number:decimal-replacement="--" number:display-factor="1000" number:decimal-places="2" number:min-decimal-places="1" number:min-integer-digits="1" number:grouping="true"><number:embedded-text number:position="1">x</number:embedded-text></number:number><number:text>]</number:text><style:map style:condition="value()&gt;=0" style:apply-style-name="positive" style:base-cell-address="Sheet1.A1"/>
            </number:number-style>
            <number:number-style style:name="s"><number:scientific-number number:min-exponent-digits="2" number:exponent-interval="3" number:forced-exponent-sign="true" number:decimal-places="4" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="false"/></number:number-style>
            <number:number-style style:name="f"><number:fraction number:min-numerator-digits="1" number:min-denominator-digits="2" number:denominator-value="8" number:max-denominator-value="64" number:min-integer-digits="0" number:grouping="false"/></number:number-style>
            <number:currency-style style:name="c" number:automatic-order="true"><number:currency-symbol number:language="fr" number:country="FR">€</number:currency-symbol><number:text> </number:text><number:number/></number:currency-style>
            <number:percentage-style style:name="p"><number:number/><number:text>%</number:text></number:percentage-style>
            <number:date-style style:name="d" number:automatic-order="false" number:format-source="language"><number:day number:style="long" number:calendar="gregorian"/><number:month number:style="short" number:textual="true" number:possessive-form="false" number:calendar="gengou"/><number:year/><number:era/><number:day-of-week/><number:week-of-year number:calendar="ROC"/><number:quarter/><number:hours/><number:minutes/><number:seconds number:style="long" number:decimal-places="3"/><number:am-pm/></number:date-style>
            <number:time-style style:name="t" number:format-source="fixed" number:truncate-on-overflow="false"><number:hours number:style="long"/><number:text>:</number:text><number:minutes/><number:seconds/><number:am-pm/></number:time-style>
            <number:boolean-style style:name="b"><number:text>?</number:text><number:boolean/><number:text>!</number:text></number:boolean-style>
            <number:text-style style:name="x"><number:text-content/><number:text> </number:text><number:text-content/></number:text-style>
        "##);
        let styles = parse_data_styles_xml(&xml, OdfDataStylePart::Flat).unwrap();
        assert_eq!(styles.styles.len(), 9);
        for style in &styles.styles {
            let fragment = style.to_xml_fragment(OdfDataStyleVersion::V1_3).unwrap();
            let reparsed = parse_data_styles_xml(
                &doc13(&fragment),
                OdfDataStylePart::Flat,
            )
            .unwrap();
            assert_eq!(reparsed.styles[0].kind, style.kind);
            assert_eq!(reparsed.styles[0].parts, style.parts);
        }
    }

    #[test]
    fn reads_libreoffice_12_aliases_but_writes_them_only_as_standard_13() {
        let xml = doc12(r#"<number:number-style style:name="n"><loext:fill-character> </loext:fill-character><number:number loext:min-decimal-places="2"/></number:number-style>"#);
        let style = parse_data_styles_xml(&xml, OdfDataStylePart::Flat)
            .unwrap()
            .styles
            .remove(0);
        assert!(style.to_xml_fragment(OdfDataStyleVersion::V1_2).is_err());
        let out = style.to_xml_fragment(OdfDataStyleVersion::V1_3).unwrap();
        assert!(out.contains("number:fill-character"));
        assert!(out.contains("number:min-decimal-places=\"2\""));
        assert!(!out.contains("loext:"));
    }

    #[test]
    fn parses_yielddisc_n122_n126_n170() {
        let fixture = include_str!("../../../3rdparty/libreoffice-core/sc/qa/unit/data/functions/financial/fods/yielddisc.fods");
        fn style<'a>(fixture: &'a str, marker: &str, close: &str) -> &'a str {
            let begin = fixture.find(marker).unwrap();
            let end = begin + fixture[begin..].find(close).unwrap() + close.len();
            &fixture[begin..end]
        }
        let body = format!(
            "{}{}{}",
            style(fixture, r#"<number:currency-style style:name="N122">"#, "</number:currency-style>"),
            style(fixture, r#"<number:text-style style:name="N126">"#, "</number:text-style>"),
            style(fixture, r#"<number:date-style style:name="N170">"#, "</number:date-style>")
        );
        let parsed = parse_data_styles_xml(&doc12(&body), OdfDataStylePart::Flat).unwrap();
        assert_eq!(parsed.styles.len(), 3);
        assert_eq!(parsed.styles[0].maps.len(), 1);
        assert_eq!(parsed.styles[1].maps.len(), 3);
        assert!(matches!(parsed.styles[2].parts[0], OdfDataStylePartToken::DayOfWeek(_)));
    }

    #[test]
    fn accepts_odfdo_default_style_shapes() {
        let body = r#"<number:boolean-style style:name="bool"><number:boolean/></number:boolean-style><number:currency-style style:name="cur"><number:text>-</number:text><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text> </number:text><number:currency-symbol number:language="fr" number:country="FR">€</number:currency-symbol></number:currency-style><number:date-style style:name="date"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style><number:number-style style:name="num"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style><number:percentage-style style:name="pct"><number:number number:decimal-places="2" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style><number:time-style style:name="time"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>"#;
        assert_eq!(
            parse_data_styles_xml(&doc12(body), OdfDataStylePart::Flat)
                .unwrap()
                .styles
                .len(),
            6
        );
    }

    #[test]
    fn rejects_wrong_namespace_order_cardinality_and_lexicals() {
        let invalid = [
            r#"<x:number-style xmlns:x="urn:wrong" style:name="n"/>"#,
            r#"<number:number-style/>"#,
            r#"<number:number-style style:name="n"><number:number/><number:currency-symbol>$</number:currency-symbol></number:number-style>"#,
            r#"<number:date-style style:name="d"/>"#,
            r#"<number:time-style style:name="t"><number:day/></number:time-style>"#,
            r#"<number:number-style style:name="n"><style:map style:condition="x" style:apply-style-name="a"/><number:number/></number:number-style>"#,
            r#"<number:number-style style:name="n"><number:number number:grouping="yes"/></number:number-style>"#,
            r#"<number:number-style style:name="n"><number:number number:decimal-places="1.5"/></number:number-style>"#,
            r#"<number:number-style style:name="n"><number:fraction number:max-denominator-value="0"/></number:number-style>"#,
            r#"<number:number-style style:name="n"><number:number/><style:map style:condition="x"/></number:number-style>"#,
            r#"<number:boolean-style style:name="b"><number:fill-character> </number:fill-character><number:boolean/></number:boolean-style>"#,
        ];
        for body in invalid {
            assert!(
                parse_data_styles_xml(&doc13(body), OdfDataStylePart::Flat).is_err(),
                "accepted {body}"
            );
        }
        assert!(parse_data_styles_xml(&doc12(r#"<number:number-style style:name="n"><number:number number:min-decimal-places="1"/></number:number-style>"#), OdfDataStylePart::Flat).is_err());
    }

    #[test]
    fn accepts_exact_xsd_integer_and_double_lexicals() {
        let body = r#"<number:number-style style:name="plus"><number:number number:decimal-places="+2" number:min-integer-digits="+1" number:display-factor="+1.5"/></number:number-style><number:number-style style:name="inf"><number:number number:display-factor="INF"/></number:number-style><number:number-style style:name="neg"><number:number number:display-factor="-INF"/></number:number-style><number:number-style style:name="nan"><number:number number:display-factor="NaN"/></number:number-style>"#;
        let parsed = parse_data_styles_xml(&doc13(body), OdfDataStylePart::Flat).unwrap();
        let factor = |index: usize| match &parsed.styles[index].parts[0] {
            OdfDataStylePartToken::Number(value) => value.display_factor.unwrap(),
            _ => panic!("expected number token"),
        };
        assert_eq!(factor(0), 1.5);
        assert_eq!(factor(1), f64::INFINITY);
        assert_eq!(factor(2), f64::NEG_INFINITY);
        assert!(factor(3).is_nan());
        assert!(parsed.styles[1].to_xml_fragment(OdfDataStyleVersion::V1_3).unwrap().contains("display-factor=\"INF\""));
        assert!(parsed.styles[2].to_xml_fragment(OdfDataStyleVersion::V1_3).unwrap().contains("display-factor=\"-INF\""));
        assert!(parsed.styles[3].to_xml_fragment(OdfDataStyleVersion::V1_3).unwrap().contains("display-factor=\"NaN\""));
        for lexical in ["inf", "-inf", "+INF", "1e", "++1"] {
            let body = format!(r#"<number:number-style style:name="bad"><number:number number:display-factor="{lexical}"/></number:number-style>"#);
            assert!(parse_data_styles_xml(&doc13(&body), OdfDataStylePart::Flat).is_err());
        }
        assert!(parse_data_styles_xml(&doc13(r#"<number:number-style style:name="bad"><number:number number:decimal-places="++1"/></number:number-style>"#), OdfDataStylePart::Flat).is_err());
    }

    #[test]
    fn lossless_insert_replace_remove_preserves_unrelated_markup() {
        let original = doc13("<!--keep--><number:number-style style:name=\"other\"><number:number/></number:number-style><x:keep xmlns:x=\"urn:x\"/>");
        let mut style = OdfDataStyle::new(
            "new",
            OdfDataStyleKind::Number,
            OdfDataStyleSection::Styles,
        )
        .unwrap();
        style.parts.push(OdfDataStylePartToken::Number(OdfNumberToken::default()));
        let inserted = set_data_style_xml(&original, &style).unwrap();
        assert!(inserted.contains("<!--keep--><number:number-style style:name=\"other\""));
        assert!(inserted.contains("<x:keep xmlns:x=\"urn:x\"/>"));
        style.parts = vec![OdfDataStylePartToken::Text("-".into())];
        let replaced = set_data_style_xml(&inserted, &style).unwrap();
        assert!(replaced.contains("<number:text>-</number:text>"));
        assert_eq!(
            remove_data_style_xml(&replaced, OdfDataStyleSection::Styles, "new").unwrap(),
            original
        );
    }

    #[test]
    fn expands_empty_target_container_and_enforces_caps() {
        let xml = format!("{HEAD_13}</office:styles><office:automatic-styles/></office:document>");
        let mut style = OdfDataStyle::new(
            "auto",
            OdfDataStyleKind::Text,
            OdfDataStyleSection::AutomaticStyles,
        )
        .unwrap();
        style.parts.push(OdfDataStylePartToken::TextContent);
        let output = set_data_style_xml(&xml, &style).unwrap();
        assert!(output.contains("<office:automatic-styles><number:text-style"));
        let huge = "x".repeat(MAX_VALUE_BYTES + 1);
        style.parts = vec![OdfDataStylePartToken::Text(huge)];
        assert!(style.validate(OdfDataStyleVersion::V1_3).is_err());
    }
}
