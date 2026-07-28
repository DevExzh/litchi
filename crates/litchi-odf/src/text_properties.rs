//! Complete typed ODF `style:text-properties` support.

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::{BTreeMap, HashSet};
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML: usize = 32 * 1024 * 1024;
const MAX_VALUE: usize = 1024 * 1024;
const MAX_ATTRIBUTES: usize = 128;
const MAX_DEPTH: usize = 128;
const MAX_STYLES: usize = 65_536;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;
fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty())
        || value.len() > MAX_VALUE
        || value.chars().any(
            |c| matches!(c,'\0'..='\u{8}'|'\u{b}'|'\u{c}'|'\u{e}'..='\u{1f}'|'\u{fffe}'|'\u{ffff}'),
        )
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn decimal(value: &str, signed: bool) -> bool {
    let value = if signed {
        value.strip_prefix('-').unwrap_or(value)
    } else {
        value
    };
    if value.is_empty() {
        return false;
    }
    let mut parts = value.split('.');
    let left = parts.next().unwrap_or_default();
    let right = parts.next();
    if parts.next().is_some() {
        return false;
    }
    match right {
        None => !left.is_empty() && left.bytes().all(|b| b.is_ascii_digit()),
        Some(right) => {
            (!left.is_empty() || !right.is_empty())
                && left.bytes().all(|b| b.is_ascii_digit())
                && right.bytes().all(|b| b.is_ascii_digit())
        },
    }
}
fn length(value: &str, signed: bool, positive: bool) -> bool {
    ["cm", "mm", "in", "pt", "pc", "px"].iter().any(|unit| {
        value.strip_suffix(unit).is_some_and(|number| {
            decimal(number, signed)
                && (!positive || number.bytes().any(|b| b.is_ascii_digit() && b != b'0'))
        })
    })
}
fn percent(value: &str) -> bool {
    value
        .strip_suffix('%')
        .is_some_and(|value| decimal(value, true))
}
fn positive_integer(value: &str) -> bool {
    let value = value.strip_prefix('+').unwrap_or(value);
    !value.is_empty()
        && value.bytes().all(|b| b.is_ascii_digit())
        && value.bytes().any(|b| b != b'0')
}
fn ncname(value: &str, empty: bool) -> bool {
    if value.is_empty() {
        return empty;
    }
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|c| c == '_' || c == '-' || c == '.' || c.is_alphanumeric())
}
fn language(value: &str) -> bool {
    !value.is_empty()
        && value.split('-').all(|part| {
            !part.is_empty() && part.len() <= 8 && part.bytes().all(|b| b.is_ascii_alphanumeric())
        })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum TextPropertyNamespace {
    Fo,
    Style,
    Text,
}
impl TextPropertyNamespace {
    pub fn prefix(self) -> &'static str {
        match self {
            Self::Fo => "fo",
            Self::Style => "style",
            Self::Text => "text",
        }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextPropertyValue {
    Boolean(bool),
    Color(String),
    Length(String),
    PositiveLength(String),
    Percent(String),
    PositiveInteger(String),
    Character(char),
    Language(String),
    CountryCode(String),
    LanguageCode(String),
    ScriptCode(String),
    TextEncoding(String),
    StyleNameRef(String),
    Keyword(String),
    Text(String),
    Compound(String),
}
impl TextPropertyValue {
    pub fn lexical(&self) -> String {
        match self {
            Self::Boolean(value) => value.to_string(),
            Self::Character(value) => value.to_string(),
            Self::Color(value)
            | Self::Length(value)
            | Self::PositiveLength(value)
            | Self::Percent(value)
            | Self::PositiveInteger(value)
            | Self::Language(value)
            | Self::CountryCode(value)
            | Self::LanguageCode(value)
            | Self::ScriptCode(value)
            | Self::TextEncoding(value)
            | Self::StyleNameRef(value)
            | Self::Keyword(value)
            | Self::Text(value)
            | Self::Compound(value) => value.clone(),
        }
    }
}
include!("text_property_specs.rs");
fn validate_ref(reference: &str, value: &str, depth: usize) -> Option<TextPropertyValue> {
    if depth > 8 {
        return None;
    }
    let direct = match reference {
        "boolean" => match value {
            "true" => Some(TextPropertyValue::Boolean(true)),
            "false" => Some(TextPropertyValue::Boolean(false)),
            _ => None,
        },
        "color" => (value.len() == 7
            && value.starts_with('#')
            && value[1..].bytes().all(|b| b.is_ascii_hexdigit()))
        .then(|| TextPropertyValue::Color(value.to_owned())),
        "length" => length(value, true, false).then(|| TextPropertyValue::Length(value.to_owned())),
        "positiveLength" => {
            length(value, false, true).then(|| TextPropertyValue::PositiveLength(value.to_owned()))
        },
        "percent" => percent(value).then(|| TextPropertyValue::Percent(value.to_owned())),
        "positiveInteger" => {
            positive_integer(value).then(|| TextPropertyValue::PositiveInteger(value.to_owned()))
        },
        "character" => {
            let mut chars = value.chars();
            let first = chars.next();
            (first.is_some() && chars.next().is_none())
                .then(|| TextPropertyValue::Character(first.unwrap()))
        },
        "language" => language(value).then(|| TextPropertyValue::Language(value.to_owned())),
        "languageCode" => {
            (value.len() >= 2 && value.len() <= 8 && value.bytes().all(|b| b.is_ascii_alphabetic()))
                .then(|| TextPropertyValue::LanguageCode(value.to_owned()))
        },
        "countryCode" => (value.len() == 2 && value.bytes().all(|b| b.is_ascii_alphabetic()))
            .then(|| TextPropertyValue::CountryCode(value.to_owned())),
        "scriptCode" => (value.len() == 4 && value.bytes().all(|b| b.is_ascii_alphabetic()))
            .then(|| TextPropertyValue::ScriptCode(value.to_owned())),
        "textEncoding" => Some(TextPropertyValue::TextEncoding(value.to_owned())),
        "styleNameRef" => {
            ncname(value, true).then(|| TextPropertyValue::StyleNameRef(value.to_owned()))
        },
        "angle" | "string" | "shadowType" => Some(TextPropertyValue::Text(value.to_owned())),
        _ => None,
    };
    if direct.is_some() {
        return direct;
    }
    let keywords = leaf_keywords(reference);
    if keywords.contains(&value) {
        return Some(TextPropertyValue::Keyword(value.to_owned()));
    }
    let references = leaf_refs(reference);
    if leaf_is_list(reference) {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.iter().all(|part| {
            keywords.contains(part)
                || references
                    .iter()
                    .any(|reference| validate_ref(reference, part, depth + 1).is_some())
        }) {
            return Some(TextPropertyValue::Compound(value.to_owned()));
        }
        return None;
    }
    references
        .iter()
        .find_map(|reference| validate_ref(reference, value, depth + 1))
}
fn validate_spec(
    value: &str,
    keywords: &[&str],
    references: &[&str],
    list: bool,
    kind: TextPropertyKind,
) -> Result<TextPropertyValue> {
    safe(value, "text property value", true)?;
    if list {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.iter().all(|part| {
            keywords.contains(part)
                || references
                    .iter()
                    .any(|reference| validate_ref(reference, part, 0).is_some())
        }) {
            return Ok(TextPropertyValue::Compound(value.to_owned()));
        }
        return Err(bad(format!(
            "invalid {}:{} list",
            kind.namespace().prefix(),
            kind.local_name()
        )));
    }
    if keywords.contains(&value) {
        return Ok(TextPropertyValue::Keyword(value.to_owned()));
    }
    for reference in references {
        if let Some(value) = validate_ref(reference, value, 0) {
            return Ok(value);
        }
    }
    Err(bad(format!(
        "invalid {}:{} value",
        kind.namespace().prefix(),
        kind.local_name()
    )))
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextProperty {
    kind: TextPropertyKind,
    value: TextPropertyValue,
}
impl TextProperty {
    pub fn new(kind: TextPropertyKind, lexical: &str) -> Result<Self> {
        Ok(Self {
            kind,
            value: kind.parse_value(lexical)?,
        })
    }
    pub fn kind(&self) -> TextPropertyKind {
        self.kind
    }
    pub fn value(&self) -> &TextPropertyValue {
        &self.value
    }
    pub fn lexical(&self) -> String {
        self.value.lexical()
    }
    pub fn qualified_name(&self) -> String {
        format!(
            "{}:{}",
            self.kind.namespace().prefix(),
            self.kind.local_name()
        )
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextStyleProperties {
    properties: BTreeMap<TextPropertyKind, TextPropertyValue>,
}
impl TextStyleProperties {
    pub fn set(&mut self, property: TextProperty) -> Option<TextPropertyValue> {
        self.properties.insert(property.kind, property.value)
    }
    pub fn set_lexical(
        &mut self,
        kind: TextPropertyKind,
        value: &str,
    ) -> Result<Option<TextPropertyValue>> {
        Ok(self.set(TextProperty::new(kind, value)?))
    }
    pub fn get(&self, kind: TextPropertyKind) -> Option<&TextPropertyValue> {
        self.properties.get(&kind)
    }
    pub fn remove(&mut self, kind: TextPropertyKind) -> Option<TextPropertyValue> {
        self.properties.remove(&kind)
    }
    pub fn iter(&self) -> impl Iterator<Item = (TextPropertyKind, &TextPropertyValue)> {
        self.properties.iter().map(|(kind, value)| (*kind, value))
    }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        let xml = format!(
            r#"<office:document xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="fragment" style:family="text">{fragment}</style:style></office:styles></office:document>"#
        );
        let mut set = parse_text_style_properties(&xml)?;
        set.styles
            .pop()
            .and_then(|style| style.properties)
            .ok_or_else(|| bad("fragment does not contain style:text-properties"))
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        let mut xml = format!(
            r#"<style:text-properties xmlns:style="{STYLE_NS}" xmlns:fo="{FO_NS}" xmlns:text="{TEXT_NS}""#
        );
        for (kind, value) in &self.properties {
            xml.push(' ');
            xml.push_str(kind.namespace().prefix());
            xml.push(':');
            xml.push_str(kind.local_name());
            xml.push_str("=\"");
            xml.push_str(&escape_xml(&value.lexical()));
            xml.push('"')
        }
        xml.push_str("/>");
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStyleRecord {
    pub name: Option<String>,
    pub family: String,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<TextStyleProperties>,
}
impl TextStyleRecord {
    pub fn named(
        name: impl Into<String>,
        family: impl Into<String>,
        properties: Option<TextStyleProperties>,
    ) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            family: family.into(),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(
        family: impl Into<String>,
        properties: Option<TextStyleProperties>,
    ) -> Result<Self> {
        let value = Self {
            name: None,
            family: family.into(),
            parent_style_name: None,
            is_default_style: true,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        safe(&self.family, "style:family", false)?;
        match (&self.name, self.is_default_style) {
            (Some(value), false) if ncname(value, false) => {},
            (None, true) => {},
            _ => return Err(bad("invalid text-property style identity")),
        }
        if let Some(value) = &self.parent_style_name {
            if self.is_default_style || !ncname(value, false) {
                return Err(bad("invalid parent style name"));
            }
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
        let mut xml = format!(
            r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="{}""#,
            escape_xml(&self.family)
        );
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)))
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ))
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"))
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextStylePropertiesSet {
    pub styles: Vec<TextStyleRecord>,
}
impl TextStylePropertiesSet {
    pub fn get(&self, family: &str, name: &str) -> Option<&TextStyleRecord> {
        self.styles
            .iter()
            .find(|style| style.family == family && style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self, family: &str) -> Option<&TextStyleRecord> {
        self.styles
            .iter()
            .find(|style| style.family == family && style.is_default_style)
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Ns {
    Office,
    Style,
    Fo,
    Text,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == FO => Ns::Fo,
        ResolveResult::Bound(value) if value.as_ref() == TEXT => Ns::Text,
        _ => Ns::Other,
    }
}
fn property_ns(value: ResolveResult<'_>) -> Option<TextPropertyNamespace> {
    match ns(value) {
        Ns::Fo => Some(TextPropertyNamespace::Fo),
        Ns::Style => Some(TextPropertyNamespace::Style),
        Ns::Text => Some(TextPropertyNamespace::Text),
        _ => None,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ns(namespace), local.as_ref().to_vec())
}
fn attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Option<TextPropertyNamespace>, Vec<u8>, String)>> {
    let mut out = Vec::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid text property attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if out.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many text property attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid text property value: {error}")))?
            .into_owned();
        safe(&value, "text property value", true)?;
        out.push((property_ns(namespace), local.as_ref().to_vec(), value))
    }
    Ok(out)
}
fn header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<TextStyleRecord> {
    let mut family = None;
    let mut name = None;
    let mut parent = None;
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid style attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if ns(namespace) != Ns::Style {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid style value: {error}")))?
            .into_owned();
        match local.as_ref() {
            b"family" => family = Some(value),
            b"name" => name = Some(value),
            b"parent-style-name" => parent = Some(value),
            _ => {},
        }
    }
    let value = TextStyleRecord {
        name,
        family: family.ok_or_else(|| bad("style with text properties requires style:family"))?,
        parent_style_name: parent,
        is_default_style: default,
        properties: None,
    };
    value.validate()?;
    Ok(value)
}
fn properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<TextStyleProperties> {
    let mut value = TextStyleProperties::default();
    let mut seen = HashSet::new();
    for (namespace, local, lexical) in attrs(reader, version, start)? {
        let namespace = namespace.ok_or_else(|| bad("unknown style:text-properties namespace"))?;
        let local = std::str::from_utf8(&local).map_err(|_| bad("malformed XML name"))?;
        let kind = TextPropertyKind::from_expanded(namespace, local)
            .ok_or_else(|| bad("unknown style:text-properties attribute or wrong namespace"))?;
        if !seen.insert(kind) {
            return Err(bad("duplicate style:text-properties attribute"));
        }
        value.set_lexical(kind, &lexical)?;
    }
    Ok(value)
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
struct Active {
    depth: usize,
    style: TextStyleRecord,
    seen: bool,
    property_depth: Option<usize>,
}
fn push(out: &mut Vec<TextStyleRecord>, style: TextStyleRecord, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out.iter().any(|value| {
            value.name == style.name
                && value.family == style.family
                && value.is_default_style == style.is_default_style
        })
    {
        return Err(bad("duplicate or excessive text-property style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("text property data is too large"));
    }
    out.push(style);
    Ok(())
}
pub fn parse_text_style_properties(xml: &str) -> Result<TextStylePropertiesSet> {
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
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("styles XML has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML is too deep"));
                }
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    active = Some(Active {
                        depth,
                        style: header(&reader, version, &start, current.1 == b"default-style")?,
                        seen: false,
                        property_depth: None,
                    });
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"text-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:text-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(properties(&reader, version, &start)?);
                        value.property_depth = Some(depth)
                    } else if current.1 == b"text-properties" {
                        return Err(bad("style:text-properties has invalid namespace or parent"));
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("style:text-properties cannot have children"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    push(
                        &mut out,
                        header(&reader, version, &start, current.1 == b"default-style")?,
                        &mut total,
                    )?;
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"text-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:text-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(properties(&reader, version, &start)?)
                    } else if current.1 == b"text-properties" {
                        return Err(bad("style:text-properties has invalid namespace or parent"));
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("style:text-properties cannot have children"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|value| value.property_depth.is_some())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("style:text-properties cannot contain text"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|value| value.property_depth.is_some())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("style:text-properties cannot contain text"));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(value) = active.as_mut() {
                    if value.property_depth == Some(depth) {
                        value.property_depth = None
                    }
                }
                if active.as_ref().is_some_and(|value| value.depth == depth) {
                    push(&mut out, active.take().unwrap().style, &mut total)?
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
    Ok(TextStylePropertiesSet { styles: out })
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
struct Target {
    style: Span,
    properties: Option<Span>,
}
fn replace(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty style"))?;
    Ok(replace(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}
pub fn set_text_style_properties_xml(xml: &str, requested: &TextStyleRecord) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut target_depth = None;
    let mut active: Option<Target> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    let style = header(&reader, version, &start, current.1 == b"default-style")?;
                    if style.name == requested.name
                        && style.family == requested.family
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target style"));
                        }
                        target_depth = Some(depth);
                        active = Some(Target {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"text-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:text-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
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
                    let style = header(&reader, version, &start, current.1 == b"default-style")?;
                    if style.name == requested.name
                        && style.family == requested.family
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target style"));
                        }
                        found = Some(Target {
                            style: span,
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"text-properties"
                    && active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:text-properties"));
                    }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|span| span.end == 0)
                        && target_depth.is_some_and(|d| depth == d + 1)
                    {
                        let span = spans.properties.as_mut().unwrap();
                        span.end_start = begin;
                        span.end = end
                    }
                    if target_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        target_depth = None
                    }
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
    let spans = found.ok_or_else(|| bad("target style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(TextStyleProperties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}
impl OpenDocumentPackage {
    pub fn text_style_properties(&self) -> Result<TextStylePropertiesSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_text_style_properties(&xml),
        )
    }
}
impl FlatOpenDocument {
    pub fn text_style_properties(&self) -> Result<TextStylePropertiesSet> {
        parse_text_style_properties(self.xml())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles>"#;
    fn doc(body: &str) -> String {
        format!("{HEAD}{body}</office:automatic-styles></office:document>")
    }
    #[test]
    fn every_kind_round_trips() {
        assert!(TextProperty::new(TextPropertyKind::StyleTextRotationAngle, "90").is_ok());
        let mut value = TextStyleProperties::default();
        for kind in TextPropertyKind::ALL {
            value.set_lexical(*kind, kind.example()).unwrap();
        }
        let fragment = value.to_xml_fragment().unwrap();
        assert_eq!(
            TextStyleProperties::from_xml_fragment(&fragment).unwrap(),
            value
        );
        assert_eq!(value.iter().count(), 84)
    }
    #[test]
    fn parses_real_libreoffice_style() {
        let fixture = include_str!(
            "../../../test-data/libreoffice-core/xmloff/qa/unit/data/tdf161327_LatheEndAngle.fodg"
        );
        let begin = fixture
            .find(r#"<style:style style:name="Text" style:family="graphic">"#)
            .unwrap();
        let end = begin + fixture[begin..].find("</style:style>").unwrap() + "</style:style>".len();
        let set = parse_text_style_properties(&doc(&fixture[begin..end])).unwrap();
        let value = set
            .get("graphic", "Text")
            .unwrap()
            .properties
            .as_ref()
            .unwrap();
        assert_eq!(
            value
                .get(TextPropertyKind::StyleFontName)
                .unwrap()
                .lexical(),
            "Noto Sans"
        )
    }
    #[test]
    fn lossless_replace_insert_remove() {
        let original = doc(
            "<!--keep--><style:style style:name=\"a\" style:family=\"text\"><x:k xmlns:x=\"urn:k\"/></style:style><style:style style:name=\"b\" style:family=\"paragraph\"><style:text-properties fo:font-size=\"12pt\"/></style:style>",
        );
        let mut value = TextStyleProperties::default();
        value
            .set_lexical(TextPropertyKind::FoFontWeight, "bold")
            .unwrap();
        let mut a = TextStyleRecord::named("a", "text", Some(value)).unwrap();
        let inserted = set_text_style_properties_xml(&original, &a).unwrap();
        assert!(inserted.contains("<x:k xmlns:x=\"urn:k\"/><style:text-properties"));
        a.properties = None;
        let restored = set_text_style_properties_xml(&inserted, &a).unwrap();
        assert_eq!(restored, original);
        let removed = set_text_style_properties_xml(
            &restored,
            &TextStyleRecord::named("b", "paragraph", None).unwrap(),
        )
        .unwrap();
        assert!(!removed.contains("fo:font-size=\"12pt\""))
    }
    #[test]
    fn rejects_lexical_namespace_placement_and_caps() {
        for body in [
            r#"<style:style style:name="a" style:family="text"><style:text-properties fo:font-weight="950"/></style:style>"#,
            r##"<style:style style:name="a" style:family="text"><style:text-properties fo:color="#fff"/></style:style>"##,
            r#"<style:style style:name="a" style:family="text"><style:text-properties fo:font-size="12pt" fo:font-size="13pt"/></style:style>"#,
            r#"<style:style style:name="a" style:family="text"><text:text-properties/></style:style>"#,
            r#"<style:style style:name="a" style:family="text"><style:text-properties><text:span/></style:text-properties></style:style>"#,
        ] {
            assert!(
                parse_text_style_properties(&doc(body)).is_err(),
                "accepted {body}"
            )
        }
        let huge = "x".repeat(MAX_VALUE + 1);
        assert!(TextProperty::new(TextPropertyKind::FoFontFamily, &huge).is_err());
        let mut attrs = String::new();
        for index in 0..=MAX_ATTRIBUTES {
            attrs.push_str(&format!(" x:a{index}=\"1\""))
        }
        assert!(parse_text_style_properties(&doc(&format!(r#"<style:style style:name="a" style:family="text"><style:text-properties xmlns:x="urn:x"{attrs}/></style:style>"#))).is_err())
    }
}
