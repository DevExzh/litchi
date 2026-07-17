//! Typed ODF `style:columns` layout properties.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashMap;
use std::fmt;
use std::str::FromStr;

const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_COLUMN_GROUPS: usize = 65_536;
const MAX_VALUE_BYTES: usize = 4_096;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;
pub const MAX_STYLE_COLUMNS: usize = 64;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct StyleColumnLength(String);

impl StyleColumnLength {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_length(&value)?;
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str { &self.0 }
}

impl fmt::Display for StyleColumnLength {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result { f.write_str(&self.0) }
}

impl FromStr for StyleColumnLength {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> { Self::new(value) }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct StyleColumn {
    pub relative_width: u64,
    pub start_indent: Option<StyleColumnLength>,
    pub end_indent: Option<StyleColumnLength>,
    pub space_before: Option<StyleColumnLength>,
    pub space_after: Option<StyleColumnLength>,
}

impl StyleColumn {
    pub fn new(relative_width: u64) -> Result<Self> {
        if relative_width == 0 { return invalid("style:rel-width must be positive"); }
        Ok(Self { relative_width, start_indent: None, end_indent: None, space_before: None, space_after: None })
    }
    pub fn validate(&self) -> Result<()> {
        if self.relative_width == 0 { return invalid("style:rel-width must be positive"); }
        for value in [&self.start_indent, &self.end_indent, &self.space_before, &self.space_after].into_iter().flatten() { validate_length(value.as_str())?; }
        Ok(())
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum StyleColumnSeparatorStyle { None, #[default] Solid, Dotted, Dashed, DotDashed }
impl StyleColumnSeparatorStyle {
    fn as_str(self) -> &'static str { match self { Self::None => "none", Self::Solid => "solid", Self::Dotted => "dotted", Self::Dashed => "dashed", Self::DotDashed => "dot-dashed" } }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum StyleColumnSeparatorAlignment { Top, #[default] Middle, Bottom }
impl StyleColumnSeparatorAlignment {
    fn as_str(self) -> &'static str { match self { Self::Top => "top", Self::Middle => "middle", Self::Bottom => "bottom" } }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct StyleColumnSeparator {
    pub width: StyleColumnLength,
    pub style: Option<StyleColumnSeparatorStyle>,
    pub height_percent: Option<u8>,
    pub vertical_alignment: Option<StyleColumnSeparatorAlignment>,
    pub color: Option<(u8, u8, u8)>,
}

impl StyleColumnSeparator {
    pub fn new(width: StyleColumnLength) -> Self {
        Self { width, style: None, height_percent: None, vertical_alignment: None, color: None }
    }
    pub fn validate(&self) -> Result<()> {
        validate_length(self.width.as_str())?;
        if self.height_percent.is_some_and(|value| value > 100) { return invalid("style:column-sep height exceeds 100%"); }
        Ok(())
    }
}

/// One `style:columns` value. No explicit children means equal-width columns.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct StyleColumns {
    pub column_count: u8,
    pub column_gap: Option<StyleColumnLength>,
    pub separator: Option<StyleColumnSeparator>,
    columns: Vec<StyleColumn>,
}

impl StyleColumns {
    pub fn new(column_count: u8) -> Result<Self> {
        if column_count == 0 || usize::from(column_count) > MAX_STYLE_COLUMNS { return invalid("fo:column-count must be in 1..=64"); }
        Ok(Self { column_count, column_gap: None, separator: None, columns: Vec::new() })
    }
    pub fn try_with_columns(column_count: u8, columns: Vec<StyleColumn>) -> Result<Self> {
        let value = Self { column_count, column_gap: None, separator: None, columns };
        value.validate()?;
        Ok(value)
    }
    pub fn columns(&self) -> &[StyleColumn] { &self.columns }
    pub fn set_columns(&mut self, columns: Vec<StyleColumn>) -> Result<()> {
        let old = std::mem::replace(&mut self.columns, columns);
        if let Err(error) = self.validate() { self.columns = old; return Err(error); }
        Ok(())
    }
    pub fn validate(&self) -> Result<()> {
        if self.column_count == 0 || usize::from(self.column_count) > MAX_STYLE_COLUMNS { return invalid("fo:column-count must be in 1..=64"); }
        if !self.columns.is_empty() && self.columns.len() != usize::from(self.column_count) { return invalid("explicit style:column count must match fo:column-count"); }
        if self.columns.len() > MAX_STYLE_COLUMNS { return invalid("style:columns exceeds 64 explicit columns"); }
        if let Some(gap) = &self.column_gap { validate_length(gap.as_str())?; }
        if let Some(separator) = &self.separator { separator.validate()?; }
        self.columns.iter().try_for_each(StyleColumn::validate)
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::from(r#"<style:columns xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0""#);
        attr(&mut xml, "fo:column-count", &self.column_count.to_string());
        if let Some(gap) = &self.column_gap { attr(&mut xml, "fo:column-gap", gap.as_str()); }
        xml.push('>');
        if let Some(separator) = &self.separator {
            xml.push_str("<style:column-sep");
            if let Some(color) = separator.color { attr(&mut xml, "style:color", &format!("#{:02X}{:02X}{:02X}", color.0, color.1, color.2)); }
            if let Some(height) = separator.height_percent { attr(&mut xml, "style:height", &format!("{height}%")); }
            if let Some(style) = separator.style { attr(&mut xml, "style:style", style.as_str()); }
            if let Some(alignment) = separator.vertical_alignment { attr(&mut xml, "style:vertical-align", alignment.as_str()); }
            attr(&mut xml, "style:width", separator.width.as_str());
            xml.push_str("/>");
        }
        for column in &self.columns {
            xml.push_str("<style:column");
            if let Some(value) = &column.end_indent { attr(&mut xml, "fo:end-indent", value.as_str()); }
            if let Some(value) = &column.space_after { attr(&mut xml, "fo:space-after", value.as_str()); }
            if let Some(value) = &column.space_before { attr(&mut xml, "fo:space-before", value.as_str()); }
            if let Some(value) = &column.start_indent { attr(&mut xml, "fo:start-indent", value.as_str()); }
            attr(&mut xml, "style:rel-width", &format!("{}*", column.relative_width));
            xml.push_str("/>");
        }
        xml.push_str("</style:columns>");
        Ok(xml)
    }
    pub(crate) fn to_page_layout_fragment(&self, name: &str) -> Result<String> {
        validate_name(name)?;
        Ok(format!(r#"<style:page-layout style:name="{}"><style:page-layout-properties>{}</style:page-layout-properties></style:page-layout>"#, escaped(name), self.to_xml_fragment()?))
    }
}

impl crate::OpenDocumentPackage {
    pub fn style_columns(&self) -> Result<Vec<StyleColumns>> {
        let mut result = parse_style_columns(self.styles_xml()?.as_deref().unwrap_or_default())?;
        result.extend(parse_style_columns(&self.content_xml()?)?);
        if result.len() > MAX_COLUMN_GROUPS { return invalid("package exceeds 65536 style:columns values"); }
        Ok(result)
    }
}

impl crate::FlatOpenDocument {
    pub fn style_columns(&self) -> Result<Vec<StyleColumns>> { parse_style_columns(self.xml()) }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum Ns { None, Style, Fo, Other }
#[derive(Clone)]
struct Frame { ns: Ns, local: String, saw_columns: bool }
struct Active { depth: usize, value: StyleColumns, saw_column: bool, child_depth: Option<usize> }
type Attributes = HashMap<(Ns, String), String>;

pub fn parse_style_columns(xml: &str) -> Result<Vec<StyleColumns>> {
    if !xml.contains("columns") { return Ok(Vec::new()); }
    if xml.len() > MAX_XML_BYTES { return invalid("style:columns XML exceeds 64 MiB"); }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<Active> = None;
    let mut result = Vec::new();
    let mut aggregate = 0usize;
    loop {
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| make_error(format!("invalid style:columns XML: {error}")))?;
        let namespace = ns(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref())?;
                spoof(namespace, &local)?;
                start(&reader, element, namespace, &local, &mut stack, &mut active, &mut result, &mut aggregate, false)?;
                stack.push(Frame { ns: namespace, local, saw_columns: false });
                if stack.len() > MAX_DEPTH { return invalid("style:columns XML exceeds 256 levels"); }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref())?;
                spoof(namespace, &local)?;
                start(&reader, element, namespace, &local, &mut stack, &mut active, &mut result, &mut aggregate, true)?;
            },
            Event::End(_) => {
                stack.pop().ok_or_else(|| make_error("style:columns XML depth underflow"))?;
                if let Some(columns) = &mut active {
                    if columns.child_depth == Some(stack.len()) { columns.child_depth = None; }
                    if columns.depth == stack.len() {
                        let columns = active.take().expect("active columns checked").value;
                        columns.validate()?;
                        if result.len() >= MAX_COLUMN_GROUPS { return invalid("XML exceeds 65536 style:columns values"); }
                        result.push(columns);
                    }
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| make_error(format!("invalid style:columns text: {error}")))?;
                if !value.chars().all(char::is_whitespace) { return invalid("style:columns children must be empty"); }
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => return invalid("style:columns cannot contain character data"),
            Event::DocType(_) | Event::PI(_) => return invalid("DTDs and processing instructions are prohibited in style:columns XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() { return invalid("unterminated style:columns XML"); }
    Ok(result)
}

pub(crate) fn parse_page_layout_property_columns(xml: &str) -> Result<Vec<StyleColumns>> {
    let (wrapped, _, _) = scoped_property_xml(xml)?;
    parse_style_columns(&wrapped)
}

fn start(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, namespace: Ns, local: &str,
    stack: &mut [Frame], active: &mut Option<Active>, result: &mut Vec<StyleColumns>,
    aggregate: &mut usize, empty: bool) -> Result<()> {
    if active.is_none() && namespace == Ns::Style && local == "columns" {
        let parent = stack.last_mut().ok_or_else(|| make_error("style:columns has no property parent"))?;
        if parent.ns != Ns::Style || !matches!(parent.local.as_str(), "section-properties" | "graphic-properties" | "page-layout-properties") { return invalid("style:columns has an invalid property parent"); }
        if parent.saw_columns { return invalid("property element has multiple style:columns children"); }
        parent.saw_columns = true;
        let value = parse_columns(reader, element, aggregate)?;
        if empty {
            value.validate()?;
            if result.len() >= MAX_COLUMN_GROUPS { return invalid("XML exceeds 65536 style:columns values"); }
            result.push(value);
        } else {
            *active = Some(Active { depth: stack.len(), value, saw_column: false, child_depth: None });
        }
        return Ok(());
    }
    let Some(columns) = active else { return Ok(()); };
    if columns.child_depth.is_some() { return invalid("style:column and style:column-sep must be empty"); }
    if stack.len() != columns.depth + 1 { return invalid("style:columns descendants have invalid nesting"); }
    match (namespace, local) {
        (Ns::Style, "column-sep") => {
            if columns.value.separator.is_some() || columns.saw_column { return invalid("style:column-sep is duplicate or follows style:column"); }
            columns.value.separator = Some(parse_separator(reader, element, aggregate)?);
        },
        (Ns::Style, "column") => {
            if columns.value.columns.len() >= MAX_STYLE_COLUMNS { return invalid("style:columns exceeds 64 explicit columns"); }
            columns.saw_column = true;
            columns.value.columns.push(parse_column(reader, element, aggregate)?);
        },
        _ => return invalid("style:columns contains an unsupported child"),
    }
    if !empty { columns.child_depth = Some(stack.len()); }
    Ok(())
}

fn parse_columns(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, aggregate: &mut usize) -> Result<StyleColumns> {
    let mut values = attributes(reader, element, aggregate)?;
    let count = take(&mut values, Ns::Fo, "column-count").ok_or_else(|| make_error("style:columns is missing fo:column-count"))?.parse::<u8>().map_err(|_| make_error("invalid fo:column-count"))?;
    let mut result = StyleColumns::new(count)?;
    result.column_gap = take(&mut values, Ns::Fo, "column-gap").map(StyleColumnLength::new).transpose()?;
    reject(&values, "style:columns")?;
    Ok(result)
}

fn parse_column(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, aggregate: &mut usize) -> Result<StyleColumn> {
    let mut values = attributes(reader, element, aggregate)?;
    let lexical = take(&mut values, Ns::Style, "rel-width").ok_or_else(|| make_error("style:column is missing style:rel-width"))?;
    let width = lexical.strip_suffix('*').filter(|value| !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())).ok_or_else(|| make_error("invalid style:rel-width"))?.parse::<u64>().map_err(|_| make_error("invalid style:rel-width"))?;
    let mut result = StyleColumn::new(width)?;
    result.start_indent = take(&mut values, Ns::Fo, "start-indent").map(StyleColumnLength::new).transpose()?;
    result.end_indent = take(&mut values, Ns::Fo, "end-indent").map(StyleColumnLength::new).transpose()?;
    result.space_before = take(&mut values, Ns::Fo, "space-before").map(StyleColumnLength::new).transpose()?;
    result.space_after = take(&mut values, Ns::Fo, "space-after").map(StyleColumnLength::new).transpose()?;
    reject(&values, "style:column")?;
    Ok(result)
}

fn parse_separator(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, aggregate: &mut usize) -> Result<StyleColumnSeparator> {
    let mut values = attributes(reader, element, aggregate)?;
    let mut result = StyleColumnSeparator::new(StyleColumnLength::new(take(&mut values, Ns::Style, "width").ok_or_else(|| make_error("style:column-sep is missing style:width"))?)?);
    result.style = take(&mut values, Ns::Style, "style").map(|value| match value.as_str() { "none" => Ok(StyleColumnSeparatorStyle::None), "solid" => Ok(StyleColumnSeparatorStyle::Solid), "dotted" => Ok(StyleColumnSeparatorStyle::Dotted), "dashed" => Ok(StyleColumnSeparatorStyle::Dashed), "dot-dashed" => Ok(StyleColumnSeparatorStyle::DotDashed), _ => invalid("invalid style:column-sep style") }).transpose()?;
    result.height_percent = take(&mut values, Ns::Style, "height").map(|value| parse_percent(&value)).transpose()?;
    result.vertical_alignment = take(&mut values, Ns::Style, "vertical-align").map(|value| match value.as_str() { "top" => Ok(StyleColumnSeparatorAlignment::Top), "middle" => Ok(StyleColumnSeparatorAlignment::Middle), "bottom" => Ok(StyleColumnSeparatorAlignment::Bottom), _ => invalid("invalid style:column-sep vertical alignment") }).transpose()?;
    result.color = take(&mut values, Ns::Style, "color").map(|value| parse_color(&value)).transpose()?;
    reject(&values, "style:column-sep")?;
    result.validate()?;
    Ok(result)
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, aggregate: &mut usize) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| make_error(format!("invalid style:columns attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(&resolved)?, decode(local.as_ref())?);
        let value = attribute.decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder()).map_err(|error| make_error(format!("invalid style:columns attribute: {error}")))?.into_owned();
        if value.len() > MAX_VALUE_BYTES { return invalid("style:columns attribute exceeds 4096 bytes"); }
        *aggregate = aggregate.checked_add(value.len()).ok_or_else(|| make_error("style:columns size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES { return invalid("style:columns values exceed 16 MiB"); }
        if result.insert(key, value).is_some() { return invalid("duplicate expanded style:columns attribute"); }
    }
    Ok(result)
}

fn ns(resolved: &ResolveResult<'_>) -> Result<Ns> {
    match resolved {
        ResolveResult::Unbound => Ok(Ns::None),
        ResolveResult::Bound(value) => { let value: &[u8] = value.as_ref(); Ok(if value == STYLE_NS { Ns::Style } else if value == FO_NS { Ns::Fo } else { Ns::Other }) },
        ResolveResult::Unknown(prefix) => invalid(format!("unbound namespace prefix '{}'", String::from_utf8_lossy(prefix.as_ref()))),
    }
}
fn spoof(namespace: Ns, local: &str) -> Result<()> { if matches!(local, "columns" | "column" | "column-sep") && namespace != Ns::Style { return invalid(format!("{local} uses the wrong namespace")); } Ok(()) }
fn take(values: &mut Attributes, namespace: Ns, local: &str) -> Option<String> { values.remove(&(namespace, local.to_owned())) }
fn reject(values: &Attributes, element: &str) -> Result<()> { if let Some(((namespace, local), _)) = values.iter().next() { return invalid(format!("unsupported {element} attribute {namespace:?}:{local}")); } Ok(()) }
fn parse_percent(value: &str) -> Result<u8> { value.strip_suffix('%').ok_or_else(|| make_error("invalid style:height"))?.parse::<u8>().map_err(|_| make_error("invalid style:height")).and_then(|value| if value <= 100 { Ok(value) } else { invalid("style:height exceeds 100%") }) }
fn parse_color(value: &str) -> Result<(u8, u8, u8)> { let hex = value.strip_prefix('#').filter(|value| value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit())).ok_or_else(|| make_error("invalid style:color"))?; Ok((u8::from_str_radix(&hex[0..2], 16).expect("hex"), u8::from_str_radix(&hex[2..4], 16).expect("hex"), u8::from_str_radix(&hex[4..6], 16).expect("hex"))) }
fn validate_length(value: &str) -> Result<()> { let unit = ["cm", "mm", "in", "pt", "pc", "px"].into_iter().find(|unit| value.ends_with(unit)).ok_or_else(|| make_error(format!("invalid ODF length '{value}'")))?; let number = &value[..value.len() - unit.len()]; let number = number.strip_prefix('+').or_else(|| number.strip_prefix('-')).unwrap_or(number); let mut parts = number.split('.'); let whole = parts.next().unwrap_or_default(); let fraction = parts.next(); if value.len() > MAX_VALUE_BYTES || parts.next().is_some() || !whole.bytes().all(|byte| byte.is_ascii_digit()) || !fraction.map_or(!whole.is_empty(), |value| !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())) { return invalid(format!("invalid ODF length '{value}'")); } Ok(()) }
fn validate_name(value: &str) -> Result<()> { if value.is_empty() || value.len() > MAX_VALUE_BYTES || value.chars().any(char::is_control) { return invalid("invalid style name"); } Ok(()) }
fn attr(xml: &mut String, name: &str, value: &str) { xml.push(' '); xml.push_str(name); xml.push_str("=\""); escape(xml, value); xml.push('"'); }
fn escaped(value: &str) -> String { let mut result = String::new(); escape(&mut result, value); result }
fn escape(xml: &mut String, value: &str) { for character in value.chars() { match character { '&' => xml.push_str("&amp;"), '<' => xml.push_str("&lt;"), '>' => xml.push_str("&gt;"), '"' => xml.push_str("&quot;"), '\'' => xml.push_str("&apos;"), _ => xml.push(character) } } }
fn decode(value: &[u8]) -> Result<String> { std::str::from_utf8(value).map(str::to_owned).map_err(|_| make_error("non-UTF-8 style:columns name")) }
fn invalid<T>(message: impl Into<String>) -> Result<T> { Err(make_error(message)) }
fn make_error(message: impl Into<String>) -> Error { Error::InvalidFormat(message.into()) }

pub(crate) fn replace_page_layout_columns(layout: &crate::PageLayout, columns: &StyleColumns) -> Result<String> {
    columns.validate()?;
    let fragment = columns.to_xml_fragment()?;
    if let Some(properties) = &layout.properties {
        let existing = parse_page_layout_property_columns(&properties.xml)?;
        let new_properties = if existing.is_empty() {
            insert_before_end(&properties.xml, &fragment, "style:page-layout-properties")?
        } else {
            replace_first_columns(&properties.xml, &fragment)?
        };
        return self_contained_layout(&layout.xml.replacen(&properties.xml, &new_properties, 1));
    }
    let properties = format!(r#"<style:page-layout-properties xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">{fragment}</style:page-layout-properties>"#);
    self_contained_layout(&insert_before_end(&layout.xml, &properties, "style:page-layout")?)
}

fn replace_first_columns(xml: &str, replacement: &str) -> Result<String> {
    let (wrapped, prefix_len, suffix_len) = scoped_property_xml(xml)?;
    let mut reader = NsReader::from_str(&wrapped);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize)> = None;
    loop {
        let start = reader.buffer_position() as usize;
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| make_error(format!("invalid columns replacement XML: {error}")))?;
        let is_columns = ns(&resolved)? == Ns::Style;
        let event = event.into_owned();
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) if active.is_none() && is_columns && element.local_name().as_ref() == b"columns" => active = Some((start, 1)),
            Event::Empty(ref element) if active.is_none() && is_columns && element.local_name().as_ref() == b"columns" => return splice_scoped(&wrapped, start, end, replacement, prefix_len, suffix_len),
            Event::Start(_) if active.is_some() => active.as_mut().expect("active").1 += 1,
            Event::End(_) if active.is_some() => { let value = active.as_mut().expect("active"); value.1 -= 1; if value.1 == 0 { return splice_scoped(&wrapped, value.0, end, replacement, prefix_len, suffix_len); } },
            Event::Eof => return invalid("page-layout properties have no style:columns to replace"),
            _ => {},
        }
        buffer.clear();
    }
}
pub(crate) fn scoped_property_xml(xml: &str) -> Result<(String, usize, usize)> {
    let prefixes = inferred_prefixes(xml)?;
    let default_style = ["page-layout-properties", "columns", "column", "column-sep"]
        .iter()
        .any(|local| xml.contains(&format!("<{local}")) || xml.contains(&format!("</{local}")));
    let mut open = String::from("<column-scope");
    if default_style { open.push_str(r#" xmlns="urn:oasis:names:tc:opendocument:xmlns:style:1.0""#); }
    for (prefix, uri) in prefixes { attr(&mut open, &format!("xmlns:{prefix}"), uri); }
    open.push('>');
    const CLOSE: &str = "</column-scope>";
    let prefix_len = open.len();
    let mut wrapped = String::with_capacity(prefix_len + xml.len() + CLOSE.len());
    wrapped.push_str(&open); wrapped.push_str(xml); wrapped.push_str(CLOSE);
    Ok((wrapped, prefix_len, CLOSE.len()))
}
fn inferred_prefixes(xml: &str) -> Result<std::collections::BTreeMap<String, &'static str>> {
    let mut prefixes = std::collections::BTreeMap::<String, &'static str>::new();
    prefixes.insert("style".to_owned(), "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
    prefixes.insert("fo".to_owned(), "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
    for local in ["page-layout", "page-layout-properties", "columns", "column", "column-sep", "footnote-sep", "rel-width", "width", "height", "style", "line-style", "adjustment", "distance-before-sep", "distance-after-sep", "vertical-align", "color"] {
        infer_prefixes(xml, local, "urn:oasis:names:tc:opendocument:xmlns:style:1.0", &mut prefixes)?;
    }
    for local in ["column-count", "column-gap", "start-indent", "end-indent", "space-before", "space-after"] {
        infer_prefixes(xml, local, "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0", &mut prefixes)?;
    }
    Ok(prefixes)
}
pub(crate) fn self_contained_layout(xml: &str) -> Result<String> {
    let prefixes = inferred_prefixes(xml)?;
    let trimmed = xml.trim_start();
    let leading = xml.len() - trimmed.len();
    let insert = trimmed.find(|character: char| character.is_whitespace() || matches!(character, '/' | '>')).ok_or_else(|| make_error("malformed page-layout fragment"))? + leading;
    let mut declarations = String::new();
    for (prefix, uri) in prefixes {
        if !xml.contains(&format!("xmlns:{prefix}=")) { attr(&mut declarations, &format!("xmlns:{prefix}"), uri); }
    }
    let mut output = String::with_capacity(xml.len() + declarations.len());
    output.push_str(&xml[..insert]); output.push_str(&declarations); output.push_str(&xml[insert..]);
    Ok(output)
}
fn infer_prefixes(xml: &str, local: &str, uri: &'static str, prefixes: &mut std::collections::BTreeMap<String, &'static str>) -> Result<()> {
    let needle = format!(":{local}");
    for (colon, _) in xml.match_indices(&needle) {
        let end = colon + needle.len();
        if xml.as_bytes().get(end).is_some_and(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')) { continue; }
        let start = xml[..colon].rfind(|character: char| !(character.is_ascii_alphanumeric() || matches!(character, '_' | '-' | '.'))).map_or(0, |index| index + 1);
        let prefix = &xml[start..colon];
        let name_context = start > 0 && (xml.as_bytes()[start - 1].is_ascii_whitespace()
            || xml.as_bytes()[start - 1] == b'<'
            || (xml.as_bytes()[start - 1] == b'/' && start > 1 && xml.as_bytes()[start - 2] == b'<'));
        if !name_context { continue; }
        if prefix == "xmlns" { continue; }
        if prefix.is_empty() || !prefix.as_bytes()[0].is_ascii_alphabetic() || !prefix.bytes().all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')) || prefix == "xml" { return invalid("invalid inherited namespace prefix in page-layout properties"); }
        if prefixes.insert(prefix.to_owned(), uri).is_some_and(|existing| existing != uri) { return invalid(format!("conflicting inherited namespace prefix '{prefix}'")); }
    }
    Ok(())
}
fn splice_scoped(xml: &str, start: usize, end: usize, replacement: &str, prefix: usize, suffix: usize) -> Result<String> { let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len()); output.push_str(&xml[..start]); output.push_str(replacement); output.push_str(&xml[end..]); Ok(output[prefix..output.len() - suffix].to_owned()) }
pub(crate) fn insert_before_end(xml: &str, fragment: &str, expected: &str) -> Result<String> {
    if let Some(index) = xml.rfind("/>") {
        if !xml[index + 2..].trim().is_empty() { return invalid(format!("malformed {expected}")); }
        let trimmed = xml.trim_start();
        let name = trimmed.strip_prefix('<')
            .and_then(|value| value.split(|character: char| character.is_whitespace() || matches!(character, '/' | '>')).next())
            .filter(|value| !value.is_empty())
            .ok_or_else(|| make_error(format!("malformed {expected}")))?;
        let mut output = String::with_capacity(xml.len() + fragment.len() + name.len());
        output.push_str(&xml[..index]);
        output.push('>');
        output.push_str(fragment);
        output.push_str("</");
        output.push_str(name);
        output.push('>');
        return Ok(output);
    }
    let index = xml.rfind("</").ok_or_else(|| make_error(format!("malformed {expected}")))?;
    let mut output = String::with_capacity(xml.len() + fragment.len());
    output.push_str(&xml[..index]);
    output.push_str(fragment);
    output.push_str(&xml[index..]);
    Ok(output)
}
