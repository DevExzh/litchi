//! Complete typed ODF `style:table-column-properties` support.
//!
//! ODF 1.3 section 17.16 allows exactly five attributes on
//! `style:table-column-properties` (`style:column-width`, `style:rel-column-width`,
//! `style:use-optimal-column-width`, `fo:break-before`, `fo:break-after`) and no child
//! elements. Unknown attributes, children, or text are rejected.

use crate::{FlatOpenDocument, OpenDocumentPackage, TableRowBreak};
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
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_ATTRIBUTES: usize = 32;

fn bad(x: impl Into<String>) -> Error {
    Error::InvalidFormat(x.into())
}
fn safe(x: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && x.is_empty()) || x.len() > MAX_VALUE || x.chars().any(char::is_control) {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}

/// A positive ODF physical length used by `style:column-width`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableColumnLength(String);
impl TableColumnLength {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        if x.len() > MAX_VALUE {
            return Err(bad("style:column-width is too large"));
        }
        let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
            .iter()
            .find_map(|unit| x.strip_suffix(unit))
        else {
            return Err(bad("style:column-width must use an ODF physical unit"));
        };
        if number.starts_with(['+', '-']) {
            return Err(bad("style:column-width cannot be signed"));
        }
        let mut parts = number.split('.');
        let whole = parts.next().unwrap_or_default();
        let fraction = parts.next();
        if parts.next().is_some()
            || whole.is_empty()
            || !whole.bytes().all(|c| c.is_ascii_digit())
            || fraction
                .is_some_and(|part| part.is_empty() || !part.bytes().all(|c| c.is_ascii_digit()))
        {
            return Err(bad("invalid style:column-width"));
        }
        if !number.bytes().any(|c| c.is_ascii_digit() && c != b'0') {
            return Err(bad("style:column-width must be positive"));
        }
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// An ODF relative column width (`relativeLength`, digits followed by `*`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableColumnRelWidth(String);
impl TableColumnRelWidth {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        let Some(number) = x.strip_suffix('*') else {
            return Err(bad("style:rel-column-width requires a trailing *"));
        };
        if x.len() > MAX_VALUE || number.is_empty() || !number.bytes().all(|c| c.is_ascii_digit()) {
            return Err(bad("invalid style:rel-column-width"));
        }
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Complete `style:table-column-properties` value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableColumnProperties {
    pub column_width: Option<TableColumnLength>,
    pub rel_column_width: Option<TableColumnRelWidth>,
    pub use_optimal_column_width: Option<bool>,
    pub break_before: Option<TableRowBreak>,
    pub break_after: Option<TableRowBreak>,
}
impl TableColumnProperties {
    pub fn validate(&self) -> Result<()> {
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:table-column-properties xmlns:style="{STYLE_NS}" xmlns:fo="{FO_NS}""#
        );
        if let Some(value) = &self.column_width {
            xml.push_str(&format!(r#" style:column-width="{}""#, value.as_str()));
        }
        if let Some(value) = &self.rel_column_width {
            xml.push_str(&format!(r#" style:rel-column-width="{}""#, value.as_str()));
        }
        if let Some(value) = self.use_optimal_column_width {
            xml.push_str(&format!(r#" style:use-optimal-column-width="{value}""#));
        }
        if let Some(value) = self.break_before {
            xml.push_str(&format!(r#" fo:break-before="{}""#, break_xml(value)));
        }
        if let Some(value) = self.break_after {
            xml.push_str(&format!(r#" fo:break-after="{}""#, break_xml(value)));
        }
        xml.push_str("/>");
        Ok(xml)
    }
}
fn break_xml(x: TableRowBreak) -> &'static str {
    match x {
        TableRowBreak::Auto => "auto",
        TableRowBreak::Column => "column",
        TableRowBreak::Page => "page",
    }
}

/// A named or default table-column style declaration carrying typed column properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableColumnStyleProperties {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<TableColumnProperties>,
}
impl TableColumnStyleProperties {
    pub fn named(
        name: impl Into<String>,
        properties: Option<TableColumnProperties>,
    ) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<TableColumnProperties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(value), false) => safe(value, "table-column style name", false)?,
            (None, true) => {},
            _ => return Err(bad("invalid table-column style identity")),
        }
        if let Some(value) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default table-column style cannot have a parent"));
            }
            safe(value, "parent table-column style name", false)?;
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
        let mut xml =
            format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="table-column""#);
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
pub struct TableColumnStylePropertiesSet {
    pub styles: Vec<TableColumnStyleProperties>,
}
impl TableColumnStylePropertiesSet {
    pub fn get(&self, name: &str) -> Option<&TableColumnStyleProperties> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&TableColumnStyleProperties> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Fo,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(x) if x.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(x) if x.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(x) if x.as_ref() == FO => Ns::Fo,
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
            attribute.map_err(|error| bad(format!("invalid table-column attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if result.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many table-column attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate table-column attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid table-column value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("table-column attribute value is too large"));
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
fn boolean(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(bad("ODF boolean must be true or false")),
    }
}
fn parse_break(value: &str) -> Result<TableRowBreak> {
    match value {
        "auto" => Ok(TableRowBreak::Auto),
        "column" => Ok(TableRowBreak::Column),
        "page" => Ok(TableRowBreak::Page),
        _ => Err(bad("invalid table-column break value")),
    }
}
fn style_header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<TableColumnStyleProperties>> {
    let mut attrs = attributes(reader, version, start)?;
    if take(&mut attrs, Ns::Style, b"family").as_deref() != Some("table-column") {
        return Ok(None);
    }
    let value = TableColumnStyleProperties {
        name: take(&mut attrs, Ns::Style, b"name"),
        parent_style_name: take(&mut attrs, Ns::Style, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    value.validate()?;
    Ok(Some(value))
}
fn column_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<TableColumnProperties> {
    let mut attrs = attributes(reader, version, start)?;
    let value = TableColumnProperties {
        column_width: take(&mut attrs, Ns::Style, b"column-width")
            .map(TableColumnLength::new)
            .transpose()?,
        rel_column_width: take(&mut attrs, Ns::Style, b"rel-column-width")
            .map(TableColumnRelWidth::new)
            .transpose()?,
        use_optimal_column_width: take(&mut attrs, Ns::Style, b"use-optimal-column-width")
            .map(|x| boolean(&x))
            .transpose()?,
        break_before: take(&mut attrs, Ns::Fo, b"break-before")
            .map(|x| parse_break(&x))
            .transpose()?,
        break_after: take(&mut attrs, Ns::Fo, b"break-after")
            .map(|x| parse_break(&x))
            .transpose()?,
    };
    if !attrs.is_empty() {
        return Err(bad("unknown style:table-column-properties attribute"));
    }
    Ok(value)
}
struct Active {
    depth: usize,
    style: TableColumnStyleProperties,
    seen_properties: bool,
    properties_depth: Option<usize>,
}
fn push_style(
    out: &mut Vec<TableColumnStyleProperties>,
    style: TableColumnStyleProperties,
    total: &mut usize,
) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out
            .iter()
            .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate or excessive table-column style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("table-column style data is too large"));
    }
    out.push(style);
    Ok(())
}

/// Parse direct table-column styles in `office:styles` and `office:automatic-styles`.
pub fn parse_table_column_style_properties(xml: &str) -> Result<TableColumnStylePropertiesSet> {
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
                        });
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"table-column-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:table-column-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(column_properties(&reader, version, &start)?);
                        state.properties_depth = Some(depth);
                    } else if current.1 == b"table-column-properties" {
                        return Err(bad(
                            "style:table-column-properties has invalid namespace or parent",
                        ));
                    } else if state.properties_depth.is_some()
                        && depth > state.properties_depth.unwrap()
                    {
                        return Err(bad("unexpected table-column property child"));
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
                        && current.1 == b"table-column-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:table-column-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(column_properties(&reader, version, &start)?);
                    } else if current.1 == b"table-column-properties" {
                        return Err(bad(
                            "style:table-column-properties has invalid namespace or parent",
                        ));
                    } else if state.properties_depth.is_some()
                        && depth > state.properties_depth.unwrap()
                    {
                        return Err(bad("unexpected table-column property child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut()
                    && state.properties_depth.is_some()
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in table-column properties"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut()
                    && state.properties_depth.is_some()
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in table-column properties"));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(state) = active.as_mut()
                    && state.properties_depth == Some(depth)
                {
                    state.properties_depth = None;
                }
                if active.as_ref().is_some_and(|x| x.depth == depth) {
                    push_style(&mut out, active.take().unwrap().style, &mut total)?;
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
    Ok(TableColumnStylePropertiesSet { styles: out })
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

/// Losslessly replace, insert, or remove one existing column style's property element.
pub fn set_table_column_style_properties_xml(
    xml: &str,
    requested: &TableColumnStyleProperties,
) -> Result<String> {
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
                            return Err(bad("duplicate target table-column style"));
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
                    && current.1 == b"table-column-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:table-column-properties"));
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
                            return Err(bad("duplicate target table-column style"));
                        }
                        found = Some(TargetSpans {
                            style: span,
                            ..Default::default()
                        });
                    }
                } else if depth_target.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"table-column-properties"
                    && active.as_mut().unwrap().properties.replace(span).is_some()
                {
                    return Err(bad("duplicate style:table-column-properties"));
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
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid styles XML: {e}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target table-column style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(TableColumnProperties::to_xml_fragment)
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

impl OpenDocumentPackage {
    pub fn table_column_style_properties(&self) -> Result<TableColumnStylePropertiesSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_table_column_style_properties(&xml),
        )
    }
}
impl FlatOpenDocument {
    pub fn table_column_style_properties(&self) -> Result<TableColumnStylePropertiesSet> {
        parse_table_column_style_properties(self.xml())
    }
}
