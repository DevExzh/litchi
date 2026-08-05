//! Bounded table-style XML scanning, semantic comparison, and encoding.

use super::model::{Attr, Conformance, Def, Id, List, PARTS, Parts, Payload};
use super::validation::{validate_name, validate_parsed};
use super::{
    A, AS, MAX_ATTRIBUTE_BYTES, MAX_ATTRIBUTES, MAX_DEPTH, MAX_NODES, MAX_STYLES, MAX_XML_BYTES,
    allocation, invalid, limit, xml_error,
};
use crate::Result;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;

#[derive(Debug, PartialEq, Eq)]
enum SemanticToken {
    Start {
        namespace: String,
        local: String,
        attributes: Vec<(String, String, String)>,
    },
    End {
        namespace: String,
        local: String,
    },
    Text(String),
    Comment(String),
}

struct SemanticCursor<'a> {
    reader: NsReader<&'a [u8]>,
    pending_end: Option<(String, String)>,
    spaces: Vec<Space>,
    nodes: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Space {
    Default,
    Preserve,
}

impl<'a> SemanticCursor<'a> {
    fn new(xml: &'a [u8]) -> Self {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        Self {
            reader,
            pending_end: None,
            spaces: Vec::new(),
            nodes: 0,
        }
    }

    fn next(&mut self) -> Result<Option<SemanticToken>> {
        if let Some((namespace, local)) = self.pending_end.take() {
            let _ = self.spaces.pop();
            return Ok(Some(SemanticToken::End { namespace, local }));
        }
        loop {
            let decoder = self.reader.decoder();
            let event = self.reader.read_event().map_err(xml_error)?.into_owned();
            let resolver = self.reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    bump_node(&mut self.nodes)?;
                    let (namespace, local, attributes) =
                        semantic_element(&resolver, &namespace, &element, decoder)?;
                    self.push_space(&attributes);
                    return Ok(Some(SemanticToken::Start {
                        namespace,
                        local,
                        attributes,
                    }));
                },
                Event::Empty(element) => {
                    bump_node(&mut self.nodes)?;
                    let (namespace, local, attributes) =
                        semantic_element(&resolver, &namespace, &element, decoder)?;
                    self.push_space(&attributes);
                    self.pending_end = Some((namespace.clone(), local.clone()));
                    return Ok(Some(SemanticToken::Start {
                        namespace,
                        local,
                        attributes,
                    }));
                },
                Event::End(element) => {
                    let _ = self.spaces.pop();
                    return Ok(Some(SemanticToken::End {
                        namespace: resolved_namespace(&namespace)?,
                        local: std::str::from_utf8(element.local_name().as_ref())
                            .map_err(xml_error)?
                            .to_owned(),
                    }));
                },
                Event::Text(text) => {
                    let decoded = text.decode().map_err(xml_error)?;
                    let text = quick_xml::escape::unescape(&decoded)
                        .map_err(xml_error)?
                        .into_owned();
                    if !text.chars().all(char::is_whitespace)
                        || self.spaces.len() >= 3
                        || self.spaces.last() == Some(&Space::Preserve)
                    {
                        return Ok(Some(SemanticToken::Text(text)));
                    }
                },
                Event::GeneralRef(reference) => {
                    return Ok(Some(SemanticToken::Text(
                        litchi_ooxml_common::xml::decode_xml_reference(&reference)?,
                    )));
                },
                Event::CData(text) => {
                    return Ok(Some(SemanticToken::Text(
                        text.decode().map_err(xml_error)?.into_owned(),
                    )));
                },
                Event::Comment(comment) => {
                    return Ok(Some(SemanticToken::Comment(
                        comment.decode().map_err(xml_error)?.into_owned(),
                    )));
                },
                Event::Decl(_) => {},
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("forbidden markup in table-style XML"));
                },
                Event::Eof => return Ok(None),
            }
        }
    }

    fn push_space(&mut self, attributes: &[(String, String, String)]) {
        const XML: &str = "http://www.w3.org/XML/1998/namespace";
        let inherited = self.spaces.last().copied().unwrap_or(Space::Default);
        let selected = attributes
            .iter()
            .find(|(namespace, local, _)| namespace == XML && local == "space")
            .map_or(inherited, |(_, _, value)| match value.as_str() {
                "default" => Space::Default,
                "preserve" => Space::Preserve,
                // The catalog scanner deliberately retains opaque attributes.
                // Unknown xml:space values must therefore disable whitespace
                // normalization rather than risk discarding caller data.
                _ => Space::Preserve,
            });
        self.spaces.push(selected);
    }
}

pub(super) fn semantic_xml_eq(left: &[u8], right: &[u8]) -> Result<bool> {
    let mut left = SemanticCursor::new(left);
    let mut right = SemanticCursor::new(right);
    loop {
        let left = left.next()?;
        let right = right.next()?;
        if left != right {
            return Ok(false);
        }
        if left.is_none() {
            return Ok(true);
        }
    }
}

fn semantic_element(
    resolver: &NamespaceResolver,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<(String, String, Vec<(String, String, String)>)> {
    let namespace = resolved_namespace(namespace)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = Vec::new();
    let mut bytes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_ATTRIBUTES {
            return Err(limit("table-style attribute count", MAX_ATTRIBUTES));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?;
        let local = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bytes = bytes
            .checked_add(namespace.len())
            .and_then(|total| total.checked_add(local.len()))
            .and_then(|total| total.checked_add(value.len()))
            .ok_or_else(|| limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES))?;
        if bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        attributes
            .try_reserve(1)
            .map_err(|source| allocation("table-style semantic attributes", source))?;
        attributes.push((namespace, local, value));
    }
    attributes.sort_unstable();
    Ok((namespace, local, attributes))
}

fn resolved_namespace(namespace: &ResolveResult<'_>) -> Result<String> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound table-style XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
#[derive(Debug)]
pub(super) struct Parsed {
    pub(super) conformance: Conformance,
    pub(super) default: Id,
    pub(super) defs: Vec<ParsedDef>,
    pub(super) root_attrs: Vec<Attr>,
}

#[derive(Debug)]
pub(super) struct ParsedDef {
    pub(super) id: Id,
    pub(super) name: String,
    pub(super) parts: Parts,
    pub(super) attrs: Vec<Attr>,
    pub(super) raw: Range<usize>,
    pub(super) body: Range<usize>,
}

struct OpenDef {
    id: Id,
    name: String,
    parts: Parts,
    attrs: Vec<Attr>,
    raw_start: usize,
    body_start: usize,
    extension: bool,
    last_child: Option<usize>,
}

pub(super) fn parse_owned(source: Vec<u8>) -> Result<List> {
    let parsed = scan(&source)?;
    let mut defs = Vec::new();
    defs.try_reserve_exact(parsed.defs.len())
        .map_err(|source| allocation("table-style index", source))?;
    for style in parsed.defs {
        defs.push(Def {
            id: style.id,
            name: style.name,
            parts: style.parts,
            attrs: style.attrs,
            payload: Payload::Shared {
                raw: style.raw,
                body: style.body,
                exact: true,
            },
        });
    }
    Ok(List {
        conformance: parsed.conformance,
        default: parsed.default,
        defs,
        root_attrs: parsed.root_attrs,
        source,
        dirty: false,
    })
}

pub(super) fn scan(xml: &[u8]) -> Result<Parsed> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table-style XML bytes", MAX_XML_BYTES));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut conformance = None;
    let mut default = None;
    let mut root_attrs = Vec::new();
    let mut defs = Vec::new();
    let mut open = None;

    loop {
        let start = xml_position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = xml_position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                bump_node(&mut nodes)?;
                depth = checked_depth(depth)?;
                if depth == 1 {
                    if saw_root || element.local_name().as_ref() != b"tblStyleLst" {
                        return Err(invalid(
                            "table-style part must contain one tblStyleLst root",
                        ));
                    }
                    let profile = drawing_conformance(&namespace)
                        .ok_or_else(|| invalid("table-style root has the wrong namespace"))?;
                    let (id, attrs) = parse_root_attrs(&element, decoder)?;
                    saw_root = true;
                    conformance = Some(profile);
                    default = Some(id);
                    root_attrs = attrs;
                } else if depth == 2 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyle")?;
                    if defs.len() >= MAX_STYLES {
                        return Err(limit("table-style count", MAX_STYLES));
                    }
                    let (id, name, attrs) = parse_def_attrs(&element, decoder)?;
                    open = Some(OpenDef {
                        id,
                        name,
                        parts: Parts::empty(),
                        attrs,
                        raw_start: start,
                        body_start: end,
                        extension: false,
                        last_child: None,
                    });
                } else if depth == 3 {
                    record_part(&namespace, conformance, element.name(), &mut open)?;
                }
            },
            Event::Empty(element) => {
                bump_node(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("table-style XML depth", MAX_DEPTH))?;
                if child_depth > MAX_DEPTH {
                    return Err(limit("table-style XML depth", MAX_DEPTH));
                }
                if child_depth == 1 {
                    if saw_root || element.local_name().as_ref() != b"tblStyleLst" {
                        return Err(invalid(
                            "table-style part must contain one tblStyleLst root",
                        ));
                    }
                    let profile = drawing_conformance(&namespace)
                        .ok_or_else(|| invalid("table-style root has the wrong namespace"))?;
                    let (id, attrs) = parse_root_attrs(&element, decoder)?;
                    saw_root = true;
                    closed_root = true;
                    conformance = Some(profile);
                    default = Some(id);
                    root_attrs = attrs;
                } else if child_depth == 2 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyle")?;
                    if defs.len() >= MAX_STYLES {
                        return Err(limit("table-style count", MAX_STYLES));
                    }
                    let (id, name, attrs) = parse_def_attrs(&element, decoder)?;
                    defs.try_reserve(1)
                        .map_err(|source| allocation("table-style parse index", source))?;
                    defs.push(ParsedDef {
                        id,
                        name,
                        parts: Parts::empty(),
                        attrs,
                        raw: start..end,
                        body: end..end,
                    });
                } else if child_depth == 3 {
                    record_part(&namespace, conformance, element.name(), &mut open)?;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("table-style XML nesting underflow"));
                }
                if depth == 2 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyle")?;
                    let style = open
                        .take()
                        .ok_or_else(|| invalid("table-style closing element has no start"))?;
                    if style.body_start > start || style.raw_start > style.body_start {
                        return Err(invalid("table-style source ranges are invalid"));
                    }
                    defs.try_reserve(1)
                        .map_err(|source| allocation("table-style parse index", source))?;
                    defs.push(ParsedDef {
                        id: style.id,
                        name: style.name,
                        parts: style.parts,
                        attrs: style.attrs,
                        raw: style.raw_start..end,
                        body: style.body_start..start,
                    });
                } else if depth == 1 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyleLst")?;
                    closed_root = true;
                }
                depth -= 1;
            },
            Event::Text(text) => {
                if depth <= 2
                    && text
                        .decode()
                        .map_err(xml_error)?
                        .chars()
                        .any(|value| !value.is_whitespace())
                {
                    return Err(invalid(
                        "table-style root or definition contains text content",
                    ));
                }
            },
            Event::GeneralRef(_) if depth <= 2 => {
                return Err(invalid(
                    "table-style root or definition contains a character reference",
                ));
            },
            Event::CData(_) => return Err(invalid("table-style XML must not contain CDATA")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "table-style XML must not contain a DTD or processing instruction",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
            _ => {},
        }
    }
    if !saw_root || !closed_root || depth != 0 || open.is_some() {
        return Err(invalid("table-style XML root is missing or unterminated"));
    }
    let parsed = Parsed {
        conformance: conformance.ok_or_else(|| invalid("missing table-style conformance"))?,
        default: default.ok_or_else(|| invalid("table-style default GUID is required"))?,
        defs,
        root_attrs,
    };
    validate_parsed(&parsed)?;
    Ok(parsed)
}

fn parse_root_attrs(element: &BytesStart<'_>, decoder: Decoder) -> Result<(Id, Vec<Attr>)> {
    let mut default = None;
    let mut extras = Vec::new();
    for (name, value) in attributes(element, decoder)? {
        if name == "def" {
            if default.replace(Id::parse(&value)?).is_some() {
                return Err(invalid("table-style root declares def twice"));
            }
        } else {
            extras
                .try_reserve(1)
                .map_err(|source| allocation("table-style root attributes", source))?;
            extras.push(Attr { name, value });
        }
    }
    Ok((
        default.ok_or_else(|| invalid("table-style default GUID is required"))?,
        extras,
    ))
}

fn parse_def_attrs(element: &BytesStart<'_>, decoder: Decoder) -> Result<(Id, String, Vec<Attr>)> {
    let mut id = None;
    let mut name = None;
    let mut extras = Vec::new();
    for (attribute, value) in attributes(element, decoder)? {
        match attribute.as_str() {
            "styleId" => {
                if id.replace(Id::parse(&value)?).is_some() {
                    return Err(invalid("table style declares styleId twice"));
                }
            },
            "styleName" => {
                validate_name(&value)?;
                if name.replace(value).is_some() {
                    return Err(invalid("table style declares styleName twice"));
                }
            },
            _ => {
                extras
                    .try_reserve(1)
                    .map_err(|source| allocation("table-style attributes", source))?;
                extras.push(Attr {
                    name: attribute,
                    value,
                });
            },
        }
    }
    Ok((
        id.ok_or_else(|| invalid("table style requires a styleId GUID"))?,
        name.ok_or_else(|| invalid("table style requires a styleName attribute"))?,
        extras,
    ))
}

fn attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<Vec<(String, String)>> {
    let mut output = Vec::new();
    let mut bytes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if output.len() >= MAX_ATTRIBUTES {
            return Err(limit("table-style attribute count", MAX_ATTRIBUTES));
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bytes = bytes
            .checked_add(name.len())
            .and_then(|value_len| value_len.checked_add(value.len()))
            .ok_or_else(|| limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES))?;
        if bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        output
            .try_reserve(1)
            .map_err(|source| allocation("table-style attribute decoding", source))?;
        output.push((name, value));
    }
    Ok(output)
}

fn record_part(
    namespace: &ResolveResult<'_>,
    conformance: Option<Conformance>,
    name: QName<'_>,
    open: &mut Option<OpenDef>,
) -> Result<()> {
    let profile = conformance.ok_or_else(|| invalid("table-style root profile is missing"))?;
    if !is_drawing(namespace, profile) {
        return Err(invalid(
            "table-style region uses the wrong DrawingML namespace",
        ));
    }
    let style = open
        .as_mut()
        .ok_or_else(|| invalid("table-style region has no owning definition"))?;
    if name.local_name().as_ref() == b"extLst" {
        if style.extension {
            return Err(invalid("table style declares extLst twice"));
        }
        style.extension = true;
        return Ok(());
    }
    if style.extension {
        return Err(invalid("table-style region appears after extLst"));
    }
    let part = Parts::from_xml_name(name.local_name().as_ref())
        .ok_or_else(|| invalid("unexpected direct child in table-style definition"))?;
    if style.parts.intersects(part) {
        return Err(invalid("table style declares a conditional region twice"));
    }
    let order = PARTS
        .iter()
        .position(|(candidate, _)| *candidate == part)
        .ok_or_else(|| invalid("table-style region has no schema order"))?;
    if style.last_child.is_some_and(|previous| previous >= order) {
        return Err(invalid(
            "table-style regions do not follow the schema sequence",
        ));
    }
    style.last_child = Some(order);
    style.parts.insert(part);
    Ok(())
}

fn drawing_conformance(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == A.as_bytes() => {
            Some(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == AS.as_bytes() => {
            Some(Conformance::Strict)
        },
        _ => None,
    }
}

fn is_drawing(namespace: &ResolveResult<'_>, conformance: Conformance) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == conformance.drawing().as_bytes())
}

fn require_drawing(
    namespace: &ResolveResult<'_>,
    conformance: Conformance,
    actual: QName<'_>,
    expected: &[u8],
) -> Result<()> {
    if is_drawing(namespace, conformance) && actual.local_name().as_ref() == expected {
        Ok(())
    } else {
        Err(invalid("unexpected table-style element or namespace"))
    }
}

pub(super) fn encode(list: &List) -> Result<Vec<u8>> {
    let mut output = String::new();
    output
        .try_reserve(list.source.len().clamp(256, MAX_XML_BYTES))
        .map_err(|source| allocation("table-style XML encoding", source))?;
    append(
        &mut output,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a=""#,
    )?;
    append(&mut output, list.conformance.drawing())?;
    append(&mut output, "\"")?;
    for attribute in &list.root_attrs {
        if attribute.name == "xmlns:a" || attribute.name == "def" {
            continue;
        }
        append_attribute(&mut output, &attribute.name, &attribute.value)?;
    }
    append(&mut output, " def=\"")?;
    list.default.write_to(&mut output)?;
    if list.defs.is_empty() {
        append(&mut output, "\"/>")?;
    } else {
        append(&mut output, "\">")?;
        for style in &list.defs {
            encode_def(&mut output, style, &list.source)?;
        }
        append(&mut output, "</a:tblStyleLst>")?;
    }
    Ok(output.into_bytes())
}

fn encode_def(output: &mut String, style: &Def, source: &[u8]) -> Result<()> {
    let (raw, body, exact) = payload(style, source)?;
    if exact {
        return append_bytes(output, raw);
    }
    append(output, "<a:tblStyle styleId=\"")?;
    style.id.write_to(output)?;
    append(output, "\" styleName=\"")?;
    append_escaped(output, &style.name)?;
    append(output, "\"")?;
    for attribute in &style.attrs {
        if matches!(attribute.name.as_str(), "styleId" | "styleName") {
            continue;
        }
        append_attribute(output, &attribute.name, &attribute.value)?;
    }
    if body.is_empty() && style.parts.is_empty() {
        return append(output, "/>");
    }
    append(output, ">")?;
    if body.is_empty() {
        for (part, name) in PARTS {
            if style.parts.contains(part) {
                append(output, "<a:")?;
                append(output, name)?;
                append(output, "/>")?;
            }
        }
    } else {
        append_bytes(output, body)?;
    }
    append(output, "</a:tblStyle>")
}

fn payload<'a>(style: &'a Def, source: &'a [u8]) -> Result<(&'a [u8], &'a [u8], bool)> {
    match &style.payload {
        Payload::Shared { raw, body, exact } => Ok((
            source
                .get(raw.clone())
                .ok_or_else(|| invalid("table-style raw source range is invalid"))?,
            source
                .get(body.clone())
                .ok_or_else(|| invalid("table-style body source range is invalid"))?,
            *exact,
        )),
        Payload::Owned { xml, body, exact } => Ok((
            xml,
            xml.get(body.clone())
                .ok_or_else(|| invalid("detached table-style body range is invalid"))?,
            *exact,
        )),
    }
}

fn append_attribute(output: &mut String, name: &str, value: &str) -> Result<()> {
    append(output, " ")?;
    append(output, name)?;
    append(output, "=\"")?;
    append_escaped(output, value)?;
    append(output, "\"")
}

fn append_escaped(output: &mut String, value: &str) -> Result<()> {
    let escaped = quick_xml::escape::escape(value);
    append(output, escaped.as_ref())
}

fn append_bytes(output: &mut String, value: &[u8]) -> Result<()> {
    append(output, std::str::from_utf8(value).map_err(xml_error)?)
}

fn append(output: &mut String, value: &str) -> Result<()> {
    let length = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| limit("encoded table-style XML bytes", MAX_XML_BYTES))?;
    if length > MAX_XML_BYTES {
        return Err(limit("encoded table-style XML bytes", MAX_XML_BYTES));
    }
    output
        .try_reserve(value.len())
        .map_err(|source| allocation("table-style XML encoding", source))?;
    output.push_str(value);
    Ok(())
}

pub(super) fn xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| limit("table-style XML depth", MAX_DEPTH))?;
    if depth > MAX_DEPTH {
        Err(limit("table-style XML depth", MAX_DEPTH))
    } else {
        Ok(depth)
    }
}

fn bump_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("table-style XML nodes", MAX_NODES))?;
    if *nodes > MAX_NODES {
        Err(limit("table-style XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("table-style XML offset exceeds usize"))
}
