//! Master-document content validation.

use litchi_core::{Error, Result};
use litchi_odf_common::compact_xml;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{borrow::Cow, collections::HashSet, ops::Range};

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_IDENTITIES: usize = 1_000_000;
const MAX_REFERENCES: usize = 1_000_000;
const MAX_REFERENCE_BYTES: usize = 16 * 1024;
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Text,
    Other,
}

#[derive(Clone, Debug)]
struct ActiveSection {
    depth: usize,
    name: String,
    child_seen: bool,
    source_seen: bool,
}

/// Bounded semantic projection retained by the package snapshot.
#[derive(Clone, Debug, Default)]
pub(crate) struct Semantics {
    references: Vec<crate::model::subdocument::Reference>,
    href_spans: Vec<Range<usize>>,
}

impl Semantics {
    pub(crate) fn references(&self) -> &[crate::model::subdocument::Reference] {
        &self.references
    }

    pub(crate) fn href_span(&self, reference: usize) -> Option<&Range<usize>> {
        self.href_spans.get(reference)
    }
}

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    parse(xml).map(|_| ())
}

/// Validates content and projects ordered, inert subdocument references.
pub(crate) fn parse(xml: &str) -> Result<Semantics> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "content.xml exceeds the family limit".to_string(),
        ));
    }
    compact_xml::validate(xml.as_bytes()).map_err(Error::from)?;
    parse_structure(xml)
}

fn parse_structure(xml: &str) -> Result<Semantics> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut body_seen = false;
    let mut master_seen = false;
    let mut body_depth = None;
    let mut master_depth = None;
    let mut section_names = HashSet::new();
    let mut xml_ids = HashSet::new();
    let mut sections: Vec<ActiveSection> = Vec::new();
    let mut references = Vec::new();
    let mut href_spans = Vec::new();
    let mut section_source_depth = None;
    loop {
        let event_start = position(&reader)?;
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM content XML: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let event = borrowed_event.into_owned();
        let event_end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                observe(
                    &reader,
                    namespace,
                    &element,
                    depth,
                    false,
                    &mut root_seen,
                    &mut body_seen,
                    &mut master_seen,
                    &mut body_depth,
                    &mut master_depth,
                    &mut section_names,
                    &mut xml_ids,
                    &mut sections,
                    &mut references,
                    &mut href_spans,
                    &mut section_source_depth,
                    xml.as_bytes()
                        .get(event_start..event_end)
                        .ok_or_else(|| invalid("ODM XML event span is outside content.xml"))?,
                    event_start,
                )?;
            },
            Event::Empty(element) => {
                let virtual_depth = checked_depth(depth)?;
                observe(
                    &reader,
                    namespace,
                    &element,
                    virtual_depth,
                    true,
                    &mut root_seen,
                    &mut body_seen,
                    &mut master_seen,
                    &mut body_depth,
                    &mut master_depth,
                    &mut section_names,
                    &mut xml_ids,
                    &mut sections,
                    &mut references,
                    &mut href_spans,
                    &mut section_source_depth,
                    xml.as_bytes()
                        .get(event_start..event_end)
                        .ok_or_else(|| invalid("ODM XML event span is outside content.xml"))?,
                    event_start,
                )?;
            },
            Event::End(element) => {
                let local = element.local_name();
                if section_source_depth == Some(depth) {
                    section_source_depth = None;
                }
                if namespace == NamespaceKind::Text && local.as_ref() == b"section" {
                    let section = sections
                        .pop()
                        .ok_or_else(|| invalid("ODM text:section nesting underflow"))?;
                    if section.depth != depth {
                        return Err(invalid("ODM text:section nesting is malformed"));
                    }
                }
                if master_depth == Some(depth) {
                    master_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODM XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM content")),
            Event::GeneralRef(_) => {
                return Err(invalid("named XML entities are not allowed in ODM content"));
            },
            Event::Text(text) if section_source_depth.is_some() && !text.as_ref().is_empty() => {
                return Err(invalid("ODM text:section-source cannot contain text"));
            },
            Event::CData(text) if section_source_depth.is_some() && !text.as_ref().is_empty() => {
                return Err(invalid("ODM text:section-source cannot contain CDATA"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if depth != 0 || !root_seen || !body_seen || !master_seen {
        return Err(invalid(
            "ODM content has an incomplete master-document structure",
        ));
    }
    if !sections.is_empty() {
        return Err(invalid("ODM text:section nesting is incomplete"));
    }
    Ok(Semantics {
        references,
        href_spans,
    })
}

#[allow(
    clippy::too_many_arguments,
    reason = "one XML event updates the bounded ODM semantic projection"
)]
fn observe(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    root_seen: &mut bool,
    body_seen: &mut bool,
    master_seen: &mut bool,
    body_depth: &mut Option<usize>,
    master_depth: &mut Option<usize>,
    section_names: &mut HashSet<String>,
    xml_ids: &mut HashSet<String>,
    sections: &mut Vec<ActiveSection>,
    references: &mut Vec<crate::model::subdocument::Reference>,
    href_spans: &mut Vec<Range<usize>>,
    section_source_depth: &mut Option<usize>,
    tag: &[u8],
    tag_start: usize,
) -> Result<()> {
    let local = element.local_name();
    if section_source_depth.is_some() {
        return Err(invalid("ODM text:section-source cannot contain elements"));
    }
    let is_section_source = namespace == NamespaceKind::Text && local.as_ref() == b"section-source";
    let is_dde_source = namespace == NamespaceKind::Office && local.as_ref() == b"dde-source";
    if let Some(parent) = sections.last_mut()
        && depth == parent.depth.saturating_add(1)
    {
        if is_section_source || is_dde_source {
            if parent.child_seen || parent.source_seen {
                return Err(invalid(
                    "ODM linked-section source must occur once as the first section child",
                ));
            }
            parent.source_seen = true;
        } else {
            parent.child_seen = true;
        }
    }
    if depth == 1 {
        if *root_seen
            || namespace != NamespaceKind::Office
            || local.as_ref() != b"document-content"
            || empty
        {
            return Err(invalid(
                "ODM content requires one office:document-content root",
            ));
        }
        *root_seen = true;
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
        if *body_seen || depth != 2 || empty {
            return Err(invalid("ODM content requires one non-empty office:body"));
        }
        *body_seen = true;
        *body_depth = Some(depth);
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"text" {
        if *master_seen || *body_depth != Some(depth - 1) {
            return Err(invalid("ODM office:text body is misplaced or duplicated"));
        }
        *master_seen = true;
        if !empty {
            *master_depth = Some(depth);
        }
    } else if namespace == NamespaceKind::Text && local.as_ref() == b"section" {
        if master_depth.is_none() {
            return Err(invalid("ODM text:section is outside the master body"));
        }
        let name = attribute(reader, element, TEXT, b"name")?
            .ok_or_else(|| invalid("ODM text:section has no text:name"))?;
        ensure_short(&name, "ODM text:section name")?;
        let section = name.clone();
        insert_identity(section_names, name, "duplicate ODM section name")?;
        if !empty {
            sections
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM section nesting",
                    source,
                })?;
            sections.push(ActiveSection {
                depth,
                name: section,
                child_seen: false,
                source_seen: false,
            });
        }
    } else if is_section_source {
        let Some(section) = sections.last() else {
            return Err(invalid("ODM text:section-source is outside a text:section"));
        };
        if depth != section.depth.saturating_add(1) {
            return Err(invalid(
                "ODM text:section-source is not a direct text:section child",
            ));
        }
        let href = attribute(reader, element, XLINK, b"href")?;
        let link_type = attribute(reader, element, XLINK, b"type")?;
        let show = attribute(reader, element, XLINK, b"show")?;
        if link_type.as_deref().is_some_and(|value| value != "simple") {
            return Err(invalid("ODM text:section-source xlink:type must be simple"));
        }
        if show.as_deref().is_some_and(|value| value != "embed") {
            return Err(invalid("ODM text:section-source xlink:show must be embed"));
        }
        if href.is_none() && (link_type.is_some() || show.is_some()) {
            return Err(invalid(
                "ODM text:section-source link attributes require xlink:href",
            ));
        }
        let source_section = attribute(reader, element, TEXT, b"section-name")?;
        let filter_name = attribute(reader, element, TEXT, b"filter-name")?;
        for (optional_value, scope) in [
            (source_section.as_deref(), "ODM source section name"),
            (filter_name.as_deref(), "ODM source filter name"),
        ] {
            if let Some(semantic_value) = optional_value {
                ensure_short(semantic_value, scope)?;
            }
        }
        if let Some(link_href) = href {
            ensure_short(&link_href, "ODM xlink:href")?;
            if references.len() >= MAX_REFERENCES {
                return Err(invalid("ODM subdocument reference count exceeds the limit"));
            }
            references
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM subdocument references",
                    source,
                })?;
            href_spans
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM subdocument reference spans",
                    source,
                })?;
            let href_key = attribute_key(reader, element, XLINK, b"href")?
                .ok_or_else(|| invalid("ODM xlink:href source spelling disappeared"))?;
            let (span_start, span_end) = attribute_value_span(tag, &href_key)?;
            references.push(crate::model::subdocument::Reference::new(
                section.name.clone(),
                link_href,
                source_section,
                filter_name,
            ));
            href_spans.push(tag_start + span_start..tag_start + span_end);
        }
        if !empty {
            *section_source_depth = Some(depth);
        }
    }
    if let Some(xml_id) = attribute(reader, element, XML, b"id")? {
        ensure_short(&xml_id, "ODM xml:id")?;
        insert_identity(xml_ids, xml_id, "duplicate ODM xml:id")?;
    }
    Ok(())
}

fn insert_identity(set: &mut HashSet<String>, value: String, message: &str) -> Result<()> {
    if set.len() >= MAX_IDENTITIES {
        return Err(invalid("ODM identity count exceeds the limit"));
    }
    set.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODM identities",
        source,
    })?;
    if !set.insert(value) {
        return Err(invalid(message));
    }
    Ok(())
}

fn ensure_short(value: &str, scope: &str) -> Result<()> {
    if value.len() > MAX_REFERENCE_BYTES {
        return Err(invalid(format!(
            "{scope} exceeds the {MAX_REFERENCE_BYTES}-byte limit"
        )));
    }
    Ok(())
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut found_value = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid ODM attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_bound(&resolved, namespace) && name.as_ref() == local {
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(Cow::into_owned)
                .map_err(|error| invalid(format!("invalid ODM attribute value: {error}")))?;
            if found_value.replace(decoded).is_some() {
                return Err(invalid("duplicate namespace-equivalent ODM attribute"));
            }
        }
    }
    Ok(found_value)
}

fn attribute_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<Vec<u8>>> {
    let mut key = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid ODM attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_bound(&resolved, namespace)
            && name.as_ref() == local
            && key.replace(attribute.key.as_ref().to_vec()).is_some()
        {
            return Err(invalid("duplicate namespace-equivalent ODM attribute"));
        }
    }
    Ok(key)
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("ODM attribute is missing '='"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid("ODM attribute value is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(invalid("ODM attribute value is unterminated"));
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    Err(invalid("ODM attribute source span is missing"))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_range_error| invalid("ODM XML source position exceeds the platform range"))
}

fn checked_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODM XML depth overflow"))?;
    if next_depth > MAX_DEPTH {
        return Err(invalid("ODM XML depth exceeds the limit"));
    }
    Ok(next_depth)
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        let text = concat!(
            r#"<office:document-content xmlns:office="#,
            r#""urn:oasis:names:tc:opendocument:xmlns:office:1.0">"#,
            "<office:body><office:text/></office:body>",
            "</office:document-content>",
        );
        let chart = concat!(
            r#"<office:document-content xmlns:office="#,
            r#""urn:oasis:names:tc:opendocument:xmlns:office:1.0">"#,
            "<office:body><office:chart/></office:body>",
            "</office:document-content>",
        );
        assert!(validate(text).is_ok());
        assert!(validate(chart).is_err());
    }
}
