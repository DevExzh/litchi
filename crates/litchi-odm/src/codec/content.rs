//! Master-document content validation.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const BODY_MARKER: &str = "<office:text";
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_IDENTITIES: usize = 1_000_000;
const MAX_REFERENCES: usize = 1_000_000;
const MAX_REFERENCE_BYTES: usize = 16 * 1024;
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";

/// Bounded semantic projection retained by the package snapshot.
#[derive(Clone, Debug, Default)]
pub(crate) struct Semantics {
    references: Vec<crate::model::subdocument::Reference>,
}

impl Semantics {
    pub(crate) fn references(&self) -> &[crate::model::subdocument::Reference] {
        &self.references
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
    if !xml.contains(BODY_MARKER) {
        return Err(Error::InvalidFormat(
            "content.xml has no master body".to_string(),
        ));
    }
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
    let mut sections = Vec::new();
    let mut references = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM content XML: {error}")))?;
        let namespace = classify(&namespace);
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
                )?;
            },
            Event::End(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Text && local.as_ref() == b"section" {
                    let section = sections
                        .pop()
                        .ok_or_else(|| invalid("ODM text:section nesting underflow"))?;
                    if section.0 != depth {
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
            Event::Eof => break,
            _ => {},
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
    Ok(Semantics { references })
}

#[allow(clippy::too_many_arguments)]
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
    sections: &mut Vec<(usize, String)>,
    references: &mut Vec<crate::model::subdocument::Reference>,
) -> Result<()> {
    let local = element.local_name();
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
            sections.push((depth, section));
        }
    } else if namespace == NamespaceKind::Text && local.as_ref() == b"section-source" {
        let Some((section_depth, section)) = sections.last() else {
            return Err(invalid("ODM text:section-source is outside a text:section"));
        };
        if depth != section_depth.saturating_add(1) {
            return Err(invalid(
                "ODM text:section-source is not a direct text:section child",
            ));
        }
        let href = attribute(reader, element, XLINK, b"href")?
            .ok_or_else(|| invalid("ODM text:section-source has no xlink:href"))?;
        ensure_short(&href, "ODM xlink:href")?;
        if references.len() >= MAX_REFERENCES {
            return Err(invalid("ODM subdocument reference count exceeds the limit"));
        }
        references
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODM subdocument references",
                source,
            })?;
        references.push(crate::model::subdocument::Reference::new(
            section.clone(),
            href,
        ));
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
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid ODM attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_bound(&resolved, namespace) && name.as_ref() == local {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| invalid(format!("invalid ODM attribute value: {error}")));
        }
    }
    Ok(None)
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODM XML depth overflow"))?;
    if depth > MAX_DEPTH {
        return Err(invalid("ODM XML depth exceeds the limit"));
    }
    Ok(depth)
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Text,
    Other,
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        _ => NamespaceKind::Other,
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
