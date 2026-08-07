//! XML codec for generated `OpenDocument` text indexes.

use super::model::{
    TextIndex, TextIndexAttribute, TextIndexContent, TextIndexElement, TextIndexKind,
};

use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, decode_reference, is_bound, namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub(super) const MAX_INDEX_DEPTH: usize = 4_096;
const MAX_INDEX_ITEMS: usize = 1_000_000;

struct ActiveIndex {
    kind: TextIndexKind,
    stack: Vec<TextIndexElement>,
    order: usize,
    item_count: usize,
}

pub(crate) fn parse_text_indexes(xml: &str) -> Result<Vec<TextIndex>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active = Vec::<ActiveIndex>::new();
    let mut indexes = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid text-index XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref source) => {
                document_depth = checked_depth(document_depth)?;
                let kind = text_element
                    .then(|| index_kind(source.local_name().as_ref()))
                    .flatten();
                if !active.is_empty() || kind.is_some() {
                    let namespace_uri = resolved_namespace(&namespace, "text-index element")?;
                    let node = element_from_start(&reader, namespace_uri, source)?;
                    for index in &mut active {
                        add_index_item(index)?;
                        index.stack.push(node.clone());
                    }
                    if let Some(kind) = kind {
                        validate_index_root(&reader, source)?;
                        if next_order >= MAX_INDEX_ITEMS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_INDEX_ITEMS} text indexes"
                            )));
                        }
                        active.push(ActiveIndex {
                            kind,
                            stack: vec![node],
                            order: next_order,
                            item_count: 1,
                        });
                        next_order += 1;
                    }
                }
            },
            Event::Empty(ref source) => {
                let kind = text_element
                    .then(|| index_kind(source.local_name().as_ref()))
                    .flatten();
                if !active.is_empty() || kind.is_some() {
                    let namespace_uri = resolved_namespace(&namespace, "text-index element")?;
                    let node = element_from_start(&reader, namespace_uri, source)?;
                    for index in &mut active {
                        add_index_item(index)?;
                        index
                            .stack
                            .last_mut()
                            .expect("active index stack")
                            .content
                            .push(TextIndexContent::Element(node.clone()));
                    }
                    if let Some(kind) = kind {
                        validate_index_root(&reader, source)?;
                        if next_order >= MAX_INDEX_ITEMS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_INDEX_ITEMS} text indexes"
                            )));
                        }
                        indexes.push((next_order, TextIndex { kind, root: node }));
                        next_order += 1;
                    }
                }
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text-index text: {error}"))
                    })?;
                append_index_text(&mut active, &value)?;
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text-index CDATA: {error}"))
                    })?;
                append_index_text(&mut active, &value)?;
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                append_index_text(&mut active, &decode_reference(reference, "text index")?)?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("text-index XML stack underflow".to_string())
                })?;
                for position in (0..active.len()).rev() {
                    let node = active[position].stack.pop().expect("active index stack");
                    if let Some(parent) = active[position].stack.last_mut() {
                        parent.content.push(TextIndexContent::Element(node));
                    } else {
                        let finished = active.remove(position);
                        indexes.push((
                            finished.order,
                            TextIndex {
                                kind: finished.kind,
                                root: node,
                            },
                        ));
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete text-index XML structure".to_string(),
        ));
    }
    indexes.sort_by_key(|(order, _)| *order);
    Ok(indexes.into_iter().map(|(_, index)| index).collect())
}

fn index_kind(local_name: &[u8]) -> Option<TextIndexKind> {
    match local_name {
        b"table-of-content" => Some(TextIndexKind::TableOfContents),
        b"illustration-index" => Some(TextIndexKind::Illustration),
        b"table-index" => Some(TextIndexKind::Table),
        b"object-index" => Some(TextIndexKind::Object),
        b"user-index" => Some(TextIndexKind::User),
        b"alphabetical-index" => Some(TextIndexKind::Alphabetical),
        b"bibliography" => Some(TextIndexKind::Bibliography),
        _ => None,
    }
}

fn validate_index_root(reader: &NsReader<&[u8]>, source: &BytesStart<'_>) -> Result<()> {
    namespaced_attribute(reader, source, TEXT_NAMESPACE, b"name", "text index")?
        .ok_or_else(|| Error::InvalidFormat("text index requires text:name".to_string()))?;
    if let Some(value) =
        namespaced_attribute(reader, source, TEXT_NAMESPACE, b"protected", "text index")?
        && !matches!(value.as_str(), "true" | "false" | "1" | "0")
    {
        return Err(Error::InvalidFormat(
            "text:protected must be true, false, 1, or 0".to_string(),
        ));
    }
    Ok(())
}

fn element_from_start(
    reader: &NsReader<&[u8]>,
    namespace_uri: Option<String>,
    source: &BytesStart<'_>,
) -> Result<TextIndexElement> {
    let local_name = std::str::from_utf8(source.local_name().as_ref())
        .map_err(|_| Error::InvalidFormat("non-UTF-8 text-index element name".to_string()))?
        .to_string();
    let attributes = expanded_attributes(reader, source, "text index")?;
    Ok(TextIndexElement {
        namespace_uri,
        local_name,
        attributes,
        content: Vec::new(),
    })
}

pub(crate) fn expanded_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    context: &str,
) -> Result<Vec<TextIndexAttribute>> {
    let mut attributes = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid {context} attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace, context)?;
        let local_name = std::str::from_utf8(local_name.as_ref())
            .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 {context} attribute name")))?
            .to_string();
        if attributes.iter().any(|existing: &TextIndexAttribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded {context} attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid {context} attribute value: {error}"))
            })?
            .into_owned();
        attributes.push(TextIndexAttribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(attributes)
}

fn resolved_namespace(namespace: &ResolveResult<'_>, context: &str) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => std::str::from_utf8(uri)
            .map(|uri| Some(uri.to_string()))
            .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 {context} namespace URI"))),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown {context} namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn append_index_text(active: &mut [ActiveIndex], value: &str) -> Result<()> {
    for index in active {
        let element = index.stack.last_mut().expect("active index stack");
        if let Some(TextIndexContent::Text(text)) = element.content.last_mut() {
            append_checked(text, value)?;
        } else {
            let mut text = String::new();
            append_checked(&mut text, value)?;
            element.content.push(TextIndexContent::Text(text));
            add_index_item(index)?;
        }
    }
    Ok(())
}

fn add_index_item(index: &mut ActiveIndex) -> Result<()> {
    index.item_count = index
        .item_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("text-index item count overflow".to_string()))?;
    if index.item_count > MAX_INDEX_ITEMS {
        return Err(Error::InvalidFormat(format!(
            "text index exceeds {MAX_INDEX_ITEMS} items"
        )));
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("text-index nesting depth overflow".to_string()))?;
    if depth > MAX_INDEX_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "text-index nesting exceeds {MAX_INDEX_DEPTH} levels"
        )));
    }
    Ok(depth)
}
