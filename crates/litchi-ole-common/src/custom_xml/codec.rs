use std::collections::HashSet;
use std::io::{Read, Seek};
use std::ops::Range;
use std::sync::Arc;

use litchi_cfb::{OleError, OleFile, OleWriter};
use litchi_core::xml::ReaderOrigin;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::model::{
    Error, Item, ItemId, Limits, Promotion, Properties, Result, Store, invalid, limit,
};
use super::xml::{
    escape_attribute, is_whitespace, normalize_encoding, reject_other_attributes,
    required_attribute, resolved_element, restore_encoding, validate_characters, validate_payload,
    validate_reference,
};
use super::{
    CUSTOM_XML_NAMESPACE, ITEM_STREAM, MODIFIED_PROMOTION_STORAGE, PROPERTIES_STREAM,
    REDUNDANT_PROMOTION_STORAGE, STORE_STORAGE,
};
use crate::xml_attributes::BytesStartExt as _;

#[derive(Default)]
struct PropertiesParseState {
    depth: usize,
    root_seen: bool,
    root_closed: bool,
    schema_refs_seen: bool,
    opaque_depth: usize,
    item_id: Option<ItemId>,
    schema_references: Vec<String>,
    element_count: usize,
}

/// Inspect a complete Custom XML data store under bounded resource limits
/// ([`Limits::default`] is the default profile).
///
/// # Errors
///
/// Returns an error if the compound file or Custom XML store is malformed, or
/// if it exceeds `limits`.
pub fn inspect<R: Read + Seek>(ole: &mut OleFile<R>, limits: Limits) -> Result<Option<Store>> {
    validate_limits(&limits)?;
    let redundant = marker_exists(ole, REDUNDANT_PROMOTION_STORAGE)?;
    let modified = marker_exists(ole, MODIFIED_PROMOTION_STORAGE)?;
    if redundant && modified {
        return Err(invalid(
            "redundant and modified promotion storages are both present",
        ));
    }
    let promotion = if redundant {
        Promotion::Redundant
    } else if modified {
        Promotion::Modified
    } else {
        Promotion::Unspecified
    };

    if !ole.directory_exists(&[STORE_STORAGE]) {
        if ole.exists(&[STORE_STORAGE]) {
            return Err(invalid("MsoDataStore is not a storage"));
        }
        if promotion != Promotion::Unspecified {
            return Err(invalid("promotion marker exists without MsoDataStore"));
        }
        return Ok(None);
    }

    let mut names = Vec::new();
    ole.visit_directory_entries(&[STORE_STORAGE], limits.max_items, |ole, sid| {
        let entry = ole
            .directory_entry_by_sid(sid)
            .ok_or_else(|| invalid("MsoDataStore child SID is not present"))?;
        if entry.entry_type != 1 {
            return Err(invalid(format!(
                "MsoDataStore child '{}' is not a storage",
                entry.name
            )));
        }
        validate_storage_name(&entry.name)?;
        names
            .try_reserve_exact(1)
            .map_err(|_| limit("custom XML item-name allocation exceeds its limit"))?;
        names.push(entry.name.clone());
        Ok(())
    })
    .map_err(map_ole_limit)?;
    names.sort();

    let mut total_bytes = 0usize;
    let mut item_ids = HashSet::with_capacity(names.len());
    let mut items = Vec::with_capacity(names.len());
    for storage_name in names {
        let mut item_stream_size = None;
        let mut properties_stream_size = None;
        let mut entry_count = 0usize;
        ole.visit_directory_entry_refs(&[STORE_STORAGE, storage_name.as_str()], 3, |entry| {
            entry_count = entry_count
                .checked_add(1)
                .ok_or_else(|| limit("custom XML stream-entry count overflows usize"))?;
            match (entry.name.as_str(), entry.entry_type) {
                (ITEM_STREAM, 2) => {
                    item_stream_size = Some(stream_size(entry.size, "Item", &limits)?);
                },
                (PROPERTIES_STREAM, 2) => {
                    properties_stream_size = Some(stream_size(entry.size, "Properties", &limits)?);
                },
                _ => {
                    return Err(invalid(format!(
                        "custom XML sub-storage '{storage_name}' has unexpected entry '{}'",
                        entry.name
                    )));
                },
            }
            Ok(())
        })
        .map_err(map_ole_limit)?;
        if entry_count != 2 {
            return Err(invalid(format!(
                "custom XML sub-storage '{storage_name}' must contain exactly Item and Properties"
            )));
        }
        let item_bytes =
            item_stream_size.ok_or_else(|| invalid("custom XML Item stream is missing"))?;
        let properties_bytes = properties_stream_size
            .ok_or_else(|| invalid("custom XML Properties stream is missing"))?;
        if item_bytes > limits.max_item_bytes {
            return Err(limit(format!(
                "Item stream has {item_bytes} bytes, limit is {}",
                limits.max_item_bytes
            )));
        }
        if properties_bytes > limits.max_properties_bytes {
            return Err(limit(format!(
                "Properties stream has {properties_bytes} bytes, limit is {}",
                limits.max_properties_bytes
            )));
        }
        total_bytes = total_bytes
            .checked_add(item_bytes)
            .and_then(|accumulated| accumulated.checked_add(properties_bytes))
            .ok_or_else(|| limit("aggregate stream size overflows usize"))?;
        if total_bytes > limits.max_total_bytes {
            return Err(limit(format!(
                "aggregate stream bytes exceed {}",
                limits.max_total_bytes
            )));
        }
        let xml = ole.open_stream(&[STORE_STORAGE, storage_name.as_str(), ITEM_STREAM])?;
        let properties_xml =
            ole.open_stream(&[STORE_STORAGE, storage_name.as_str(), PROPERTIES_STREAM])?;
        let root_name = validate_payload(&xml, &limits)?;
        let properties = parse_properties_with_limits(&properties_xml, &limits)?;
        if !item_ids.insert(properties.item_id) {
            return Err(invalid(format!(
                "itemID {} is used by multiple custom XML items",
                properties.item_id
            )));
        }
        items.push(Item {
            storage_name,
            xml: Arc::from(xml),
            root_name,
            properties_xml: Arc::from(properties_xml),
            properties,
        });
    }
    Ok(Some(Store { promotion, items }))
}

/// Materialize a validated store in a newly assembled OLE writer under the
/// caller-selected Custom XML policy.
///
/// # Errors
///
/// Returns an error if `store` is invalid or the writer rejects an OLE entry.
pub fn write(writer: &mut OleWriter, store: &Store, limits: &Limits) -> Result<()> {
    validate_store(store, limits)?;
    writer.create_storage(&[STORE_STORAGE])?;
    match store.promotion {
        Promotion::Unspecified => {},
        Promotion::Redundant => {
            writer.create_storage(&[REDUNDANT_PROMOTION_STORAGE])?;
        },
        Promotion::Modified => {
            writer.create_storage(&[MODIFIED_PROMOTION_STORAGE])?;
        },
    }
    for item in &store.items {
        writer.create_storage(&[STORE_STORAGE, item.storage_name.as_str()])?;
        writer.create_stream(
            &[STORE_STORAGE, item.storage_name.as_str(), ITEM_STREAM],
            &item.xml,
        )?;
        writer.create_stream(
            &[STORE_STORAGE, item.storage_name.as_str(), PROPERTIES_STREAM],
            &item.properties_xml,
        )?;
    }
    Ok(())
}

/// Parse the schema-defined Custom XML data-store Properties stream under the
/// caller-selected module policy.
///
/// # Errors
///
/// Returns an error if `xml` is not a valid bounded Properties document.
pub fn parse_properties(xml: &[u8], limits: &Limits) -> Result<Properties> {
    parse_properties_with_limits(xml, limits)
}

/// Serialize Custom XML data-store Properties in stable schema order under the
/// caller-selected module policy.
///
/// # Errors
///
/// Returns an error if `properties` violates the supported format limits.
pub fn write_properties(properties: &Properties, limits: &Limits) -> Result<Vec<u8>> {
    write_properties_with_limits(properties, limits)
}

pub(crate) fn write_properties_with_limits(
    properties: &Properties,
    limits: &Limits,
) -> Result<Vec<u8>> {
    validate_properties(properties, limits)?;
    let capacity = serialized_properties_len(properties)?;
    if capacity > limits.max_properties_bytes {
        return Err(limit("serialized Properties XML exceeds its byte limit"));
    }
    let item_id = properties.item_id.to_string();
    let mut output = Vec::with_capacity(capacity);
    output.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    output.extend_from_slice(b"<ds:datastoreItem xmlns:ds=\"");
    output.extend_from_slice(CUSTOM_XML_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\" ds:itemID=\"");
    output.extend_from_slice(item_id.as_bytes());
    output.extend_from_slice(b"\">");
    if !properties.schema_references.is_empty() {
        output.extend_from_slice(b"<ds:schemaRefs>");
        for uri in &properties.schema_references {
            output.extend_from_slice(b"<ds:schemaRef ds:uri=\"");
            escape_attribute(&mut output, uri);
            output.extend_from_slice(b"\"/>");
        }
        output.extend_from_slice(b"</ds:schemaRefs>");
    }
    output.extend_from_slice(b"</ds:datastoreItem>");
    debug_assert_eq!(output.len(), capacity);
    Ok(output)
}

fn serialized_properties_len(properties: &Properties) -> Result<usize> {
    let mut length = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        .len()
        .checked_add(b"<ds:datastoreItem xmlns:ds=\"".len())
        .and_then(|value| value.checked_add(CUSTOM_XML_NAMESPACE.len()))
        .and_then(|value| value.checked_add(b"\" ds:itemID=\"".len()))
        .and_then(|value| value.checked_add(38))
        .and_then(|value| value.checked_add(b"\">".len()))
        .ok_or_else(|| limit("serialized Properties XML size overflows usize"))?;
    if !properties.schema_references.is_empty() {
        length = length
            .checked_add(b"<ds:schemaRefs>".len())
            .ok_or_else(|| limit("serialized Properties XML size overflows usize"))?;
        for uri in &properties.schema_references {
            length = length
                .checked_add(b"<ds:schemaRef ds:uri=\"".len())
                .ok_or_else(|| limit("serialized Properties XML size overflows usize"))?;
            length = length
                .checked_add(escaped_attribute_len(uri)?)
                .and_then(|value| value.checked_add(b"\"/>".len()))
                .ok_or_else(|| limit("serialized Properties XML size overflows usize"))?;
        }
        length = length
            .checked_add(b"</ds:schemaRefs>".len())
            .ok_or_else(|| limit("serialized Properties XML size overflows usize"))?;
    }
    length
        .checked_add(b"</ds:datastoreItem>".len())
        .ok_or_else(|| limit("serialized Properties XML size overflows usize"))
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    value
        .bytes()
        .try_fold(0usize, |length, byte| {
            length.checked_add(match byte {
                b'&' => 5,
                b'<' => 4,
                b'"' => 6,
                b'\t' | b'\n' | b'\r' => 5,
                _ => 1,
            })
        })
        .ok_or_else(|| limit("escaped Properties XML attribute size overflows usize"))
}

/// Rewrite known properties attributes while retaining the source XML around
/// them. Schema-reference count changes splice only the affected
/// schemaRef elements, so extension attributes, comments, whitespace, and
/// namespace declarations remain source-owned.
pub(crate) fn rewrite_properties(
    source: &[u8],
    before: &Properties,
    after: &Properties,
    limits: &Limits,
) -> Result<Vec<u8>> {
    if before == after {
        return Ok(source.to_vec());
    }
    if parse_properties_with_limits(source, limits)? != *before {
        return Err(invalid("custom XML Properties source projection is stale"));
    }
    validate_properties(after, limits)?;

    let normalized = normalize_encoding(source)?;
    let spans = property_attribute_spans(normalized.as_ref(), limits)?;
    if spans.item_id.is_some() {
        let output = splice_property_spans(normalized.as_ref(), &spans, after, limits)?;
        if parse_properties_with_limits(&output, limits).is_ok_and(|value| value == *after) {
            return restore_encoding(source, &output, limits.max_properties_bytes);
        }
    }

    let canonical = write_properties_with_limits(after, limits)?;
    restore_encoding(source, &canonical, limits.max_properties_bytes)
}

#[derive(Debug)]
struct PropertySpans {
    root: Option<ElementSpan>,
    item_id: Option<Range<usize>>,
    schema_refs: Option<ElementSpan>,
    schema_references: Vec<SchemaReferenceSpan>,
}

#[derive(Debug)]
struct ElementSpan {
    open: Range<usize>,
    close: Option<Range<usize>>,
    name: Vec<u8>,
}

#[derive(Debug)]
struct SchemaReferenceSpan {
    element: Range<usize>,
    value: Range<usize>,
    name: Vec<u8>,
    attribute: Vec<u8>,
}

fn splice_property_spans(
    source: &[u8],
    spans: &PropertySpans,
    after: &Properties,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(
            spans
                .schema_references
                .len()
                .checked_add(after.schema_references.len())
                .and_then(|value| value.checked_add(2))
                .ok_or_else(|| limit("Properties XML replacement count overflows usize"))?,
        )
        .map_err(|_| limit("Properties XML replacement allocation exceeds its limit"))?;
    let item_id_span = spans
        .item_id
        .clone()
        .ok_or_else(|| invalid("Properties XML itemID span is missing"))?;
    let expand_empty_root_without_schema_refs =
        spans.schema_refs.is_none() && spans.root.as_ref().is_some_and(|root| root.close.is_none());
    if !expand_empty_root_without_schema_refs
        || after.schema_references.len() <= spans.schema_references.len()
    {
        replacements.push((
            item_id_span.clone(),
            escaped_attribute(after.item_id.to_string().as_str()),
        ));
    }

    let existing = &spans.schema_references;
    let common = existing.len().min(after.schema_references.len());
    for (span, value) in existing.iter().zip(&after.schema_references).take(common) {
        replacements.push((span.value.clone(), escaped_attribute(value)));
    }
    for span in existing.iter().skip(after.schema_references.len()) {
        replacements.push((span.element.clone(), Vec::new()));
    }

    if after.schema_references.len() > existing.len() {
        let schema_ref_name = existing
            .first()
            .map(|span| span.name.clone())
            .or_else(|| {
                spans
                    .schema_refs
                    .as_ref()
                    .and_then(|span| qualified_name_with_local(&span.name, b"schemaRef"))
            })
            .or_else(|| {
                spans
                    .root
                    .as_ref()
                    .and_then(|span| qualified_name_with_local(&span.name, b"schemaRef"))
            });
        let insertion = schema_ref_name
            .ok_or_else(|| invalid("Properties XML schemaRef namespace prefix is missing"))?;
        let attributes = existing
            .first()
            .map(|span| span.attribute.clone())
            .or_else(|| qualified_attribute_name(&insertion, b"uri"))
            .or_else(|| {
                spans
                    .schema_refs
                    .as_ref()
                    .and_then(|span| custom_namespace_attribute(source, &span.open))
            })
            .or_else(|| {
                spans
                    .root
                    .as_ref()
                    .and_then(|span| custom_namespace_attribute(source, &span.open))
            })
            .ok_or_else(|| invalid("Properties XML schemaRef attribute prefix is missing"))?;
        let mut added = Vec::new();
        for uri in &after.schema_references[existing.len()..] {
            append_schema_reference(&mut added, &insertion, &attributes, uri);
        }
        if let Some(schema_refs) = spans.schema_refs.as_ref() {
            if let Some(close) = &schema_refs.close {
                replacements.push((close.start..close.start, added));
            } else {
                let replacement =
                    expand_empty_element(source, &schema_refs.open, &schema_refs.name, &added)?;
                replacements.push((schema_refs.open.clone(), replacement));
            }
        } else {
            let root = spans
                .root
                .as_ref()
                .ok_or_else(|| invalid("Properties XML datastoreItem root span is missing"))?;
            let schema_refs_name = qualified_name_with_local(&root.name, b"schemaRefs")
                .ok_or_else(|| invalid("Properties XML schemaRefs namespace prefix is missing"))?;
            let mut replacement = Vec::new();
            append_element_start(&mut replacement, &schema_refs_name);
            replacement.extend_from_slice(&added);
            append_element_end(&mut replacement, &schema_refs_name);
            if let Some(close) = &root.close {
                replacements.push((close.start..close.start, replacement));
            } else {
                let replacement = if expand_empty_root_without_schema_refs {
                    expand_empty_element_with_attribute(
                        source,
                        &root.open,
                        &root.name,
                        &replacement,
                        &item_id_span,
                        &escaped_attribute(after.item_id.to_string().as_str()),
                    )?
                } else {
                    expand_empty_element(source, &root.open, &root.name, &replacement)?
                };
                replacements.push((root.open.clone(), replacement));
            }
        }
    }

    splice_attribute_values(source, replacements, limits)
}

fn property_attribute_spans(source: &[u8], limits: &Limits) -> Result<PropertySpans> {
    let mut reader = NsReader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut opaque_depth = 0usize;
    let mut schema_refs_seen = false;
    let mut root_open = false;
    let mut schema_refs_open = false;
    let mut root = None;
    let mut item_id = None;
    let mut schema_refs = None;
    let mut schema_references = Vec::new();
    let mut open_schema_reference = None;

    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("Properties XML source offset overflows usize"))?;
        buffer.clear();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| super::model::xml_error(error.to_string()))?;
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("Properties XML source offset overflows usize"))?;
        match event {
            Event::Start(element) => {
                let local_name = element.local_name();
                let (namespace, _) = resolved_element(&reader, &element)?;
                let raw = source
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("Properties XML element range is outside its source"))?;
                let is_custom = namespace.as_deref() == Some(CUSTOM_XML_NAMESPACE.as_bytes());
                if opaque_depth > 0 {
                    opaque_depth = opaque_depth
                        .checked_add(1)
                        .ok_or_else(|| limit("Properties XML opaque depth overflows usize"))?;
                } else if depth == 0
                    && is_custom
                    && local_name.as_ref() == b"datastoreItem"
                    && root.is_none()
                {
                    let name = qualified_name(raw)?.to_vec();
                    root = Some(ElementSpan {
                        open: event_start..event_end,
                        close: None,
                        name,
                    });
                    item_id = find_resolved_attribute_value_span(
                        &reader,
                        &element,
                        raw,
                        event_start,
                        CUSTOM_XML_NAMESPACE,
                        b"itemID",
                    )?
                    .map(|(span, _)| span);
                    root_open = true;
                } else if depth == 1
                    && is_custom
                    && local_name.as_ref() == b"schemaRefs"
                    && root.is_some()
                    && !schema_refs_seen
                {
                    schema_refs = Some(ElementSpan {
                        open: event_start..event_end,
                        close: None,
                        name: qualified_name(raw)?.to_vec(),
                    });
                    schema_refs_seen = true;
                    schema_refs_open = true;
                } else if depth == 2
                    && is_custom
                    && local_name.as_ref() == b"schemaRef"
                    && schema_refs_seen
                {
                    if schema_references.len() >= limits.max_schema_references {
                        return Err(limit("schema reference count exceeds its limit"));
                    }
                    let (value, attribute) = find_resolved_attribute_value_span(
                        &reader,
                        &element,
                        raw,
                        event_start,
                        CUSTOM_XML_NAMESPACE,
                        b"uri",
                    )?
                    .ok_or_else(|| invalid("schemaRef lacks its uri attribute"))?;
                    schema_references.push(SchemaReferenceSpan {
                        element: event_start..event_end,
                        value,
                        name: qualified_name(raw)?.to_vec(),
                        attribute,
                    });
                    open_schema_reference = Some(schema_references.len() - 1);
                } else if depth > 0 {
                    opaque_depth = 1;
                } else {
                    return Err(invalid("Properties XML violates datastoreItem grammar"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("Properties XML depth overflows usize"))?;
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                let (namespace, _) = resolved_element(&reader, &element)?;
                let raw = source
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("Properties XML element range is outside its source"))?;
                let is_custom = namespace.as_deref() == Some(CUSTOM_XML_NAMESPACE.as_bytes());
                if opaque_depth > 0 {
                    // The parser treats empty elements inside an opaque
                    // subtree as opaque without changing the subtree depth.
                } else if depth == 0
                    && is_custom
                    && local_name.as_ref() == b"datastoreItem"
                    && root.is_none()
                {
                    item_id = find_resolved_attribute_value_span(
                        &reader,
                        &element,
                        raw,
                        event_start,
                        CUSTOM_XML_NAMESPACE,
                        b"itemID",
                    )?
                    .map(|(span, _)| span);
                    root = Some(ElementSpan {
                        open: event_start..event_end,
                        close: None,
                        name: qualified_name(raw)?.to_vec(),
                    });
                } else if depth == 1
                    && is_custom
                    && local_name.as_ref() == b"schemaRefs"
                    && root.is_some()
                    && !schema_refs_seen
                {
                    schema_refs = Some(ElementSpan {
                        open: event_start..event_end,
                        close: None,
                        name: qualified_name(raw)?.to_vec(),
                    });
                    schema_refs_seen = true;
                } else if depth == 2
                    && is_custom
                    && local_name.as_ref() == b"schemaRef"
                    && schema_refs_seen
                {
                    if schema_references.len() >= limits.max_schema_references {
                        return Err(limit("schema reference count exceeds its limit"));
                    }
                    let (value, attribute) = find_resolved_attribute_value_span(
                        &reader,
                        &element,
                        raw,
                        event_start,
                        CUSTOM_XML_NAMESPACE,
                        b"uri",
                    )?
                    .ok_or_else(|| invalid("schemaRef lacks its uri attribute"))?;
                    schema_references.push(SchemaReferenceSpan {
                        element: event_start..event_end,
                        value,
                        name: qualified_name(raw)?.to_vec(),
                        attribute,
                    });
                }
            },
            Event::End(_) => {
                if opaque_depth > 0 {
                    opaque_depth -= 1;
                } else if depth == 3 {
                    if let Some(index) = open_schema_reference.take() {
                        schema_references[index].element.end = event_end;
                    }
                } else if depth == 2 && schema_refs_open {
                    if let Some(schema_refs) = schema_refs.as_mut() {
                        schema_refs.close = Some(event_start..event_end);
                    }
                    schema_refs_open = false;
                } else if depth == 1 && root_open {
                    if let Some(root) = root.as_mut() {
                        root.close = Some(event_start..event_end);
                    }
                    root_open = false;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Properties XML has an unexpected closing tag"))?;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(PropertySpans {
        root,
        item_id,
        schema_refs,
        schema_references,
    })
}

fn qualified_name(tag: &[u8]) -> Result<&[u8]> {
    if tag.len() < 3 || tag[0] != b'<' {
        return Err(invalid("Properties XML element tag is malformed"));
    }
    let mut end = 1usize;
    while end < tag.len() && !tag[end].is_ascii_whitespace() && !matches!(tag[end], b'/' | b'>') {
        end += 1;
    }
    if end == 1 {
        return Err(invalid("Properties XML element name is empty"));
    }
    Ok(&tag[1..end])
}

fn qualified_name_with_local(name: &[u8], local: &[u8]) -> Option<Vec<u8>> {
    let mut output = name
        .iter()
        .rposition(|byte| *byte == b':')
        .map_or_else(Vec::new, |index| name[..=index].to_vec());
    output.extend_from_slice(local);
    Some(output)
}

fn qualified_attribute_name(name: &[u8], local: &[u8]) -> Option<Vec<u8>> {
    let prefix = name
        .iter()
        .rposition(|byte| *byte == b':')
        .map(|index| &name[..=index])?;
    let mut output = prefix.to_vec();
    output.extend_from_slice(local);
    Some(output)
}

fn custom_namespace_attribute(source: &[u8], open: &Range<usize>) -> Option<Vec<u8>> {
    let tag = source.get(open.clone())?;
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
        let name = &tag[name_start..cursor];
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] != b'=' {
            return None;
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let delimiter = *tag.get(cursor)?;
        if !matches!(delimiter, b'\'' | b'"') {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != delimiter {
            cursor += 1;
        }
        let value = tag.get(value_start..cursor)?;
        if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            if value == CUSTOM_XML_NAMESPACE.as_bytes() {
                let mut attribute = prefix.to_vec();
                attribute.extend_from_slice(b":uri");
                return Some(attribute);
            }
        }
        cursor += 1;
    }
    None
}

fn append_element_start(output: &mut Vec<u8>, name: &[u8]) {
    output.push(b'<');
    output.extend_from_slice(name);
    output.push(b'>');
}

fn append_element_end(output: &mut Vec<u8>, name: &[u8]) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
}

fn append_schema_reference(
    output: &mut Vec<u8>,
    element_name: &[u8],
    attribute_name: &[u8],
    uri: &str,
) {
    output.push(b'<');
    output.extend_from_slice(element_name);
    output.push(b' ');
    output.extend_from_slice(attribute_name);
    output.extend_from_slice(b"=\"");
    escape_attribute(output, uri);
    output.extend_from_slice(b"\"/>");
}

fn expand_empty_element(
    source: &[u8],
    open: &Range<usize>,
    name: &[u8],
    content: &[u8],
) -> Result<Vec<u8>> {
    let tag = source
        .get(open.clone())
        .ok_or_else(|| invalid("Properties XML empty element range is outside its source"))?;
    if !tag.ends_with(b"/>") {
        return Err(invalid("Properties XML element is not self-closing"));
    }
    let mut output = Vec::new();
    output.extend_from_slice(&tag[..tag.len() - 2]);
    output.push(b'>');
    output.extend_from_slice(content);
    append_element_end(&mut output, name);
    Ok(output)
}

fn expand_empty_element_with_attribute(
    source: &[u8],
    open: &Range<usize>,
    name: &[u8],
    content: &[u8],
    attribute: &Range<usize>,
    value: &[u8],
) -> Result<Vec<u8>> {
    let tag = source
        .get(open.clone())
        .ok_or_else(|| invalid("Properties XML empty element range is outside its source"))?;
    if !tag.ends_with(b"/>") {
        return Err(invalid("Properties XML element is not self-closing"));
    }
    if attribute.start < open.start || attribute.end > open.end || attribute.end < attribute.start {
        return Err(invalid("Properties XML itemID range is outside its root"));
    }
    let mut output = Vec::new();
    output.extend_from_slice(&source[open.start..attribute.start]);
    output.extend_from_slice(value);
    output.extend_from_slice(&source[attribute.end..open.end - 2]);
    output.push(b'>');
    output.extend_from_slice(content);
    append_element_end(&mut output, name);
    Ok(output)
}

fn find_resolved_attribute_value_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
    namespace: &str,
    wanted_local_name: &[u8],
) -> Result<Option<(Range<usize>, Vec<u8>)>> {
    let mut result = None;
    for attribute_result in element.checked_attributes() {
        let attribute =
            attribute_result.map_err(|error| super::model::xml_error(error.to_string()))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != wanted_local_name
            || !matches!(resolved, ResolveResult::Bound(bound) if bound.as_ref() == namespace.as_bytes())
        {
            continue;
        }
        let name = attribute.key.as_ref();
        let span = find_attribute_value_span_exact(tag, tag_start, name)?
            .ok_or_else(|| invalid("Properties XML resolved attribute has no source span"))?;
        if result.is_some() {
            return Err(invalid("Properties XML has a duplicate attribute"));
        }
        result = Some((span, name.to_vec()));
    }
    Ok(result)
}

fn find_attribute_value_span_exact(
    tag: &[u8],
    tag_start: usize,
    wanted_name: &[u8],
) -> Result<Option<Range<usize>>> {
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
        let name = &tag[name_start..cursor];
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] != b'=' {
            return Err(invalid("Properties XML attribute is malformed"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let delimiter = *tag
            .get(cursor)
            .ok_or_else(|| invalid("Properties XML attribute has no value"))?;
        if !matches!(delimiter, b'\'' | b'"') {
            return Err(invalid("Properties XML attribute is unquoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != delimiter {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(invalid("Properties XML attribute is unterminated"));
        }
        if name == wanted_name {
            return Ok(Some(tag_start + value_start..tag_start + cursor));
        }
        cursor += 1;
    }
    Ok(None)
}

fn escaped_attribute(value: &str) -> Vec<u8> {
    let mut output = Vec::with_capacity(value.len());
    escape_attribute(&mut output, value);
    output
}

fn splice_attribute_values(
    source: &[u8],
    mut replacements: Vec<(Range<usize>, Vec<u8>)>,
    limits: &Limits,
) -> Result<Vec<u8>> {
    replacements.sort_unstable_by_key(|(range, _)| range.start);
    let mut output_len = source.len();
    let mut cursor = 0usize;
    for (range, replacement) in &replacements {
        if range.start < cursor || range.end < range.start || range.end > source.len() {
            return Err(invalid(
                "Properties XML attribute replacement ranges overlap",
            ));
        }
        output_len = output_len
            .checked_sub(range.end - range.start)
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or_else(|| limit("Properties XML rewrite size overflows usize"))?;
        cursor = range.end;
    }
    if output_len > limits.max_properties_bytes {
        return Err(limit("serialized Properties XML exceeds its byte limit"));
    }
    let mut output = Vec::with_capacity(output_len);
    cursor = 0;
    for (range, replacement) in replacements {
        output.extend_from_slice(&source[cursor..range.start]);
        output.extend_from_slice(&replacement);
        cursor = range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

pub(crate) fn parse_properties_with_limits(xml: &[u8], limits: &Limits) -> Result<Properties> {
    if xml.len() > limits.max_properties_bytes {
        return Err(limit("Properties XML exceeds its byte limit"));
    }
    let normalized = normalize_encoding(xml)?;
    let mut reader = NsReader::from_reader(normalized.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut state = PropertiesParseState::default();

    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| super::model::xml_error(error.to_string()))?
        {
            Event::Start(element) => {
                count_properties_element(&mut state, limits)?;
                if state.opaque_depth > 0 {
                    state.opaque_depth = state
                        .opaque_depth
                        .checked_add(1)
                        .ok_or_else(|| limit("Properties XML opaque depth overflows usize"))?;
                } else {
                    process_properties_element(&reader, &element, limits, &mut state, true)?;
                }
                state.depth += 1;
                if state.depth > limits.max_xml_depth {
                    return Err(limit("Properties XML depth exceeds its limit"));
                }
            },
            Event::Empty(element) => {
                count_properties_element(&mut state, limits)?;
                if state.opaque_depth == 0 {
                    process_properties_element(&reader, &element, limits, &mut state, false)?;
                }
                if state.depth == 0 {
                    state.root_closed = true;
                }
            },
            Event::End(_) => {
                if state.opaque_depth > 0 {
                    state.opaque_depth -= 1;
                }
                state.depth = state
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Properties XML has an unexpected closing tag"))?;
                if state.depth == 0 {
                    state.root_closed = true;
                }
            },
            Event::Text(text) if state.opaque_depth == 0 && !is_whitespace(text.as_ref()) => {
                return Err(invalid("Properties XML contains text content"));
            },
            Event::CData(text) if state.opaque_depth == 0 && !is_whitespace(text.as_ref()) => {
                return Err(invalid("Properties XML contains CDATA content"));
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD and general entity references are forbidden in Properties XML",
                ));
            },
            Event::GeneralRef(reference) => {
                if state.depth == 0 {
                    return Err(invalid("Properties XML has a reference outside its root"));
                }
                let whitespace = validate_reference(&reference)?;
                if state.opaque_depth == 0 && !whitespace {
                    return Err(invalid("Properties XML contains text content"));
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if !state.root_seen || !state.root_closed || state.depth != 0 {
        return Err(invalid("Properties XML has no complete datastoreItem root"));
    }
    let properties = Properties {
        item_id: state
            .item_id
            .ok_or_else(|| invalid("datastoreItem lacks itemID"))?,
        schema_references: state.schema_references,
    };
    validate_properties(&properties, limits)?;
    Ok(properties)
}

fn process_properties_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &Limits,
    state: &mut PropertiesParseState,
    capture_opaque: bool,
) -> Result<()> {
    let (namespace, local_name) = resolved_element(reader, element)?;
    match (state.depth, local_name.as_slice(), namespace.as_deref()) {
        (0, b"datastoreItem", Some(resolved_namespace))
            if resolved_namespace == CUSTOM_XML_NAMESPACE.as_bytes() && !state.root_seen =>
        {
            state.item_id = Some(
                required_attribute(reader, element, CUSTOM_XML_NAMESPACE, b"itemID", limits)?
                    .parse()?,
            );
            reject_other_attributes(
                reader,
                element,
                &[(CUSTOM_XML_NAMESPACE, b"itemID")],
                limits,
            )?;
            state.root_seen = true;
        },
        (1, b"schemaRefs", Some(resolved_namespace))
            if resolved_namespace == CUSTOM_XML_NAMESPACE.as_bytes()
                && state.root_seen
                && !state.schema_refs_seen =>
        {
            reject_other_attributes(reader, element, &[], limits)?;
            state.schema_refs_seen = true;
        },
        (2, b"schemaRef", Some(resolved_namespace))
            if resolved_namespace == CUSTOM_XML_NAMESPACE.as_bytes() && state.schema_refs_seen =>
        {
            if state.schema_references.len() >= limits.max_schema_references {
                return Err(limit("schema reference count exceeds its limit"));
            }
            state.schema_references.push(required_attribute(
                reader,
                element,
                CUSTOM_XML_NAMESPACE,
                b"uri",
                limits,
            )?);
            reject_other_attributes(reader, element, &[(CUSTOM_XML_NAMESPACE, b"uri")], limits)?;
        },
        _ if state.depth > 0 => {
            if capture_opaque {
                state.opaque_depth = 1;
            }
        },
        _ => return Err(invalid("Properties XML violates datastoreItem grammar")),
    }
    Ok(())
}

fn count_properties_element(state: &mut PropertiesParseState, limits: &Limits) -> Result<()> {
    state.element_count = state
        .element_count
        .checked_add(1)
        .ok_or_else(|| limit("Properties XML element count overflows usize"))?;
    if state.element_count > limits.max_xml_elements {
        return Err(limit("Properties XML element count exceeds its limit"));
    }
    Ok(())
}

pub(crate) fn validate_store(store: &Store, limits: &Limits) -> Result<()> {
    validate_limits(limits)?;
    if store.items.len() > limits.max_items {
        return Err(limit("item count exceeds its limit"));
    }
    let mut names = HashSet::with_capacity(store.items.len());
    let mut ids = HashSet::with_capacity(store.items.len());
    let mut total = 0usize;
    for item in &store.items {
        validate_storage_name(&item.storage_name)?;
        if !names.insert(item.storage_name.to_uppercase()) {
            return Err(invalid("custom XML storage name is duplicated"));
        }
        if !ids.insert(item.properties.item_id) {
            return Err(invalid("custom XML itemID is duplicated"));
        }
        let root = validate_payload(&item.xml, limits)?;
        if root != item.root_name {
            return Err(invalid("cached Item root name disagrees with Item XML"));
        }
        let properties = parse_properties_with_limits(&item.properties_xml, limits)?;
        if properties != item.properties {
            return Err(invalid(
                "typed properties disagree with preserved Properties XML",
            ));
        }
        total = total
            .checked_add(item.xml.len())
            .and_then(|value| value.checked_add(item.properties_xml.len()))
            .ok_or_else(|| limit("aggregate stream size overflows usize"))?;
        if total > limits.max_total_bytes {
            return Err(limit("aggregate stream size exceeds its limit"));
        }
    }
    Ok(())
}

pub(crate) fn validate_properties(properties: &Properties, limits: &Limits) -> Result<()> {
    if properties.schema_references.len() > limits.max_schema_references {
        return Err(limit("schema reference count exceeds its limit"));
    }
    let total =
        properties
            .schema_references
            .iter()
            .try_fold(0usize, |total, value| -> Result<usize> {
                validate_characters(value)?;
                total
                    .checked_add(value.len())
                    .ok_or_else(|| limit("schema reference strings overflow usize"))
            })?;
    if total > limits.max_string_bytes {
        return Err(limit("schema reference strings exceed their byte limit"));
    }
    Ok(())
}

pub(crate) fn validate_storage_name(value: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if value.is_empty() || units > 31 || value.contains('\0') {
        return Err(invalid(
            "custom XML sub-storage name is empty, too long, or contains NUL",
        ));
    }
    Ok(())
}

fn validate_limits(limits: &Limits) -> Result<()> {
    if limits.max_item_bytes == 0
        || limits.max_properties_bytes == 0
        || limits.max_total_bytes == 0
        || limits.max_xml_depth == 0
        || limits.max_xml_elements == 0
        || limits.max_string_bytes == 0
    {
        return Err(limit(
            "configured byte, depth, and element limits must be nonzero",
        ));
    }
    Ok(())
}

fn stream_size(stream_size: u64, label: &str, limits: &Limits) -> Result<usize> {
    let size = usize::try_from(stream_size)
        .map_err(|_conversion_error| limit(format!("{label} size overflows usize")))?;
    if size > limits.max_total_bytes {
        return Err(limit(format!("{label} size exceeds aggregate byte limit")));
    }
    Ok(size)
}

fn marker_exists<R: Read + Seek>(ole: &OleFile<R>, name: &str) -> Result<bool> {
    if ole.directory_exists(&[name]) {
        return Ok(true);
    }
    if ole.exists(&[name]) {
        return Err(invalid(format!("{name} is not a storage")));
    }
    Ok(false)
}

fn map_ole_limit(error: Error) -> Error {
    match error {
        Error::Ole(OleError::LimitExceeded {
            resource,
            observed,
            maximum,
        }) => limit(format!(
            "{resource} count {observed} exceeds the configured limit {maximum}"
        )),
        other => other,
    }
}
