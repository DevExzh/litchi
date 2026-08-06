use std::collections::HashSet;
use std::io::{Read, Seek};
use std::ops::Range;
use std::sync::Arc;

use litchi_cfb::{OleFile, OleWriter};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use super::model::{
    Error, Item, ItemId, Limits, Promotion, Properties, Result, Store, invalid, limit,
};
use super::xml::{
    escape_attribute, is_whitespace, normalize_encoding, reject_other_attributes,
    required_attribute, resolved_element, validate_characters, validate_payload,
};
use super::{
    CUSTOM_XML_NAMESPACE, ITEM_STREAM, MODIFIED_PROMOTION_STORAGE, PROPERTIES_STREAM,
    REDUNDANT_PROMOTION_STORAGE, STORE_STORAGE,
};

/// Inspect a complete Custom XML data store with default limits.
pub fn inspect<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Option<Store>> {
    inspect_with_limits(ole, Limits::default())
}

/// Inspect a complete Custom XML data store with caller-selected limits.
pub fn inspect_with_limits<R: Read + Seek>(
    ole: &mut OleFile<R>,
    limits: Limits,
) -> Result<Option<Store>> {
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

    let entries = ole.list_directory_entries(&[STORE_STORAGE])?;
    if entries.len() > limits.max_items {
        return Err(limit(format!(
            "item count {} exceeds {}",
            entries.len(),
            limits.max_items
        )));
    }
    let mut names = Vec::with_capacity(entries.len());
    for entry in entries {
        if entry.entry_type != 1 {
            return Err(invalid(format!(
                "MsoDataStore child '{}' is not a storage",
                entry.name
            )));
        }
        validate_storage_name(&entry.name)?;
        names.push(entry.name.clone());
    }
    names.sort();

    let mut total_bytes = 0usize;
    let mut item_ids = HashSet::with_capacity(names.len());
    let mut items = Vec::with_capacity(names.len());
    for storage_name in names {
        let entries = ole.list_directory_entries(&[STORE_STORAGE, storage_name.as_str()])?;
        if entries.len() != 2 {
            return Err(invalid(format!(
                "custom XML sub-storage '{storage_name}' must contain exactly Item and Properties"
            )));
        }
        let mut item_size = None;
        let mut properties_size = None;
        for entry in entries {
            match (entry.name.as_str(), entry.entry_type) {
                (ITEM_STREAM, 2) => item_size = Some(stream_size(entry.size, "Item", &limits)?),
                (PROPERTIES_STREAM, 2) => {
                    properties_size = Some(stream_size(entry.size, "Properties", &limits)?)
                },
                _ => {
                    return Err(invalid(format!(
                        "custom XML sub-storage '{storage_name}' has unexpected entry '{}'",
                        entry.name
                    )));
                },
            }
        }
        let item_size = item_size.ok_or_else(|| invalid("custom XML Item stream is missing"))?;
        let properties_size =
            properties_size.ok_or_else(|| invalid("custom XML Properties stream is missing"))?;
        if item_size > limits.max_item_bytes {
            return Err(limit(format!(
                "Item stream has {item_size} bytes, limit is {}",
                limits.max_item_bytes
            )));
        }
        if properties_size > limits.max_properties_bytes {
            return Err(limit(format!(
                "Properties stream has {properties_size} bytes, limit is {}",
                limits.max_properties_bytes
            )));
        }
        total_bytes = total_bytes
            .checked_add(item_size)
            .and_then(|value| value.checked_add(properties_size))
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

/// Materialize a validated store in a newly assembled OLE writer.
pub fn write(writer: &mut OleWriter, store: &Store) -> Result<()> {
    validate_store(store, &Limits::default())?;
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

/// Parse the schema-defined Custom XML data-store Properties stream.
pub fn parse_properties(xml: &[u8]) -> Result<Properties> {
    parse_properties_with_limits(xml, &Limits::default())
}

/// Serialize Custom XML data-store Properties in stable schema order.
pub fn write_properties(properties: &Properties) -> Result<Vec<u8>> {
    write_properties_with_limits(properties, &Limits::default())
}

pub(crate) fn write_properties_with_limits(
    properties: &Properties,
    limits: &Limits,
) -> Result<Vec<u8>> {
    validate_properties(properties, limits)?;
    let mut output = Vec::with_capacity(256);
    output.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    output.extend_from_slice(b"<ds:datastoreItem xmlns:ds=\"");
    output.extend_from_slice(CUSTOM_XML_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\" ds:itemID=\"");
    output.extend_from_slice(properties.item_id.to_string().as_bytes());
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
    if output.len() > limits.max_properties_bytes {
        return Err(limit("serialized Properties XML exceeds its byte limit"));
    }
    Ok(output)
}

/// Rewrite known properties attributes while retaining the source XML around
/// them. A structural schema-reference edit falls back to the deterministic
/// writer; an unchanged projection always returns the exact source allocation.
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

    if std::str::from_utf8(source).is_ok() {
        let (item_id, schema_references) = property_attribute_spans(source, limits)?;
        if let Some(item_id) = item_id
            && schema_references.len() == after.schema_references.len()
        {
            let mut replacements = Vec::with_capacity(schema_references.len() + 1);
            replacements.push((
                item_id,
                escaped_attribute(after.item_id.to_string().as_str()),
            ));
            for (span, value) in schema_references.into_iter().zip(&after.schema_references) {
                replacements.push((span, escaped_attribute(value)));
            }
            let output = splice_attribute_values(source, replacements)?;
            if output.len() <= limits.max_properties_bytes
                && parse_properties_with_limits(&output, limits).is_ok_and(|value| value == *after)
            {
                return Ok(output);
            }
        }
    }

    write_properties_with_limits(after, limits)
}

fn property_attribute_spans(
    source: &[u8],
    limits: &Limits,
) -> Result<(Option<Range<usize>>, Vec<Range<usize>>)> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut item_id = None;
    let mut schema_references = Vec::new();

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| limit("Properties XML source offset overflows usize"))?;
        buffer.clear();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| super::model::xml_error(error.to_string()))?;
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| limit("Properties XML source offset overflows usize"))?;
        match event {
            Event::Start(element) => {
                let local_name = element.local_name();
                let raw = source
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("Properties XML element range is outside its source"))?;
                if depth == 0 && local_name.as_ref() == b"datastoreItem" {
                    item_id = find_attribute_value_span(raw, event_start, b"itemID")?;
                } else if depth == 2 && local_name.as_ref() == b"schemaRef" {
                    if schema_references.len() >= limits.max_schema_references {
                        return Err(limit("schema reference count exceeds its limit"));
                    }
                    if let Some(span) = find_attribute_value_span(raw, event_start, b"uri")? {
                        schema_references.push(span);
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("Properties XML depth overflows usize"))?;
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                let raw = source
                    .get(event_start..event_end)
                    .ok_or_else(|| invalid("Properties XML element range is outside its source"))?;
                if depth == 0 && local_name.as_ref() == b"datastoreItem" {
                    item_id = find_attribute_value_span(raw, event_start, b"itemID")?;
                } else if depth == 2 && local_name.as_ref() == b"schemaRef" {
                    if schema_references.len() >= limits.max_schema_references {
                        return Err(limit("schema reference count exceeds its limit"));
                    }
                    if let Some(span) = find_attribute_value_span(raw, event_start, b"uri")? {
                        schema_references.push(span);
                    }
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Properties XML has an unexpected closing tag"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((item_id, schema_references))
}

fn find_attribute_value_span(
    tag: &[u8],
    tag_start: usize,
    wanted_local_name: &[u8],
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
        let local_name = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
        if local_name == wanted_local_name {
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
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD and general entity references are forbidden in Properties XML",
                ));
            },
            Event::Eof => break,
            _ => {},
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

fn process_properties_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &Limits,
    state: &mut PropertiesParseState,
    capture_opaque: bool,
) -> Result<()> {
    let (namespace, local_name) = resolved_element(reader, element)?;
    match (state.depth, local_name.as_slice(), namespace.as_deref()) {
        (0, b"datastoreItem", Some(namespace))
            if namespace == CUSTOM_XML_NAMESPACE.as_bytes() && !state.root_seen =>
        {
            state.item_id = Some(
                required_attribute(reader, element, CUSTOM_XML_NAMESPACE, b"itemID")?.parse()?,
            );
            reject_other_attributes(reader, element, &[(CUSTOM_XML_NAMESPACE, b"itemID")])?;
            state.root_seen = true;
        },
        (1, b"schemaRefs", Some(namespace))
            if namespace == CUSTOM_XML_NAMESPACE.as_bytes()
                && state.root_seen
                && !state.schema_refs_seen =>
        {
            reject_other_attributes(reader, element, &[])?;
            state.schema_refs_seen = true;
        },
        (2, b"schemaRef", Some(namespace))
            if namespace == CUSTOM_XML_NAMESPACE.as_bytes() && state.schema_refs_seen =>
        {
            if state.schema_references.len() >= limits.max_schema_references {
                return Err(limit("schema reference count exceeds its limit"));
            }
            state.schema_references.push(required_attribute(
                reader,
                element,
                CUSTOM_XML_NAMESPACE,
                b"uri",
            )?);
            reject_other_attributes(reader, element, &[(CUSTOM_XML_NAMESPACE, b"uri")])?;
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

fn stream_size(value: u64, label: &str, limits: &Limits) -> Result<usize> {
    let value =
        usize::try_from(value).map_err(|_| limit(format!("{label} size overflows usize")))?;
    if value > limits.max_total_bytes {
        return Err(limit(format!("{label} size exceeds aggregate byte limit")));
    }
    Ok(value)
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

#[allow(
    dead_code,
    reason = "keeps the module error type explicit for API audits"
)]
fn _error_type_marker(_: Error) {}
