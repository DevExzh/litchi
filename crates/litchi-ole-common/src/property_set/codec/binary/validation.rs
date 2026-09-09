//! Section and stream structure validation and serialization.

use super::super::super::model::{
    CodePage, DEFAULT_CODEPAGE, Guid, PID_CODEPAGE, PID_DICTIONARY, Section, Stream,
    UNICODE_CODEPAGE, Value, align4_len, checked_add, checked_mul, checked_u32, invalid,
    try_clone_vec, try_hash_map_with_capacity, try_hash_set_with_capacity, try_vec_with_capacity,
    valid_named_property_identifier, valid_property_identifier,
};
use super::super::semantic::{MAX_PROPERTY_COUNT, validate_section};
use super::super::support::allocation;
use super::semantic::{
    append_typed_for_property, serialize_typed_for_property,
    serialize_typed_for_property_with_limit,
};
use super::wire::{
    ByteSink, ByteWriter, CountingWriter, ValueReader, append_ansi, append_bytes, append_u32,
    checked_range, decode_ansi, decode_utf16, encoded_ansi_len, pad4, read_guid, read_u16,
    read_u32, reserve_bytes,
};
use super::{parse_typed_property, parse_typed_property_for_property};
use litchi_cfb::OleError;
use litchi_cfb::consts::VT_I2;
use std::collections::HashMap;

const PROPERTY_SET_HEADER_SIZE: usize = 28;
const SECTION_DESCRIPTOR_SIZE: usize = 20;
const SECTION_HEADER_SIZE: usize = 8;
const PROPERTY_DESCRIPTOR_SIZE: usize = 8;

pub(super) fn parse_stream(data: &[u8]) -> Result<Stream, OleError> {
    parse_stream_with_mode(data, false)
}

pub(super) fn parse_non_simple_stream(data: &[u8]) -> Result<Stream, OleError> {
    parse_stream_with_mode(data, true)
}

fn parse_stream_with_mode(data: &[u8], allow_indirect: bool) -> Result<Stream, OleError> {
    if data.len() < PROPERTY_SET_HEADER_SIZE + SECTION_DESCRIPTOR_SIZE {
        return Err(invalid("Property Set stream is too short"));
    }
    if read_u16(data, 0, "byte order")? != 0xfffe {
        return Err(invalid("Property Set byte order must be 0xFFFE"));
    }
    let version = read_u16(data, 2, "version")?;
    if !matches!(version, Stream::VERSION_0 | Stream::VERSION_1) {
        return Err(invalid(format!(
            "Unsupported Property Set version {version}"
        )));
    }
    let system_identifier = read_u32(data, 4, "system identifier")?;
    let class_identifier = read_guid(data, 8, "class identifier")?;
    let section_count = usize::try_from(read_u32(data, 24, "section count")?)
        .map_err(|_conversion_error| invalid("Property Set section count is too large"))?;
    if !(1..=2).contains(&section_count) {
        return Err(invalid(format!(
            "Property Set section count {section_count} is outside 1..=2"
        )));
    }
    let descriptor_end = section_count
        .checked_mul(SECTION_DESCRIPTOR_SIZE)
        .and_then(|size| PROPERTY_SET_HEADER_SIZE.checked_add(size))
        .ok_or_else(|| invalid("Property Set descriptor table overflow"))?;
    checked_range(data, 0, descriptor_end, "descriptor table")?;

    let mut descriptors = try_vec_with_capacity(section_count, "property-set descriptors")?;
    let mut format_ids = try_hash_set_with_capacity(section_count, "property-set format IDs")?;
    let mut offsets = try_hash_set_with_capacity(section_count, "property-set offsets")?;
    for index in 0..section_count {
        let descriptor = checked_add(
            PROPERTY_SET_HEADER_SIZE,
            checked_mul(index, SECTION_DESCRIPTOR_SIZE, "section descriptor")?,
            "section descriptor",
        )?;
        let format_identifier = read_guid(data, descriptor, "section format identifier")?;
        let offset_field = checked_add(descriptor, 16, "section descriptor")?;
        let offset = usize::try_from(read_u32(data, offset_field, "section offset")?)
            .map_err(|_conversion_error| invalid("Property Set section offset is too large"))?;
        if offset < descriptor_end || offset % 4 != 0 || offset >= data.len() {
            return Err(invalid(format!(
                "Invalid Property Set section offset {offset}"
            )));
        }
        if !format_ids.insert(format_identifier) || !offsets.insert(offset) {
            return Err(invalid("Duplicate Property Set section descriptor"));
        }
        let size = usize::try_from(read_u32(data, offset, "section size")?)
            .map_err(|_conversion_error| invalid("Property Set section size is too large"))?;
        let end = offset
            .checked_add(size)
            .filter(|end| *end <= data.len())
            .ok_or_else(|| invalid("Property Set section exceeds its stream"))?;
        descriptors.push((format_identifier, offset, end));
    }

    let mut ranges = try_vec_with_capacity(descriptors.len(), "property-set section ranges")?;
    ranges.extend(descriptors.iter().map(|(_, start, end)| (*start, *end)));
    ranges.sort_unstable();
    if ranges.windows(2).any(|pair| pair[1].0 < pair[0].1) {
        return Err(invalid("Property Set sections overlap"));
    }

    let mut sections = try_vec_with_capacity(section_count, "property-set sections")?;
    for (format_identifier, start, end) in descriptors {
        sections.push(parse_section(
            &data[start..end],
            format_identifier,
            version,
            allow_indirect,
        )?);
    }
    Ok(Stream {
        version,
        system_identifier,
        class_identifier,
        sections,
    })
}

pub(super) fn serialize_stream(stream: &Stream) -> Result<Vec<u8>, OleError> {
    let descriptor_end = validate_stream(stream)?;
    serialize_stream_impl(stream, descriptor_end, None)
}

pub(super) fn serialize_stream_with_limit(
    stream: &Stream,
    maximum: u64,
) -> Result<Vec<u8>, OleError> {
    let descriptor_end = validate_stream(stream)?;
    let expected = serialized_stream_len(stream, descriptor_end)?;
    ensure_limit(expected, maximum, "serialized Property Set stream")?;
    serialize_stream_impl(stream, descriptor_end, Some(maximum))
}

fn validate_stream(stream: &Stream) -> Result<usize, OleError> {
    if !matches!(stream.version, Stream::VERSION_0 | Stream::VERSION_1)
        || !(1..=2).contains(&stream.sections.len())
    {
        return Err(invalid(
            "Property Set must have version zero or one and one or two sections",
        ));
    }
    let mut ids = try_hash_set_with_capacity(stream.sections.len(), "property-set format IDs")?;
    for section in &stream.sections {
        if !ids.insert(section.format_identifier) {
            return Err(invalid("Duplicate section format identifier"));
        }
        validate_section(section, stream.version)?;
    }
    checked_add(
        PROPERTY_SET_HEADER_SIZE,
        checked_mul(
            stream.sections.len(),
            SECTION_DESCRIPTOR_SIZE,
            "Property Set descriptor table",
        )?,
        "Property Set descriptor table",
    )
}

fn serialize_stream_impl(
    stream: &Stream,
    descriptor_end: usize,
    maximum: Option<u64>,
) -> Result<Vec<u8>, OleError> {
    let mut sections = try_vec_with_capacity(stream.sections.len(), "serialized sections")?;
    for section in &stream.sections {
        sections.push(serialize_section(section, maximum)?);
    }
    let descriptor_size = align4_len(descriptor_end, "Property Set descriptor table")?;
    let mut offsets = try_vec_with_capacity(sections.len(), "section offsets")?;
    let mut cursor = descriptor_size;
    for section in &sections {
        offsets.push(cursor);
        cursor = align4_len(
            checked_add(cursor, section.len(), "Property Set size")?,
            "Property Set size",
        )?;
    }
    let mut out = ByteWriter::zeroed(cursor, maximum, "serialized Property Set stream")?;
    let bytes = out.as_mut_slice();
    bytes[0..2].copy_from_slice(&0xfffeu16.to_le_bytes());
    bytes[2..4].copy_from_slice(&stream.version.to_le_bytes());
    bytes[4..8].copy_from_slice(&stream.system_identifier.to_le_bytes());
    bytes[8..24].copy_from_slice(stream.class_identifier.as_bytes());
    bytes[24..28]
        .copy_from_slice(&checked_u32(stream.sections.len(), "section count")?.to_le_bytes());
    for (index, section) in stream.sections.iter().enumerate() {
        let base = checked_add(
            PROPERTY_SET_HEADER_SIZE,
            checked_mul(index, SECTION_DESCRIPTOR_SIZE, "section descriptor")?,
            "section descriptor",
        )?;
        bytes[base..base + 16].copy_from_slice(section.format_identifier.as_bytes());
        bytes[base + 16..base + 20]
            .copy_from_slice(&checked_u32(offsets[index], "Section offset")?.to_le_bytes());
    }
    for (index, section) in sections.iter().enumerate() {
        let start = offsets[index];
        let end = checked_add(start, section.len(), "Property Set section")?;
        bytes[start..end].copy_from_slice(section);
    }
    Ok(out.into_bytes())
}

fn serialization_order(
    property_order: &[u32],
    properties: &HashMap<u32, Value>,
    dictionary: &HashMap<u32, String>,
) -> Result<Vec<u32>, OleError> {
    let mut order = try_clone_vec(property_order, "property order")?;
    order
        .try_reserve(properties.len())
        .map_err(|source| allocation("property order", source))?;
    if !dictionary.is_empty() && !order.contains(&PID_DICTIONARY) {
        order.insert(0, PID_DICTIONARY);
    }
    for id in properties.keys() {
        if !order.contains(id) {
            order.push(*id);
        }
    }
    order.retain(|id| {
        *id == PID_DICTIONARY && !dictionary.is_empty() || properties.contains_key(id)
    });
    let mut order_ids = try_hash_set_with_capacity(order.len(), "property order set")?;
    for identifier in order.iter().copied() {
        order_ids.insert(identifier);
    }
    if order_ids.len() != order.len() {
        return Err(invalid("Duplicate property order identifier"));
    }
    Ok(order)
}

fn serialize_section(section: &Section, maximum: Option<u64>) -> Result<Vec<u8>, OleError> {
    let order = serialization_order(
        &section.property_order,
        &section.properties,
        &section.dictionary,
    )?;
    let table_end = checked_add(
        SECTION_HEADER_SIZE,
        checked_mul(
            order.len(),
            PROPERTY_DESCRIPTOR_SIZE,
            "property descriptor table",
        )?,
        "property descriptor table",
    )?;
    let mut values = try_vec_with_capacity(order.len(), "serialized property values")?;
    let codepage = section.codepage.unwrap_or(CodePage::WINDOWS_1252).id();
    for id in &order {
        values.push(if *id == PID_DICTIONARY {
            serialize_dictionary(section, codepage, maximum)?
        } else {
            let value = section
                .properties
                .get(id)
                .ok_or_else(|| invalid(format!("Property order references missing PID {id}")))?;
            match maximum {
                Some(maximum) => {
                    serialize_typed_for_property_with_limit(*id, value, codepage, maximum)?
                },
                None => serialize_typed_for_property(*id, value, codepage)?,
            }
        });
    }
    let table_size = align4_len(table_end, "property descriptor table")?;
    let mut offsets = try_vec_with_capacity(values.len(), "property offsets")?;
    let mut cursor = table_size;
    for value in &values {
        offsets.push(cursor);
        cursor = align4_len(
            checked_add(cursor, value.len(), "Property Set section size")?,
            "Property Set section size",
        )?;
    }
    let mut out = ByteWriter::zeroed(cursor, maximum, "serialized Property Set section")?;
    let bytes = out.as_mut_slice();
    bytes[4..8].copy_from_slice(&checked_u32(order.len(), "property count")?.to_le_bytes());
    for (index, id) in order.iter().enumerate() {
        let base = checked_add(
            SECTION_HEADER_SIZE,
            checked_mul(index, PROPERTY_DESCRIPTOR_SIZE, "property descriptor")?,
            "property descriptor",
        )?;
        bytes[base..base + 4].copy_from_slice(&id.to_le_bytes());
        bytes[base + 4..base + 8]
            .copy_from_slice(&checked_u32(offsets[index], "Property offset")?.to_le_bytes());
    }
    for (index, value) in values.iter().enumerate() {
        let start = offsets[index];
        let end = checked_add(start, value.len(), "Property Set value")?;
        bytes[start..end].copy_from_slice(value);
    }
    let section_len = checked_u32(bytes.len(), "Section length")?;
    bytes[0..4].copy_from_slice(&section_len.to_le_bytes());
    Ok(out.into_bytes())
}

fn append_dictionary<S: ByteSink + ?Sized>(
    out: &mut S,
    section: &Section,
    codepage: u16,
) -> Result<(), OleError> {
    let mut order = try_clone_vec(&section.dictionary_order, "dictionary order")?;
    order
        .try_reserve(section.dictionary.len())
        .map_err(|source| allocation("dictionary order", source))?;
    for id in section.dictionary.keys() {
        if !order.contains(id) {
            order.push(*id);
        }
    }
    order.retain(|id| section.dictionary.contains_key(id));
    let mut order_ids = try_hash_set_with_capacity(order.len(), "dictionary order set")?;
    for identifier in order.iter().copied() {
        order_ids.insert(identifier);
    }
    if order_ids.len() != order.len() {
        return Err(invalid("Duplicate dictionary order identifier"));
    }
    append_u32(
        out,
        checked_u32(order.len(), "dictionary entry count")?,
        "serialized property dictionary",
    )?;
    for id in order {
        let name = section
            .dictionary
            .get(&id)
            .ok_or_else(|| invalid(format!("Dictionary order references missing PID {id}")))?;
        append_u32(out, id, "serialized property dictionary")?;
        if codepage == UNICODE_CODEPAGE {
            let units = checked_add(name.encode_utf16().count(), 1, "Dictionary name length")?;
            let byte_len = checked_mul(units, 2, "Dictionary name length")?;
            append_u32(
                out,
                checked_u32(units, "Dictionary name length")?,
                "serialized property dictionary",
            )?;
            reserve_bytes(out, byte_len, "serialized property dictionary")?;
            for unit in name.encode_utf16() {
                append_bytes(out, &unit.to_le_bytes(), "serialized property dictionary")?;
            }
            append_bytes(out, &0u16.to_le_bytes(), "serialized property dictionary")?;
            pad4(out)?;
        } else {
            let byte_len = checked_add(
                encoded_ansi_len(name, codepage)?,
                1,
                "Dictionary name length",
            )?;
            append_u32(
                out,
                checked_u32(byte_len, "Dictionary name length")?,
                "serialized property dictionary",
            )?;
            reserve_bytes(out, byte_len, "serialized property dictionary")?;
            append_ansi(out, name, codepage, "serialized property dictionary")?;
            out.push_zero("serialized property dictionary")?;
        }
    }
    Ok(())
}

fn serialize_dictionary(
    section: &Section,
    codepage: u16,
    maximum: Option<u64>,
) -> Result<Vec<u8>, OleError> {
    let mut out = ByteWriter::new(4, maximum, "serialized property dictionary")?;
    append_dictionary(&mut out, section, codepage)?;
    Ok(out.into_bytes())
}

fn serialized_stream_len(stream: &Stream, descriptor_end: usize) -> Result<usize, OleError> {
    serialized_stream_layout_len(
        stream.sections.len(),
        descriptor_end,
        stream.sections.iter().map(serialized_section_len),
    )
}

fn serialized_section_len(section: &Section) -> Result<usize, OleError> {
    let order = serialization_order(
        &section.property_order,
        &section.properties,
        &section.dictionary,
    )?;
    let table_end = checked_add(
        SECTION_HEADER_SIZE,
        checked_mul(
            order.len(),
            PROPERTY_DESCRIPTOR_SIZE,
            "property descriptor table",
        )?,
        "property descriptor table",
    )?;
    let codepage = section.codepage.unwrap_or(CodePage::WINDOWS_1252).id();
    serialized_section_layout_len(
        order.len(),
        table_end,
        order.into_iter().map(|id| {
            if id == PID_DICTIONARY {
                serialized_dictionary_len(section, codepage)
            } else {
                let value = section.properties.get(&id).ok_or_else(|| {
                    invalid(format!("Property order references missing PID {id}"))
                })?;
                serialized_typed_for_property_len(id, value, codepage)
            }
        }),
    )
}

pub(super) fn serialized_stream_layout_len<I>(
    section_count: usize,
    descriptor_end: usize,
    section_lengths: I,
) -> Result<usize, OleError>
where
    I: IntoIterator<Item = Result<usize, OleError>>,
{
    checked_u32(section_count, "section count")?;
    let descriptor_size = align4_len(descriptor_end, "Property Set descriptor table")?;
    let mut cursor = descriptor_size;
    let mut seen = 0usize;
    for section_len in section_lengths {
        let section_len = section_len?;
        checked_u32(cursor, "Section offset")?;
        cursor = align4_len(
            checked_add(cursor, section_len, "Property Set size")?,
            "Property Set size",
        )?;
        seen = checked_add(seen, 1, "Property Set section count")?;
    }
    if seen != section_count {
        return Err(invalid("Property Set section length count mismatch"));
    }
    Ok(cursor)
}

pub(super) fn serialized_section_layout_len<I>(
    property_count: usize,
    table_end: usize,
    value_lengths: I,
) -> Result<usize, OleError>
where
    I: IntoIterator<Item = Result<usize, OleError>>,
{
    checked_u32(property_count, "property count")?;
    let mut cursor = align4_len(table_end, "property descriptor table")?;
    let mut seen = 0usize;
    for value_len in value_lengths {
        let value_len = value_len?;
        checked_u32(cursor, "Property offset")?;
        cursor = align4_len(
            checked_add(cursor, value_len, "Property Set section size")?,
            "Property Set section size",
        )?;
        seen = checked_add(seen, 1, "Property count")?;
    }
    if seen != property_count {
        return Err(invalid("Property value length count mismatch"));
    }
    checked_u32(cursor, "Section length")?;
    Ok(cursor)
}

fn serialized_dictionary_len(section: &Section, codepage: u16) -> Result<usize, OleError> {
    let mut out = CountingWriter::new();
    append_dictionary(&mut out, section, codepage)?;
    Ok(out.len())
}

fn serialized_typed_for_property_len(
    property_identifier: u32,
    value: &Value,
    codepage: u16,
) -> Result<usize, OleError> {
    let mut out = CountingWriter::new();
    append_typed_for_property(&mut out, property_identifier, value, codepage)?;
    Ok(out.len())
}

fn ensure_limit(observed: usize, maximum: u64, resource: &'static str) -> Result<(), OleError> {
    let observed = u64::try_from(observed).unwrap_or(u64::MAX);
    if observed > maximum {
        return Err(OleError::LimitExceeded {
            resource,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn parse_section(
    data: &[u8],
    format_identifier: Guid,
    version: u16,
    allow_indirect: bool,
) -> Result<Section, OleError> {
    checked_range(data, 0, SECTION_HEADER_SIZE, "section header")?;
    let declared_size = usize::try_from(read_u32(data, 0, "section size")?)
        .map_err(|_conversion_error| invalid("Property Set section size is too large"))?;
    if declared_size != data.len() {
        return Err(invalid(
            "Property Set section range does not match its size",
        ));
    }
    let property_count = usize::try_from(read_u32(data, 4, "property count")?)
        .map_err(|_conversion_error| invalid("Property count is too large"))?;
    if property_count > MAX_PROPERTY_COUNT {
        return Err(invalid(format!(
            "Property count {property_count} exceeds the safety limit"
        )));
    }
    let table_end = property_count
        .checked_mul(PROPERTY_DESCRIPTOR_SIZE)
        .and_then(|size| SECTION_HEADER_SIZE.checked_add(size))
        .ok_or_else(|| invalid("Property table size overflow"))?;
    if table_end > data.len() {
        return Err(invalid("Property table exceeds its section"));
    }

    let mut identifiers = try_hash_set_with_capacity(property_count, "property identifiers")?;
    let mut offsets = try_hash_set_with_capacity(property_count, "property offsets")?;
    let mut descriptors = try_vec_with_capacity(property_count, "property descriptors")?;
    for index in 0..property_count {
        let descriptor = checked_add(
            SECTION_HEADER_SIZE,
            checked_mul(index, PROPERTY_DESCRIPTOR_SIZE, "property descriptor")?,
            "property descriptor",
        )?;
        let identifier = read_u32(data, descriptor, "property identifier")?;
        if !valid_property_identifier(identifier) {
            return Err(invalid(format!(
                "Property identifier {identifier} is outside the Property Set range"
            )));
        }
        let offset_field = checked_add(descriptor, 4, "property descriptor")?;
        let offset = usize::try_from(read_u32(data, offset_field, "property offset")?)
            .map_err(|_conversion_error| invalid("Property offset is too large"))?;
        // Older Office producers emit valid properties at non-DWORD offsets.
        // Bounds and uniqueness remain mandatory, but alignment is advisory.
        if offset < table_end || offset >= data.len() {
            return Err(invalid(format!("Invalid property offset {offset}")));
        }
        if !identifiers.insert(identifier) {
            return Err(invalid(format!(
                "Duplicate property identifier {identifier}"
            )));
        }
        if !offsets.insert(offset) {
            return Err(invalid(format!("Duplicate property offset {offset}")));
        }
        descriptors.push((offset, identifier));
    }
    descriptors.sort_unstable();

    let property_slice = |identifier: u32| -> Result<Option<(usize, &[u8])>, OleError> {
        let Some(index) = descriptors.iter().position(|(_, id)| *id == identifier) else {
            return Ok(None);
        };
        let start = descriptors[index].0;
        let end = descriptors
            .get(index + 1)
            .map_or(data.len(), |(offset, _)| *offset);
        if end <= start {
            return Err(invalid("Property has an empty or inverted range"));
        }
        Ok(Some((start, &data[start..end])))
    };

    let (codepage, codepage_value) = match property_slice(PID_CODEPAGE)? {
        Some((offset, bytes)) => {
            if read_u16(bytes, 0, "codepage type")? != VT_I2
                || read_u16(bytes, 2, "codepage reserved field")? != 0
            {
                return Err(invalid("PID 1 must be a VT_I2 codepage"));
            }
            let value = parse_typed_property(bytes, DEFAULT_CODEPAGE, offset)?;
            let Value::I2(signed) = value else {
                return Err(invalid("PID 1 must contain a VT_I2 codepage"));
            };
            let codepage = CodePage::try_from(u16::from_ne_bytes(signed.to_ne_bytes()))?;
            (Some(codepage), Some(Value::I2(signed)))
        },
        None => (None, None),
    };
    let effective_codepage = codepage.unwrap_or(CodePage::WINDOWS_1252).id();

    let (dictionary, dictionary_order) = match property_slice(PID_DICTIONARY)? {
        Some((offset, bytes)) => parse_dictionary(bytes, effective_codepage, offset)?,
        None => (HashMap::new(), Vec::new()),
    };
    let mut properties =
        try_hash_map_with_capacity(property_count.saturating_sub(1), "property values")?;
    if let Some(value) = codepage_value {
        properties.insert(PID_CODEPAGE, value);
    }
    for (index, (start, identifier)) in descriptors.iter().enumerate() {
        if matches!(*identifier, PID_DICTIONARY | PID_CODEPAGE) {
            continue;
        }
        let end = descriptors
            .get(index + 1)
            .map_or(data.len(), |(offset, _)| *offset);
        let bytes = &data[*start..end];
        let value = if allow_indirect {
            super::parse_typed_property_for_non_simple_property(
                bytes,
                effective_codepage,
                *start,
                *identifier,
            )
        } else {
            parse_typed_property_for_property(bytes, effective_codepage, *start, *identifier)
        }
        .map_err(|error| {
            invalid(format!(
                "Property {identifier} at section offset {start} is invalid: {error}"
            ))
        })?;
        properties.insert(*identifier, value);
    }
    let mut property_order = try_vec_with_capacity(descriptors.len(), "property order")?;
    property_order.extend(descriptors.iter().map(|(_, identifier)| *identifier));
    let section = Section {
        format_identifier,
        codepage,
        dictionary,
        properties,
        property_order,
        dictionary_order,
    };
    validate_section(&section, version)?;
    Ok(section)
}

fn parse_dictionary(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<(HashMap<u32, String>, Vec<u32>), OleError> {
    let mut reader = ValueReader::new(data, property_offset);
    let count = usize::try_from(reader.read_u32("dictionary count")?)
        .map_err(|_conversion_error| invalid("Dictionary count is too large"))?;
    if count > MAX_PROPERTY_COUNT {
        return Err(invalid(format!(
            "Dictionary count {count} exceeds the safety limit"
        )));
    }
    if count > reader.remaining_len() / 8 {
        return Err(invalid("Dictionary count exceeds its property range"));
    }
    let mut dictionary = try_hash_map_with_capacity(count, "property dictionary")?;
    let mut order = try_vec_with_capacity(count, "dictionary order")?;
    for _ in 0..count {
        let identifier = reader.read_u32("dictionary property identifier")?;
        if !valid_named_property_identifier(identifier) {
            return Err(invalid(format!(
                "Dictionary property identifier {identifier} is outside the normal range"
            )));
        }
        let length = usize::try_from(reader.read_u32("dictionary string length")?)
            .map_err(|_conversion_error| invalid("Dictionary string length is too large"))?;
        if length == 0 {
            return Err(invalid("Dictionary strings must include a terminator"));
        }
        let name = if codepage == UNICODE_CODEPAGE {
            let payload_units = length - 1;
            let payload_bytes = payload_units
                .checked_mul(2)
                .ok_or_else(|| invalid("Dictionary Unicode length overflow"))?;
            let raw = reader.take(payload_bytes, "dictionary Unicode string")?;
            if reader.read_u16("dictionary Unicode terminator")? != 0 {
                return Err(invalid("Dictionary Unicode string is not terminated"));
            }
            let decoded = decode_utf16(raw, "dictionary Unicode string")?;
            reader.align4(false, "dictionary Unicode padding")?;
            decoded
        } else {
            let raw = reader.take(length - 1, "dictionary string")?;
            if reader.read_u8("dictionary terminator")? != 0 {
                return Err(invalid("Dictionary string is not terminated"));
            }
            decode_ansi(raw, codepage, "dictionary string")?
        };
        if dictionary.insert(identifier, name).is_some() {
            return Err(invalid(format!(
                "Duplicate dictionary property identifier {identifier}"
            )));
        }
        order.push(identifier);
    }
    reader.finish_zero_padding("dictionary padding")?;
    Ok((dictionary, order))
}
