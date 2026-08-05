//! Binary MS-OLEPS stream and VARIANT codec.

use super::super::model::*;
use super::semantic::{MAX_PROPERTY_COUNT, validate_section};
use super::support::allocation;
use chrono::{DateTime, Duration, Utc};
use litchi_cfb::OleError;
use litchi_cfb::consts::*;
use litchi_codepage::Mbcs;
use std::borrow::Cow;
use std::collections::HashMap;

const PROPERTY_SET_HEADER_SIZE: usize = 28;
const SECTION_DESCRIPTOR_SIZE: usize = 20;
const SECTION_HEADER_SIZE: usize = 8;
const PROPERTY_DESCRIPTOR_SIZE: usize = 8;
const MAX_VECTOR_ELEMENTS: usize = 1_000_000;

fn try_zeroed_vec(len: usize, resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut values = try_vec_with_capacity(len, resource)?;
    values.resize(len, 0);
    Ok(values)
}

fn reserve_bytes(
    output: &mut Vec<u8>,
    additional: usize,
    resource: &'static str,
) -> Result<(), OleError> {
    output
        .len()
        .checked_add(additional)
        .ok_or_else(|| invalid("serialized property size overflow"))?;
    output
        .try_reserve(additional)
        .map_err(|source| allocation(resource, source))?;
    Ok(())
}

fn append_bytes(
    output: &mut Vec<u8>,
    bytes: &[u8],
    resource: &'static str,
) -> Result<(), OleError> {
    reserve_bytes(output, bytes.len(), resource)?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_u16(output: &mut Vec<u8>, value: u16, resource: &'static str) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

fn append_u32(output: &mut Vec<u8>, value: u32, resource: &'static str) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

fn append_u64(output: &mut Vec<u8>, value: u64, resource: &'static str) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

impl Stream {
    /// Parse a complete Property Set stream with section-local bounds.
    pub fn parse(data: &[u8]) -> Result<Self, OleError> {
        if data.len() < PROPERTY_SET_HEADER_SIZE + SECTION_DESCRIPTOR_SIZE {
            return Err(invalid("Property Set stream is too short"));
        }
        if read_u16(data, 0, "byte order")? != 0xfffe {
            return Err(invalid("Property Set byte order must be 0xFFFE"));
        }
        let version = read_u16(data, 2, "version")?;
        if !matches!(version, Self::VERSION_0 | Self::VERSION_1) {
            return Err(invalid(format!(
                "Unsupported Property Set version {version}"
            )));
        }
        let system_identifier = read_u32(data, 4, "system identifier")?;
        let class_identifier = read_guid(data, 8, "class identifier")?;
        let section_count = usize::try_from(read_u32(data, 24, "section count")?)
            .map_err(|_| invalid("Property Set section count is too large"))?;
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
                .map_err(|_| invalid("Property Set section offset is too large"))?;
            if offset < descriptor_end || offset % 4 != 0 || offset >= data.len() {
                return Err(invalid(format!(
                    "Invalid Property Set section offset {offset}"
                )));
            }
            if !format_ids.insert(format_identifier) || !offsets.insert(offset) {
                return Err(invalid("Duplicate Property Set section descriptor"));
            }
            let size = usize::try_from(read_u32(data, offset, "section size")?)
                .map_err(|_| invalid("Property Set section size is too large"))?;
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
            )?);
        }
        Ok(Self {
            version,
            system_identifier,
            class_identifier,
            sections,
        })
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        serialize_property_set_stream(self)
    }
}

fn serialize_property_set_stream(stream: &Stream) -> Result<Vec<u8>, OleError> {
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
        validate_section(section, stream.version)?
    }
    let descriptor_end = checked_add(
        PROPERTY_SET_HEADER_SIZE,
        checked_mul(
            stream.sections.len(),
            SECTION_DESCRIPTOR_SIZE,
            "Property Set descriptor table",
        )?,
        "Property Set descriptor table",
    )?;
    let mut sections = try_vec_with_capacity(stream.sections.len(), "serialized sections")?;
    for section in &stream.sections {
        sections.push(serialize_section(section)?)
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
    let mut out = try_zeroed_vec(cursor, "serialized Property Set stream")?;
    out[0..2].copy_from_slice(&0xfffeu16.to_le_bytes());
    out[2..4].copy_from_slice(&stream.version.to_le_bytes());
    out[4..8].copy_from_slice(&stream.system_identifier.to_le_bytes());
    out[8..24].copy_from_slice(stream.class_identifier.as_bytes());
    out[24..28]
        .copy_from_slice(&checked_u32(stream.sections.len(), "section count")?.to_le_bytes());
    for (index, section) in stream.sections.iter().enumerate() {
        let base = checked_add(
            PROPERTY_SET_HEADER_SIZE,
            checked_mul(index, SECTION_DESCRIPTOR_SIZE, "section descriptor")?,
            "section descriptor",
        )?;
        out[base..base + 16].copy_from_slice(section.format_identifier.as_bytes());
        out[base + 16..base + 20]
            .copy_from_slice(&checked_u32(offsets[index], "Section offset")?.to_le_bytes())
    }
    for (index, section) in sections.iter().enumerate() {
        let start = offsets[index];
        let end = checked_add(start, section.len(), "Property Set section")?;
        out[start..end].copy_from_slice(section);
    }
    Ok(out)
}

fn serialize_section(section: &Section) -> Result<Vec<u8>, OleError> {
    let mut order = try_clone_vec(&section.property_order, "property order")?;
    order
        .try_reserve(section.properties.len())
        .map_err(|source| allocation("property order", source))?;
    if !section.dictionary.is_empty() && !order.contains(&PID_DICTIONARY) {
        order.insert(0, PID_DICTIONARY)
    }
    for id in section.properties.keys() {
        if !order.contains(id) {
            order.push(*id)
        }
    }
    order.retain(|id| {
        *id == PID_DICTIONARY && !section.dictionary.is_empty()
            || section.properties.contains_key(id)
    });
    let mut order_ids = try_hash_set_with_capacity(order.len(), "property order set")?;
    for identifier in order.iter().copied() {
        order_ids.insert(identifier);
    }
    if order_ids.len() != order.len() {
        return Err(invalid("Duplicate property order identifier"));
    }
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
            serialize_dictionary(section, codepage)?
        } else {
            let value = section
                .properties
                .get(id)
                .ok_or_else(|| invalid(format!("Property order references missing PID {id}")))?;
            serialize_typed(value, codepage)?
        })
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
    let mut out = try_zeroed_vec(cursor, "serialized Property Set section")?;
    out[4..8].copy_from_slice(&checked_u32(order.len(), "property count")?.to_le_bytes());
    for (index, id) in order.iter().enumerate() {
        let base = checked_add(
            SECTION_HEADER_SIZE,
            checked_mul(index, PROPERTY_DESCRIPTOR_SIZE, "property descriptor")?,
            "property descriptor",
        )?;
        out[base..base + 4].copy_from_slice(&id.to_le_bytes());
        out[base + 4..base + 8]
            .copy_from_slice(&checked_u32(offsets[index], "Property offset")?.to_le_bytes())
    }
    for (index, value) in values.iter().enumerate() {
        let start = offsets[index];
        let end = checked_add(start, value.len(), "Property Set value")?;
        out[start..end].copy_from_slice(value);
    }
    let section_len = checked_u32(out.len(), "Section length")?;
    out[0..4].copy_from_slice(&section_len.to_le_bytes());
    Ok(out)
}

fn serialize_dictionary(section: &Section, codepage: u16) -> Result<Vec<u8>, OleError> {
    let mut order = try_clone_vec(&section.dictionary_order, "dictionary order")?;
    order
        .try_reserve(section.dictionary.len())
        .map_err(|source| allocation("dictionary order", source))?;
    for id in section.dictionary.keys() {
        if !order.contains(id) {
            order.push(*id)
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
    let mut out = try_vec_with_capacity(4, "serialized property dictionary")?;
    append_u32(
        &mut out,
        checked_u32(order.len(), "dictionary entry count")?,
        "serialized property dictionary",
    )?;
    for id in order {
        let name = section
            .dictionary
            .get(&id)
            .ok_or_else(|| invalid(format!("Dictionary order references missing PID {id}")))?;
        append_u32(&mut out, id, "serialized property dictionary")?;
        if codepage == UNICODE_CODEPAGE {
            let units = checked_add(name.encode_utf16().count(), 1, "Dictionary name length")?;
            let byte_len = checked_mul(units, 2, "Dictionary name length")?;
            append_u32(
                &mut out,
                checked_u32(units, "Dictionary name length")?,
                "serialized property dictionary",
            )?;
            reserve_bytes(&mut out, byte_len, "serialized property dictionary")?;
            for unit in name.encode_utf16() {
                out.extend_from_slice(&unit.to_le_bytes());
            }
            out.extend_from_slice(&0u16.to_le_bytes());
            pad4(&mut out)?;
        } else {
            let bytes = encode_ansi(name, codepage)?;
            let byte_len = checked_add(bytes.len(), 1, "Dictionary name length")?;
            append_u32(
                &mut out,
                checked_u32(byte_len, "Dictionary name length")?,
                "serialized property dictionary",
            )?;
            reserve_bytes(&mut out, byte_len, "serialized property dictionary")?;
            out.extend_from_slice(&bytes);
            out.push(0)
        }
    }
    Ok(out)
}

fn serialize_typed(value: &Value, codepage: u16) -> Result<Vec<u8>, OleError> {
    let mut out = try_vec_with_capacity(4, "serialized property value")?;
    append_typed(&mut out, value, codepage)?;
    Ok(out)
}
fn append_typed(out: &mut Vec<u8>, value: &Value, codepage: u16) -> Result<(), OleError> {
    let vt = variant_type(value);
    append_u16(out, vt, "serialized property value")?;
    append_u16(out, 0, "serialized property value")?;
    append_body(out, value, codepage)?;
    Ok(())
}
fn variant_type(value: &Value) -> u16 {
    match value {
        Value::Empty => VT_EMPTY,
        Value::Null => VT_NULL,
        Value::I1(_) => VT_I1,
        Value::UI1(_) => VT_UI1,
        Value::I2(_) => VT_I2,
        Value::UI2(_) => VT_UI2,
        Value::I4(_) => VT_I4,
        Value::UI4(_) => VT_UI4,
        Value::I8(_) => VT_I8,
        Value::UI8(_) => VT_UI8,
        Value::Int(_) => VT_INT,
        Value::UInt(_) => VT_UINT,
        Value::R4(_) => VT_R4,
        Value::R8(_) => VT_R8,
        Value::Currency(_) => VT_CY,
        Value::Date(_) => VT_DATE,
        Value::Bstr(_) => VT_BSTR,
        Value::Error(_) => VT_ERROR,
        Value::Bool(_) => VT_BOOL,
        Value::Decimal(_) => VT_DECIMAL,
        Value::Lpstr(_) => VT_LPSTR,
        Value::Lpwstr(_) => VT_LPWSTR,
        Value::Filetime(_) => VT_FILETIME,
        Value::Blob(_) => VT_BLOB,
        Value::Clipboard { .. } => VT_CF,
        Value::Clsid(_) => VT_CLSID,
        Value::Vector(_) => VT_VECTOR | VT_VARIANT,
        Value::Unknown { variant_type, .. } => *variant_type,
    }
}
fn append_body(out: &mut Vec<u8>, value: &Value, codepage: u16) -> Result<(), OleError> {
    match value {
        Value::Empty | Value::Null => {},
        Value::I1(v) => append_bytes(out, &[*v as u8], "serialized property value")?,
        Value::UI1(v) => append_bytes(out, &[*v], "serialized property value")?,
        Value::I2(v) => append_bytes(out, &v.to_le_bytes(), "serialized property value")?,
        Value::UI2(v) => append_bytes(out, &v.to_le_bytes(), "serialized property value")?,
        Value::I4(v) | Value::Int(v) => {
            append_bytes(out, &v.to_le_bytes(), "serialized property value")?
        },
        Value::UI4(v) | Value::UInt(v) | Value::Error(v) => {
            append_u32(out, *v, "serialized property value")?
        },
        Value::I8(v) | Value::Currency(v) => {
            append_bytes(out, &v.to_le_bytes(), "serialized property value")?
        },
        Value::UI8(v) | Value::Filetime(v) => append_u64(out, *v, "serialized property value")?,
        Value::R4(v) => append_u32(out, v.to_bits(), "serialized property value")?,
        Value::R8(v) | Value::Date(v) => append_u64(out, v.to_bits(), "serialized property value")?,
        Value::Bool(v) => append_bytes(
            out,
            &(if *v { -1i16 } else { 0 }).to_le_bytes(),
            "serialized property value",
        )?,
        Value::Decimal(v) => append_bytes(out, v, "serialized property value")?,
        Value::Bstr(v) | Value::Lpstr(v) => append_codepage_string(out, v, codepage)?,
        Value::Lpwstr(v) => {
            let units = checked_add(v.encode_utf16().count(), 1, "LPWSTR length")?;
            let byte_len = checked_mul(units, 2, "LPWSTR length")?;
            append_u32(
                out,
                checked_u32(units, "LPWSTR length")?,
                "serialized property value",
            )?;
            reserve_bytes(out, byte_len, "serialized property value")?;
            for unit in v.encode_utf16() {
                out.extend_from_slice(&unit.to_le_bytes());
            }
            out.extend_from_slice(&0u16.to_le_bytes());
        },
        Value::Blob(v) => {
            append_u32(
                out,
                checked_u32(v.len(), "Blob size")?,
                "serialized property value",
            )?;
            append_bytes(out, v, "serialized property value")?;
            pad4(out)?;
        },
        Value::Clipboard { format, data } => {
            let size = checked_add(data.len(), 4, "Clipboard size")?;
            append_u32(
                out,
                checked_u32(size, "Clipboard size")?,
                "serialized property value",
            )?;
            append_bytes(out, &format.to_le_bytes(), "serialized property value")?;
            append_bytes(out, data, "serialized property value")?;
            pad4(out)?;
        },
        Value::Clsid(v) => append_bytes(out, v.as_bytes(), "serialized property value")?,
        Value::Vector(values) => {
            if values.len() > MAX_VECTOR_ELEMENTS {
                return Err(invalid("Vector exceeds safety limit"));
            }
            append_u32(
                out,
                checked_u32(values.len(), "Vector element count")?,
                "serialized property value",
            )?;
            for value in values {
                append_typed(out, value, codepage)?;
                pad4(out)?;
            }
        },
        Value::Unknown { data, .. } => append_bytes(out, data, "serialized property value")?,
    }
    Ok(())
}
fn append_codepage_string(out: &mut Vec<u8>, value: &str, codepage: u16) -> Result<(), OleError> {
    if codepage == UNICODE_CODEPAGE {
        let units = checked_add(value.encode_utf16().count(), 1, "CodePageString length")?;
        let byte_len = checked_mul(units, 2, "CodePageString length")?;
        append_u32(
            out,
            checked_u32(byte_len, "CodePageString length")?,
            "serialized CodePageString",
        )?;
        reserve_bytes(out, byte_len, "serialized CodePageString")?;
        for unit in value.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out.extend_from_slice(&[0, 0]);
    } else {
        let bytes = encode_ansi(value, codepage)?;
        let byte_len = checked_add(bytes.len(), 1, "CodePageString length")?;
        append_u32(
            out,
            checked_u32(byte_len, "CodePageString length")?,
            "serialized CodePageString",
        )?;
        reserve_bytes(out, byte_len, "serialized CodePageString")?;
        out.extend_from_slice(&bytes);
        out.push(0);
    }
    pad4(out)?;
    Ok(())
}
fn pad4(out: &mut Vec<u8>) -> Result<(), OleError> {
    let padding = (4 - (out.len() & 3)) & 3;
    reserve_bytes(out, padding, "serialized property padding")?;
    for _ in 0..padding {
        out.push(0);
    }
    Ok(())
}
fn encode_ansi(value: &str, codepage: u16) -> Result<Vec<u8>, OleError> {
    let page = Mbcs::require(u32::from(codepage)).map_err(|error| invalid(error.to_string()))?;
    page.encode(value)
        .map(Cow::into_owned)
        .map_err(|error| invalid(error.to_string()))
}

fn parse_section(data: &[u8], format_identifier: Guid, version: u16) -> Result<Section, OleError> {
    checked_range(data, 0, SECTION_HEADER_SIZE, "section header")?;
    let declared_size = usize::try_from(read_u32(data, 0, "section size")?)
        .map_err(|_| invalid("Property Set section size is too large"))?;
    if declared_size != data.len() {
        return Err(invalid(
            "Property Set section range does not match its size",
        ));
    }
    let property_count = usize::try_from(read_u32(data, 4, "property count")?)
        .map_err(|_| invalid("Property count is too large"))?;
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
            .map_err(|_| invalid("Property offset is too large"))?;
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
            let codepage = CodePage::try_from(signed as u16)?;
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
        let value = parse_typed_property(bytes, effective_codepage, *start).map_err(|error| {
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
        .map_err(|_| invalid("Dictionary count is too large"))?;
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
            .map_err(|_| invalid("Dictionary string length is too large"))?;
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

pub(crate) fn parse_typed_property(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<Value, OleError> {
    if data.len() < 4 {
        return Err(invalid("Typed property value is truncated"));
    }
    let variant_type = read_u16(data, 0, "property variant type")?;
    if read_u16(data, 2, "property reserved field")? != 0 {
        return Err(invalid("Typed property reserved field must be zero"));
    }
    let value_offset = property_offset
        .checked_add(4)
        .ok_or_else(|| invalid("Property value offset overflow"))?;
    let mut reader = ValueReader::new(&data[4..], value_offset);
    let value = parse_value_body(&mut reader, variant_type, codepage, 0)?;
    let allows_opaque_tail = matches!(
        variant_type,
        VT_EMPTY | VT_NULL | VT_BSTR | VT_LPSTR | VT_LPWSTR
    );
    if reader.remaining_len() <= 3 && allows_opaque_tail {
        // Some producers retain up to three opaque bytes after a top-level
        // value, especially when a non-DWORD property offset is used. The
        // property offset table still bounds the value.
        reader.take_remaining();
    } else {
        reader
            .finish_zero_padding("property padding")
            .map_err(|error| {
                invalid(format!(
                    "variant 0x{variant_type:04x} has {} trailing bytes: {error}",
                    reader.remaining_len()
                ))
            })?;
    }
    Ok(value)
}

fn parse_value_body(
    reader: &mut ValueReader<'_>,
    variant_type: u16,
    codepage: u16,
    depth: usize,
) -> Result<Value, OleError> {
    if depth > 8 {
        return Err(invalid("Property value nesting exceeds the safety limit"));
    }
    if variant_type & VT_VECTOR != 0 {
        let base_type = variant_type & !VT_VECTOR;
        if !is_supported_vector_base(base_type) {
            if depth != 0 {
                return Err(invalid(format!(
                    "Unsupported nested vector type 0x{variant_type:04x}"
                )));
            }
            return Ok(Value::Unknown {
                variant_type,
                data: try_copy_bytes(reader.take_remaining(), "unknown vector property")?,
            });
        }
        return parse_vector(reader, base_type, codepage, depth + 1);
    }
    match variant_type {
        VT_EMPTY => Ok(Value::Empty),
        VT_NULL => Ok(Value::Null),
        VT_I1 => Ok(Value::I1(reader.read_i8("I1 value")?)),
        VT_UI1 => Ok(Value::UI1(reader.read_u8("UI1 value")?)),
        VT_I2 => Ok(Value::I2(reader.read_i16("I2 value")?)),
        VT_UI2 => Ok(Value::UI2(reader.read_u16("UI2 value")?)),
        VT_I4 => Ok(Value::I4(reader.read_i32("I4 value")?)),
        VT_UI4 => Ok(Value::UI4(reader.read_u32("UI4 value")?)),
        VT_I8 => Ok(Value::I8(reader.read_i64("I8 value")?)),
        VT_UI8 => Ok(Value::UI8(reader.read_u64("UI8 value")?)),
        VT_INT => Ok(Value::Int(reader.read_i32("INT value")?)),
        VT_UINT => Ok(Value::UInt(reader.read_u32("UINT value")?)),
        VT_R4 => Ok(Value::R4(f32::from_bits(reader.read_u32("R4 value")?))),
        VT_R8 => Ok(Value::R8(f64::from_bits(reader.read_u64("R8 value")?))),
        VT_CY => Ok(Value::Currency(reader.read_i64("CY value")?)),
        VT_DATE => Ok(Value::Date(f64::from_bits(reader.read_u64("DATE value")?))),
        VT_BSTR => Ok(Value::Bstr(read_codepage_string(
            reader,
            codepage,
            "BSTR value",
            depth == 0,
        )?)),
        VT_ERROR => Ok(Value::Error(reader.read_u32("ERROR value")?)),
        // Office producers sometimes use 1 for TRUE even though VARIANT_BOOL
        // conventionally uses -1. Preserve the established nonzero mapping.
        VT_BOOL => Ok(Value::Bool(reader.read_i16("BOOL value")? != 0)),
        VT_DECIMAL => {
            let raw = reader.take(16, "DECIMAL value")?;
            let mut value = [0u8; 16];
            value.copy_from_slice(raw);
            Ok(Value::Decimal(value))
        },
        VT_LPSTR => Ok(Value::Lpstr(read_codepage_string(
            reader,
            codepage,
            "LPSTR value",
            depth == 0,
        )?)),
        VT_LPWSTR => Ok(Value::Lpwstr(read_unicode_string(
            reader,
            "LPWSTR value",
            depth == 0,
        )?)),
        VT_FILETIME => Ok(Value::Filetime(reader.read_u64("FILETIME value")?)),
        VT_BLOB | VT_BLOB_OBJECT => {
            let size = usize::try_from(reader.read_u32("blob size")?)
                .map_err(|_| invalid("Blob size is too large"))?;
            let blob = try_copy_bytes(reader.take(size, "blob data")?, "blob data")?;
            reader.align4(depth == 0, "blob padding")?;
            Ok(Value::Blob(blob))
        },
        VT_CF => {
            let size = usize::try_from(reader.read_u32("clipboard size")?)
                .map_err(|_| invalid("Clipboard size is too large"))?;
            if size < 4 {
                return Err(invalid("Clipboard size must include its format field"));
            }
            let format = reader.read_i32("clipboard format")?;
            let payload =
                try_copy_bytes(reader.take(size - 4, "clipboard data")?, "clipboard data")?;
            reader.align4(depth == 0, "clipboard padding")?;
            Ok(Value::Clipboard {
                format,
                data: payload,
            })
        },
        VT_CLSID => {
            let raw = reader.take(16, "CLSID value")?;
            let mut value = [0u8; 16];
            value.copy_from_slice(raw);
            Ok(Value::Clsid(Guid::from_bytes(value)))
        },
        _ if depth == 0 => Ok(Value::Unknown {
            variant_type,
            data: try_copy_bytes(reader.take_remaining(), "unknown property value")?,
        }),
        _ => Err(invalid(format!(
            "Unsupported nested property variant 0x{variant_type:04x}"
        ))),
    }
}

fn parse_vector(
    reader: &mut ValueReader<'_>,
    base_type: u16,
    codepage: u16,
    depth: usize,
) -> Result<Value, OleError> {
    let count = usize::try_from(reader.read_u32("vector element count")?)
        .map_err(|_| invalid("Vector element count is too large"))?;
    if count > MAX_VECTOR_ELEMENTS {
        return Err(invalid(format!(
            "Vector element count {count} exceeds the safety limit"
        )));
    }
    let minimum = minimum_value_size(base_type).unwrap_or(0);
    if minimum != 0
        && count
            .checked_mul(minimum)
            .is_none_or(|size| size > reader.remaining_len())
    {
        return Err(invalid("Vector element count exceeds its property range"));
    }
    if base_type == VT_VARIANT && count > reader.remaining_len() / 4 {
        return Err(invalid("Variant vector count exceeds its property range"));
    }
    let mut values = try_vec_with_capacity(count, "property vector")?;
    for _ in 0..count {
        let value = if base_type == VT_VARIANT {
            let nested_type = reader.read_u16("variant vector element type")?;
            if reader.read_u16("variant vector reserved field")? != 0 {
                return Err(invalid("Variant vector reserved field must be zero"));
            }
            let value = parse_value_body(reader, nested_type, codepage, depth + 1)?;
            reader.align4(false, "variant vector element padding")?;
            value
        } else {
            parse_value_body(reader, base_type, codepage, depth + 1)?
        };
        values.push(value);
    }
    Ok(Value::Vector(values))
}

fn is_supported_vector_base(variant_type: u16) -> bool {
    matches!(
        variant_type,
        VT_I1
            | VT_UI1
            | VT_I2
            | VT_UI2
            | VT_I4
            | VT_UI4
            | VT_I8
            | VT_UI8
            | VT_R4
            | VT_R8
            | VT_CY
            | VT_DATE
            | VT_BSTR
            | VT_ERROR
            | VT_BOOL
            | VT_VARIANT
            | VT_LPSTR
            | VT_LPWSTR
            | VT_FILETIME
            | VT_CF
            | VT_CLSID
    )
}

fn minimum_value_size(variant_type: u16) -> Option<usize> {
    match variant_type {
        VT_I1 | VT_UI1 => Some(1),
        VT_I2 | VT_UI2 | VT_BOOL => Some(2),
        VT_I4 | VT_UI4 | VT_INT | VT_UINT | VT_R4 | VT_ERROR => Some(4),
        VT_I8 | VT_UI8 | VT_R8 | VT_CY | VT_DATE | VT_FILETIME => Some(8),
        VT_CLSID => Some(16),
        VT_VARIANT | VT_BSTR | VT_LPSTR | VT_LPWSTR | VT_CF => Some(4),
        _ => None,
    }
}

fn read_codepage_string(
    reader: &mut ValueReader<'_>,
    codepage: u16,
    description: &str,
    top_level: bool,
) -> Result<String, OleError> {
    let size = usize::try_from(reader.read_u32(description)?)
        .map_err(|_| invalid(format!("{description} is too large")))?;
    let raw = reader.take(size, description)?;
    let value = if size == 0 {
        String::new()
    } else if codepage == UNICODE_CODEPAGE {
        if size % 2 != 0 || !raw.ends_with(&[0, 0]) {
            return Err(invalid(format!("{description} is not terminated UTF-16LE")));
        }
        let end = raw
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .map_or(raw.len(), |units| units * 2);
        decode_utf16(&raw[..end], description)?
    } else {
        if raw.last() != Some(&0) {
            return Err(invalid(format!("{description} is not NUL-terminated")));
        }
        let end = raw.iter().position(|byte| *byte == 0).unwrap_or(raw.len());
        decode_ansi(&raw[..end], codepage, description)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

fn read_unicode_string(
    reader: &mut ValueReader<'_>,
    description: &str,
    top_level: bool,
) -> Result<String, OleError> {
    let units = usize::try_from(reader.read_u32(description)?)
        .map_err(|_| invalid(format!("{description} is too large")))?;
    let byte_len = units
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{description} length overflow")))?;
    let raw = reader.take(byte_len, description)?;
    let value = if units == 0 {
        String::new()
    } else {
        if !raw.ends_with(&[0, 0]) {
            return Err(invalid(format!("{description} is not NUL-terminated")));
        }
        let end = raw
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .map_or(raw.len(), |units| units * 2);
        decode_utf16(&raw[..end], description)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

fn decode_utf16(data: &[u8], description: &str) -> Result<String, OleError> {
    if !data.len().is_multiple_of(2) {
        return Err(invalid(format!("{description} has an odd byte length")));
    }
    let mut utf8_len = 0usize;
    for decoded in std::char::decode_utf16(
        data.chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
    ) {
        let character =
            decoded.map_err(|_| invalid(format!("{description} contains invalid UTF-16")))?;
        utf8_len = checked_add(utf8_len, character.len_utf8(), description)?;
    }
    let mut value = String::new();
    value
        .try_reserve_exact(utf8_len)
        .map_err(|source| allocation("decoded UTF-16 string", source))?;
    for decoded in std::char::decode_utf16(
        data.chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
    ) {
        let character =
            decoded.map_err(|_| invalid(format!("{description} contains invalid UTF-16")))?;
        value.push(character);
    }
    Ok(value)
}

fn decode_ansi(data: &[u8], codepage: u16, description: &str) -> Result<String, OleError> {
    let page = Mbcs::require(u32::from(codepage))
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))?;
    let decoded = page
        .decode(data)
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))?;
    match decoded {
        Cow::Borrowed(value) => try_clone_string(value, "decoded ANSI string"),
        Cow::Owned(value) => Ok(value),
    }
}

struct ValueReader<'a> {
    data: &'a [u8],
    position: usize,
    alignment_base: usize,
}

impl<'a> ValueReader<'a> {
    const fn new(data: &'a [u8], alignment_base: usize) -> Self {
        Self {
            data,
            position: 0,
            alignment_base,
        }
    }

    fn remaining_len(&self) -> usize {
        self.data.len() - self.position
    }

    fn take(&mut self, length: usize, description: &str) -> Result<&'a [u8], OleError> {
        let start = self.position;
        let end = start
            .checked_add(length)
            .filter(|end| *end <= self.data.len())
            .ok_or_else(|| invalid(format!("{description} exceeds its property range")))?;
        self.position = end;
        Ok(&self.data[start..end])
    }

    fn take_remaining(&mut self) -> &'a [u8] {
        let remaining = &self.data[self.position..];
        self.position = self.data.len();
        remaining
    }

    fn read_u8(&mut self, description: &str) -> Result<u8, OleError> {
        Ok(self.take(1, description)?[0])
    }

    fn read_i8(&mut self, description: &str) -> Result<i8, OleError> {
        Ok(self.read_u8(description)? as i8)
    }

    fn read_u16(&mut self, description: &str) -> Result<u16, OleError> {
        let bytes = self.take(2, description)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_i16(&mut self, description: &str) -> Result<i16, OleError> {
        let bytes = self.take(2, description)?;
        Ok(i16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_u32(&mut self, description: &str) -> Result<u32, OleError> {
        let bytes = self.take(4, description)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn read_i32(&mut self, description: &str) -> Result<i32, OleError> {
        let bytes = self.take(4, description)?;
        Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn read_u64(&mut self, description: &str) -> Result<u64, OleError> {
        let bytes = self.take(8, description)?;
        Ok(u64::from_le_bytes([
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
        ]))
    }

    fn read_i64(&mut self, description: &str) -> Result<i64, OleError> {
        let bytes = self.take(8, description)?;
        Ok(i64::from_le_bytes([
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
        ]))
    }

    fn align4(&mut self, top_level: bool, description: &str) -> Result<(), OleError> {
        let absolute_position = self
            .alignment_base
            .checked_add(self.position)
            .ok_or_else(|| invalid(format!("{description} position overflow")))?;
        let padding = (4 - (absolute_position & 3)) & 3;
        let available = padding.min(self.remaining_len());
        let end = self
            .position
            .checked_add(available)
            .ok_or_else(|| invalid(format!("{description} range overflow")))?;
        let candidate = &self.data[self.position..end];
        let consumed = if top_level {
            // Top-level property offsets are authoritative. Several Office
            // producers omit filler or write nonzero filler between values.
            available
        } else {
            // Inside a vector there is no offset table. Match Office readers:
            // skip zero filler only and stop before the next nonzero field.
            candidate.iter().take_while(|byte| **byte == 0).count()
        };
        self.take(consumed, description)?;
        Ok(())
    }

    fn finish_zero_padding(&mut self, description: &str) -> Result<(), OleError> {
        let remaining = &self.data[self.position..];
        if remaining.iter().any(|byte| *byte != 0) {
            return Err(invalid(format!("{description} must be zero")));
        }
        self.position = self.data.len();
        Ok(())
    }
}

fn checked_range<'a>(
    data: &'a [u8],
    offset: usize,
    length: usize,
    description: &str,
) -> Result<&'a [u8], OleError> {
    let end = offset
        .checked_add(length)
        .filter(|end| *end <= data.len())
        .ok_or_else(|| invalid(format!("{description} exceeds its enclosing range")))?;
    Ok(&data[offset..end])
}

fn read_u16(data: &[u8], offset: usize, description: &str) -> Result<u16, OleError> {
    let bytes = checked_range(data, offset, 2, description)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, description: &str) -> Result<u32, OleError> {
    let bytes = checked_range(data, offset, 4, description)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_guid(data: &[u8], offset: usize, description: &str) -> Result<Guid, OleError> {
    let bytes = checked_range(data, offset, 16, description)?;
    let mut guid = [0u8; 16];
    guid.copy_from_slice(bytes);
    Ok(Guid::from_bytes(guid))
}

pub(crate) fn filetime_to_date(filetime: u64) -> Option<DateTime<Utc>> {
    const EPOCH_DIFF: i64 = 116_444_736_000_000_000;
    let doc_epoch = i64::try_from(filetime).ok()?;
    let nanos = doc_epoch.checked_sub(EPOCH_DIFF)?.checked_mul(100)?;
    Some(DateTime::from_timestamp_nanos(nanos))
}

pub(crate) fn filetime_to_duration(filetime: u64) -> Option<Duration> {
    let nanos = filetime.checked_mul(100)?;
    Some(Duration::nanoseconds(i64::try_from(nanos).ok()?))
}
