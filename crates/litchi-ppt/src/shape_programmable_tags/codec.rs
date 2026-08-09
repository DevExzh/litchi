//! Binary record codec for shape-scoped programmable tags.
//!
//! This module implements MS-PPT sections 2.7.14 through 2.7.20. Parsing is
//! inert: binary tag payloads are never executed, loaded, or resolved.

use crate::consts::RecordType;

use super::model::{
    ShapeBinaryTag, ShapeBinaryTagPayload, ShapeBinaryTagVersion, ShapeProgrammableTag,
    ShapeProgrammableTagLimits, ShapeProgrammableTags, ShapeStringTag, ShapeStyleAtom,
};
use crate::package::{Error, Result};
use crate::records::Record;
use crate::text_extensions::{TextStyleExtension9, TextStyleExtension10, TextStyleExtension11};

impl ShapeProgrammableTags {
    /// Parse and validate a complete `ShapeProgTagsContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record, limits: ShapeProgrammableTagLimits) -> Result<Self> {
        if record.record_type != RecordType::ProgTags
            || record.record_type_raw != RecordType::ProgTags.as_u16()
            || record.version != 0x0f
        {
            return corrupted("Invalid ShapeProgTagsContainer record header");
        }
        let declared = usize::try_from(record.data_length)
            .map_err(|_err| Error::Corrupted("ShapeProgTagsContainer size overflow".into()))?;
        if declared != record.data.len() {
            return corrupted("ShapeProgTagsContainer length does not match its payload");
        }
        Self::parse_payload(&record.data, record.instance, limits)
    }

    /// Parse a `ShapeProgTagsContainer` payload and its record instance.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_payload(
        data: &[u8],
        instance: u16,
        limits: ShapeProgrammableTagLimits,
    ) -> Result<Self> {
        check_limit(
            data.len(),
            limits.max_container_bytes,
            "ShapeProgTagsContainer payload",
        )?;
        let records = parse_sequence(data, limits.max_tags, "ShapeProgTagsContainer")?;
        let mut tags = Vec::with_capacity(records.len());
        let mut seen9 = false;
        let mut seen10 = false;
        let mut seen11 = false;

        for record in records {
            check_limit(
                record.data.len(),
                limits.max_tag_bytes,
                "shape programmable tag payload",
            )?;
            let tag = match record.record_type {
                value if value == RecordType::ProgStringTag.as_u16() => {
                    ShapeProgrammableTag::String(parse_string_tag(&record, limits)?)
                },
                value if value == RecordType::ProgBinaryTag.as_u16() => {
                    let tag = parse_binary_tag(&record, limits)?;
                    let duplicate = match tag.version {
                        ShapeBinaryTagVersion::PowerPoint9 => std::mem::replace(&mut seen9, true),
                        ShapeBinaryTagVersion::PowerPoint10 => std::mem::replace(&mut seen10, true),
                        ShapeBinaryTagVersion::PowerPoint11 => std::mem::replace(&mut seen11, true),
                        ShapeBinaryTagVersion::Unknown => false,
                    };
                    if duplicate {
                        return corrupted(
                            "ShapeProgTagsContainer contains a duplicate versioned binary tag",
                        );
                    }
                    ShapeProgrammableTag::Binary(tag)
                },
                _ => {
                    return corrupted(
                        "ShapeProgTagsContainer contains a disallowed child record type",
                    );
                },
            };
            tags.push(tag);
        }
        Ok(Self { instance, tags })
    }

    /// Serialize a complete `ShapeProgTagsContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn to_bytes(&self, limits: ShapeProgrammableTagLimits) -> Result<Vec<u8>> {
        let payload = self.to_payload(limits)?;
        encode_record(0x0f, self.instance, RecordType::ProgTags.as_u16(), &payload)
    }

    /// Serialize only the `ShapeProgTagsContainer` payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_payload(&self, limits: ShapeProgrammableTagLimits) -> Result<Vec<u8>> {
        check_limit(
            self.tags.len(),
            limits.max_tags,
            "shape programmable tag count",
        )?;
        let mut payload = Vec::new();
        for tag in &self.tags {
            let encoded = match tag {
                ShapeProgrammableTag::String(string_tag) => encode_string_tag(string_tag)?,
                ShapeProgrammableTag::Binary(binary_tag) => encode_binary_tag(binary_tag)?,
            };
            check_limit(
                encoded.len().saturating_sub(8),
                limits.max_tag_bytes,
                "shape programmable tag payload",
            )?;
            payload.extend_from_slice(&encoded);
        }
        check_limit(
            payload.len(),
            limits.max_container_bytes,
            "ShapeProgTagsContainer payload",
        )?;
        // Reparse before returning so public-field mutations cannot serialize an
        // invalid duplicate/version combination or bypass any current limit.
        Self::parse_payload(&payload, self.instance, limits)?;
        Ok(payload)
    }

    /// Build a generic PPT record for insertion into `OfficeArt` client data.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_record(&self, limits: ShapeProgrammableTagLimits) -> Result<Record> {
        let data = self.to_payload(limits)?;
        let data_length = u32::try_from(data.len()).map_err(|_err| {
            Error::Corrupted("ShapeProgTagsContainer payload exceeds u32".into())
        })?;
        Ok(Record {
            record_type: RecordType::ProgTags,
            record_type_raw: RecordType::ProgTags.as_u16(),
            version: 0x0f,
            instance: self.instance,
            data_length,
            data,
            children: Vec::new(),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawRecord {
    version: u16,
    instance: u16,
    record_type: u16,
    data: Vec<u8>,
}

impl From<RawRecord> for ShapeStyleAtom {
    fn from(value: RawRecord) -> Self {
        Self {
            version: value.version,
            instance: value.instance,
            record_type: value.record_type,
            data: value.data,
        }
    }
}

fn parse_string_tag(
    record: &RawRecord,
    limits: ShapeProgrammableTagLimits,
) -> Result<ShapeStringTag> {
    require_container_header(record, RecordType::ProgStringTag, "ProgStringTagContainer")?;
    let children = parse_sequence(&record.data, 2, "ProgStringTagContainer")?;
    if children.is_empty() || children.len() > 2 {
        return corrupted("ProgStringTagContainer must contain a name and at most one value");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let value_units = children
        .get(1)
        .map(|value| parse_cstring_atom(value, 1, false, limits))
        .transpose()?;
    Ok(ShapeStringTag {
        name: decode_units(&name_units, "programmable tag name")?,
        value: value_units
            .as_ref()
            .map(|units| decode_units(units, "programmable tag value"))
            .transpose()?,
        name_units,
        value_units,
    })
}

fn parse_binary_tag(
    record: &RawRecord,
    limits: ShapeProgrammableTagLimits,
) -> Result<ShapeBinaryTag> {
    require_container_header(
        record,
        RecordType::ProgBinaryTag,
        "ShapeProgBinaryTagContainer",
    )?;
    let children = parse_sequence(&record.data, 2, "ShapeProgBinaryTagContainer")?;
    if children.len() != 2 {
        return corrupted("ShapeProgBinaryTagContainer must contain exactly one CString/blob pair");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let name = decode_units(&name_units, "binary programmable tag name")?;
    let blob = &children[1];
    require_atom_header(blob, RecordType::BinaryTagData, 0, "BinaryTagDataBlob")?;

    let (version, payload) = match name.as_str() {
        "___PPT9" => (
            ShapeBinaryTagVersion::PowerPoint9,
            parse_known_payload(
                blob,
                RecordType::StyleTextProp9Atom,
                14,
                &name_units,
                limits,
                |atom| {
                    let style = TextStyleExtension9::parse(&atom.data)?;
                    check_limit(
                        style.runs.len(),
                        limits.max_style_runs,
                        "PowerPoint 9 style run count",
                    )?;
                    Ok(ShapeBinaryTagPayload::PowerPoint9 {
                        style,
                        atom: atom.into(),
                    })
                },
            )?,
        ),
        "___PPT10" => (
            ShapeBinaryTagVersion::PowerPoint10,
            parse_known_payload(
                blob,
                RecordType::StyleTextProp10Atom,
                16,
                &name_units,
                limits,
                |atom| {
                    let style = TextStyleExtension10::parse(&atom.data)?;
                    check_limit(
                        style.runs.len(),
                        limits.max_style_runs,
                        "PowerPoint 10 style run count",
                    )?;
                    Ok(ShapeBinaryTagPayload::PowerPoint10 {
                        style,
                        atom: atom.into(),
                    })
                },
            )?,
        ),
        "___PPT11" => (
            ShapeBinaryTagVersion::PowerPoint11,
            parse_known_payload(
                blob,
                RecordType::StyleTextProp11Atom,
                16,
                &name_units,
                limits,
                |atom| {
                    let style = TextStyleExtension11::parse(&atom.data)?;
                    check_limit(
                        style.runs.len(),
                        limits.max_style_runs,
                        "PowerPoint 11 style run count",
                    )?;
                    Ok(ShapeBinaryTagPayload::PowerPoint11 {
                        style,
                        atom: atom.into(),
                    })
                },
            )?,
        ),
        _ => {
            check_limit(
                blob.data.len(),
                limits.max_unknown_binary_bytes,
                "unknown binary programmable tag payload",
            )?;
            (
                ShapeBinaryTagVersion::Unknown,
                ShapeBinaryTagPayload::Unknown(blob.data.clone()),
            )
        },
    };

    Ok(ShapeBinaryTag {
        name,
        version,
        payload,
        name_units,
    })
}

fn parse_known_payload<F>(
    blob: &RawRecord,
    expected_type: RecordType,
    expected_name_bytes: usize,
    name_units: &[u16],
    limits: ShapeProgrammableTagLimits,
    decode: F,
) -> Result<ShapeBinaryTagPayload>
where
    F: FnOnce(RawRecord) -> Result<ShapeBinaryTagPayload>,
{
    if name_units.len().checked_mul(2) != Some(expected_name_bytes) {
        return corrupted("Versioned shape binary tag has an invalid CString length");
    }
    check_limit(
        blob.data.len(),
        limits.max_style_payload_bytes.saturating_add(8),
        "versioned shape binary tag blob",
    )?;
    let (atom, consumed) = parse_one(&blob.data, 0, "versioned shape binary tag blob")?;
    if consumed != blob.data.len() {
        return corrupted("Versioned shape binary tag must contain exactly one style atom");
    }
    require_atom_header(&atom, expected_type, 0, "versioned shape style atom")?;
    check_limit(
        atom.data.len(),
        limits.max_style_payload_bytes,
        "versioned shape style payload",
    )?;
    decode(atom)
}

fn encode_string_tag(tag: &ShapeStringTag) -> Result<Vec<u8>> {
    let name = encode_cstring_atom(0, &tag.name_units)?;
    let mut payload = name;
    if let Some(value) = &tag.value_units {
        payload.extend_from_slice(&encode_cstring_atom(1, value)?);
    }
    encode_record(0x0f, 0, RecordType::ProgStringTag.as_u16(), &payload)
}

fn encode_binary_tag(tag: &ShapeBinaryTag) -> Result<Vec<u8>> {
    let mut payload = encode_cstring_atom(0, &tag.name_units)?;
    let blob_data = match &tag.payload {
        ShapeBinaryTagPayload::PowerPoint9 { atom, .. }
        | ShapeBinaryTagPayload::PowerPoint10 { atom, .. }
        | ShapeBinaryTagPayload::PowerPoint11 { atom, .. } => {
            encode_record(atom.version, atom.instance, atom.record_type, &atom.data)?
        },
        ShapeBinaryTagPayload::Unknown(data) => data.clone(),
    };
    payload.extend_from_slice(&encode_record(
        0,
        0,
        RecordType::BinaryTagData.as_u16(),
        &blob_data,
    )?);
    encode_record(0x0f, 0, RecordType::ProgBinaryTag.as_u16(), &payload)
}

fn encode_cstring_atom(instance: u16, units: &[u16]) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(units.len().saturating_mul(2));
    for unit in units {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    encode_record(0, instance, RecordType::CString.as_u16(), &data)
}

fn parse_cstring_atom(
    record: &RawRecord,
    instance: u16,
    printable: bool,
    limits: ShapeProgrammableTagLimits,
) -> Result<Vec<u16>> {
    require_atom_header(
        record,
        RecordType::CString,
        instance,
        "programmable tag CString",
    )?;
    if record.data.is_empty() || !record.data.len().is_multiple_of(2) {
        return corrupted("Programmable tag CString length must be positive and even");
    }
    let count = record.data.len() / 2;
    check_limit(
        count,
        limits.max_string_code_units,
        "programmable tag CString code-unit count",
    )?;
    let units: Vec<u16> = record
        .data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect();
    if printable {
        let end = units
            .iter()
            .position(|&unit| unit == 0)
            .unwrap_or(units.len());
        if units[..end]
            .iter()
            .any(|unit| matches!(*unit, 0x0000..=0x001f | 0x007f..=0x009f))
        {
            return corrupted("PrintableUnicodeString contains a control character");
        }
    }
    decode_units(&units, "programmable tag CString")?;
    Ok(units)
}

fn decode_units(units: &[u16], field: &str) -> Result<String> {
    let end = units
        .iter()
        .position(|&unit| unit == 0)
        .unwrap_or(units.len());
    String::from_utf16(&units[..end])
        .map_err(|_err| Error::Corrupted(format!("{field} contains invalid UTF-16")))
}

fn require_container_header(record: &RawRecord, kind: RecordType, name: &str) -> Result<()> {
    if record.version != 0x0f || record.instance != 0 || record.record_type != kind.as_u16() {
        return corrupted(format!("Invalid {name} record header"));
    }
    Ok(())
}

fn require_atom_header(
    record: &RawRecord,
    kind: RecordType,
    instance: u16,
    name: &str,
) -> Result<()> {
    if record.version != 0 || record.instance != instance || record.record_type != kind.as_u16() {
        return corrupted(format!("Invalid {name} record header"));
    }
    Ok(())
}

fn parse_sequence(data: &[u8], max_records: usize, context: &str) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        if records.len() >= max_records {
            return corrupted(format!("{context} exceeds its record-count limit"));
        }
        let (record, consumed) = parse_one(data, offset, context)?;
        records.push(record);
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| Error::Corrupted(format!("{context} offset overflow")))?;
    }
    Ok(records)
}

fn parse_one(data: &[u8], offset: usize, context: &str) -> Result<(RawRecord, usize)> {
    let header_end = offset
        .checked_add(8)
        .ok_or_else(|| Error::Corrupted(format!("{context} header offset overflow")))?;
    if header_end > data.len() {
        return corrupted(format!("Truncated record header in {context}"));
    }
    let version_instance = u16::from_le_bytes([data[offset], data[offset + 1]]);
    let record_type = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
    let length_u32 = u32::from_le_bytes([
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]);
    let length = usize::try_from(length_u32)
        .map_err(|_err| Error::Corrupted(format!("{context} record length overflow")))?;
    let end = header_end
        .checked_add(length)
        .ok_or_else(|| Error::Corrupted(format!("{context} record end overflow")))?;
    if end > data.len() {
        return corrupted(format!("Record extends beyond {context}"));
    }
    Ok((
        RawRecord {
            version: version_instance & 0x000f,
            instance: version_instance >> 4,
            record_type,
            data: data[header_end..end].to_vec(),
        },
        end - offset,
    ))
}

pub(super) fn encode_record(
    version: u16,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    if version > 0x0f || instance > 0x0fff {
        return corrupted("PPT record version or instance exceeds its bit field");
    }
    let length = u32::try_from(data.len())
        .map_err(|_err| Error::Corrupted("PPT record payload exceeds u32".into()))?;
    let mut result = Vec::with_capacity(8usize.saturating_add(data.len()));
    result.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    result.extend_from_slice(&record_type.to_le_bytes());
    result.extend_from_slice(&length.to_le_bytes());
    result.extend_from_slice(data);
    Ok(result)
}

fn check_limit(actual: usize, limit: usize, field: &str) -> Result<()> {
    if actual > limit {
        corrupted(format!("{field} exceeds its configured limit"))
    } else {
        Ok(())
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
