//! Shape-scoped programmable tags from OfficeArt `ClientData` records.
//!
//! This module implements MS-PPT sections 2.7.14 through 2.7.20. Parsing is
//! inert: binary tag payloads are never executed, loaded, or resolved.

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::records::PptRecord;
use super::text_extensions::{TextStyleExtension9, TextStyleExtension10, TextStyleExtension11};

/// Resource limits for shape programmable-tag parsing and serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointShapeProgrammableTagLimits {
    /// Maximum `ShapeProgTagsContainer` payload size.
    pub max_container_bytes: usize,
    /// Maximum number of direct string or binary tags.
    pub max_tags: usize,
    /// Maximum payload size of one `ProgStringTag` or `ProgBinaryTag`.
    pub max_tag_bytes: usize,
    /// Maximum number of UTF-16 code units in one tag name or value.
    pub max_string_code_units: usize,
    /// Maximum known PP9/PP10/PP11 style-atom payload size.
    pub max_style_payload_bytes: usize,
    /// Maximum decoded style runs in one known extension.
    pub max_style_runs: usize,
    /// Maximum opaque payload size for an unknown binary tag.
    pub max_unknown_binary_bytes: usize,
}

impl Default for PowerPointShapeProgrammableTagLimits {
    fn default() -> Self {
        Self {
            max_container_bytes: 4 * 1024 * 1024,
            max_tags: 1024,
            max_tag_bytes: 1024 * 1024,
            max_string_code_units: 64 * 1024,
            max_style_payload_bytes: 1024 * 1024,
            max_style_runs: 64 * 1024,
            max_unknown_binary_bytes: 1024 * 1024,
        }
    }
}

/// Raw atom retained alongside a typed versioned style payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeStyleAtom {
    /// Original record version.
    pub version: u16,
    /// Original record instance.
    pub instance: u16,
    /// Original numeric record type.
    pub record_type: u16,
    /// Original atom payload.
    pub data: Vec<u8>,
}

/// Discriminant of a shape binary programmable tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointShapeBinaryTagVersion {
    /// `___PPT9` / `PP9ShapeBinaryTagExtension`.
    PowerPoint9,
    /// `___PPT10` / `PP10ShapeBinaryTagExtension`.
    PowerPoint10,
    /// `___PPT11` / `PP11ShapeBinaryTagExtension`.
    PowerPoint11,
    /// Any tag name not assigned by sections 2.7.18 through 2.7.20.
    Unknown,
}

/// Typed or preserved payload of a shape binary programmable tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointShapeBinaryTagPayload {
    /// A `StyleTextProp9Atom` and its typed text-style data.
    PowerPoint9 {
        /// Decoded style data.
        style: TextStyleExtension9,
        /// Original style atom retained for byte-exact serialization.
        atom: PowerPointShapeStyleAtom,
    },
    /// A `StyleTextProp10Atom` and its typed text-style data.
    PowerPoint10 {
        /// Decoded style data.
        style: TextStyleExtension10,
        /// Original style atom retained for byte-exact serialization.
        atom: PowerPointShapeStyleAtom,
    },
    /// A `StyleTextProp11Atom` and its typed text-style data.
    PowerPoint11 {
        /// Decoded style data.
        style: TextStyleExtension11,
        /// Original style atom retained for byte-exact serialization.
        atom: PowerPointShapeStyleAtom,
    },
    /// An unassigned binary tag value, preserved without interpretation.
    Unknown(Vec<u8>),
}

/// One `ShapeProgBinaryTagContainer` and its CString/data-blob pair.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeBinaryTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Typed tag-name discriminant.
    pub version: PowerPointShapeBinaryTagVersion,
    /// Typed or preserved value.
    pub payload: PowerPointShapeBinaryTagPayload,
    name_units: Vec<u16>,
}

/// One `ProgStringTagContainer` allowed inside shape programmable tags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeStringTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Optional decoded Unicode value.
    pub value: Option<String>,
    name_units: Vec<u16>,
    value_units: Option<Vec<u16>>,
}

/// Direct child of a `ShapeProgTagsContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointShapeProgrammableTag {
    /// Unicode name/value programmable tag.
    String(PowerPointShapeStringTag),
    /// Binary programmable tag.
    Binary(PowerPointShapeBinaryTag),
}

/// Typed shape programmable tags retained from one OfficeArt `ClientData`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeProgrammableTags {
    /// Original `ShapeProgTagsContainer` record instance. Section 2.7.14 says
    /// this SHOULD be zero, so a nonzero value is preserved rather than rejected.
    pub instance: u16,
    /// Direct tags in file order.
    pub tags: Vec<PowerPointShapeProgrammableTag>,
}

/// Shape-level result returned by [`crate::ppt::slide::Slide`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeProgrammableTagsEntry {
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// Programmable tags owned by that shape.
    pub programmable_tags: PowerPointShapeProgrammableTags,
}

/// Presentation-level shape programmable-tag result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointPresentationShapeProgrammableTagsEntry {
    /// One-based slide number.
    pub slide_number: usize,
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// Programmable tags owned by that shape.
    pub programmable_tags: PowerPointShapeProgrammableTags,
}

impl PowerPointShapeProgrammableTags {
    /// Parse and validate a complete `ShapeProgTagsContainer` record.
    pub fn parse(record: &PptRecord, limits: PowerPointShapeProgrammableTagLimits) -> Result<Self> {
        if record.record_type != PptRecordType::ProgTags
            || record.record_type_raw != PptRecordType::ProgTags.as_u16()
            || record.version != 0x0f
        {
            return corrupted("Invalid ShapeProgTagsContainer record header");
        }
        let declared = usize::try_from(record.data_length)
            .map_err(|_| PptError::Corrupted("ShapeProgTagsContainer size overflow".into()))?;
        if declared != record.data.len() {
            return corrupted("ShapeProgTagsContainer length does not match its payload");
        }
        Self::parse_payload(&record.data, record.instance, limits)
    }

    /// Parse a `ShapeProgTagsContainer` payload and its record instance.
    pub fn parse_payload(
        data: &[u8],
        instance: u16,
        limits: PowerPointShapeProgrammableTagLimits,
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
                value if value == PptRecordType::ProgStringTag.as_u16() => {
                    PowerPointShapeProgrammableTag::String(parse_string_tag(record, limits)?)
                },
                value if value == PptRecordType::ProgBinaryTag.as_u16() => {
                    let tag = parse_binary_tag(record, limits)?;
                    let duplicate = match tag.version {
                        PowerPointShapeBinaryTagVersion::PowerPoint9 => {
                            std::mem::replace(&mut seen9, true)
                        },
                        PowerPointShapeBinaryTagVersion::PowerPoint10 => {
                            std::mem::replace(&mut seen10, true)
                        },
                        PowerPointShapeBinaryTagVersion::PowerPoint11 => {
                            std::mem::replace(&mut seen11, true)
                        },
                        PowerPointShapeBinaryTagVersion::Unknown => false,
                    };
                    if duplicate {
                        return corrupted(
                            "ShapeProgTagsContainer contains a duplicate versioned binary tag",
                        );
                    }
                    PowerPointShapeProgrammableTag::Binary(tag)
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
    pub fn to_bytes(&self, limits: PowerPointShapeProgrammableTagLimits) -> Result<Vec<u8>> {
        let payload = self.to_payload(limits)?;
        encode_record(
            0x0f,
            self.instance,
            PptRecordType::ProgTags.as_u16(),
            &payload,
        )
    }

    /// Serialize only the `ShapeProgTagsContainer` payload.
    pub fn to_payload(&self, limits: PowerPointShapeProgrammableTagLimits) -> Result<Vec<u8>> {
        check_limit(
            self.tags.len(),
            limits.max_tags,
            "shape programmable tag count",
        )?;
        let mut payload = Vec::new();
        for tag in &self.tags {
            let encoded = match tag {
                PowerPointShapeProgrammableTag::String(tag) => encode_string_tag(tag)?,
                PowerPointShapeProgrammableTag::Binary(tag) => encode_binary_tag(tag)?,
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

    /// Build a generic PPT record for insertion into OfficeArt client data.
    pub fn to_record(&self, limits: PowerPointShapeProgrammableTagLimits) -> Result<PptRecord> {
        let data = self.to_payload(limits)?;
        let data_length = u32::try_from(data.len()).map_err(|_| {
            PptError::Corrupted("ShapeProgTagsContainer payload exceeds u32".into())
        })?;
        Ok(PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: PptRecordType::ProgTags.as_u16(),
            version: 0x0f,
            instance: self.instance,
            data_length,
            data,
            children: Vec::new(),
        })
    }

    /// Return the decoded PowerPoint 9 style payload, when present.
    pub fn powerpoint9(&self) -> Option<&TextStyleExtension9> {
        self.tags.iter().find_map(|tag| match tag {
            PowerPointShapeProgrammableTag::Binary(PowerPointShapeBinaryTag {
                payload: PowerPointShapeBinaryTagPayload::PowerPoint9 { style, .. },
                ..
            }) => Some(style),
            _ => None,
        })
    }

    /// Return the decoded PowerPoint 10 style payload, when present.
    pub fn powerpoint10(&self) -> Option<&TextStyleExtension10> {
        self.tags.iter().find_map(|tag| match tag {
            PowerPointShapeProgrammableTag::Binary(PowerPointShapeBinaryTag {
                payload: PowerPointShapeBinaryTagPayload::PowerPoint10 { style, .. },
                ..
            }) => Some(style),
            _ => None,
        })
    }

    /// Return the decoded PowerPoint 11 style payload, when present.
    pub fn powerpoint11(&self) -> Option<&TextStyleExtension11> {
        self.tags.iter().find_map(|tag| match tag {
            PowerPointShapeProgrammableTag::Binary(PowerPointShapeBinaryTag {
                payload: PowerPointShapeBinaryTagPayload::PowerPoint11 { style, .. },
                ..
            }) => Some(style),
            _ => None,
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

fn parse_string_tag(
    record: RawRecord,
    limits: PowerPointShapeProgrammableTagLimits,
) -> Result<PowerPointShapeStringTag> {
    require_container_header(
        &record,
        PptRecordType::ProgStringTag,
        "ProgStringTagContainer",
    )?;
    let children = parse_sequence(&record.data, 2, "ProgStringTagContainer")?;
    if children.is_empty() || children.len() > 2 {
        return corrupted("ProgStringTagContainer must contain a name and at most one value");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let value_units = children
        .get(1)
        .map(|value| parse_cstring_atom(value, 1, false, limits))
        .transpose()?;
    Ok(PowerPointShapeStringTag {
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
    record: RawRecord,
    limits: PowerPointShapeProgrammableTagLimits,
) -> Result<PowerPointShapeBinaryTag> {
    require_container_header(
        &record,
        PptRecordType::ProgBinaryTag,
        "ShapeProgBinaryTagContainer",
    )?;
    let children = parse_sequence(&record.data, 2, "ShapeProgBinaryTagContainer")?;
    if children.len() != 2 {
        return corrupted("ShapeProgBinaryTagContainer must contain exactly one CString/blob pair");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let name = decode_units(&name_units, "binary programmable tag name")?;
    let blob = &children[1];
    require_atom_header(blob, PptRecordType::BinaryTagData, 0, "BinaryTagDataBlob")?;

    let (version, payload) = match name.as_str() {
        "___PPT9" => (
            PowerPointShapeBinaryTagVersion::PowerPoint9,
            parse_known_payload(
                blob,
                PptRecordType::StyleTextProp9Atom,
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
                    Ok(PowerPointShapeBinaryTagPayload::PowerPoint9 {
                        style,
                        atom: atom.into(),
                    })
                },
            )?,
        ),
        "___PPT10" => (
            PowerPointShapeBinaryTagVersion::PowerPoint10,
            parse_known_payload(
                blob,
                PptRecordType::StyleTextProp10Atom,
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
                    Ok(PowerPointShapeBinaryTagPayload::PowerPoint10 {
                        style,
                        atom: atom.into(),
                    })
                },
            )?,
        ),
        "___PPT11" => (
            PowerPointShapeBinaryTagVersion::PowerPoint11,
            parse_known_payload(
                blob,
                PptRecordType::StyleTextProp11Atom,
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
                    Ok(PowerPointShapeBinaryTagPayload::PowerPoint11 {
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
                PowerPointShapeBinaryTagVersion::Unknown,
                PowerPointShapeBinaryTagPayload::Unknown(blob.data.clone()),
            )
        },
    };

    Ok(PowerPointShapeBinaryTag {
        name,
        version,
        payload,
        name_units,
    })
}

fn parse_known_payload<F>(
    blob: &RawRecord,
    expected_type: PptRecordType,
    expected_name_bytes: usize,
    name_units: &[u16],
    limits: PowerPointShapeProgrammableTagLimits,
    decode: F,
) -> Result<PowerPointShapeBinaryTagPayload>
where
    F: FnOnce(RawRecord) -> Result<PowerPointShapeBinaryTagPayload>,
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

impl From<RawRecord> for PowerPointShapeStyleAtom {
    fn from(value: RawRecord) -> Self {
        Self {
            version: value.version,
            instance: value.instance,
            record_type: value.record_type,
            data: value.data,
        }
    }
}

fn encode_string_tag(tag: &PowerPointShapeStringTag) -> Result<Vec<u8>> {
    let name = encode_cstring_atom(0, &tag.name_units)?;
    let mut payload = name;
    if let Some(value) = &tag.value_units {
        payload.extend_from_slice(&encode_cstring_atom(1, value)?);
    }
    encode_record(0x0f, 0, PptRecordType::ProgStringTag.as_u16(), &payload)
}

fn encode_binary_tag(tag: &PowerPointShapeBinaryTag) -> Result<Vec<u8>> {
    let mut payload = encode_cstring_atom(0, &tag.name_units)?;
    let blob_data = match &tag.payload {
        PowerPointShapeBinaryTagPayload::PowerPoint9 { atom, .. }
        | PowerPointShapeBinaryTagPayload::PowerPoint10 { atom, .. }
        | PowerPointShapeBinaryTagPayload::PowerPoint11 { atom, .. } => {
            encode_record(atom.version, atom.instance, atom.record_type, &atom.data)?
        },
        PowerPointShapeBinaryTagPayload::Unknown(data) => data.clone(),
    };
    payload.extend_from_slice(&encode_record(
        0,
        0,
        PptRecordType::BinaryTagData.as_u16(),
        &blob_data,
    )?);
    encode_record(0x0f, 0, PptRecordType::ProgBinaryTag.as_u16(), &payload)
}

fn encode_cstring_atom(instance: u16, units: &[u16]) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(units.len().saturating_mul(2));
    for unit in units {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    encode_record(0, instance, PptRecordType::CString.as_u16(), &data)
}

fn parse_cstring_atom(
    record: &RawRecord,
    instance: u16,
    printable: bool,
    limits: PowerPointShapeProgrammableTagLimits,
) -> Result<Vec<u16>> {
    require_atom_header(
        record,
        PptRecordType::CString,
        instance,
        "programmable tag CString",
    )?;
    if record.data.is_empty() || record.data.len() % 2 != 0 {
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
        .map_err(|_| PptError::Corrupted(format!("{field} contains invalid UTF-16")))
}

fn require_container_header(record: &RawRecord, kind: PptRecordType, name: &str) -> Result<()> {
    if record.version != 0x0f || record.instance != 0 || record.record_type != kind.as_u16() {
        return corrupted(format!("Invalid {name} record header"));
    }
    Ok(())
}

fn require_atom_header(
    record: &RawRecord,
    kind: PptRecordType,
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
            .ok_or_else(|| PptError::Corrupted(format!("{context} offset overflow")))?;
    }
    Ok(records)
}

fn parse_one(data: &[u8], offset: usize, context: &str) -> Result<(RawRecord, usize)> {
    let header_end = offset
        .checked_add(8)
        .ok_or_else(|| PptError::Corrupted(format!("{context} header offset overflow")))?;
    if header_end > data.len() {
        return corrupted(format!("Truncated record header in {context}"));
    }
    let version_instance = u16::from_le_bytes([data[offset], data[offset + 1]]);
    let record_type = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
    let length = u32::from_le_bytes([
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]);
    let length = usize::try_from(length)
        .map_err(|_| PptError::Corrupted(format!("{context} record length overflow")))?;
    let end = header_end
        .checked_add(length)
        .ok_or_else(|| PptError::Corrupted(format!("{context} record end overflow")))?;
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

fn encode_record(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    if version > 0x0f || instance > 0x0fff {
        return corrupted("PPT record version or instance exceeds its bit field");
    }
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PPT record payload exceeds u32".into()))?;
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
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        encode_record(version, instance, kind, payload).unwrap()
    }

    fn string_tag(name: &str, value: Option<&str>) -> Vec<u8> {
        let units = |text: &str| {
            text.encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>()
        };
        let mut data = record(0, 0, PptRecordType::CString.as_u16(), &units(name));
        if let Some(value) = value {
            data.extend_from_slice(&record(
                0,
                1,
                PptRecordType::CString.as_u16(),
                &units(value),
            ));
        }
        record(0x0f, 0, PptRecordType::ProgStringTag.as_u16(), &data)
    }

    fn binary_tag(name: &str, style_type: Option<PptRecordType>, payload: &[u8]) -> Vec<u8> {
        let name_data: Vec<u8> = name.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let mut data = record(0, 0, PptRecordType::CString.as_u16(), &name_data);
        let blob_data = style_type
            .map(|kind| record(0, 0, kind.as_u16(), payload))
            .unwrap_or_else(|| payload.to_vec());
        data.extend_from_slice(&record(
            0,
            0,
            PptRecordType::BinaryTagData.as_u16(),
            &blob_data,
        ));
        record(0x0f, 0, PptRecordType::ProgBinaryTag.as_u16(), &data)
    }

    fn complete_record(payload: &[u8]) -> (Vec<u8>, PptRecord) {
        let bytes = record(0x0f, 7, PptRecordType::ProgTags.as_u16(), payload);
        let (parsed, consumed) = PptRecord::parse_strict(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        (bytes, parsed)
    }

    #[test]
    fn parses_all_defined_variants_and_round_trips_exactly() {
        let mut payload = string_tag("author", Some("Ada"));
        payload.extend_from_slice(&binary_tag(
            "___PPT9",
            Some(PptRecordType::StyleTextProp9Atom),
            &[0; 12],
        ));
        payload.extend_from_slice(&binary_tag(
            "___PPT10",
            Some(PptRecordType::StyleTextProp10Atom),
            &[0; 4],
        ));
        payload.extend_from_slice(&binary_tag(
            "___PPT11",
            Some(PptRecordType::StyleTextProp11Atom),
            &[0; 4],
        ));
        payload.extend_from_slice(&binary_tag("vendor", None, &[1, 2, 3, 4, 5]));
        let (bytes, record) = complete_record(&payload);

        let parsed = PowerPointShapeProgrammableTags::parse(
            &record,
            PowerPointShapeProgrammableTagLimits::default(),
        )
        .unwrap();

        assert_eq!(parsed.instance, 7);
        assert_eq!(parsed.tags.len(), 5);
        assert_eq!(parsed.powerpoint9().unwrap().runs.len(), 1);
        assert_eq!(parsed.powerpoint10().unwrap().runs.len(), 1);
        assert_eq!(parsed.powerpoint11().unwrap().runs.len(), 1);
        assert_eq!(
            parsed.tags.iter().find_map(|tag| match tag {
                PowerPointShapeProgrammableTag::String(tag) => tag.value.as_deref(),
                _ => None,
            }),
            Some("Ada")
        );
        assert_eq!(
            parsed
                .to_bytes(PowerPointShapeProgrammableTagLimits::default())
                .unwrap(),
            bytes
        );
    }

    #[test]
    fn enforces_container_pair_and_version_ownership() {
        let limits = PowerPointShapeProgrammableTagLimits::default();
        let duplicate = [
            binary_tag("___PPT9", Some(PptRecordType::StyleTextProp9Atom), &[0; 12]),
            binary_tag("___PPT9", Some(PptRecordType::StyleTextProp9Atom), &[0; 12]),
        ]
        .concat();
        assert!(PowerPointShapeProgrammableTags::parse_payload(&duplicate, 0, limits).is_err());

        let wrong_style = binary_tag(
            "___PPT10",
            Some(PptRecordType::StyleTextProp9Atom),
            &[0; 12],
        );
        assert!(PowerPointShapeProgrammableTags::parse_payload(&wrong_style, 0, limits).is_err());

        let mut missing_blob = record(
            0,
            0,
            PptRecordType::CString.as_u16(),
            &"vendor"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>(),
        );
        missing_blob = record(
            0x0f,
            0,
            PptRecordType::ProgBinaryTag.as_u16(),
            &missing_blob,
        );
        assert!(PowerPointShapeProgrammableTags::parse_payload(&missing_blob, 0, limits).is_err());

        let disallowed = record(0, 0, PptRecordType::CString.as_u16(), &[65, 0]);
        assert!(PowerPointShapeProgrammableTags::parse_payload(&disallowed, 0, limits).is_err());
    }

    #[test]
    fn rejects_malformed_strings_headers_truncation_and_every_limit() {
        let defaults = PowerPointShapeProgrammableTagLimits::default();
        let valid = binary_tag("vendor", None, &[1, 2, 3, 4]);

        let mut truncated = valid.clone();
        truncated.pop();
        assert!(PowerPointShapeProgrammableTags::parse_payload(&truncated, 0, defaults).is_err());

        let invalid_utf16 = record(0, 0, PptRecordType::CString.as_u16(), &[0x00, 0xd8]);
        let invalid_utf16 = record(
            0x0f,
            0,
            PptRecordType::ProgStringTag.as_u16(),
            &invalid_utf16,
        );
        assert!(
            PowerPointShapeProgrammableTags::parse_payload(&invalid_utf16, 0, defaults).is_err()
        );

        let control_name = string_tag("bad\nname", None);
        assert!(
            PowerPointShapeProgrammableTags::parse_payload(&control_name, 0, defaults).is_err()
        );

        let cases = [
            PowerPointShapeProgrammableTagLimits {
                max_container_bytes: valid.len() - 1,
                ..defaults
            },
            PowerPointShapeProgrammableTagLimits {
                max_tags: 0,
                ..defaults
            },
            PowerPointShapeProgrammableTagLimits {
                max_tag_bytes: 0,
                ..defaults
            },
            PowerPointShapeProgrammableTagLimits {
                max_string_code_units: 1,
                ..defaults
            },
            PowerPointShapeProgrammableTagLimits {
                max_unknown_binary_bytes: 3,
                ..defaults
            },
        ];
        for limits in cases {
            assert!(PowerPointShapeProgrammableTags::parse_payload(&valid, 0, limits).is_err());
        }

        let known = binary_tag(
            "___PPT10",
            Some(PptRecordType::StyleTextProp10Atom),
            &[0; 4],
        );
        assert!(
            PowerPointShapeProgrammableTags::parse_payload(
                &known,
                0,
                PowerPointShapeProgrammableTagLimits {
                    max_style_payload_bytes: 3,
                    ..defaults
                },
            )
            .is_err()
        );
        assert!(
            PowerPointShapeProgrammableTags::parse_payload(
                &known,
                0,
                PowerPointShapeProgrammableTagLimits {
                    max_style_runs: 0,
                    ..defaults
                },
            )
            .is_err()
        );
    }
}
