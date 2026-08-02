//! Document- and slide-level programmable tags.
//!
//! This module implements the `DocProgTagsContainer` family (MS-PPT sections
//! 2.4.23.1 through 2.4.23.4) and the `SlideProgTagsContainer` family (MS-PPT
//! sections 2.5.19 through 2.5.22), together with the shared
//! `ProgStringTagContainer`/`TagNameAtom`/`TagValueAtom` records (sections
//! 2.11.30 through 2.11.32) and `UnknownBinaryTag` (section 2.11.33).
//!
//! Parsing is inert: versioned binary payloads are validated as strict record
//! sequences and retained byte-for-byte, but they are never executed, loaded,
//! or resolved. Shape-scoped programmable tags (sections 2.7.14 through
//! 2.7.20) live in [`super::shape_programmable_tags`]; typed decoding of the
//! versioned extension payloads lives in [`super::prog_tag_extensions`].

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::records::PptRecord;

/// Resource limits for document/slide programmable-tag parsing and serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointProgTagLimits {
    /// Maximum `DocProgTagsContainer`/`SlideProgTagsContainer` payload size.
    pub max_container_bytes: usize,
    /// Maximum number of direct string or binary tags.
    pub max_tags: usize,
    /// Maximum payload size of one `ProgStringTag` or `ProgBinaryTag` container.
    pub max_tag_bytes: usize,
    /// Maximum number of UTF-16 code units in one tag name or value.
    pub max_string_code_units: usize,
    /// Maximum payload size of one `BinaryTagDataBlob`.
    pub max_binary_payload_bytes: usize,
    /// Maximum number of records inside one versioned `BinaryTagDataBlob`.
    pub max_binary_records: usize,
}

impl Default for PowerPointProgTagLimits {
    fn default() -> Self {
        Self {
            max_container_bytes: 16 * 1024 * 1024,
            max_tags: 1024,
            max_tag_bytes: 8 * 1024 * 1024,
            max_string_code_units: 64 * 1024,
            max_binary_payload_bytes: 8 * 1024 * 1024,
            max_binary_records: 64 * 1024,
        }
    }
}

/// The record family a `ProgTags` container belongs to.
///
/// The record type is identical (`RT_ProgTags`) in both scopes, but the set of
/// assigned versioned binary-tag names differs: document tags assign
/// `___PPT9` through `___PPT12` (section 2.4.23.4) while slide tags assign
/// only `___PPT9`, `___PPT10`, and `___PPT12` (section 2.5.22). Any other
/// name, including `___PPT11` at slide scope, is an `UnknownBinaryTag`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointProgTagScope {
    /// `DocProgTagsContainer` inside the `DocumentContainer` (section 2.4.23.1).
    Document,
    /// `SlideProgTagsContainer` inside a slide, notes, handout, or main-master
    /// container (section 2.5.19).
    Slide,
}

impl PowerPointProgTagScope {
    /// Return the assigned version for a tag name in this scope, if any.
    fn known_version(self, name: &str) -> Option<PowerPointProgBinaryTagVersion> {
        match (self, name) {
            (_, "___PPT9") => Some(PowerPointProgBinaryTagVersion::PowerPoint9),
            (_, "___PPT10") => Some(PowerPointProgBinaryTagVersion::PowerPoint10),
            (Self::Document, "___PPT11") => Some(PowerPointProgBinaryTagVersion::PowerPoint11),
            (_, "___PPT12") => Some(PowerPointProgBinaryTagVersion::PowerPoint12),
            _ => None,
        }
    }
}

/// Discriminant of a document/slide binary programmable tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointProgBinaryTagVersion {
    /// `___PPT9` / `PP9DocBinaryTagExtension` or `PP9SlideBinaryTagExtension`.
    PowerPoint9,
    /// `___PPT10` / `PP10DocBinaryTagExtension` or `PP10SlideBinaryTagExtension`.
    PowerPoint10,
    /// `___PPT11` / `PP11DocBinaryTagExtension` (document scope only).
    PowerPoint11,
    /// `___PPT12` / `PP12DocBinaryTagExtension` or `PP12SlideBinaryTagExtension`.
    PowerPoint12,
    /// Any tag name not assigned by section 2.4.23.4 or 2.5.22 for the scope.
    Unknown,
}

/// One `ProgStringTagContainer` (section 2.11.30) and its name/value pair.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointProgStringTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Optional decoded Unicode value.
    pub value: Option<String>,
    name_units: Vec<u16>,
    value_units: Option<Vec<u16>>,
}

/// One `DocProgBinaryTagContainer`/`SlideProgBinaryTagContainer` record pair.
///
/// The `BinaryTagDataBlob` payload is retained byte-for-byte in `payload`.
/// For versioned tags the payload is validated as a strict record sequence at
/// parse time; use [`PowerPointProgBinaryTag::records`] to decode it into
/// typed records. Unknown tags are preserved without any interpretation, as
/// required by sections 2.4.23.4 and 2.5.22.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointProgBinaryTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Typed tag-name discriminant for the container scope.
    pub version: PowerPointProgBinaryTagVersion,
    /// Raw `BinaryTagDataBlob` payload, preserved for byte-exact serialization.
    pub payload: Vec<u8>,
    name_units: Vec<u16>,
}

impl PowerPointProgBinaryTag {
    /// Decode the retained `BinaryTagDataBlob` payload as a strict sequence of
    /// PPT records.
    ///
    /// Versioned tag payloads are guaranteed to decode successfully because
    /// they were validated at parse time; unknown payloads may hold arbitrary
    /// bytes and can fail here.
    pub fn records(&self) -> Result<Vec<PptRecord>> {
        PptRecord::parse_sequence_strict(&self.payload, "BinaryTagDataBlob payload")
    }
}

/// Direct child of a `DocProgTagsContainer`/`SlideProgTagsContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointProgTag {
    /// Unicode name/value programmable tag.
    String(PowerPointProgStringTag),
    /// Binary programmable tag.
    Binary(PowerPointProgBinaryTag),
}

/// Typed programmable tags of one document- or slide-level `ProgTags` container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointProgTags {
    /// The record family this container belongs to.
    pub scope: PowerPointProgTagScope,
    /// Original container record instance. Sections 2.4.23.1 and 2.5.19 say
    /// this SHOULD be zero, so a nonzero value is preserved rather than rejected.
    pub instance: u16,
    /// Direct tags in file order.
    pub tags: Vec<PowerPointProgTag>,
}

impl PowerPointProgTags {
    /// Parse the `DocProgTagsContainer` of a `DocumentContainer` record.
    ///
    /// The container is an optional child of the document's
    /// `DocInfoListContainer` (MS-PPT section 2.4.4); each record type MUST
    /// NOT occur there more than once, so duplicates are an error.
    pub fn parse_document(
        document: &PptRecord,
        limits: PowerPointProgTagLimits,
    ) -> Result<Option<Self>> {
        if document.record_type != PptRecordType::Document {
            return corrupted("Document ProgTags require a DocumentContainer record");
        }
        let Some(doc_info_list) = single_child(document, PptRecordType::DocInfoList)? else {
            return Ok(None);
        };
        let Some(record) = single_child(doc_info_list, PptRecordType::ProgTags)? else {
            return Ok(None);
        };
        Self::parse(record, PowerPointProgTagScope::Document, limits).map(Some)
    }

    /// Parse the `SlideProgTagsContainer` of a slide-like container record
    /// (`SlideContainer`, `NotesContainer`, `MainMasterContainer`, or
    /// `HandoutContainer`; MS-PPT section 2.5.19).
    pub fn parse_slide(
        container: &PptRecord,
        limits: PowerPointProgTagLimits,
    ) -> Result<Option<Self>> {
        let Some(record) = single_child(container, PptRecordType::ProgTags)? else {
            return Ok(None);
        };
        Self::parse(record, PowerPointProgTagScope::Slide, limits).map(Some)
    }

    /// Parse and validate a complete `DocProgTagsContainer` or
    /// `SlideProgTagsContainer` record.
    pub fn parse(
        record: &PptRecord,
        scope: PowerPointProgTagScope,
        limits: PowerPointProgTagLimits,
    ) -> Result<Self> {
        if record.record_type != PptRecordType::ProgTags
            || record.record_type_raw != PptRecordType::ProgTags.as_u16()
            || record.version != 0x0f
        {
            return corrupted("Invalid ProgTags container record header");
        }
        let declared = usize::try_from(record.data_length)
            .map_err(|_| PptError::Corrupted("ProgTags container size overflow".into()))?;
        if declared != record.data.len() {
            return corrupted("ProgTags container length does not match its payload");
        }
        Self::parse_payload(&record.data, record.instance, scope, limits)
    }

    /// Parse a `ProgTags` container payload and its record instance.
    pub fn parse_payload(
        data: &[u8],
        instance: u16,
        scope: PowerPointProgTagScope,
        limits: PowerPointProgTagLimits,
    ) -> Result<Self> {
        check_limit(data.len(), limits.max_container_bytes, "ProgTags payload")?;
        let records = parse_sequence(data, limits.max_tags, "ProgTags")?;
        let mut tags = Vec::with_capacity(records.len());
        let mut seen9 = false;
        let mut seen10 = false;
        let mut seen11 = false;
        let mut seen12 = false;

        for record in records {
            check_limit(
                record.data.len(),
                limits.max_tag_bytes,
                "programmable tag payload",
            )?;
            let tag = match record.record_type {
                value if value == PptRecordType::ProgStringTag.as_u16() => {
                    PowerPointProgTag::String(parse_string_tag(record, limits)?)
                },
                value if value == PptRecordType::ProgBinaryTag.as_u16() => {
                    let tag = parse_binary_tag(record, scope, limits)?;
                    // Sections 2.4.23.1 and 2.5.19: the array MUST NOT contain
                    // more than one of each versioned extension.
                    let duplicate = match tag.version {
                        PowerPointProgBinaryTagVersion::PowerPoint9 => {
                            std::mem::replace(&mut seen9, true)
                        },
                        PowerPointProgBinaryTagVersion::PowerPoint10 => {
                            std::mem::replace(&mut seen10, true)
                        },
                        PowerPointProgBinaryTagVersion::PowerPoint11 => {
                            std::mem::replace(&mut seen11, true)
                        },
                        PowerPointProgBinaryTagVersion::PowerPoint12 => {
                            std::mem::replace(&mut seen12, true)
                        },
                        PowerPointProgBinaryTagVersion::Unknown => false,
                    };
                    if duplicate {
                        return corrupted(
                            "ProgTags container contains a duplicate versioned binary tag",
                        );
                    }
                    PowerPointProgTag::Binary(tag)
                },
                _ => {
                    return corrupted("ProgTags container contains a disallowed child record type");
                },
            };
            tags.push(tag);
        }
        Ok(Self {
            scope,
            instance,
            tags,
        })
    }

    /// Serialize a complete `ProgTags` container record.
    pub fn to_bytes(&self, limits: PowerPointProgTagLimits) -> Result<Vec<u8>> {
        let payload = self.to_payload(limits)?;
        encode_record(
            0x0f,
            self.instance,
            PptRecordType::ProgTags.as_u16(),
            &payload,
        )
    }

    /// Serialize only the `ProgTags` container payload.
    pub fn to_payload(&self, limits: PowerPointProgTagLimits) -> Result<Vec<u8>> {
        check_limit(self.tags.len(), limits.max_tags, "programmable tag count")?;
        let mut payload = Vec::new();
        for tag in &self.tags {
            let encoded = match tag {
                PowerPointProgTag::String(tag) => encode_string_tag(tag)?,
                PowerPointProgTag::Binary(tag) => encode_binary_tag(tag)?,
            };
            check_limit(
                encoded.len().saturating_sub(8),
                limits.max_tag_bytes,
                "programmable tag payload",
            )?;
            payload.extend_from_slice(&encoded);
        }
        check_limit(
            payload.len(),
            limits.max_container_bytes,
            "ProgTags payload",
        )?;
        // Reparse before returning so public-field mutations cannot serialize an
        // invalid duplicate/version combination or bypass any current limit.
        Self::parse_payload(&payload, self.instance, self.scope, limits)?;
        Ok(payload)
    }

    /// Build a generic PPT record for insertion into a document or slide tree.
    pub fn to_record(&self, limits: PowerPointProgTagLimits) -> Result<PptRecord> {
        let data = self.to_payload(limits)?;
        let data_length = u32::try_from(data.len())
            .map_err(|_| PptError::Corrupted("ProgTags payload exceeds u32".into()))?;
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

    /// Return the binary tag for an assigned version, when present.
    pub fn binary_tag(
        &self,
        version: PowerPointProgBinaryTagVersion,
    ) -> Option<&PowerPointProgBinaryTag> {
        self.tags.iter().find_map(|tag| match tag {
            PowerPointProgTag::Binary(tag) if tag.version == version => Some(tag),
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
    limits: PowerPointProgTagLimits,
) -> Result<PowerPointProgStringTag> {
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
    Ok(PowerPointProgStringTag {
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
    scope: PowerPointProgTagScope,
    limits: PowerPointProgTagLimits,
) -> Result<PowerPointProgBinaryTag> {
    require_container_header(
        &record,
        PptRecordType::ProgBinaryTag,
        "ProgBinaryTagContainer",
    )?;
    let children = parse_sequence(&record.data, 2, "ProgBinaryTagContainer")?;
    if children.len() != 2 {
        return corrupted("ProgBinaryTagContainer must contain exactly one CString/blob pair");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let name = decode_units(&name_units, "binary programmable tag name")?;
    let blob = &children[1];
    require_atom_header(blob, PptRecordType::BinaryTagData, 0, "BinaryTagDataBlob")?;
    check_limit(
        blob.data.len(),
        limits.max_binary_payload_bytes,
        "binary programmable tag payload",
    )?;

    let version = scope
        .known_version(name.as_str())
        .unwrap_or(PowerPointProgBinaryTagVersion::Unknown);
    if version != PowerPointProgBinaryTagVersion::Unknown {
        // Sections 2.4.23.5 and 2.5.23: the versioned tag name has a fixed
        // length with no NUL padding, and the blob is a strict record sequence.
        if name_units.len() != name.len() {
            return corrupted("Versioned binary tag has an invalid CString length");
        }
        let records = PptRecord::parse_sequence_strict(&blob.data, "versioned BinaryTagDataBlob")?;
        check_limit(
            records.len(),
            limits.max_binary_records,
            "versioned binary tag record count",
        )?;
    }

    Ok(PowerPointProgBinaryTag {
        name,
        version,
        payload: blob.data.clone(),
        name_units,
    })
}

fn encode_string_tag(tag: &PowerPointProgStringTag) -> Result<Vec<u8>> {
    let name = encode_cstring_atom(0, &tag.name_units)?;
    let mut payload = name;
    if let Some(value) = &tag.value_units {
        payload.extend_from_slice(&encode_cstring_atom(1, value)?);
    }
    encode_record(0x0f, 0, PptRecordType::ProgStringTag.as_u16(), &payload)
}

fn encode_binary_tag(tag: &PowerPointProgBinaryTag) -> Result<Vec<u8>> {
    let mut payload = encode_cstring_atom(0, &tag.name_units)?;
    payload.extend_from_slice(&encode_record(
        0,
        0,
        PptRecordType::BinaryTagData.as_u16(),
        &tag.payload,
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
    limits: PowerPointProgTagLimits,
) -> Result<Vec<u16>> {
    require_atom_header(
        record,
        PptRecordType::CString,
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

/// Return the single direct child of a record type, rejecting duplicates.
fn single_child(record: &PptRecord, kind: PptRecordType) -> Result<Option<&PptRecord>> {
    let mut matches = record
        .children
        .iter()
        .filter(|child| child.record_type == kind);
    let first = matches.next();
    if matches.next().is_some() {
        return corrupted(format!(
            "container contains more than one {kind:?} child record"
        ));
    }
    Ok(first)
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

    fn binary_tag(name: &str, payload: &[u8]) -> Vec<u8> {
        let name_data: Vec<u8> = name.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let mut data = record(0, 0, PptRecordType::CString.as_u16(), &name_data);
        data.extend_from_slice(&record(
            0,
            0,
            PptRecordType::BinaryTagData.as_u16(),
            payload,
        ));
        record(0x0f, 0, PptRecordType::ProgBinaryTag.as_u16(), &data)
    }

    fn complete_record(payload: &[u8]) -> (Vec<u8>, PptRecord) {
        let bytes = record(0x0f, 3, PptRecordType::ProgTags.as_u16(), payload);
        let (parsed, consumed) = PptRecord::parse_strict(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        (bytes, parsed)
    }

    fn versioned_payload() -> Vec<u8> {
        // Two arbitrary atom records forming a valid strict record sequence.
        let mut payload = record(0, 0, PptRecordType::TextHeaderAtom.as_u16(), &[0; 4]);
        payload.extend_from_slice(&record(
            0,
            0,
            PptRecordType::StyleTextProp9Atom.as_u16(),
            &[0; 12],
        ));
        payload
    }

    #[test]
    fn parses_document_scope_variants_and_round_trips_exactly() {
        let limits = PowerPointProgTagLimits::default();
        let mut payload = string_tag("author", Some("Ada"));
        payload.extend_from_slice(&binary_tag("___PPT9", &versioned_payload()));
        payload.extend_from_slice(&binary_tag("___PPT10", &versioned_payload()));
        payload.extend_from_slice(&binary_tag("___PPT11", &versioned_payload()));
        payload.extend_from_slice(&binary_tag("___PPT12", &versioned_payload()));
        payload.extend_from_slice(&binary_tag("vendor", &[1, 2, 3]));
        let (bytes, record) = complete_record(&payload);

        let parsed =
            PowerPointProgTags::parse(&record, PowerPointProgTagScope::Document, limits).unwrap();

        assert_eq!(parsed.scope, PowerPointProgTagScope::Document);
        assert_eq!(parsed.instance, 3);
        assert_eq!(parsed.tags.len(), 6);
        for version in [
            PowerPointProgBinaryTagVersion::PowerPoint9,
            PowerPointProgBinaryTagVersion::PowerPoint10,
            PowerPointProgBinaryTagVersion::PowerPoint11,
            PowerPointProgBinaryTagVersion::PowerPoint12,
        ] {
            let tag = parsed.binary_tag(version).unwrap();
            assert_eq!(tag.records().unwrap().len(), 2);
        }
        let unknown = parsed
            .binary_tag(PowerPointProgBinaryTagVersion::Unknown)
            .unwrap();
        assert_eq!(unknown.name, "vendor");
        assert_eq!(unknown.payload, [1, 2, 3]);
        assert!(unknown.records().is_err());
        assert_eq!(
            parsed.tags.iter().find_map(|tag| match tag {
                PowerPointProgTag::String(tag) => tag.value.as_deref(),
                _ => None,
            }),
            Some("Ada")
        );
        assert_eq!(parsed.to_bytes(limits).unwrap(), bytes);
    }

    #[test]
    fn slide_scope_treats_ppt11_as_unknown() {
        let limits = PowerPointProgTagLimits::default();
        let mut payload = binary_tag("___PPT9", &versioned_payload());
        payload.extend_from_slice(&binary_tag("___PPT11", &versioned_payload()));
        let (bytes, record) = complete_record(&payload);

        let parsed =
            PowerPointProgTags::parse(&record, PowerPointProgTagScope::Slide, limits).unwrap();

        assert_eq!(
            parsed
                .binary_tag(PowerPointProgBinaryTagVersion::PowerPoint9)
                .unwrap()
                .name,
            "___PPT9"
        );
        assert!(
            parsed
                .binary_tag(PowerPointProgBinaryTagVersion::PowerPoint11)
                .is_none()
        );
        let unknown = parsed
            .binary_tag(PowerPointProgBinaryTagVersion::Unknown)
            .unwrap();
        assert_eq!(unknown.name, "___PPT11");
        assert_eq!(parsed.to_bytes(limits).unwrap(), bytes);
    }

    #[test]
    fn enforces_duplicate_pair_and_versioned_payload_rules() {
        let limits = PowerPointProgTagLimits::default();
        let duplicate = [
            binary_tag("___PPT9", &versioned_payload()),
            binary_tag("___PPT9", &versioned_payload()),
        ]
        .concat();
        assert!(
            PowerPointProgTags::parse_payload(
                &duplicate,
                0,
                PowerPointProgTagScope::Document,
                limits
            )
            .is_err()
        );

        // A versioned tag whose blob is not a strict record sequence is invalid.
        let invalid_blob = binary_tag("___PPT10", &[1, 2, 3]);
        assert!(
            PowerPointProgTags::parse_payload(
                &invalid_blob,
                0,
                PowerPointProgTagScope::Document,
                limits
            )
            .is_err()
        );

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
        assert!(
            PowerPointProgTags::parse_payload(
                &missing_blob,
                0,
                PowerPointProgTagScope::Document,
                limits
            )
            .is_err()
        );

        let disallowed = record(0, 0, PptRecordType::CString.as_u16(), &[65, 0]);
        assert!(
            PowerPointProgTags::parse_payload(
                &disallowed,
                0,
                PowerPointProgTagScope::Document,
                limits
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_malformed_strings_headers_truncation_and_every_limit() {
        let defaults = PowerPointProgTagLimits::default();
        let valid = binary_tag("vendor", &[1, 2, 3, 4]);

        let mut truncated = valid.clone();
        truncated.pop();
        assert!(
            PowerPointProgTags::parse_payload(
                &truncated,
                0,
                PowerPointProgTagScope::Document,
                defaults
            )
            .is_err()
        );

        let invalid_utf16 = record(0, 0, PptRecordType::CString.as_u16(), &[0x00, 0xd8]);
        let invalid_utf16 = record(
            0x0f,
            0,
            PptRecordType::ProgStringTag.as_u16(),
            &invalid_utf16,
        );
        assert!(
            PowerPointProgTags::parse_payload(
                &invalid_utf16,
                0,
                PowerPointProgTagScope::Document,
                defaults
            )
            .is_err()
        );

        let control_name = string_tag("bad\nname", None);
        assert!(
            PowerPointProgTags::parse_payload(
                &control_name,
                0,
                PowerPointProgTagScope::Document,
                defaults
            )
            .is_err()
        );

        let cases = [
            PowerPointProgTagLimits {
                max_container_bytes: valid.len() - 1,
                ..defaults
            },
            PowerPointProgTagLimits {
                max_tags: 0,
                ..defaults
            },
            PowerPointProgTagLimits {
                max_tag_bytes: 0,
                ..defaults
            },
            PowerPointProgTagLimits {
                max_string_code_units: 1,
                ..defaults
            },
            PowerPointProgTagLimits {
                max_binary_payload_bytes: 3,
                ..defaults
            },
        ];
        for limits in cases {
            assert!(
                PowerPointProgTags::parse_payload(
                    &valid,
                    0,
                    PowerPointProgTagScope::Document,
                    limits
                )
                .is_err()
            );
        }

        let known = binary_tag("___PPT12", &versioned_payload());
        assert!(
            PowerPointProgTags::parse_payload(
                &known,
                0,
                PowerPointProgTagScope::Document,
                PowerPointProgTagLimits {
                    max_binary_records: 1,
                    ..defaults
                }
            )
            .is_err()
        );
    }

    fn parsed_container(version: u16, kind: PptRecordType, payload: &[u8]) -> PptRecord {
        let bytes = record(version, 0, kind.as_u16(), payload);
        let (parsed, consumed) = PptRecord::parse_strict(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        parsed
    }

    #[test]
    fn parse_document_locates_prog_tags_inside_doc_info_list() {
        let limits = PowerPointProgTagLimits::default();
        let tags_payload = string_tag("author", None);
        let prog_tags = record(0x0f, 0, PptRecordType::ProgTags.as_u16(), &tags_payload);
        let doc_info_list = record(0x0f, 0, PptRecordType::DocInfoList.as_u16(), &prog_tags);
        let document = parsed_container(0x0f, PptRecordType::Document, &doc_info_list);

        let parsed = PowerPointProgTags::parse_document(&document, limits)
            .unwrap()
            .unwrap();
        assert_eq!(parsed.scope, PowerPointProgTagScope::Document);
        assert_eq!(parsed.tags.len(), 1);

        // A document without a DocInfoListContainer has no tags.
        let bare = parsed_container(0x0f, PptRecordType::Document, &[]);
        assert!(
            PowerPointProgTags::parse_document(&bare, limits)
                .unwrap()
                .is_none()
        );

        // A DocInfoListContainer without ProgTags has no tags.
        let empty_list = record(0x0f, 0, PptRecordType::DocInfoList.as_u16(), &[]);
        let no_tags = parsed_container(0x0f, PptRecordType::Document, &empty_list);
        assert!(
            PowerPointProgTags::parse_document(&no_tags, limits)
                .unwrap()
                .is_none()
        );

        // Duplicate DocInfoListContainer or ProgTags children are rejected.
        let duplicate_list = parsed_container(
            0x0f,
            PptRecordType::Document,
            &[doc_info_list.clone(), doc_info_list.clone()].concat(),
        );
        assert!(PowerPointProgTags::parse_document(&duplicate_list, limits).is_err());
        let duplicate_tags = record(
            0x0f,
            0,
            PptRecordType::DocInfoList.as_u16(),
            &[prog_tags.clone(), prog_tags].concat(),
        );
        let duplicate_tags = parsed_container(0x0f, PptRecordType::Document, &duplicate_tags);
        assert!(PowerPointProgTags::parse_document(&duplicate_tags, limits).is_err());

        // A non-Document record cannot provide document tags.
        let slide = parsed_container(0x0f, PptRecordType::Slide, &[]);
        assert!(PowerPointProgTags::parse_document(&slide, limits).is_err());
    }

    #[test]
    fn parse_slide_locates_direct_prog_tags_child() {
        let limits = PowerPointProgTagLimits::default();
        let prog_tags = record(
            0x0f,
            0,
            PptRecordType::ProgTags.as_u16(),
            &binary_tag("___PPT9", &versioned_payload()),
        );
        let slide = parsed_container(0x0f, PptRecordType::Slide, &prog_tags);

        let parsed = PowerPointProgTags::parse_slide(&slide, limits)
            .unwrap()
            .unwrap();
        assert_eq!(parsed.scope, PowerPointProgTagScope::Slide);
        assert!(
            parsed
                .binary_tag(PowerPointProgBinaryTagVersion::PowerPoint9)
                .is_some()
        );

        let bare = parsed_container(0x0f, PptRecordType::Slide, &[]);
        assert!(
            PowerPointProgTags::parse_slide(&bare, limits)
                .unwrap()
                .is_none()
        );

        let duplicate = parsed_container(
            0x0f,
            PptRecordType::Slide,
            &[prog_tags.clone(), prog_tags].concat(),
        );
        assert!(PowerPointProgTags::parse_slide(&duplicate, limits).is_err());
    }
}
