use super::model::{
    ProgBinaryTag, ProgBinaryTagVersion, ProgStringTag, ProgTag, ProgTagLimits, ProgTagScope,
    ProgTags,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

impl ProgTagScope {
    /// Return the assigned version for a tag name in this scope, if any.
    fn known_version(self, name: &str) -> Option<ProgBinaryTagVersion> {
        match (self, name) {
            (_, "___PPT9") => Some(ProgBinaryTagVersion::PowerPoint9),
            (_, "___PPT10") => Some(ProgBinaryTagVersion::PowerPoint10),
            (Self::Document, "___PPT11") => Some(ProgBinaryTagVersion::PowerPoint11),
            (_, "___PPT12") => Some(ProgBinaryTagVersion::PowerPoint12),
            _ => None,
        }
    }
}

impl ProgBinaryTag {
    /// Decode the retained `BinaryTagDataBlob` payload as a strict sequence of
    /// PPT records.
    ///
    /// Versioned tag payloads are guaranteed to decode successfully because
    /// they were validated at parse time; unknown payloads may hold arbitrary
    /// bytes and can fail here.
    pub fn records(&self) -> Result<Vec<Record>> {
        Record::parse_sequence_strict(&self.payload, "BinaryTagDataBlob payload")
    }
}

impl ProgTags {
    /// Parse the `DocProgTagsContainer` of a `DocumentContainer` record.
    ///
    /// The container is an optional child of the document's
    /// `DocInfoListContainer` (MS-PPT section 2.4.4); each record type MUST
    /// NOT occur there more than once, so duplicates are an error.
    pub fn parse_document(document: &Record, limits: ProgTagLimits) -> Result<Option<Self>> {
        if document.record_type != RecordType::Document {
            return corrupted("Document ProgTags require a DocumentContainer record");
        }
        let Some(doc_info_list) = single_child(document, RecordType::DocInfoList)? else {
            return Ok(None);
        };
        let Some(record) = single_child(doc_info_list, RecordType::ProgTags)? else {
            return Ok(None);
        };
        Self::parse(record, ProgTagScope::Document, limits).map(Some)
    }

    /// Parse the `SlideProgTagsContainer` of a slide-like container record
    /// (`SlideContainer`, `NotesContainer`, `MainMasterContainer`, or
    /// `HandoutContainer`; MS-PPT section 2.5.19).
    pub fn parse_slide(container: &Record, limits: ProgTagLimits) -> Result<Option<Self>> {
        let Some(record) = single_child(container, RecordType::ProgTags)? else {
            return Ok(None);
        };
        Self::parse(record, ProgTagScope::Slide, limits).map(Some)
    }

    /// Parse and validate a complete `DocProgTagsContainer` or
    /// `SlideProgTagsContainer` record.
    pub fn parse(record: &Record, scope: ProgTagScope, limits: ProgTagLimits) -> Result<Self> {
        if record.record_type != RecordType::ProgTags
            || record.record_type_raw != RecordType::ProgTags.as_u16()
            || record.version != 0x0f
        {
            return corrupted("Invalid ProgTags container record header");
        }
        let declared = usize::try_from(record.data_length)
            .map_err(|_| Error::Corrupted("ProgTags container size overflow".into()))?;
        if declared != record.data.len() {
            return corrupted("ProgTags container length does not match its payload");
        }
        Self::parse_payload(&record.data, record.instance, scope, limits)
    }

    /// Parse a `ProgTags` container payload and its record instance.
    pub fn parse_payload(
        data: &[u8],
        instance: u16,
        scope: ProgTagScope,
        limits: ProgTagLimits,
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
                value if value == RecordType::ProgStringTag.as_u16() => {
                    ProgTag::String(parse_string_tag(record, limits)?)
                },
                value if value == RecordType::ProgBinaryTag.as_u16() => {
                    let tag = parse_binary_tag(record, scope, limits)?;
                    // Sections 2.4.23.1 and 2.5.19: the array MUST NOT contain
                    // more than one of each versioned extension.
                    let duplicate = match tag.version {
                        ProgBinaryTagVersion::PowerPoint9 => std::mem::replace(&mut seen9, true),
                        ProgBinaryTagVersion::PowerPoint10 => std::mem::replace(&mut seen10, true),
                        ProgBinaryTagVersion::PowerPoint11 => std::mem::replace(&mut seen11, true),
                        ProgBinaryTagVersion::PowerPoint12 => std::mem::replace(&mut seen12, true),
                        ProgBinaryTagVersion::Unknown => false,
                    };
                    if duplicate {
                        return corrupted(
                            "ProgTags container contains a duplicate versioned binary tag",
                        );
                    }
                    ProgTag::Binary(tag)
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
    pub fn to_bytes(&self, limits: ProgTagLimits) -> Result<Vec<u8>> {
        let payload = self.to_payload(limits)?;
        encode_record(0x0f, self.instance, RecordType::ProgTags.as_u16(), &payload)
    }

    /// Serialize only the `ProgTags` container payload.
    pub fn to_payload(&self, limits: ProgTagLimits) -> Result<Vec<u8>> {
        check_limit(self.tags.len(), limits.max_tags, "programmable tag count")?;
        let mut payload = Vec::new();
        for tag in &self.tags {
            let encoded = match tag {
                ProgTag::String(tag) => encode_string_tag(tag)?,
                ProgTag::Binary(tag) => encode_binary_tag(tag)?,
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
    pub fn to_record(&self, limits: ProgTagLimits) -> Result<Record> {
        let data = self.to_payload(limits)?;
        let data_length = u32::try_from(data.len())
            .map_err(|_| Error::Corrupted("ProgTags payload exceeds u32".into()))?;
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

    /// Return the binary tag for an assigned version, when present.
    pub fn binary_tag(&self, version: ProgBinaryTagVersion) -> Option<&ProgBinaryTag> {
        self.tags.iter().find_map(|tag| match tag {
            ProgTag::Binary(tag) if tag.version == version => Some(tag),
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

fn parse_string_tag(record: RawRecord, limits: ProgTagLimits) -> Result<ProgStringTag> {
    require_container_header(&record, RecordType::ProgStringTag, "ProgStringTagContainer")?;
    let children = parse_sequence(&record.data, 2, "ProgStringTagContainer")?;
    if children.is_empty() || children.len() > 2 {
        return corrupted("ProgStringTagContainer must contain a name and at most one value");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let value_units = children
        .get(1)
        .map(|value| parse_cstring_atom(value, 1, false, limits))
        .transpose()?;
    Ok(ProgStringTag {
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
    scope: ProgTagScope,
    limits: ProgTagLimits,
) -> Result<ProgBinaryTag> {
    require_container_header(&record, RecordType::ProgBinaryTag, "ProgBinaryTagContainer")?;
    let children = parse_sequence(&record.data, 2, "ProgBinaryTagContainer")?;
    if children.len() != 2 {
        return corrupted("ProgBinaryTagContainer must contain exactly one CString/blob pair");
    }
    let name_units = parse_cstring_atom(&children[0], 0, true, limits)?;
    let name = decode_units(&name_units, "binary programmable tag name")?;
    let blob = &children[1];
    require_atom_header(blob, RecordType::BinaryTagData, 0, "BinaryTagDataBlob")?;
    check_limit(
        blob.data.len(),
        limits.max_binary_payload_bytes,
        "binary programmable tag payload",
    )?;

    let version = scope
        .known_version(name.as_str())
        .unwrap_or(ProgBinaryTagVersion::Unknown);
    if version != ProgBinaryTagVersion::Unknown {
        // Sections 2.4.23.5 and 2.5.23: the versioned tag name has a fixed
        // length with no NUL padding, and the blob is a strict record sequence.
        if name_units.len() != name.len() {
            return corrupted("Versioned binary tag has an invalid CString length");
        }
        let records = Record::parse_sequence_strict(&blob.data, "versioned BinaryTagDataBlob")?;
        check_limit(
            records.len(),
            limits.max_binary_records,
            "versioned binary tag record count",
        )?;
    }

    Ok(ProgBinaryTag {
        name,
        version,
        payload: blob.data.clone(),
        name_units,
    })
}

fn encode_string_tag(tag: &ProgStringTag) -> Result<Vec<u8>> {
    let name = encode_cstring_atom(0, &tag.name_units)?;
    let mut payload = name;
    if let Some(value) = &tag.value_units {
        payload.extend_from_slice(&encode_cstring_atom(1, value)?);
    }
    encode_record(0x0f, 0, RecordType::ProgStringTag.as_u16(), &payload)
}

fn encode_binary_tag(tag: &ProgBinaryTag) -> Result<Vec<u8>> {
    let mut payload = encode_cstring_atom(0, &tag.name_units)?;
    payload.extend_from_slice(&encode_record(
        0,
        0,
        RecordType::BinaryTagData.as_u16(),
        &tag.payload,
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
    limits: ProgTagLimits,
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
        .map_err(|_| Error::Corrupted(format!("{field} contains invalid UTF-16")))
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
    let length = u32::from_le_bytes([
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]);
    let length = usize::try_from(length)
        .map_err(|_| Error::Corrupted(format!("{context} record length overflow")))?;
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
        .map_err(|_| Error::Corrupted("PPT record payload exceeds u32".into()))?;
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
fn single_child(record: &Record, kind: RecordType) -> Result<Option<&Record>> {
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
    Err(Error::Corrupted(message.into()))
}
