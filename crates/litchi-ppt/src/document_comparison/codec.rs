//! Binary record parsing and serialization for document-comparison metadata.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

use super::model::{
    DiffFlags, DiffNode, DiffRecordHeaders, DiffTree10, DiffType, Entry, Limits,
    POWERPOINT_DIFF_MAX_DEPTH, POWERPOINT_DIFF_MAX_RECORDS, Review, ReviewingToolbarStates,
    SlideCreationEntry, SlideListTable10, Unknown,
};
use super::validation::{
    MAX_REVIEWER_NAME_BYTES, corrupted, validate_atom, validate_count, validate_reviewer_name,
};

pub(super) const RECORD_HEADER_SIZE: usize = 8;
pub(super) const SIZE_ATOM_PAYLOAD_SIZE: usize = 4;
pub(super) const ENTRY_PAYLOAD_SIZE: usize = 12;
pub(super) const SIZE_ATOM_RECORD_SIZE: usize = RECORD_HEADER_SIZE + SIZE_ATOM_PAYLOAD_SIZE;
pub(super) const ENTRY_RECORD_SIZE: usize = RECORD_HEADER_SIZE + ENTRY_PAYLOAD_SIZE;
pub(super) const DIFF_HEADER_SIZE: usize = 28;
const DIFF_FIXED_SIZE: usize = 32;
const REVIEWER_NAME_RECORD_TYPE: u16 = 0x0FBA;

#[derive(Debug, Clone, Copy)]
struct DiffLimits {
    max_depth: usize,
    max_records: usize,
}

impl DiffNode {
    /// Serialize this node as a `Diff10` container record with its children.
    ///
    /// # Errors
    ///
    /// Returns an error if the tree exceeds the nesting or record-count
    /// limits, or if a node fails structural validation.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut count = 0;
        self.encode(0, &mut count)
    }

    fn parse(
        data: &[u8],
        depth: usize,
        limits: DiffLimits,
        count: &mut usize,
    ) -> Result<(Self, usize)> {
        if depth > limits.max_depth {
            return corrupted("document-comparison tree exceeds the nesting limit");
        }
        *count = count.checked_add(1).ok_or_else(|| {
            Error::Corrupted("document-comparison record count overflow".to_string())
        })?;
        if *count > limits.max_records {
            return corrupted("document-comparison tree exceeds the record-count limit");
        }
        if data.len() < DIFF_FIXED_SIZE {
            return corrupted("truncated Diff10 container");
        }
        let (version, instance) = unpack_version_instance(read_u16(data, 0)?);
        if version != 0xF || instance != 0 || read_u16(data, 2)? != RecordType::Diff10.as_u16() {
            return corrupted("invalid Diff10 record header");
        }
        let payload_len = usize::try_from(read_u32(data, 4)?)
            .map_err(|_err| Error::Corrupted("Diff10 length does not fit usize".to_string()))?;
        let total_len = RECORD_HEADER_SIZE
            .checked_add(payload_len)
            .ok_or_else(|| Error::Corrupted("Diff10 length overflow".to_string()))?;
        if total_len > data.len() || payload_len < DIFF_FIXED_SIZE - RECORD_HEADER_SIZE {
            return corrupted("Diff10 record extends beyond its parent");
        }
        let (atom_version, atom_instance) = unpack_version_instance(read_u16(data, 8)?);
        if atom_version != 0
            || atom_instance != 0
            || read_u16(data, 10)? != RecordType::Diff10Atom.as_u16()
            || read_u32(data, 12)? != 12
        {
            return corrupted("invalid Diff10Atom header");
        }
        let index = match data[16] {
            0 => false,
            1 => true,
            _ => return corrupted("Diff10Atom fIndex is not a bool1"),
        };
        let diff_type = DiffType::try_from(read_u32(data, 20)?)?;
        if index
            && !matches!(
                diff_type,
                DiffType::HeaderFooter | DiffType::InteractiveInfo
            )
        {
            return corrupted("Diff10 fIndex is invalid for its diff type");
        }
        let headers = DiffRecordHeaders {
            index,
            diff_type,
            ignored_prefix: [data[17], data[18], data[19]],
            ignored_tail: read_u32(data, 24)?,
        };
        let raw_flags = read_u32(data, DIFF_HEADER_SIZE)?;
        let (flags, ignored_flag_bits) = DiffFlags::from_raw(diff_type, raw_flags);
        let mut children = Vec::new();
        let mut offset = DIFF_FIXED_SIZE;
        while offset < total_len {
            let (child, consumed) =
                Self::parse(&data[offset..total_len], depth + 1, limits, count)?;
            if consumed == 0 {
                return corrupted("zero-length Diff10 child");
            }
            children.push(child);
            offset += consumed;
        }
        let node = Self {
            headers,
            flags,
            ignored_flag_bits,
            children,
        };
        node.validate_node()?;
        Ok((node, total_len))
    }

    fn encode(&self, depth: usize, count: &mut usize) -> Result<Vec<u8>> {
        if depth > POWERPOINT_DIFF_MAX_DEPTH {
            return corrupted("document-comparison tree exceeds the nesting limit");
        }
        *count = count.checked_add(1).ok_or_else(|| {
            Error::Corrupted("document-comparison record count overflow".to_string())
        })?;
        if *count > POWERPOINT_DIFF_MAX_RECORDS {
            return corrupted("document-comparison tree exceeds the record-count limit");
        }
        self.validate_node()?;
        let mut atom = Vec::with_capacity(12);
        atom.push(u8::from(self.headers.index));
        atom.extend_from_slice(&self.headers.ignored_prefix);
        atom.extend_from_slice(&self.headers.diff_type.as_u32().to_le_bytes());
        atom.extend_from_slice(&self.headers.ignored_tail.to_le_bytes());
        let mut payload = encode_record(0, 0, RecordType::Diff10Atom, &atom);
        payload.extend_from_slice(&(self.flags.to_raw() | self.ignored_flag_bits).to_le_bytes());
        for child in &self.children {
            payload.extend_from_slice(&child.encode(depth + 1, count)?);
        }
        Ok(encode_record(0xF, 0, RecordType::Diff10, &payload))
    }
}

impl DiffTree10 {
    /// Parse a strict `DiffTree10Container` record under the default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the container header, reviewer-name atom, or any
    /// nested `Diff10` record is malformed or truncated.
    pub fn parse(record: &Record) -> Result<Self> {
        Self::parse_with_limits(
            record,
            POWERPOINT_DIFF_MAX_DEPTH,
            POWERPOINT_DIFF_MAX_RECORDS,
        )
    }

    /// Parse a strict `DiffTree10Container` record under explicit limits.
    ///
    /// The limits are clamped to the supported maximums before parsing.
    ///
    /// # Errors
    ///
    /// Returns an error if `max_records` is zero, or if the container header,
    /// reviewer-name atom, or any nested `Diff10` record is malformed or
    /// truncated.
    pub fn parse_with_limits(
        record: &Record,
        max_depth: usize,
        max_records: usize,
    ) -> Result<Self> {
        if max_records == 0 {
            return corrupted("document-comparison record limit must be nonzero");
        }
        if record.record_type_raw != RecordType::DiffTree10.as_u16()
            || record.version != 0xF
            || record.instance != 0
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
        {
            return corrupted("invalid DiffTree10 container header");
        }
        let limits = DiffLimits {
            max_depth: max_depth.min(POWERPOINT_DIFF_MAX_DEPTH),
            max_records: max_records.min(POWERPOINT_DIFF_MAX_RECORDS),
        };
        let (reviewer_name, reviewer_len) = parse_reviewer_name(&record.data)?;
        let mut count = 0;
        let (document_diff, diff_len) =
            DiffNode::parse(&record.data[reviewer_len..], 0, limits, &mut count)?;
        if document_diff.diff_type() != DiffType::Document {
            return corrupted("DiffTree10 root is not a DocDiff10 container");
        }
        if reviewer_len + diff_len != record.data.len() {
            return corrupted("DiffTree10 has trailing records");
        }
        Ok(Self {
            reviewer_name,
            document_diff,
        })
    }

    /// Serialize this tree as a `DiffTree10Container` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the reviewer name is not a valid
    /// `PrintableUnicodeString`, if the root node is not a `DocDiff10`
    /// container, or if a nested node fails structural validation.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        validate_reviewer_name(&self.reviewer_name)?;
        if self.document_diff.diff_type() != DiffType::Document {
            return corrupted("DiffTree10 root is not a DocDiff10 container");
        }
        let mut reviewer_payload = Vec::new();
        for unit in self.reviewer_name.encode_utf16() {
            reviewer_payload.extend_from_slice(&unit.to_le_bytes());
        }
        let mut payload = encode_record_raw(0, 0, REVIEWER_NAME_RECORD_TYPE, &reviewer_payload);
        payload.extend_from_slice(&self.document_diff.to_record_bytes()?);
        Ok(encode_record(0xF, 0, RecordType::DiffTree10, &payload))
    }
}

impl ReviewingToolbarStates {
    /// Construct reviewing UI state. Reserved bits are emitted as zero.
    #[must_use]
    pub const fn new(show_reviewing_toolbar: bool, show_reviewing_gallery: bool) -> Self {
        Self {
            show_reviewing_toolbar,
            show_reviewing_gallery,
            ignored_reserved_bits: 0,
        }
    }

    /// Parse a strict `DocToolbarStates10Atom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record) -> Result<Self> {
        validate_atom(record, RecordType::DocToolbarStates10Atom, 1)?;
        let value = record.data[0];
        Ok(Self {
            show_reviewing_toolbar: value & 0x01 != 0,
            show_reviewing_gallery: value & 0x02 != 0,
            // MS-PPT requires writers to zero these bits and readers to ignore
            // them. Retaining them keeps a parsed record byte-stable.
            ignored_reserved_bits: value & 0xfc,
        })
    }

    #[must_use]
    pub const fn show_reviewing_toolbar(self) -> bool {
        self.show_reviewing_toolbar
    }

    #[must_use]
    pub const fn show_reviewing_gallery(self) -> bool {
        self.show_reviewing_gallery
    }

    /// Raw ignored bits retained from an existing record.
    #[must_use]
    pub const fn ignored_reserved_bits(self) -> u8 {
        self.ignored_reserved_bits
    }

    pub fn set_show_reviewing_toolbar(&mut self, value: bool) {
        self.show_reviewing_toolbar = value;
    }

    pub fn set_show_reviewing_gallery(&mut self, value: bool) {
        self.show_reviewing_gallery = value;
    }

    /// Serialize the exact atom, preserving ignored bits from parsed input.
    #[must_use]
    pub fn to_record_bytes(self) -> Vec<u8> {
        let value = self.ignored_reserved_bits
            | u8::from(self.show_reviewing_toolbar)
            | (u8::from(self.show_reviewing_gallery) << 1);
        encode_record(0, 0, RecordType::DocToolbarStates10Atom, &[value])
    }
}

impl SlideCreationEntry {
    fn parse_payload(payload: &[u8]) -> Self {
        let slide_id_ref = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
        let high = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
        let low = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
        Self {
            slide_id_ref,
            file_time: (u64::from(high) << 32) | u64::from(low),
        }
    }

    /// Serialize the fixed-size entry payload.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the low 32 bits of the FILETIME are intentionally written to the final payload word"
    )]
    fn payload(self) -> [u8; ENTRY_PAYLOAD_SIZE] {
        let mut payload = [0; ENTRY_PAYLOAD_SIZE];
        payload[..4].copy_from_slice(&self.slide_id_ref.to_le_bytes());
        payload[4..8].copy_from_slice(&(self.file_time >> 32).to_le_bytes()[..4]);
        payload[8..].copy_from_slice(&(self.file_time as u32).to_le_bytes());
        payload
    }

    /// Serialize a strict `SlideListEntry10Atom` record.
    #[must_use]
    pub fn to_record_bytes(self) -> Vec<u8> {
        encode_record(0, 0, RecordType::SlideListEntry10Atom, &self.payload())
    }
}

impl SlideListTable10 {
    /// Parse a strict container without allocating intermediate child records.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    ///
    /// # Panics
    ///
    /// Panics if internal invariants are violated.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::SlideListTable10
            || record.version != 0x0f
            || record.instance != 0
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
        {
            return corrupted("SlideListTable10Container has an invalid record header");
        }
        let data = record.data.as_slice();
        if data.len() < SIZE_ATOM_RECORD_SIZE {
            return corrupted("SlideListTable10Container is truncated");
        }
        let size_payload = child_payload(
            data,
            0,
            RecordType::SlideListTableSize10Atom,
            SIZE_ATOM_PAYLOAD_SIZE,
        )?;
        let signed_count = i32::from_le_bytes([
            size_payload[0],
            size_payload[1],
            size_payload[2],
            size_payload[3],
        ]);
        let count = usize::try_from(signed_count)
            .map_err(|_err| Error::Corrupted("negative slide-list table count".to_string()))?;
        validate_count(count)?;
        let expected = count
            .checked_mul(ENTRY_RECORD_SIZE)
            .and_then(|size| size.checked_add(SIZE_ATOM_RECORD_SIZE))
            .ok_or_else(|| Error::Corrupted("slide-list table size overflow".to_string()))?;
        if data.len() != expected {
            return corrupted("SlideListTable10Container count does not match its payload");
        }

        let mut entries = Vec::with_capacity(count);
        let mut offset = SIZE_ATOM_RECORD_SIZE;
        for _ in 0..count {
            let payload = child_payload(
                data,
                offset,
                RecordType::SlideListEntry10Atom,
                ENTRY_PAYLOAD_SIZE,
            )?;
            entries.push(SlideCreationEntry::parse_payload(payload));
            offset += ENTRY_RECORD_SIZE;
        }
        Ok(Self { entries })
    }

    /// Serialize the size atom followed by the exact declared entry array.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        validate_count(self.entries.len())?;
        let count = i32::try_from(self.entries.len())
            .map_err(|_err| Error::Corrupted("slide-list table count overflow".to_string()))?;
        let capacity = SIZE_ATOM_RECORD_SIZE
            .checked_add(self.entries.len().saturating_mul(ENTRY_RECORD_SIZE))
            .ok_or_else(|| Error::Corrupted("slide-list table size overflow".to_string()))?;
        let mut payload = Vec::with_capacity(capacity);
        payload.extend_from_slice(&encode_record(
            0,
            0,
            RecordType::SlideListTableSize10Atom,
            &count.to_le_bytes(),
        ));
        for entry in &self.entries {
            payload.extend_from_slice(&entry.to_record_bytes());
        }
        Ok(encode_record(
            0x0f,
            0,
            RecordType::SlideListTable10,
            &payload,
        ))
    }
}

/// Read the review-owned records from the document's `___PPT10` payload.
pub(crate) fn read_review(root: &Record, limits: Limits) -> Result<Review> {
    let Some(path) = find_pp10_blob(root)? else {
        return Ok(Review::default());
    };
    let blob = record_at(root, &path)?;
    parse_payload(&blob.data, limits)
}

/// Replace the inert PP10 payload represented by a review view.
pub(crate) fn write_review(root: &mut Record, review: &Review, limits: Limits) -> Result<()> {
    let path = find_pp10_blob(root)?
        .ok_or_else(|| Error::InvalidFormat("document has no ___PPT10 review payload".into()))?;
    let payload = encode_payload(review, limits)?;
    let blob = record_at_mut(root, &path)?;
    blob.data_length = u32::try_from(payload.len())
        .map_err(|_err| Error::InvalidFormat("review payload exceeds u32".into()))?;
    blob.data = payload;
    Ok(())
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "`RecordType` has hundreds of variants; every record type not owned by the review payload is deliberately retained as an opaque `Unknown` entry"
)]
fn parse_payload(data: &[u8], limits: Limits) -> Result<Review> {
    if data.len() > limits.max_bytes {
        return Err(Error::InvalidFormat(
            "review payload exceeds the snapshot byte limit".into(),
        ));
    }
    let mut entries = Vec::new();
    let mut offset = 0usize;
    let mut count = 0usize;
    let mut last_rank = 0usize;
    let mut seen_toolbar = false;
    let mut seen_slide_list = false;
    while offset < data.len() {
        count = count
            .checked_add(1)
            .ok_or_else(|| Error::Corrupted("review record count overflow".into()))?;
        if count > limits.max_review_records {
            return Err(Error::InvalidFormat(
                "review payload exceeds the record limit".into(),
            ));
        }
        let header_end = offset
            .checked_add(RECORD_HEADER_SIZE)
            .ok_or_else(|| Error::Corrupted("review record offset overflow".into()))?;
        if header_end > data.len() {
            return corrupted("review payload ends with a truncated record header");
        }
        let length = usize::try_from(read_u32(data, offset + 4)?)
            .map_err(|_err| Error::Corrupted("review record length overflow".into()))?;
        let end = header_end
            .checked_add(length)
            .ok_or_else(|| Error::Corrupted("review record length overflow".into()))?;
        if end > data.len() {
            return corrupted("review record extends beyond its payload");
        }
        let (record, consumed) = Record::parse_strict(&data[offset..end], 0)?;
        if consumed != end - offset {
            return corrupted("review record was only partially parsed");
        }
        let raw = data[offset..end].to_vec();
        let (entry, rank) = match record.record_type {
            RecordType::DocToolbarStates10Atom => {
                if seen_toolbar {
                    return corrupted("review payload contains duplicate toolbar state atoms");
                }
                seen_toolbar = true;
                (Entry::Toolbar(ReviewingToolbarStates::parse(&record)?), 0)
            },
            RecordType::SlideListTable10 => {
                if seen_slide_list {
                    return corrupted("review payload contains duplicate slide-list tables");
                }
                seen_slide_list = true;
                (Entry::SlideList(SlideListTable10::parse(&record)?), 1)
            },
            RecordType::DiffTree10 => (Entry::Diff(DiffTree10::parse(&record)?), 2),
            _ => (
                Entry::Unknown(Unknown::new(
                    record.record_type_raw,
                    record.version,
                    record.instance,
                    raw,
                )),
                last_rank,
            ),
        };
        if !matches!(entry, Entry::Unknown(_)) {
            if rank < last_rank {
                return corrupted("review records are out of order");
            }
            last_rank = rank;
        }
        entries.push(entry);
        offset = end;
    }
    Ok(Review { entries })
}

fn encode_payload(review: &Review, limits: Limits) -> Result<Vec<u8>> {
    if review.entries.len() > limits.max_review_records {
        return Err(Error::InvalidFormat(
            "review payload exceeds the record limit".into(),
        ));
    }
    let mut payload = Vec::new();
    let mut toolbar = false;
    let mut slide_list = false;
    let mut last_rank = 0usize;
    for entry in &review.entries {
        let (bytes, rank) = match entry {
            Entry::Toolbar(value) => {
                if toolbar {
                    return corrupted("review payload contains duplicate toolbar state atoms");
                }
                toolbar = true;
                (value.to_record_bytes(), 0)
            },
            Entry::SlideList(value) => {
                if slide_list {
                    return corrupted("review payload contains duplicate slide-list tables");
                }
                slide_list = true;
                (value.to_record_bytes()?, 1)
            },
            Entry::Diff(value) => (value.to_record_bytes()?, 2),
            Entry::Unknown(value) => (value.bytes().to_vec(), last_rank),
        };
        if !matches!(entry, Entry::Unknown(_)) {
            if rank < last_rank {
                return corrupted("review records are out of order");
            }
            last_rank = rank;
        }
        payload.extend_from_slice(&bytes);
        if payload.len() > limits.max_bytes {
            return Err(Error::InvalidFormat(
                "review payload exceeds the snapshot byte limit".into(),
            ));
        }
    }
    Ok(payload)
}

pub(crate) fn encode_document(root: &Record) -> Result<Vec<u8>> {
    let payload = if root.children.is_empty() {
        root.data.clone()
    } else {
        let mut payload = Vec::new();
        for child in &root.children {
            payload.extend_from_slice(&encode_document(child)?);
        }
        payload
    };
    if root.version > 0x0f || root.instance > 0x0fff {
        return Err(Error::InvalidFormat(
            "document-comparison record header fields exceed their wire widths".into(),
        ));
    }
    let length = u32::try_from(payload.len())
        .map_err(|_err| Error::InvalidFormat("document-comparison record exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(RECORD_HEADER_SIZE + payload.len());
    bytes.extend_from_slice(&(root.version | (root.instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&root.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(&payload);
    Ok(bytes)
}

fn find_pp10_blob(root: &Record) -> Result<Option<Vec<usize>>> {
    if root.record_type != RecordType::Document {
        return corrupted("document-comparison owner requires a DocumentContainer root");
    }

    let mut doc_info_iter = root
        .children
        .iter()
        .enumerate()
        .filter(|(_, child)| child.record_type == RecordType::DocInfoList);
    let Some((doc_info_index, doc_info)) = doc_info_iter.next() else {
        return Ok(None);
    };
    if doc_info_iter.next().is_some() {
        return corrupted("document contains duplicate DocInfoList containers");
    }

    let mut prog_tags_iter = doc_info
        .children
        .iter()
        .enumerate()
        .filter(|(_, child)| child.record_type == RecordType::ProgTags);
    let Some((prog_tags_index, prog_tags)) = prog_tags_iter.next() else {
        return Ok(None);
    };
    if prog_tags_iter.next().is_some() {
        return corrupted("document info contains duplicate ProgTags containers");
    }

    let mut match_path = None;
    for (tag_index, tag) in prog_tags.children.iter().enumerate() {
        if tag.record_type != RecordType::ProgBinaryTag || !is_pp10_tag(tag)? {
            continue;
        }
        if match_path.is_some() {
            return corrupted("document contains duplicate ___PPT10 review payloads");
        }
        let mut blob = None;
        for (child_index, child) in tag.children.iter().enumerate() {
            if child.record_type == RecordType::BinaryTagData && blob.replace(child_index).is_some()
            {
                return corrupted("___PPT10 tag contains duplicate BinaryTagData records");
            }
        }
        let blob_index =
            blob.ok_or_else(|| Error::Corrupted("___PPT10 tag is missing BinaryTagData".into()))?;
        match_path = Some(vec![doc_info_index, prog_tags_index, tag_index, blob_index]);
    }
    Ok(match_path)
}

fn is_pp10_tag(record: &Record) -> Result<bool> {
    let Some(name) = record
        .children
        .iter()
        .find(|child| child.record_type == RecordType::CString)
    else {
        return Ok(false);
    };
    if name.version != 0 || name.instance != 0 || name.data.len() % 2 != 0 {
        return corrupted("ProgBinaryTag has an invalid tag-name record");
    }
    let units = name
        .data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    Ok(String::from_utf16(&units).ok().as_deref() == Some("___PPT10"))
}

fn record_at<'a>(root: &'a Record, path: &[usize]) -> Result<&'a Record> {
    let mut record = root;
    for index in path {
        record = record
            .children
            .get(*index)
            .ok_or_else(|| Error::Corrupted("review record path is out of range".into()))?;
    }
    Ok(record)
}

fn record_at_mut<'a>(root: &'a mut Record, path: &[usize]) -> Result<&'a mut Record> {
    let mut record = root;
    for index in path {
        record = record
            .children
            .get_mut(*index)
            .ok_or_else(|| Error::InvalidFormat("review record path is out of range".into()))?;
    }
    Ok(record)
}

fn parse_reviewer_name(data: &[u8]) -> Result<(String, usize)> {
    if data.len() < RECORD_HEADER_SIZE {
        return corrupted("DiffTree10 is missing ReviewerNameAtom");
    }
    let (version, instance) = unpack_version_instance(read_u16(data, 0)?);
    let byte_len = usize::try_from(read_u32(data, 4)?)
        .map_err(|_err| Error::Corrupted("reviewer-name length does not fit usize".to_string()))?;
    let total_len = RECORD_HEADER_SIZE
        .checked_add(byte_len)
        .ok_or_else(|| Error::Corrupted("reviewer-name length overflow".to_string()))?;
    if version != 0
        || instance != 0
        || read_u16(data, 2)? != REVIEWER_NAME_RECORD_TYPE
        || byte_len > MAX_REVIEWER_NAME_BYTES
        || byte_len % 2 != 0
        || total_len > data.len()
    {
        return corrupted("invalid ReviewerNameAtom");
    }
    let mut units = Vec::with_capacity(byte_len / 2);
    for chunk in data[RECORD_HEADER_SIZE..total_len].chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    let name = String::from_utf16(&units)
        .map_err(|_err| Error::Corrupted("ReviewerNameAtom contains invalid UTF-16".to_string()))?;
    validate_reviewer_name(&name)?;
    Ok((name, total_len))
}

fn unpack_version_instance(value: u16) -> (u16, u16) {
    (value & 0x000F, value >> 4)
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison field".to_string()))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison field".to_string()))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn child_payload(
    data: &[u8],
    offset: usize,
    kind: RecordType,
    payload_size: usize,
) -> Result<&[u8]> {
    let header_end = offset
        .checked_add(RECORD_HEADER_SIZE)
        .ok_or_else(|| Error::Corrupted("record offset overflow".to_string()))?;
    let header = data
        .get(offset..header_end)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison child".to_string()))?;
    let version_instance = u16::from_le_bytes([header[0], header[1]]);
    let raw_type = u16::from_le_bytes([header[2], header[3]]);
    let length = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
    if version_instance != 0 || raw_type != kind.as_u16() || length as usize != payload_size {
        return corrupted("document-comparison child has an invalid record header");
    }
    let end = header_end
        .checked_add(payload_size)
        .ok_or_else(|| Error::Corrupted("record length overflow".to_string()))?;
    data.get(header_end..end)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison payload".to_string()))
}

fn encode_record(version: u16, instance: u16, kind: RecordType, payload: &[u8]) -> Vec<u8> {
    encode_record_raw(version, instance, kind.as_u16(), payload)
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "payloads are bounded to the 16 MiB snapshot byte limit enforced by callers, well below u32::MAX"
)]
fn encode_record_raw(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(RECORD_HEADER_SIZE + payload.len());
    bytes.extend_from_slice(&((instance << 4) | (version & 0xF)).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}
