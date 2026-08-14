//! Core PPT record structure and parsing.
//!
//! This module implements the fundamental PPT record parsing based on
//! Apache POI's HSLF Record.java implementation.

use super::{DocumentInfo, SlideAtomsSet, SlideInfo};
use crate::consts::RecordType;
use crate::package::{Error, RecordLimits, Result};
use crate::text::extractor::{parse_cstring, parse_text_bytes_atom, parse_text_chars_atom};
use zerocopy::{
    FromBytes,
    byteorder::{LittleEndian, U16, U32},
};

/// A PPT record containing binary data and metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Record {
    /// Record type
    pub record_type: RecordType,
    /// Original record type value (for unknown types)
    pub record_type_raw: u16,
    /// Record version
    pub version: u16,
    /// Record instance (sub-type)
    pub instance: u16,
    /// Record data length
    pub data_length: u32,
    /// Record data
    pub data: Vec<u8>,
    /// Child records (for container records)
    pub children: Vec<Record>,
}

impl Record {
    /// Parse a PPT record from binary data.
    ///
    /// # Arguments
    ///
    /// * `data` - Binary data containing the record
    /// * `offset` - Starting offset in the data
    ///
    /// # Returns
    ///
    /// Tuple of (`parsed_record`, `bytes_consumed`)
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        Self::parse_with_limits(data, offset, RecordLimits::default())
    }

    /// Parse a PPT record with explicit finite resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(
        data: &[u8],
        offset: usize,
        limits: RecordLimits,
    ) -> Result<(Self, usize)> {
        let mut budget = ParseBudget::new(limits, data.len())?;
        Self::parse_impl(data, offset, false, 0, &mut budget)
    }

    /// Parse a record without truncation recovery or byte resynchronization.
    pub(crate) fn parse_strict(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        Self::parse_strict_with_limits(data, offset, RecordLimits::default())
    }

    pub(crate) fn parse_strict_with_limits(
        data: &[u8],
        offset: usize,
        limits: RecordLimits,
    ) -> Result<(Self, usize)> {
        let mut budget = ParseBudget::new(limits, data.len())?;
        Self::parse_impl(data, offset, true, 0, &mut budget)
    }

    pub(crate) fn parse_with_budget(
        data: &[u8],
        offset: usize,
        strict: bool,
        budget: &mut ParseBudget,
    ) -> Result<(Self, usize)> {
        Self::parse_impl(data, offset, strict, 0, budget)
    }

    fn parse_impl(
        data: &[u8],
        offset: usize,
        strict: bool,
        depth: usize,
        budget: &mut ParseBudget,
    ) -> Result<(Self, usize)> {
        const HEADER_LEN: usize = 8;

        let header_end = offset
            .checked_add(HEADER_LEN)
            .ok_or_else(|| Error::Corrupted("PPT record header offset overflow".to_string()))?;
        let header: &[u8; HEADER_LEN] = data
            .get(offset..header_end)
            .and_then(|bytes| bytes.try_into().ok())
            .ok_or_else(|| Error::Corrupted("Not enough data for PPT record header".to_string()))?;

        // Read record header (8 bytes) - little-endian format
        // PPT Record Header format (based on POI's Record.java):
        // Bytes 0-1: Version and Instance packed together
        // Bytes 2-3: Record Type
        // Bytes 4-7: Data Length

        let [
            version_lo,
            version_hi,
            type_lo,
            type_hi,
            len_0,
            len_1,
            len_2,
            len_3,
        ] = *header;

        // Read version/instance field (bytes 0-1)
        let version_instance = u16::from_le_bytes([version_lo, version_hi]);

        // Read record type (bytes 2-3)
        let record_type = u16::from_le_bytes([type_lo, type_hi]);

        // Read data length (bytes 4-7)
        let data_length = u32::from_le_bytes([len_0, len_1, len_2, len_3]);
        let declared_data_size = usize::try_from(data_length).map_err(|_err| {
            Error::Corrupted(format!(
                "PPT record at offset {offset} has a data length that exceeds this platform"
            ))
        })?;
        let available_data_size = data.len() - header_end;
        if strict && declared_data_size > available_data_size {
            return Err(Error::Corrupted(format!(
                "PPT record at offset {offset} extends beyond its containing data"
            )));
        }
        budget.charge_record(depth, declared_data_size)?;

        // Extract version and instance from the packed field
        // Format: bits 0-3 = version, bits 4-15 = instance (POI's format)
        let version = version_instance & 0x000F; // Low 4 bits for version
        let instance = (version_instance >> 4) & 0x0FFF; // High 12 bits for instance

        let record_type_enum = RecordType::from(record_type);

        // Check if record data extends beyond available data
        if declared_data_size > available_data_size {
            // If this is a container record and we have at least some data, try to parse partially
            if Self::is_container_record(record_type_enum) && available_data_size > 0 {
                // For container records, we can still parse what we have
            } else if available_data_size == 0 {
                return Err(Error::Corrupted(
                    "Record extends beyond data bounds and no data available".to_string(),
                ));
            }
        }

        // Use available data size, but don't exceed what the record claims to need
        let actual_data_size = available_data_size.min(declared_data_size);
        let record_end = header_end.checked_add(actual_data_size).ok_or_else(|| {
            Error::Corrupted(format!("PPT record at offset {offset} size overflow"))
        })?;
        let source = data.get(header_end..record_end).ok_or_else(|| {
            Error::Corrupted(format!(
                "PPT record at offset {offset} extends beyond its containing data"
            ))
        })?;
        budget.charge_copy(source.len())?;
        let mut record_data = Vec::new();
        record_data
            .try_reserve_exact(source.len())
            .map_err(|_err| Error::AllocationFailed("PPT record payload"))?;
        record_data.extend_from_slice(source);

        let mut record = Record {
            record_type: record_type_enum,
            record_type_raw: record_type,
            version,
            instance,
            data_length,
            data: record_data,
            children: Vec::new(),
        };

        // Parse children if this is a container record
        if Self::is_container_record(record_type_enum) && actual_data_size > 0 {
            let children_data = data.get(header_end..record_end).ok_or_else(|| {
                Error::Corrupted(format!(
                    "PPT record at offset {offset} extends beyond its containing data"
                ))
            })?;
            record.children = if strict {
                Self::parse_container_children_strict(children_data, depth, budget)?
            } else {
                Self::parse_container_children(children_data, depth, budget)?
            };
        }

        let consumed = record_end.checked_sub(offset).ok_or_else(|| {
            Error::Corrupted(format!("PPT record at offset {offset} size underflow"))
        })?;
        Ok((record, consumed))
    }

    /// Check if a record type is a container that can hold child records.
    pub(crate) fn is_container_record(record_type: RecordType) -> bool {
        matches!(
            record_type,
            RecordType::Document
                | RecordType::Slide
                | RecordType::Notes
                | RecordType::Handout
                | RecordType::MainMaster
                | RecordType::HeadersFooters
                | RecordType::DocInfoList
                | RecordType::SlideViewInfo
                | RecordType::ExObjList
                | RecordType::VBAInfo
                | RecordType::SlideListWithText
                | RecordType::NormalViewSetInfo9
                | RecordType::NotesTextViewInfo9
                | RecordType::Environment
                | RecordType::FontCollection
                | RecordType::FontCollection10
                | RecordType::BlipCollection9
                | RecordType::Kinsoku
                | RecordType::ExternalHyperlink
                | RecordType::ExternalHyperlink9
                | RecordType::InteractiveInfo
                | RecordType::AnimationInfo
                | RecordType::ProgTags
                | RecordType::ProgStringTag
                | RecordType::ProgBinaryTag
                | RecordType::RoundTripSlideSyncInfo12
                | RecordType::OutlineTextProps9
                | RecordType::OutlineTextProps10
                | RecordType::OutlineTextProps11
                | RecordType::Comment2000
                | RecordType::CommentIndex10
                | RecordType::BuildList
                | RecordType::ChartBuild
                | RecordType::DiagramBuild
                | RecordType::ParaBuild
                | RecordType::ExtTimeNode
                | RecordType::TimeSubEffectContainer
                | RecordType::TimeConditionContainer
                | RecordType::TimeBehaviorContainer
                | RecordType::TimeAnimateBehaviorContainer
                | RecordType::TimeColorBehaviorContainer
                | RecordType::TimeEffectBehaviorContainer
                | RecordType::TimeMotionBehaviorContainer
                | RecordType::TimeRotationBehaviorContainer
                | RecordType::TimeScaleBehaviorContainer
                | RecordType::TimeSetBehaviorContainer
                | RecordType::TimeCommandBehaviorContainer
                | RecordType::TimeClientVisualElement
                | RecordType::TimePropertyList
                | RecordType::TimeVariantList
                | RecordType::TimeAnimationValueList
        )
    }

    /// Parse child records from a container record.
    fn parse_container_children(
        data: &[u8],
        parent_depth: usize,
        budget: &mut ParseBudget,
    ) -> Result<Vec<Record>> {
        let mut children = Vec::new();
        let mut offset = 0;

        while offset + 8 <= data.len() {
            match Self::parse_impl(data, offset, false, parent_depth + 1, budget) {
                Ok((child, consumed)) => {
                    children
                        .try_reserve(1)
                        .map_err(|_err| Error::AllocationFailed("PPT child-record table"))?;
                    children.push(child);
                    offset += consumed;

                    if consumed == 0 {
                        break;
                    }
                },
                Err(error @ (Error::ResourceLimit(_) | Error::AllocationFailed(_))) => {
                    return Err(error);
                },
                Err(_) => {
                    offset += 1;
                    if offset + 8 > data.len() {
                        break;
                    }
                },
            }
        }

        Ok(children)
    }

    fn parse_container_children_strict(
        data: &[u8],
        parent_depth: usize,
        budget: &mut ParseBudget,
    ) -> Result<Vec<Record>> {
        let mut children = Vec::new();
        let mut offset = 0usize;
        while offset < data.len() {
            if data.len() - offset < 8 {
                return Err(Error::Corrupted(
                    "container ends with a truncated record header".to_string(),
                ));
            }
            let (child, consumed) = Self::parse_impl(data, offset, true, parent_depth + 1, budget)?;
            if consumed == 0 {
                return Err(Error::Corrupted(
                    "zero-length progress while parsing a PPT container".to_string(),
                ));
            }
            children
                .try_reserve(1)
                .map_err(|_err| Error::AllocationFailed("PPT child-record table"))?;
            children.push(child);
            offset += consumed;
        }
        Ok(children)
    }

    /// Find a child record of a specific type.
    #[must_use]
    pub fn find_child(&self, record_type: RecordType) -> Option<&Record> {
        self.children
            .iter()
            .find(|child| child.record_type == record_type)
    }

    /// Find all child records of a specific type.
    #[must_use]
    pub fn find_children(&self, record_type: RecordType) -> Vec<&Record> {
        self.children
            .iter()
            .filter(|child| child.record_type == record_type)
            .collect()
    }

    /// Return records stored in every matching `___PPTn` programmable tag.
    ///
    /// `BinaryTagData` is an atom whose payload is itself a strict sequence of
    /// PPT records, so these records do not appear in the ordinary child tree.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn versioned_binary_tag_records(&self, version: u8) -> Result<Vec<Record>> {
        self.versioned_binary_tag_records_with_limits(version, RecordLimits::default())
    }

    /// Return versioned programmable-tag records with explicit parse limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn versioned_binary_tag_records_with_limits(
        &self,
        version: u8,
        limits: RecordLimits,
    ) -> Result<Vec<Record>> {
        if !matches!(version, 9..=12) {
            return Err(Error::Corrupted(
                "Unsupported PowerPoint programmable-tag version".to_string(),
            ));
        }
        let expected_tag_name = format!("___PPT{version}");
        let expected_name: Vec<u16> = expected_tag_name.encode_utf16().collect();
        let mut session = RecordParseSession::new(limits, self.data.len())?;
        let prog_tags = collect_prog_tags(self, &mut session)?;
        let mut records = Vec::new();

        for (container, depth) in prog_tags {
            let tag_depth = checked_logical_depth(depth, 1)?;
            let tag_child_depth = checked_logical_depth(depth, 2)?;
            let blob_depth = checked_logical_depth(depth, 3)?;
            for tag in session.parse_sequence(
                &container.data,
                &format!("PPT{version} ProgTags"),
                tag_depth,
            )? {
                if tag.record_type != RecordType::ProgBinaryTag {
                    continue;
                }
                let children = session.parse_sequence(
                    &tag.data,
                    &format!("PPT{version} ProgBinaryTag"),
                    tag_child_depth,
                )?;
                let Some(name) = children
                    .iter()
                    .find(|child| child.record_type == RecordType::CString)
                else {
                    continue;
                };
                if name.version != 0
                    || name.instance != 0
                    || name.data.len() != expected_name.len() * 2
                    || !name
                        .data
                        .chunks_exact(2)
                        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                        .eq(expected_name.iter().copied())
                {
                    continue;
                }
                let blob = children
                    .iter()
                    .find(|child| child.record_type == RecordType::BinaryTagData)
                    .ok_or_else(|| {
                        Error::Corrupted(format!(
                            "___PPT{version} programmable tag is missing BinaryTagData"
                        ))
                    })?;
                let decoded = session.parse_sequence(
                    &blob.data,
                    &format!("___PPT{version} BinaryTagData"),
                    blob_depth,
                )?;
                records
                    .try_reserve(decoded.len())
                    .map_err(|_err| Error::AllocationFailed("versioned binary-tag records"))?;
                records.extend(decoded);
            }
        }
        Ok(records)
    }

    pub(crate) fn parse_sequence_strict(data: &[u8], context: &str) -> Result<Vec<Record>> {
        Self::parse_sequence_strict_with_limits(data, context, RecordLimits::default())
    }

    pub(crate) fn parse_sequence_strict_with_limits(
        data: &[u8],
        context: &str,
        limits: RecordLimits,
    ) -> Result<Vec<Record>> {
        let mut session = RecordParseSession::new(limits, data.len())?;
        session.parse_sequence(data, context, 0)
    }

    fn parse_sequence_strict_with_budget(
        data: &[u8],
        context: &str,
        depth: usize,
        budget: &mut ParseBudget,
    ) -> Result<Vec<Record>> {
        let mut records = Vec::new();
        let mut offset = 0usize;
        while offset < data.len() {
            let header_end = offset.checked_add(8).ok_or_else(|| {
                Error::Corrupted(format!("{context} record header offset overflow"))
            })?;
            if header_end > data.len() {
                return Err(Error::Corrupted(format!(
                    "Truncated record header in {context}"
                )));
            }
            let declared_length = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]);
            let length = usize::try_from(declared_length)
                .map_err(|_err| Error::Corrupted(format!("{context} record size overflow")))?;
            let record_end = header_end
                .checked_add(length)
                .ok_or_else(|| Error::Corrupted(format!("{context} record size overflow")))?;
            if record_end > data.len() {
                return Err(Error::Corrupted(format!("Record extends beyond {context}")));
            }
            let (record, consumed) = Self::parse_impl(data, offset, true, depth, budget)?;
            if consumed != record_end - offset {
                return Err(Error::Corrupted(format!(
                    "Record in {context} was only partially parsed"
                )));
            }
            records
                .try_reserve(1)
                .map_err(|_err| Error::AllocationFailed("PPT strict record sequence"))?;
            records.push(record);
            offset = record_end;
        }
        Ok(records)
    }

    /// Extract slide data from this record.
    #[must_use]
    pub fn extract_slide_data(&self) -> Option<Vec<u8>> {
        if let Some(ppdrawing) = self.find_child(RecordType::PPDrawing) {
            return Some(ppdrawing.data.clone());
        }

        if self.record_type == RecordType::Slide && !self.data.is_empty() && self.data.len() > 8 {
            let first_record_type =
                U16::<LittleEndian>::read_from_bytes(&self.data[0..2]).map_or(0, U16::get);
            if first_record_type >= 0xF000 {
                return Some(self.data.clone());
            }
        }

        None
    }

    /// Extract document information from this record.
    #[must_use]
    pub fn extract_document_info(&self) -> Option<DocumentInfo> {
        if self.record_type != RecordType::Document {
            return None;
        }

        let mut info = DocumentInfo::default();

        if let Some(document_atom) = self.find_child(RecordType::DocumentAtom) {
            info = Self::parse_document_atom(document_atom);
        }

        if self.find_child(RecordType::Environment).is_some() {
            info.has_environment = true;
        }

        if self.find_child(RecordType::PPDrawingGroup).is_some() {
            info.has_drawing_group = true;
        }

        Some(info)
    }

    /// Parse `DocumentAtom` record data.
    fn parse_document_atom(record: &Record) -> DocumentInfo {
        let mut info = DocumentInfo::default();

        if record.data.len() >= 20 {
            info.slide_width =
                U32::<LittleEndian>::read_from_bytes(&record.data[0..4]).map_or(0, U32::get);
            info.slide_height =
                U32::<LittleEndian>::read_from_bytes(&record.data[4..8]).map_or(0, U32::get);
            info.slide_count = U32::<LittleEndian>::read_from_bytes(&record.data[8..12])
                .map_or(0, |v| v.get() as usize);
            info.notes_count = U32::<LittleEndian>::read_from_bytes(&record.data[12..16])
                .map_or(0, |v| v.get() as usize);
            info.master_count = U32::<LittleEndian>::read_from_bytes(&record.data[16..20])
                .map_or(0, |v| v.get() as usize);
        }
        if record.data.len() >= 28 {
            info.notes_master_persist_id_ref =
                U32::<LittleEndian>::read_from_bytes(&record.data[24..28]).map_or(0, U32::get);
        }
        if record.data.len() >= 32 {
            info.handout_master_persist_id_ref =
                U32::<LittleEndian>::read_from_bytes(&record.data[28..32]).map_or(0, U32::get);
        }

        info
    }

    /// Extract slide information from this record.
    #[must_use]
    pub fn extract_slide_info(&self) -> Option<SlideInfo> {
        if self.record_type != RecordType::Slide {
            return None;
        }

        let mut info = SlideInfo::default();

        if let Some(slide_atom) = self.find_child(RecordType::SlideAtom) {
            info = Self::parse_slide_atom(slide_atom);
        }

        if self.find_child(RecordType::PPDrawing).is_some() {
            info.has_drawing = true;
        }

        info.has_notes = info.notes_id != 0;

        Some(info)
    }

    /// Parse `SlideAtom` record data.
    fn parse_slide_atom(record: &Record) -> SlideInfo {
        let mut info = SlideInfo::default();

        if record.data.len() >= 20 {
            info.layout_id =
                U32::<LittleEndian>::read_from_bytes(&record.data[0..4]).map_or(0, U32::get);
            info.master_id =
                U32::<LittleEndian>::read_from_bytes(&record.data[12..16]).map_or(0, U32::get);
            info.notes_id =
                U32::<LittleEndian>::read_from_bytes(&record.data[16..20]).map_or(0, U32::get);
        }

        info
    }

    /// Extract text content from this record and its children.
    ///
    /// # Errors
    ///
    /// Currently this function never fails: malformed atom payloads are skipped
    /// rather than reported. The `Result` return type is kept so future parse
    /// failures can surface without breaking the public API.
    pub fn extract_text(&self) -> Result<String> {
        if matches!(
            self.record_type,
            RecordType::ProgTags
                | RecordType::ProgStringTag
                | RecordType::ProgBinaryTag
                | RecordType::Comment2000
        ) {
            return Ok(String::new());
        }
        let mut text_parts = Vec::new();

        // Extract text from text-related records
        if self.record_type == RecordType::TextCharsAtom {
            if let Ok(text) = parse_text_chars_atom(&self.data) {
                text_parts.push(text);
            }
        } else if self.record_type == RecordType::TextBytesAtom {
            if let Ok(text) = parse_text_bytes_atom(&self.data) {
                text_parts.push(text);
            }
        } else if self.record_type == RecordType::CString
            && let Ok(text) = parse_cstring(&self.data)
        {
            text_parts.push(text);
        }

        // Recursively extract text from children
        for child in &self.children {
            if let Ok(child_text) = child.extract_text()
                && !child_text.is_empty()
            {
                text_parts.push(child_text);
            }
        }

        Ok(text_parts.join("\n"))
    }

    /// Extract `SlideListWithText` records from Document record.
    #[must_use]
    pub fn extract_slide_list_with_texts(&self) -> Vec<&Record> {
        if self.record_type != RecordType::Document {
            return Vec::new();
        }

        self.children
            .iter()
            .filter(|child| child.record_type == RecordType::SlideListWithText)
            .collect()
    }

    /// Get the instance field from the record header.
    #[must_use]
    pub fn get_instance(&self) -> u16 {
        self.instance
    }

    /// Group children into `SlideAtomsSets`.
    #[must_use]
    pub fn group_into_slide_atoms_sets(&self) -> Vec<SlideAtomsSet<'_>> {
        if self.record_type != RecordType::SlideListWithText {
            return Vec::new();
        }

        let mut sets = Vec::new();
        let mut i = 0;

        while i < self.children.len() {
            if self.children[i].record_type == RecordType::SlidePersistAtom {
                let slide_persist_atom = &self.children[i];

                let mut end_pos = i + 1;
                while end_pos < self.children.len()
                    && self.children[end_pos].record_type != RecordType::SlidePersistAtom
                {
                    end_pos += 1;
                }

                let associated_records: Vec<&Record> =
                    self.children[i + 1..end_pos].iter().collect();

                sets.push(SlideAtomsSet {
                    slide_persist_atom,
                    slide_records: associated_records,
                });

                i = end_pos;
            } else {
                i += 1;
            }
        }

        sets
    }

    /// Get the slide ID from a `SlidePersistAtom` record.
    #[must_use]
    pub fn get_slide_id(&self) -> Option<u32> {
        if self.record_type == RecordType::SlidePersistAtom && self.data.len() >= 4 {
            Some(U32::<LittleEndian>::read_from_bytes(&self.data[0..4]).map_or(0, U32::get))
        } else {
            None
        }
    }
}

pub(crate) struct ParseBudget {
    limits: RecordLimits,
    records: usize,
    copied_payload_bytes: usize,
}

/// Shared budget for semantic record sequences embedded in already-parsed records.
///
/// A session must be reused across every nested byte slice so record, depth,
/// and copied-payload ceilings cannot be multiplied by restarting a parser.
pub(crate) struct RecordParseSession {
    budget: ParseBudget,
}

impl RecordParseSession {
    pub(crate) fn new(limits: RecordLimits, input_bytes: usize) -> Result<Self> {
        Ok(Self {
            budget: ParseBudget::new(limits, input_bytes)?,
        })
    }

    pub(crate) fn parse_sequence(
        &mut self,
        data: &[u8],
        context: &str,
        logical_depth: usize,
    ) -> Result<Vec<Record>> {
        Record::parse_sequence_strict_with_budget(data, context, logical_depth, &mut self.budget)
    }

    /// Charge traversal of an existing record without charging a payload copy.
    pub(crate) fn account_existing(&mut self, logical_depth: usize) -> Result<()> {
        self.budget.charge_visit(logical_depth)
    }

    /// Charge a record that a range-aware resolver validates from its header
    /// without copying its payload.
    pub(crate) fn account_existing_header(
        &mut self,
        logical_depth: usize,
        payload_bytes: usize,
    ) -> Result<()> {
        self.budget.charge_record(logical_depth, payload_bytes)
    }

    /// Charge a record whose complete payload was materialized without using
    /// the semantic record parser.
    pub(crate) fn account_materialized_record(
        &mut self,
        logical_depth: usize,
        payload_bytes: usize,
    ) -> Result<()> {
        self.budget.charge_record(logical_depth, payload_bytes)?;
        self.budget.charge_copy(payload_bytes)
    }

    /// Parse one complete strict record while retaining this session's shared
    /// record/depth/copy budget.
    pub(crate) fn parse_strict_record(
        &mut self,
        data: &[u8],
        offset: usize,
    ) -> Result<(Record, usize)> {
        self.parse_strict_record_at_depth(data, offset, 0)
    }

    /// Parse one strict record at an already-known logical nesting depth.
    pub(crate) fn parse_strict_record_at_depth(
        &mut self,
        data: &[u8],
        offset: usize,
        logical_depth: usize,
    ) -> Result<(Record, usize)> {
        Record::parse_impl(data, offset, true, logical_depth, &mut self.budget)
    }
}

impl ParseBudget {
    pub(crate) fn new(limits: RecordLimits, input_bytes: usize) -> Result<Self> {
        if input_bytes > limits.max_input_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT input size {input_bytes} exceeds limit {}",
                limits.max_input_bytes
            )));
        }
        Ok(Self {
            limits,
            records: 0,
            copied_payload_bytes: 0,
        })
    }

    fn charge_record(&mut self, depth: usize, payload_bytes: usize) -> Result<()> {
        if depth > self.limits.max_depth {
            return Err(Error::ResourceLimit(format!(
                "PPT record nesting depth {depth} exceeds limit {}",
                self.limits.max_depth
            )));
        }
        if payload_bytes > self.limits.max_record_payload_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT record payload size {payload_bytes} exceeds limit {}",
                self.limits.max_record_payload_bytes
            )));
        }
        let encoded_bytes = payload_bytes
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("PPT encoded record size overflow".to_string()))?;
        if encoded_bytes > self.limits.max_record_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT record size {encoded_bytes} exceeds limit {}",
                self.limits.max_record_bytes
            )));
        }
        self.charge_visit(depth)
    }

    fn charge_visit(&mut self, depth: usize) -> Result<()> {
        if depth > self.limits.max_depth {
            return Err(Error::ResourceLimit(format!(
                "PPT record nesting depth {depth} exceeds limit {}",
                self.limits.max_depth
            )));
        }
        self.records = self
            .records
            .checked_add(1)
            .ok_or_else(|| Error::Corrupted("PPT record count overflow".to_string()))?;
        if self.records > self.limits.max_records {
            return Err(Error::ResourceLimit(format!(
                "PPT record count exceeds limit {}",
                self.limits.max_records
            )));
        }
        Ok(())
    }

    fn charge_copy(&mut self, bytes: usize) -> Result<()> {
        self.copied_payload_bytes = self
            .copied_payload_bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::Corrupted("PPT copied payload size overflow".to_string()))?;
        if self.copied_payload_bytes > self.limits.max_copied_payload_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT copied payload bytes exceed limit {}",
                self.limits.max_copied_payload_bytes
            )));
        }
        Ok(())
    }
}

fn collect_prog_tags<'a>(
    record: &'a Record,
    session: &mut RecordParseSession,
) -> Result<Vec<(&'a Record, usize)>> {
    let mut output = Vec::new();
    let mut pending = Vec::new();
    pending
        .try_reserve(1)
        .map_err(|_err| Error::AllocationFailed("programmable-tag traversal stack"))?;
    pending.push((record, 0usize));
    while let Some((current, depth)) = pending.pop() {
        session.account_existing(depth)?;
        if current.record_type == RecordType::ProgTags {
            output
                .try_reserve(1)
                .map_err(|_err| Error::AllocationFailed("programmable-tag table"))?;
            output.push((current, depth));
            continue;
        }
        pending
            .try_reserve(current.children.len())
            .map_err(|_err| Error::AllocationFailed("programmable-tag traversal stack"))?;
        let child_depth = checked_logical_depth(depth, 1)?;
        pending.extend(
            current
                .children
                .iter()
                .rev()
                .map(|child| (child, child_depth)),
        );
    }
    Ok(output)
}

fn checked_logical_depth(depth: usize, increment: usize) -> Result<usize> {
    depth
        .checked_add(increment)
        .ok_or_else(|| Error::ResourceLimit("PPT logical record depth overflow".to_string()))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn atom(record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + payload.len());
        bytes.extend_from_slice(&0x1234_u16.to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn canonical_record(record_type: RecordType, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + payload.len());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&record_type.as_u16().to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn versioned_tag_root(tag_count: usize) -> Record {
        let mut name_bytes = Vec::new();
        for code_unit in "___PPT10".encode_utf16() {
            name_bytes.extend_from_slice(&code_unit.to_le_bytes());
        }
        let name = canonical_record(RecordType::CString, &name_bytes);
        let leaf = canonical_record(RecordType::Unknown, &[0u8; 16]);
        let blob = canonical_record(RecordType::BinaryTagData, &leaf);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = canonical_record(RecordType::ProgBinaryTag, &tag_payload);
        let mut tag_bytes = Vec::new();
        for _ in 0..tag_count {
            tag_bytes.extend_from_slice(&tag);
        }
        let tags = canonical_record(RecordType::ProgTags, &tag_bytes);
        let root = canonical_record(RecordType::Document, &tags);
        Record::parse(&root, 0).unwrap().0
    }

    #[test]
    fn test_record_creation() {
        let record = Record {
            record_type: RecordType::Document,
            record_type_raw: 1000,
            version: 1,
            instance: 0,
            data_length: 16,
            data: vec![1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
            children: Vec::new(),
        };

        assert_eq!(record.record_type, RecordType::Document);
        assert_eq!(record.version, 1);
        assert_eq!(record.data_length, 16);
        assert_eq!(record.data.len(), 16);
    }

    #[test]
    fn parse_rejects_header_offset_overflow() {
        let error = Record::parse(&[0; 8], usize::MAX).unwrap_err();

        assert!(matches!(
            error,
            Error::Corrupted(message) if message.contains("header offset overflow")
        ));
    }

    #[test]
    fn parse_rejects_header_offset_just_below_overflow() {
        let error = Record::parse(&[0; 8], usize::MAX - 3).unwrap_err();

        assert!(matches!(
            error,
            Error::Corrupted(message) if message.contains("header offset overflow")
        ));
    }

    #[test]
    fn parse_rejects_offset_past_input() {
        let error = Record::parse(&[0; 8], 9).unwrap_err();

        assert!(matches!(error, Error::Corrupted(_)));
    }

    #[test]
    fn parse_handles_a_checked_nonzero_offset() {
        let payload = [0xAA, 0xBB, 0xCC];
        let mut bytes = vec![0xFF; 5];
        bytes.extend_from_slice(&atom(0x2222, &payload));

        let (record, consumed) = Record::parse(&bytes, 5).unwrap();

        assert_eq!(consumed, 8 + payload.len());
        assert_eq!(record.record_type_raw, 0x2222);
        assert_eq!(record.version, 4);
        assert_eq!(record.instance, 0x123);
        assert_eq!(record.data, payload);
    }

    #[test]
    fn strict_parse_rejects_an_oversized_payload_without_slicing() {
        let mut bytes = atom(0x2222, &[]);
        bytes[4..8].copy_from_slice(&u32::MAX.to_le_bytes());

        let error = Record::parse_strict(&bytes, 0).unwrap_err();

        assert!(matches!(
            error,
            Error::Corrupted(message)
                if message.contains("extends beyond its containing data")
        ));
    }

    #[test]
    fn explicit_payload_and_record_limits_accept_exact_boundaries() {
        let bytes = atom(0x2222, &[1, 2, 3, 4]);
        let limits = RecordLimits {
            max_package_bytes: bytes.len(),
            max_input_bytes: bytes.len(),
            max_aggregate_input_bytes: bytes.len(),
            max_record_bytes: bytes.len(),
            max_record_payload_bytes: 4,
            max_copied_payload_bytes: 4,
            max_records: 1,
            max_depth: 0,
        };

        let (record, consumed) = Record::parse_with_limits(&bytes, 0, limits).unwrap();
        assert_eq!(record.data, [1, 2, 3, 4]);
        assert_eq!(consumed, bytes.len());
    }

    #[test]
    fn explicit_limits_reject_input_payload_and_copied_bytes_over_boundary() {
        let bytes = atom(0x2222, &[1, 2, 3, 4]);
        for limits in [
            RecordLimits {
                max_input_bytes: bytes.len() - 1,
                ..RecordLimits::default()
            },
            RecordLimits {
                max_record_payload_bytes: 3,
                ..RecordLimits::default()
            },
            RecordLimits {
                max_copied_payload_bytes: 3,
                ..RecordLimits::default()
            },
        ] {
            assert!(Record::parse_with_limits(&bytes, 0, limits).is_err());
        }
    }

    #[test]
    fn explicit_limits_bound_depth_record_count_and_aggregate_container_copies() {
        let leaf = atom(0x2222, &[1]);
        let container = atom(RecordType::Document.as_u16(), &leaf);

        assert!(
            Record::parse_with_limits(
                &container,
                0,
                RecordLimits {
                    max_depth: 0,
                    ..RecordLimits::default()
                }
            )
            .is_err()
        );
        assert!(
            Record::parse_with_limits(
                &container,
                0,
                RecordLimits {
                    max_records: 1,
                    ..RecordLimits::default()
                }
            )
            .is_err()
        );
        assert!(
            Record::parse_with_limits(
                &container,
                0,
                RecordLimits {
                    max_copied_payload_bytes: leaf.len(),
                    ..RecordLimits::default()
                }
            )
            .is_err()
        );
    }

    #[test]
    fn semantic_session_shares_count_copy_and_logical_depth_across_sequences() {
        let bytes = atom(0x2222, &[1, 2, 3, 4]);
        let mut session = RecordParseSession::new(
            RecordLimits {
                max_records: 2,
                max_copied_payload_bytes: 7,
                max_depth: 2,
                ..RecordLimits::default()
            },
            bytes.len(),
        )
        .unwrap();

        session.parse_sequence(&bytes, "first", 1).unwrap();
        let error = session.parse_sequence(&bytes, "second", 2).unwrap_err();
        assert!(
            matches!(error, Error::ResourceLimit(message) if message.contains("copied payload"))
        );

        let mut count_session = RecordParseSession::new(
            RecordLimits {
                max_records: 1,
                ..RecordLimits::default()
            },
            0,
        )
        .unwrap();
        count_session.account_existing(0).unwrap();
        assert!(matches!(
            count_session.account_existing(0),
            Err(Error::ResourceLimit(_))
        ));

        let mut depth_session = RecordParseSession::new(
            RecordLimits {
                max_depth: 1,
                ..RecordLimits::default()
            },
            bytes.len(),
        )
        .unwrap();
        assert!(matches!(
            depth_session.parse_sequence(&bytes, "deep", 2),
            Err(Error::ResourceLimit(_))
        ));
    }

    #[test]
    fn header_only_accounting_does_not_consume_copy_budget() {
        let bytes = atom(0x2222, &[1, 2, 3, 4]);
        let mut header_session = RecordParseSession::new(
            RecordLimits {
                max_copied_payload_bytes: 0,
                max_record_payload_bytes: 4,
                max_record_bytes: bytes.len(),
                max_records: 1,
                ..RecordLimits::default()
            },
            0,
        )
        .unwrap();
        header_session.account_existing_header(0, 4).unwrap();

        let mut exact_session = RecordParseSession::new(
            RecordLimits {
                max_copied_payload_bytes: 4,
                max_record_payload_bytes: 4,
                max_record_bytes: bytes.len(),
                max_records: 1,
                ..RecordLimits::default()
            },
            bytes.len(),
        )
        .unwrap();
        exact_session.parse_strict_record(&bytes, 0).unwrap();

        let mut one_under_session = RecordParseSession::new(
            RecordLimits {
                max_copied_payload_bytes: 3,
                max_record_payload_bytes: 4,
                max_record_bytes: bytes.len(),
                max_records: 1,
                ..RecordLimits::default()
            },
            bytes.len(),
        )
        .unwrap();
        assert!(matches!(
            one_under_session.parse_strict_record(&bytes, 0),
            Err(Error::ResourceLimit(message)) if message.contains("copied payload")
        ));
    }

    #[test]
    fn versioned_tags_share_record_count_across_many_nested_payloads() {
        let one = versioned_tag_root(1);
        assert_eq!(
            one.versioned_binary_tag_records_with_limits(
                10,
                RecordLimits {
                    max_records: 10,
                    max_depth: 8,
                    ..RecordLimits::default()
                },
            )
            .unwrap()
            .len(),
            1
        );
        let many = versioned_tag_root(2);
        assert!(matches!(
            many.versioned_binary_tag_records_with_limits(
                10,
                RecordLimits {
                    max_records: 10,
                    max_depth: 8,
                    ..RecordLimits::default()
                },
            ),
            Err(Error::ResourceLimit(message)) if message.contains("record count")
        ));
    }

    #[test]
    fn versioned_tags_share_copied_payload_budget_across_nested_payloads() {
        let one = versioned_tag_root(1);
        one.versioned_binary_tag_records_with_limits(
            10,
            RecordLimits {
                max_records: 100,
                max_depth: 8,
                max_copied_payload_bytes: 200,
                ..RecordLimits::default()
            },
        )
        .unwrap();
        let many = versioned_tag_root(2);
        assert!(matches!(
            many.versioned_binary_tag_records_with_limits(
                10,
                RecordLimits {
                    max_records: 100,
                    max_depth: 8,
                    max_copied_payload_bytes: 200,
                    ..RecordLimits::default()
                },
            ),
            Err(Error::ResourceLimit(message)) if message.contains("copied payload")
        ));
    }
}
