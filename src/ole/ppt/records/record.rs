//! Core PPT record structure and parsing.
//!
//! This module implements the fundamental PPT record parsing based on
//! Apache POI's HSLF Record.java implementation.

use super::{DocumentInfo, SlideAtomsSet, SlideInfo};
use crate::ole::consts::PptRecordType;
use crate::ole::ppt::package::{PptError, Result};
use crate::ole::ppt::text::extractor::{
    parse_cstring, parse_text_bytes_atom, parse_text_chars_atom,
};
use zerocopy::{
    FromBytes,
    byteorder::{LittleEndian, U16, U32},
};

/// A PPT record containing binary data and metadata.
#[derive(Debug, Clone)]
pub struct PptRecord {
    /// Record type
    pub record_type: PptRecordType,
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
    pub children: Vec<PptRecord>,
}

impl PptRecord {
    /// Parse a PPT record from binary data.
    ///
    /// # Arguments
    ///
    /// * `data` - Binary data containing the record
    /// * `offset` - Starting offset in the data
    ///
    /// # Returns
    ///
    /// Tuple of (parsed_record, bytes_consumed)
    pub fn parse(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        if offset + 8 > data.len() {
            return Err(PptError::Corrupted(
                "Not enough data for PPT record header".to_string(),
            ));
        }

        // Read record header (8 bytes) - little-endian format
        // PPT Record Header format (based on POI's Record.java):
        // Bytes 0-1: Version and Instance packed together
        // Bytes 2-3: Record Type
        // Bytes 4-7: Data Length

        // Read version/instance field (bytes 0-1)
        let version_instance = U16::<LittleEndian>::read_from_bytes(&data[offset..offset + 2])
            .map(|v| v.get())
            .unwrap_or(0);

        // Read record type (bytes 2-3)
        let record_type = U16::<LittleEndian>::read_from_bytes(&data[offset + 2..offset + 4])
            .map(|v| v.get())
            .unwrap_or(0);

        // Read data length (bytes 4-7)
        let data_length = U32::<LittleEndian>::read_from_bytes(&data[offset + 4..offset + 8])
            .map(|v| v.get())
            .unwrap_or(0);

        // Extract version and instance from the packed field
        // Format: bits 0-3 = version, bits 4-15 = instance (POI's format)
        let version = version_instance & 0x000F; // Low 4 bits for version
        let instance = (version_instance >> 4) & 0x0FFF; // High 12 bits for instance

        let record_type_enum = PptRecordType::from(record_type);

        // Check if record data extends beyond available data
        let available_data_size = data.len().saturating_sub(offset + 8);
        if data_length as usize > available_data_size {
            // If this is a container record and we have at least some data, try to parse partially
            if Self::is_container_record(record_type_enum) && available_data_size > 0 {
                // For container records, we can still parse what we have
            } else if available_data_size == 0 {
                return Err(PptError::Corrupted(
                    "Record extends beyond data bounds and no data available".to_string(),
                ));
            }
        }

        // Use available data size, but don't exceed what the record claims to need
        let actual_data_size = available_data_size.min(data_length as usize);
        let record_data = data[offset + 8..offset + 8 + actual_data_size].to_vec();

        let mut record = PptRecord {
            record_type: record_type_enum,
            record_type_raw: record_type,
            version,
            instance,
            data_length: actual_data_size as u32,
            data: record_data,
            children: Vec::new(),
        };

        // Parse children if this is a container record
        if Self::is_container_record(record_type_enum) && actual_data_size > 0 {
            record.children =
                Self::parse_container_children(&data[offset + 8..offset + 8 + actual_data_size])?;
        }

        Ok((record, 8 + actual_data_size))
    }

    /// Check if a record type is a container that can hold child records.
    fn is_container_record(record_type: PptRecordType) -> bool {
        matches!(
            record_type,
            PptRecordType::Document
                | PptRecordType::Slide
                | PptRecordType::Notes
                | PptRecordType::MainMaster
                | PptRecordType::HeadersFooters
                | PptRecordType::ExObjList
                | PptRecordType::VBAInfo
                | PptRecordType::SlideListWithText
                | PptRecordType::PersistPtrHolder
                | PptRecordType::Environment
                | PptRecordType::InteractiveInfo
                | PptRecordType::AnimationInfo
        )
    }

    /// Parse child records from a container record.
    fn parse_container_children(data: &[u8]) -> Result<Vec<PptRecord>> {
        let mut children = Vec::new();
        let mut offset = 0;

        while offset + 8 <= data.len() {
            match Self::parse(data, offset) {
                Ok((child, consumed)) => {
                    children.push(child);
                    offset += consumed;

                    if consumed == 0 {
                        break;
                    }
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

    /// Find a child record of a specific type.
    pub fn find_child(&self, record_type: PptRecordType) -> Option<&PptRecord> {
        self.children
            .iter()
            .find(|child| child.record_type == record_type)
    }

    /// Find all child records of a specific type.
    pub fn find_children(&self, record_type: PptRecordType) -> Vec<&PptRecord> {
        self.children
            .iter()
            .filter(|child| child.record_type == record_type)
            .collect()
    }

    /// Extract slide data from this record.
    pub fn extract_slide_data(&self) -> Option<Vec<u8>> {
        if let Some(ppdrawing) = self.find_child(PptRecordType::PPDrawing) {
            return Some(ppdrawing.data.clone());
        }

        if self.record_type == PptRecordType::Slide && !self.data.is_empty() && self.data.len() > 8
        {
            let first_record_type = U16::<LittleEndian>::read_from_bytes(&self.data[0..2])
                .map(|v| v.get())
                .unwrap_or(0);
            if first_record_type >= 0xF000 {
                return Some(self.data.clone());
            }
        }

        None
    }

    /// Extract document information from this record.
    pub fn extract_document_info(&self) -> Option<DocumentInfo> {
        if self.record_type != PptRecordType::Document {
            return None;
        }

        let mut info = DocumentInfo::default();

        if let Some(document_atom) = self.find_child(PptRecordType::DocumentAtom) {
            info = Self::parse_document_atom(document_atom);
        }

        if self.find_child(PptRecordType::Environment).is_some() {
            info.has_environment = true;
        }

        if self.find_child(PptRecordType::PPDrawingGroup).is_some() {
            info.has_drawing_group = true;
        }

        Some(info)
    }

    /// Parse DocumentAtom record data.
    fn parse_document_atom(record: &PptRecord) -> DocumentInfo {
        let mut info = DocumentInfo::default();

        if record.data.len() >= 20 {
            info.slide_width = U32::<LittleEndian>::read_from_bytes(&record.data[0..4])
                .map(|v| v.get())
                .unwrap_or(0);
            info.slide_height = U32::<LittleEndian>::read_from_bytes(&record.data[4..8])
                .map(|v| v.get())
                .unwrap_or(0);
            info.slide_count = U32::<LittleEndian>::read_from_bytes(&record.data[8..12])
                .map(|v| v.get() as usize)
                .unwrap_or(0);
            info.notes_count = U32::<LittleEndian>::read_from_bytes(&record.data[12..16])
                .map(|v| v.get() as usize)
                .unwrap_or(0);
            info.master_count = U32::<LittleEndian>::read_from_bytes(&record.data[16..20])
                .map(|v| v.get() as usize)
                .unwrap_or(0);
        }

        info
    }

    /// Extract slide information from this record.
    pub fn extract_slide_info(&self) -> Option<SlideInfo> {
        if self.record_type != PptRecordType::Slide {
            return None;
        }

        let mut info = SlideInfo::default();

        if let Some(slide_atom) = self.find_child(PptRecordType::SlideAtom) {
            info = Self::parse_slide_atom(slide_atom);
        }

        if self.find_child(PptRecordType::PPDrawing).is_some() {
            info.has_drawing = true;
        }

        if self.find_child(PptRecordType::Notes).is_some() {
            info.has_notes = true;
        }

        Some(info)
    }

    /// Parse SlideAtom record data.
    fn parse_slide_atom(record: &PptRecord) -> SlideInfo {
        let mut info = SlideInfo::default();

        if record.data.len() >= 12 {
            info.layout_id = U32::<LittleEndian>::read_from_bytes(&record.data[0..4])
                .map(|v| v.get())
                .unwrap_or(0);
            info.master_id = U32::<LittleEndian>::read_from_bytes(&record.data[4..8])
                .map(|v| v.get())
                .unwrap_or(0);
            info.notes_id = U32::<LittleEndian>::read_from_bytes(&record.data[8..12])
                .map(|v| v.get())
                .unwrap_or(0);
        }

        info
    }

    /// Extract text content from this record and its children.
    pub fn extract_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        // Extract text from text-related records
        match self.record_type {
            PptRecordType::TextCharsAtom => {
                if let Ok(text) = parse_text_chars_atom(&self.data) {
                    text_parts.push(text);
                }
            },
            PptRecordType::TextBytesAtom => {
                if let Ok(text) = parse_text_bytes_atom(&self.data) {
                    text_parts.push(text);
                }
            },
            PptRecordType::CString => {
                if let Ok(text) = parse_cstring(&self.data) {
                    text_parts.push(text);
                }
            },
            _ => {},
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

    /// Extract SlideListWithText records from Document record.
    pub fn extract_slide_list_with_texts(&self) -> Vec<&PptRecord> {
        if self.record_type != PptRecordType::Document {
            return Vec::new();
        }

        self.children
            .iter()
            .filter(|child| child.record_type == PptRecordType::SlideListWithText)
            .collect()
    }

    /// Get the instance field from the record header.
    pub fn get_instance(&self) -> u16 {
        self.instance
    }

    /// Group children into SlideAtomsSets.
    pub fn group_into_slide_atoms_sets<'a>(&'a self) -> Vec<SlideAtomsSet<'a>> {
        if self.record_type != PptRecordType::SlideListWithText {
            return Vec::new();
        }

        let mut sets = Vec::new();
        let mut i = 0;

        while i < self.children.len() {
            if self.children[i].record_type == PptRecordType::SlidePersistAtom {
                let slide_persist_atom = &self.children[i];

                let mut end_pos = i + 1;
                while end_pos < self.children.len()
                    && self.children[end_pos].record_type != PptRecordType::SlidePersistAtom
                {
                    end_pos += 1;
                }

                let associated_records: Vec<&PptRecord> =
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

    /// Get the slide ID from a SlidePersistAtom record.
    pub fn get_slide_id(&self) -> Option<u32> {
        if self.record_type == PptRecordType::SlidePersistAtom && self.data.len() >= 4 {
            Some(
                U32::<LittleEndian>::read_from_bytes(&self.data[0..4])
                    .map(|v| v.get())
                    .unwrap_or(0),
            )
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_record_creation() {
        let record = PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 1,
            instance: 0,
            data_length: 16,
            data: vec![1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
            children: Vec::new(),
        };

        assert_eq!(record.record_type, PptRecordType::Document);
        assert_eq!(record.version, 1);
        assert_eq!(record.data_length, 16);
        assert_eq!(record.data.len(), 16);
    }
}
