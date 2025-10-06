/// PPT record parser for binary format parsing.
///
/// This module provides functionality to parse PowerPoint binary records
/// according to the MS-PPT specification, following POI's HSLF record parsing logic.
use super::super::consts::PptRecordType;
use super::package::{PptError, Result};

/// A PPT record containing binary data and metadata.
#[derive(Debug, Clone)]
pub struct PptRecord {
    /// Record type
    pub record_type: PptRecordType,
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

/// Actions for processing UTF-16LE text characters.
#[derive(Debug, Clone, Copy)]
enum TextCharAction {
    /// Add the character to the result
    Add(char),
    /// Stop processing (null terminator found)
    Stop,
    /// Skip this character (invalid)
    Skip,
}

impl TextCharAction {
    /// Process a UTF-16LE code unit and determine the appropriate action.
    fn process_utf16_char(code_unit: u16) -> Self {
        match code_unit {
            // Null terminator - stop processing
            0 => TextCharAction::Stop,
            // ASCII range (0x00-0x7F) - add as character
            0x01..=0x7F => {
                if let Some(ch) = char::from_u32(code_unit as u32) {
                    TextCharAction::Add(ch)
                } else {
                    TextCharAction::Skip
                }
            }
            // Unicode range (0x80 and above) - try to decode as Unicode
            0x80.. => {
                if let Some(ch) = char::from_u32(code_unit as u32) {
                    TextCharAction::Add(ch)
                } else {
                    TextCharAction::Skip
                }
            }
        }
    }
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
            return Err(PptError::Corrupted("Not enough data for PPT record header".to_string()));
        }

        // Read record header (8 bytes) - little-endian format
        let record_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let data_length = u32::from_le_bytes([
            data[offset + 2],
            data[offset + 3],
            data[offset + 4],
            data[offset + 5],
        ]);

        // Version and instance are packed in the same 16-bit field
        // Format: VVVV VVVV IIII IIII (V = version bits, I = instance bits)
        let version_instance = u16::from_le_bytes([data[offset + 6], data[offset + 7]]);
        let version = (version_instance >> 4) & 0x0FFF;  // High 12 bits for version
        let instance = version_instance & 0x0FFF;        // Low 12 bits for instance

        let record_type_enum = PptRecordType::from(record_type);
        let _total_size = 8 + data_length as usize;

        // Improved bounds checking: allow for truncated records at end of data
        // Only fail if we don't have the complete header or if the record claims
        // to extend beyond available data but we need at least some data
        if offset + 8 > data.len() {
            return Err(PptError::Corrupted("Not enough data for PPT record header".to_string()));
        }

        // Check if record data extends beyond available data
        let available_data_size = data.len().saturating_sub(offset + 8);
        if data_length as usize > available_data_size {
            // If this is a container record and we have at least some data, try to parse partially
            if Self::is_container_record(record_type_enum) && available_data_size > 0 {
                // For container records, we can still parse what we have
            } else if available_data_size == 0 {
                return Err(PptError::Corrupted("Record extends beyond data bounds and no data available".to_string()));
            }
        }

        // Use available data size, but don't exceed what the record claims to need
        let actual_data_size = available_data_size.min(data_length as usize);
        let record_data = data[offset + 8..offset + 8 + actual_data_size].to_vec();

        let mut record = PptRecord {
            record_type: record_type_enum,
            version,
            instance,
            data_length: actual_data_size as u32, // Store actual data length, not claimed length
            data: record_data,
            children: Vec::new(),
        };

        // Parse children if this is a container record or specific record types that contain children
        if Self::is_container_record(record_type_enum) && actual_data_size > 0 {
            record.children = Self::parse_container_children(&data[offset + 8..offset + 8 + actual_data_size])?;
        }

        Ok((record, 8 + actual_data_size))
    }

    /// Check if a record type is a container that can hold child records.
    fn is_container_record(record_type: PptRecordType) -> bool {
        matches!(
            record_type,
            PptRecordType::Document |
            PptRecordType::Slide |
            PptRecordType::Notes |
            PptRecordType::MainMaster |
            PptRecordType::HeadersFooters |
            PptRecordType::ExObjList |
            PptRecordType::VBAInfo
        )
    }

    /// Parse child records from a container record.
    fn parse_container_children(data: &[u8]) -> Result<Vec<PptRecord>> {
        let mut children = Vec::new();
        let mut offset = 0;

        while offset + 8 <= data.len() {
            // Try to parse a child record, but handle errors gracefully
            match Self::parse(data, offset) {
                Ok((child, consumed)) => {
                    children.push(child);
                    offset += consumed;

                    // Prevent infinite loops by ensuring we make progress
                    if consumed == 0 {
                        break;
                    }
                }
                Err(_) => {
                    // If we can't parse a record, skip to the next possible position
                    // This handles corrupted or truncated records gracefully
                    offset += 1;
                    if offset + 8 > data.len() {
                        break;
                    }
                }
            }
        }

        Ok(children)
    }

    /// Find a child record of a specific type.
    pub fn find_child(&self, record_type: PptRecordType) -> Option<&PptRecord> {
        self.children.iter().find(|child| child.record_type == record_type)
    }

    /// Find all child records of a specific type.
    pub fn find_children(&self, record_type: PptRecordType) -> Vec<&PptRecord> {
        self.children.iter().filter(|child| child.record_type == record_type).collect()
    }

    /// Extract slide data from this record.
    /// This follows POI's logic for extracting slide content.
    pub fn extract_slide_data(&self) -> Option<Vec<u8>> {
        // Look for PPDrawing record which contains Escher data
        if let Some(ppdrawing) = self.find_child(PptRecordType::PPDrawing) {
            return Some(ppdrawing.data.clone());
        }

        // For Slide records, check if they contain Escher data directly
        if self.record_type == PptRecordType::Slide {
            // Slide records may contain Escher data in their data section
            if !self.data.is_empty() && self.data.len() > 8 {
                // Check if the data contains Escher records (records >= 0xF000)
                let first_record_type = u16::from_le_bytes([self.data[0], self.data[1]]);
                if first_record_type >= 0xF000 {
                    return Some(self.data.clone());
                }
            }
        }

        None
    }

    /// Extract document information from this record.
    /// This follows POI's Document record parsing logic.
    pub fn extract_document_info(&self) -> Option<DocumentInfo> {
        if self.record_type != PptRecordType::Document {
            return None;
        }

        let mut info = DocumentInfo::default();

        // Find DocumentAtom child record
        if let Some(document_atom) = self.find_child(PptRecordType::DocumentAtom) {
            info = Self::parse_document_atom(document_atom);
        }

        // Find Environment record
        if let Some(_env) = self.find_child(PptRecordType::Environment) {
            // Parse environment information
            // This would include slide size, color scheme, etc.
            info.has_environment = true;
        }

        // Find PPDrawingGroup record
        if let Some(_pp_drawing_group) = self.find_child(PptRecordType::PPDrawingGroup) {
            info.has_drawing_group = true;
        }

        Some(info)
    }

    /// Parse DocumentAtom record data.
    /// Based on POI's DocumentAtom parsing.
    fn parse_document_atom(record: &PptRecord) -> DocumentInfo {
        let mut info = DocumentInfo::default();

        if record.data.len() >= 20 { // DocumentAtom should have at least 20 bytes
            // Parse slide size information
            info.slide_width = u32::from_le_bytes([
                record.data[0], record.data[1], record.data[2], record.data[3]
            ]);
            info.slide_height = u32::from_le_bytes([
                record.data[4], record.data[5], record.data[6], record.data[7]
            ]);

            // Parse slide count
            info.slide_count = u32::from_le_bytes([
                record.data[8], record.data[9], record.data[10], record.data[11]
            ]) as usize;

            // Parse notes count
            info.notes_count = u32::from_le_bytes([
                record.data[12], record.data[13], record.data[14], record.data[15]
            ]) as usize;

            // Parse master count (usually 1)
            info.master_count = u32::from_le_bytes([
                record.data[16], record.data[17], record.data[18], record.data[19]
            ]) as usize;
        }

        info
    }

    /// Extract slide information from this record.
    /// This follows POI's Slide record parsing logic.
    pub fn extract_slide_info(&self) -> Option<SlideInfo> {
        if self.record_type != PptRecordType::Slide {
            return None;
        }

        let mut info = SlideInfo::default();

        // Find SlideAtom child record
        if let Some(slide_atom) = self.find_child(PptRecordType::SlideAtom) {
            info = Self::parse_slide_atom(slide_atom);
        }

        // Check if slide has drawing data
        if self.find_child(PptRecordType::PPDrawing).is_some() {
            info.has_drawing = true;
        }

        // Check if slide has notes
        if self.find_child(PptRecordType::Notes).is_some() {
            info.has_notes = true;
        }

        Some(info)
    }

    /// Parse SlideAtom record data.
    /// Based on POI's SlideAtom parsing.
    fn parse_slide_atom(record: &PptRecord) -> SlideInfo {
        let mut info = SlideInfo::default();

        if record.data.len() >= 12 { // SlideAtom should have at least 12 bytes
            // Parse slide layout (master slide reference)
            info.layout_id = u32::from_le_bytes([
                record.data[0], record.data[1], record.data[2], record.data[3]
            ]);

            // Parse master slide ID
            info.master_id = u32::from_le_bytes([
                record.data[4], record.data[5], record.data[6], record.data[7]
            ]);

            // Parse notes ID
            info.notes_id = u32::from_le_bytes([
                record.data[8], record.data[9], record.data[10], record.data[11]
            ]);
        }

        info
    }

    /// Extract text content from this record and its children.
    /// This follows POI's text extraction logic.
    pub fn extract_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        // Extract text from text-related records
        match self.record_type {
            PptRecordType::TextCharsAtom => {
                if let Ok(text) = Self::parse_text_chars_atom(&self.data) {
                    text_parts.push(text);
                }
            }
            PptRecordType::TextBytesAtom => {
                if let Ok(text) = Self::parse_text_bytes_atom(&self.data) {
                    text_parts.push(text);
                }
            }
            PptRecordType::CString => {
                if let Ok(text) = Self::parse_cstring(&self.data) {
                    text_parts.push(text);
                }
            }
            _ => {}
        }

        // Recursively extract text from children
        for child in &self.children {
            if let Ok(child_text) = child.extract_text() {
                if !child_text.is_empty() {
                    text_parts.push(child_text);
                }
            }
        }

        Ok(text_parts.join("\n"))
    }

    /// Parse TextCharsAtom record (UTF-16LE text content).
    /// Based on POI's TextCharsAtom.getText() method.
    fn parse_text_chars_atom(data: &[u8]) -> Result<String> {
        if data.is_empty() {
            return Ok(String::new());
        }

        // TextCharsAtom contains UTF-16LE encoded text (little-endian)
        // Use the same logic as POI's StringUtil.getFromUnicodeLE
        let text = Self::from_utf16le_lossy(data);

        // POI strips the trailing return character and null terminator if present
        let text = text.trim_end_matches('\r').trim_end_matches('\u{0}').to_string();

        Ok(text)
    }


    /// Convert UTF-16LE bytes to String (lossy conversion).
    /// This follows POI's StringUtil.getFromUnicodeLE logic.
    /// Optimized for performance with minimal allocations.
    fn from_utf16le_lossy(bytes: &[u8]) -> String {
        if bytes.is_empty() {
            return String::new();
        }

        // Pre-calculate capacity for the result string
        let estimated_chars = bytes.len() / 2;
        let mut result = String::with_capacity(estimated_chars);

        // Process in chunks of 2 bytes for better performance
        let mut i = 0;
        while i + 1 < bytes.len() {
            let code_unit = unsafe { u16::from_le_bytes(*(&bytes[i..i + 2] as *const [u8] as *const [u8; 2])) };
            i += 2;

            // Use match expression for cleaner character processing
            match TextCharAction::process_utf16_char(code_unit) {
                TextCharAction::Add(ch) => result.push(ch),
                TextCharAction::Stop => break,
                TextCharAction::Skip => continue,
            }
        }

        // Shrink to fit if we over-allocated
        result.shrink_to_fit();
        result
    }

    /// Parse TextBytesAtom record (byte text content).
    /// Based on POI's TextBytesAtom.getText() method.
    fn parse_text_bytes_atom(data: &[u8]) -> Result<String> {
        if data.is_empty() {
            return Ok(String::new());
        }

        // TextBytesAtom contains text in "compressed unicode" format (ISO-8859-1)
        // This follows POI's StringUtil.getFromCompressedUnicode logic
        let text = data.iter().map(|&b| b as char).collect::<String>();

        // POI strips the trailing return character and null terminator if present
        let text = text.trim_end_matches('\r').trim_end_matches('\u{0}').to_string();

        Ok(text)
    }

    /// Parse CString record (null-terminated string).
    fn parse_cstring(data: &[u8]) -> Result<String> {
        // CString contains null-terminated ASCII text
        let null_pos = data.iter().position(|&b| b == 0).unwrap_or(data.len());
        let text = String::from_utf8_lossy(&data[..null_pos]).to_string();

        // POI strips the trailing return character if present
        let text = text.trim_end_matches('\r').to_string();

        Ok(text)
    }
}

/// Information extracted from a Document record.
/// Based on POI's Document and DocumentAtom record parsing.
#[derive(Debug, Clone, Default)]
pub struct DocumentInfo {
    /// Slide width in EMUs (English Metric Units)
    pub slide_width: u32,
    /// Slide height in EMUs
    pub slide_height: u32,
    /// Number of slides in the presentation
    pub slide_count: usize,
    /// Number of notes slides
    pub notes_count: usize,
    /// Number of master slides
    pub master_count: usize,
    /// Whether the document has an Environment record
    pub has_environment: bool,
    /// Whether the document has a PPDrawingGroup record
    pub has_drawing_group: bool,
}

/// Information extracted from a Slide record.
/// Based on POI's Slide and SlideAtom record parsing.
#[derive(Debug, Clone, Default)]
pub struct SlideInfo {
    /// Layout ID (reference to master slide)
    pub layout_id: u32,
    /// Master slide ID
    pub master_id: u32,
    /// Notes slide ID (0 if no notes)
    pub notes_id: u32,
    /// Whether the slide has drawing data (PPDrawing record)
    pub has_drawing: bool,
    /// Whether the slide has notes
    pub has_notes: bool,
}

/// Parser for PPT binary format that extracts document structure and content.
pub struct PptRecordParser {
    /// All parsed records
    records: Vec<PptRecord>,
    /// Slide records specifically
    slides: Vec<Vec<u8>>,
}

impl PptRecordParser {
    /// Create a new PPT record parser.
    pub fn new() -> Self {
        Self {
            records: Vec::new(),
            slides: Vec::new(),
        }
    }

    /// Parse a complete PPT document.
    ///
    /// # Arguments
    ///
    /// * `data` - The complete PowerPoint document data
    pub fn parse_document(&mut self, data: &[u8]) -> Result<()> {
        if data.is_empty() {
            return Ok(());
        }

        let mut offset = 0;
        while offset + 8 <= data.len() {
            // Try to parse a record, but handle errors gracefully
            match PptRecord::parse(data, offset) {
                Ok((record, consumed)) => {
                    self.records.push(record.clone());
                    offset += consumed;

                    // If this is a slide record, extract slide data
                    if record.record_type == PptRecordType::Slide {
                        if let Some(slide_data) = record.extract_slide_data() {
                            self.slides.push(slide_data);
                        }
                    }

                    // Prevent infinite loops by ensuring we make progress
                    if consumed == 0 {
                        break;
                    }
                }
                Err(_) => {
                    // If we can't parse a record, skip ahead by 1 byte
                    // This handles corrupted or truncated records gracefully
                    offset += 1;
                    if offset + 8 > data.len() {
                        break;
                    }
                }
            }
        }

        Ok(())
    }

    /// Get all slide data extracted from the document.
    pub fn slides(&self) -> &[Vec<u8>] {
        &self.slides
    }

    /// Get the number of slides in the document.
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Find a record of a specific type.
    pub fn find_record(&self, record_type: PptRecordType) -> Option<&PptRecord> {
        self.records.iter().find(|record| record.record_type == record_type)
    }

    /// Find all records of a specific type.
    pub fn find_records(&self, record_type: PptRecordType) -> Vec<&PptRecord> {
        self.records.iter().filter(|record| record.record_type == record_type).collect()
    }

    /// Extract all text content from the document.
    pub fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        for record in &self.records {
            if let Ok(record_text) = record.extract_text() {
                if !record_text.is_empty() {
                    text_parts.push(record_text);
                }
            }
        }

        if text_parts.is_empty() {
            Ok("No text content found".to_string())
        } else {
            Ok(text_parts.join("\n\n"))
        }
    }

    /// Extract text content from slide data.
    /// This follows POI's text extraction logic for PPT slides.
    pub fn extract_text_from_slide_data(slide_data: &[u8]) -> Result<String> {
        if slide_data.is_empty() {
            return Ok(String::new());
        }

        let mut text_parts = Vec::new();

        // Parse Escher records in the slide data to find text content
        let mut offset = 0;
        while offset + 8 <= slide_data.len() {
            // Check if this is an Escher record (records >= 0xF000)
            let record_type = u16::from_le_bytes([slide_data[offset], slide_data[offset + 1]]);
            if record_type >= 0xF000 {
                // This is an Escher record, parse it
                if let Ok((record, consumed)) = super::shapes::escher::EscherRecord::parse(slide_data, offset) {
                    // Extract text from this record
                    let record_text = record.extract_text().unwrap_or_default();
                    if !record_text.is_empty() {
                        text_parts.push(record_text);
                    }
                    offset += consumed;
                } else {
                    break; // Stop parsing if we can't parse a record
                }
            } else {
                // This is a PPT record - parse it using PptRecord parser
                match PptRecord::parse(slide_data, offset) {
                    Ok((ppt_record, consumed)) => {
                        // Extract text from PPT record
                        if let Ok(ppt_text) = ppt_record.extract_text() {
                            if !ppt_text.is_empty() {
                                text_parts.push(ppt_text);
                            }
                        }
                        offset += consumed;
                    }
                    Err(_) => {
                        // If parsing fails, try to skip the record gracefully
                        let data_length = if offset + 6 <= slide_data.len() {
                            u32::from_le_bytes([
                                slide_data[offset + 2],
                                slide_data[offset + 3],
                                slide_data[offset + 4],
                                slide_data[offset + 5],
                            ])
                        } else {
                            break;
                        };

                        let total_size = 8 + data_length as usize;
                        if offset + total_size > slide_data.len() {
                            break;
                        }
                        offset += total_size;
                    }
                }
            }
        }

        Ok(text_parts.join("\n"))
    }
}

impl Default for PptRecordParser {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ppt_record_creation() {
        let record = PptRecord {
            record_type: PptRecordType::Document,
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

    #[test]
    fn test_ppt_record_type_conversion() {
        assert_eq!(PptRecordType::from(1000), PptRecordType::Document);
        assert_eq!(PptRecordType::from(1006), PptRecordType::Slide);
        assert_eq!(PptRecordType::from(1036), PptRecordType::PPDrawing);
        assert_eq!(PptRecordType::from(4000), PptRecordType::TextCharsAtom);
        assert_eq!(PptRecordType::from(999), PptRecordType::Unknown);
    }

    #[test]
    fn test_text_parsing() {
        // Test TextCharsAtom parsing
        let text_data = vec![
            0x48, 0x00, // 'H'
            0x65, 0x00, // 'e'
            0x6C, 0x00, // 'l'
            0x6C, 0x00, // 'l'
            0x6F, 0x00, // 'o'
            0x00, 0x00, // null terminator
        ];

        let text = PptRecord::parse_text_chars_atom(&text_data).unwrap();
        assert_eq!(text, "Hello");
    }

    #[test]
    fn test_parser_creation() {
        let parser = PptRecordParser::new();
        assert_eq!(parser.slide_count(), 0);
        assert!(parser.slides().is_empty());
    }
}
