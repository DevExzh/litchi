/// PPT record parser - orchestrates document parsing and text extraction.
///
/// Based on Apache POI's HSLFSlideShow and QuickButCruddyTextExtractor.

use crate::ole::consts::PptRecordType;
use crate::ole::ppt::package::{PptError, Result};
use crate::ole::ppt::records::PptRecord;

/// Parser for PPT binary format that extracts document structure and content.
pub struct PptRecordParser {
    /// All parsed records
    records: Vec<PptRecord>,
    /// Slide text organized by SlideAtomsSets (following POI's architecture)
    slide_atoms_sets: Vec<Vec<u8>>,
    /// Document record if found
    document_record: Option<PptRecord>,
}

impl PptRecordParser {
    /// Create a new PPT record parser.
    pub fn new() -> Self {
        Self {
            records: Vec::new(),
            slide_atoms_sets: Vec::new(),
            document_record: None,
        }
    }

    /// Parse a complete PPT document.
    ///
    /// This method follows POI's parsing approach:
    /// 1. Parse all records in the document
    /// 2. Find the Document record
    /// 3. Extract text from all records
    pub fn parse_document(&mut self, data: &[u8]) -> Result<()> {
        if data.is_empty() {
            return Ok(());
        }

        // Parse all top-level records
        let mut offset = 0;
        while offset + 8 <= data.len() {
            match PptRecord::parse(data, offset) {
                Ok((record, consumed)) => {
                    // Save Document record for later processing
                    if record.record_type == PptRecordType::Document {
                        self.document_record = Some(record.clone());
                    }

                    self.records.push(record);
                    offset += consumed;

                    if consumed == 0 {
                        break;
                    }
                }
                Err(_) => {
                    offset += 1;
                    if offset + 8 > data.len() {
                        break;
                    }
                }
            }
        }

        // Extract slide text
        self.extract_slide_text_from_document()?;

        Ok(())
    }

    /// Extract slide text from the document.
    /// Based on POI's QuickButCruddyTextExtractor approach.
    fn extract_slide_text_from_document(&mut self) -> Result<()> {
        let all_text = self.extract_all_text()?;
        
        // For now, treat all text as a single slide
        // TODO: Properly associate text with individual slides
        if !all_text.is_empty() && all_text != "No text content found" {
            self.slide_atoms_sets.push(all_text.into_bytes());
        }

        Ok(())
    }

    /// Get all slide text data extracted from the document.
    pub fn slides(&self) -> &[Vec<u8>] {
        &self.slide_atoms_sets
    }

    /// Get the number of slides in the document.
    pub fn slide_count(&self) -> usize {
        self.slide_atoms_sets.len()
    }

    /// Find a record of a specific type.
    pub fn find_record(&self, record_type: PptRecordType) -> Option<&PptRecord> {
        self.records.iter().find(|record| record.record_type == record_type)
    }

    /// Get all records recursively (for building persist mapping).
    pub fn find_records(&self, _record_type: PptRecordType) -> Vec<PptRecord> {
        // Collect all records recursively
        let mut all_records = Vec::new();
        self.collect_records_recursive(&self.records, &mut all_records);
        all_records
    }
    
    /// Recursively collect all records including children.
    fn collect_records_recursive(&self, records: &[PptRecord], collector: &mut Vec<PptRecord>) {
        for record in records {
            collector.push(record.clone());
            if !record.children.is_empty() {
                self.collect_records_recursive(&record.children, collector);
            }
        }
    }
    
    /// Find all records matching a specific type (filtered).
    pub fn filter_records(&self, record_type: PptRecordType) -> impl Iterator<Item = &PptRecord> {
        self.records.iter().filter(move |record| record.record_type == record_type)
    }

    /// Extract all text content from the document.
    pub fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        for record in &self.records {
            match record.extract_text() {
                Ok(record_text) => {
                    if !record_text.is_empty() {
                        for line in record_text.lines() {
                            let trimmed = line.trim();
                            if !trimmed.is_empty() {
                                text_parts.push(trimmed.to_string());
                            }
                        }
                    }
                }
                Err(_) => continue,
            }
        }

        if text_parts.is_empty() {
            Ok("No text content found".to_string())
        } else {
            Ok(text_parts.join("\n"))
        }
    }

    /// Extract text content from slide data.
    pub fn extract_text_from_slide_data(slide_data: &[u8]) -> Result<String> {
        if slide_data.is_empty() {
            return Ok(String::new());
        }

        String::from_utf8(slide_data.to_vec())
            .map_err(|e| PptError::InvalidFormat(format!("Invalid UTF-8 in slide text: {}", e)))
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
    fn test_parser_creation() {
        let parser = PptRecordParser::new();
        assert_eq!(parser.slide_count(), 0);
        assert!(parser.slides().is_empty());
    }
}

