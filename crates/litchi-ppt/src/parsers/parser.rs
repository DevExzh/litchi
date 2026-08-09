//! PPT record parser - orchestrates document parsing and text extraction.
//!
//! Based on Apache POI's `HSLFSlideShow` and `QuickButCruddyTextExtractor`.

use crate::consts::RecordType;
use crate::package::{Error, RecordLimits, Result};
use crate::records::Record;
use crate::records::record::ParseBudget;

/// Parser for PPT binary format that extracts document structure and content.
#[allow(
    clippy::module_name_repetitions,
    reason = "`RecordParser` is the established public API name re-exported from the crate root; renaming it would break downstream crates"
)]
pub struct RecordParser {
    /// All parsed records
    records: Vec<Record>,
    /// Slide text organized by `SlideAtomsSets` (following POI's architecture)
    slide_atoms_sets: Vec<Vec<u8>>,
}

impl RecordParser {
    /// Create a new PPT record parser.
    #[must_use]
    pub fn new() -> Self {
        Self {
            records: Vec::new(),
            slide_atoms_sets: Vec::new(),
        }
    }

    /// Parse a complete PPT document.
    ///
    /// This method follows POI's parsing approach:
    /// 1. Parse all records in the document
    /// 2. Find the Document record
    /// 3. Extract text from all records
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_document(&mut self, data: &[u8]) -> Result<()> {
        self.parse_document_with_limits(data, RecordLimits::default())
    }

    /// Parse a complete PPT document with a shared finite record budget.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_document_with_limits(&mut self, data: &[u8], limits: RecordLimits) -> Result<()> {
        if data.is_empty() {
            return Ok(());
        }

        let mut budget = ParseBudget::new(limits, data.len())?;

        // Parse all top-level records
        let mut offset = 0;
        while offset + 8 <= data.len() {
            match Record::parse_with_budget(data, offset, false, &mut budget) {
                Ok((record, consumed)) => {
                    self.records
                        .try_reserve(1)
                        .map_err(|_err| Error::AllocationFailed("PPT top-level record table"))?;
                    self.records.push(record);
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

        // Extract slide text
        self.extract_slide_text_from_document()?;

        Ok(())
    }

    /// Parse only the validated live persist objects of an encrypted document.
    #[cfg(feature = "encryption")]
    pub(crate) fn parse_document_at_offsets_with_limits(
        &mut self,
        data: &[u8],
        offsets: &[usize],
        limits: RecordLimits,
    ) -> Result<()> {
        let mut budget = ParseBudget::new(limits, data.len())?;
        for &offset in offsets {
            let (record, _) = Record::parse_with_budget(data, offset, true, &mut budget)?;
            self.records
                .try_reserve(1)
                .map_err(|_err| Error::AllocationFailed("PPT live-record table"))?;
            self.records.push(record);
        }
        self.extract_slide_text_from_document()
    }

    /// Extract slide text from the document.
    /// Based on POI's `QuickButCruddyTextExtractor` approach.
    fn extract_slide_text_from_document(&mut self) -> Result<()> {
        let all_text = self.extract_all_text()?;

        // For now, treat all text as a single slide
        // Note: Full slide association would require parsing SlideListWithText structure
        if !all_text.is_empty() && all_text != "No text content found" {
            self.slide_atoms_sets.push(all_text.into_bytes());
        }

        Ok(())
    }

    /// Get all slide text data extracted from the document.
    #[must_use]
    pub fn slides(&self) -> &[Vec<u8>] {
        &self.slide_atoms_sets
    }

    /// Get the number of slides in the document.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slide_atoms_sets.len()
    }

    /// Find a record of a specific type.
    #[must_use]
    pub fn find_record(&self, record_type: RecordType) -> Option<&Record> {
        self.records
            .iter()
            .find(|record| record.record_type == record_type)
    }

    /// Get all records recursively (for building persist mapping).
    ///
    /// # Deprecated
    ///
    /// This method clones all records including their data. Use `find_records_ref()` instead
    /// for zero-copy access.
    #[deprecated(note = "Use find_records_ref() instead to avoid expensive cloning")]
    #[allow(dead_code, reason = "deprecated API retained for compatibility")]
    #[allow(
        deprecated,
        reason = "the deprecated shim may reference deprecated internals"
    )]
    #[must_use]
    pub fn find_records(&self, _record_type: RecordType) -> Vec<Record> {
        // Collect all records recursively
        let mut all_records = Vec::new();
        Self::collect_records_recursive(&self.records, &mut all_records);
        all_records
    }

    /// Get all record references recursively (zero-copy version).
    ///
    /// # Performance
    ///
    /// - Zero-copy: returns references instead of cloning records
    /// - Significantly faster than `find_records()` for large presentations
    /// - Preferred method for building persist mappings and iterating records
    #[must_use]
    pub fn find_records_ref(&self) -> Vec<&Record> {
        let mut all_records = Vec::new();
        Self::collect_records_recursive_ref(&self.records, &mut all_records);
        all_records
    }

    /// Collect all record references with fallible traversal allocation.
    pub(crate) fn try_find_records_ref(&self) -> Result<Vec<&Record>> {
        let mut all_records = Vec::new();
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(self.records.len())
            .map_err(|_err| Error::AllocationFailed("PPT record traversal stack"))?;
        pending.extend(self.records.iter().rev());
        while let Some(record) = pending.pop() {
            all_records
                .try_reserve(1)
                .map_err(|_err| Error::AllocationFailed("PPT record reference table"))?;
            all_records.push(record);
            pending
                .try_reserve(record.children.len())
                .map_err(|_err| Error::AllocationFailed("PPT record traversal stack"))?;
            pending.extend(record.children.iter().rev());
        }
        Ok(all_records)
    }

    /// Recursively collect all records including children (cloning version).
    ///
    /// # Deprecated
    ///
    /// This clones all records including their Vec<u8> data, causing expensive memory copies.
    /// Use `collect_records_recursive_ref` instead for zero-copy collection.
    #[deprecated(note = "Use collect_records_recursive_ref instead to avoid expensive cloning")]
    #[allow(dead_code, reason = "deprecated API retained for compatibility")]
    #[allow(
        deprecated,
        reason = "the deprecated shim may reference deprecated internals"
    )]
    fn collect_records_recursive(records: &[Record], collector: &mut Vec<Record>) {
        for record in records {
            collector.push(record.clone());
            if !record.children.is_empty() {
                Self::collect_records_recursive(&record.children, collector);
            }
        }
    }

    /// Recursively collect all record references including children (zero-copy).
    fn collect_records_recursive_ref<'a>(records: &'a [Record], collector: &mut Vec<&'a Record>) {
        let mut pending: Vec<&Record> = records.iter().rev().collect();
        while let Some(record) = pending.pop() {
            collector.push(record);
            pending.extend(record.children.iter().rev());
        }
    }

    /// Find all records matching a specific type (filtered).
    pub fn filter_records(&self, record_type: RecordType) -> impl Iterator<Item = &Record> {
        self.records
            .iter()
            .filter(move |record| record.record_type == record_type)
    }

    /// Extract all text content from the document.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        for record in &self.records {
            if let Ok(record_text) = record.extract_text()
                && !record_text.is_empty()
            {
                for line in record_text.lines() {
                    let trimmed = line.trim();
                    if !trimmed.is_empty() {
                        text_parts.push(trimmed.to_string());
                    }
                }
            }
        }

        if text_parts.is_empty() {
            Ok("No text content found".to_string())
        } else {
            Ok(text_parts.join("\n"))
        }
    }

    /// Extract text content from slide data.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn extract_text_from_slide_data(slide_data: &[u8]) -> Result<String> {
        if slide_data.is_empty() {
            return Ok(String::new());
        }

        String::from_utf8(slide_data.to_vec())
            .map_err(|e| Error::InvalidFormat(format!("Invalid UTF-8 in slide text: {e}")))
    }
}

impl Default for RecordParser {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn test_parser_creation() {
        let parser = RecordParser::new();
        assert_eq!(parser.slide_count(), 0);
        assert!(parser.slides().is_empty());
    }

    #[test]
    fn document_parser_shares_count_across_top_level_records() {
        let mut data = Vec::new();
        for _ in 0..2 {
            data.extend_from_slice(&0u16.to_le_bytes());
            data.extend_from_slice(&0x2222u16.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
        }
        let mut parser = RecordParser::new();
        let error = parser
            .parse_document_with_limits(
                &data,
                RecordLimits {
                    max_records: 1,
                    ..RecordLimits::default()
                },
            )
            .unwrap_err();
        assert!(matches!(error, Error::ResourceLimit(message) if message.contains("record count")));
    }
}
