//! High-Level Text Extraction API
//!
//! Provides utilities for extracting text from iWork document objects.

use crate::iwa::bundle::Bundle;
use crate::iwa::archive::ArchiveObject;
use crate::iwa::Result;
use super::storage::{TextStorage, parse_storage_archive};

/// Text extractor for iWork documents
pub struct TextExtractor {
    /// Extracted text storages
    storages: Vec<TextStorage>,
}

impl TextExtractor {
    /// Create a new text extractor
    pub fn new() -> Self {
        Self {
            storages: Vec::new(),
        }
    }

    /// Extract text from a bundle
    pub fn extract_from_bundle(&mut self, bundle: &Bundle) -> Result<()> {
        // Find all TSWP storage objects (message types 200-205, 2001-2022)
        let storage_types = [
            200, 201, 202, 203, 204, 205,
            2001, 2002, 2003, 2004, 2005,
            2011, 2012, 2022,
        ];

        for type_id in storage_types {
            let objects = bundle.find_objects_by_type(type_id);
            for (_archive_name, object) in objects {
                if let Ok(storage) = self.extract_from_object(object)
                    && !storage.is_empty() {
                        self.storages.push(storage);
                    }
            }
        }

        Ok(())
    }

    /// Extract text from a single archive object
    pub fn extract_from_object(&self, object: &ArchiveObject) -> Result<TextStorage> {
        // Extract text from decoded messages
        let text_lines = object.extract_text();
        
        if text_lines.is_empty() {
            return Ok(TextStorage::new());
        }

        parse_storage_archive(&text_lines)
    }

    /// Get all extracted text as a single string
    pub fn get_text(&self) -> String {
        self.storages
            .iter()
            .map(|s| s.plain_text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get all text storages
    pub fn storages(&self) -> &[TextStorage] {
        &self.storages
    }

    /// Get number of text storages found
    pub fn storage_count(&self) -> usize {
        self.storages.len()
    }

    /// Clear all extracted text
    pub fn clear(&mut self) {
        self.storages.clear();
    }
}

impl Default for TextExtractor {
    fn default() -> Self {
        Self::new()
    }
}

/// Quick text extraction function for convenience
pub fn extract_text_from_bundle(bundle: &Bundle) -> Result<String> {
    let mut extractor = TextExtractor::new();
    extractor.extract_from_bundle(bundle)?;
    Ok(extractor.get_text())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_extractor_creation() {
        let extractor = TextExtractor::new();
        assert_eq!(extractor.storage_count(), 0);
        assert_eq!(extractor.get_text(), "");
    }

    #[test]
    fn test_text_extractor_clear() {
        let mut extractor = TextExtractor::new();
        extractor.storages.push(TextStorage::from_text("Test".to_string()));
        assert_eq!(extractor.storage_count(), 1);
        
        extractor.clear();
        assert_eq!(extractor.storage_count(), 0);
    }
}

