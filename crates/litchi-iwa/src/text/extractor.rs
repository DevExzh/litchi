//! High-Level Text Extraction API
//!
//! Provides utilities for extracting text from iWork document objects.

use super::{TextStorage, parse_storage_archive};
use crate::Result;
use crate::archive::{ArchiveObject, extract_text};
use crate::bundle::Bundle;

const fn is_storage_message_type(type_id: u32) -> bool {
    matches!(
        type_id,
        200 | 201 | 202 | 203 | 204 | 205 | 2001 | 2002 | 2003 | 2004 | 2005 | 2011 | 2012 | 2022
    )
}

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
        // Find all TSWP storage objects in one archive traversal. An object
        // may carry more than one recognized message type; extracting it once
        // avoids duplicate text while removing fourteen full-bundle scans.
        for (_archive_name, object) in bundle.iter_objects().filter(|(_, object)| {
            object
                .messages
                .iter()
                .any(|message| is_storage_message_type(message.type_))
        }) {
            if let Ok(storage) = self.extract_from_object(object)
                && !storage.is_empty()
            {
                self.storages.push(storage);
            }
        }

        Ok(())
    }

    /// Extract text from a single archive object
    pub fn extract_from_object(&self, object: &ArchiveObject) -> Result<TextStorage> {
        // Extract text from decoded messages
        let text_lines = extract_text(object);

        if text_lines.is_empty() {
            return Ok(TextStorage::new());
        }

        Ok(parse_storage_archive(&text_lines))
    }

    /// Get all extracted text as a single string
    pub fn get_text(&self) -> String {
        let capacity = self
            .storages
            .len()
            .saturating_sub(1)
            .saturating_add(self.storages.iter().map(TextStorage::len).sum());
        let mut text = String::with_capacity(capacity);
        for (index, storage) in self.storages.iter().enumerate() {
            if index != 0 {
                text.push('\n');
            }
            text.push_str(storage.plain_text());
        }
        text
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
    fn recognizes_all_supported_storage_message_types_without_nearby_types() {
        for type_id in [
            200, 201, 202, 203, 204, 205, 2001, 2002, 2003, 2004, 2005, 2011, 2012, 2022,
        ] {
            assert!(is_storage_message_type(type_id));
        }
        for type_id in [0, 199, 206, 2000, 2006, 2010, 2013, 2021, 2023] {
            assert!(!is_storage_message_type(type_id));
        }
    }

    #[test]
    fn test_text_extractor_creation() {
        let extractor = TextExtractor::new();
        assert_eq!(extractor.storage_count(), 0);
        assert_eq!(extractor.get_text(), "");
    }

    #[test]
    fn test_text_extractor_clear() {
        let mut extractor = TextExtractor::new();
        extractor
            .storages
            .push(TextStorage::from_text("Test".to_string()));
        assert_eq!(extractor.storage_count(), 1);

        extractor.clear();
        assert_eq!(extractor.storage_count(), 0);
    }
}
