//! High-Level Text Extraction API
//!
//! Provides utilities for extracting text from iWork document objects.

use super::Storage;
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
    storages: Vec<Storage>,
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
    pub fn extract_from_object(&self, object: &ArchiveObject) -> Result<Storage> {
        // Extract text from decoded messages
        let text_lines = extract_text(object);

        if text_lines.is_empty() {
            return Ok(Storage::new());
        }

        Ok(Storage::from_text(text_lines.join("\n")))
    }

    /// Get all extracted text as a single string
    pub fn get_text(&self) -> String {
        let capacity = self
            .storages
            .len()
            .saturating_sub(1)
            .saturating_add(self.storages.iter().map(Storage::len).sum());
        let mut text = String::with_capacity(capacity);
        for (index, storage) in self.storages.iter().enumerate() {
            if index != 0 {
                text.push('\n');
            }
            text.push_str(storage.text());
        }
        text
    }

    /// Get all text storages
    pub fn storages(&self) -> &[Storage] {
        &self.storages
    }
}

impl Default for TextExtractor {
    fn default() -> Self {
        Self::new()
    }
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
    fn adapter_joins_decoded_storage_lines_without_native_metadata() {
        use prost::Message;

        let payload = crate::protobuf::tswp::StorageArchive {
            text: vec!["first".to_owned(), "second".to_owned()],
            ..Default::default()
        }
        .encode_to_vec();
        let object = ArchiveObject::new(
            42,
            vec![crate::archive::RawMessage {
                type_: 2001,
                data: payload,
            }],
        )
        .unwrap_or_else(|error| panic!("valid archive object: {error}"));

        let storage = TextExtractor::new()
            .extract_from_object(&object)
            .unwrap_or_else(|error| panic!("valid storage payload: {error}"));
        assert_eq!(storage.text(), "first\nsecond");
        assert_eq!(storage.runs().len(), 1);
        assert_eq!(storage.runs()[0].start(), 0);
        assert_eq!(storage.runs()[0].len(), "first\nsecond".len());
    }
}
