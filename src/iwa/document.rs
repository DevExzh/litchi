//! High-Level iWork Document API
//!
//! Provides user-friendly interfaces for working with iWork documents
//! (Pages, Keynote, Numbers) similar to the high-level APIs for
//! Microsoft Office formats.
//!
//! This module provides a unified `Document` interface that works with all
//! iWork formats. For application-specific features, use the specialized
//! modules:
//!
//! - `crate::iwa::pages::PagesDocument` for Pages-specific features
//! - `crate::iwa::numbers::NumbersDocument` for Numbers-specific features
//! - `crate::iwa::keynote::KeynoteDocument` for Keynote-specific features

use std::collections::HashMap;
use std::path::Path;

use crate::iwa::bundle::Bundle;
use crate::iwa::media::{MediaManager, MediaStats};
use crate::iwa::object_index::{ObjectIndex, ResolvedObject};
use crate::iwa::registry::{Application, detect_application};
use crate::iwa::structured::{self, StructuredData};
use crate::iwa::text::TextExtractor;
use crate::iwa::{Error, Result};

/// Unified iWork document interface
#[derive(Debug)]
pub struct Document {
    /// The underlying bundle
    bundle: Bundle,
    /// Object index for cross-referencing
    object_index: ObjectIndex,
    /// Detected application type
    application: Application,
    /// Media manager for assets
    media_manager: Option<MediaManager>,
}

impl Document {
    /// Open an iWork document from a bundle path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path_ref = path.as_ref();
        let bundle = Bundle::open(path_ref)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        // Detect application type from message types
        let all_message_types: Vec<u32> = bundle
            .archives()
            .values()
            .flat_map(|archive| &archive.objects)
            .flat_map(|obj| &obj.messages)
            .map(|msg| msg.type_)
            .collect();

        let application = detect_application(&all_message_types).unwrap_or(Application::Common);

        // Try to create media manager (may fail for single-file bundles)
        let media_manager = MediaManager::new(path_ref).ok();

        Ok(Document {
            bundle,
            object_index,
            application,
            media_manager,
        })
    }

    /// Open an iWork document from raw bytes
    ///
    /// This allows parsing iWork documents directly from memory without
    /// requiring file system access. Note that media extraction is not
    /// available when opening from bytes.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::Document;
    /// use std::fs;
    ///
    /// let data = fs::read("document.pages")?;
    /// let doc = Document::from_bytes(&data)?;
    /// let text = doc.text()?;
    /// println!("Extracted text: {}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let bundle = Bundle::from_bytes(bytes)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        // Detect application type from message types
        let all_message_types: Vec<u32> = bundle
            .archives()
            .values()
            .flat_map(|archive| &archive.objects)
            .flat_map(|obj| &obj.messages)
            .map(|msg| msg.type_)
            .collect();

        let application = detect_application(&all_message_types).unwrap_or(Application::Common);

        Ok(Document {
            bundle,
            object_index,
            application,
            media_manager: None, // No media access from bytes
        })
    }

    /// Get the document's text content
    ///
    /// This method uses the modern text extraction API that efficiently
    /// processes TSWP storage objects across all iWork applications.
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.bundle)?;
        Ok(extractor.get_text())
    }

    /// Get all objects in the document
    pub fn objects(&self) -> Vec<ResolvedObject> {
        self.object_index
            .all_object_ids()
            .iter()
            .filter_map(|&id| {
                self.object_index
                    .resolve_object(&self.bundle, id)
                    .ok()
                    .flatten()
            })
            .collect()
    }

    /// Get an object by ID
    pub fn get_object(&self, id: u64) -> Result<Option<ResolvedObject>> {
        self.object_index.resolve_object(&self.bundle, id)
    }

    /// Get the application type
    pub fn application(&self) -> Application {
        self.application
    }

    /// Get the underlying bundle
    pub fn bundle(&self) -> &Bundle {
        &self.bundle
    }

    /// Get document metadata
    pub fn metadata(&self) -> &crate::iwa::bundle::BundleMetadata {
        self.bundle.metadata()
    }

    /// Get the media manager (if available)
    pub fn media_manager(&self) -> Option<&MediaManager> {
        self.media_manager.as_ref()
    }

    /// Get media statistics
    pub fn media_stats(&self) -> Option<MediaStats> {
        self.media_manager.as_ref().map(|m| m.stats())
    }

    /// Extract a media asset by filename
    pub fn extract_media(&self, filename: &str) -> Result<Vec<u8>> {
        let manager = self
            .media_manager
            .as_ref()
            .ok_or_else(|| Error::Bundle("Media manager not available".to_string()))?;
        manager.extract(filename)
    }

    /// Extract structured data from the document
    ///
    /// This returns tables, slides, sections, and other structured content
    /// depending on the document type (Numbers, Keynote, or Pages).
    pub fn extract_structured_data(&self) -> Result<StructuredData> {
        structured::extract_all(&self.bundle, &self.object_index)
    }

    /// Get document statistics
    pub fn stats(&self) -> DocumentStats {
        let total_objects = self.object_index.all_object_ids().len();
        let archives_count = self.bundle.archives().len();

        let mut message_type_counts = HashMap::new();
        for object in self.objects() {
            for &msg_type in &object.message_types() {
                *message_type_counts.entry(msg_type).or_insert(0) += 1;
            }
        }

        let media_stats = self.media_stats();

        DocumentStats {
            total_objects,
            archives_count,
            message_type_counts,
            application: self.application,
            media_stats,
        }
    }
}

/// Statistics about a document
#[derive(Debug, Clone)]
pub struct DocumentStats {
    /// Total number of objects
    pub total_objects: usize,
    /// Number of archives
    pub archives_count: usize,
    /// Count of each message type
    pub message_type_counts: HashMap<u32, usize>,
    /// Application type
    pub application: Application,
    /// Media statistics (if available)
    pub media_stats: Option<MediaStats>,
}

impl DocumentStats {
    /// Get the most common message type
    pub fn most_common_message_type(&self) -> Option<(u32, usize)> {
        self.message_type_counts
            .iter()
            .max_by_key(|&(_, count)| count)
            .map(|(&type_, &count)| (type_, count))
    }

    /// Get message type distribution as a string
    pub fn message_type_summary(&self) -> String {
        let mut types: Vec<_> = self.message_type_counts.iter().collect();
        types.sort_by_key(|&(_, count)| std::cmp::Reverse(*count));

        let top_types: Vec<String> = types
            .into_iter()
            .take(5)
            .map(|(type_, count)| format!("{}: {}", type_, count))
            .collect();

        if top_types.len() < self.message_type_counts.len() {
            format!(
                "{} (and {} more)",
                top_types.join(", "),
                self.message_type_counts.len() - top_types.len()
            )
        } else {
            top_types.join(", ")
        }
    }
}

// Note: Application-specific document types have been moved to dedicated modules:
// - crate::iwa::pages::PagesDocument
// - crate::iwa::numbers::NumbersDocument
// - crate::iwa::keynote::KeynoteDocument
//
// The unified Document type above works with all formats and provides
// common functionality. For application-specific features, use the
// specialized document types in their respective modules.

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_document_stats() {
        let mut message_counts = HashMap::new();
        message_counts.insert(1, 10);
        message_counts.insert(2, 5);
        message_counts.insert(3, 15);

        let stats = DocumentStats {
            total_objects: 25,
            archives_count: 3,
            message_type_counts: message_counts,
            application: Application::Pages,
            media_stats: None,
        };

        assert_eq!(stats.total_objects, 25);
        assert_eq!(stats.archives_count, 3);
        assert_eq!(stats.most_common_message_type(), Some((3, 15)));

        let summary = stats.message_type_summary();
        assert!(summary.contains("3: 15"));
        assert!(summary.contains("1: 10"));
    }

    #[test]
    fn test_application_detection() {
        // Test Keynote detection (should work with current registry)
        let keynote_types = vec![101, 102, 103]; // KN.* types
        let keynote_result = detect_application(&keynote_types);
        assert!(keynote_result.is_some()); // Should detect some application

        // Test with mixed types
        let mixed_types = vec![1, 1, 1, 101]; // Mostly common types, one Keynote type
        let mixed_result = detect_application(&mixed_types);
        assert!(mixed_result.is_some()); // Should detect something

        // Test empty input
        assert_eq!(detect_application(&[]), None);
    }

    #[test]
    fn test_document_parsing() {
        let doc_path = std::path::Path::new("test.pages");
        if !doc_path.exists() {
            // Skip test if test file doesn't exist
            return;
        }

        let doc_result = Document::open(doc_path);
        assert!(
            doc_result.is_ok(),
            "Failed to open document: {:?}",
            doc_result.err()
        );

        let doc = doc_result.unwrap();

        // Verify we can get objects
        let objects = doc.objects();
        assert!(!objects.is_empty(), "Document should contain objects");

        // Verify we can get stats
        let stats = doc.stats();
        assert!(stats.total_objects > 0, "Document should have objects");

        // Test text extraction
        let text_result = doc.text();
        assert!(text_result.is_ok());
    }

    #[test]
    fn test_text_extraction() {
        let doc_path = std::path::Path::new("test.pages");
        if !doc_path.exists() {
            return;
        }

        let doc = Document::open(doc_path).unwrap();
        let text_result = doc.text();
        assert!(text_result.is_ok());

        // Text extraction should succeed even if result is empty
        let _text = text_result.unwrap();
    }
}
