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
//! - `crate::pages::PagesDocument` for Pages-specific features
//! - `crate::numbers::NumbersDocument` for Numbers-specific features
//! - `crate::keynote::KeynoteDocument` for Keynote-specific features

use std::collections::HashMap;
use std::io::Write;
use std::path::Path;
use std::sync::Arc;

use crate::bundle::{Bundle, BundleLimits};
use crate::media::{MediaLimits, MediaManager, MediaStats};
use crate::object_index::{ObjectIndex, ResolvedObject};
use crate::package::PackageLimits;
use crate::ref_graph::ObjectId;
use crate::registry::{Application, detect_application_from_document};
use crate::structured::{self, StructuredData};
use crate::text::TextExtractor;
use crate::{Error, Result};

/// Unified iWork document interface
#[derive(Debug, Clone)]
pub struct Document {
    state: Arc<DocumentState>,
}

#[derive(Debug)]
struct DocumentState {
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
        Self::open_with_limits(path, BundleLimits::default())
    }

    /// Open an iWork document under caller-selected bundle ingress ceilings.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: BundleLimits) -> Result<Self> {
        let path_ref = path.as_ref();
        let bundle = Bundle::open_with_limits(path_ref, limits)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        // Application ownership is established only by the validated root
        // DocumentArchive envelope. Numeric message IDs overlap across iWork
        // applications and are not safe evidence for a fallback guess.
        let application = detect_bundle_application(&bundle).unwrap_or(Application::Common);

        // Keep media discovery and later file-backed extraction under the same
        // caller-selected package and aggregate media ceilings as document
        // ingress. Media is optional, so a source without a usable Data tree
        // remains a valid document.
        let (media_limits, package_limits) = media_profiles(limits)?;
        let media_manager = MediaManager::new_with_limits_and_package_limits(
            path_ref,
            media_limits,
            package_limits,
        )
        .ok();

        Ok(Self::from_parts(
            bundle,
            object_index,
            application,
            media_manager,
        ))
    }

    /// Open an iWork document from raw bytes
    ///
    /// This allows parsing iWork documents directly from memory without
    /// requiring file system access. Media extraction remains available when
    /// the in-memory package contains materialized `Data/*` assets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::Document;
    /// use std::fs;
    ///
    /// let data = fs::read("document.pages")?;
    /// let doc = Document::from_bytes(&data)?;
    /// let text = doc.text()?;
    /// println!("Extracted text: {}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, BundleLimits::default())
    }

    /// Open an iWork document from bytes under caller-selected bundle limits.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::from_bytes_with_limits(bytes, limits)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        // Application ownership is established only by the validated root
        // DocumentArchive envelope. Numeric message IDs overlap across iWork
        // applications and are not safe evidence for a fallback guess.
        let application = detect_bundle_application(&bundle).unwrap_or(Application::Common);

        let (media_limits, package_limits) = media_profiles(limits)?;
        let media_manager = MediaManager::from_bytes_with_limits_and_package_limits(
            bytes,
            media_limits,
            package_limits,
        )
        .ok();

        Ok(Self::from_parts(
            bundle,
            object_index,
            application,
            media_manager,
        ))
    }

    fn from_parts(
        bundle: Bundle,
        object_index: ObjectIndex,
        application: Application,
        media_manager: Option<MediaManager>,
    ) -> Self {
        Self {
            state: Arc::new(DocumentState {
                bundle,
                object_index,
                application,
                media_manager,
            }),
        }
    }

    /// Capture a cheap immutable snapshot that shares all parsed document state.
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Get the document's text content
    ///
    /// This method uses the modern text extraction API that efficiently
    /// processes TSWP storage objects across all iWork applications.
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.state.bundle)?;
        Ok(extractor.get_text())
    }

    /// Resolve all objects in the document in deterministic object-ID order.
    pub fn objects(&self) -> Result<Vec<ResolvedObject>> {
        self.try_objects()
    }

    /// Resolve all indexed objects in deterministic object-ID order.
    ///
    /// This named compatibility entry point remains available for callers
    /// that prefer an explicitly fallible method name. It uses the typed
    /// object index and batches archive lookups.
    pub fn try_objects(&self) -> Result<Vec<ResolvedObject>> {
        let object_ids = self.state.object_index.object_ids()?;
        self.state
            .object_index
            .resolve_many(&self.state.bundle, &object_ids)
    }

    /// Get an object through the validated identity API.
    pub fn object(&self, object_id: ObjectId) -> Result<Option<ResolvedObject>> {
        self.state
            .object_index
            .resolve(&self.state.bundle, object_id)
    }

    /// Get an object by its legacy raw numeric ID.
    #[deprecated(note = "use object(ObjectId) for checked identity semantics")]
    pub fn get_object(&self, id: u64) -> Result<Option<ResolvedObject>> {
        self.state
            .object_index
            .resolve_object(&self.state.bundle, id)
    }

    /// Get the immutable typed object index backing this document.
    pub fn object_index(&self) -> &ObjectIndex {
        &self.state.object_index
    }

    /// Get all validated object identities in deterministic numeric order.
    pub fn object_ids(&self) -> Result<Vec<ObjectId>> {
        self.state.object_index.object_ids()
    }

    /// Get the application type
    pub fn application(&self) -> Application {
        self.state.application
    }

    /// Get the underlying bundle
    pub fn bundle(&self) -> &Bundle {
        &self.state.bundle
    }

    /// Return a bounded, deterministic validation report for this snapshot.
    pub fn validation_report(&self) -> crate::bundle::BundleValidationReport {
        self.state.bundle.validation_report()
    }

    /// Validate this immutable snapshot without mutating it or emitting
    /// process-wide diagnostics.
    pub fn validate(&self) -> Result<()> {
        self.validation_report().as_result()
    }

    /// Get document metadata
    pub fn metadata(&self) -> &crate::bundle::BundleMetadata {
        self.state.bundle.metadata()
    }

    /// Get the media manager (if available)
    pub fn media_manager(&self) -> Option<&MediaManager> {
        self.state.media_manager.as_ref()
    }

    /// Get media statistics
    pub fn media_stats(&self) -> Option<MediaStats> {
        self.state.media_manager.as_ref().map(|m| m.stats())
    }

    /// Extract a media asset by filename
    pub fn extract_media(&self, filename: &str) -> Result<Vec<u8>> {
        let manager = self
            .state
            .media_manager
            .as_ref()
            .ok_or_else(|| Error::Bundle("Media manager not available".to_string()))?;
        manager.extract(filename)
    }

    /// Stream a materialized media asset to a caller-owned sequential sink.
    pub fn extract_media_to_writer<W: Write>(&self, filename: &str, sink: W) -> Result<()> {
        let manager = self
            .state
            .media_manager
            .as_ref()
            .ok_or_else(|| Error::Bundle("Media manager not available".to_string()))?;
        manager.extract_to_writer(filename, sink)
    }

    /// Atomically extract a materialized media asset to a regular file.
    pub fn extract_media_to_file<P: AsRef<Path>>(
        &self,
        filename: &str,
        output_path: P,
    ) -> Result<()> {
        let manager = self
            .state
            .media_manager
            .as_ref()
            .ok_or_else(|| Error::Bundle("Media manager not available".to_string()))?;
        manager.extract_to_file(filename, output_path.as_ref())
    }

    /// Extract structured data from the document
    ///
    /// This returns tables, slides, sections, and other structured content
    /// depending on the document type (Numbers, Keynote, or Pages).
    pub fn extract_structured_data(&self) -> Result<StructuredData> {
        structured::extract_all(&self.state.bundle, &self.state.object_index)
    }

    /// Get document statistics after resolving the indexed object set.
    pub fn stats(&self) -> Result<DocumentStats> {
        let total_objects = self.state.object_index.object_ids()?.len();
        let archives_count = self.state.bundle.iter_archives().count();

        let mut message_type_counts = HashMap::new();
        for object in self.objects()? {
            for &msg_type in &object.message_types() {
                *message_type_counts.entry(msg_type).or_insert(0) += 1;
            }
        }

        let media_stats = self.media_stats();

        Ok(DocumentStats {
            total_objects,
            archives_count,
            message_type_counts,
            application: self.state.application,
            media_stats,
        })
    }
}

fn media_profiles(limits: BundleLimits) -> Result<(MediaLimits, PackageLimits)> {
    let package_limits = PackageLimits::new_with_limits(
        limits.max_input_bytes(),
        limits.max_entries(),
        limits.max_entry_bytes(),
        limits.max_total_bytes(),
        limits.max_iwa_stream_bytes(),
    )?;
    let media_limits = MediaLimits::new(
        limits.max_entries(),
        limits.max_entry_bytes(),
        limits.max_total_bytes(),
    )?;
    Ok((media_limits, package_limits))
}

fn detect_bundle_application(bundle: &Bundle) -> Option<Application> {
    bundle
        .iter_archives()
        .filter(|(name, _)| name.ends_with("/Document.iwa") || *name == "Document.iwa")
        .flat_map(|(_, archive)| &archive.objects)
        .filter(|object| object.archive_info.identifier == Some(1))
        .flat_map(|object| &object.messages)
        .find_map(|message| detect_application_from_document(&message.data))
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
            .max_by(|(left_type, left_count), (right_type, right_count)| {
                left_count
                    .cmp(right_count)
                    .then_with(|| right_type.cmp(left_type))
            })
            .map(|(&type_, &count)| (type_, count))
    }

    /// Get message type distribution as a deterministic summary string.
    pub fn message_type_summary(&self) -> String {
        let mut types: Vec<_> = self.message_type_counts.iter().collect();
        types.sort_by(|(left_type, left_count), (right_type, right_count)| {
            right_count
                .cmp(left_count)
                .then_with(|| left_type.cmp(right_type))
        });

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
// - crate::pages::PagesDocument
// - crate::numbers::NumbersDocument
// - crate::keynote::KeynoteDocument
//
// The unified Document type above works with all formats and provides
// common functionality. For application-specific features, use the
// specialized document types in their respective modules.

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
    use crate::registry::detect_application;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn documents_are_send_and_sync() {
        assert_send_sync::<Document>();
    }

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
    fn document_statistics_break_ties_by_message_type() {
        let message_type_counts =
            HashMap::from([(7, 10), (3, 10), (5, 10), (1, 10), (9, 10), (2, 10)]);
        let stats = DocumentStats {
            total_objects: 0,
            archives_count: 0,
            message_type_counts,
            application: Application::Common,
            media_stats: None,
        };

        assert_eq!(stats.most_common_message_type(), Some((1, 10)));
        assert_eq!(
            stats.message_type_summary(),
            "1: 10, 2: 10, 3: 10, 5: 10, 7: 10 (and 1 more)"
        );
    }

    #[test]
    fn test_application_detection() {
        // Use IDs that are unique to Keynote's registry entries. Shared IDs
        // cannot safely identify document ownership.
        let keynote_types = vec![145, 146, 147, 148];
        let keynote_result = detect_application(&keynote_types);
        assert_eq!(keynote_result, Some(Application::Keynote));

        // Shared and application-overlapping IDs fail closed.
        assert_eq!(detect_application(&[1, 1, 1, 101]), None);

        // Test empty input
        assert_eq!(detect_application(&[]), None);
    }

    #[test]
    fn document_application_detection_fails_closed_on_overlapping_message_ids() {
        let object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 101,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let mut package = crate::IWorkPackage::new();
        package
            .replace_archive(
                "Index/Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        let document = Document::from_bytes(&package.to_bytes().unwrap()).unwrap();
        assert_eq!(document.application(), Application::Common);
        assert!(document.validate().is_ok());
        assert!(document.validation_report().is_valid());
    }

    #[test]
    fn document_exposes_streaming_media_extraction() {
        let object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 101,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let mut package = crate::IWorkPackage::new();
        package
            .replace_archive(
                "Index/Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();
        package
            .insert_entry("Data/image.png", b"media-bytes".to_vec())
            .unwrap();

        let document = Document::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let mut streamed = Vec::new();
        document
            .extract_media_to_writer("image.png", &mut streamed)
            .unwrap();
        assert_eq!(streamed, b"media-bytes");
    }

    #[test]
    fn document_media_profiles_follow_bundle_limits() {
        let limits = BundleLimits::new(31, 7, 11, 23, 47).unwrap();
        let (media_limits, package_limits) = media_profiles(limits).unwrap();

        assert_eq!(media_limits.max_assets(), 7);
        assert_eq!(media_limits.max_asset_bytes(), 11);
        assert_eq!(media_limits.max_total_bytes(), 23);
        assert_eq!(package_limits.max_input_bytes(), 31);
        assert_eq!(package_limits.max_entries(), 7);
        assert_eq!(package_limits.max_entry_bytes(), 11);
        assert_eq!(package_limits.max_total_bytes(), 23);
        assert_eq!(package_limits.max_iwa_stream_bytes(), 47);
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
        let objects = doc.objects().expect("test document should resolve objects");
        assert!(!objects.is_empty(), "Document should contain objects");

        // Verify we can get stats
        let stats = doc
            .stats()
            .expect("test document should produce statistics");
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
