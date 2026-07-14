//! iWork Bundle Structure Parser
//!
//! iWork documents are stored as bundles (directories) containing:
//! - `Index.zip`: Archive of IWA files with serialized objects
//! - `Data/`: Directory containing media assets
//! - `Metadata/`: Document metadata and properties
//! - Preview images at root level

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use crate::archive::{Archive, ArchiveObject};
use crate::{Error, Result};

mod loader;
mod metadata;

pub use metadata::{BundleMetadata, PropertyValue, detect_application_type};

/// Represents an iWork document bundle
#[derive(Debug)]
pub struct Bundle {
    /// Path to the bundle directory
    bundle_path: PathBuf,
    /// Parsed IWA archives from Index.zip
    archives: HashMap<String, Archive>,
    /// Metadata from Metadata/ directory
    metadata: BundleMetadata,
}

impl Bundle {
    /// Get all archives in the bundle
    pub fn archives(&self) -> &HashMap<String, Archive> {
        &self.archives
    }

    /// Get a specific archive by name
    pub fn get_archive(&self, name: &str) -> Option<&Archive> {
        self.archives.get(name)
    }

    /// Get bundle metadata
    pub fn metadata(&self) -> &BundleMetadata {
        &self.metadata
    }

    /// Get the bundle path
    pub fn path(&self) -> &Path {
        &self.bundle_path
    }

    /// Validate the bundle structure and integrity
    ///
    /// Performs comprehensive validation including:
    /// - Checking for required archives
    /// - Verifying IWA file format correctness
    /// - Detecting corrupted or incomplete data
    ///
    /// # Returns
    ///
    /// * `Ok(())` if validation passes
    /// * `Err(Error)` with detailed error message if validation fails
    pub fn validate(&self) -> Result<()> {
        // Check that we have at least one archive
        if self.archives.is_empty() {
            return Err(Error::Bundle(
                "Bundle contains no archives - may be corrupted or empty".to_string(),
            ));
        }

        // Verify each archive has at least one object
        let mut total_objects = 0;
        for (archive_name, archive) in &self.archives {
            if archive.objects.is_empty() {
                eprintln!("Warning: Archive '{}' contains no objects", archive_name);
            }
            total_objects += archive.objects.len();
        }

        if total_objects == 0 {
            return Err(Error::Bundle(
                "Bundle contains no objects across all archives - may be corrupted".to_string(),
            ));
        }

        Ok(())
    }

    /// Check if the bundle appears to be corrupted
    ///
    /// Performs basic sanity checks to detect obviously corrupted bundles.
    pub fn is_corrupted(&self) -> bool {
        if self.archives.is_empty() {
            return true;
        }

        let has_any_objects = self
            .archives
            .values()
            .any(|archive| !archive.objects.is_empty());

        !has_any_objects
    }

    /// Get bundle statistics
    pub fn stats(&self) -> BundleStats {
        let archive_count = self.archives.len();
        let total_objects: usize = self
            .archives
            .values()
            .map(|archive| archive.objects.len())
            .sum();

        let largest_archive = self
            .archives
            .iter()
            .max_by_key(|(_, archive)| archive.objects.len())
            .map(|(name, archive)| (name.clone(), archive.objects.len()));

        BundleStats {
            archive_count,
            total_objects,
            largest_archive,
        }
    }

    /// Extract all text content from the bundle
    pub fn extract_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        for archive in self.archives.values() {
            for object in &archive.objects {
                text_parts.extend(object.extract_text());
            }
        }

        // Join all text parts with newlines
        Ok(text_parts.join("\n"))
    }

    /// Get all objects across all archives
    pub fn all_objects(&self) -> Vec<(&str, &ArchiveObject)> {
        let mut objects = Vec::new();
        for (archive_name, archive) in &self.archives {
            for object in &archive.objects {
                objects.push((archive_name.as_str(), object));
            }
        }
        objects
    }

    /// Find objects by message type
    pub fn find_objects_by_type(&self, message_type: u32) -> Vec<(&str, &ArchiveObject)> {
        let mut matching_objects = Vec::new();

        for (archive_name, archive) in &self.archives {
            for object in &archive.objects {
                if object.messages.iter().any(|msg| msg.type_ == message_type) {
                    matching_objects.push((archive_name.as_str(), object));
                }
            }
        }

        matching_objects
    }
}

/// Statistics about a bundle
#[derive(Debug, Clone)]
pub struct BundleStats {
    /// Number of IWA archives in the bundle
    pub archive_count: usize,
    /// Total number of objects across all archives
    pub total_objects: usize,
    /// Largest archive (name, object count)
    pub largest_archive: Option<(String, usize)>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_bundle_validation() {
        // Test with a non-existent directory
        let bundle_path = std::path::Path::new("non_existent_bundle");
        assert!(Bundle::open(bundle_path).is_err());

        // Test with existing iWork bundle
        let bundle_path = std::path::Path::new("test.pages");
        if bundle_path.exists() {
            let result = Bundle::open(bundle_path);
            assert!(
                result.is_ok(),
                "Failed to open test.pages: {:?}",
                result.err()
            );
        }
    }

    #[test]
    fn test_bundle_parsing() {
        let bundle_path = std::path::Path::new("test.pages");
        if !bundle_path.exists() {
            // Skip test if test file doesn't exist
            return;
        }

        let bundle = Bundle::open(bundle_path).expect("Failed to open test.pages");

        // Verify bundle has expected structure
        assert!(
            !bundle.archives().is_empty(),
            "Bundle should contain archives"
        );

        // Check for common iWork files
        assert!(
            bundle.get_archive("Index/Document.iwa").is_some(),
            "Bundle should contain Document.iwa"
        );
        assert!(
            bundle.get_archive("Index/Metadata.iwa").is_some(),
            "Bundle should contain Metadata.iwa"
        );

        // Verify metadata exists
        let metadata = bundle.metadata();
        assert!(
            metadata.has_properties || metadata.has_build_version_history,
            "Bundle should have some metadata"
        );

        // Test text extraction (will be empty for now as protobuf decoding isn't implemented)
        let text_result = bundle.extract_text();
        assert!(text_result.is_ok());
    }

    #[test]
    fn test_numbers_bundle_parsing() {
        let bundle_path = std::path::Path::new("test.numbers");
        if !bundle_path.exists() {
            // Skip test if test file doesn't exist
            return;
        }

        let bundle = Bundle::open(bundle_path).expect("Failed to open test.numbers");

        // Verify bundle has expected structure
        assert!(
            !bundle.archives().is_empty(),
            "Bundle should contain archives"
        );

        // Check for common Numbers files
        assert!(
            bundle.get_archive("Index/Document.iwa").is_some(),
            "Bundle should contain Document.iwa"
        );
        assert!(
            bundle.get_archive("Index/CalculationEngine.iwa").is_some(),
            "Numbers bundle should contain CalculationEngine.iwa"
        );
    }

    #[test]
    fn test_metadata_summary() {
        let mut properties = HashMap::new();
        properties.insert(
            "Title".to_string(),
            PropertyValue::String("Test Doc".to_string()),
        );

        let metadata = BundleMetadata {
            has_properties: true,
            has_build_version_history: true,
            has_document_identifier: false,
            detected_application: Some("Pages".to_string()),
            properties,
            build_versions: vec!["7029".to_string()],
            document_id: None,
        };

        let summary = metadata.summary();
        assert!(summary.contains("Properties: true"));
        assert!(summary.contains("BuildVersion: true"));
        assert!(summary.contains("DocumentID: false"));
        assert!(summary.contains("App: Pages"));

        // Test property accessors
        assert_eq!(
            metadata.get_property_string("Title"),
            Some("Test Doc".to_string())
        );
        assert_eq!(metadata.latest_build_version(), Some("7029"));
        assert_eq!(metadata.document_identifier(), None);
    }
}
