//! iWork Bundle Structure Parser
//!
//! iWork documents are stored as bundles (directories) containing:
//! - `Index.zip`: Archive of IWA files with serialized objects
//! - `Data/`: Directory containing media assets
//! - `Metadata/`: Document metadata and properties
//! - Preview images at root level

use std::collections::HashMap;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use plist::Value;
use zip::ZipArchive;

use crate::iwa::archive::{Archive, ArchiveObject};
use crate::iwa::zip_utils::parse_iwa_files_from_zip;
use crate::iwa::{Error, Result};

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
    /// Open an iWork bundle from a path (directory or zip file)
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bundle_path = path.as_ref().to_path_buf();

        if bundle_path.is_dir() {
            // Traditional bundle directory structure
            Self::open_directory_bundle(&bundle_path)
        } else if bundle_path.is_file() {
            // Single file bundle (zip archive)
            Self::open_file_bundle(&bundle_path)
        } else {
            Err(Error::Bundle("Path does not exist".to_string()))
        }
    }

    /// Open an iWork bundle from raw bytes (single-file zip archive)
    ///
    /// This function can parse iWork documents that are stored as ZIP archives
    /// directly from memory, without requiring file system access.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Raw bytes of the iWork ZIP archive
    ///
    /// # Returns
    ///
    /// * `Result<Self>` - Parsed bundle on success, error on failure
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::Bundle;
    /// use std::fs;
    ///
    /// let data = fs::read("document.pages")?;
    /// let bundle = Bundle::from_bytes(&data)?;
    /// println!("Archives: {}", bundle.archives().len());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        // Parse the ZIP archive directly from bytes
        let archives = Self::parse_zip_bytes(bytes)?;

        // For single-file bundles, metadata is typically embedded
        let metadata = BundleMetadata {
            has_properties: true, // Assume it has properties
            has_build_version_history: true,
            has_document_identifier: true,
            detected_application: None,
            properties: HashMap::new(),
            build_versions: Vec::new(),
            document_id: None,
        };

        Ok(Bundle {
            bundle_path: std::path::PathBuf::from("<bytes>"), // Placeholder path
            archives,
            metadata,
        })
    }

    /// Open a traditional directory-based bundle
    fn open_directory_bundle(bundle_path: &Path) -> Result<Self> {
        // Check for required bundle structure
        Self::validate_bundle_structure(bundle_path)?;

        // Parse Index.zip
        let archives = Self::parse_index_zip(bundle_path)?;

        // Parse metadata
        let metadata = Self::parse_metadata(bundle_path)?;

        Ok(Bundle {
            bundle_path: bundle_path.to_path_buf(),
            archives,
            metadata,
        })
    }

    /// Open a single-file bundle (zip archive)
    fn open_file_bundle(bundle_path: &Path) -> Result<Self> {
        // Parse the zip file directly
        let archives = Self::parse_zip_bundle(bundle_path)?;

        // For single-file bundles, metadata is typically embedded
        let metadata = BundleMetadata {
            has_properties: true, // Assume it has properties
            has_build_version_history: true,
            has_document_identifier: true,
            detected_application: None,
            properties: HashMap::new(),
            build_versions: Vec::new(),
            document_id: None,
        };

        Ok(Bundle {
            bundle_path: bundle_path.to_path_buf(),
            archives,
            metadata,
        })
    }

    /// Validate that the path contains a valid iWork bundle structure
    fn validate_bundle_structure(bundle_path: &Path) -> Result<()> {
        // Check for Index.zip
        let index_zip = bundle_path.join("Index.zip");
        if !index_zip.exists() {
            return Err(Error::Bundle("Index.zip not found in bundle".to_string()));
        }

        // Check for Metadata directory (optional but common)
        let metadata_dir = bundle_path.join("Metadata");
        if !metadata_dir.exists() || !metadata_dir.is_dir() {
            // Some bundles might not have metadata, continue anyway
        }

        Ok(())
    }

    /// Parse Index.zip and extract all IWA files
    fn parse_index_zip(bundle_path: &Path) -> Result<HashMap<String, Archive>> {
        let index_zip_path = bundle_path.join("Index.zip");
        let file = fs::File::open(&index_zip_path).map_err(Error::Io)?;

        let mut zip_archive = ZipArchive::new(file)
            .map_err(|e| Error::Bundle(format!("Failed to open Index.zip: {}", e)))?;

        parse_iwa_files_from_zip(&mut zip_archive)
    }

    /// Parse a single-file bundle (zip archive) and extract all IWA files
    fn parse_zip_bundle(bundle_path: &Path) -> Result<HashMap<String, Archive>> {
        let file = fs::File::open(bundle_path).map_err(Error::Io)?;

        let mut zip_archive = ZipArchive::new(file)
            .map_err(|e| Error::Bundle(format!("Failed to open bundle file: {}", e)))?;

        parse_iwa_files_from_zip(&mut zip_archive)
    }

    /// Parse a ZIP archive from raw bytes and extract all IWA files
    fn parse_zip_bytes(bytes: &[u8]) -> Result<HashMap<String, Archive>> {
        let cursor = Cursor::new(bytes);
        let mut zip_archive = ZipArchive::new(cursor)
            .map_err(|e| Error::Bundle(format!("Failed to open ZIP archive from bytes: {}", e)))?;

        parse_iwa_files_from_zip(&mut zip_archive)
    }

    /// Parse metadata from Metadata/ directory
    fn parse_metadata(bundle_path: &Path) -> Result<BundleMetadata> {
        let metadata_dir = bundle_path.join("Metadata");
        let mut metadata = BundleMetadata::default();

        if !metadata_dir.exists() {
            return Ok(metadata);
        }

        // Parse Properties.plist
        let properties_path = metadata_dir.join("Properties.plist");
        if properties_path.exists() {
            metadata.has_properties = true;
            if let Ok(value) = Value::from_file(&properties_path) {
                metadata.properties = Self::parse_plist_value(&value);

                // Try to detect application from properties
                if let Some(PropertyValue::String(app_name)) =
                    metadata.properties.get("Application")
                {
                    metadata.detected_application = Some(app_name.clone());
                }
            }
        }

        // Parse BuildVersionHistory.plist
        let build_version_path = metadata_dir.join("BuildVersionHistory.plist");
        if build_version_path.exists() {
            metadata.has_build_version_history = true;
            if let Ok(value) = Value::from_file(&build_version_path) {
                metadata.build_versions = Self::parse_build_versions(&value);
            }
        }

        // Read DocumentIdentifier
        let doc_id_path = metadata_dir.join("DocumentIdentifier");
        if doc_id_path.exists() {
            metadata.has_document_identifier = true;
            if let Ok(id) = fs::read_to_string(&doc_id_path) {
                metadata.document_id = Some(id.trim().to_string());
            }
        }

        Ok(metadata)
    }

    /// Parse a plist Value into our PropertyValue structure
    fn parse_plist_value(value: &Value) -> HashMap<String, PropertyValue> {
        let mut result = HashMap::new();

        if let Value::Dictionary(dict) = value {
            for (key, val) in dict {
                result.insert(key.clone(), Self::convert_plist_value(val));
            }
        }

        result
    }

    /// Convert a plist Value to PropertyValue
    fn convert_plist_value(value: &Value) -> PropertyValue {
        match value {
            Value::String(s) => PropertyValue::String(s.clone()),
            Value::Integer(i) => PropertyValue::Integer(i.as_signed().unwrap_or(0)),
            Value::Real(r) => PropertyValue::Real(*r),
            Value::Boolean(b) => PropertyValue::Boolean(*b),
            Value::Date(d) => PropertyValue::Date(format!("{:?}", d)),
            Value::Array(arr) => {
                PropertyValue::Array(arr.iter().map(Self::convert_plist_value).collect())
            },
            Value::Dictionary(dict) => {
                let mut map = HashMap::new();
                for (k, v) in dict {
                    map.insert(k.clone(), Self::convert_plist_value(v));
                }
                PropertyValue::Dictionary(map)
            },
            Value::Data(_) => PropertyValue::String("<binary data>".to_string()),
            _ => PropertyValue::String("<unknown>".to_string()),
        }
    }

    /// Parse build versions from BuildVersionHistory.plist
    fn parse_build_versions(value: &Value) -> Vec<String> {
        let mut versions = Vec::new();

        if let Value::Array(arr) = value {
            for item in arr {
                if let Value::String(version) = item {
                    versions.push(version.clone());
                } else if let Value::Dictionary(dict) = item {
                    // BuildVersionHistory might be an array of dictionaries with version info
                    if let Some(Value::String(version)) = dict.get("Version") {
                        versions.push(version.clone());
                    } else if let Some(Value::String(build)) = dict.get("Build") {
                        versions.push(build.clone());
                    }
                }
            }
        }

        versions
    }

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

/// Metadata associated with an iWork bundle
#[derive(Debug, Clone, Default)]
pub struct BundleMetadata {
    /// Whether Properties.plist exists
    pub has_properties: bool,
    /// Whether BuildVersionHistory.plist exists
    pub has_build_version_history: bool,
    /// Whether DocumentIdentifier exists
    pub has_document_identifier: bool,
    /// Application type detected from the bundle
    pub detected_application: Option<String>,
    /// Parsed properties from Properties.plist
    pub properties: HashMap<String, PropertyValue>,
    /// Build version history
    pub build_versions: Vec<String>,
    /// Document identifier
    pub document_id: Option<String>,
}

/// Represents a property value from plist
#[derive(Debug, Clone)]
pub enum PropertyValue {
    /// String value
    String(String),
    /// Integer value
    Integer(i64),
    /// Real/float value
    Real(f64),
    /// Boolean value
    Boolean(bool),
    /// Date value
    Date(String),
    /// Array of values
    Array(Vec<PropertyValue>),
    /// Dictionary of values
    Dictionary(HashMap<String, PropertyValue>),
}

impl BundleMetadata {
    /// Get a summary of the metadata
    pub fn summary(&self) -> String {
        format!(
            "Properties: {}, BuildVersion: {}, DocumentID: {}, App: {}",
            self.has_properties,
            self.has_build_version_history,
            self.has_document_identifier,
            self.detected_application.as_deref().unwrap_or("unknown")
        )
    }

    /// Get a property value as a string
    pub fn get_property_string(&self, key: &str) -> Option<String> {
        match self.properties.get(key)? {
            PropertyValue::String(s) => Some(s.clone()),
            PropertyValue::Integer(i) => Some(i.to_string()),
            PropertyValue::Real(r) => Some(r.to_string()),
            PropertyValue::Boolean(b) => Some(b.to_string()),
            PropertyValue::Date(d) => Some(d.clone()),
            _ => None,
        }
    }

    /// Get a property value as an integer
    pub fn get_property_int(&self, key: &str) -> Option<i64> {
        match self.properties.get(key)? {
            PropertyValue::Integer(i) => Some(*i),
            _ => None,
        }
    }

    /// Get a property value as a boolean
    pub fn get_property_bool(&self, key: &str) -> Option<bool> {
        match self.properties.get(key)? {
            PropertyValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the document identifier
    pub fn document_identifier(&self) -> Option<&str> {
        self.document_id.as_deref()
    }

    /// Get the build versions
    pub fn build_version_history(&self) -> &[String] {
        &self.build_versions
    }

    /// Get the latest build version
    pub fn latest_build_version(&self) -> Option<&str> {
        self.build_versions.last().map(|s| s.as_str())
    }
}

/// Detect the application type from a bundle path
pub fn detect_application_type<P: AsRef<Path>>(bundle_path: P) -> Result<String> {
    let path = bundle_path.as_ref();

    // Check file extension or directory structure
    if let Some(extension) = path.extension() {
        match extension.to_str() {
            Some("pages") => return Ok("Pages".to_string()),
            Some("key") => return Ok("Keynote".to_string()),
            Some("numbers") => return Ok("Numbers".to_string()),
            _ => {},
        }
    }

    // Check for application-specific files in Index.zip
    if path.is_dir() {
        let index_zip = path.join("Index.zip");
        if index_zip.exists() {
            // This would require opening the zip and checking for app-specific files
            // For now, return "Unknown"
        }
    }

    Ok("Unknown".to_string())
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
