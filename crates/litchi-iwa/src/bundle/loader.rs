//! Loading directory and zip-backed iWork bundles.

use std::collections::HashMap;
use std::fs;
use std::path::Path;

use plist::Value;
use soapberry_zip::office::ArchiveReader;

use super::{Bundle, BundleMetadata, PropertyValue};
use crate::archive::Archive;
use crate::zip_utils::parse_iwa_files_from_archive;
use crate::{Error, Result};

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
    /// use litchi_iwa::Bundle;
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

    /// Create a Bundle from raw bytes (ZIP archive).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been parsed during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: &[u8]) -> Result<Self> {
        // Parse IWA files from the ZIP archive
        let archive = ArchiveReader::new(bytes)
            .map_err(|e| Error::Bundle(format!("Failed to open ZIP archive: {}", e)))?;
        let archives = parse_iwa_files_from_archive(&archive)?;

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
            bundle_path: std::path::PathBuf::from("<zip_archive>"), // Placeholder path
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
        let data = fs::read(&index_zip_path).map_err(Error::Io)?;

        let archive = ArchiveReader::new(&data)
            .map_err(|e| Error::Bundle(format!("Failed to open Index.zip: {}", e)))?;

        parse_iwa_files_from_archive(&archive)
    }

    /// Parse a single-file bundle (zip archive) and extract all IWA files
    fn parse_zip_bundle(bundle_path: &Path) -> Result<HashMap<String, Archive>> {
        let data = fs::read(bundle_path).map_err(Error::Io)?;

        let archive = ArchiveReader::new(&data)
            .map_err(|e| Error::Bundle(format!("Failed to open bundle file: {}", e)))?;

        parse_iwa_files_from_archive(&archive)
    }

    /// Parse a ZIP archive from raw bytes and extract all IWA files
    fn parse_zip_bytes(bytes: &[u8]) -> Result<HashMap<String, Archive>> {
        let archive = ArchiveReader::new(bytes)
            .map_err(|e| Error::Bundle(format!("Failed to open ZIP archive from bytes: {}", e)))?;

        parse_iwa_files_from_archive(&archive)
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
}
