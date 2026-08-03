//! iWork Bundle Structure Parser
//!
//! iWork documents are stored as bundles (directories) containing:
//! - `Index.zip`: Archive of IWA files with serialized objects
//! - `Data/`: Directory containing media assets
//! - `Metadata/`: Document metadata and properties
//! - Preview images at root level

use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use plist::Value;
use soapberry_zip::office::{ArchiveLimits, ArchiveReader};

use crate::archive::{Archive, ArchiveObject};
use crate::snappy::{SnappyLimits, SnappyStream};
use crate::zip_utils::parse_iwa_files_from_archive_with_limits;
use crate::{Error, Result};

/// Represents an iWork document bundle
#[derive(Debug, Clone)]
pub struct Bundle {
    state: Arc<BundleState>,
}

/// Resource ceilings for one parsed iWork bundle.
///
/// The profile bounds the bytes read from a filesystem path, ZIP central
/// directory metadata, every nested `Index.zip`, and each decompressed IWA
/// component. Limits can only be tightened below the hard format ceilings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BundleLimits {
    max_input_bytes: u64,
    max_entries: usize,
    max_entry_bytes: u64,
    max_total_bytes: u64,
    max_iwa_stream_bytes: usize,
}

impl BundleLimits {
    /// Hard ceiling for bytes read from one bundle file or `Index.zip`.
    pub const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for non-directory ZIP members in one archive.
    pub const MAX_ENTRIES: usize = 100_000;
    /// Hard ceiling for one declared uncompressed ZIP member.
    pub const MAX_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
    /// Hard ceiling for the declared uncompressed size of one ZIP archive.
    pub const MAX_TOTAL_BYTES: u64 = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for one decompressed IWA component.
    pub const MAX_IWA_STREAM_BYTES: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;

    /// Build checked bundle-ingress ceilings.
    pub fn new(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_total_bytes: u64,
        max_iwa_stream_bytes: usize,
    ) -> Result<Self> {
        if max_input_bytes == 0
            || max_entries == 0
            || max_entry_bytes == 0
            || max_total_bytes == 0
            || max_iwa_stream_bytes == 0
        {
            return Err(Error::InvalidFormat(
                "iWork bundle limits must be non-zero".to_owned(),
            ));
        }
        if max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork bundle input limit exceeds the {} byte hard ceiling",
                Self::MAX_INPUT_BYTES
            )));
        }
        if max_entries > Self::MAX_ENTRIES {
            return Err(Error::InvalidFormat(format!(
                "iWork bundle entry limit exceeds the {} entry hard ceiling",
                Self::MAX_ENTRIES
            )));
        }
        if max_entry_bytes > Self::MAX_ENTRY_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork bundle entry limit exceeds the {} byte hard ceiling",
                Self::MAX_ENTRY_BYTES
            )));
        }
        if max_total_bytes > Self::MAX_TOTAL_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork bundle total limit exceeds the {} byte hard ceiling",
                Self::MAX_TOTAL_BYTES
            )));
        }
        if max_iwa_stream_bytes > Self::MAX_IWA_STREAM_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork IWA stream limit exceeds the {} byte hard ceiling",
                Self::MAX_IWA_STREAM_BYTES
            )));
        }

        Ok(Self {
            max_input_bytes,
            max_entries,
            max_entry_bytes,
            max_total_bytes,
            max_iwa_stream_bytes,
        })
    }

    /// Maximum bytes read from one bundle file or nested index.
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum number of non-directory ZIP members in one archive.
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }

    /// Maximum declared uncompressed size of one ZIP member.
    pub const fn max_entry_bytes(self) -> u64 {
        self.max_entry_bytes
    }

    /// Maximum declared uncompressed size of one ZIP archive.
    pub const fn max_total_bytes(self) -> u64 {
        self.max_total_bytes
    }

    /// Maximum decompressed size of one IWA component.
    pub const fn max_iwa_stream_bytes(self) -> usize {
        self.max_iwa_stream_bytes
    }

    pub(crate) fn archive_limits(self) -> ArchiveLimits {
        ArchiveLimits {
            max_files: self.max_entries,
            max_entry_size: self.max_entry_bytes,
            max_total_size: self.max_total_bytes,
        }
    }

    pub(crate) fn snappy_limits(self) -> Result<SnappyLimits> {
        SnappyLimits::new(
            self.max_iwa_stream_bytes
                .min(SnappyStream::MAX_UNCOMPRESSED_CHUNK),
            self.max_iwa_stream_bytes,
        )
    }

    pub(crate) fn check_input_size(self, size: u64, label: &str) -> Result<()> {
        if size > self.max_input_bytes {
            return Err(Error::InvalidFormat(format!(
                "{label} is {size} bytes, exceeding the {} byte limit",
                self.max_input_bytes
            )));
        }
        Ok(())
    }
}

impl Default for BundleLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
            max_entries: Self::MAX_ENTRIES,
            max_entry_bytes: Self::MAX_ENTRY_BYTES,
            max_total_bytes: Self::MAX_TOTAL_BYTES,
            max_iwa_stream_bytes: Self::MAX_IWA_STREAM_BYTES,
        }
    }
}

#[derive(Debug)]
struct BundleState {
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
        Self::open_with_limits(path, BundleLimits::default())
    }

    /// Open an iWork bundle under caller-selected ingress ceilings.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: BundleLimits) -> Result<Self> {
        let bundle_path = path.as_ref().to_path_buf();

        let metadata = fs::symlink_metadata(&bundle_path).map_err(|error| {
            if error.kind() == std::io::ErrorKind::NotFound {
                Error::Bundle(format!("Path does not exist: {}", bundle_path.display()))
            } else {
                error.into()
            }
        })?;
        if metadata.file_type().is_symlink() {
            return Err(Error::Bundle(format!(
                "Bundle path must not be a symbolic link: {}",
                bundle_path.display()
            )));
        }
        if metadata.is_dir() {
            // Traditional bundle directory structure
            Self::open_directory_bundle(&bundle_path, limits)
        } else if metadata.is_file() {
            // Single file bundle (zip archive)
            Self::open_file_bundle(&bundle_path, limits)
        } else {
            Err(Error::Bundle(format!(
                "Bundle path is not a regular file or directory: {}",
                bundle_path.display()
            )))
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
        Self::from_bytes_with_limits(bytes, BundleLimits::default())
    }

    /// Open a single-file bundle from bytes under caller-selected limits.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let input_size = u64::try_from(bytes.len()).map_err(|_| {
            Error::InvalidFormat("iWork bundle input length does not fit u64".to_owned())
        })?;
        limits.check_input_size(input_size, "iWork bundle input")?;

        // Parse the ZIP archive directly from bytes
        let archives = Self::parse_zip_bytes(bytes, limits)?;

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

        Ok(Self::from_parts(
            std::path::PathBuf::from("<bytes>"), // Placeholder path
            archives,
            metadata,
        ))
    }

    /// Create a Bundle from raw bytes (ZIP archive).
    ///
    /// This method parses the archive and its IWA members from the supplied
    /// bytes; it does not accept or reuse a previously parsed archive owner.
    pub fn from_archive_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_archive_bytes_with_limits(bytes, BundleLimits::default())
    }

    /// Parse a ZIP archive from bytes under caller-selected ingress ceilings.
    pub fn from_archive_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let input_size = u64::try_from(bytes.len()).map_err(|_| {
            Error::InvalidFormat("iWork bundle input length does not fit u64".to_owned())
        })?;
        limits.check_input_size(input_size, "iWork bundle input")?;

        // Parse IWA files from the ZIP archive
        let archive = ArchiveReader::new_with_limits(bytes, limits.archive_limits())
            .map_err(|e| Error::Bundle(format!("Failed to open ZIP archive: {}", e)))?;
        let archives = parse_iwa_files_from_archive_with_limits(&archive, limits)?;

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

        Ok(Self::from_parts(
            std::path::PathBuf::from("<zip_archive>"), // Placeholder path
            archives,
            metadata,
        ))
    }

    /// Open a traditional directory-based bundle
    fn open_directory_bundle(bundle_path: &Path, limits: BundleLimits) -> Result<Self> {
        // Check for required bundle structure
        Self::validate_bundle_structure(bundle_path)?;

        // Parse Index.zip
        let archives = Self::parse_index_zip(bundle_path, limits)?;

        // Parse metadata
        let metadata = Self::parse_metadata(bundle_path)?;

        Ok(Self::from_parts(
            bundle_path.to_path_buf(),
            archives,
            metadata,
        ))
    }

    /// Open a single-file bundle (zip archive)
    fn open_file_bundle(bundle_path: &Path, limits: BundleLimits) -> Result<Self> {
        // Parse the zip file directly
        let archives = Self::parse_zip_bundle(bundle_path, limits)?;

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

        Ok(Self::from_parts(
            bundle_path.to_path_buf(),
            archives,
            metadata,
        ))
    }

    fn from_parts(
        bundle_path: PathBuf,
        archives: HashMap<String, Archive>,
        metadata: BundleMetadata,
    ) -> Self {
        Self {
            state: Arc::new(BundleState {
                bundle_path,
                archives,
                metadata,
            }),
        }
    }

    fn archives_in_order(&self) -> Vec<(&str, &Archive)> {
        let mut archives: Vec<_> = self
            .state
            .archives
            .iter()
            .map(|(name, archive)| (name.as_str(), archive))
            .collect();
        archives.sort_unstable_by_key(|(name, _)| *name);
        archives
    }

    /// Capture a cheap immutable snapshot that shares all parsed bundle state.
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Validate that the path contains a valid iWork bundle structure
    fn validate_bundle_structure(bundle_path: &Path) -> Result<()> {
        // Check for Index.zip
        let index_zip = bundle_path.join("Index.zip");
        ensure_regular_file(&index_zip, "Index.zip")?;

        // Check for Metadata directory (optional but common)
        let metadata_dir = bundle_path.join("Metadata");
        match fs::symlink_metadata(&metadata_dir) {
            Ok(metadata) if metadata.file_type().is_symlink() => {
                return Err(Error::Bundle(format!(
                    "Metadata path must not be a symbolic link: {}",
                    metadata_dir.display()
                )));
            },
            Ok(metadata) if !metadata.is_dir() => {
                return Err(Error::Bundle(format!(
                    "Metadata path is not a directory: {}",
                    metadata_dir.display()
                )));
            },
            Ok(_) => {},
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {},
            Err(error) => return Err(error.into()),
        }

        Ok(())
    }

    /// Parse Index.zip and extract all IWA files
    fn parse_index_zip(
        bundle_path: &Path,
        limits: BundleLimits,
    ) -> Result<HashMap<String, Archive>> {
        let index_zip_path = bundle_path.join("Index.zip");
        let data = read_bounded_file(&index_zip_path, limits, "iWork Index.zip")?;

        let archive = ArchiveReader::new_with_limits(&data, limits.archive_limits())
            .map_err(|e| Error::Bundle(format!("Failed to open Index.zip: {}", e)))?;

        parse_iwa_files_from_archive_with_limits(&archive, limits)
    }

    /// Parse a single-file bundle (zip archive) and extract all IWA files
    fn parse_zip_bundle(
        bundle_path: &Path,
        limits: BundleLimits,
    ) -> Result<HashMap<String, Archive>> {
        let data = read_bounded_file(bundle_path, limits, "iWork bundle")?;

        let archive = ArchiveReader::new_with_limits(&data, limits.archive_limits())
            .map_err(|e| Error::Bundle(format!("Failed to open bundle file: {}", e)))?;

        parse_iwa_files_from_archive_with_limits(&archive, limits)
    }

    /// Parse a ZIP archive from raw bytes and extract all IWA files
    fn parse_zip_bytes(bytes: &[u8], limits: BundleLimits) -> Result<HashMap<String, Archive>> {
        let archive = ArchiveReader::new_with_limits(bytes, limits.archive_limits())
            .map_err(|e| Error::Bundle(format!("Failed to open ZIP archive from bytes: {}", e)))?;

        parse_iwa_files_from_archive_with_limits(&archive, limits)
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
        if optional_regular_file(&properties_path, "Properties.plist")? {
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
        if optional_regular_file(&build_version_path, "BuildVersionHistory.plist")? {
            metadata.has_build_version_history = true;
            if let Ok(value) = Value::from_file(&build_version_path) {
                metadata.build_versions = Self::parse_build_versions(&value);
            }
        }

        // Read DocumentIdentifier
        let doc_id_path = metadata_dir.join("DocumentIdentifier");
        if optional_regular_file(&doc_id_path, "DocumentIdentifier")? {
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
        &self.state.archives
    }

    /// Get a specific archive by name
    pub fn get_archive(&self, name: &str) -> Option<&Archive> {
        self.state.archives.get(name)
    }

    /// Get bundle metadata
    pub fn metadata(&self) -> &BundleMetadata {
        &self.state.metadata
    }

    /// Get the bundle path
    pub fn path(&self) -> &Path {
        &self.state.bundle_path
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
        if self.state.archives.is_empty() {
            return Err(Error::Bundle(
                "Bundle contains no archives - may be corrupted or empty".to_string(),
            ));
        }

        // Verify each archive has at least one object
        let mut total_objects = 0;
        for (archive_name, archive) in self.archives_in_order() {
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
        if self.state.archives.is_empty() {
            return true;
        }

        let has_any_objects = self
            .state
            .archives
            .values()
            .any(|archive| !archive.objects.is_empty());

        !has_any_objects
    }

    /// Get bundle statistics
    pub fn stats(&self) -> BundleStats {
        let archive_count = self.state.archives.len();
        let total_objects: usize = self
            .state
            .archives
            .values()
            .map(|archive| archive.objects.len())
            .sum();

        let largest_archive = self
            .archives_in_order()
            .into_iter()
            .max_by(|(left_name, left_archive), (right_name, right_archive)| {
                left_archive
                    .objects
                    .len()
                    .cmp(&right_archive.objects.len())
                    .then_with(|| right_name.cmp(left_name))
            })
            .map(|(name, archive)| (name.to_owned(), archive.objects.len()));

        BundleStats {
            archive_count,
            total_objects,
            largest_archive,
        }
    }

    /// Extract all text content from the bundle in deterministic archive order.
    pub fn extract_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        for (_, archive) in self.archives_in_order() {
            for object in &archive.objects {
                text_parts.extend(object.extract_text());
            }
        }

        // Join all text parts with newlines
        Ok(text_parts.join("\n"))
    }

    /// Get all objects across all archives in archive-name/source-object order.
    pub fn all_objects(&self) -> Vec<(&str, &ArchiveObject)> {
        let mut objects = Vec::new();
        for (archive_name, archive) in self.archives_in_order() {
            for object in &archive.objects {
                objects.push((archive_name, object));
            }
        }
        objects
    }

    /// Find objects by message type in archive-name/source-object order.
    pub fn find_objects_by_type(&self, message_type: u32) -> Vec<(&str, &ArchiveObject)> {
        let mut matching_objects = Vec::new();

        for (archive_name, archive) in self.archives_in_order() {
            for object in &archive.objects {
                if object.messages.iter().any(|msg| msg.type_ == message_type) {
                    matching_objects.push((archive_name, object));
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
    Ok(match crate::detect::path(bundle_path)? {
        Some(crate::detect::Format::Pages) => "Pages",
        Some(crate::detect::Format::Keynote) => "Keynote",
        Some(crate::detect::Format::Numbers) => "Numbers",
        None => "Unknown",
    }
    .to_owned())
}

fn ensure_regular_file(path: &Path, label: &str) -> Result<()> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(Error::Bundle(format!(
            "{label} must not be a symbolic link: {}",
            path.display()
        ))),
        Ok(metadata) if metadata.is_file() => Ok(()),
        Ok(_) => Err(Error::Bundle(format!(
            "{label} is not a regular file: {}",
            path.display()
        ))),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Err(Error::Bundle(format!(
            "{label} not found in bundle: {}",
            path.display()
        ))),
        Err(error) => Err(error.into()),
    }
}

fn read_bounded_file(path: &Path, limits: BundleLimits, label: &str) -> Result<Vec<u8>> {
    let size = fs::metadata(path)?.len();
    limits.check_input_size(size, label)?;
    fs::read(path).map_err(Error::Io)
}

fn optional_regular_file(path: &Path, label: &str) -> Result<bool> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(Error::Bundle(format!(
            "{label} must not be a symbolic link: {}",
            path.display()
        ))),
        Ok(metadata) if metadata.is_file() => Ok(true),
        Ok(_) => Err(Error::Bundle(format!(
            "{label} is not a regular file: {}",
            path.display()
        ))),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(false),
        Err(error) => Err(error.into()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
    use std::fs;

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

    #[test]
    fn bundle_limits_are_checked_and_bound_nested_iwa_decompression() -> crate::Result<()> {
        let limits = BundleLimits::new(7, 11, 13, 17, 19)?;
        assert_eq!(limits.max_input_bytes(), 7);
        assert_eq!(limits.max_entries(), 11);
        assert_eq!(limits.max_entry_bytes(), 13);
        assert_eq!(limits.max_total_bytes(), 17);
        assert_eq!(limits.max_iwa_stream_bytes(), 19);
        assert!(BundleLimits::new(0, 1, 1, 1, 1).is_err());
        assert!(BundleLimits::new(1, 1, 1, 1, 0).is_err());
        assert!(BundleLimits::new(BundleLimits::MAX_INPUT_BYTES + 1, 1, 1, 1, 1).is_err());

        let compressed = SnappyStream::compress(&[0_u8; 64])?;
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored("Index/Document.iwa", &compressed)
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let tight_stream = BundleLimits::new(
            bytes.len() as u64,
            BundleLimits::MAX_ENTRIES,
            BundleLimits::MAX_ENTRY_BYTES,
            BundleLimits::MAX_TOTAL_BYTES,
            8,
        )?;
        let error = Bundle::from_bytes_with_limits(&bytes, tight_stream).unwrap_err();
        assert!(error.to_string().contains("Snappy block expands"));

        let tight_input = BundleLimits::new(
            (bytes.len() - 1) as u64,
            BundleLimits::MAX_ENTRIES,
            BundleLimits::MAX_ENTRY_BYTES,
            BundleLimits::MAX_TOTAL_BYTES,
            BundleLimits::MAX_IWA_STREAM_BYTES,
        )?;
        let error = Bundle::from_bytes_with_limits(&bytes, tight_input).unwrap_err();
        assert!(error.to_string().contains("iWork bundle input"));

        let file = tempfile::NamedTempFile::new()?;
        fs::write(file.path(), &bytes)?;
        let error = Bundle::open_with_limits(file.path(), tight_input).unwrap_err();
        assert!(error.to_string().contains("iWork bundle is"));
        Ok(())
    }

    #[test]
    fn application_detection_uses_content_not_filename_extension() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let file = temp.path().join("looks-like.numbers");
        fs::write(&file, b"not an iWork archive")?;
        assert_eq!(detect_application_type(&file)?, "Unknown");

        let directory = temp.path().join("looks-like.pages");
        fs::create_dir(&directory)?;
        fs::write(directory.join("index.apxl"), [])?;
        assert_eq!(detect_application_type(&directory)?, "Keynote");
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn bundle_ingress_rejects_symbolic_links() -> std::io::Result<()> {
        use std::os::unix::fs::symlink;

        let temp = tempfile::tempdir()?;
        let target = temp.path().join("target");
        fs::create_dir(&target)?;
        fs::write(target.join("Index.zip"), b"not a zip")?;

        let index_link = target.join("Index-link.zip");
        symlink(target.join("Index.zip"), &index_link)?;
        assert!(ensure_regular_file(&index_link, "Index.zip").is_err());

        let bundle_link = temp.path().join("bundle.pages");
        symlink(&target, &bundle_link)?;
        let error = Bundle::open(&bundle_link).unwrap_err();
        assert!(error.to_string().contains("symbolic link"));
        Ok(())
    }

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn snapshots_share_immutable_bundle_state() {
        let bundle = Bundle::from_parts(
            PathBuf::from("<test>"),
            HashMap::new(),
            BundleMetadata::default(),
        );
        let snapshot = bundle.snapshot();

        assert!(Arc::ptr_eq(&bundle.state, &snapshot.state));
    }

    #[test]
    fn bundles_are_send_and_sync() {
        assert_send_sync::<Bundle>();
    }

    #[test]
    fn bundle_queries_are_deterministic_across_archive_hash_map_order() -> Result<()> {
        let archive_a = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 7,
                        data: Vec::new(),
                    }],
                )?,
                ArchiveObject::new(
                    3,
                    vec![RawMessage {
                        type_: 8,
                        data: Vec::new(),
                    }],
                )?,
            ],
        };
        let archive_z = Archive {
            objects: vec![
                ArchiveObject::new(
                    2,
                    vec![RawMessage {
                        type_: 7,
                        data: Vec::new(),
                    }],
                )?,
                ArchiveObject::new(
                    4,
                    vec![RawMessage {
                        type_: 9,
                        data: Vec::new(),
                    }],
                )?,
            ],
        };
        let bundle = Bundle::from_parts(
            PathBuf::from("<test>"),
            HashMap::from([
                ("Index/Z.iwa".to_owned(), archive_z),
                ("Index/A.iwa".to_owned(), archive_a),
            ]),
            BundleMetadata::default(),
        );

        let object_order: Vec<_> = bundle
            .all_objects()
            .into_iter()
            .map(|(archive, object)| (archive, object.archive_info.identifier))
            .collect();
        assert_eq!(
            object_order,
            vec![
                ("Index/A.iwa", Some(1)),
                ("Index/A.iwa", Some(3)),
                ("Index/Z.iwa", Some(2)),
                ("Index/Z.iwa", Some(4)),
            ]
        );

        let matching_order: Vec<_> = bundle
            .find_objects_by_type(7)
            .into_iter()
            .map(|(archive, object)| (archive, object.archive_info.identifier))
            .collect();
        assert_eq!(
            matching_order,
            vec![("Index/A.iwa", Some(1)), ("Index/Z.iwa", Some(2))]
        );
        assert_eq!(
            bundle.stats().largest_archive,
            Some(("Index/A.iwa".to_owned(), 2))
        );
        Ok(())
    }
}
