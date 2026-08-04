//! iWork Bundle Structure Parser
//!
//! iWork documents are stored as bundles (directories) containing:
//! - `Index.zip`: Archive of IWA files with serialized objects
//! - `Data/`: Directory containing media assets
//! - `Metadata/`: Document metadata and properties
//! - Preview images at root level

use std::collections::{BTreeMap, HashMap, HashSet};
use std::fmt;
use std::fs;
use std::io::{Cursor, Read};
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
    ///
    /// The vector is sorted once at ingress. Keeping the names beside their
    /// archive values avoids exposing the parser's lookup map and avoids a
    /// second owned name catalog solely for deterministic traversal.
    archives: Vec<(String, Archive)>,
    /// Metadata from Metadata/ directory
    metadata: BundleMetadata,
}

/// Severity attached to one deterministic bundle validation finding.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum BundleValidationSeverity {
    /// The bundle can still be consumed, but the producer emitted an unusual
    /// or incomplete structure.
    Warning,
    /// The bundle violates an invariant required by the safe facade.
    Error,
}

impl fmt::Display for BundleValidationSeverity {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Warning => "warning",
            Self::Error => "error",
        })
    }
}

/// Stable code identifying one bundle validation rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum BundleValidationCode {
    /// No IWA archives were found in the bundle.
    EmptyBundle,
    /// An individual IWA archive contains no objects.
    EmptyArchive,
    /// No objects were found across any archive.
    NoObjects,
    /// An object has no archive identifier.
    MissingObjectIdentifier,
    /// An object uses the protobuf null identifier.
    NullObjectIdentifier,
    /// An object identifier occurs more than once in the bundle.
    DuplicateObjectIdentifier,
    /// Archive metadata and payload counts differ.
    MessageInfoCountMismatch,
    /// A message's declared length differs from its payload length.
    MessageLengthMismatch,
}

impl fmt::Display for BundleValidationCode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::EmptyBundle => "empty-bundle",
            Self::EmptyArchive => "empty-archive",
            Self::NoObjects => "no-objects",
            Self::MissingObjectIdentifier => "missing-object-identifier",
            Self::NullObjectIdentifier => "null-object-identifier",
            Self::DuplicateObjectIdentifier => "duplicate-object-identifier",
            Self::MessageInfoCountMismatch => "message-info-count-mismatch",
            Self::MessageLengthMismatch => "message-length-mismatch",
        })
    }
}

/// One structured, source-located bundle validation finding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BundleValidationIssue {
    severity: BundleValidationSeverity,
    code: BundleValidationCode,
    archive_name: Option<String>,
    object_id: Option<u64>,
}

impl BundleValidationIssue {
    /// Return the finding severity.
    pub const fn severity(&self) -> BundleValidationSeverity {
        self.severity
    }

    /// Return the stable validation rule code.
    pub const fn code(&self) -> BundleValidationCode {
        self.code
    }

    /// Return the deterministic archive location, when applicable.
    pub fn archive_name(&self) -> Option<&str> {
        self.archive_name.as_deref()
    }

    /// Return the native object identifier, when applicable.
    pub const fn object_id(&self) -> Option<u64> {
        self.object_id
    }
}

impl fmt::Display for BundleValidationIssue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.severity, self.code)?;
        if let Some(archive_name) = &self.archive_name {
            write!(formatter, " archive={archive_name}")?;
        }
        if let Some(object_id) = self.object_id {
            write!(formatter, " object={object_id}")?;
        }
        Ok(())
    }
}

/// Bounded, deterministic validation output for one parsed iWork bundle.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct BundleValidationReport {
    issues: Vec<BundleValidationIssue>,
    truncated: bool,
}

impl BundleValidationReport {
    /// Maximum number of findings retained by one report.
    pub const MAX_ISSUES: usize = 4096;

    /// Return findings in stable archive/object/source order.
    pub fn issues(&self) -> &[BundleValidationIssue] {
        &self.issues
    }

    /// Return whether the scan completed without hitting the diagnostic cap.
    pub const fn is_complete(&self) -> bool {
        !self.truncated
    }

    /// Return whether no errors were found and the scan completed.
    pub fn is_valid(&self) -> bool {
        !self.truncated
            && self
                .issues
                .iter()
                .all(|issue| !matches!(issue.severity, BundleValidationSeverity::Error))
    }

    /// Return whether at least one error finding was retained.
    pub fn has_errors(&self) -> bool {
        self.issues
            .iter()
            .any(|issue| issue.severity == BundleValidationSeverity::Error)
    }

    /// Count retained findings of one severity.
    pub fn count(&self, severity: BundleValidationSeverity) -> usize {
        self.issues
            .iter()
            .filter(|issue| issue.severity == severity)
            .count()
    }

    /// Convert this report to the crate's compatibility result API.
    pub fn as_result(&self) -> Result<()> {
        if self.is_valid() {
            Ok(())
        } else {
            Err(Error::Bundle(self.to_string()))
        }
    }

    fn push(
        &mut self,
        severity: BundleValidationSeverity,
        code: BundleValidationCode,
        archive_name: Option<&str>,
        object_id: Option<u64>,
    ) {
        if self.issues.len() >= Self::MAX_ISSUES {
            self.truncated = true;
            return;
        }
        self.issues.push(BundleValidationIssue {
            severity,
            code,
            archive_name: archive_name.map(str::to_owned),
            object_id,
        });
    }
}

impl fmt::Display for BundleValidationReport {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.is_valid() && self.issues.is_empty() {
            return formatter.write_str("bundle validation passed");
        }
        write!(
            formatter,
            "bundle validation found {} issue(s)",
            self.issues.len()
        )?;
        if self.truncated {
            formatter.write_str(" (report truncated)")?;
        }
        for issue in &self.issues {
            write!(formatter, "; {issue}")?;
        }
        Ok(())
    }
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
    /// println!("Archives: {}", bundle.iter_archives().count());
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
            properties: PropertyMap::default(),
            build_versions: Vec::new().into_boxed_slice(),
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
            properties: PropertyMap::default(),
            build_versions: Vec::new().into_boxed_slice(),
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
        let metadata = Self::parse_metadata(bundle_path, limits)?;

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
            properties: PropertyMap::default(),
            build_versions: Vec::new().into_boxed_slice(),
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
        let mut archives: Vec<_> = archives.into_iter().collect();
        archives.sort_unstable_by(|(left, _), (right, _)| left.cmp(right));
        Self {
            state: Arc::new(BundleState {
                bundle_path,
                archives,
                metadata,
            }),
        }
    }

    /// Enumerate archives in deterministic lexicographic name order.
    ///
    /// The iterator borrows the immutable catalog and performs no allocation.
    /// This is the preferred collection view for selectors, diagnostics, and
    /// other consumers that must not depend on parser storage details.
    pub fn iter_archives(&self) -> impl Iterator<Item = (&str, &Archive)> {
        self.state
            .archives
            .iter()
            .map(|(name, archive)| (name.as_str(), archive))
    }

    /// Collect the deterministic archive view into a caller-owned vector.
    ///
    /// Use [`Self::iter_archives`] when a borrowed traversal is sufficient.
    pub fn archives_in_order(&self) -> Vec<(&str, &Archive)> {
        self.iter_archives().collect()
    }

    /// Enumerate every object in deterministic archive-name/source-object order.
    ///
    /// The iterator borrows the immutable archive catalog and performs no
    /// allocation. Use [`Self::all_objects`] when an owned collection is
    /// required by a compatibility or batch API.
    pub fn iter_objects(&self) -> impl Iterator<Item = (&str, &ArchiveObject)> {
        self.iter_archives().flat_map(|(archive_name, archive)| {
            archive
                .objects
                .iter()
                .map(move |object| (archive_name, object))
        })
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
    fn parse_metadata(bundle_path: &Path, limits: BundleLimits) -> Result<BundleMetadata> {
        let metadata_dir = bundle_path.join("Metadata");
        let mut metadata = BundleMetadata::default();

        if !metadata_dir.exists() {
            return Ok(metadata);
        }

        // Parse Properties.plist
        let properties_path = metadata_dir.join("Properties.plist");
        if optional_regular_file(&properties_path, "Properties.plist")? {
            metadata.has_properties = true;
            let bytes = read_bounded_file(&properties_path, limits, "Properties.plist")?;
            let value = Self::parse_plist_bytes(&bytes, "Properties.plist")?;
            metadata.properties = Self::parse_plist_value(&value, "Properties.plist")?;

            // Try to detect application from properties.
            if let Some(app_name) = metadata.properties.get("Application") {
                let PropertyValue::String(app_name) = app_name else {
                    return Err(Error::InvalidFormat(
                        "Properties.plist Application must be a string".to_owned(),
                    ));
                };
                metadata.detected_application = Some(app_name.clone());
            }
        }

        // Parse BuildVersionHistory.plist
        let build_version_path = metadata_dir.join("BuildVersionHistory.plist");
        if optional_regular_file(&build_version_path, "BuildVersionHistory.plist")? {
            metadata.has_build_version_history = true;
            let bytes =
                read_bounded_file(&build_version_path, limits, "BuildVersionHistory.plist")?;
            let value = Self::parse_plist_bytes(&bytes, "BuildVersionHistory.plist")?;
            metadata.build_versions =
                Self::parse_build_versions(&value, "BuildVersionHistory.plist")?.into_boxed_slice();
        }

        // Read DocumentIdentifier
        let doc_id_path = metadata_dir.join("DocumentIdentifier");
        if optional_regular_file(&doc_id_path, "DocumentIdentifier")? {
            metadata.has_document_identifier = true;
            let bytes = read_bounded_file(&doc_id_path, limits, "DocumentIdentifier")?;
            let id = String::from_utf8(bytes).map_err(|error| {
                Error::InvalidFormat(format!("DocumentIdentifier is not valid UTF-8: {error}"))
            })?;
            let id = id.trim();
            if id.is_empty() {
                return Err(Error::InvalidFormat(
                    "DocumentIdentifier must not be empty".to_owned(),
                ));
            }
            metadata.document_id = Some(id.to_owned());
        }

        Ok(metadata)
    }

    /// Parse a plist Value into our PropertyValue structure
    fn parse_plist_bytes(bytes: &[u8], label: &str) -> Result<Value> {
        Value::from_reader(Cursor::new(bytes))
            .map_err(|error| Error::InvalidFormat(format!("failed to parse {label}: {error}")))
    }

    /// Parse a plist Value into our PropertyValue structure.
    fn parse_plist_value(value: &Value, label: &str) -> Result<PropertyMap> {
        match value {
            Value::Dictionary(dict) => dict
                .iter()
                .map(|(key, value)| {
                    let context = format!("{label}.{key}");
                    Ok((key.clone(), Self::convert_plist_value(value, &context)?))
                })
                .collect(),
            _ => Err(Error::InvalidFormat(format!(
                "{label} must contain a dictionary at its root"
            ))),
        }
    }

    /// Convert a plist Value to PropertyValue.
    fn convert_plist_value(value: &Value, context: &str) -> Result<PropertyValue> {
        match value {
            Value::String(s) => Ok(PropertyValue::String(s.clone())),
            Value::Integer(i) => i.as_signed().map(PropertyValue::Integer).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "{context} contains an unsigned integer outside the supported i64 range"
                ))
            }),
            Value::Real(r) => Ok(PropertyValue::Real(*r)),
            Value::Boolean(b) => Ok(PropertyValue::Boolean(*b)),
            Value::Date(d) => Ok(PropertyValue::Date(format!("{d:?}"))),
            Value::Array(arr) => arr
                .iter()
                .enumerate()
                .map(|(index, value)| {
                    Self::convert_plist_value(value, &format!("{context}[{index}]"))
                })
                .collect::<Result<Vec<_>>>()
                .map(PropertyValue::Array),
            Value::Dictionary(dict) => dict
                .iter()
                .map(|(key, value)| {
                    let nested_context = format!("{context}.{key}");
                    Ok((
                        key.clone(),
                        Self::convert_plist_value(value, &nested_context)?,
                    ))
                })
                .collect::<Result<PropertyMap>>()
                .map(PropertyValue::Dictionary),
            Value::Data(_) => Err(Error::InvalidFormat(format!(
                "{context} contains unsupported binary data"
            ))),
            Value::Uid(_) => Err(Error::InvalidFormat(format!(
                "{context} contains an unsupported UID"
            ))),
            _ => Err(Error::InvalidFormat(format!(
                "{context} contains an unsupported plist value"
            ))),
        }
    }

    /// Parse build versions from BuildVersionHistory.plist.
    fn parse_build_versions(value: &Value, label: &str) -> Result<Vec<String>> {
        let Value::Array(arr) = value else {
            return Err(Error::InvalidFormat(format!(
                "{label} must contain an array at its root"
            )));
        };

        arr.iter()
            .enumerate()
            .map(|(index, item)| match item {
                Value::String(version) => Ok(version.clone()),
                Value::Dictionary(dict) => {
                    let version = dict
                        .get("Version")
                        .or_else(|| dict.get("Build"))
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{label}[{index}] dictionary has neither Version nor Build"
                            ))
                        })?;
                    match version {
                        Value::String(version) => Ok(version.clone()),
                        _ => Err(Error::InvalidFormat(format!(
                            "{label}[{index}] Version/Build must be a string"
                        ))),
                    }
                },
                _ => Err(Error::InvalidFormat(format!(
                    "{label}[{index}] must be a string or dictionary"
                ))),
            })
            .collect()
    }

    /// Get a specific archive by name
    pub fn get_archive(&self, name: &str) -> Option<&Archive> {
        let position = self
            .state
            .archives
            .binary_search_by(|(archive_name, _)| archive_name.as_str().cmp(name))
            .ok()?;
        Some(&self.state.archives[position].1)
    }

    /// Get bundle metadata
    pub fn metadata(&self) -> &BundleMetadata {
        &self.state.metadata
    }

    /// Get the bundle path
    pub fn path(&self) -> &Path {
        &self.state.bundle_path
    }

    /// Return a bounded, deterministic validation report without mutating the
    /// bundle or emitting diagnostics to the process-wide stderr stream.
    pub fn validation_report(&self) -> BundleValidationReport {
        let mut report = BundleValidationReport::default();
        if self.state.archives.is_empty() {
            report.push(
                BundleValidationSeverity::Error,
                BundleValidationCode::EmptyBundle,
                None,
                None,
            );
            return report;
        }

        let mut identifiers = HashSet::new();
        let mut total_objects = 0usize;
        for (archive_name, archive) in self.iter_archives() {
            if archive.objects.is_empty() {
                report.push(
                    BundleValidationSeverity::Warning,
                    BundleValidationCode::EmptyArchive,
                    Some(archive_name),
                    None,
                );
            }

            total_objects = total_objects.saturating_add(archive.objects.len());
            for object in &archive.objects {
                let object_id = object.archive_info.identifier;
                match object_id {
                    None => report.push(
                        BundleValidationSeverity::Error,
                        BundleValidationCode::MissingObjectIdentifier,
                        Some(archive_name),
                        None,
                    ),
                    Some(0) => report.push(
                        BundleValidationSeverity::Error,
                        BundleValidationCode::NullObjectIdentifier,
                        Some(archive_name),
                        Some(0),
                    ),
                    Some(identifier) => {
                        if !identifiers.insert(identifier) {
                            report.push(
                                BundleValidationSeverity::Error,
                                BundleValidationCode::DuplicateObjectIdentifier,
                                Some(archive_name),
                                Some(identifier),
                            );
                        }
                    },
                }

                if object.archive_info.message_infos.len() != object.messages.len() {
                    report.push(
                        BundleValidationSeverity::Error,
                        BundleValidationCode::MessageInfoCountMismatch,
                        Some(archive_name),
                        object_id,
                    );
                }
                for (message_info, message) in object
                    .archive_info
                    .message_infos
                    .iter()
                    .zip(&object.messages)
                {
                    if usize::try_from(message_info.length).ok() != Some(message.data.len()) {
                        report.push(
                            BundleValidationSeverity::Error,
                            BundleValidationCode::MessageLengthMismatch,
                            Some(archive_name),
                            object_id,
                        );
                    }
                }
            }
        }

        if total_objects == 0 {
            report.push(
                BundleValidationSeverity::Error,
                BundleValidationCode::NoObjects,
                None,
                None,
            );
        }
        report
    }

    /// Validate the bundle structure and integrity without side effects.
    pub fn validate(&self) -> Result<()> {
        self.validation_report().as_result()
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
            .iter()
            .any(|(_, archive)| !archive.objects.is_empty());

        !has_any_objects
    }

    /// Get bundle statistics
    pub fn stats(&self) -> BundleStats {
        let archive_count = self.state.archives.len();
        let total_objects: usize = self
            .state
            .archives
            .iter()
            .map(|(_, archive)| archive.objects.len())
            .sum();

        let largest_archive = self
            .iter_archives()
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

        for (_, archive) in self.iter_archives() {
            for object in &archive.objects {
                text_parts.extend(object.extract_text());
            }
        }

        // Join all text parts with newlines
        Ok(text_parts.join("\n"))
    }

    /// Get all objects across all archives in archive-name/source-object order.
    pub fn all_objects(&self) -> Vec<(&str, &ArchiveObject)> {
        self.iter_objects().collect()
    }

    /// Find objects by message type in archive-name/source-object order.
    pub fn find_objects_by_type(&self, message_type: u32) -> Vec<(&str, &ArchiveObject)> {
        self.iter_objects()
            .filter(|(_, object)| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == message_type)
            })
            .collect()
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

/// An immutable, deterministic property catalog parsed from an iWork plist.
///
/// The map keeps its storage private so callers use validated lookup and
/// borrowed traversal rather than depending on the parser's hash-table
/// representation. Entries are ordered lexicographically by key.
#[derive(Debug, Clone, Default)]
pub struct PropertyMap {
    entries: Box<[(String, PropertyValue)]>,
}

impl PropertyMap {
    fn from_entries(entries: impl IntoIterator<Item = (String, PropertyValue)>) -> Self {
        let mut ordered = BTreeMap::new();
        for (key, value) in entries {
            ordered.insert(key, value);
        }
        Self {
            entries: ordered.into_iter().collect(),
        }
    }

    /// Return the value associated with `key`, if present.
    pub fn get(&self, key: &str) -> Option<&PropertyValue> {
        self.entries
            .binary_search_by(|(entry_key, _)| entry_key.as_str().cmp(key))
            .ok()
            .map(|index| &self.entries[index].1)
    }

    /// Iterate over properties in deterministic key order.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &PropertyValue)> + '_ {
        self.entries
            .iter()
            .map(|(key, value)| (key.as_str(), value))
    }

    /// Return the number of properties in the catalog.
    #[inline]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether the catalog contains no properties.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

impl FromIterator<(String, PropertyValue)> for PropertyMap {
    fn from_iter<T: IntoIterator<Item = (String, PropertyValue)>>(iter: T) -> Self {
        Self::from_entries(iter)
    }
}

/// Metadata associated with an iWork bundle.
#[derive(Debug, Clone, Default)]
pub struct BundleMetadata {
    has_properties: bool,
    has_build_version_history: bool,
    has_document_identifier: bool,
    detected_application: Option<String>,
    properties: PropertyMap,
    build_versions: Box<[String]>,
    document_id: Option<String>,
}

/// Represents a property value from plist.
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
    Dictionary(PropertyMap),
}

impl BundleMetadata {
    /// Return whether `Properties.plist` was present.
    pub const fn has_properties(&self) -> bool {
        self.has_properties
    }

    /// Return whether `BuildVersionHistory.plist` was present.
    pub const fn has_build_version_history(&self) -> bool {
        self.has_build_version_history
    }

    /// Return whether `DocumentIdentifier` was present.
    pub const fn has_document_identifier(&self) -> bool {
        self.has_document_identifier
    }

    /// Return the detected source application, if metadata identified one.
    pub fn detected_application(&self) -> Option<&str> {
        self.detected_application.as_deref()
    }

    /// Borrow the immutable property catalog.
    pub const fn properties(&self) -> &PropertyMap {
        &self.properties
    }

    /// Look up one parsed property without exposing the backing container.
    pub fn property(&self, key: &str) -> Option<&PropertyValue> {
        self.properties.get(key)
    }

    /// Get a summary of the metadata
    pub fn summary(&self) -> String {
        format!(
            "Properties: {}, BuildVersion: {}, DocumentID: {}, App: {}",
            self.has_properties(),
            self.has_build_version_history(),
            self.has_document_identifier(),
            self.detected_application().unwrap_or("unknown")
        )
    }

    /// Get a property value as a string
    pub fn get_property_string(&self, key: &str) -> Option<String> {
        match self.property(key)? {
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
        match self.property(key)? {
            PropertyValue::Integer(i) => Some(*i),
            _ => None,
        }
    }

    /// Get a property value as a boolean
    pub fn get_property_bool(&self, key: &str) -> Option<bool> {
        match self.property(key)? {
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
    let file = fs::File::open(path)?;
    let declared_size = file.metadata()?.len();
    limits.check_input_size(declared_size, label)?;

    let max_read = limits
        .max_input_bytes()
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat(format!("{label} size limit overflow")))?;
    let mut bytes = Vec::new();
    file.take(max_read).read_to_end(&mut bytes)?;
    let actual_size = u64::try_from(bytes.len())
        .map_err(|_| Error::InvalidFormat(format!("{label} size does not fit u64")))?;
    limits.check_input_size(actual_size, label)?;
    Ok(bytes)
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
            bundle.iter_archives().next().is_some(),
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
            bundle.iter_archives().next().is_some(),
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
        let properties = PropertyMap::from_entries([(
            "Title".to_string(),
            PropertyValue::String("Test Doc".to_string()),
        )]);

        let metadata = BundleMetadata {
            has_properties: true,
            has_build_version_history: true,
            has_document_identifier: false,
            detected_application: Some("Pages".to_string()),
            properties,
            build_versions: vec!["7029".to_string()].into_boxed_slice(),
            document_id: None,
        };

        let summary = metadata.summary();
        assert!(summary.contains("Properties: true"));
        assert!(summary.contains("BuildVersion: true"));
        assert!(summary.contains("DocumentID: false"));
        assert!(summary.contains("App: Pages"));
        assert!(metadata.has_properties());
        assert!(metadata.has_build_version_history());
        assert!(!metadata.has_document_identifier());
        assert_eq!(metadata.detected_application(), Some("Pages"));
        assert_eq!(metadata.properties().len(), 1);
        assert!(matches!(
            metadata.property("Title"),
            Some(PropertyValue::String(value)) if value == "Test Doc"
        ));

        // Test property accessors
        assert_eq!(
            metadata.get_property_string("Title"),
            Some("Test Doc".to_string())
        );
        assert_eq!(metadata.latest_build_version(), Some("7029"));
        assert_eq!(metadata.document_identifier(), None);
    }

    #[test]
    fn property_map_is_sorted_and_duplicate_keys_are_deterministic() {
        let properties: PropertyMap = [
            (
                "zeta".to_string(),
                PropertyValue::String("first".to_string()),
            ),
            ("alpha".to_string(), PropertyValue::Integer(7)),
            ("zeta".to_string(), PropertyValue::Boolean(true)),
        ]
        .into_iter()
        .collect();

        assert_eq!(properties.len(), 2);
        assert_eq!(
            properties.iter().map(|(key, _)| key).collect::<Vec<_>>(),
            vec!["alpha", "zeta"]
        );
        assert!(matches!(
            properties.get("zeta"),
            Some(PropertyValue::Boolean(true))
        ));
        assert!(properties.get("missing").is_none());
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
    fn metadata_ingress_preserves_valid_values() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let metadata_dir = temp.path().join("Metadata");
        fs::create_dir(&metadata_dir)?;

        let properties = plist::Dictionary::from_iter([
            ("Application".to_owned(), Value::String("Pages".to_owned())),
            ("Revision".to_owned(), Value::Integer(42_i64.into())),
        ]);
        write_test_plist(
            &metadata_dir.join("Properties.plist"),
            &Value::Dictionary(properties),
        )?;

        let history = Value::Array(vec![
            Value::String("14.4.1".to_owned()),
            Value::Dictionary(plist::Dictionary::from_iter([(
                "Build".to_owned(),
                Value::String("7029".to_owned()),
            )])),
        ]);
        write_test_plist(&metadata_dir.join("BuildVersionHistory.plist"), &history)?;
        fs::write(metadata_dir.join("DocumentIdentifier"), b"document-id\n")?;

        let metadata = Bundle::parse_metadata(temp.path(), BundleLimits::default())?;
        assert_eq!(metadata.detected_application(), Some("Pages"));
        assert_eq!(metadata.get_property_int("Revision"), Some(42));
        assert_eq!(metadata.build_version_history(), ["14.4.1", "7029"]);
        assert_eq!(metadata.document_identifier(), Some("document-id"));
        Ok(())
    }

    #[test]
    fn metadata_ingress_rejects_malformed_plist() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let metadata_dir = temp.path().join("Metadata");
        fs::create_dir(&metadata_dir)?;
        fs::write(metadata_dir.join("Properties.plist"), b"not a plist")?;

        let result = Bundle::parse_metadata(temp.path(), BundleLimits::default());
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("malformed plist was accepted".to_owned()))?;
        assert!(error.to_string().contains("Properties.plist"));
        Ok(())
    }

    #[test]
    fn metadata_ingress_rejects_semantically_invalid_history() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let metadata_dir = temp.path().join("Metadata");
        fs::create_dir(&metadata_dir)?;
        write_test_plist(
            &metadata_dir.join("BuildVersionHistory.plist"),
            &Value::Dictionary(plist::Dictionary::new()),
        )?;

        let result = Bundle::parse_metadata(temp.path(), BundleLimits::default());
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("invalid history was accepted".to_owned()))?;
        assert!(error.to_string().contains("BuildVersionHistory.plist"));
        assert!(error.to_string().contains("array"));
        Ok(())
    }

    #[test]
    fn metadata_ingress_rejects_invalid_document_identifier_utf8() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let metadata_dir = temp.path().join("Metadata");
        fs::create_dir(&metadata_dir)?;
        fs::write(metadata_dir.join("DocumentIdentifier"), [0xff_u8, 0xfe])?;

        let result = Bundle::parse_metadata(temp.path(), BundleLimits::default());
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("invalid UTF-8 was accepted".to_owned()))?;
        assert!(error.to_string().contains("DocumentIdentifier"));
        assert!(error.to_string().contains("UTF-8"));
        Ok(())
    }

    #[test]
    fn metadata_ingress_rejects_unsigned_integer_outside_property_range() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let metadata_dir = temp.path().join("Metadata");
        fs::create_dir(&metadata_dir)?;
        let properties = plist::Dictionary::from_iter([(
            "TooLarge".to_owned(),
            Value::Integer(u64::MAX.into()),
        )]);
        write_test_plist(
            &metadata_dir.join("Properties.plist"),
            &Value::Dictionary(properties),
        )?;

        let result = Bundle::parse_metadata(temp.path(), BundleLimits::default());
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("oversized integer was accepted".to_owned()))?;
        assert!(error.to_string().contains("TooLarge"));
        assert!(
            error
                .to_string()
                .contains("outside the supported i64 range")
        );
        Ok(())
    }

    #[test]
    fn metadata_ingress_bounds_each_metadata_file() -> crate::Result<()> {
        let temp = tempfile::tempdir()?;
        let metadata_dir = temp.path().join("Metadata");
        fs::create_dir(&metadata_dir)?;
        fs::write(metadata_dir.join("DocumentIdentifier"), [b'x'; 32])?;
        let limits = BundleLimits::new(
            8,
            BundleLimits::MAX_ENTRIES,
            BundleLimits::MAX_ENTRY_BYTES,
            BundleLimits::MAX_TOTAL_BYTES,
            BundleLimits::MAX_IWA_STREAM_BYTES,
        )?;

        let result = Bundle::parse_metadata(temp.path(), limits);
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("oversized metadata was accepted".to_owned()))?;
        assert!(error.to_string().contains("DocumentIdentifier"));
        assert!(error.to_string().contains("8 byte limit"));
        Ok(())
    }

    fn write_test_plist(path: &Path, value: &Value) -> crate::Result<()> {
        let mut bytes = Vec::new();
        value.to_writer_binary(&mut bytes).map_err(|error| {
            Error::InvalidFormat(format!("failed to encode test plist: {error}"))
        })?;
        fs::write(path, bytes)?;
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

        let borrowed_object_order: Vec<_> = bundle
            .iter_objects()
            .map(|(archive, object)| (archive, object.archive_info.identifier))
            .collect();
        assert_eq!(borrowed_object_order, object_order);

        let archive_order: Vec<_> = bundle
            .archives_in_order()
            .into_iter()
            .map(|(name, _)| name)
            .collect();
        assert_eq!(archive_order, vec!["Index/A.iwa", "Index/Z.iwa"]);

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

    #[test]
    fn validation_report_is_structured_and_deterministic() -> Result<()> {
        let empty = Archive {
            objects: Vec::new(),
        };
        let mut length_mismatch = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 7,
                data: vec![1],
            }],
        )?;
        length_mismatch.archive_info.message_infos[0].length = 9;
        let mut duplicate = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 8,
                data: Vec::new(),
            }],
        )?;
        duplicate.archive_info.message_infos.clear();

        let bundle = Bundle::from_parts(
            PathBuf::from("<test>"),
            HashMap::from([
                (
                    "Index/C.iwa".to_owned(),
                    Archive {
                        objects: vec![duplicate],
                    },
                ),
                ("Index/A.iwa".to_owned(), empty),
                (
                    "Index/B.iwa".to_owned(),
                    Archive {
                        objects: vec![length_mismatch],
                    },
                ),
            ]),
            BundleMetadata::default(),
        );

        let report = bundle.validation_report();
        let findings: Vec<_> = report
            .issues()
            .iter()
            .map(|issue| (issue.code(), issue.archive_name(), issue.object_id()))
            .collect();
        assert_eq!(
            findings,
            vec![
                (
                    BundleValidationCode::EmptyArchive,
                    Some("Index/A.iwa"),
                    None
                ),
                (
                    BundleValidationCode::MessageLengthMismatch,
                    Some("Index/B.iwa"),
                    Some(1)
                ),
                (
                    BundleValidationCode::DuplicateObjectIdentifier,
                    Some("Index/C.iwa"),
                    Some(1)
                ),
                (
                    BundleValidationCode::MessageInfoCountMismatch,
                    Some("Index/C.iwa"),
                    Some(1)
                ),
            ]
        );
        assert_eq!(report.count(BundleValidationSeverity::Warning), 1);
        assert_eq!(report.count(BundleValidationSeverity::Error), 3);
        assert!(report.is_complete());
        assert!(!report.is_valid());
        assert!(bundle.validate().is_err());
        Ok(())
    }

    #[test]
    fn validation_report_is_bounded() {
        let archives = (0..=BundleValidationReport::MAX_ISSUES)
            .map(|index| {
                (
                    format!("Index/{index:04}.iwa"),
                    Archive {
                        objects: Vec::new(),
                    },
                )
            })
            .collect();
        let bundle =
            Bundle::from_parts(PathBuf::from("<test>"), archives, BundleMetadata::default());

        let report = bundle.validation_report();
        assert_eq!(report.issues().len(), BundleValidationReport::MAX_ISSUES);
        assert!(!report.is_complete());
        assert!(!report.is_valid());
        assert!(report.as_result().is_err());
    }
}
