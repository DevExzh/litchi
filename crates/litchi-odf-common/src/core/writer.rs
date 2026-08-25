//! ODF package writing functionality.
//!
//! This module provides utilities for creating and writing ODF files as ZIP archives,
//! including generating manifests and proper file structure.
//!
//! Uses soapberry-zip for high-performance ZIP writing.

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64_STANDARD};
use chrono::NaiveDate;
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use soapberry_zip::office::{
    StreamingArchiveFailure, StreamingArchiveWriter, StreamingLimitExceeded, StreamingLimitResource,
};
use soapberry_zip::{ErrorKind as ZipErrorKind, LimitResource as ZipLimitResource};
use std::collections::HashSet;
use std::fmt;
use std::io::{self, Read, Write};
use zeroize::Zeroizing;

const MAX_MANIFEST_TEXT_BYTES: usize = 1024;
const MANIFEST_PATH: &str = "META-INF/manifest.xml";
const LOEXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

fn ensure_source_manifest_rewritable(source: &OwnedPackage) -> Result<()> {
    if source.has_zip_encrypted_entries() {
        return Err(Error::Unsupported(
            "ZIP-encrypted ODF entries cannot be copied by the package writer".to_string(),
        ));
    }
    let has_canonical_manifest = source.has_file(MANIFEST_PATH)?;
    let has_legacy_manifest = source.has_file("manifest.xml")?;
    if has_canonical_manifest && has_legacy_manifest {
        return Err(Error::Unsupported(
            "ODF source contains both canonical and legacy manifest members".to_string(),
        ));
    }
    let package = source.package()?;
    let manifest_bytes = source
        .get_file(MANIFEST_PATH)
        .or_else(|_| source.get_file("manifest.xml"))?;
    ensure_supported_manifest_metadata(&manifest_bytes)?;
    let manifest = package.manifest();
    if manifest
        .entries
        .values()
        .any(|entry| entry.size.is_some() && entry.encryption.is_none())
    {
        return Err(Error::Unsupported(
            "ODF unencrypted manifest:size metadata cannot be preserved by the package writer"
                .to_string(),
        ));
    }
    Ok(())
}

fn ensure_supported_manifest_metadata(bytes: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid source manifest XML: {error}"))
            })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let ResolveResult::Bound(Namespace(uri)) = namespace else {
                    return Err(Error::Unsupported(
                        "ODF manifest contains unsupported non-manifest metadata".to_string(),
                    ));
                };
                if uri != MANIFEST_NAMESPACE {
                    return Err(Error::Unsupported(
                        "ODF manifest contains unsupported non-manifest metadata".to_string(),
                    ));
                }
                let element_name = element.local_name();
                if !matches!(
                    element_name.as_ref(),
                    b"manifest"
                        | b"file-entry"
                        | b"encryption-data"
                        | b"algorithm"
                        | b"start-key-generation"
                        | b"key-derivation"
                ) {
                    return Err(Error::Unsupported(
                        "ODF manifest contains unsupported element metadata".to_string(),
                    ));
                }
                let mut full_path = None;
                let mut has_version = false;
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid source manifest attribute: {error}"))
                    })?;
                    if attribute.key.as_ref() == b"xmlns"
                        || attribute.key.as_ref().starts_with(b"xmlns:")
                    {
                        continue;
                    }
                    let (attribute_namespace, local) =
                        reader.resolver().resolve_attribute(attribute.key);
                    let local = local.as_ref();
                    let allowed = match attribute_namespace {
                        ResolveResult::Bound(Namespace(attribute_uri))
                            if attribute_uri == MANIFEST_NAMESPACE =>
                        {
                            let allowed = match element_name.as_ref() {
                                b"manifest" => local == b"version",
                                b"file-entry" => {
                                    matches!(
                                        local,
                                        b"full-path" | b"media-type" | b"size" | b"version"
                                    )
                                },
                                b"encryption-data" => {
                                    matches!(local, b"checksum-type" | b"checksum")
                                },
                                b"algorithm" => {
                                    matches!(local, b"algorithm-name" | b"initialisation-vector")
                                },
                                b"start-key-generation" => {
                                    matches!(local, b"start-key-generation-name" | b"key-size")
                                },
                                b"key-derivation" => matches!(
                                    local,
                                    b"key-derivation-name"
                                        | b"salt"
                                        | b"iteration-count"
                                        | b"key-size"
                                        | b"argon2-iterations"
                                        | b"argon2-memory"
                                        | b"argon2-lanes"
                                ),
                                _ => false,
                            };
                            if local == b"full-path" {
                                full_path = Some(
                                    attribute
                                        .decoded_and_normalized_value(
                                            XmlVersion::Implicit1_0,
                                            reader.decoder(),
                                        )
                                        .map_err(|error| {
                                            Error::InvalidFormat(format!(
                                                "invalid source manifest path: {error}"
                                            ))
                                        })?
                                        .into_owned(),
                                );
                            }
                            if local == b"version" {
                                has_version = true;
                            }
                            allowed
                        },
                        ResolveResult::Bound(Namespace(attribute_uri))
                            if attribute_uri == LOEXT_NAMESPACE
                                && element_name.as_ref() == b"key-derivation" =>
                        {
                            matches!(
                                local,
                                b"argon2-iterations" | b"argon2-memory" | b"argon2-lanes"
                            )
                        },
                        _ => false,
                    };
                    if !allowed {
                        return Err(Error::Unsupported(
                            "ODF manifest contains unsupported metadata".to_string(),
                        ));
                    }
                }
                if element_name.as_ref() == b"file-entry"
                    && has_version
                    && full_path.as_deref() != Some("/")
                {
                    return Err(Error::Unsupported(
                        "ODF manifest version metadata is only supported on the root entry"
                            .to_string(),
                    ));
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(Error::Unsupported(
                    "ODF manifest contains unsupported DTD or entity metadata".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

/// Finite limits applied to sequential ODF package publication.
///
/// These limits bound ZIP transport bytes and the writer's retained manifest
/// bookkeeping. They do not make arbitrary caller-owned readers seekable or
/// reserve the complete payload in memory.
pub use soapberry_zip::office::StreamingArchiveLimits as PackageWriterLimits;

/// Compression methods supported by opaque streamed ODF package members.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PackageCompression {
    /// Keep the member uncompressed.
    Stored,
    /// Compress the member with ZIP Deflate.
    Deflated,
}

/// A finite ZIP resource bounded by sequential ODF publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PackageWriterLimitResource {
    /// Number of file members accepted by the package.
    FileCount,
    /// UTF-8 bytes in one member name.
    MemberNameBytes,
    /// Aggregate variable central-directory metadata bytes.
    MetadataBytes,
    /// Compressed bytes in one member.
    CompressedSize,
    /// Uncompressed bytes in one member.
    EntrySize,
    /// Aggregate uncompressed bytes across members.
    TotalSize,
    /// Complete bytes accepted by the output sink.
    OutputBytes,
    /// A future transport resource not known by this crate yet.
    Other,
}

impl fmt::Display for PackageWriterLimitResource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::FileCount => "file count",
            Self::MemberNameBytes => "member name bytes",
            Self::MetadataBytes => "metadata bytes",
            Self::CompressedSize => "compressed member size",
            Self::EntrySize => "uncompressed member size",
            Self::TotalSize => "total uncompressed size",
            Self::OutputBytes => "output bytes",
            Self::Other => "other streaming resource",
        })
    }
}

/// Typed attribution for a finite ODF package publication ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PackageWriterLimitExceeded {
    resource: PackageWriterLimitResource,
    actual: u64,
    maximum: u64,
}

impl PackageWriterLimitExceeded {
    /// The bounded resource that caused the failure.
    #[must_use]
    pub const fn resource(self) -> PackageWriterLimitResource {
        self.resource
    }

    /// The attempted or observed value.
    #[must_use]
    pub const fn actual(self) -> u64 {
        self.actual
    }

    /// The configured finite ceiling.
    #[must_use]
    pub const fn maximum(self) -> u64 {
        self.maximum
    }
}

impl fmt::Display for PackageWriterLimitExceeded {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "ODF package {resource} limit exceeded: attempted {actual}, maximum {maximum}",
            resource = self.resource,
            actual = self.actual,
            maximum = self.maximum,
        )
    }
}

/// A bounded ODF package publication failure.
///
/// Existing in-memory writer methods continue to return [`litchi_core::Error`].
/// The sequential reader and sink APIs use this error so callers can inspect
/// whether a caller-owned sink contains an incomplete prefix and how many
/// bytes it accepted.
#[derive(Debug)]
#[non_exhaustive]
pub enum PackageWriterError {
    /// A preflight, source, or archive error without a typed streaming prefix.
    Core(Error),
    /// The original ZIP/archive failure, retained for callers that need its
    /// concrete source chain instead of only a compatibility string.
    Archive(soapberry_zip::Error),
    /// A finalization failure retaining the complete typed archive report.
    ArchiveFailure(StreamingArchiveFailure),
    /// The caller-owned sink accepted a prefix before publication failed.
    IncompleteOutput {
        /// Bytes accepted by the output sink before failure.
        written: u64,
        /// The underlying failure.
        source: Box<Self>,
    },
    /// A low-level streaming byte ceiling was reached after publication began.
    LimitExceeded {
        /// Bytes accepted by the output sink before failure.
        written: u64,
        /// The bounded streaming resource that caused the failure.
        limit: PackageWriterLimitExceeded,
        /// The underlying failure.
        source: Box<Self>,
    },
}

impl PackageWriterError {
    /// Return the number of output bytes accepted before this failure, if any.
    #[must_use]
    pub const fn written(&self) -> Option<u64> {
        match self {
            Self::Core(_) | Self::Archive(_) | Self::ArchiveFailure(_) => None,
            Self::IncompleteOutput { written, .. } | Self::LimitExceeded { written, .. } => {
                Some(*written)
            },
        }
    }

    /// Return the low-level byte limit attribution, if this was a limit error.
    #[must_use]
    pub const fn limit(&self) -> Option<PackageWriterLimitExceeded> {
        match self {
            Self::LimitExceeded { limit, .. } => Some(*limit),
            Self::Core(_)
            | Self::Archive(_)
            | Self::ArchiveFailure(_)
            | Self::IncompleteOutput { .. } => None,
        }
    }

    /// Consume this error and recover the compatibility core error.
    #[must_use]
    pub fn into_core_error(self) -> Error {
        match self {
            Self::Core(error) => error,
            Self::Archive(error) => Error::ZipError(error.to_string()),
            Self::ArchiveFailure(failure) => Error::ZipError(failure.to_string()),
            Self::IncompleteOutput { source, .. } | Self::LimitExceeded { source, .. } => {
                source.into_core_error()
            },
        }
    }
}

impl From<Error> for PackageWriterError {
    fn from(error: Error) -> Self {
        Self::Core(error)
    }
}

impl From<PackageWriterError> for Error {
    fn from(error: PackageWriterError) -> Self {
        error.into_core_error()
    }
}

impl fmt::Display for PackageWriterError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Core(error) => error.fmt(formatter),
            Self::Archive(error) => error.fmt(formatter),
            Self::ArchiveFailure(failure) => failure.fmt(formatter),
            Self::IncompleteOutput { written, source } => {
                write!(
                    formatter,
                    "incomplete ODF output after {written} byte(s): {source}"
                )
            },
            Self::LimitExceeded {
                written,
                limit,
                source,
            } => write!(
                formatter,
                "ODF streaming limit {limit} after {written} byte(s): {source}"
            ),
        }
    }
}

impl std::error::Error for PackageWriterError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Core(error) => Some(error),
            Self::Archive(error) => match error.kind() {
                ZipErrorKind::IO(source) | ZipErrorKind::Io(source) => Some(source),
                _ => Some(error),
            },
            Self::ArchiveFailure(failure) => match failure.error().kind() {
                ZipErrorKind::IO(source) | ZipErrorKind::Io(source) => Some(source),
                _ => Some(failure),
            },
            Self::IncompleteOutput { source, .. } | Self::LimitExceeded { source, .. } => {
                Some(source)
            },
        }
    }
}

/// Result returned by sequential ODF package publication methods.
pub type PackageWriterResult<T> = std::result::Result<T, PackageWriterError>;

use super::encryption::{Profile, encrypt_entry};
use super::manifest::{
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm,
    ManifestKeyDerivation, ManifestStartKeyGeneration,
};
use super::package::{OwnedPackage, is_signature_owner_path};
use super::xml_splice::XmlSplicePublication;
use crate::package::validate_manifest_path;

const MANIFEST_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

fn normalized_manifest_path(path: &str) -> &str {
    if path == "/" {
        "/"
    } else {
        path.strip_prefix('/').unwrap_or(path)
    }
}

fn invalid_odf_datetime(value: &str, field: &str) -> Error {
    Error::InvalidFormat(format!("invalid ODF {field} date-time '{value}'"))
}

fn parse_ascii_digits(value: &[u8]) -> Option<u32> {
    if value.is_empty() || !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    value.iter().try_fold(0_u32, |number, digit| {
        number
            .checked_mul(10)
            .and_then(|number| number.checked_add(u32::from(digit - b'0')))
    })
}

/// Validate the canonical writer date-time profile emitted by fresh ODF metadata.
///
/// The shared reader codec intentionally accepts a wider compatibility space.
/// New writer input is narrower: it requires a four-digit year in
/// `0001..=9999`, `YYYY-MM-DDThh:mm:ss`, optional decimal seconds, and either
/// no timezone, `Z`, or a numeric offset in the inclusive range
/// `-14:00..=+14:00`. Broader XSD lexical forms such as negative or five-digit
/// years and `24:00:00` are intentionally refused by this writer profile.
fn validate_canonical_odf_datetime(value: &str, field: &str) -> Result<()> {
    let bytes = value.as_bytes();
    if bytes.len() < 19
        || bytes.get(4) != Some(&b'-')
        || bytes.get(7) != Some(&b'-')
        || bytes.get(10) != Some(&b'T')
        || bytes.get(13) != Some(&b':')
        || bytes.get(16) != Some(&b':')
    {
        return Err(invalid_odf_datetime(value, field));
    }

    let year = parse_ascii_digits(&bytes[0..4]);
    let month = parse_ascii_digits(&bytes[5..7]);
    let day = parse_ascii_digits(&bytes[8..10]);
    let hour = parse_ascii_digits(&bytes[11..13]);
    let minute = parse_ascii_digits(&bytes[14..16]);
    let second = parse_ascii_digits(&bytes[17..19]);
    let (Some(year), Some(month), Some(day), Some(hour), Some(minute), Some(second)) =
        (year, month, day, hour, minute, second)
    else {
        return Err(invalid_odf_datetime(value, field));
    };

    if year == 0
        || NaiveDate::from_ymd_opt(year as i32, month, day).is_none()
        || hour > 23
        || minute > 59
        || second > 59
    {
        return Err(invalid_odf_datetime(value, field));
    }

    let mut index = 19;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        let fraction_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if fraction_start == index {
            return Err(invalid_odf_datetime(value, field));
        }
    }

    if index < bytes.len() {
        match bytes[index] {
            b'Z' => index += 1,
            b'+' | b'-' => {
                if bytes.len() != index + 6 || bytes.get(index + 3) != Some(&b':') {
                    return Err(invalid_odf_datetime(value, field));
                }
                let offset_hour = parse_ascii_digits(&bytes[index + 1..index + 3]);
                let offset_minute = parse_ascii_digits(&bytes[index + 4..index + 6]);
                let (Some(offset_hour), Some(offset_minute)) = (offset_hour, offset_minute) else {
                    return Err(invalid_odf_datetime(value, field));
                };
                if offset_hour > 14
                    || offset_minute > 59
                    || (offset_hour == 14 && offset_minute != 0)
                {
                    return Err(invalid_odf_datetime(value, field));
                }
                index += 6;
            },
            _ => return Err(invalid_odf_datetime(value, field)),
        }
    }

    if index == bytes.len() {
        Ok(())
    } else {
        Err(invalid_odf_datetime(value, field))
    }
}

/// Builder for creating ODF packages (ZIP archives)
///
/// This struct helps create valid ODF files by managing the ZIP archive structure,
/// manifest, and required files.
///
/// # Examples
///
/// ```ignore
/// # use litchi_odf::core::PackageWriter;
/// # use litchi_core::Result;
/// # fn example() -> Result<()> {
/// let mut writer = PackageWriter::new();
/// writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
/// writer.add_file("content.xml", b"<office:document-content>...</office:document-content>")?;
/// writer.add_file("styles.xml", b"<office:document-styles>...</office:document-styles>")?;
/// writer.add_file("meta.xml", b"<office:document-meta>...</office:document-meta>")?;
///
/// let bytes = writer.finish()?;
/// std::fs::write("document.odt", bytes)?;
/// # Ok(())
/// # }
/// ```
#[allow(
    clippy::module_name_repetitions,
    reason = "`PackageWriter` is the established public ODF package writer name."
)]
pub struct PackageWriter<W: Write = io::Cursor<Vec<u8>>> {
    // The locally retained limits also bound generated ODF manifest XML and
    // manifest-only entries; ZIP transport metadata alone is insufficient.
    zip_writer: StreamingArchiveWriter<W>,
    limits: PackageWriterLimits,
    mimetype: Option<String>,
    manifest_entries: Vec<ManifestEntry>,
    manifest_paths: HashSet<String>,
    member_paths: HashSet<String>,
    manifest_metadata_bytes: u64,
    archive_entry_count: usize,
    manifest_version: String,
    preserved_manifest: Option<PreservedManifest>,
    wrote_any_entry: bool,
    wrote_mimetype: bool,
    wrote_payload_entry: bool,
    encryption: Option<WriterEncryption>,
    document_signer: Option<crate::signature::DocumentSigner>,
}

struct WriterEncryption {
    profile: Profile,
    password: Zeroizing<String>,
}

#[derive(Clone, Copy)]
enum PayloadOrigin {
    AuthoredOrChanged,
    CheckedSplice,
    ExactSource,
}

/// Entry in the ODF manifest
#[derive(Debug, Clone)]
struct ManifestEntry {
    full_path: String,
    media_type: String,
    size: Option<u64>,
    encryption: Option<ManifestEncryption>,
}

struct PreservedManifest {
    bytes: Vec<u8>,
    entries: std::collections::HashMap<String, ManifestEntry>,
    physical_paths: HashSet<String>,
}

/// Helper to create standard ODF directory structure.
pub struct Structure;

/// A bounded, fallibly growing in-memory package sink.
///
/// Every write checks the configured byte limit before an amortized capacity
/// increase capped at that limit. This is intended for package publication
/// paths that must never first materialize an oversized archive.
pub struct BoundedBytes {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedBytes {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for BoundedBytes {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("ODF bounded package output size overflow"))?;
        if length > self.limit {
            return Err(io::Error::other(
                "ODF bounded package output exceeds its limit",
            ));
        }
        if length > self.bytes.capacity() {
            let doubled_capacity = self.bytes.capacity().saturating_mul(2).max(1);
            let desired_capacity = doubled_capacity.max(length).min(self.limit);
            let additional = desired_capacity.saturating_sub(self.bytes.capacity());
            self.bytes.try_reserve_exact(additional).map_err(|error| {
                io::Error::other(format!(
                    "ODF bounded package output allocation failed: {error}"
                ))
            })?;
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl PackageWriter<io::Cursor<Vec<u8>>> {
    /// Create a new package writer that writes to memory
    #[must_use]
    pub fn new() -> Self {
        Self {
            zip_writer: StreamingArchiveWriter::new(),
            limits: PackageWriterLimits::default(),
            mimetype: None,
            manifest_entries: Vec::new(),
            manifest_paths: HashSet::new(),
            member_paths: HashSet::new(),
            manifest_metadata_bytes: 0,
            archive_entry_count: 0,
            manifest_version: "1.3".to_string(),
            preserved_manifest: None,
            wrote_any_entry: false,
            wrote_mimetype: false,
            wrote_payload_entry: false,
            encryption: None,
            document_signer: None,
        }
    }

    /// Create a writer whose archive bytes are bounded before materialization.
    #[must_use]
    pub fn new_bounded(limit: usize) -> PackageWriter<BoundedBytes> {
        PackageWriter::with_writer(BoundedBytes::new(limit))
    }

    /// Create an in-memory writer with explicit finite ZIP transport limits.
    #[must_use]
    pub fn new_with_limits(limits: PackageWriterLimits) -> Self {
        Self::with_writer_and_limits(io::Cursor::new(Vec::new()), limits)
    }
}

impl Default for PackageWriter<io::Cursor<Vec<u8>>> {
    fn default() -> Self {
        Self::new()
    }
}

impl<W: Write> PackageWriter<W> {
    /// Create a package writer over an arbitrary output sink.
    pub fn with_writer(writer: W) -> Self {
        Self::with_writer_and_limits(writer, PackageWriterLimits::default())
    }

    /// Create a package writer over an arbitrary sink with explicit finite
    /// ZIP transport limits.
    pub fn with_writer_and_limits(writer: W, limits: PackageWriterLimits) -> Self {
        Self {
            zip_writer: StreamingArchiveWriter::with_writer_and_limits(writer, limits),
            limits,
            mimetype: None,
            manifest_entries: Vec::new(),
            manifest_paths: HashSet::new(),
            member_paths: HashSet::new(),
            manifest_metadata_bytes: 0,
            archive_entry_count: 0,
            manifest_version: "1.3".to_string(),
            preserved_manifest: None,
            wrote_any_entry: false,
            wrote_mimetype: false,
            wrote_payload_entry: false,
            encryption: None,
            document_signer: None,
        }
    }

    /// Configure a document signature generated after every other package entry is final.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn set_document_signer(&mut self, signer: crate::signature::DocumentSigner) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF signing must be configured before payload entries".to_string(),
            ));
        }
        self.document_signer = Some(signer);
        Ok(())
    }

    /// Clear document signing before any payload entry is written.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn clear_document_signer(&mut self) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF signing cannot be changed after payload entries".to_string(),
            ));
        }
        self.document_signer = None;
        Ok(())
    }

    /// Configure encryption for subsequently written payload entries.
    ///
    /// This may be called after `mimetype`, but not after any payload entry was emitted.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn set_encryption(&mut self, password: impl Into<String>, profile: Profile) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF encryption must be configured before payload entries".to_string(),
            ));
        }
        // Profiles can only be constructed after validation; evaluate the password before
        // mutating state so a late call remains atomic.
        let secret = Zeroizing::new(password.into());
        self.encryption = Some(WriterEncryption {
            profile,
            password: secret,
        });
        Ok(())
    }

    /// Clear encryption before any payload entry is written.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn clear_encryption(&mut self) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF encryption cannot be changed after payload entries".to_string(),
            ));
        }
        self.encryption = None;
        Ok(())
    }

    /// Set the MIME type for the document
    ///
    /// This sets both the mimetype file and the root manifest entry.
    ///
    /// # Arguments
    ///
    /// * `mimetype` - MIME type string (e.g., "application/vnd.oasis.opendocument.text")
    ///
    /// # Errors
    ///
    /// Returns an error when the MIME type has already been written, another
    /// package entry was written first, or the archive write fails.
    pub fn set_mimetype(&mut self, mimetype: &str) -> Result<()> {
        self.validate_mimetype_publication(mimetype)
            .map_err(PackageWriterError::into_core_error)?;

        let entry = ManifestEntry {
            full_path: "/".to_string(),
            media_type: mimetype.to_string(),
            size: None,
            encryption: None,
        };
        let entry_bytes = self
            .validate_manifest_candidate(&entry)
            .map_err(PackageWriterError::into_core_error)?;

        self.zip_writer
            .write_stored("mimetype", mimetype.as_bytes())
            .map_err(|e| Error::ZipError(e.to_string()))?;

        self.mimetype = Some(mimetype.to_string());
        self.wrote_any_entry = true;
        self.wrote_mimetype = true;
        self.archive_entry_count += 1;
        self.member_paths.insert("mimetype".to_string());
        self.record_manifest_entry(entry, entry_bytes);
        Ok(())
    }

    /// Set the MIME type using the typed sequential publication error.
    ///
    /// This is the caller-owned-sink counterpart to [`Self::set_mimetype`].
    /// It validates the MIME metadata and all bounded bookkeeping before the
    /// `mimetype` local header is emitted; sink and archive failures retain
    /// the original ZIP error as their source.
    pub fn set_mimetype_streaming(&mut self, mimetype: &str) -> PackageWriterResult<()> {
        self.validate_mimetype_publication(mimetype)?;

        let entry = ManifestEntry {
            full_path: "/".to_string(),
            media_type: mimetype.to_string(),
            size: None,
            encryption: None,
        };
        let entry_bytes = self.validate_manifest_candidate(&entry)?;

        if let Err(error) = self
            .zip_writer
            .write_stored("mimetype", mimetype.as_bytes())
        {
            return Err(self.map_archive_error(error));
        }

        self.mimetype = Some(mimetype.to_string());
        self.wrote_any_entry = true;
        self.wrote_mimetype = true;
        self.archive_entry_count += 1;
        self.member_paths.insert("mimetype".to_string());
        self.record_manifest_entry(entry, entry_bytes);
        Ok(())
    }

    fn validate_mimetype_publication(&self, mimetype: &str) -> PackageWriterResult<()> {
        if self.wrote_mimetype {
            return Err(PackageWriterError::Core(Error::InvalidFormat(
                "MIME type already set".to_string(),
            )));
        }
        if self.wrote_any_entry {
            return Err(PackageWriterError::Core(Error::InvalidFormat(
                "Cannot set MIME type after writing other files".to_string(),
            )));
        }
        Self::validate_media_type(mimetype, false, "MIME type").map_err(PackageWriterError::Core)
    }

    fn validate_media_type(value: &str, allow_empty: bool, field: &str) -> Result<()> {
        if allow_empty && value.is_empty() {
            return Ok(());
        }
        if value.is_empty() || value.len() > MAX_MANIFEST_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODF {field} must be non-empty and at most {MAX_MANIFEST_TEXT_BYTES} bytes"
            )));
        }
        if !value.is_ascii()
            || value
                .bytes()
                .any(|byte| (byte < 0x20 && byte != b'\t') || byte == 0x7f)
        {
            return Err(Error::InvalidFormat(format!(
                "ODF {field} must be ASCII and contain no unsafe control characters"
            )));
        }

        let mut segments = value.split(';');
        let essence = segments.next().unwrap_or_default().trim();
        let mut parts = essence.split('/');
        let top_level = parts.next().unwrap_or_default();
        let subtype = parts.next();
        let valid_token = |token: &str| {
            !token.is_empty()
                && token.bytes().all(|byte| {
                    byte.is_ascii_alphanumeric()
                        || matches!(
                            byte,
                            b'!' | b'#'
                                | b'$'
                                | b'%'
                                | b'&'
                                | b'\''
                                | b'*'
                                | b'+'
                                | b'-'
                                | b'.'
                                | b'^'
                                | b'_'
                                | b'`'
                                | b'|'
                                | b'~'
                        )
                })
        };
        if !valid_token(top_level)
            || !subtype.is_some_and(valid_token)
            || parts.next().is_some()
            || segments.any(|parameter| parameter.trim().is_empty())
        {
            return Err(Error::InvalidFormat(format!(
                "ODF {field} must contain a valid type/subtype and non-empty parameters"
            )));
        }
        Ok(())
    }

    fn validate_manifest_version(value: &str) -> Result<()> {
        if value.is_empty() || value.len() > MAX_MANIFEST_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODF manifest version must be non-empty and at most {MAX_MANIFEST_TEXT_BYTES} bytes"
            )));
        }
        if value.chars().any(char::is_control) {
            return Err(Error::InvalidFormat(
                "ODF manifest version contains a control character".to_string(),
            ));
        }
        Ok(())
    }

    fn validate_member_path(&self, path: &str, allow_directory: bool) -> Result<()> {
        if path.is_empty() {
            return Err(Error::InvalidFormat(
                "ODF member path must not be empty".to_string(),
            ));
        }
        if path == "/" {
            return Err(Error::InvalidFormat(
                "ODF member path must be relative and non-root".to_string(),
            ));
        }
        if !allow_directory && path.ends_with('/') {
            return Err(Error::InvalidFormat(
                "ODF file member path must not end with '/'".to_string(),
            ));
        }
        // The shared ODF validator is stricter than ZIP's normalization: it
        // rejects URI aliases (`?`, `#`, percent escapes, dot segments, and
        // separators) before a raw member name can reach the transport.
        validate_manifest_path(path)?;
        if path.len() as u64 > self.limits.max_member_name_bytes {
            return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed: path.len() as u64,
                limit: self.limits.max_member_name_bytes,
                scope: std::sync::Arc::from("ODF member name"),
            }));
        }
        Ok(())
    }

    fn is_reserved_admin_path(path: &str) -> bool {
        matches!(
            path,
            "mimetype" | "manifest.xml" | "META-INF" | "META-INF/" | MANIFEST_PATH
        ) || is_signature_owner_path(path)
    }

    fn is_reserved_write_path(path: &str) -> bool {
        Self::is_reserved_admin_path(path) && !is_signature_owner_path(path)
    }

    fn validate_reader_path(&self, path: &str) -> PackageWriterResult<()> {
        let path_bytes = u64::try_from(path.len()).unwrap_or(u64::MAX);
        if path_bytes > self.limits.max_member_name_bytes {
            return Err(self.manifest_limit_error(
                PackageWriterLimitResource::MemberNameBytes,
                path_bytes,
                self.limits.max_member_name_bytes,
            ));
        }
        self.validate_member_path(path, false)
            .map_err(PackageWriterError::Core)?;
        if Self::is_reserved_admin_path(path) {
            return Err(PackageWriterError::Core(Error::InvalidFormat(format!(
                "ODF reader publication reserves administrative member '{path}'"
            ))));
        }
        Ok(())
    }

    fn manifest_limit_error(
        &self,
        resource: PackageWriterLimitResource,
        actual: u64,
        maximum: u64,
    ) -> PackageWriterError {
        let limit = PackageWriterLimitExceeded {
            resource,
            actual,
            maximum,
        };
        let source = Box::new(PackageWriterError::Core(Error::InvalidFormat(
            limit.to_string(),
        )));
        PackageWriterError::LimitExceeded {
            written: self.zip_writer.output_bytes(),
            limit,
            source,
        }
    }

    fn manifest_fixed_metadata_bytes(&self) -> Result<u64> {
        let mut prefix = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version=""#,
        );
        prefix.push_str(&escape_xml(&self.manifest_version));
        prefix.push('>');
        let bytes = prefix
            .len()
            .checked_add("</manifest:manifest>".len())
            .ok_or_else(|| {
                Error::InvalidFormat("ODF manifest metadata size overflow".to_string())
            })?;
        u64::try_from(bytes)
            .map_err(|_| Error::InvalidFormat("ODF manifest metadata is too large".to_string()))
    }

    fn manifest_entry_metadata_bytes(entry: &ManifestEntry, version: &str) -> Result<u64> {
        let mut xml = String::new();
        let estimate = entry
            .full_path
            .len()
            .saturating_add(entry.media_type.len())
            .saturating_add(1_024);
        xml.try_reserve(estimate)
            .map_err(|source| Error::Allocation {
                resource: "ODF manifest entry metadata",
                source,
            })?;
        Self::write_manifest_entry(&mut xml, entry, version);
        u64::try_from(xml.len())
            .map_err(|_| Error::InvalidFormat("ODF manifest metadata is too large".to_string()))
    }

    fn validate_manifest_candidate(&self, entry: &ManifestEntry) -> PackageWriterResult<u64> {
        if entry.full_path == "/" {
            Self::validate_media_type(&entry.media_type, false, "MIME type")
                .map_err(PackageWriterError::Core)?;
        } else {
            if Self::is_reserved_write_path(&entry.full_path) {
                return Err(PackageWriterError::Core(Error::InvalidFormat(format!(
                    "ODF manifest path '{}' is reserved for generated package metadata",
                    entry.full_path
                ))));
            }
            self.validate_member_path(&entry.full_path, entry.full_path.ends_with('/'))
                .map_err(PackageWriterError::Core)?;
            Self::validate_media_type(&entry.media_type, true, "manifest media type")
                .map_err(PackageWriterError::Core)?;
        }
        if self.manifest_paths.contains(&entry.full_path)
            || self.member_paths.contains(&entry.full_path)
        {
            return Err(PackageWriterError::Core(Error::InvalidFormat(format!(
                "ODF manifest/member path collision: '{}'",
                entry.full_path
            ))));
        }

        let next_entries = self.manifest_entries.len().checked_add(1).ok_or_else(|| {
            PackageWriterError::Core(Error::InvalidFormat(
                "ODF manifest entry count overflow".to_string(),
            ))
        })?;
        let next_entries_u64 = u64::try_from(next_entries).unwrap_or(u64::MAX);
        let maximum_entries = u64::try_from(self.limits.max_entries).unwrap_or(u64::MAX);
        if next_entries > self.limits.max_entries {
            return Err(self.manifest_limit_error(
                PackageWriterLimitResource::FileCount,
                next_entries_u64,
                maximum_entries,
            ));
        }

        let entry_bytes = Self::manifest_entry_metadata_bytes(entry, &self.manifest_version)
            .map_err(PackageWriterError::Core)?;
        let fixed_bytes = self
            .manifest_fixed_metadata_bytes()
            .map_err(PackageWriterError::Core)?;
        let next_metadata = self
            .manifest_metadata_bytes
            .checked_add(entry_bytes)
            .and_then(|bytes| fixed_bytes.checked_add(bytes))
            .ok_or_else(|| {
                PackageWriterError::Core(Error::InvalidFormat(
                    "ODF manifest metadata size overflow".to_string(),
                ))
            })?;
        if next_metadata > self.limits.max_metadata_bytes {
            return Err(self.manifest_limit_error(
                PackageWriterLimitResource::MetadataBytes,
                next_metadata,
                self.limits.max_metadata_bytes,
            ));
        }
        Ok(entry_bytes)
    }

    fn record_manifest_entry(&mut self, entry: ManifestEntry, entry_bytes: u64) {
        self.manifest_metadata_bytes = self.manifest_metadata_bytes.saturating_add(entry_bytes);
        let inserted = self.manifest_paths.insert(entry.full_path.clone());
        debug_assert!(inserted);
        self.manifest_entries.push(entry);
    }

    fn record_member_path(&mut self, path: &str) {
        let inserted = self.member_paths.insert(path.to_string());
        debug_assert!(inserted);
        self.archive_entry_count += 1;
    }

    /// Add a file to the package
    ///
    /// # Arguments
    ///
    /// * `path` - Path within the ZIP archive (e.g., "content.xml", "Pictures/image1.png")
    /// * `content` - File content as bytes
    ///
    /// # Note
    ///
    /// This method automatically adds the file to the manifest with an appropriate media type.
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured, the path is reserved,
    /// encryption fails, or the archive write fails.
    pub fn add_file(&mut self, path: &str, content: &[u8]) -> Result<()> {
        if path == "mimetype" {
            return Err(Error::InvalidFormat(
                "mimetype is written via set_mimetype()".to_string(),
            ));
        }
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }

        // Determine media type based on file extension
        let media_type = Self::guess_media_type(path);

        self.write_file(path, content, media_type, PayloadOrigin::AuthoredOrChanged)
    }

    /// Add a file to the package with a specific media type
    ///
    /// # Arguments
    ///
    /// * `path` - Path within the ZIP archive
    /// * `content` - File content as bytes
    /// * `media_type` - MIME type for the manifest entry
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured, the path is reserved,
    /// encryption fails, or the archive write fails.
    pub fn add_file_with_media_type(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
    ) -> Result<()> {
        if path == "mimetype" {
            return Err(Error::InvalidFormat(
                "mimetype is written via set_mimetype()".to_string(),
            ));
        }
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }

        self.write_file(path, content, media_type, PayloadOrigin::AuthoredOrChanged)
    }

    /// Add an opaque, non-XML file by consuming a reader incrementally.
    ///
    /// The reader is not retained after this method returns. The member uses
    /// Deflate compression, and the ZIP transport applies its configured
    /// finite entry, aggregate, metadata, and output limits.
    ///
    /// XML-classified members, encryption, and document signing are refused
    /// before a new ZIP local header is emitted. A source or sink failure
    /// permanently poisons this writer; the returned error reports accepted
    /// output bytes when publication had already started.
    pub fn add_file_reader<R: Read>(&mut self, path: &str, reader: R) -> PackageWriterResult<()> {
        let media_type = Self::guess_media_type(path);
        self.add_file_reader_with_media_type(path, reader, media_type)
    }

    /// Add an opaque, non-XML file by consuming a reader incrementally with an
    /// explicit manifest media type.
    pub fn add_file_reader_with_media_type<R: Read>(
        &mut self,
        path: &str,
        reader: R,
        media_type: &str,
    ) -> PackageWriterResult<()> {
        self.add_file_reader_with_media_type_and_compression(
            path,
            reader,
            media_type,
            PackageCompression::Deflated,
        )
    }

    /// Add an opaque, non-XML file by consuming a reader incrementally with an
    /// explicit media type and ZIP compression method.
    pub fn add_file_reader_with_media_type_and_compression<R: Read>(
        &mut self,
        path: &str,
        reader: R,
        media_type: &str,
        compression: PackageCompression,
    ) -> PackageWriterResult<()> {
        self.validate_reader_publication(path, media_type)?;

        let entry = ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
            size: None,
            encryption: None,
        };
        let entry_bytes = self.validate_manifest_candidate(&entry)?;

        let result = match compression {
            PackageCompression::Stored => self.zip_writer.write_stored_stream(path, reader),
            PackageCompression::Deflated => self.zip_writer.write_deflated_stream(path, reader),
        };
        if let Err(error) = result {
            return Err(self.map_archive_error(error));
        }

        self.record_manifest_entry(entry, entry_bytes);
        self.wrote_any_entry = true;
        self.record_member_path(path);
        if !path.starts_with("META-INF/") {
            self.wrote_payload_entry = true;
        }
        Ok(())
    }

    fn validate_reader_publication(&self, path: &str, media_type: &str) -> PackageWriterResult<()> {
        self.validate_reader_path(path)?;
        if !self.wrote_mimetype {
            return Err(PackageWriterError::Core(Error::InvalidFormat(
                "MIME type not set".to_string(),
            )));
        }
        Self::validate_media_type(media_type, false, "manifest media type")
            .map_err(PackageWriterError::Core)?;
        if xml_minifier::audit::package::is_xml_part(path, media_type) {
            return Err(PackageWriterError::Core(Error::InvalidFormat(format!(
                "ODF reader publication rejects XML member '{path}'; use add_file() or add_file_with_media_type()"
            ))));
        }
        if self.encryption.is_some() {
            return Err(PackageWriterError::Core(Error::InvalidFormat(
                "ODF reader publication does not support encryption".to_string(),
            )));
        }
        if self.document_signer.is_some() {
            return Err(PackageWriterError::Core(Error::InvalidFormat(
                "ODF reader publication does not support document signing".to_string(),
            )));
        }
        if self.member_paths.contains(path) || self.manifest_paths.contains(path) {
            return Err(PackageWriterError::Core(Error::InvalidFormat(format!(
                "ODF manifest/member path collision: '{path}'"
            ))));
        }
        Ok(())
    }

    fn write_file(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
        origin: PayloadOrigin,
    ) -> Result<()> {
        self.validate_member_path(path, false)?;
        if Self::is_reserved_write_path(path) {
            return Err(Error::InvalidFormat(format!(
                "ODF member path '{path}' is reserved for generated package metadata"
            )));
        }
        Self::validate_media_type(media_type, true, "manifest media type")?;
        if matches!(origin, PayloadOrigin::AuthoredOrChanged) {
            Self::validate_authored_xml(path, content, media_type)?;
        }
        // Validate path/media/collision/count/manifest bookkeeping before any
        // encryption work can materialize a second payload buffer. The final
        // encrypted descriptor is checked again below because it contributes
        // additional bounded manifest metadata.
        let mut entry = ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
            size: None,
            encryption: None,
        };
        let mut entry_bytes = self
            .validate_manifest_candidate(&entry)
            .map_err(PackageWriterError::into_core_error)?;
        let encrypt = self
            .encryption
            .as_ref()
            .filter(|_| !path.starts_with("META-INF/"));
        let encrypted_content = if let Some(settings) = encrypt {
            let (ciphertext, descriptor) =
                encrypt_entry(content, settings.password.as_str(), settings.profile)?;
            let plaintext_size = u64::try_from(content.len()).map_err(|error| {
                Error::InvalidFormat(format!("ODF plaintext entry is too large: {error}"))
            })?;
            entry.size = Some(plaintext_size);
            entry.encryption = Some(descriptor);
            entry_bytes = self
                .validate_manifest_candidate(&entry)
                .map_err(PackageWriterError::into_core_error)?;
            Some(ciphertext)
        } else {
            None
        };
        if let Some(ciphertext) = encrypted_content {
            self.zip_writer
                .write_stored(path, &ciphertext)
                .map_err(|e| Error::ZipError(e.to_string()))?;
        } else {
            self.zip_writer
                .write_deflated_sized(path, content)
                .map_err(|e| Error::ZipError(e.to_string()))?;
        };
        self.record_manifest_entry(entry, entry_bytes);
        self.wrote_any_entry = true;
        self.record_member_path(path);
        if !path.starts_with("META-INF/") {
            self.wrote_payload_entry = true;
        }
        Ok(())
    }

    fn write_exact_source_file(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
    ) -> Result<()> {
        self.write_file(path, content, media_type, PayloadOrigin::ExactSource)
    }

    /// Add a provenance-bearing, individually audited XML splice publication.
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured, the checked part
    /// cannot be assembled, or the archive write fails.
    pub fn add_spliced_xml(&mut self, publication: XmlSplicePublication) -> Result<()> {
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }
        let (path, content, media_type) = publication.assemble()?;
        self.write_file(&path, &content, &media_type, PayloadOrigin::CheckedSplice)
    }

    /// Copy every exact source member except regenerated metadata, signatures,
    /// and the explicitly excluded replacement paths.
    ///
    /// Unlike [`Self::copy_auxiliary_files_from`], this preserves source-loaded
    /// core parts such as `styles.xml` and `meta.xml` during a splice rebuild.
    ///
    /// # Errors
    ///
    /// Returns an error for unsupported source encryption or archive failures.
    pub(crate) fn copy_source_files_from_except(
        &mut self,
        source: &OwnedPackage,
        excluded_paths: &HashSet<String>,
    ) -> Result<()> {
        self.inherit_manifest_version(source)?;
        ensure_source_manifest_rewritable(source)?;
        let package = source.package()?;
        if package.manifest().has_encrypted_entries() && self.encryption.is_none() {
            return Err(Error::InvalidFormat(
                "Rewriting encrypted ODF entries requires writer encryption".to_string(),
            ));
        }
        for (path, entry) in &package.manifest().entries {
            if path.ends_with('/')
                && !matches!(path.as_str(), "/" | "META-INF/")
                && !excluded_paths.contains(path)
            {
                self.add_manifest_entry(path, &entry.media_type)?;
            }
        }
        for path in package.files()? {
            if path.ends_with('/')
                || matches!(path.as_str(), "mimetype" | "META-INF/manifest.xml")
                || is_signature_owner_path(&path)
                || excluded_paths.contains(&path)
            {
                continue;
            }
            let bytes = package.get_file(&path)?;
            let media_type = package
                .manifest()
                .get_media_type(&path)
                .unwrap_or_else(|| Self::guess_media_type(&path));
            self.write_exact_source_file(&path, &bytes, media_type)?;
        }
        Ok(())
    }

    fn validate_authored_xml(path: &str, content: &[u8], media_type: &str) -> Result<()> {
        if !xml_minifier::audit::package::is_xml_part(path, media_type) {
            return Ok(());
        }
        xml_minifier::audit::verify_authored(content, xml_minifier::audit::Limits::default())
            .map(|_report| ())
            .map_err(|source| {
                Error::InvalidFormat(format!("XML publication rejected for '{path}': {source}"))
            })
    }

    /// Add an entry to the package manifest without writing a ZIP member.
    ///
    /// ODF uses manifest-only entries for directories such as embedded objects.
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured or the path is
    /// reserved or empty.
    pub fn add_manifest_entry(&mut self, path: &str, media_type: &str) -> Result<()> {
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }
        if path.is_empty() || path == "mimetype" {
            return Err(Error::InvalidFormat(
                "Invalid manifest-only path".to_string(),
            ));
        }
        if Self::is_reserved_admin_path(path) {
            return Err(Error::InvalidFormat(format!(
                "ODF manifest-only path '{path}' is reserved for generated package metadata or signing"
            )));
        }
        self.validate_member_path(path, path.ends_with('/'))?;
        Self::validate_media_type(media_type, true, "manifest media type")?;

        let entry = ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
            size: None,
            encryption: None,
        };
        let entry_bytes = self
            .validate_manifest_candidate(&entry)
            .map_err(PackageWriterError::into_core_error)?;
        self.record_manifest_entry(entry, entry_bytes);
        Ok(())
    }

    /// Copy all non-core parts from an existing ODF package.
    ///
    /// Core XML parts are regenerated by mutable format writers. Digital
    /// signatures are deliberately omitted because changing those parts
    /// invalidates the signatures. Encrypted parts cannot be reconstructed
    /// faithfully with the current manifest writer and are rejected.
    ///
    /// # Errors
    ///
    /// Returns an error when the source package cannot be read, contains
    /// unsupported encrypted entries, or an entry cannot be copied.
    pub fn copy_auxiliary_files_from(&mut self, source: &OwnedPackage) -> Result<()> {
        self.inherit_manifest_version(source)?;
        ensure_source_manifest_rewritable(source)?;
        let package = source.package()?;
        if package.manifest().has_encrypted_entries() && self.encryption.is_none() {
            return Err(Error::InvalidFormat(
                "Rewriting encrypted ODF entries requires writer encryption".to_string(),
            ));
        }

        for (path, entry) in &package.manifest().entries {
            if path.ends_with('/') && !Self::is_regenerated_package_part(path) {
                self.add_manifest_entry(path, &entry.media_type)?;
            }
        }

        for path in package.files()? {
            if path.ends_with('/') || Self::is_regenerated_package_part(&path) {
                continue;
            }
            let bytes = package.get_file(&path)?;
            let media_type = package
                .manifest()
                .get_media_type(&path)
                .unwrap_or_else(|| Self::guess_media_type(&path));
            self.write_exact_source_file(&path, &bytes, media_type)?;
        }

        Ok(())
    }

    /// Add a directory entry to the generated manifest without a ZIP payload.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is not a safe relative directory or the
    /// manifest entry cannot be added.
    pub fn add_manifest_directory(&mut self, path: &str, media_type: &str) -> Result<()> {
        if !path.ends_with('/') || path == "/" {
            return Err(Error::InvalidFormat(
                "invalid embedded-object manifest directory".to_string(),
            ));
        }
        self.add_manifest_entry(path, media_type)
    }

    /// Copy auxiliary entries except selected exact paths and directory trees.
    ///
    /// # Errors
    ///
    /// Returns an error when the source package cannot be read, contains
    /// unsupported encrypted entries, or an entry cannot be copied.
    pub fn copy_auxiliary_files_from_except(
        &mut self,
        source: &OwnedPackage,
        excluded_paths: &[String],
        excluded_prefixes: &[String],
    ) -> Result<()> {
        self.inherit_manifest_version(source)?;
        ensure_source_manifest_rewritable(source)?;
        let package = source.package()?;
        if package.manifest().has_encrypted_entries() && self.encryption.is_none() {
            return Err(Error::InvalidFormat(
                "Rewriting encrypted ODF entries requires writer encryption".to_string(),
            ));
        }
        let excluded = |path: &str| {
            excluded_paths.iter().any(|candidate| candidate == path)
                || excluded_prefixes
                    .iter()
                    .any(|candidate| path.starts_with(candidate))
        };

        for (path, entry) in &package.manifest().entries {
            if path.ends_with('/') && !Self::is_regenerated_package_part(path) && !excluded(path) {
                self.add_manifest_entry(path, &entry.media_type)?;
            }
        }
        for path in package.files()? {
            if path.ends_with('/') || Self::is_regenerated_package_part(&path) || excluded(&path) {
                continue;
            }
            let bytes = package.get_file(&path)?;
            let media_type = package
                .manifest()
                .get_media_type(&path)
                .unwrap_or_else(|| Self::guess_media_type(&path));
            self.write_exact_source_file(&path, &bytes, media_type)?;
        }
        Ok(())
    }

    /// Request exact publication of a validated source manifest when the
    /// final staged inventory remains equivalent to that source inventory.
    ///
    /// This is an opt-in preservation slot rather than an immediate write:
    /// later member additions or removals are compared at finalization. The
    /// request is deliberately ignored for legacy-only manifest locations,
    /// signed sources, encrypted metadata, and raw metadata larger than the
    /// writer's bounded metadata ceiling. Those cases continue through
    /// canonical manifest generation. Source rebuilds that copy auxiliary
    /// members retain the copy path's explicit refusal for unsupported
    /// unencrypted `manifest:size` metadata.
    pub fn preserve_source_manifest(&mut self, source: &OwnedPackage) -> Result<()> {
        self.preserved_manifest = None;

        let has_canonical_manifest = source.has_file(MANIFEST_PATH)?;
        if !has_canonical_manifest {
            return Ok(());
        }
        ensure_source_manifest_rewritable(source)?;
        let source_files = source.files()?;
        if source_files
            .iter()
            .any(|path| is_signature_owner_path(path))
        {
            return Ok(());
        }

        let package = source.package()?;
        let source_manifest = package.manifest();
        if source_manifest.has_encrypted_entries()
            || source_manifest
                .entries
                .values()
                .any(|entry| entry.size.is_some())
        {
            return Ok(());
        }

        let bytes = source.get_file(MANIFEST_PATH)?;
        let byte_count = u64::try_from(bytes.len()).map_err(|error| {
            Error::InvalidFormat(format!("ODF source manifest is too large: {error}"))
        })?;
        if byte_count > self.limits.max_metadata_bytes {
            return Ok(());
        }

        let mut physical_paths = HashSet::new();
        physical_paths
            .try_reserve(source_files.len())
            .map_err(|source| Error::Allocation {
                resource: "ODF preserved physical member paths",
                source,
            })?;
        for path in &source_files {
            if !physical_paths.insert(normalized_manifest_path(path).to_string()) {
                return Ok(());
            }
        }

        let mut entries = std::collections::HashMap::new();
        entries
            .try_reserve(source_manifest.entries.len())
            .map_err(|source| Error::Allocation {
                resource: "ODF preserved manifest entries",
                source,
            })?;
        for (path, entry) in &source_manifest.entries {
            let normalized_path = normalized_manifest_path(path).to_string();
            if entries.contains_key(&normalized_path) {
                return Ok(());
            }
            entries.insert(
                normalized_path,
                ManifestEntry {
                    full_path: path.clone(),
                    media_type: entry.media_type.clone(),
                    size: entry.size,
                    encryption: entry.encryption.clone(),
                },
            );
        }

        self.preserved_manifest = Some(PreservedManifest {
            bytes,
            entries,
            physical_paths,
        });
        Ok(())
    }

    fn inherit_manifest_version(&mut self, source: &OwnedPackage) -> Result<()> {
        let bytes = source
            .get_file("META-INF/manifest.xml")
            .or_else(|_| source.get_file("manifest.xml"))?;
        if let Some(version) = source_manifest_version(&bytes)? {
            Self::validate_manifest_version(&version)?;
            self.manifest_version = version;
        }
        Ok(())
    }

    fn preserved_manifest_if_equivalent(&self) -> Option<&[u8]> {
        let preserved = self.preserved_manifest.as_ref()?;
        if self.document_signer.is_some() || preserved.entries.len() != self.manifest_entries.len()
        {
            return None;
        }
        let expected_physical_count = self.member_paths.len().checked_add(1)?;
        if preserved.physical_paths.len() != expected_physical_count
            || !preserved.physical_paths.contains(MANIFEST_PATH)
            || !self.member_paths.iter().all(|path| {
                preserved
                    .physical_paths
                    .contains(normalized_manifest_path(path))
            })
        {
            return None;
        }
        for entry in &self.manifest_entries {
            let normalized_path = normalized_manifest_path(&entry.full_path);
            let source_entry = preserved.entries.get(normalized_path)?;
            if source_entry.media_type != entry.media_type
                || source_entry.size != entry.size
                || source_entry.encryption != entry.encryption
            {
                return None;
            }
        }
        Some(preserved.bytes.as_slice())
    }

    fn final_manifest_bytes(&self) -> Result<Vec<u8>> {
        if let Some(bytes) = self.preserved_manifest_if_equivalent() {
            return Ok(bytes.to_vec());
        }
        Ok(self.generate_manifest()?.into_bytes())
    }

    /// Generate the manifest.xml content
    fn generate_manifest(&self) -> Result<String> {
        let total_bytes = self
            .manifest_fixed_metadata_bytes()?
            .checked_add(self.manifest_metadata_bytes)
            .ok_or_else(|| {
                Error::InvalidFormat("ODF manifest metadata size overflow".to_string())
            })?;
        let capacity = usize::try_from(total_bytes)
            .map_err(|_| Error::InvalidFormat("ODF manifest metadata is too large".to_string()))?;
        let mut manifest = String::new();
        manifest
            .try_reserve(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODF manifest metadata",
                source,
            })?;
        manifest.push_str(
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version=""#,
        );
        manifest.push_str(&escape_xml(&self.manifest_version));
        manifest.push_str("\">");

        // Add manifest entries
        for entry in &self.manifest_entries {
            Self::write_manifest_entry(&mut manifest, entry, &self.manifest_version);
        }

        manifest.push_str("</manifest:manifest>");
        Ok(manifest)
    }

    /// Guess media type from file path
    fn guess_media_type(path: &str) -> &'static str {
        if path.ends_with('/') {
            return "";
        }
        let extension = path.rsplit('.').next().unwrap_or_default();
        if extension.eq_ignore_ascii_case("xml") {
            "text/xml"
        } else if extension.eq_ignore_ascii_case("rdf") {
            "application/rdf+xml"
        } else if extension.eq_ignore_ascii_case("png") {
            "image/png"
        } else if extension.eq_ignore_ascii_case("jpg") || extension.eq_ignore_ascii_case("jpeg") {
            "image/jpeg"
        } else if extension.eq_ignore_ascii_case("gif") {
            "image/gif"
        } else if extension.eq_ignore_ascii_case("svg") {
            "image/svg+xml"
        } else {
            "application/octet-stream"
        }
    }

    fn map_archive_error(&self, error: soapberry_zip::Error) -> PackageWriterError {
        let poisoned = self.zip_writer.is_poisoned();
        let written = self.zip_writer.output_bytes();
        let limit = self
            .zip_writer
            .last_limit()
            .map(Self::map_streaming_limit)
            .or_else(|| Self::map_zip_limit(&error));
        let source = PackageWriterError::Archive(error);
        Self::map_stream_failure(
            source,
            if poisoned || limit.is_some() {
                written
            } else {
                0
            },
            limit,
        )
    }

    fn map_archive_failure(failure: StreamingArchiveFailure) -> PackageWriterError {
        let written = failure.progress().output_bytes();
        let limit = failure
            .limit()
            .map(Self::map_streaming_limit)
            .or_else(|| Self::map_zip_limit(failure.error()));
        Self::map_stream_failure(PackageWriterError::ArchiveFailure(failure), written, limit)
    }

    fn map_streaming_limit(limit: StreamingLimitExceeded) -> PackageWriterLimitExceeded {
        let resource = match limit.resource() {
            StreamingLimitResource::CompressedBytes => PackageWriterLimitResource::CompressedSize,
            StreamingLimitResource::OutputBytes => PackageWriterLimitResource::OutputBytes,
            _ => PackageWriterLimitResource::Other,
        };
        PackageWriterLimitExceeded {
            resource,
            actual: limit.actual(),
            maximum: limit.maximum(),
        }
    }

    fn map_zip_limit(error: &soapberry_zip::Error) -> Option<PackageWriterLimitExceeded> {
        let ZipErrorKind::LimitExceeded {
            resource,
            actual,
            maximum,
        } = error.kind()
        else {
            return None;
        };
        let resource = match resource {
            ZipLimitResource::FileCount => PackageWriterLimitResource::FileCount,
            ZipLimitResource::MemberNameBytes => PackageWriterLimitResource::MemberNameBytes,
            ZipLimitResource::MetadataBytes => PackageWriterLimitResource::MetadataBytes,
            ZipLimitResource::CompressedSize => PackageWriterLimitResource::CompressedSize,
            ZipLimitResource::EntrySize => PackageWriterLimitResource::EntrySize,
            ZipLimitResource::TotalSize => PackageWriterLimitResource::TotalSize,
        };
        Some(PackageWriterLimitExceeded {
            resource,
            actual: *actual,
            maximum: *maximum,
        })
    }

    fn map_stream_error(
        source: Error,
        written: u64,
        limit: Option<PackageWriterLimitExceeded>,
    ) -> PackageWriterError {
        Self::map_stream_failure(PackageWriterError::Core(source), written, limit)
    }

    fn map_stream_failure(
        source: PackageWriterError,
        written: u64,
        limit: Option<PackageWriterLimitExceeded>,
    ) -> PackageWriterError {
        let source = Box::new(source);
        if let Some(limit) = limit {
            PackageWriterError::LimitExceeded {
                written,
                limit,
                source,
            }
        } else if written != 0 {
            PackageWriterError::IncompleteOutput { written, source }
        } else {
            *source
        }
    }

    fn validate_finish_publication(&self) -> PackageWriterResult<()> {
        if !self.wrote_mimetype {
            return Err(PackageWriterError::Core(Error::InvalidFormat(
                "MIME type not set".to_string(),
            )));
        }
        let manifest_bytes = if let Some(raw_manifest) = self.preserved_manifest_if_equivalent() {
            u64::try_from(raw_manifest.len()).map_err(|error| {
                PackageWriterError::Core(Error::InvalidFormat(format!(
                    "ODF preserved manifest metadata is too large: {error}"
                )))
            })?
        } else {
            let fixed_bytes = self
                .manifest_fixed_metadata_bytes()
                .map_err(PackageWriterError::Core)?;
            fixed_bytes
                .checked_add(self.manifest_metadata_bytes)
                .ok_or_else(|| {
                    PackageWriterError::Core(Error::InvalidFormat(
                        "ODF manifest metadata size overflow".to_string(),
                    ))
                })?
        };
        if manifest_bytes > self.limits.max_metadata_bytes {
            return Err(self.manifest_limit_error(
                PackageWriterLimitResource::MetadataBytes,
                manifest_bytes,
                self.limits.max_metadata_bytes,
            ));
        }
        let next_archive_entries = self.archive_entry_count.checked_add(1).ok_or_else(|| {
            PackageWriterError::Core(Error::InvalidFormat(
                "ODF archive entry count overflow".to_string(),
            ))
        })?;
        if next_archive_entries > self.limits.max_entries {
            return Err(self.manifest_limit_error(
                PackageWriterLimitResource::FileCount,
                u64::try_from(next_archive_entries).unwrap_or(u64::MAX),
                u64::try_from(self.limits.max_entries).unwrap_or(u64::MAX),
            ));
        }
        Ok(())
    }

    fn finish_into_writer_with_progress(
        mut self,
    ) -> PackageWriterResult<(W, Option<crate::signature::DocumentSigner>)> {
        self.validate_finish_publication()?;

        // Select and write the manifest only after every staged entry has been
        // accounted for. A retained source payload is already validated by
        // `OwnedPackage`; generated content still goes through the authored
        // XML publication audit.
        let manifest_content = match self.final_manifest_bytes() {
            Ok(content) => content,
            Err(error) => {
                return Err(Self::map_stream_error(
                    error,
                    self.zip_writer.output_bytes(),
                    self.zip_writer.last_limit().map(Self::map_streaming_limit),
                ));
            },
        };
        if self.preserved_manifest_if_equivalent().is_none()
            && let Err(error) =
                Self::validate_authored_xml("META-INF/manifest.xml", &manifest_content, "text/xml")
        {
            return Err(Self::map_stream_error(
                error,
                self.zip_writer.output_bytes(),
                self.zip_writer.last_limit().map(Self::map_streaming_limit),
            ));
        }
        if let Err(error) = self
            .zip_writer
            .write_deflated_sized("META-INF/manifest.xml", &manifest_content)
        {
            return Err(self.map_archive_error(error));
        }

        let signer = self.document_signer.take();
        self.zip_writer
            .finish_with_progress()
            .map(|(writer, _progress)| (writer, signer))
            .map_err(Self::map_archive_failure)
    }

    /// Finish writing the package and return the bytes.
    ///
    /// This method writes the mimetype file, manifest, and finalizes the ZIP archive.
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - No MIME type has been set
    /// - Writing to the ZIP archive fails
    fn finish_into_writer(self) -> Result<(W, Option<crate::signature::DocumentSigner>)> {
        self.finish_into_writer_with_progress()
            .map_err(PackageWriterError::into_core_error)
    }

    /// Finish the package on a caller-owned sequential sink.
    ///
    /// This method never signs the package and refuses a configured document
    /// signer before writing the manifest. If the sink has accepted bytes,
    /// failures preserve the accepted count in [`PackageWriterError`].
    pub fn finish_to_writer(self) -> PackageWriterResult<W> {
        if self.document_signer.is_some() {
            let source = Error::InvalidFormat(
                "ODF sequential package output does not support document signing".to_string(),
            );
            let written = self.zip_writer.output_bytes();
            return Err(Self::map_stream_error(source, written, None));
        }
        let (writer, signer) = self.finish_into_writer_with_progress()?;
        debug_assert!(signer.is_none());
        Ok(writer)
    }
}

impl PackageWriter<io::Cursor<Vec<u8>>> {
    /// Finish writing the package and return the bytes.
    ///
    /// This method writes the mimetype file, manifest, and finalizes the ZIP archive.
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - No MIME type has been set
    /// - Writing to the ZIP archive fails
    pub fn finish(self) -> Result<Vec<u8>> {
        let (cursor, document_signer) = self.finish_into_writer()?;
        let bytes = cursor.into_inner();
        if let Some(signer) = &document_signer {
            crate::signature::sign_package(&bytes, signer)
        } else {
            Ok(bytes)
        }
    }

    /// Alias for `finish()` for API compatibility.
    ///
    /// # Errors
    ///
    /// Returns an error under the same conditions as [`Self::finish`].
    pub fn finish_to_bytes(self) -> Result<Vec<u8>> {
        self.finish()
    }
}

impl PackageWriter<BoundedBytes> {
    /// Finish a package into the configured bounded sink.
    ///
    /// Document signing is refused because signing can create a second package
    /// representation outside this sink.
    ///
    /// # Errors
    ///
    /// Returns an error when signing is configured, a MIME type is missing, or
    /// the archive cannot be finalized into the bounded sink.
    pub fn finish_to_bounded_bytes(self) -> Result<Vec<u8>> {
        if self.document_signer.is_some() {
            return Err(Error::InvalidFormat(
                "ODF bounded package output does not support document signing".to_string(),
            ));
        }
        let (sink, signer) = self.finish_into_writer()?;
        debug_assert!(signer.is_none());
        Ok(sink.into_inner())
    }
}

impl<W: Write> PackageWriter<W> {
    fn write_manifest_entry(xml: &mut String, entry: &ManifestEntry, manifest_version: &str) {
        xml.push_str("<manifest:file-entry manifest:full-path=\"");
        xml.push_str(&escape_xml(&entry.full_path));
        if entry.full_path == "/" {
            xml.push_str("\" manifest:version=\"");
            xml.push_str(&escape_xml(manifest_version));
        }
        xml.push_str("\" manifest:media-type=\"");
        xml.push_str(&escape_xml(&entry.media_type));
        xml.push('"');
        if let Some(size) = entry.size {
            xml.push_str(" manifest:size=\"");
            xml.push_str(&size.to_string());
            xml.push('"');
        }
        let Some(encryption) = &entry.encryption else {
            xml.push_str("/>");
            return;
        };
        xml.push('>');
        xml.push_str("<manifest:encryption-data");
        if let Some(checksum) = &encryption.checksum {
            let algorithm = match checksum.algorithm {
                ManifestChecksumAlgorithm::Sha1First1024 => "SHA1/1K",
                ManifestChecksumAlgorithm::Sha256First1024 => {
                    "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k"
                },
            };
            xml.push_str(" manifest:checksum-type=\"");
            xml.push_str(algorithm);
            xml.push_str("\" manifest:checksum=\"");
            xml.push_str(&BASE64_STANDARD.encode(&checksum.value));
            xml.push('"');
        }
        xml.push('>');

        let (algorithm_name, iv): (&str, &[u8]) = match &encryption.algorithm {
            ManifestEncryptionAlgorithm::Aes128Cbc { iv } => {
                ("http://www.w3.org/2001/04/xmlenc#aes128-cbc", iv)
            },
            ManifestEncryptionAlgorithm::Aes192Cbc { iv } => {
                ("http://www.w3.org/2001/04/xmlenc#aes192-cbc", iv)
            },
            ManifestEncryptionAlgorithm::Aes256Cbc { iv } => {
                ("http://www.w3.org/2001/04/xmlenc#aes256-cbc", iv)
            },
            ManifestEncryptionAlgorithm::Aes128Gcm { iv } => {
                ("http://www.w3.org/2009/xmlenc11#aes128-gcm", iv)
            },
            ManifestEncryptionAlgorithm::Aes192Gcm { iv } => {
                ("http://www.w3.org/2009/xmlenc11#aes192-gcm", iv)
            },
            ManifestEncryptionAlgorithm::Aes256Gcm { iv } => {
                ("http://www.w3.org/2009/xmlenc11#aes256-gcm", iv)
            },
            ManifestEncryptionAlgorithm::BlowfishCfb8 { iv } => ("Blowfish CFB", iv),
        };
        xml.push_str("<manifest:algorithm manifest:algorithm-name=\"");
        xml.push_str(algorithm_name);
        xml.push_str("\" manifest:initialisation-vector=\"");
        xml.push_str(&BASE64_STANDARD.encode(iv));
        xml.push_str("\"/>");

        let (start_name, start_size) = match encryption.start_key {
            ManifestStartKeyGeneration::Sha1 => ("SHA1", 20),
            ManifestStartKeyGeneration::Sha256 => ("http://www.w3.org/2001/04/xmlenc#sha256", 32),
        };
        xml.push_str("<manifest:start-key-generation manifest:start-key-generation-name=\"");
        xml.push_str(start_name);
        xml.push_str("\" manifest:key-size=\"");
        xml.push_str(&start_size.to_string());
        xml.push_str("\"/>");

        match &encryption.key_derivation {
            ManifestKeyDerivation::Pbkdf2 {
                salt,
                iterations,
                key_size,
            } => {
                xml.push_str(
                "<manifest:key-derivation manifest:key-derivation-name=\"PBKDF2\" manifest:salt=\"",
            );
                xml.push_str(&BASE64_STANDARD.encode(salt));
                xml.push_str("\" manifest:iteration-count=\"");
                xml.push_str(&iterations.get().to_string());
                xml.push_str("\" manifest:key-size=\"");
                xml.push_str(&key_size.to_string());
                xml.push_str("\"/>");
            },
            ManifestKeyDerivation::Argon2id {
                salt,
                iterations,
                memory_kib,
                lanes,
                key_size,
            } => {
                xml.push_str("<manifest:key-derivation manifest:key-derivation-name=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.5#argon2id\" manifest:salt=\"");
                xml.push_str(&BASE64_STANDARD.encode(salt));
                xml.push_str("\" manifest:argon2-iterations=\"");
                xml.push_str(&iterations.get().to_string());
                xml.push_str("\" manifest:argon2-memory=\"");
                xml.push_str(&memory_kib.get().to_string());
                xml.push_str("\" manifest:argon2-lanes=\"");
                xml.push_str(&lanes.get().to_string());
                xml.push('"');
                if let Some(optional_key_size) = key_size {
                    xml.push_str(" manifest:key-size=\"");
                    xml.push_str(&optional_key_size.to_string());
                    xml.push('"');
                }
                xml.push_str("/>");
            },
        }
        xml.push_str("</manifest:encryption-data></manifest:file-entry>");
    }

    fn is_regenerated_package_part(path: &str) -> bool {
        matches!(
            path,
            "/" | "mimetype"
                | "content.xml"
                | "styles.xml"
                | "meta.xml"
                | "manifest.xml"
                | "META-INF/"
                | "META-INF/manifest.xml"
        ) || is_signature_owner_path(path)
    }
}

impl Structure {
    /// Generate a default content.xml skeleton
    #[must_use]
    pub fn default_content_xml(office_type: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><{office_type}></{office_type}></office:body></office:document-content>"#
        )
    }

    /// Generate a default styles.xml skeleton
    #[must_use]
    pub fn default_styles_xml() -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.3"><office:font-face-decls/><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#.to_string()
    }

    /// Generate a default meta.xml skeleton
    #[must_use]
    pub fn default_meta_xml() -> String {
        Self::default_meta_xml_inner(None, None)
    }

    /// Generate a metadata skeleton from explicitly supplied canonical writer
    /// date-time values.
    ///
    /// The no-argument [`Self::default_meta_xml`] intentionally contains no
    /// creation or modification timestamp. New package output must not read
    /// the host clock. Callers that need document dates can provide checked
    /// ODF date-time values here; invalid values are rejected before any XML
    /// is returned. The accepted writer profile is exactly a four-digit year
    /// in `0001..=9999`, `YYYY-MM-DDThh:mm:ss`, an optional fractional
    /// second, and either no timezone, `Z`, or an offset in `±14:00`. Broader
    /// XSD lexical forms are intentionally refused. Other metadata fields are
    /// deliberately not accepted by this minimal skeleton API, so they cannot
    /// be silently dropped.
    ///
    /// # Errors
    ///
    /// Returns an error when a supplied creation or modification date is not
    /// valid under the canonical writer profile.
    pub fn default_meta_xml_with_dates(
        creation_date: Option<&str>,
        modification_date: Option<&str>,
    ) -> Result<String> {
        if let Some(value) = creation_date {
            validate_canonical_odf_datetime(value, "creation")?;
        }
        if let Some(value) = modification_date {
            validate_canonical_odf_datetime(value, "modification")?;
        }

        Ok(Self::default_meta_xml_inner(
            creation_date,
            modification_date,
        ))
    }

    fn default_meta_xml_inner(
        creation_date: Option<&str>,
        modification_date: Option<&str>,
    ) -> String {
        let mut dates = String::new();
        if let Some(value) = creation_date {
            dates.push_str("<meta:creation-date>");
            dates.push_str(value);
            dates.push_str("</meta:creation-date>");
        }
        if let Some(value) = modification_date {
            dates.push_str("<dc:date>");
            dates.push_str(value);
            dates.push_str("</dc:date>");
        }
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.3"><office:meta><meta:generator>{}</meta:generator>{}</office:meta></office:document-meta>"#,
            "Litchi/0.0.1", dates,
        )
    }

    /// Generate a default settings.xml skeleton
    #[must_use]
    pub fn default_settings_xml() -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewAreaTop" config:type="long">0</config:config-item><config:config-item config:name="ViewAreaLeft" config:type="long">0</config:config-item><config:config-item config:name="ViewAreaWidth" config:type="long">1</config:config-item><config:config-item config:name="ViewAreaHeight" config:type="long">1</config:config-item></config:config-item-set></office:settings></office:document-settings>"#.to_string()
    }
}

fn source_manifest_version(bytes: &[u8]) -> Result<Option<String>> {
    let mut reader = NsReader::from_reader(bytes);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid source manifest XML: {error}"))
            })?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
                    && element.local_name().as_ref() == b"manifest" =>
            {
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid source manifest attribute: {error}"))
                    })?;
                    let (attribute_namespace, local) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if matches!(attribute_namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE)
                        && local.as_ref() == b"version"
                    {
                        let version = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid source manifest version: {error}"
                                ))
                            })?
                            .into_owned();
                        if !matches!(version.as_str(), "1.0" | "1.1" | "1.2" | "1.3" | "1.4") {
                            return Err(Error::InvalidFormat(format!(
                                "unsupported source manifest version '{version}'"
                            )));
                        }
                        return Ok(Some(version));
                    }
                }
                return Ok(None);
            },
            Event::Eof => return Ok(None),
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible ZIP setup operations so assertions can focus on writer behavior."
)]
mod tests {
    use super::*;
    use soapberry_zip::office::ArchiveReader;
    use std::io::{Cursor, Write};

    #[test]
    fn test_package_writer_new() {
        let writer = PackageWriter::new();
        assert!(!writer.wrote_mimetype);
        assert!(!writer.wrote_any_entry);
        assert!(writer.mimetype.is_none());
    }

    #[test]
    fn test_package_writer_default() {
        let writer = PackageWriter::default();
        assert!(!writer.wrote_mimetype);
    }

    #[test]
    fn test_package_writer_set_mimetype() {
        let mut writer = PackageWriter::new();
        assert!(
            writer
                .set_mimetype("application/vnd.oasis.opendocument.text")
                .is_ok()
        );
        assert!(writer.wrote_mimetype);
        assert_eq!(
            writer.mimetype,
            Some("application/vnd.oasis.opendocument.text".to_string())
        );
    }

    #[test]
    fn test_package_writer_set_mimetype_twice() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(
            writer
                .set_mimetype("application/vnd.oasis.opendocument.spreadsheet")
                .is_err()
        );
    }

    #[test]
    fn test_package_writer_add_file_without_mimetype() {
        let mut writer = PackageWriter::new();
        assert!(writer.add_file("content.xml", b"test").is_err());
    }

    #[test]
    fn test_package_writer_add_mimetype_as_file() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(writer.add_file("mimetype", b"test").is_err());
    }

    #[test]
    fn test_package_writer_finish_without_mimetype() {
        let writer = PackageWriter::new();
        assert!(writer.finish().is_err());
    }

    #[test]
    fn test_package_writer_finish_to_bytes() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        let result = writer.finish_to_bytes();
        assert!(result.is_ok());
        let bytes = result.unwrap();
        assert!(!bytes.is_empty());
    }

    #[test]
    fn test_package_writer_add_file_with_media_type() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(
            writer
                .add_file_with_media_type("custom.dat", b"data", "application/octet-stream")
                .is_ok()
        );
    }

    #[test]
    fn arbitrary_writer_bytes_are_refused_for_every_xml_classification() {
        for (path, media_type) in [
            ("manifest.rdf", "application/octet-stream"),
            ("custom/metadata", "application/rdf+xml"),
            (
                "META-INF/custom-signature",
                "application/vnd.oasis.opendocument.digital-signature+xml",
            ),
        ] {
            let mut writer = PackageWriter::new();
            writer
                .set_mimetype("application/vnd.oasis.opendocument.text")
                .unwrap();
            let error = writer
                .add_file_with_media_type(path, b"<root> <child/></root>", media_type)
                .unwrap_err();
            assert!(error.to_string().contains("XML publication rejected"));
        }
    }

    #[test]
    fn real_package_enumeration_includes_rdf_and_manifest_declared_xml() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type(
                "manifest.rdf",
                b"<rdf:RDF xmlns:rdf=\"urn:test\"><rdf:Description/></rdf:RDF>",
                "application/rdf+xml",
            )
            .unwrap();
        writer
            .add_file_with_media_type(
                "custom/metadata",
                b"<metadata><value>content</value></metadata>",
                "application/vnd.example.metadata+xml",
            )
            .unwrap();
        let bytes = writer.finish().unwrap();
        let archive = ArchiveReader::new(&bytes).unwrap();
        let manifest_xml = archive.read("META-INF/manifest.xml").unwrap();
        let manifest_text = std::str::from_utf8(&manifest_xml).unwrap();
        let manifest = super::super::manifest::Manifest::parse(manifest_text).unwrap();
        let mut audited = Vec::new();
        for name in archive.file_names() {
            let media_type = manifest.get_media_type(name).unwrap_or_default();
            if xml_minifier::audit::package::is_xml_part(name, media_type) {
                let payload = archive.read(name).unwrap();
                let _report = xml_minifier::audit::verify_authored(
                    &payload,
                    xml_minifier::audit::Limits::default(),
                )
                .unwrap();
                audited.push(name.to_string());
            }
        }
        audited.sort();
        assert_eq!(
            audited,
            ["META-INF/manifest.xml", "custom/metadata", "manifest.rdf"]
        );
    }

    #[test]
    fn auxiliary_copy_preserves_noncompact_source_rdf_exactly() {
        let mut source_bytes = Vec::new();
        let source_rdf = b"<rdf:RDF xmlns:rdf=\"urn:test\">\n <rdf:Description/>\n</rdf:RDF>";
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut source_bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("manifest.rdf", options).unwrap();
            zip.write_all(source_rdf).unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/></manifest:manifest>"#).unwrap();
            zip.finish().unwrap();
        }
        let source = OwnedPackage::from_bytes(source_bytes).unwrap();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer.copy_auxiliary_files_from(&source).unwrap();
        let output = writer.finish().unwrap();
        let archive = ArchiveReader::new(&output).unwrap();
        assert_eq!(archive.read("manifest.rdf").unwrap(), source_rdf);
    }

    #[test]
    fn copying_auxiliary_files_rejects_encrypted_manifest_entries() {
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(b"encrypted payload").unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA256" manifest:checksum="checksum"/></manifest:file-entry></manifest:manifest>"#).unwrap();
            zip.finish().unwrap();
        }

        let source = OwnedPackage::from_bytes(bytes).unwrap();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        let error = writer.copy_auxiliary_files_from(&source).unwrap_err();
        assert!(error.to_string().contains("encrypted entries"));
    }

    #[test]
    fn source_manifest_accepts_loext_key_derivation_attributes() {
        let accepted = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><m:key-derivation loext:argon2-iterations="1"/></m:manifest>"#;
        assert!(ensure_supported_manifest_metadata(accepted).is_ok());

        let rejected = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:manifest:1.0"><m:key-derivation loext:argon2-iterations="1"/></m:manifest>"#;
        assert!(ensure_supported_manifest_metadata(rejected).is_err());
    }

    #[test]
    fn copying_auxiliary_files_rejects_unencrypted_manifest_sizes() {
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(b"<content/>").unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml" manifest:size="10"/></manifest:manifest>"#).unwrap();
            zip.finish().unwrap();
        }
        let source = OwnedPackage::from_bytes(bytes).unwrap();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        let error = writer.copy_auxiliary_files_from(&source).unwrap_err();
        assert!(error.to_string().contains("manifest:size"));
    }

    #[test]
    fn test_guess_media_type() {
        type MemoryWriter = PackageWriter<Cursor<Vec<u8>>>;
        assert_eq!(MemoryWriter::guess_media_type("content.xml"), "text/xml");
        assert_eq!(
            MemoryWriter::guess_media_type("manifest.rdf"),
            "application/rdf+xml"
        );
        assert_eq!(MemoryWriter::guess_media_type("image.png"), "image/png");
        assert_eq!(MemoryWriter::guess_media_type("image.jpg"), "image/jpeg");
        assert_eq!(MemoryWriter::guess_media_type("image.jpeg"), "image/jpeg");
        assert_eq!(MemoryWriter::guess_media_type("image.gif"), "image/gif");
        assert_eq!(MemoryWriter::guess_media_type("image.svg"), "image/svg+xml");
        assert_eq!(MemoryWriter::guess_media_type("META-INF/"), "");
        assert_eq!(
            MemoryWriter::guess_media_type("data.bin"),
            "application/octet-stream"
        );
    }

    #[test]
    fn test_generate_manifest() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer.add_file("content.xml", b"<content/>").unwrap();

        let manifest = writer.generate_manifest().unwrap();
        assert!(manifest.contains("manifest:manifest"));
        assert!(manifest.contains("content.xml"));
        assert!(manifest.contains("text/xml"));
        assert!(manifest.contains("manifest:full-path=\"/\" manifest:version=\"1.3\""));
        assert!(!manifest.contains("manifest:full-path=\"META-INF/\""));
        assert!(!manifest.contains("manifest:full-path=\"META-INF/manifest.xml\""));
    }

    #[test]
    fn finalized_manifest_does_not_describe_itself_or_meta_inf() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.graphics")
            .unwrap();
        writer.add_file("content.xml", b"<content/>").unwrap();

        let bytes = writer.finish().unwrap();
        let archive = ArchiveReader::new(&bytes).unwrap();
        let manifest = archive.read("META-INF/manifest.xml").unwrap();
        let manifest = std::str::from_utf8(&manifest).unwrap();

        assert!(!manifest.contains("manifest:full-path=\"META-INF/\""));
        assert!(!manifest.contains("manifest:full-path=\"META-INF/manifest.xml\""));
    }

    #[test]
    fn source_manifest_version_is_retained_during_auxiliary_copy() {
        let mut source_writer = PackageWriter::new();
        source_writer.manifest_version = "1.2".to_string();
        source_writer
            .set_mimetype("application/vnd.oasis.opendocument.graphics")
            .unwrap();
        source_writer
            .add_file_with_media_type("Pictures/pixel.png", b"pixel", "image/png")
            .unwrap();
        let source = OwnedPackage::from_bytes(source_writer.finish().unwrap()).unwrap();

        let mut destination = PackageWriter::new();
        destination
            .set_mimetype("application/vnd.oasis.opendocument.graphics")
            .unwrap();
        destination.copy_auxiliary_files_from(&source).unwrap();
        let bytes = destination.finish().unwrap();
        let archive = ArchiveReader::new(&bytes).unwrap();
        let manifest = archive.read("META-INF/manifest.xml").unwrap();
        let manifest = std::str::from_utf8(&manifest).unwrap();

        assert!(manifest.contains("manifest:manifest"));
        assert!(manifest.contains("manifest:version=\"1.2\""));
        assert!(manifest.contains("manifest:full-path=\"/\" manifest:version=\"1.2\""));
    }

    #[test]
    fn source_manifest_version_accepts_odf_1_4_and_rejects_unknown_versions() {
        let namespace = String::from_utf8_lossy(MANIFEST_NAMESPACE);
        let manifest = |version: &str| {
            format!(
                "<manifest:manifest xmlns:manifest=\"{namespace}\" manifest:version=\"{version}\"/>"
            )
        };

        assert_eq!(
            source_manifest_version(manifest("1.4").as_bytes()).unwrap(),
            Some("1.4".to_string())
        );
        assert!(source_manifest_version(manifest("1.5").as_bytes()).is_err());
    }

    #[test]
    fn test_odf_structure_default_styles_xml() {
        let styles = Structure::default_styles_xml();
        assert!(styles.contains("office:document-styles"));
        assert!(styles.contains("office:styles"));
    }

    #[test]
    fn test_odf_structure_default_meta_xml() {
        let meta = Structure::default_meta_xml();
        assert!(meta.contains("office:document-meta"));
        assert!(meta.contains("Litchi"));
        assert!(!meta.contains("meta:creation-date"));
        assert!(!meta.contains("<dc:date>"));
    }

    #[test]
    fn default_meta_xml_is_byte_identical_without_ambient_time() {
        assert_eq!(Structure::default_meta_xml(), Structure::default_meta_xml());

        let build = || {
            let mut writer = PackageWriter::new();
            writer
                .set_mimetype("application/vnd.oasis.opendocument.text")
                .unwrap();
            writer
                .add_file(
                    "content.xml",
                    Structure::default_content_xml("office:text").as_bytes(),
                )
                .unwrap();
            writer
                .add_file("styles.xml", Structure::default_styles_xml().as_bytes())
                .unwrap();
            writer
                .add_file("meta.xml", Structure::default_meta_xml().as_bytes())
                .unwrap();
            writer.finish_to_bytes().unwrap()
        };
        assert_eq!(build(), build());
    }

    #[test]
    fn default_meta_xml_serializes_checked_explicit_metadata() {
        let meta = Structure::default_meta_xml_with_dates(
            Some("2024-01-02T03:04:05.123456Z"),
            Some("2024-06-07T08:09:10+02:00"),
        )
        .unwrap();
        assert!(meta.contains("<meta:generator>Litchi/0.0.1</meta:generator>"));
        assert!(
            meta.contains("<meta:creation-date>2024-01-02T03:04:05.123456Z</meta:creation-date>")
        );
        assert!(meta.contains("<dc:date>2024-06-07T08:09:10+02:00</dc:date>"));
    }

    #[test]
    fn default_meta_xml_rejects_invalid_explicit_metadata_dates() {
        for value in [
            "not-a-date",
            "2024-01-02 03:04:05Z",
            "2024-01-02T03:04:05z",
            "2024-01-02T03:04:05+14:01",
            "2024-01-02T03:04:05+23:59",
            "2024-01-02T03:04:05.Z",
            "2024-02-30T03:04:05Z",
            "0000-01-02T03:04:05Z",
            "-0001-01-02T03:04:05Z",
            "10000-01-02T03:04:05Z",
            "2024-01-02T24:00:00Z",
        ] {
            assert!(
                Structure::default_meta_xml_with_dates(Some(value), None).is_err(),
                "accepted invalid date-time {value}"
            );
        }

        for value in [
            "2024-01-02T03:04:05Z",
            "2024-01-02T03:04:05.123456Z",
            "2024-01-02T03:04:05+14:00",
            "2024-01-02T03:04:05-14:00",
            "2024-01-02T03:04:05",
        ] {
            assert!(
                Structure::default_meta_xml_with_dates(Some(value), None).is_ok(),
                "rejected valid date-time {value}"
            );
        }
    }

    #[test]
    fn test_odf_structure_default_settings_xml() {
        let settings = Structure::default_settings_xml();
        assert!(settings.contains("office:document-settings"));
        assert!(settings.contains("config:config-item"));
    }

    #[test]
    fn test_odf_structure_default_content_xml() {
        let content = Structure::default_content_xml("office:text");
        assert!(content.contains("office:document-content"));
        assert!(content.contains("office:text"));
        assert!(content.contains("office:body"));
    }

    #[test]
    fn test_manifest_entry_debug() {
        let entry = ManifestEntry {
            full_path: "content.xml".to_string(),
            media_type: "text/xml".to_string(),
            size: None,
            encryption: None,
        };
        let debug_str = format!("{entry:?}");
        assert!(debug_str.contains("content.xml"));
        assert!(debug_str.contains("text/xml"));
    }

    #[test]
    fn test_package_writer_full_package() {
        let mut writer = PackageWriter::new();

        // Set mimetype
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();

        // Add files
        writer
            .add_file("content.xml", b"<office:document-content/>")
            .unwrap();
        writer
            .add_file("styles.xml", b"<office:document-styles/>")
            .unwrap();
        writer
            .add_file("meta.xml", b"<office:document-meta/>")
            .unwrap();

        // Finish
        let bytes = writer.finish().unwrap();
        assert!(!bytes.is_empty());

        // Verify it's a valid ZIP (starts with PK)
        assert_eq!(&bytes[0..2], b"PK");
    }

    fn physical_inventory_source() -> (OwnedPackage, Vec<u8>) {
        let manifest = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="Pictures/" m:media-type=""/></m:manifest>"#;
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(b"<content/>").unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest).unwrap();
            zip.finish().unwrap();
        }
        (OwnedPackage::from_bytes(bytes).unwrap(), manifest.to_vec())
    }

    #[test]
    fn preserved_manifest_requires_physical_inventory_equivalence() {
        let (source, source_manifest) = physical_inventory_source();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type("content.xml", b"<content/>", "text/xml")
            .unwrap();
        writer.copy_auxiliary_files_from(&source).unwrap();
        writer.preserve_source_manifest(&source).unwrap();
        let output = writer.finish().unwrap();
        let output_package = OwnedPackage::from_bytes(output).unwrap();
        assert_eq!(
            output_package.get_file(MANIFEST_PATH).unwrap(),
            source_manifest
        );

        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type("content.xml", b"<content/>", "text/xml")
            .unwrap();
        writer.copy_auxiliary_files_from(&source).unwrap();
        writer.preserve_source_manifest(&source).unwrap();
        writer
            .add_file_with_media_type("extra.bin", b"extra", "application/octet-stream")
            .unwrap();
        let output_package = OwnedPackage::from_bytes(writer.finish().unwrap()).unwrap();
        assert_ne!(
            output_package.get_file(MANIFEST_PATH).unwrap(),
            source_manifest
        );
    }

    #[test]
    fn copying_auxiliary_files_rejects_canonical_and_legacy_manifest_ambiguity() {
        let canonical = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/></m:manifest>"#;
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file(MANIFEST_PATH, options).unwrap();
            zip.write_all(canonical).unwrap();
            zip.start_file("manifest.xml", options).unwrap();
            zip.write_all(canonical).unwrap();
            zip.finish().unwrap();
        }
        let source = OwnedPackage::from_bytes(bytes).unwrap();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        let error = writer.copy_auxiliary_files_from(&source).unwrap_err();
        assert!(error.to_string().contains("both canonical and legacy"));
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible encrypted-package setup so assertions can focus on rewrite behavior."
)]
mod encrypted_copy_tests {
    use super::*;

    #[test]
    fn encrypted_auxiliary_entries_require_plaintext_and_are_reencrypted() {
        let mut source_writer = PackageWriter::new();
        source_writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        source_writer
            .set_encryption("source-password", Profile::compatible())
            .unwrap();
        source_writer
            .add_file("Pictures/asset.bin", b"encrypted auxiliary bytes")
            .unwrap();
        source_writer
            .add_file("META-INF/documentsignatures.xml", b"<signatures/>")
            .unwrap();
        let source = OwnedPackage::from_bytes_with_password(
            source_writer.finish().unwrap(),
            "source-password",
        )
        .unwrap();

        let mut unencrypted = PackageWriter::new();
        unencrypted
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(unencrypted.copy_auxiliary_files_from(&source).is_err());

        let mut destination = PackageWriter::new();
        destination
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        destination
            .set_encryption("new-password", Profile::compatible())
            .unwrap();
        destination.copy_auxiliary_files_from(&source).unwrap();
        let rewritten =
            OwnedPackage::from_bytes_with_password(destination.finish().unwrap(), "new-password")
                .unwrap();
        assert_eq!(
            rewritten.get_file("Pictures/asset.bin").unwrap(),
            b"encrypted auxiliary bytes"
        );
        assert!(
            !rewritten
                .has_file("META-INF/documentsignatures.xml")
                .unwrap()
        );
    }
}
