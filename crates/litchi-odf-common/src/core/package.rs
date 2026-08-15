//! ODF package (ZIP archive) handling functionality.
//!
//! This module provides utilities for working with ODF files as ZIP archives,
//! including reading files, checking existence, and basic package operations.
//!
//! Uses soapberry-zip for high-performance zero-copy ZIP parsing.

use crate::package::{self, Archive, PreparedArchive};
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, ReadAt, Resource, ResourceLimit, Result, SourceVersion};
use soapberry_zip::office::{
    ArchiveLimits as SourceArchiveLimits, ArchiveValidationPolicy as SourceArchiveValidationPolicy,
    IndexedArchive as SourceIndexedArchive,
};
use soapberry_zip::{ErrorKind as ZipErrorKind, ReaderAt as ZipReaderAt};
use std::io::{self, Read};
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::Arc;
use zeroize::Zeroizing;

/// An ODF package (ZIP file containing XML documents).
///
/// Uses soapberry-zip for efficient lazy decompression.
#[allow(
    clippy::module_name_repetitions,
    reason = "`Package` is the established public ODF archive reader name."
)]
pub struct Package<'data> {
    archive: Archive<'data>,
    manifest: super::manifest::Manifest,
    mimetype: String,
    password: Option<&'data str>,
}

/// Owned version of [`Package`] that owns the data buffer.
#[derive(Clone)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`OwnedPackage` distinguishes the owning public archive handle."
)]
pub struct OwnedPackage {
    data: Arc<Vec<u8>>,
    index: PreparedArchive,
    password: Option<Zeroizing<String>>,
}

/// Resource limits for a positional ODF package.
///
/// `max_source_bytes` bounds the physical source before ZIP indexing starts;
/// [`Self::archive`] bounds the central-directory and declared member sizes.
/// Both limits are finite by default and are checked before payloads other
/// than the mandatory `mimetype` and manifest members are read. The mandatory
/// member ceilings are independently hard-capped so a hostile ZIP declaration
/// cannot turn package opening into an unbounded metadata allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourcePackageLimits {
    /// Maximum physical source length accepted by the positional owner.
    pub max_source_bytes: u64,
    /// ZIP index and declared member limits.
    pub archive: SourceArchiveLimits,
    /// Maximum declared and materialized `mimetype` bytes.
    pub max_mimetype_bytes: u64,
    /// Maximum declared and materialized manifest bytes.
    pub max_manifest_bytes: u64,
}

impl SourcePackageLimits {
    /// Construct a positional package policy from a source ceiling and ZIP
    /// limits.
    #[must_use]
    pub const fn new(max_source_bytes: u64, archive: SourceArchiveLimits) -> Self {
        Self {
            max_source_bytes,
            archive,
            max_mimetype_bytes: 4 * 1024,
            max_manifest_bytes: 16 * 1024 * 1024,
        }
    }

    /// Return the configured physical source ceiling.
    #[must_use]
    pub const fn max_source_bytes(self) -> u64 {
        self.max_source_bytes
    }

    /// Return the configured ZIP limits.
    #[must_use]
    pub const fn archive_limits(self) -> SourceArchiveLimits {
        self.archive
    }

    /// Tighten the physical source ceiling.
    #[must_use]
    pub const fn with_max_source_bytes(mut self, maximum: u64) -> Self {
        self.max_source_bytes = maximum;
        self
    }

    /// Set the ZIP limits; package opening validates that callers only tighten
    /// the hard default ceilings.
    #[must_use]
    pub const fn with_archive_limits(mut self, archive: SourceArchiveLimits) -> Self {
        self.archive = archive;
        self
    }

    /// Return the configured `mimetype` payload ceiling.
    #[must_use]
    pub const fn max_mimetype_bytes(self) -> u64 {
        self.max_mimetype_bytes
    }

    /// Return the configured manifest payload ceiling.
    #[must_use]
    pub const fn max_manifest_bytes(self) -> u64 {
        self.max_manifest_bytes
    }

    /// Tighten the mandatory `mimetype` payload ceiling.
    #[must_use]
    pub const fn with_max_mimetype_bytes(mut self, maximum: u64) -> Self {
        self.max_mimetype_bytes = maximum;
        self
    }

    /// Tighten the mandatory manifest payload ceiling.
    #[must_use]
    pub const fn with_max_manifest_bytes(mut self, maximum: u64) -> Self {
        self.max_manifest_bytes = maximum;
        self
    }
}

impl Default for SourcePackageLimits {
    fn default() -> Self {
        Self {
            max_source_bytes: 2 * 1024 * 1024 * 1024,
            archive: SourceArchiveLimits::default(),
            max_mimetype_bytes: 4 * 1024,
            max_manifest_bytes: 16 * 1024 * 1024,
        }
    }
}

/// An immutable, lazily read ODF package backed by a positional source.
///
/// Opening scans and validates the ZIP central directory using the strict
/// package policy, then reads only `mimetype` and the manifest. Ordinary
/// member payloads remain in the source until [`Self::get_file`] or
/// [`Self::materialize`] is requested. Every source read checks the version
/// captured at open and reports [`Error::SourceChanged`] when it no longer
/// identifies the same source snapshot.
pub struct SourceBackedPackage {
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    source_length: u64,
    limits: SourcePackageLimits,
    archive: SourceIndexedArchive<SourceReader>,
    manifest: super::manifest::Manifest,
    mimetype: String,
    password: Option<Zeroizing<String>>,
}

impl std::fmt::Debug for SourceBackedPackage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceBackedPackage")
            .field("source_version", &self.source_version)
            .field("source_length", &self.source_length)
            .field("limits", &self.limits)
            .field("file_count", &self.archive.len())
            .finish_non_exhaustive()
    }
}

#[derive(Debug)]
struct SourceChangedIo {
    expected: SourceVersion,
    observed: SourceVersion,
}

impl std::fmt::Display for SourceChangedIo {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "positional source changed (expected {:?}, observed {:?})",
            self.expected, self.observed
        )
    }
}

impl std::error::Error for SourceChangedIo {}

/// Adapter from Litchi's positional source contract to soapberry ZIP's
/// positional reader contract.
#[derive(Clone)]
struct SourceReader {
    source: Arc<dyn ReadAt>,
    expected: SourceVersion,
    length: u64,
}

impl SourceReader {
    fn check_version(&self) -> io::Result<()> {
        let observed = self.source.version()?;
        if observed != self.expected {
            return Err(io::Error::other(SourceChangedIo {
                expected: self.expected,
                observed,
            }));
        }
        Ok(())
    }
}

impl ZipReaderAt for SourceReader {
    fn read_at(&self, output: &mut [u8], offset: u64) -> io::Result<usize> {
        self.check_version()?;
        let Some(available) = self.length.checked_sub(offset) else {
            self.check_version()?;
            return Ok(0);
        };
        let available = usize::try_from(available).unwrap_or(usize::MAX);
        let requested = output.len().min(available);
        let read = self.source.read_at(offset, &mut output[..requested])?;
        if read > requested {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional source reported more bytes than requested",
            ));
        }
        self.check_version()?;
        Ok(read)
    }
}

impl SourceBackedPackage {
    /// Open a positional ODF package with the default finite limits.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, SourcePackageLimits::default())
    }

    /// Open a positional ODF package with explicit finite limits.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SourcePackageLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_password_inner(source, limits, None)
    }

    /// Open a positional ODF package and retain a password for lazy entry
    /// decryption.
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let password = Zeroizing::new(password.into());
        Self::from_read_at_with_limits_and_password_inner(
            source,
            SourcePackageLimits::default(),
            Some(password),
        )
    }

    /// Open a positional ODF package with explicit limits and a retained
    /// password for lazy entry decryption.
    pub fn from_read_at_with_limits_and_password(
        source: Arc<dyn ReadAt>,
        limits: SourcePackageLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let password = Zeroizing::new(password.into());
        Self::from_read_at_with_limits_and_password_inner(source, limits, Some(password))
    }

    fn from_read_at_with_limits_and_password_inner(
        source: Arc<dyn ReadAt>,
        limits: SourcePackageLimits,
        password: Option<Zeroizing<String>>,
    ) -> Result<Self> {
        validate_source_limits(limits)?;

        let source_version = source.version()?;
        let source_length = source.len()?;
        let observed = source.version()?;
        if observed != source_version {
            return Err(Error::SourceChanged {
                expected: source_version,
                observed,
            });
        }
        if source_length > limits.max_source_bytes {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::InputBytes,
                observed: source_length,
                limit: limits.max_source_bytes,
                scope: Arc::from("ODF positional package"),
            }));
        }

        let reader = SourceReader {
            source: Arc::clone(&source),
            expected: source_version,
            length: source_length,
        };
        let archive_result = SourceIndexedArchive::from_reader_with_limits_and_policy(
            reader,
            source_length,
            limits.archive,
            SourceArchiveValidationPolicy::StrictPackage,
        );
        let archive = match archive_result {
            Ok(archive) => archive,
            Err(error) => {
                let mapped = map_zip_error(error);
                match ensure_source_version(source.as_ref(), source_version) {
                    Err(changed @ Error::SourceChanged { .. }) => return Err(changed),
                    Err(other) => return Err(other),
                    Ok(()) => return Err(mapped),
                }
            },
        };

        ensure_source_version(source.as_ref(), source_version)?;
        let mimetype_stored = prefer_current(
            source.as_ref(),
            source_version,
            archive.is_stored("mimetype").map_err(map_zip_error),
        )?;
        if !mimetype_stored {
            return Err(Error::InvalidFormat(
                "ODF mimetype member must use ZIP Store compression".to_string(),
            ));
        }
        let mimetype = read_indexed_string_with_limit(
            &archive,
            "mimetype",
            limits.max_mimetype_bytes,
            "ODF mimetype bytes",
            source.as_ref(),
            source_version,
        )?;
        let mimetype = fallible_string(mimetype.trim(), "ODF positional package mimetype")?;
        ensure_source_version(source.as_ref(), source_version)?;

        let manifest_xml = if prefer_current(
            source.as_ref(),
            source_version,
            Ok(archive.contains("META-INF/manifest.xml")),
        )? {
            read_indexed_string_with_limit(
                &archive,
                "META-INF/manifest.xml",
                limits.max_manifest_bytes,
                "ODF manifest bytes",
                source.as_ref(),
                source_version,
            )?
        } else if prefer_current(
            source.as_ref(),
            source_version,
            Ok(archive.contains("manifest.xml")),
        )? {
            read_indexed_string_with_limit(
                &archive,
                "manifest.xml",
                limits.max_manifest_bytes,
                "ODF manifest bytes",
                source.as_ref(),
                source_version,
            )?
        } else {
            return Err(Error::InvalidFormat(
                "No manifest.xml found in ODF package".to_string(),
            ));
        };
        let manifest = prefer_current(
            source.as_ref(),
            source_version,
            super::manifest::Manifest::parse(&manifest_xml),
        )?;
        let mime_result = if manifest.mimetype == mimetype {
            Ok(())
        } else {
            Err(Error::InvalidFormat(format!(
                "ODF manifest root MIME type '{}' does not match mimetype '{}'",
                manifest.mimetype, mimetype
            )))
        };
        prefer_current(source.as_ref(), source_version, mime_result)?;

        Ok(Self {
            source,
            source_version,
            source_length,
            limits,
            archive,
            manifest,
            mimetype,
            password,
        })
    }

    /// Open a filesystem-backed ODF package without slurping the source.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(Arc::new(FileSource::open(path)?))
    }

    /// Open a filesystem-backed ODF package with explicit finite limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(
        path: impl AsRef<Path>,
        limits: SourcePackageLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits(Arc::new(FileSource::open(path)?), limits)
    }

    /// Open a filesystem-backed ODF package with a retained decryption
    /// password.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let password = Zeroizing::new(password.into());
        Self::from_read_at_with_limits_and_password_inner(
            Arc::new(FileSource::open(path)?),
            SourcePackageLimits::default(),
            Some(password),
        )
    }

    /// Open a filesystem-backed ODF package with explicit limits and password.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_password(
        path: impl AsRef<Path>,
        limits: SourcePackageLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let password = Zeroizing::new(password.into());
        Self::from_read_at_with_limits_and_password_inner(
            Arc::new(FileSource::open(path)?),
            limits,
            Some(password),
        )
    }

    /// Return the captured source version after checking that it remains
    /// current.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.ensure_current()?;
        Ok(self.source_version)
    }

    /// Return the physical source length captured at open.
    #[must_use]
    pub const fn len(&self) -> u64 {
        self.source_length
    }

    /// Return whether the captured source is empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.source_length == 0
    }

    /// Return the source package limits.
    #[must_use]
    pub const fn limits(&self) -> SourcePackageLimits {
        self.limits
    }

    /// Return the package MIME type after checking source identity.
    pub fn mimetype(&self) -> Result<&str> {
        self.ensure_current()?;
        Ok(self.mimetype.as_str())
    }

    /// Return the parsed ODF manifest after checking source identity.
    pub fn manifest(&self) -> Result<&super::manifest::Manifest> {
        self.ensure_current()?;
        Ok(&self.manifest)
    }

    /// Check whether an archive member exists.
    pub fn has_file(&self, path: &str) -> Result<bool> {
        let path = normalize_member_path(path)?;
        self.ensure_current()?;
        let present = self.archive.contains(path);
        self.ensure_current()?;
        Ok(present)
    }

    /// Check whether an archive member uses ZIP Store compression.
    pub fn is_stored(&self, path: &str) -> Result<bool> {
        let path = normalize_member_path(path)?;
        self.ensure_current()?;
        let stored = prefer_current(
            self.source.as_ref(),
            self.source_version,
            self.archive.is_stored(path).map_err(map_zip_error),
        )?;
        Ok(stored)
    }

    /// List package members without reading payload bytes.
    pub fn files(&self) -> Result<Vec<String>> {
        self.ensure_current()?;
        let mut files = Vec::new();
        files
            .try_reserve_exact(self.archive.len())
            .map_err(|error| Error::Allocation {
                resource: "ODF positional package file list",
                source: error,
            })?;
        for name in self.archive.file_names() {
            files.push(fallible_string(name, "ODF positional package file name")?);
        }
        self.ensure_current()?;
        Ok(files)
    }

    /// List embedded media members without reading their payloads.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let files = self.files()?;
        let mut media = Vec::new();
        media
            .try_reserve(files.len())
            .map_err(|error| Error::Allocation {
                resource: "ODF positional package media list",
                source: error,
            })?;
        media.extend(
            files
                .into_iter()
                .filter(|path| package::is_media_path(path)),
        );
        self.ensure_current()?;
        Ok(media)
    }

    /// Read and verify one archive member, decrypting it when its manifest
    /// entry requires a retained password.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let path = normalize_member_path(path)?;
        let bytes = self.read_entry(path)?;
        let Some(entry) = manifest_entry_for_path(&self.manifest, path)? else {
            self.ensure_current()?;
            return Ok(bytes);
        };
        let Some(encryption) = &entry.encryption else {
            self.ensure_current()?;
            return Ok(bytes);
        };
        if !self.is_stored(path)? {
            return Err(Error::InvalidFormat(format!(
                "Encrypted ODF entry '{path}' must use ZIP Store"
            )));
        }
        let password = self.password.as_ref().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Password required for encrypted ODF entry '{path}'"
            ))
        })?;
        let size = entry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Encrypted ODF entry '{path}' has no plaintext size"
            ))
        })?;
        let decrypted = super::encryption::decrypt_entry(&bytes, password, encryption, size);
        match self.ensure_current() {
            Err(changed @ Error::SourceChanged { .. }) => Err(changed),
            Err(other) => Err(other),
            Ok(()) => decrypted,
        }
    }

    /// Materialize an exact, source-checked owned package.
    ///
    /// This is the explicit transition from positional lazy access to the
    /// existing mutable/serialization owner. It reads exactly the captured
    /// source length, preserves the retained password credential, and
    /// revalidates source identity before returning.
    pub fn materialize(&self) -> Result<OwnedPackage> {
        self.ensure_current()?;
        let length = usize::try_from(self.source_length).map_err(|_| {
            Error::InvalidFormat("ODF positional source exceeds platform limits".to_string())
        })?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|error| Error::Allocation {
                resource: "ODF positional source materialization",
                source: error,
            })?;
        bytes.resize(length, 0);
        if let Err(error) = self.source.read_exact_at(0, &mut bytes) {
            match self.ensure_current() {
                Err(changed @ Error::SourceChanged { .. }) => return Err(changed),
                Err(other) => return Err(other),
                Ok(()) => return Err(error.into()),
            }
        }
        self.ensure_current()?;
        let package = prefer_current(
            self.source.as_ref(),
            self.source_version,
            OwnedPackage::from_shared_bytes_with_strict_policy(Arc::new(bytes)),
        )?;
        let mut package = package;
        package.password = self.password.clone();
        self.ensure_current()?;
        Ok(package)
    }

    /// Read inert document and macro signature metadata.
    pub fn digital_signatures(&self) -> Result<crate::signature::DigitalSignatures> {
        use crate::signature::{
            DOCUMENT_SIGNATURE_PATH, MACRO_SIGNATURE_PATH, parse_signature_container,
        };

        let mut result = crate::signature::DigitalSignatures::default();
        if self.has_file(DOCUMENT_SIGNATURE_PATH)? {
            result.document_signatures =
                parse_signature_container(&self.get_file(DOCUMENT_SIGNATURE_PATH)?)?;
        }
        if self.has_file(MACRO_SIGNATURE_PATH)? {
            result.macro_signatures =
                parse_signature_container(&self.get_file(MACRO_SIGNATURE_PATH)?)?;
        }
        self.ensure_current()?;
        Ok(result)
    }

    fn ensure_current(&self) -> Result<()> {
        ensure_source_version(self.source.as_ref(), self.source_version)
    }

    fn read_entry(&self, path: &str) -> Result<Vec<u8>> {
        self.ensure_current()?;
        let result = self.archive.read(path).map_err(map_zip_error);
        match result {
            Ok(bytes) => {
                self.ensure_current()?;
                Ok(bytes)
            },
            Err(error) => match self.ensure_current() {
                Err(changed @ Error::SourceChanged { .. }) => Err(changed),
                Err(other) => Err(other),
                Ok(()) => Err(error),
            },
        }
    }
}

fn ensure_source_version(source: &dyn ReadAt, expected: SourceVersion) -> Result<()> {
    let observed = source.version()?;
    if observed != expected {
        return Err(Error::SourceChanged { expected, observed });
    }
    Ok(())
}

fn normalize_member_path(path: &str) -> Result<&str> {
    let normalized = path.strip_prefix('/').unwrap_or(path);
    if normalized.is_empty()
        || normalized.starts_with('/')
        || normalized.contains('\\')
        || normalized
            .as_bytes()
            .iter()
            .any(|byte| byte.is_ascii_control())
        || normalized
            .split('/')
            .any(|segment| segment.is_empty() || segment == "." || segment == "..")
    {
        return Err(Error::InvalidFormat(format!(
            "unsafe ODF package member path '{path}'"
        )));
    }
    Ok(normalized)
}

fn manifest_entry_for_path<'manifest>(
    manifest: &'manifest super::manifest::Manifest,
    path: &str,
) -> Result<Option<&'manifest super::manifest::ManifestEntry>> {
    let mut found = None;
    for (candidate, entry) in &manifest.entries {
        let matches = candidate == path
            || candidate
                .strip_prefix('/')
                .is_some_and(|candidate| !candidate.is_empty() && candidate == path);
        if matches {
            if found.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "manifest contains ambiguous aliases for '{path}'"
                )));
            }
            found = Some(entry);
        }
    }
    Ok(found)
}

fn fallible_string(value: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

fn validate_source_limits(limits: SourcePackageLimits) -> Result<()> {
    let hard = SourcePackageLimits::default();
    validate_limit(
        "ODF positional source bytes",
        limits.max_source_bytes,
        hard.max_source_bytes,
    )?;
    validate_limit(
        "ODF mimetype bytes",
        limits.max_mimetype_bytes,
        hard.max_mimetype_bytes,
    )?;
    validate_limit(
        "ODF manifest bytes",
        limits.max_manifest_bytes,
        hard.max_manifest_bytes,
    )?;
    if limits.archive.max_files == 0 || limits.archive.max_files > hard.archive.max_files {
        return Err(Error::InvalidFormat(format!(
            "ODF archive max_files must be between 1 and {}",
            hard.archive.max_files
        )));
    }
    validate_limit(
        "ODF archive max_member_name_bytes",
        limits.archive.max_member_name_bytes,
        hard.archive.max_member_name_bytes,
    )?;
    validate_limit(
        "ODF archive max_metadata_bytes",
        limits.archive.max_metadata_bytes,
        hard.archive.max_metadata_bytes,
    )?;
    validate_limit(
        "ODF archive max_compressed_size",
        limits.archive.max_compressed_size,
        hard.archive.max_compressed_size,
    )?;
    validate_limit(
        "ODF archive max_entry_size",
        limits.archive.max_entry_size,
        hard.archive.max_entry_size,
    )?;
    validate_limit(
        "ODF archive max_total_size",
        limits.archive.max_total_size,
        hard.archive.max_total_size,
    )?;
    Ok(())
}

fn validate_limit(name: &str, actual: u64, maximum: u64) -> Result<()> {
    if actual == 0 || actual > maximum {
        return Err(Error::InvalidFormat(format!(
            "{name} must be between 1 and {maximum}"
        )));
    }
    Ok(())
}

fn prefer_current<T>(source: &dyn ReadAt, expected: SourceVersion, result: Result<T>) -> Result<T> {
    // The adapter checks the version before and after each physical read, but
    // a high-level operation can still return an arbitrary ZIP/parser error
    // after that read. Re-check after the operation so a stale source always
    // wins over the secondary error.
    match ensure_source_version(source, expected) {
        Err(error) => Err(error),
        Ok(()) => result,
    }
}

fn read_indexed_string_with_limit<R: ZipReaderAt>(
    archive: &SourceIndexedArchive<R>,
    path: &str,
    maximum: u64,
    resource: &'static str,
    source: &dyn ReadAt,
    expected: SourceVersion,
) -> Result<String> {
    let metadata = prefer_current(
        source,
        expected,
        archive.metadata(path).map_err(map_zip_error),
    )?;
    if metadata.uncompressed_size() > maximum {
        return Err(Error::ResourceLimit(ResourceLimit {
            resource: Resource::InputBytes,
            observed: metadata.uncompressed_size(),
            limit: maximum,
            scope: Arc::from(resource),
        }));
    }
    let text = archive.read(path).map_err(map_zip_error).and_then(|bytes| {
        String::from_utf8(bytes)
            .map_err(|error| Error::InvalidFormat(format!("Invalid UTF-8 in '{path}': {error}")))
    });
    prefer_current(source, expected, text)
}

fn map_zip_error(error: soapberry_zip::Error) -> Error {
    match error.kind() {
        ZipErrorKind::LimitExceeded {
            resource,
            actual,
            maximum,
        } => zip_limit_error(*resource, *actual, *maximum),
        ZipErrorKind::IO(io_error) | ZipErrorKind::Io(io_error) => {
            if let Some(source) = io_error
                .get_ref()
                .and_then(|source| source.downcast_ref::<SourceChangedIo>())
            {
                Error::SourceChanged {
                    expected: source.expected,
                    observed: source.observed,
                }
            } else {
                Error::Io(io::Error::new(io_error.kind(), io_error.to_string()))
            }
        },
        _ => Error::InvalidFormat(error.to_string()),
    }
}

fn zip_limit_error(resource: soapberry_zip::LimitResource, actual: u64, maximum: u64) -> Error {
    let (dimension, scope) = match resource {
        soapberry_zip::LimitResource::FileCount => (Resource::Objects, "ODF ZIP file count"),
        soapberry_zip::LimitResource::MemberNameBytes => {
            (Resource::InputBytes, "ODF ZIP member name bytes")
        },
        soapberry_zip::LimitResource::MetadataBytes => {
            (Resource::InputBytes, "ODF ZIP metadata bytes")
        },
        soapberry_zip::LimitResource::CompressedSize => {
            (Resource::InputBytes, "ODF ZIP compressed bytes")
        },
        soapberry_zip::LimitResource::EntrySize => (Resource::InputBytes, "ODF ZIP entry bytes"),
        soapberry_zip::LimitResource::TotalSize => (Resource::InputBytes, "ODF ZIP total bytes"),
    };
    Error::ResourceLimit(ResourceLimit {
        resource: dimension,
        observed: actual,
        limit: maximum,
        scope: Arc::from(scope),
    })
}

fn map_owned_zip_error(error: soapberry_zip::Error) -> Error {
    match map_zip_error(error) {
        limit @ Error::ResourceLimit(_) => limit,
        error => Error::InvalidFormat(format!("Invalid ZIP archive: {error}")),
    }
}

impl OwnedPackage {
    /// Open an ODF package from a reader.
    ///
    /// # Errors
    ///
    /// Returns an error when reading the input or parsing the ZIP archive
    /// fails.
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes(data)
    }

    /// Create an ODF package from bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes do not form a valid ZIP archive.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Self::from_shared_bytes(Arc::new(data))
    }

    /// Adopt shared ODF package bytes without copying the archive buffer.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes do not form a valid ZIP archive.
    pub fn from_shared_bytes(data: Arc<Vec<u8>>) -> Result<Self> {
        Self::from_shared_bytes_with_policy(
            data,
            soapberry_zip::office::ArchiveValidationPolicy::Normalized,
        )
    }

    pub(crate) fn from_prepared_bytes_or_recover(
        data: Vec<u8>,
    ) -> std::result::Result<Self, Vec<u8>> {
        let data = Arc::new(data);
        match Self::from_shared_bytes_with_strict_policy(Arc::clone(&data)) {
            Ok(package) => Ok(package),
            Err(_error) => {
                let Some(data) = Arc::into_inner(data) else {
                    unreachable!("failed ODF preparation must release its temporary source handle")
                };
                Err(data)
            },
        }
    }

    fn from_shared_bytes_with_policy(
        data: Arc<Vec<u8>>,
        policy: soapberry_zip::office::ArchiveValidationPolicy,
    ) -> Result<Self> {
        #[cfg(test)]
        package::note_index_build();
        let index = Arc::new(
            soapberry_zip::office::IndexedArchive::from_reader_with_limits_and_policy(
                Arc::clone(&data),
                u64::try_from(data.len()).map_err(|_| {
                    Error::InvalidFormat("ODF package length exceeds ZIP reader limits".to_string())
                })?,
                soapberry_zip::office::ArchiveLimits::default(),
                policy,
            )
            .map_err(map_owned_zip_error)?,
        );
        Ok(Self {
            data,
            index,
            password: None,
        })
    }

    /// Build a strict prepared archive through the fallible ZIP index path.
    ///
    /// Positional materialization uses this explicit constructor so the
    /// strict package policy and structured resource-limit errors are retained
    /// at the transition to the owning representation.
    pub(crate) fn from_shared_bytes_with_strict_policy(data: Arc<Vec<u8>>) -> Result<Self> {
        Self::from_shared_bytes_with_policy(
            data,
            soapberry_zip::office::ArchiveValidationPolicy::StrictPackage,
        )
    }

    /// Open an ODF package and retain a password for lazy entry decryption.
    ///
    /// # Errors
    ///
    /// Returns an error when reading the input or parsing the ZIP archive
    /// fails.
    pub fn from_reader_with_password<R: Read>(
        mut reader: R,
        password: impl Into<String>,
    ) -> Result<Self> {
        let password = Zeroizing::new(password.into());
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes_with_zeroizing_password(data, password)
    }

    /// Open ODF bytes and retain a password for lazy entry decryption.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes do not form a valid ZIP archive.
    pub fn from_bytes_with_password(data: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        let password = Zeroizing::new(password.into());
        Self::from_bytes_with_zeroizing_password(data, password)
    }

    fn from_bytes_with_zeroizing_password(
        data: Vec<u8>,
        password: Zeroizing<String>,
    ) -> Result<Self> {
        let mut package = Self::from_shared_bytes(Arc::new(data))?;
        package.password = Some(password);
        Ok(package)
    }

    /// Get a borrowed [`Package`] for accessing archive contents.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive MIME type or manifest cannot be
    /// decoded.
    pub fn package(&self) -> Result<Package<'_>> {
        Package::new_with_prepared(
            Arc::clone(&self.index),
            self.password.as_ref().map(|password| password.as_str()),
        )
    }

    /// Get the underlying data.
    #[must_use]
    pub fn into_inner(self) -> Vec<u8> {
        let Self {
            data,
            index,
            password: _,
        } = self;
        match Arc::try_unwrap(index) {
            Ok(index) => {
                drop(data);
                match Arc::try_unwrap(index.into_zip_archive().into_inner()) {
                    Ok(data) => data,
                    Err(data) => (*data).clone(),
                }
            },
            Err(index) => {
                drop(index);
                Arc::try_unwrap(data).unwrap_or_else(|data| (*data).clone())
            },
        }
    }

    /// Get a reference to the underlying data.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// Clone the shared handle to the exact archive allocation.
    #[must_use]
    pub fn shared_bytes(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.data)
    }

    /// Clone the immutable archive/index handles without retaining a
    /// decryption credential. Snapshot owners use this when the credential is
    /// intentionally scoped to the opening document and must not become part
    /// of a public, cloneable snapshot.
    #[doc(hidden)]
    pub fn clone_without_password(&self) -> Self {
        Self {
            data: Arc::clone(&self.data),
            index: Arc::clone(&self.index),
            password: None,
        }
    }

    /// Return a stable identity for the retained archive index.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        Arc::as_ptr(&self.index) as usize
    }

    // Convenience methods that delegate to Package

    /// Get the MIME type from the mimetype file.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn mimetype(&self) -> Result<String> {
        let package = self.package()?;
        Ok(package.mimetype().to_string())
    }

    /// Get a file from the package by path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata, requested entry, or entry
    /// decryption is invalid.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let package = self.package()?;
        package.get_file(path)
    }

    /// Check if a file exists in the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn has_file(&self, path: &str) -> Result<bool> {
        let package = self.package()?;
        Ok(package.has_file(path))
    }

    /// Check whether a package member uses ZIP Store compression.
    pub fn is_stored(&self, path: &str) -> Result<bool> {
        self.index
            .is_stored(path)
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    /// List all files in the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.files()
    }

    /// Get all embedded media files from the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.media_files()
    }

    /// Read inert document and macro signature metadata from the package.
    ///
    /// This does not verify cryptographic signatures or execute macro content.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or its signature metadata cannot be
    /// decoded.
    pub fn digital_signatures(&self) -> Result<crate::signature::DigitalSignatures> {
        self.package()?.digital_signatures()
    }

    /// Cryptographically verify document signatures without making any PKI trust claim.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or its signature metadata cannot be
    /// decoded or verified.
    pub fn verify_document_signatures(
        &self,
    ) -> Result<Vec<crate::signature::SignatureVerification>> {
        crate::signature::verify_package(&self.data)
    }
}

impl<'data> Package<'data> {
    /// Create a new [`Package`] from a byte slice.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive MIME type or manifest cannot be
    /// decoded.
    pub fn new(data: &'data [u8]) -> Result<Self> {
        Self::new_with_password(data, None)
    }

    fn new_with_password(data: &'data [u8], password: Option<&'data str>) -> Result<Self> {
        let archive = Archive::new(data)?;

        Self::from_archive(archive, password)
    }

    fn new_with_prepared(index: PreparedArchive, password: Option<&'data str>) -> Result<Self> {
        let archive = Archive::from_prepared(index);

        Self::from_archive(archive, password)
    }

    fn from_archive(archive: Archive<'data>, password: Option<&'data str>) -> Result<Self> {
        // Read MIME type from mimetype file
        let mimetype = archive
            .read_string("mimetype")
            .map_err(|error| {
                Error::InvalidFormat(format!("No mimetype file found in ODF package: {error}"))
            })?
            .trim()
            .to_string();

        // Parse the manifest
        let manifest = super::manifest::Manifest::parse(&archive.read_manifest_xml()?)?;

        Ok(Self {
            archive,
            manifest,
            mimetype,
            password,
        })
    }

    /// Get the MIME type from the mimetype file.
    #[must_use]
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Get a file from the package by path.
    ///
    /// # Errors
    ///
    /// Returns an error when the entry does not exist, does not meet encrypted
    /// package requirements, or cannot be decrypted.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let path = normalize_member_path(path)?;
        let bytes = self
            .archive
            .read(path)
            .map_err(|error| Error::InvalidFormat(format!("File not found: {path}: {error}")))?;
        let Some(entry) = manifest_entry_for_path(&self.manifest, path)? else {
            return Ok(bytes);
        };
        let Some(encryption) = &entry.encryption else {
            return Ok(bytes);
        };
        if !self.archive.is_stored(path).map_err(|error| {
            Error::InvalidFormat(format!(
                "Unable to inspect encrypted ODF entry '{path}': {error}"
            ))
        })? {
            return Err(Error::InvalidFormat(format!(
                "Encrypted ODF entry '{path}' must use ZIP Store"
            )));
        }
        let password = self.password.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Password required for encrypted ODF entry '{path}'"
            ))
        })?;
        let size = entry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Encrypted ODF entry '{path}' has no plaintext size"
            ))
        })?;
        super::encryption::decrypt_entry(&bytes, password, encryption, size)
    }

    /// Check if a file exists in the package.
    #[must_use]
    pub fn has_file(&self, path: &str) -> bool {
        self.archive.contains(path)
    }

    /// Check whether a package member uses ZIP Store compression.
    pub fn is_stored(&self, path: &str) -> Result<bool> {
        self.archive.is_stored(path)
    }

    /// Get the manifest.
    #[must_use]
    pub fn manifest(&self) -> &super::manifest::Manifest {
        &self.manifest
    }

    /// List all files in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the archive cannot enumerate its entries.
    pub fn files(&self) -> Result<Vec<String>> {
        Ok(self.archive.file_names().map(String::from).collect())
    }

    /// Get all embedded media files (images, etc.) from the package.
    ///
    /// This returns paths to all files in the Pictures/ directory and other media directories.
    ///
    /// # Errors
    ///
    /// Returns an error if the archive cannot enumerate its entries.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let all_files = self.files()?;
        Ok(all_files
            .into_iter()
            .filter(|path| package::is_media_path(path))
            .collect())
    }

    /// Check if the package contains any media files.
    ///
    /// # Errors
    ///
    /// Returns an error if the archive cannot enumerate its entries.
    pub fn has_media(&self) -> Result<bool> {
        Ok(!self.media_files()?.is_empty())
    }

    /// Read inert document and macro signature metadata from the package.
    ///
    /// This does not verify cryptographic signatures or execute macro content.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or its signature metadata cannot be
    /// decoded.
    pub fn digital_signatures(&self) -> Result<crate::signature::DigitalSignatures> {
        use crate::signature::{
            DOCUMENT_SIGNATURE_PATH, MACRO_SIGNATURE_PATH, parse_signature_container,
        };

        let mut result = crate::signature::DigitalSignatures::default();
        if self.has_file(DOCUMENT_SIGNATURE_PATH) {
            result.document_signatures =
                parse_signature_container(&self.get_file(DOCUMENT_SIGNATURE_PATH)?)?;
        }
        if self.has_file(MACRO_SIGNATURE_PATH) {
            result.macro_signatures =
                parse_signature_container(&self.get_file(MACRO_SIGNATURE_PATH)?)?;
        }
        Ok(result)
    }
}

impl package::PackageLookup for Package<'_> {
    fn has_file(&self, path: &str) -> bool {
        self.has_file(path)
    }

    fn media_type(&self, path: &str) -> Option<&str> {
        self.manifest.get_media_type(path)
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible ZIP setup operations so assertions can focus on package behavior."
)]
mod tests {
    use super::*;
    use std::io::Cursor;
    use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

    // Helper function to create a minimal ODF package (ZIP with mimetype and manifest)
    fn create_test_odf_package(mimetype: &str) -> Vec<u8> {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            // Write mimetype file (must be first and uncompressed for ODF)
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(mimetype.as_bytes()).unwrap();

            // Write manifest.xml
            let manifest_xml = format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/>
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Pictures/image.png" manifest:media-type="image/png"/>
</manifest:manifest>"#
            );
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest_xml.as_bytes()).unwrap();

            // Write content.xml
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(b"<office:document-content/>").unwrap();

            // Write styles.xml
            zip.start_file("styles.xml", options).unwrap();
            zip.write_all(b"<office:document-styles/>").unwrap();

            // Write a picture
            zip.start_file("Pictures/image.png", options).unwrap();
            zip.write_all(b"PNG\x89\x50\x4e\x47\x0d\x0a\x1a\x0a")
                .unwrap();

            zip.finish().unwrap();
        }
        zip_buffer
    }

    fn create_manifest_location_package(
        canonical: Option<&[u8]>,
        legacy: Option<&[u8]>,
    ) -> Vec<u8> {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            if let Some(manifest) = canonical {
                zip.start_file("META-INF/manifest.xml", options).unwrap();
                zip.write_all(manifest).unwrap();
            }
            if let Some(manifest) = legacy {
                zip.start_file("manifest.xml", options).unwrap();
                zip.write_all(manifest).unwrap();
            }
            zip.finish().unwrap();
        }
        zip_buffer
    }

    fn valid_manifest_xml(extra_entry: Option<&str>) -> Vec<u8> {
        let extra_entry = extra_entry.map_or_else(String::new, |path| {
            format!(r#"<m:file-entry m:full-path="{path}" m:media-type="text/plain"/>"#)
        });
        format!(
            r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/>{extra_entry}</m:manifest>"#
        )
        .into_bytes()
    }

    fn create_test_ods_package() -> Vec<u8> {
        create_test_odf_package("application/vnd.oasis.opendocument.spreadsheet")
    }

    fn create_test_odp_package() -> Vec<u8> {
        create_test_odf_package("application/vnd.oasis.opendocument.presentation")
    }

    struct CountingSource {
        bytes: Arc<Vec<u8>>,
        reads: AtomicUsize,
        revision: AtomicU64,
        mutate_on_read: AtomicBool,
        forbidden_range: std::sync::Mutex<Option<(u64, u64)>>,
        forbidden_reads: AtomicUsize,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Arc<Self> {
            Arc::new(Self {
                bytes: Arc::new(bytes),
                reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
                mutate_on_read: AtomicBool::new(false),
                forbidden_range: std::sync::Mutex::new(None),
                forbidden_reads: AtomicUsize::new(0),
            })
        }

        fn read_count(&self) -> usize {
            self.reads.load(Ordering::Relaxed)
        }

        fn bump_version(&self) {
            self.revision.fetch_add(1, Ordering::Relaxed);
        }

        fn mutate_next_read(&self) {
            self.mutate_on_read.store(true, Ordering::Relaxed);
        }

        fn forbid_range(&self, range: (u64, u64)) {
            *self.forbidden_range.lock().unwrap() = Some(range);
        }

        fn forbidden_read_count(&self) -> usize {
            self.forbidden_reads.load(Ordering::Relaxed)
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::Relaxed);
            if let Some((start, end)) = *self.forbidden_range.lock().unwrap() {
                let requested_end = offset.saturating_add(output.len() as u64);
                if offset < end && requested_end > start {
                    self.forbidden_reads.fetch_add(1, Ordering::Relaxed);
                    return Err(io::Error::other("test payload range was read"));
                }
            }
            let Ok(start) = usize::try_from(offset) else {
                return Ok(0);
            };
            let Some(input) = self.bytes.get(start..) else {
                return Ok(0);
            };
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            if self.mutate_on_read.swap(false, Ordering::Relaxed) {
                self.bump_version();
            }
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x4f44_462d_5445_5354,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }

    #[test]
    fn source_backed_open_is_lazy_for_unselected_payloads() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let source = CountingSource::new(data);
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let after_open = source.read_count();

        assert!(package.has_file("Pictures/image.png").unwrap());
        assert_eq!(source.read_count(), after_open);
        assert_eq!(
            package.get_file("Pictures/image.png").unwrap(),
            b"PNG\x89\x50\x4e\x47\x0d\x0a\x1a\x0a"
        );
        assert!(source.read_count() > after_open);
    }

    #[test]
    fn source_backed_selected_read_defers_crc_validation() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let marker = b"PNG\x89\x50\x4e\x47\x0d\x0a\x1a\x0a";
        let offset = data
            .windows(marker.len())
            .position(|window| window == marker)
            .unwrap();
        let mut corrupted = data;
        corrupted[offset] ^= 0x01;

        let source = CountingSource::new(corrupted);
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        assert!(package.get_file("Pictures/image.png").is_err());
    }

    #[test]
    fn source_backed_rejects_stale_sources_before_and_during_reads() {
        let source = CountingSource::new(create_test_odf_package(
            "application/vnd.oasis.opendocument.text",
        ));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();

        source.bump_version();
        assert!(matches!(
            package.get_file("content.xml"),
            Err(Error::SourceChanged { .. })
        ));

        let source = CountingSource::new(create_test_odf_package(
            "application/vnd.oasis.opendocument.text",
        ));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        source.mutate_next_read();
        assert!(matches!(
            package.get_file("content.xml"),
            Err(Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn source_backed_enforces_exact_source_ceiling_and_materializes() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let source = CountingSource::new(data.clone());
        let limits = SourcePackageLimits::new(
            u64::try_from(data.len() - 1).unwrap(),
            SourceArchiveLimits::default(),
        );
        assert!(SourceBackedPackage::from_read_at_with_limits(source, limits).is_err());

        let source = CountingSource::new(data.clone());
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let materialized = package.materialize().unwrap();
        assert_eq!(materialized.as_bytes(), data.as_slice());
    }

    #[test]
    fn source_backed_rejects_unbounded_or_oversized_policies_before_reads() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let source = CountingSource::new(data.clone());
        let limits = SourcePackageLimits::new(
            u64::try_from(data.len()).unwrap(),
            SourceArchiveLimits::UNBOUNDED,
        );
        assert!(SourceBackedPackage::from_read_at_with_limits(source.clone(), limits).is_err());
        assert_eq!(source.read_count(), 0);

        let source = CountingSource::new(data);
        let limits = SourcePackageLimits::new(
            SourcePackageLimits::default().max_source_bytes + 1,
            SourceArchiveLimits::default(),
        );
        assert!(SourceBackedPackage::from_read_at_with_limits(source.clone(), limits).is_err());
        assert_eq!(source.read_count(), 0);
    }

    #[test]
    fn source_backed_rejects_oversized_mimetype_before_payload_read() {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(&vec![b'x'; 4 * 1024 + 1]).unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(
                br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/></m:manifest>"#,
            )
            .unwrap();
            zip.start_file("Pictures/pad.bin", options).unwrap();
            zip.write_all(&vec![0_u8; 128 * 1024]).unwrap();
            zip.finish().unwrap();
        }
        let archive = soapberry_zip::ZipArchive::from_slice(&zip_buffer).unwrap();
        let mimetype_range = archive
            .entries()
            .filter_map(|entry| {
                let entry = entry.ok()?;
                (entry.file_path().as_ref() == b"mimetype").then(|| {
                    archive
                        .get_entry(entry.wayfinder())
                        .unwrap()
                        .compressed_data_range()
                })
            })
            .next()
            .unwrap();

        let source = CountingSource::new(zip_buffer);
        source.forbid_range(mimetype_range);
        let error = SourceBackedPackage::from_read_at(source.clone()).unwrap_err();
        assert!(matches!(error, Error::ResourceLimit(_)));
        assert_eq!(source.forbidden_read_count(), 0);
    }

    #[test]
    fn source_backed_strict_policy_requires_offset_zero_mimetype() {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(
                br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/></m:manifest>"#,
            )
            .unwrap();
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.finish().unwrap();
        }

        let source = CountingSource::new(zip_buffer);
        assert!(SourceBackedPackage::from_read_at(source).is_err());
    }

    #[test]
    fn source_backed_manifest_fallback_requires_canonical_absence() {
        let legacy = valid_manifest_xml(None);

        let source = CountingSource::new(create_manifest_location_package(
            Some(b"<not-a-manifest>"),
            Some(&legacy),
        ));
        assert!(matches!(
            SourceBackedPackage::from_read_at(source),
            Err(Error::InvalidFormat(_))
        ));

        let canonical = valid_manifest_xml(Some("CANONICAL_MARKER"));
        let mut corrupted = create_manifest_location_package(Some(&canonical), Some(&legacy));
        let marker_offset = corrupted
            .windows(b"CANONICAL_MARKER".len())
            .position(|window| window == b"CANONICAL_MARKER")
            .unwrap();
        corrupted[marker_offset] ^= 1;
        let source = CountingSource::new(corrupted);
        let error = SourceBackedPackage::from_read_at(source).unwrap_err();
        assert!(matches!(error, Error::Io(_) | Error::InvalidFormat(_)));

        let oversized =
            vec![b'x'; SourcePackageLimits::default().max_manifest_bytes() as usize + 1];
        let source = CountingSource::new(create_manifest_location_package(
            Some(&oversized),
            Some(&legacy),
        ));
        assert!(matches!(
            SourceBackedPackage::from_read_at(source),
            Err(Error::ResourceLimit(_))
        ));
    }

    #[test]
    fn source_backed_manifest_fallback_accepts_legacy_only_location() {
        let legacy = valid_manifest_xml(None);
        let source = CountingSource::new(create_manifest_location_package(None, Some(&legacy)));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.text"
        );
        assert_eq!(
            package.manifest().unwrap().mimetype,
            package.mimetype().unwrap()
        );
    }

    #[test]
    fn source_backed_maps_each_zip_limit_to_structured_resource_limit() {
        for resource in [
            soapberry_zip::LimitResource::FileCount,
            soapberry_zip::LimitResource::MemberNameBytes,
            soapberry_zip::LimitResource::MetadataBytes,
            soapberry_zip::LimitResource::CompressedSize,
            soapberry_zip::LimitResource::EntrySize,
            soapberry_zip::LimitResource::TotalSize,
        ] {
            let error = map_zip_error(soapberry_zip::Error::from(ZipErrorKind::LimitExceeded {
                resource,
                actual: 7,
                maximum: 6,
            }));
            assert!(matches!(
                error,
                Error::ResourceLimit(ResourceLimit {
                    observed: 7,
                    limit: 6,
                    ..
                })
            ));
        }
    }

    #[test]
    fn source_backed_preserves_ordinary_zip_io_errors() {
        let error = map_zip_error(soapberry_zip::Error::from(io::Error::new(
            io::ErrorKind::UnexpectedEof,
            "positional read failed",
        )));
        assert!(matches!(
            error,
            Error::Io(error) if error.kind() == io::ErrorKind::UnexpectedEof
        ));
    }

    #[test]
    fn source_backed_normalizes_safe_leading_slash_for_archive_and_manifest() {
        assert_eq!(
            normalize_member_path("/content.xml").unwrap(),
            "content.xml"
        );
        assert!(normalize_member_path("/../content.xml").is_err());
        assert!(normalize_member_path("content\\.xml").is_err());
    }

    #[test]
    fn manifest_aliases_cannot_bypass_encryption_metadata() {
        use std::io::Write;

        for alias in [
            "/content.xml",
            "./content.xml",
            "foo/../content.xml",
            "content%2Exml",
            "C:content.xml",
            "foo:content.xml",
        ] {
            let manifest = format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="{alias}" m:media-type="text/xml" m:size="1"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="http://www.w3.org/2001/04/xmlenc#sha256" m:key-size="32"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#
            );
            let mut bytes = Vec::new();
            {
                let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
                let options = zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Stored);
                zip.start_file("mimetype", options).unwrap();
                zip.write_all(b"application/vnd.oasis.opendocument.text")
                    .unwrap();
                zip.start_file("META-INF/manifest.xml", options).unwrap();
                zip.write_all(manifest.as_bytes()).unwrap();
                zip.start_file("content.xml", options).unwrap();
                zip.write_all(&[0_u8; 32]).unwrap();
                zip.finish().unwrap();
            }

            assert!(SourceBackedPackage::from_read_at(CountingSource::new(bytes.clone())).is_err());
            assert!(
                SourceBackedPackage::from_read_at_with_password(
                    CountingSource::new(bytes.clone()),
                    "wrong",
                )
                .is_err()
            );
            let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
            assert!(package.package().is_err());
            let package = OwnedPackage::from_bytes_with_password(bytes, "wrong").unwrap();
            assert!(package.package().is_err());
        }
    }

    #[test]
    fn test_owned_package_from_bytes() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data);
        assert!(package.is_ok());
    }

    #[test]
    fn test_owned_package_from_reader() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let cursor = Cursor::new(data);
        let package = OwnedPackage::from_reader(cursor);
        assert!(package.is_ok());
    }

    #[test]
    fn test_owned_package_invalid_data() {
        let invalid_data = b"not a zip file".to_vec();
        let result = OwnedPackage::from_bytes(invalid_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_owned_package_into_inner() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data.clone()).unwrap();
        let inner = package.into_inner();
        assert!(!inner.is_empty());
    }

    #[test]
    fn test_owned_package_as_bytes() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data.clone()).unwrap();
        let bytes = package.as_bytes();
        assert!(!bytes.is_empty());
    }

    #[test]
    fn test_owned_package_mimetype() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.text"
        );
    }

    #[test]
    fn test_owned_package_mimetype_ods() {
        let data = create_test_ods_package();
        let package = OwnedPackage::from_bytes(data).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.spreadsheet"
        );
    }

    #[test]
    fn test_owned_package_mimetype_odp() {
        let data = create_test_odp_package();
        let package = OwnedPackage::from_bytes(data).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.presentation"
        );
    }

    #[test]
    fn test_owned_package_get_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let content = package.get_file("content.xml");
        assert!(content.is_ok());
        assert_eq!(content.unwrap(), b"<office:document-content/>");
    }

    #[test]
    fn test_owned_package_get_file_not_found() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let result = package.get_file("nonexistent.xml");
        assert!(result.is_err());
    }

    #[test]
    fn test_owned_package_has_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        assert!(package.has_file("content.xml").unwrap());
        assert!(package.has_file("styles.xml").unwrap());
        assert!(!package.has_file("nonexistent.xml").unwrap());
    }

    #[test]
    fn test_owned_package_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let files = package.files().unwrap();
        assert!(files.contains(&"mimetype".to_string()));
        assert!(files.contains(&"content.xml".to_string()));
        assert!(files.contains(&"styles.xml".to_string()));
        assert!(files.contains(&"META-INF/manifest.xml".to_string()));
        assert!(files.contains(&"Pictures/image.png".to_string()));
    }

    #[test]
    fn test_owned_package_media_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let media_files = package.media_files().unwrap();
        assert!(media_files.contains(&"Pictures/image.png".to_string()));
    }

    #[test]
    fn test_package_new() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data);
        assert!(package.is_ok());
    }

    #[test]
    fn test_package_new_invalid_data() {
        let invalid_data = b"not a zip file";
        let result = Package::new(invalid_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_package_mimetype() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();
        assert_eq!(
            package.mimetype(),
            "application/vnd.oasis.opendocument.text"
        );
    }

    #[test]
    fn test_package_get_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let content = package.get_file("content.xml").unwrap();
        assert_eq!(content, b"<office:document-content/>");
    }

    #[test]
    fn test_package_get_file_not_found() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let result = package.get_file("nonexistent.xml");
        assert!(result.is_err());
    }

    #[test]
    fn test_package_has_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        assert!(package.has_file("content.xml"));
        assert!(!package.has_file("nonexistent.xml"));
    }

    #[test]
    fn test_package_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let files = package.files().unwrap();
        assert!(!files.is_empty());
        assert!(files.contains(&"content.xml".to_string()));
    }

    #[test]
    fn test_package_media_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let media_files = package.media_files().unwrap();
        assert!(media_files.contains(&"Pictures/image.png".to_string()));
    }

    #[test]
    fn test_package_has_media() -> Result<()> {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data)?;

        assert!(package.has_media()?);
        Ok(())
    }

    #[test]
    fn test_package_manifest() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let manifest = package.manifest();
        assert_eq!(manifest.mimetype, "application/vnd.oasis.opendocument.text");
    }

    #[test]
    fn test_owned_package_package_method() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let owned = OwnedPackage::from_bytes(data).unwrap();

        let package = owned.package();
        assert!(package.is_ok());
        assert_eq!(
            package.unwrap().mimetype(),
            "application/vnd.oasis.opendocument.text"
        );
    }

    #[test]
    fn test_package_media_files_various_formats() {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            // Write mimetype file
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();

            // Write manifest.xml
            let manifest_xml = r#"<?xml version="1.0"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
</manifest:manifest>"#;
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest_xml.as_bytes()).unwrap();

            // Write various media files
            zip.start_file("Pictures/photo.jpg", options).unwrap();
            zip.write_all(b"fake jpg data").unwrap();

            zip.start_file("Pictures/chart.jpeg", options).unwrap();
            zip.write_all(b"fake jpeg data").unwrap();

            zip.start_file("media/animation.gif", options).unwrap();
            zip.write_all(b"fake gif data").unwrap();

            zip.start_file("Object/image.svg", options).unwrap();
            zip.write_all(b"<svg/>").unwrap();

            zip.start_file("media/diagram.png", options).unwrap();
            zip.write_all(b"fake png data").unwrap();

            zip.finish().unwrap();
        }

        let package = Package::new(&zip_buffer).unwrap();
        let media_files = package.media_files().unwrap();

        assert!(media_files.contains(&"Pictures/photo.jpg".to_string()));
        assert!(media_files.contains(&"Pictures/chart.jpeg".to_string()));
        assert!(media_files.contains(&"media/animation.gif".to_string()));
        assert!(media_files.contains(&"Object/image.svg".to_string()));
        assert!(media_files.contains(&"media/diagram.png".to_string()));
    }
}
