//! Provides a general interface to a physical OPC package (ZIP file).
//!
//! This module handles the low-level reading of OPC packages from ZIP archives,
//! providing efficient access to package contents with minimal memory allocation.
//!
//! Uses the high-performance soapberry-zip library for zero-copy ZIP parsing.

use crate::error::{OpcError, Result};
use crate::limits::{ReadLimits, ReadResource};
use crate::packuri::{PackURI, PartNameConflict};
use soapberry_zip::CompressionMethod;
use soapberry_zip::office::{LazyArchiveReader, LimitResource};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use std::sync::{Arc, Mutex};

/// Aggregate limits for OPC part materialization.
///
/// A successful read is charged per materialization, not per distinct archive
/// member: reading the same part again consumes another `max_parts` slot.
#[derive(Default)]
struct PartBudget {
    reserved_parts: usize,
    reserved_declared: u64,
    materialized_parts: usize,
    materialized_actual: u64,
}

#[derive(Clone, Copy)]
struct PartReservation {
    parts: usize,
    declared: u64,
}

#[derive(Default)]
struct RelationshipBudget {
    reserved_parts: usize,
    reserved_bytes: u64,
    materialized_parts: usize,
    materialized_bytes: u64,
}

/// Physical package reader that provides access to parts in a ZIP-based OPC package.
///
/// Uses `soapberry_zip` for high-performance zero-copy ZIP parsing with lazy decompression.
/// File contents are decompressed on-demand and cached for efficiency. This enables
/// pipelining of decompression with XML parsing for better throughput.
#[allow(
    clippy::module_name_repetitions,
    reason = "name mirrors the OPC 'physical package' concept and is part of the public API"
)]
pub struct PhysPkgReader<'data> {
    /// The underlying ZIP archive reader (lazy decompression with caching)
    archive: LazyArchiveReader<'data>,
    /// Validated policy retained for package-level parsing.
    limits: ReadLimits,
    /// Shared across readers borrowed from one owned physical package.
    part_budget: Arc<Mutex<PartBudget>>,
    /// Public relationship-manifest reads share one aggregate budget.
    relationship_budget: Arc<Mutex<RelationshipBudget>>,
}

/// Owned version of `PhysPkgReader` that owns the data buffer.
///
/// This is used when reading from files or readers where we need to own the data.
pub struct OwnedPhysPkgReader {
    /// The owned data buffer
    data: Vec<u8>,
    /// Validated policy retained when creating borrowed readers.
    limits: ReadLimits,
    /// Shared with each borrowed reader so repeated owned reads cannot reset
    /// materialized-part budgets.
    part_budget: Arc<Mutex<PartBudget>>,
    /// Shared with each borrowed reader so public relationship reads cannot
    /// reset their aggregate count or byte budget.
    relationship_budget: Arc<Mutex<RelationshipBudget>>,
}

impl OwnedPhysPkgReader {
    /// Open an OPC package from a file path.
    ///
    /// # Arguments
    /// * `path` - Path to the OPC package file (.docx, .xlsx, .pptx, etc.)
    ///
    /// # Returns
    /// A new `OwnedPhysPkgReader` instance
    ///
    /// # Errors
    /// Returns an error if the file doesn't exist, isn't a valid ZIP file,
    /// or cannot be opened.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open an OPC package from a file path with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the file does not exist, cannot be read, exceeds the
    /// input-byte limit, or is not a valid ZIP archive.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: ReadLimits) -> Result<Self> {
        let path_ref = path.as_ref();

        if !path_ref.exists() {
            return Err(OpcError::PackageNotFound(path_ref.display().to_string()));
        }

        let metadata = std::fs::metadata(path_ref)?;
        if metadata.is_file() {
            limits.check_input_bytes(metadata.len())?;
        }
        let data = read_limited(std::fs::File::open(path_ref)?, limits)?;
        Self::from_bytes_with_limits(data, limits)
    }

    /// Create a new `OwnedPhysPkgReader` from owned bytes.
    ///
    /// # Errors
    /// Returns an error if `data` is not a valid ZIP archive under the default limits.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(data, ReadLimits::default())
    }

    /// Create a new owned physical reader with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if `data` exceeds the input-byte limit or is not a valid
    /// ZIP archive under `limits`.
    pub fn from_bytes_with_limits(data: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        limits.check_input_bytes(data.len() as u64)?;
        // Validate the ZIP archive can be parsed
        let _ = LazyArchiveReader::new_with_limits(&data, limits.zip_limits())
            .map_err(|error| map_archive_error(&error))?;
        Ok(Self {
            data,
            limits,
            part_budget: Arc::new(Mutex::new(PartBudget::default())),
            relationship_budget: Arc::new(Mutex::new(RelationshipBudget::default())),
        })
    }

    /// Create a new `OwnedPhysPkgReader` from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read
    ///
    /// # Returns
    /// A new `OwnedPhysPkgReader` instance
    ///
    /// # Errors
    /// Returns an error if the stream cannot be read or is not a valid ZIP archive
    /// under the default limits.
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        Self::from_reader_with_limits(&mut reader, ReadLimits::default())
    }

    /// Create a new owned physical reader from a stream with explicit limits.
    ///
    /// # Errors
    /// Returns an error if the stream cannot be read, exceeds `limits`, or is not
    /// a valid ZIP archive.
    pub fn from_reader_with_limits<R: Read>(reader: R, limits: ReadLimits) -> Result<Self> {
        let data = read_limited(reader, limits)?;
        Self::from_bytes_with_limits(data, limits)
    }

    /// Get a borrowed reader for accessing archive contents.
    ///
    /// # Errors
    /// Returns an error if the owned data fails archive validation under the
    /// retained limits.
    #[inline]
    pub fn reader(&self) -> Result<PhysPkgReader<'_>> {
        PhysPkgReader::new_with_limits_and_budget(
            &self.data,
            self.limits,
            Arc::clone(&self.part_budget),
            Arc::clone(&self.relationship_budget),
        )
    }

    /// Get the binary content for a part by its `PackURI`.
    ///
    /// # Errors
    /// Returns an error if the part is missing, unreadable, or exceeds the
    /// configured part limits.
    #[inline]
    pub fn blob_for(&self, pack_uri: &PackURI) -> Result<Vec<u8>> {
        self.reader()?.blob_for(pack_uri)
    }

    /// Get the `[Content_Types].xml` content.
    ///
    /// # Errors
    /// Returns an error if the part is missing, unreadable, or exceeds the
    /// content-types byte limit.
    #[inline]
    pub fn content_types_xml(&self) -> Result<Vec<u8>> {
        self.reader()?.content_types_xml()
    }

    /// Get the relationships XML for a specific source URI.
    ///
    /// # Errors
    /// Returns an error if the relationships part is unreadable or exceeds the
    /// relationship limits.
    #[inline]
    pub fn rels_xml_for(&self, source_uri: &PackURI) -> Result<Option<Vec<u8>>> {
        self.reader()?.rels_xml_for(source_uri)
    }

    /// Get the number of files in the package.
    ///
    /// # Errors
    /// Returns an error if the borrowed reader cannot be created.
    #[inline]
    pub fn len(&self) -> Result<usize> {
        Ok(self.reader()?.len())
    }

    /// Check if the package is empty.
    ///
    /// # Errors
    /// Returns an error if the borrowed reader cannot be created.
    #[inline]
    pub fn is_empty(&self) -> Result<bool> {
        Ok(self.reader()?.is_empty())
    }

    /// List all member names in the package.
    ///
    /// # Errors
    /// Returns an error if the borrowed reader cannot be created.
    #[inline]
    pub fn member_names(&self) -> Result<Vec<String>> {
        self.reader()?.member_names()
    }

    /// Check if a specific member exists in the package.
    ///
    /// # Errors
    /// Returns an error if the borrowed reader cannot be created.
    #[inline]
    pub fn contains(&self, pack_uri: &PackURI) -> Result<bool> {
        Ok(self.reader()?.contains(pack_uri))
    }

    /// Read an archive member by its normalized ZIP name.
    ///
    /// This is physical ZIP access: it observes input and ZIP archive limits,
    /// but deliberately does not charge the OPC materialized-part budget. Use
    /// [`Self::blob_for`] for an OPC part.
    ///
    /// # Errors
    /// Returns an error if the member is missing or unreadable.
    pub fn read_member(&self, name: &str) -> Result<Vec<u8>> {
        self.reader()?.read_member(name)
    }

    /// Consume self and return the underlying data.
    #[inline]
    #[must_use]
    pub fn into_inner(self) -> Vec<u8> {
        self.data
    }

    /// Get a reference to the underlying data.
    #[inline]
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.data
    }
}

impl<'data> PhysPkgReader<'data> {
    /// Create a new `PhysPkgReader` from a byte slice.
    ///
    /// # Arguments
    /// * `data` - The ZIP archive data as a byte slice
    ///
    /// # Returns
    /// A new `PhysPkgReader` instance
    ///
    /// # Errors
    /// Returns an error if `data` is not a valid ZIP archive under the default limits.
    pub fn new(data: &'data [u8]) -> Result<Self> {
        Self::new_with_limits(data, ReadLimits::default())
    }

    /// Create a physical reader from a byte slice with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if `data` exceeds `limits` or is not a valid ZIP archive.
    pub fn new_with_limits(data: &'data [u8], limits: ReadLimits) -> Result<Self> {
        Self::new_with_limits_and_budget(
            data,
            limits,
            Arc::new(Mutex::new(PartBudget::default())),
            Arc::new(Mutex::new(RelationshipBudget::default())),
        )
    }

    fn new_with_limits_and_budget(
        data: &'data [u8],
        limits: ReadLimits,
        part_budget: Arc<Mutex<PartBudget>>,
        relationship_budget: Arc<Mutex<RelationshipBudget>>,
    ) -> Result<Self> {
        limits.check_input_bytes(data.len() as u64)?;
        let archive = LazyArchiveReader::new_with_limits(data, limits.zip_limits())
            .map_err(|error| map_archive_error(&error))?;
        Ok(Self {
            archive,
            limits,
            part_budget,
            relationship_budget,
        })
    }

    /// Get the binary content for a part by its `PackURI`.
    ///
    /// Uses efficient lazy decompression. The returned vector contains
    /// the decompressed content. Every successful materialization consumes one
    /// `max_parts` slot, including repeated requests for the same URI.
    ///
    /// # Arguments
    /// * `pack_uri` - The `PackURI` of the part to read
    ///
    /// # Returns
    /// The binary content of the part
    ///
    /// # Errors
    /// Returns an error if the part is missing, unreadable, or exceeds the
    /// configured part limits.
    pub fn blob_for(&self, pack_uri: &PackURI) -> Result<Vec<u8>> {
        let membername = pack_uri.membername();
        let label = pack_uri.to_string();
        let declared = self.declared_part_bytes(membername, &label)?;
        let reservation = self.reserve_declared_parts(&[declared])?;
        match self.archive.read(membername) {
            Ok(blob) => {
                self.commit_actual_parts(reservation, std::slice::from_ref(&blob))?;
                Ok(blob)
            },
            Err(error) => {
                self.release_declared_parts(reservation);
                Err(map_part_error(&label, &error))
            },
        }
    }

    /// Read an archive member by its normalized ZIP name.
    ///
    /// This low-level physical operation is archive-budgeted only. It does not
    /// consume `max_part_bytes` or `max_total_part_bytes`; callers loading OPC
    /// parts must use [`Self::blob_for`] instead.
    ///
    /// # Errors
    /// Returns an error if the member is missing or unreadable.
    pub fn read_member(&self, name: &str) -> Result<Vec<u8>> {
        self.archive
            .read(name)
            .map_err(|error| map_archive_error(&error))
    }

    /// Get the `[Content_Types].xml` content.
    ///
    /// This is a required part of every OPC package that maps parts to content types.
    ///
    /// # Errors
    /// Returns an error if the part is missing, unreadable, or exceeds the
    /// content-types byte limit.
    pub fn content_types_xml(&self) -> Result<Vec<u8>> {
        let content_types_uri =
            PackURI::new(crate::packuri::CONTENT_TYPES_URI).map_err(OpcError::InvalidPackUri)?;
        let member_name =
            crate::pkgreader::PackageReader::locate_content_types_member(&self.archive)?;
        self.read_bounded_member(
            member_name,
            content_types_uri.as_ref(),
            ReadResource::ContentTypesBytes,
            self.limits.max_content_types_bytes() as u64,
        )
    }

    /// Get the relationships XML for a specific source URI.
    ///
    /// Relationships files are stored in _rels directories and have a .rels extension.
    /// Returns None if the source has no relationships file.
    ///
    /// # Arguments
    /// * `source_uri` - The `PackURI` of the source (part or package)
    ///
    /// # Errors
    /// Returns an error if the relationships part is unreadable or exceeds the
    /// relationship limits.
    pub fn rels_xml_for(&self, source_uri: &PackURI) -> Result<Option<Vec<u8>>> {
        let rels_uri = source_uri.rels_uri().map_err(OpcError::InvalidPackUri)?;

        match self.read_relationship_member(rels_uri.membername(), rels_uri.as_ref()) {
            Ok(blob) => Ok(Some(blob)),
            Err(OpcError::PartNotFound(_)) => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// Get the number of files in the package (excluding directories).
    pub fn len(&self) -> usize {
        self.archive.len()
    }

    /// Check if the package is empty.
    pub fn is_empty(&self) -> bool {
        self.archive.is_empty()
    }

    /// List all member names in the package.
    ///
    /// Returns all file names in the ZIP archive (excluding directories).
    /// Useful for debugging or exploring package contents.
    ///
    /// # Errors
    /// This function currently always succeeds; the `Result` is retained for
    /// API symmetry with the other readers.
    pub fn member_names(&self) -> Result<Vec<String>> {
        Ok(self.archive.file_names().map(String::from).collect())
    }

    /// Check if a specific member exists in the package.
    ///
    /// Uses the pre-built index for O(1) lookup.
    ///
    /// # Arguments
    /// * `pack_uri` - The `PackURI` to check
    pub fn contains(&self, pack_uri: &PackURI) -> bool {
        let membername = pack_uri.membername();
        self.archive.contains(membername)
    }

    /// Read multiple blobs serially.
    ///
    /// This compatibility API retains its established name but no longer uses
    /// an implicit global scheduler. Explicit bounded eager package opens use
    /// [`crate::OpenSession`] instead.
    ///
    /// # Arguments
    /// * `uris` - Slice of `PackURIs` to read
    ///
    /// # Returns
    /// A `HashMap` mapping member names to their decompressed contents.
    ///
    /// The operation is all-or-nothing: an error for any requested part drops
    /// every successful buffer and returns that error. Each URI occurrence
    /// consumes one `max_parts` slot, including repeated URIs.
    ///
    /// # Errors
    /// Returns an error if any requested part is missing, unreadable, fails
    /// allocation, or exceeds the configured limits.
    pub fn blobs_parallel(&self, uris: &[PackURI]) -> Result<HashMap<String, Vec<u8>>> {
        self.limits.check(
            ReadResource::Parts,
            uris.len() as u64,
            self.limits.max_parts() as u64,
        )?;
        let mut names = Vec::new();
        names
            .try_reserve(uris.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC parallel part names",
                source,
            })?;
        let mut declared = Vec::new();
        declared
            .try_reserve(uris.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC parallel part metadata",
                source,
            })?;
        for uri in uris {
            let name = uri.membername();
            declared.push(self.declared_part_bytes(name, uri.as_str())?);
            names.push(name);
        }
        let reservation = self.reserve_declared_parts(&declared)?;
        let results = names
            .iter()
            .map(|name| (*name, self.archive.read(name)))
            .collect::<Vec<_>>();
        let mut materialized = Vec::new();
        if let Err(source) = materialized.try_reserve(results.len()) {
            self.release_declared_parts(reservation);
            return Err(OpcError::Allocation {
                resource: "OPC parallel part results",
                source,
            });
        }
        for (name, result) in results {
            match result {
                Ok(blob) => materialized.push(blob),
                Err(error) => {
                    self.release_declared_parts(reservation);
                    return Err(map_part_error(name, &error));
                },
            }
        }
        if materialized.len() != names.len() {
            self.release_declared_parts(reservation);
            return Err(OpcError::ZipError(
                "ZIP bulk reader returned an incomplete result set".to_owned(),
            ));
        }
        let mut keys = Vec::new();
        if let Err(source) = keys.try_reserve(names.len()) {
            self.release_declared_parts(reservation);
            return Err(OpcError::Allocation {
                resource: "OPC parallel part keys",
                source,
            });
        }
        for name in &names {
            keys.push((*name).to_owned());
        }
        let mut blobs = HashMap::new();
        if let Err(source) = blobs.try_reserve(materialized.len()) {
            self.release_declared_parts(reservation);
            return Err(OpcError::Allocation {
                resource: "OPC parallel part map",
                source,
            });
        }
        self.commit_actual_parts(reservation, &materialized)?;
        for (key, blob) in keys.into_iter().zip(materialized) {
            blobs.insert(key, blob);
        }
        Ok(blobs)
    }

    /// Get a reference to the underlying lazy archive reader.
    ///
    /// The lazy reader decompresses files on-demand and caches results.
    /// This enables pipelining of decompression with parsing.
    #[inline]
    pub(crate) fn archive(&self) -> &LazyArchiveReader<'data> {
        &self.archive
    }

    /// Return the validated policy retained by this reader.
    #[inline]
    #[allow(
        dead_code,
        reason = "pkgreader consumes the retained profile as parser limits are migrated"
    )]
    pub(crate) const fn limits(&self) -> ReadLimits {
        self.limits
    }

    fn declared_part_bytes(&self, name: &str, label: &str) -> Result<u64> {
        self.archive
            .metadata(name)
            .map(|metadata| metadata.uncompressed_size())
            .map_err(|error| map_part_error(label, &error))
    }

    fn read_bounded_member(
        &self,
        name: &str,
        label: &str,
        resource: ReadResource,
        maximum: u64,
    ) -> Result<Vec<u8>> {
        let declared = self
            .archive
            .metadata(name)
            .map(|metadata| metadata.uncompressed_size())
            .map_err(|error| map_part_error(label, &error))?;
        self.limits.check(resource, declared, maximum)?;
        let blob = self
            .archive
            .read(name)
            .map_err(|error| map_part_error(label, &error))?;
        self.limits.check(resource, blob.len() as u64, maximum)?;
        Ok(blob)
    }

    fn read_relationship_member(&self, name: &str, label: &str) -> Result<Vec<u8>> {
        let declared = self
            .archive
            .metadata(name)
            .map(|metadata| metadata.uncompressed_size())
            .map_err(|error| map_part_error(label, &error))?;
        self.limits.check(
            ReadResource::RelationshipXmlBytes,
            declared,
            self.limits.max_relationship_xml_bytes() as u64,
        )?;
        self.reserve_declared_relationship(declared)?;
        match self.archive.read(name) {
            Ok(blob) => {
                self.commit_relationship(declared, blob.len() as u64)?;
                Ok(blob)
            },
            Err(error) => {
                self.release_declared_relationship(declared);
                Err(map_part_error(label, &error))
            },
        }
    }

    fn reserve_declared_relationship(&self, declared: u64) -> Result<()> {
        let mut budget = self
            .relationship_budget
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let parts = budget
            .materialized_parts
            .checked_add(budget.reserved_parts)
            .and_then(|count| count.checked_add(1))
            .ok_or(OpcError::ReadLimit {
                resource: ReadResource::RelationshipParts,
                actual: u64::MAX,
                maximum: self.limits.max_relationship_parts() as u64,
            })?;
        self.limits.check(
            ReadResource::RelationshipParts,
            parts as u64,
            self.limits.max_relationship_parts() as u64,
        )?;
        let bytes = budget
            .materialized_bytes
            .checked_add(budget.reserved_bytes)
            .and_then(|bytes| bytes.checked_add(declared))
            .ok_or(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlBytes,
                actual: u64::MAX,
                maximum: self.limits.max_total_relationship_xml_bytes() as u64,
            })?;
        self.limits.check(
            ReadResource::TotalRelationshipXmlBytes,
            bytes,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )?;
        budget.reserved_parts =
            budget
                .reserved_parts
                .checked_add(1)
                .ok_or(OpcError::ReadLimit {
                    resource: ReadResource::RelationshipParts,
                    actual: u64::MAX,
                    maximum: self.limits.max_relationship_parts() as u64,
                })?;
        budget.reserved_bytes =
            budget
                .reserved_bytes
                .checked_add(declared)
                .ok_or(OpcError::ReadLimit {
                    resource: ReadResource::TotalRelationshipXmlBytes,
                    actual: u64::MAX,
                    maximum: self.limits.max_total_relationship_xml_bytes() as u64,
                })?;
        Ok(())
    }

    fn release_declared_relationship(&self, declared: u64) {
        let mut budget = self
            .relationship_budget
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        budget.reserved_parts = budget.reserved_parts.saturating_sub(1);
        budget.reserved_bytes = budget.reserved_bytes.saturating_sub(declared);
    }

    fn commit_relationship(&self, declared: u64, actual: u64) -> Result<()> {
        let result = (|| {
            self.limits.check(
                ReadResource::RelationshipXmlBytes,
                actual,
                self.limits.max_relationship_xml_bytes() as u64,
            )?;
            let mut budget = self
                .relationship_budget
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            let reserved_parts = budget
                .reserved_parts
                .checked_sub(1)
                .ok_or(OpcError::ZipError(
                    "OPC relationship budget reservation underflow".to_owned(),
                ))?;
            let reserved_bytes =
                budget
                    .reserved_bytes
                    .checked_sub(declared)
                    .ok_or(OpcError::ZipError(
                        "OPC relationship byte reservation underflow".to_owned(),
                    ))?;
            let materialized_parts =
                budget
                    .materialized_parts
                    .checked_add(1)
                    .ok_or(OpcError::ReadLimit {
                        resource: ReadResource::RelationshipParts,
                        actual: u64::MAX,
                        maximum: self.limits.max_relationship_parts() as u64,
                    })?;
            let materialized_bytes =
                budget
                    .materialized_bytes
                    .checked_add(actual)
                    .ok_or(OpcError::ReadLimit {
                        resource: ReadResource::TotalRelationshipXmlBytes,
                        actual: u64::MAX,
                        maximum: self.limits.max_total_relationship_xml_bytes() as u64,
                    })?;
            let parts =
                materialized_parts
                    .checked_add(reserved_parts)
                    .ok_or(OpcError::ReadLimit {
                        resource: ReadResource::RelationshipParts,
                        actual: u64::MAX,
                        maximum: self.limits.max_relationship_parts() as u64,
                    })?;
            self.limits.check(
                ReadResource::RelationshipParts,
                parts as u64,
                self.limits.max_relationship_parts() as u64,
            )?;
            let bytes =
                materialized_bytes
                    .checked_add(reserved_bytes)
                    .ok_or(OpcError::ReadLimit {
                        resource: ReadResource::TotalRelationshipXmlBytes,
                        actual: u64::MAX,
                        maximum: self.limits.max_total_relationship_xml_bytes() as u64,
                    })?;
            self.limits.check(
                ReadResource::TotalRelationshipXmlBytes,
                bytes,
                self.limits.max_total_relationship_xml_bytes() as u64,
            )?;
            budget.reserved_parts = reserved_parts;
            budget.reserved_bytes = reserved_bytes;
            budget.materialized_parts = materialized_parts;
            budget.materialized_bytes = materialized_bytes;
            Ok(())
        })();
        if result.is_err() {
            self.release_declared_relationship(declared);
        }
        result
    }

    fn reserve_declared_parts(&self, parts: &[u64]) -> Result<PartReservation> {
        let requested_parts = parts.len();
        let mut requested = 0u64;
        for &bytes in parts {
            self.limits
                .check(ReadResource::PartBytes, bytes, self.limits.max_part_bytes())?;
            requested = requested.checked_add(bytes).ok_or(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                actual: u64::MAX,
                maximum: self.limits.max_total_part_bytes(),
            })?;
        }
        let mut budget = self
            .part_budget
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let count = budget
            .materialized_parts
            .checked_add(budget.reserved_parts)
            .and_then(|count| count.checked_add(requested_parts))
            .ok_or(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                actual: u64::MAX,
                maximum: self.limits.max_parts() as u64,
            })?;
        self.limits.check(
            ReadResource::Parts,
            count as u64,
            self.limits.max_parts() as u64,
        )?;
        let in_flight = budget
            .materialized_actual
            .checked_add(budget.reserved_declared)
            .and_then(|bytes| bytes.checked_add(requested))
            .ok_or(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                actual: u64::MAX,
                maximum: self.limits.max_total_part_bytes(),
            })?;
        self.limits.check(
            ReadResource::TotalPartBytes,
            in_flight,
            self.limits.max_total_part_bytes(),
        )?;
        let reserved_parts =
            budget
                .reserved_parts
                .checked_add(requested_parts)
                .ok_or(OpcError::ReadLimit {
                    resource: ReadResource::Parts,
                    actual: u64::MAX,
                    maximum: self.limits.max_parts() as u64,
                })?;
        let reserved_declared =
            budget
                .reserved_declared
                .checked_add(requested)
                .ok_or(OpcError::ReadLimit {
                    resource: ReadResource::TotalPartBytes,
                    actual: u64::MAX,
                    maximum: self.limits.max_total_part_bytes(),
                })?;
        budget.reserved_parts = reserved_parts;
        budget.reserved_declared = reserved_declared;
        Ok(PartReservation {
            parts: requested_parts,
            declared: requested,
        })
    }

    fn release_declared_parts(&self, reservation: PartReservation) {
        let mut budget = self
            .part_budget
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        budget.reserved_parts = budget.reserved_parts.saturating_sub(reservation.parts);
        budget.reserved_declared = budget
            .reserved_declared
            .saturating_sub(reservation.declared);
    }

    fn commit_actual_parts(&self, reservation: PartReservation, blobs: &[Vec<u8>]) -> Result<()> {
        let result =
            (|| {
                let actual_parts = blobs.len();
                let mut actual = 0u64;
                for blob in blobs {
                    let bytes = blob.len() as u64;
                    self.limits.check(
                        ReadResource::PartBytes,
                        bytes,
                        self.limits.max_part_bytes(),
                    )?;
                    actual = actual.checked_add(bytes).ok_or(OpcError::ReadLimit {
                        resource: ReadResource::TotalPartBytes,
                        actual: u64::MAX,
                        maximum: self.limits.max_total_part_bytes(),
                    })?;
                }

                let mut budget = self
                    .part_budget
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner);
                let remaining_parts = budget.reserved_parts.checked_sub(reservation.parts).ok_or(
                    OpcError::ZipError("OPC part count reservation underflow".to_owned()),
                )?;
                let remaining_bytes = budget
                    .reserved_declared
                    .checked_sub(reservation.declared)
                    .ok_or(OpcError::ZipError(
                        "OPC part byte reservation underflow".to_owned(),
                    ))?;
                let next_parts = budget.materialized_parts.checked_add(actual_parts).ok_or(
                    OpcError::ReadLimit {
                        resource: ReadResource::Parts,
                        actual: u64::MAX,
                        maximum: self.limits.max_parts() as u64,
                    },
                )?;
                let next_actual =
                    budget
                        .materialized_actual
                        .checked_add(actual)
                        .ok_or(OpcError::ReadLimit {
                            resource: ReadResource::TotalPartBytes,
                            actual: u64::MAX,
                            maximum: self.limits.max_total_part_bytes(),
                        })?;
                let parts = next_parts
                    .checked_add(remaining_parts)
                    .ok_or(OpcError::ReadLimit {
                        resource: ReadResource::Parts,
                        actual: u64::MAX,
                        maximum: self.limits.max_parts() as u64,
                    })?;
                self.limits.check(
                    ReadResource::Parts,
                    parts as u64,
                    self.limits.max_parts() as u64,
                )?;
                let bytes =
                    next_actual
                        .checked_add(remaining_bytes)
                        .ok_or(OpcError::ReadLimit {
                            resource: ReadResource::TotalPartBytes,
                            actual: u64::MAX,
                            maximum: self.limits.max_total_part_bytes(),
                        })?;
                self.limits.check(
                    ReadResource::TotalPartBytes,
                    bytes,
                    self.limits.max_total_part_bytes(),
                )?;
                budget.reserved_parts = remaining_parts;
                budget.reserved_declared = remaining_bytes;
                budget.materialized_parts = next_parts;
                budget.materialized_actual = next_actual;
                Ok(())
            })();
        if result.is_err() {
            self.release_declared_parts(reservation);
        }
        result
    }
}

#[derive(Default)]
struct PartNameSet {
    /// Full part names folded with the OPC ASCII-case-equivalence rule.
    names: HashMap<String, PackURI>,
    /// For each folded ancestor path, one known descendant full name.
    ///
    /// OPC topology validation rejects any ancestor/descendant pair. Keeping
    /// this index makes validation proportional to path depth rather than the
    /// number of already-emitted members.
    descendants: HashMap<String, String>,
}

struct PreparedPartName {
    folded: String,
    /// Each item is `(folded ancestor, folded full candidate name)`.
    ancestors: Vec<(String, String)>,
}

impl PartNameSet {
    fn prepare(candidate: &PackURI) -> Result<PreparedPartName> {
        let folded = fold_part_name(candidate.as_str())?;
        let mut ancestors = Vec::new();
        let slash_count = folded.bytes().filter(|byte| *byte == b'/').count();
        ancestors
            .try_reserve(slash_count.saturating_sub(1))
            .map_err(|source| OpcError::Allocation {
                resource: "OPC physical package name ancestors",
                source,
            })?;
        for (index, byte) in folded.bytes().enumerate() {
            if byte != b'/' || index == 0 {
                continue;
            }
            let ancestor = try_owned_string(&folded[..index], "OPC physical package ancestor")?;
            let descendant = try_owned_string(&folded, "OPC physical package descendant")?;
            ancestors.push((ancestor, descendant));
        }
        Ok(PreparedPartName { folded, ancestors })
    }

    fn validate(&self, candidate: &PackURI, prepared: &PreparedPartName) -> Result<()> {
        if let Some(existing) = self.names.get(&prepared.folded) {
            let conflict = if existing.as_str() == candidate.as_str() {
                PartNameConflict::Duplicate
            } else {
                PartNameConflict::Equivalent
            };
            return Err(part_name_conflict_error(existing, candidate, conflict));
        }
        for (ancestor, _) in &prepared.ancestors {
            if let Some(existing) = self.names.get(ancestor) {
                return Err(part_name_conflict_error(
                    existing,
                    candidate,
                    PartNameConflict::Derived,
                ));
            }
        }
        if let Some(descendant) = self.descendants.get(&prepared.folded) {
            if let Some(existing) = self.names.get(descendant) {
                return Err(part_name_conflict_error(
                    existing,
                    candidate,
                    PartNameConflict::Derived,
                ));
            }
        }
        Ok(())
    }

    fn reserve(&mut self, prepared: &PreparedPartName) -> Result<()> {
        self.names
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC physical package names",
                source,
            })?;
        self.descendants
            .try_reserve(prepared.ancestors.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC physical package descendant index",
                source,
            })?;
        Ok(())
    }

    fn insert(&mut self, partname: PackURI, prepared: PreparedPartName) {
        let PreparedPartName { folded, ancestors } = prepared;
        self.names.insert(folded, partname);
        for (ancestor, descendant) in ancestors {
            self.descendants.entry(ancestor).or_insert(descendant);
        }
    }
}

fn try_owned_string(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| OpcError::Allocation { resource, source })?;
    owned.push_str(value);
    Ok(owned)
}

fn fold_part_name(value: &str) -> Result<String> {
    let mut folded = String::new();
    folded
        .try_reserve_exact(value.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC physical package folded name",
            source,
        })?;
    for byte in value.bytes() {
        folded.push(char::from(byte.to_ascii_lowercase()));
    }
    Ok(folded)
}

/// Physical package writer for creating OPC packages.
///
/// Handles the low-level writing of parts to a ZIP archive with optimal compression.
/// Uses soapberry-zip's high-performance writer.
#[allow(
    clippy::module_name_repetitions,
    reason = "name mirrors the OPC 'physical package' concept and is part of the public API"
)]
pub struct PhysPkgWriter<W: Write = Cursor<Vec<u8>>> {
    /// The underlying ZIP archive writer
    archive: soapberry_zip::office::StreamingArchiveWriter<W>,
    /// Validated OPC member names already committed to this archive.
    ///
    /// Keeping this state at the OPC boundary prevents a streaming caller from
    /// publishing duplicate, ASCII-equivalent, or derived part names while the
    /// underlying ZIP transport remains format-neutral.
    part_names: PartNameSet,
}

/// An owned, sequential writer for one OPC part.
///
/// The physical package writer is moved into this value while the part is
/// being emitted. The part accepts uncompressed bytes through [`Write`] and
/// returns the package writer from [`Self::finish`], so callers never need to
/// hold a borrow into a ZIP archive or see a ZIP entry type. Dropping an
/// unfinished part abandons the consuming package writer and leaves any
/// already-published sequential sink bytes incomplete.
pub struct PartWriter<W: Write> {
    entry: Option<soapberry_zip::office::StreamingArchiveEntry<W>>,
    partname: PackURI,
    part_names: PartNameSet,
    prepared_name: PreparedPartName,
}

impl<W: Write> Write for PartWriter<W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.entry
            .as_mut()
            .ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "OPC Part writer has already been finished",
                )
            })?
            .write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.entry
            .as_mut()
            .ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "OPC Part writer has already been finished",
                )
            })?
            .flush()
    }
}

impl<W: Write> PartWriter<W> {
    /// Total package bytes accepted by the output sink so far.
    ///
    /// This includes local headers and all earlier Parts, not only the active
    /// Part payload. It is useful for diagnostics before [`Self::finish`] is
    /// called. A sink failure after publication begins is reported as
    /// [`OpcError::IncompleteOutput`] from `finish`.
    #[must_use]
    pub fn output_bytes(&self) -> u64 {
        self.entry
            .as_ref()
            .map(|entry| entry.progress().output_bytes())
            .unwrap_or(0)
    }

    /// Number of uncompressed bytes accepted by this part.
    #[must_use]
    pub fn uncompressed_bytes(&self) -> u64 {
        self.entry
            .as_ref()
            .map(soapberry_zip::office::StreamingArchiveEntry::uncompressed_bytes)
            .unwrap_or(0)
    }

    /// Finish this part and recover the physical package writer.
    ///
    /// The part name becomes committed only after the ZIP entry is finalized.
    /// If finalization or the sink fails after bytes were accepted, the error
    /// carries the exact accepted byte count in [`OpcError::IncompleteOutput`].
    pub fn finish(mut self) -> Result<PhysPkgWriter<W>> {
        let entry = self.entry.take().ok_or_else(|| {
            OpcError::ZipError("OPC Part writer has already been finished".to_string())
        })?;
        let (archive, _progress) = entry
            .finish_with_progress()
            .map_err(map_streaming_failure)?;
        let partname = self.partname;
        self.part_names.insert(partname, self.prepared_name);
        Ok(PhysPkgWriter {
            archive,
            part_names: self.part_names,
        })
    }
}

impl PhysPkgWriter<Cursor<Vec<u8>>> {
    /// Create a new package writer that writes to memory.
    #[must_use]
    pub fn new() -> Self {
        Self {
            archive: soapberry_zip::office::StreamingArchiveWriter::new(),
            part_names: PartNameSet::default(),
        }
    }

    /// Finish writing and return the package bytes.
    ///
    /// Consumes the writer and returns the complete ZIP archive.
    ///
    /// # Errors
    /// Returns an error if the ZIP archive cannot be finalized.
    pub fn finish(self) -> Result<Vec<u8>> {
        let (writer, _progress) = self
            .archive
            .finish_with_progress()
            .map_err(map_streaming_failure)?;
        Ok(writer.into_inner())
    }
}

impl<W: Write> PhysPkgWriter<W> {
    /// Create a physical package writer over a sequential sink.
    pub fn with_writer(writer: W) -> Self {
        Self {
            archive: soapberry_zip::office::StreamingArchiveWriter::with_writer(writer),
            part_names: PartNameSet::default(),
        }
    }

    /// Start a Deflate-compressed OPC Part without buffering its payload.
    ///
    /// The package writer is moved into the returned part writer while the
    /// payload is emitted. Call [`PartWriter::finish`] to recover it. A
    /// preflight failure consumes this writer without publishing a new local
    /// header; dropping a successfully started Part leaves the sequential
    /// output incomplete.
    /// `partname` must be a checked, non-root [`PackURI`]. Duplicate,
    /// ASCII-equivalent, and derived names are rejected before the ZIP local
    /// header is written.
    pub fn start_part(self, partname: &PackURI) -> Result<PartWriter<W>> {
        self.start_part_inner(partname, CompressionMethod::Deflate)
    }

    /// Start a stored OPC Part without buffering its payload.
    ///
    /// This is intended for package members whose format contract requires
    /// Store compression, such as an ODF `mimetype` member. The same OPC name
    /// validation and bounded ZIP transport policy as [`Self::start_part`]
    /// applies.
    pub fn start_stored_part(self, partname: &PackURI) -> Result<PartWriter<W>> {
        self.start_part_inner(partname, CompressionMethod::Store)
    }

    /// Number of bytes accepted by the sequential output sink so far.
    #[must_use]
    pub fn output_bytes(&self) -> u64 {
        self.archive.output_bytes()
    }

    fn start_part_inner(
        mut self,
        partname: &PackURI,
        compression_method: CompressionMethod,
    ) -> Result<PartWriter<W>> {
        let prepared_name = PartNameSet::prepare(partname)?;
        validate_part_name(&self.part_names, partname, &prepared_name)?;
        self.part_names.reserve(&prepared_name)?;
        let owned_partname = clone_pack_uri(partname)?;
        let entry = self
            .archive
            .start_entry(partname.membername(), compression_method)
            .map_err(map_streaming_failure)?;
        Ok(PartWriter {
            entry: Some(entry),
            partname: owned_partname,
            part_names: self.part_names,
            prepared_name,
        })
    }

    /// Write a part to the package with Deflate compression.
    ///
    /// # Arguments
    /// * `pack_uri` - The `PackURI` for the part
    /// * `blob` - The binary content to write
    ///
    /// # Errors
    /// Returns an error if the part cannot be written to the archive.
    pub fn write(&mut self, pack_uri: &PackURI, blob: &[u8]) -> Result<()> {
        let prepared_name = PartNameSet::prepare(pack_uri)?;
        validate_part_name(&self.part_names, pack_uri, &prepared_name)?;
        self.part_names.reserve(&prepared_name)?;
        let owned_partname = clone_pack_uri(pack_uri)?;
        self.archive
            .write_deflated(pack_uri.membername(), blob)
            .map_err(|error| map_archive_error(&error))?;
        self.part_names.insert(owned_partname, prepared_name);
        Ok(())
    }

    /// Write a part to the package without compression (stored).
    ///
    /// # Arguments
    /// * `pack_uri` - The `PackURI` for the part
    /// * `blob` - The binary content to write
    ///
    /// # Errors
    /// Returns an error if the part cannot be written to the archive.
    pub fn write_stored(&mut self, pack_uri: &PackURI, blob: &[u8]) -> Result<()> {
        let prepared_name = PartNameSet::prepare(pack_uri)?;
        validate_part_name(&self.part_names, pack_uri, &prepared_name)?;
        self.part_names.reserve(&prepared_name)?;
        let owned_partname = clone_pack_uri(pack_uri)?;
        self.archive
            .write_stored(pack_uri.membername(), blob)
            .map_err(|error| map_archive_error(&error))?;
        self.part_names.insert(owned_partname, prepared_name);
        Ok(())
    }

    /// Finalize the archive and recover the sequential sink.
    ///
    /// # Errors
    /// Returns an error if the ZIP archive cannot be finalized.
    pub fn finish_into_inner(self) -> Result<W> {
        self.archive
            .finish_with_progress()
            .map(|(writer, _progress)| writer)
            .map_err(map_streaming_failure)
    }
}

impl Default for PhysPkgWriter<Cursor<Vec<u8>>> {
    fn default() -> Self {
        Self::new()
    }
}

pub(crate) fn read_limited<R: Read>(mut reader: R, limits: ReadLimits) -> Result<Vec<u8>> {
    let Ok(maximum) = usize::try_from(limits.max_input_bytes()) else {
        return Err(OpcError::InvalidReadLimit {
            resource: ReadResource::InputBytes,
            value: limits.max_input_bytes(),
        });
    };
    let mut data = Vec::new();
    data.try_reserve_exact(maximum.min(8 * 1024))
        .map_err(|source| OpcError::Allocation {
            resource: "OPC package input",
            source,
        })?;

    let mut buffer = [0u8; 8 * 1024];
    loop {
        let remaining = maximum.saturating_sub(data.len());
        if remaining == 0 {
            let mut extra = [0u8; 1];
            let read = loop {
                match reader.read(&mut extra) {
                    Ok(read) => break read,
                    Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                    Err(error) => return Err(error.into()),
                }
            };
            if read != 0 {
                return Err(OpcError::ReadLimit {
                    resource: ReadResource::InputBytes,
                    actual: maximum as u64 + 1,
                    maximum: maximum as u64,
                });
            }
            return Ok(data);
        }

        let chunk = remaining.min(buffer.len());
        let read = loop {
            match reader.read(&mut buffer[..chunk]) {
                Ok(read) => break read,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) => return Err(error.into()),
            }
        };
        if read > chunk {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "reader returned more bytes than the provided buffer",
            )
            .into());
        }
        if read == 0 {
            return Ok(data);
        }
        data.try_reserve_exact(read)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC package input",
                source,
            })?;
        data.extend_from_slice(&buffer[..read]);
    }
}

fn map_streaming_failure(failure: soapberry_zip::office::StreamingArchiveFailure) -> OpcError {
    let written = failure.progress().output_bytes();
    let source = OpcError::from(failure.into_error());
    if written == 0 {
        source
    } else {
        OpcError::IncompleteOutput {
            written,
            source: Box::new(source),
        }
    }
}

fn validate_part_name(
    part_names: &PartNameSet,
    candidate: &PackURI,
    prepared: &PreparedPartName,
) -> Result<()> {
    if candidate.as_str() == "/" {
        return Err(OpcError::InvalidPackUri(
            "an OPC Part cannot use the package root URI".to_string(),
        ));
    }
    part_names.validate(candidate, prepared)
}

fn clone_pack_uri(candidate: &PackURI) -> Result<PackURI> {
    let owned = try_owned_string(candidate.as_str(), "OPC physical package part name")?;
    PackURI::new(owned).map_err(OpcError::InvalidPackUri)
}

fn part_name_conflict_error(
    existing: &PackURI,
    candidate: &PackURI,
    conflict: PartNameConflict,
) -> OpcError {
    match conflict {
        PartNameConflict::Duplicate => OpcError::DuplicatePartName(candidate.to_string()),
        PartNameConflict::Equivalent => OpcError::EquivalentPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
        PartNameConflict::Derived => OpcError::DerivedPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
    }
}

fn map_archive_error(error: &soapberry_zip::Error) -> OpcError {
    if let soapberry_zip::ErrorKind::LimitExceeded {
        resource,
        actual,
        maximum,
    } = error.kind()
    {
        let opc_resource = match resource {
            LimitResource::FileCount => ReadResource::ArchiveMembers,
            LimitResource::MemberNameBytes => ReadResource::ArchiveMemberNameBytes,
            LimitResource::MetadataBytes => ReadResource::ArchiveMetadataBytes,
            LimitResource::CompressedSize => ReadResource::ArchiveCompressedBytes,
            LimitResource::EntrySize => ReadResource::ArchiveEntryBytes,
            LimitResource::TotalSize => ReadResource::ArchiveTotalBytes,
        };
        return OpcError::ReadLimit {
            resource: opc_resource,
            actual: *actual,
            maximum: *maximum,
        };
    }
    OpcError::ZipError(error.to_string())
}

fn map_part_error(label: &str, error: &soapberry_zip::Error) -> OpcError {
    if matches!(error.kind(), soapberry_zip::ErrorKind::FileNotFound(_)) {
        OpcError::PartNotFound(label.to_owned())
    } else {
        map_archive_error(error)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    fn stored_archive(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = PhysPkgWriter::new();
        for (name, bytes) in entries {
            let uri = PackURI::new(format!("/{name}")).unwrap();
            writer.write_stored(&uri, bytes).unwrap();
        }
        writer.finish().unwrap()
    }

    #[test]
    fn target_limit_cannot_exceed_the_universal_xml_attribute_limit() {
        assert!(matches!(
            ReadLimits::builder()
                .max_xml_attribute_bytes(3)
                .unwrap()
                .max_relationship_target_bytes(4)
                .unwrap()
                .build(),
            Err(OpcError::InvalidReadLimit {
                resource: ReadResource::RelationshipTargetBytes,
                value: 4,
            })
        ));
    }

    #[test]
    fn dedicated_xml_reads_do_not_consume_generic_part_budgets() {
        let bytes = stored_archive(&[("[Content_Types].xml", b"types"), ("_rels/.rels", b"rels")]);
        let root = PackURI::new("/").unwrap();
        let independent = ReadLimits::builder()
            .max_part_bytes(1)
            .unwrap()
            .max_total_part_bytes(1)
            .unwrap()
            .max_content_types_bytes(5)
            .unwrap()
            .max_relationship_xml_bytes(4)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, independent).unwrap();
        assert_eq!(reader.content_types_xml().unwrap(), b"types");
        assert_eq!(reader.rels_xml_for(&root).unwrap(), Some(b"rels".to_vec()));

        let content_types_over = ReadLimits::builder()
            .max_part_bytes(1)
            .unwrap()
            .max_total_part_bytes(1)
            .unwrap()
            .max_content_types_bytes(4)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            PhysPkgReader::new_with_limits(&bytes, content_types_over)
                .unwrap()
                .content_types_xml(),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ContentTypesBytes,
                actual: 5,
                maximum: 4,
            })
        ));

        let relationships_over = ReadLimits::builder()
            .max_part_bytes(1)
            .unwrap()
            .max_total_part_bytes(1)
            .unwrap()
            .max_relationship_xml_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            PhysPkgReader::new_with_limits(&bytes, relationships_over)
                .unwrap()
                .rels_xml_for(&root),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                actual: 4,
                maximum: 3,
            })
        ));
    }

    #[test]
    fn relationship_reads_share_count_and_byte_budgets_without_charging_failures() {
        let bytes = stored_archive(&[("_rels/.rels", b"rels")]);
        let root = PackURI::new("/").unwrap();
        let exact = ReadLimits::builder()
            .max_relationship_parts(1)
            .unwrap()
            .max_relationship_xml_bytes(4)
            .unwrap()
            .max_total_relationship_xml_bytes(4)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, exact).unwrap();
        assert_eq!(reader.rels_xml_for(&root).unwrap(), Some(b"rels".to_vec()));
        assert!(matches!(
            reader.rels_xml_for(&root),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipParts,
                actual: 2,
                maximum: 1,
            })
        ));

        let aggregate = ReadLimits::builder()
            .max_relationship_parts(2)
            .unwrap()
            .max_relationship_xml_bytes(4)
            .unwrap()
            .max_total_relationship_xml_bytes(4)
            .unwrap()
            .build()
            .unwrap();
        let aggregate_reader = PhysPkgReader::new_with_limits(&bytes, aggregate).unwrap();
        assert_eq!(
            aggregate_reader.rels_xml_for(&root).unwrap(),
            Some(b"rels".to_vec())
        );
        assert!(matches!(
            aggregate_reader.rels_xml_for(&root),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlBytes,
                actual: 8,
                maximum: 4,
            })
        ));

        let byte_over = ReadLimits::builder()
            .max_relationship_parts(2)
            .unwrap()
            .max_relationship_xml_bytes(4)
            .unwrap()
            .max_total_relationship_xml_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            PhysPkgReader::new_with_limits(&bytes, byte_over)
                .unwrap()
                .rels_xml_for(&root),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlBytes,
                actual: 4,
                maximum: 3,
            })
        ));

        let owned = OwnedPhysPkgReader::from_bytes_with_limits(bytes, exact).unwrap();
        assert_eq!(
            owned.reader().unwrap().rels_xml_for(&root).unwrap(),
            Some(b"rels".to_vec())
        );
        assert!(matches!(
            owned.reader().unwrap().rels_xml_for(&root),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipParts,
                actual: 2,
                maximum: 1,
            })
        ));

        let payload = b"\xa5\x5a\xc3";
        let mut corrupt = stored_archive(&[("_rels/.rels", payload)]);
        let offset = corrupt
            .windows(payload.len())
            .position(|window| window == payload)
            .unwrap();
        corrupt[offset] ^= 0xff;
        let rollback = ReadLimits::builder()
            .max_relationship_parts(1)
            .unwrap()
            .max_relationship_xml_bytes(3)
            .unwrap()
            .max_total_relationship_xml_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        let rollback_reader = PhysPkgReader::new_with_limits(&corrupt, rollback).unwrap();
        assert!(matches!(
            rollback_reader.rels_xml_for(&root),
            Err(OpcError::ZipError(_))
        ));
        assert!(matches!(
            rollback_reader.rels_xml_for(&root),
            Err(OpcError::ZipError(_))
        ));
    }

    #[test]
    fn test_round_trip() {
        // Create a ZIP archive with soapberry-zip
        let mut writer = PhysPkgWriter::new();
        let pack_uri = PackURI::new("/test.txt").unwrap();
        writer.write(&pack_uri, b"Hello, World!").unwrap();
        let zip_data = writer.finish().unwrap();

        // Read the ZIP archive
        let reader = PhysPkgReader::new(&zip_data).unwrap();
        let content = reader.blob_for(&pack_uri).unwrap();
        assert_eq!(content, b"Hello, World!");
    }

    #[test]
    fn test_multiple_parts() {
        let mut writer = PhysPkgWriter::new();

        let content_types = PackURI::new("/[Content_Types].xml").unwrap();
        let rels = PackURI::new("/_rels/.rels").unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();

        writer.write(&content_types, b"<Types/>").unwrap();
        writer.write(&rels, b"<Relationships/>").unwrap();
        writer.write(&document, b"<document/>").unwrap();

        let zip_data = writer.finish().unwrap();
        let reader = PhysPkgReader::new(&zip_data).unwrap();

        assert!(reader.contains(&content_types));
        assert!(reader.contains(&rels));
        assert!(reader.contains(&document));
        assert_eq!(reader.blob_for(&document).unwrap(), b"<document/>");
    }

    #[test]
    fn owned_part_writer_streams_and_recovers_the_physical_writer() {
        let document = PackURI::new("/word/document.xml").unwrap();
        let styles = PackURI::new("/word/styles.xml").unwrap();
        let mut writer = PhysPkgWriter::new();

        let mut document_writer = writer.start_part(&document).unwrap();
        document_writer
            .write_all(b"<document>streamed</document>")
            .unwrap();
        assert_eq!(document_writer.uncompressed_bytes(), 29);
        writer = document_writer.finish().unwrap();

        let mut styles_writer = writer.start_stored_part(&styles).unwrap();
        styles_writer.write_all(b"<styles/>").unwrap();
        writer = styles_writer.finish().unwrap();

        let bytes = writer.finish().unwrap();
        let reader = PhysPkgReader::new(&bytes).unwrap();
        assert_eq!(
            reader.blob_for(&document).unwrap(),
            b"<document>streamed</document>"
        );
        assert_eq!(reader.blob_for(&styles).unwrap(), b"<styles/>");
    }

    #[test]
    fn physical_writer_rejects_duplicate_equivalent_and_derived_names() {
        let mut writer = PhysPkgWriter::new();
        let document = PackURI::new("/word/document.xml").unwrap();
        writer.write(&document, b"document").unwrap();

        assert!(matches!(
            writer.write(&document, b"duplicate"),
            Err(OpcError::DuplicatePartName(_))
        ));
        assert!(matches!(
            writer.write(&PackURI::new("/WORD/DOCUMENT.XML").unwrap(), b"equivalent"),
            Err(OpcError::EquivalentPartNames { .. })
        ));
        assert!(matches!(
            writer.write(
                &PackURI::new("/word/document.xml/image.bin").unwrap(),
                b"derived"
            ),
            Err(OpcError::DerivedPartNames { .. })
        ));
        let root_writer = PhysPkgWriter::new();
        assert!(matches!(
            root_writer.start_part(&PackURI::new("/").unwrap()),
            Err(OpcError::InvalidPackUri(_))
        ));

        let bytes = writer.finish().unwrap();
        let reader = PhysPkgReader::new(&bytes).unwrap();
        assert_eq!(reader.blob_for(&document).unwrap(), b"document");

        let mut child_first = PhysPkgWriter::new();
        child_first
            .write(
                &PackURI::new("/word/document.xml/image.bin").unwrap(),
                b"child",
            )
            .unwrap();
        assert!(matches!(
            child_first.write(&document, b"parent"),
            Err(OpcError::DerivedPartNames { .. })
        ));
    }

    #[test]
    fn owned_part_name_conflict_is_rejected_before_local_header_output() {
        let mut sink = Vec::new();
        let parent = PackURI::new("/word/document.xml").unwrap();
        let child = PackURI::new("/word/document.xml/image.bin").unwrap();
        let mut writer = PhysPkgWriter::with_writer(&mut sink);
        let mut part = writer.start_part(&parent).unwrap();
        part.write_all(b"parent").unwrap();
        writer = part.finish().unwrap();
        let bytes_before = writer.output_bytes() as usize;

        let conflict = writer.start_part(&child);
        assert!(matches!(conflict, Err(OpcError::DerivedPartNames { .. })));
        drop(conflict);
        assert_eq!(sink.len(), bytes_before);
    }

    struct PartialSink {
        bytes: Vec<u8>,
        limit: usize,
    }

    impl Write for PartialSink {
        fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
            let remaining = self.limit.saturating_sub(self.bytes.len());
            if remaining == 0 {
                return Err(std::io::Error::other("partial OPC sink failure"));
            }
            let accepted = remaining.min(buffer.len());
            self.bytes.extend_from_slice(&buffer[..accepted]);
            if accepted < buffer.len() {
                return Err(std::io::Error::other("partial OPC sink failure"));
            }
            Ok(accepted)
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn owned_part_writer_reports_typed_partial_progress() {
        let sink = PartialSink {
            bytes: Vec::new(),
            limit: 128,
        };
        let partname = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        let mut part = PhysPkgWriter::with_writer(sink)
            .start_stored_part(&partname)
            .unwrap();
        let write_error = part.write_all(&vec![b'x'; 4096]).unwrap_err();
        assert_eq!(write_error.kind(), std::io::ErrorKind::Other);

        let finish_result = part.finish();
        match finish_result {
            Err(OpcError::IncompleteOutput { written, source }) => {
                assert!(written > 0);
                assert!(written < 128);
                assert!(matches!(*source, OpcError::ZipError(_)));
            },
            Ok(_) => panic!("partial output unexpectedly finished"),
            Err(other) => panic!("unexpected partial output error: {other:?}"),
        }
    }

    #[test]
    fn bounded_ingress_rejects_before_zip_parsing() {
        let limits = ReadLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            PhysPkgReader::new_with_limits(b"four", limits),
            Err(OpcError::ReadLimit {
                resource: ReadResource::InputBytes,
                actual: 4,
                maximum: 3,
            })
        ));
        assert!(matches!(
            OwnedPhysPkgReader::from_reader_with_limits(Cursor::new(b"four"), limits),
            Err(OpcError::ReadLimit {
                resource: ReadResource::InputBytes,
                actual: 4,
                maximum: 3,
            })
        ));
    }

    #[test]
    fn maps_zip_member_limit_to_opc_resource() {
        let mut writer = PhysPkgWriter::new();
        writer
            .write(&PackURI::new("/first.xml").unwrap(), b"a")
            .unwrap();
        writer
            .write(&PackURI::new("/second.xml").unwrap(), b"b")
            .unwrap();
        let archive = writer.finish().unwrap();
        let limits = ReadLimits::builder()
            .max_archive_members(1)
            .unwrap()
            .max_parts(1)
            .unwrap()
            .max_relationship_parts(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            PhysPkgReader::new_with_limits(&archive, limits),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveMembers,
                actual: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn part_reads_enforce_declared_per_part_and_total_budgets() {
        let bytes = stored_archive(&[("first.bin", b"one"), ("second.bin", b"four")]);
        let first = PackURI::new("/first.bin").unwrap();
        let second = PackURI::new("/second.bin").unwrap();

        let exact = ReadLimits::builder()
            .max_part_bytes(4)
            .unwrap()
            .max_total_part_bytes(7)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, exact).unwrap();
        assert_eq!(reader.blob_for(&first).unwrap(), b"one");
        assert_eq!(reader.blob_for(&second).unwrap(), b"four");

        let per_part = ReadLimits::builder()
            .max_part_bytes(3)
            .unwrap()
            .max_total_part_bytes(7)
            .unwrap()
            .build()
            .unwrap();
        let per_part_reader = PhysPkgReader::new_with_limits(&bytes, per_part).unwrap();
        assert!(matches!(
            per_part_reader.blob_for(&second),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 4,
                maximum: 3,
            })
        ));

        let total = ReadLimits::builder()
            .max_total_part_bytes(6)
            .unwrap()
            .build()
            .unwrap();
        let total_reader = PhysPkgReader::new_with_limits(&bytes, total).unwrap();
        assert_eq!(total_reader.blob_for(&first).unwrap(), b"one");
        assert!(matches!(
            total_reader.blob_for(&second),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                actual: 7,
                maximum: 6,
            })
        ));

        let owned = OwnedPhysPkgReader::from_bytes_with_limits(bytes, per_part).unwrap();
        assert!(matches!(
            owned.blob_for(&second),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 4,
                maximum: 3,
            })
        ));
    }

    #[test]
    fn part_materializations_enforce_exact_max_parts_bound() {
        let bytes = stored_archive(&[("first.bin", b"one"), ("second.bin", b"two")]);
        let first = PackURI::new("/first.bin").unwrap();
        let second = PackURI::new("/second.bin").unwrap();
        let limits = ReadLimits::builder().max_parts(2).unwrap().build().unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, limits).unwrap();

        assert_eq!(reader.blob_for(&first).unwrap(), b"one");
        assert_eq!(reader.blob_for(&second).unwrap(), b"two");
        assert!(matches!(
            reader.blob_for(&first),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                actual: 3,
                maximum: 2,
            })
        ));
    }

    #[test]
    fn sequential_part_materializations_charge_distinct_and_repeated_uris() {
        let bytes = stored_archive(&[("first.bin", b"one"), ("second.bin", b"two")]);
        let first = PackURI::new("/first.bin").unwrap();
        let second = PackURI::new("/second.bin").unwrap();
        let limits = ReadLimits::builder().max_parts(1).unwrap().build().unwrap();

        let reader = PhysPkgReader::new_with_limits(&bytes, limits).unwrap();
        assert_eq!(reader.blob_for(&first).unwrap(), b"one");
        assert!(matches!(
            reader.blob_for(&second),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                actual: 2,
                maximum: 1,
            })
        ));

        let repeat_reader = PhysPkgReader::new_with_limits(&bytes, limits).unwrap();
        assert_eq!(repeat_reader.blob_for(&first).unwrap(), b"one");
        assert!(matches!(
            repeat_reader.blob_for(&first),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                actual: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn repeated_single_item_parallel_batches_charge_max_parts() {
        let bytes = stored_archive(&[("part.bin", b"one")]);
        let part = PackURI::new("/part.bin").unwrap();
        let limits = ReadLimits::builder().max_parts(1).unwrap().build().unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, limits).unwrap();

        assert_eq!(
            reader.blobs_parallel(std::slice::from_ref(&part)).unwrap(),
            HashMap::from([("part.bin".to_owned(), b"one".to_vec())])
        );
        assert!(matches!(
            reader.blobs_parallel(std::slice::from_ref(&part)),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                actual: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn corrupt_parallel_read_releases_part_count_and_byte_reservations() {
        let bad_payload = b"\xa5\x5a\xc3";
        let mut bytes = stored_archive(&[("good.bin", b"good"), ("bad.bin", bad_payload)]);
        let offset = bytes
            .windows(bad_payload.len())
            .position(|window| window == bad_payload)
            .unwrap();
        bytes[offset] ^= 0xff;
        let good = PackURI::new("/good.bin").unwrap();
        let bad = PackURI::new("/bad.bin").unwrap();
        let limits = ReadLimits::builder()
            .max_parts(2)
            .unwrap()
            .max_total_part_bytes(7)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, limits).unwrap();

        assert!(matches!(
            reader.blobs_parallel(&[good.clone(), bad]),
            Err(OpcError::ZipError(_))
        ));
        let budget = reader.part_budget.lock().unwrap();
        assert_eq!(budget.reserved_parts, 0);
        assert_eq!(budget.reserved_declared, 0);
        assert_eq!(budget.materialized_parts, 0);
        assert_eq!(budget.materialized_actual, 0);
        drop(budget);
        assert_eq!(reader.blob_for(&good).unwrap(), b"good");
    }

    #[test]
    fn owned_borrowed_readers_share_part_materialization_budget() {
        let bytes = stored_archive(&[("first.bin", b"one"), ("second.bin", b"two")]);
        let first = PackURI::new("/first.bin").unwrap();
        let second = PackURI::new("/second.bin").unwrap();
        let limits = ReadLimits::builder().max_parts(1).unwrap().build().unwrap();
        let owned = OwnedPhysPkgReader::from_bytes_with_limits(bytes, limits).unwrap();

        assert_eq!(owned.reader().unwrap().blob_for(&first).unwrap(), b"one");
        assert!(matches!(
            owned.reader().unwrap().blob_for(&second),
            Err(OpcError::ReadLimit {
                resource: ReadResource::Parts,
                actual: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn parallel_part_reads_are_exact_or_fail_without_partial_results() {
        let bytes = stored_archive(&[("first.bin", b"one"), ("second.bin", b"four")]);
        let first = PackURI::new("/first.bin").unwrap();
        let second = PackURI::new("/second.bin").unwrap();
        let exact = ReadLimits::builder()
            .max_part_bytes(4)
            .unwrap()
            .max_total_part_bytes(7)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, exact).unwrap();
        let blobs = reader
            .blobs_parallel(&[first.clone(), second.clone()])
            .unwrap();
        assert_eq!(blobs.get("first.bin").unwrap(), b"one");
        assert_eq!(blobs.get("second.bin").unwrap(), b"four");

        let over = ReadLimits::builder()
            .max_total_part_bytes(6)
            .unwrap()
            .build()
            .unwrap();
        let over_reader = PhysPkgReader::new_with_limits(&bytes, over).unwrap();
        assert!(matches!(
            over_reader.blobs_parallel(&[first, second]),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                actual: 7,
                maximum: 6,
            })
        ));
    }

    #[test]
    fn post_materialization_failure_releases_its_declared_reservation() {
        let bytes = stored_archive(&[("part.bin", b"one")]);
        let limits = ReadLimits::builder()
            .max_part_bytes(3)
            .unwrap()
            .max_total_part_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new_with_limits(&bytes, limits).unwrap();
        let reservation = reader.reserve_declared_parts(&[3]).unwrap();
        assert!(matches!(
            reader.commit_actual_parts(reservation, &[vec![0; 4]]),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 4,
                maximum: 3,
            })
        ));
        assert!(reader.reserve_declared_parts(&[3]).is_ok());
    }

    #[test]
    fn corrupt_present_parts_stay_zip_errors_and_bulk_reads_do_not_hide_them() {
        let bad_payload = b"\xa5\x5a\xc3";
        let mut bytes = stored_archive(&[("good.bin", b"good"), ("bad.bin", bad_payload)]);
        let offset = bytes
            .windows(bad_payload.len())
            .position(|window| window == bad_payload)
            .unwrap();
        bytes[offset] ^= 0xff;
        let good = PackURI::new("/good.bin").unwrap();
        let bad = PackURI::new("/bad.bin").unwrap();
        let missing = PackURI::new("/missing.bin").unwrap();
        let reader = PhysPkgReader::new(&bytes).unwrap();

        assert_eq!(reader.blob_for(&good).unwrap(), b"good");
        assert!(matches!(reader.blob_for(&bad), Err(OpcError::ZipError(_))));
        assert!(matches!(
            reader.blob_for(&missing),
            Err(OpcError::PartNotFound(name)) if name == "/missing.bin"
        ));

        let bulk_reader = PhysPkgReader::new(&bytes).unwrap();
        assert!(matches!(
            bulk_reader.blobs_parallel(&[good, bad]),
            Err(OpcError::ZipError(_))
        ));
        assert_eq!(
            bulk_reader
                .blob_for(&PackURI::new("/good.bin").unwrap())
                .unwrap(),
            b"good"
        );
    }
}
