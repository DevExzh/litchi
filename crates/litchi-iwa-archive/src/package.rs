//! Bounded raw package-entry ingress.
//!
//! This module owns the physical ZIP envelope used by mutable iWork package
//! snapshots. It returns ordered, owned entries and deliberately does not
//! validate format-specific paths or decode IWA messages; those policies stay
//! with the facade and the neutral component catalog respectively.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::io::{self, Write};
use std::ops::Range;
use std::sync::Arc;

use flate2::Compression;
use flate2::write::DeflateEncoder;
use litchi_core::ReadAt;

use crate::zip::{PhysicalEntry, PhysicalHeader, ZipArchive};
use crate::{Error, Limits, Result};

#[allow(
    clippy::module_name_repetitions,
    reason = "PackageState is the explicit cache-coherent state for physical package snapshots."
)]
pub use crate::package_state::{GetOrInsertError, PackageState, ParseError};
use soapberry_zip::office::StreamingArchiveWriter;

/// An immutable, clone-cheap owner for exact package bytes.
///
/// This owner accepts both the historical shared-slice representation and an
/// owned [`Vec`] representation. Converting a `Vec` retains its existing
/// payload allocation: only the small [`Arc`] control allocation is new.
/// Package facades can therefore publish a bounded reassembly buffer without
/// copying the complete package into a second allocation.
#[doc(hidden)]
#[derive(Clone)]
pub struct SharedBytes {
    storage: SharedBytesStorage,
}

#[derive(Clone)]
enum SharedBytesStorage {
    Slice(Arc<[u8]>),
    Vec(Arc<Vec<u8>>),
}

impl SharedBytes {
    /// Retain an existing immutable shared slice without copying it.
    #[must_use]
    pub const fn from_shared_slice(source: Arc<[u8]>) -> Self {
        Self {
            storage: SharedBytesStorage::Slice(source),
        }
    }

    /// Retain an owned byte vector without copying its payload allocation.
    #[must_use]
    pub(crate) fn from_owned_vec(source: Vec<u8>) -> Self {
        Self {
            storage: SharedBytesStorage::Vec(Arc::new(source)),
        }
    }

    /// Borrow the exact immutable bytes.
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        match &self.storage {
            SharedBytesStorage::Slice(source) => source,
            SharedBytesStorage::Vec(source) => source,
        }
    }

    /// Return whether two handles retain the same underlying allocation.
    #[must_use]
    pub fn ptr_eq(left: &Self, right: &Self) -> bool {
        match (&left.storage, &right.storage) {
            (SharedBytesStorage::Slice(left), SharedBytesStorage::Slice(right)) => {
                Arc::ptr_eq(left, right)
            },
            (SharedBytesStorage::Vec(left), SharedBytesStorage::Vec(right)) => {
                Arc::ptr_eq(left, right)
            },
            (SharedBytesStorage::Slice(_), SharedBytesStorage::Vec(_))
            | (SharedBytesStorage::Vec(_), SharedBytesStorage::Slice(_)) => false,
        }
    }

    /// Clone the historical shared-slice handle when this owner has one.
    ///
    /// Owned-vector backing deliberately returns `None`; materializing an
    /// `Arc<[u8]>` would require another package-sized allocation and copy.
    #[must_use]
    pub(crate) fn shared_slice(&self) -> Option<Arc<[u8]>> {
        match &self.storage {
            SharedBytesStorage::Slice(source) => Some(Arc::clone(source)),
            SharedBytesStorage::Vec(_) => None,
        }
    }
}

impl AsRef<[u8]> for SharedBytes {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

impl std::ops::Deref for SharedBytes {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.as_slice()
    }
}

impl fmt::Debug for SharedBytes {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let backing = match self.storage {
            SharedBytesStorage::Slice(_) => "shared-slice",
            SharedBytesStorage::Vec(_) => "owned-vector",
        };
        formatter
            .debug_struct("SharedBytes")
            .field("len", &self.len())
            .field("backing", &backing)
            .finish()
    }
}

impl PartialEq for SharedBytes {
    fn eq(&self, other: &Self) -> bool {
        Self::ptr_eq(self, other) || self.as_slice() == other.as_slice()
    }
}

impl Eq for SharedBytes {}

#[derive(Debug)]
struct PreparedEdit {
    compressed: Vec<u8>,
    crc32: u32,
    uncompressed_size: u64,
    descriptor_start: Option<usize>,
}

#[derive(Debug, Clone, Copy)]
struct ReassemblyShape {
    base_offset: u64,
}

#[derive(Debug, Clone, Copy)]
struct MatchedEdit<'a> {
    index: usize,
    edit: EntryEdit<'a>,
    descriptor_start: Option<usize>,
}

#[derive(Debug)]
struct BoundedBuffer {
    bytes: Vec<u8>,
    maximum: usize,
}

impl BoundedBuffer {
    const fn new(maximum: usize) -> Self {
        Self {
            bytes: Vec::new(),
            maximum,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for BoundedBuffer {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let new_len = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("bounded ZIP buffer length overflows usize"))?;
        if new_len > self.maximum {
            return Err(io::Error::other("bounded ZIP output limit exceeded"));
        }
        self.bytes
            .try_reserve(bytes.len())
            .map_err(|error| io::Error::other(format!("could not allocate ZIP output: {error}")))?;
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// A raw ZIP entry record retained for an exact preserve-mode write.
///
/// The local record includes the local header, compressed bytes, data
/// descriptor, and any bytes before the next local record or central
/// directory. The central record is retained separately because ZIP stores it
/// in a different part of the archive. Neither view is normalized.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawEntryRecord {
    source: SharedBytes,
    local_record: Range<usize>,
    compressed_data: Range<usize>,
    central_directory_record: Range<usize>,
}

impl RawEntryRecord {
    fn new(source: SharedBytes, entry: &PhysicalEntry) -> Self {
        Self {
            source,
            local_record: entry.local_record(),
            compressed_data: entry.compressed_data_range(),
            central_directory_record: entry.central_record(),
        }
    }

    /// Borrow the exact local record bytes.
    #[must_use]
    pub fn local_record(&self) -> &[u8] {
        &self.source[self.local_record.clone()]
    }

    /// Borrow the exact compressed data bytes.
    #[must_use]
    pub fn compressed_data(&self) -> &[u8] {
        &self.source[self.compressed_data.clone()]
    }

    /// Borrow the exact central-directory record bytes.
    #[must_use]
    pub fn central_directory_record(&self) -> &[u8] {
        &self.source[self.central_directory_record.clone()]
    }
}

/// The raw timestamp fields from one ZIP header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DosDateTime {
    time: u16,
    date: u16,
}

impl DosDateTime {
    /// Return the original packed DOS time field.
    #[must_use]
    pub const fn time(self) -> u16 {
        self.time
    }

    /// Return the original packed DOS date field.
    #[must_use]
    pub const fn date(self) -> u16 {
        self.date
    }
}

/// Physical metadata from one local or central ZIP header.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderMetadata {
    version_needed: u16,
    flags: u16,
    compression_method: u16,
    last_modified: DosDateTime,
    name: Box<[u8]>,
    extra: Box<[u8]>,
    comment: Box<[u8]>,
}

impl HeaderMetadata {
    /// Return the ZIP version needed to extract this header's member.
    #[must_use]
    pub const fn version_needed(&self) -> u16 {
        self.version_needed
    }

    /// Return the original general-purpose bit flags.
    #[must_use]
    pub const fn flags(&self) -> u16 {
        self.flags
    }

    /// Return the original numeric ZIP compression method.
    #[must_use]
    pub const fn compression_method(&self) -> u16 {
        self.compression_method
    }

    /// Return the original packed DOS modification timestamp.
    #[must_use]
    pub const fn last_modified(&self) -> DosDateTime {
        self.last_modified
    }

    /// Borrow the exact header filename bytes.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        &self.name
    }

    /// Borrow the exact header extra-field bytes.
    #[must_use]
    pub fn extra(&self) -> &[u8] {
        &self.extra
    }

    /// Borrow the exact header comment bytes.
    #[must_use]
    pub fn comment(&self) -> &[u8] {
        &self.comment
    }
}

/// All physical metadata needed to describe one ZIP member.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EntryMetadata {
    local: HeaderMetadata,
    central: HeaderMetadata,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
}

impl EntryMetadata {
    fn new(entry: &PhysicalEntry) -> Self {
        Self {
            local: header_metadata(entry.local_header()),
            central: header_metadata(entry.central_header()),
            compressed_size: entry.compressed_size(),
            uncompressed_size: entry.uncompressed_size(),
            crc32: entry.crc32(),
        }
    }

    /// Return metadata from the local file header.
    #[must_use]
    pub const fn local(&self) -> &HeaderMetadata {
        &self.local
    }

    /// Return metadata from the central-directory header.
    #[must_use]
    pub const fn central(&self) -> &HeaderMetadata {
        &self.central
    }

    /// Return the central directory's declared compressed size.
    #[must_use]
    pub const fn compressed_size(&self) -> u64 {
        self.compressed_size
    }

    /// Return the central directory's declared uncompressed size.
    #[must_use]
    pub const fn uncompressed_size(&self) -> u64 {
        self.uncompressed_size
    }

    /// Return the central directory's declared CRC-32.
    #[must_use]
    pub const fn crc32(&self) -> u32 {
        self.crc32
    }
}

/// The safe distinction between decoded content and an unsupported physical
/// member.
#[derive(Debug, Clone, Copy)]
pub enum EntryPayload<'a> {
    /// The member was decoded by the bounded ZIP reader.
    Decoded(&'a [u8]),
    /// The member's compression method is unsupported; use its raw record.
    Opaque(&'a RawEntryRecord),
}

/// Replace the decoded payload of one existing ZIP member during a bounded
/// physical reassembly.
///
/// An edit is addressed by the catalog's normalized member name. It cannot
/// add, remove, rename, or reorder members. Store and Deflate members are
/// re-encoded with their original compression method; unsupported methods are
/// rejected when selected for editing and remain byte-for-byte untouched when
/// not selected.
#[derive(Debug, Clone, Copy)]
pub struct EntryEdit<'a> {
    name: &'a str,
    data: &'a [u8],
}

impl<'a> EntryEdit<'a> {
    /// Build an edit for one existing normalized member name.
    #[must_use]
    pub const fn new(name: &'a str, data: &'a [u8]) -> Self {
        Self { name, data }
    }

    /// Return the normalized member name used for lookup.
    #[must_use]
    pub const fn name(self) -> &'a str {
        self.name
    }

    /// Borrow the replacement decoded payload.
    #[must_use]
    pub const fn data(self) -> &'a [u8] {
        self.data
    }
}

/// A process-local pair of exact immutable shared-slice artifacts.
#[derive(Clone, PartialEq, Eq)]
pub struct ExactArtifacts {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
}

impl fmt::Debug for ExactArtifacts {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ExactArtifacts")
            .field("source_fingerprint", &self.source_fingerprint)
            .field("target_fingerprint", &self.target_fingerprint)
            .finish_non_exhaustive()
    }
}

impl ExactArtifacts {
    /// Retain one exact source/target artifact pair.
    /// Return the source artifact's compact diagnostic fingerprint.
    #[must_use]
    pub fn new(source: Arc<[u8]>, target: Arc<[u8]>) -> Self {
        let source_fingerprint = diagnostic_fingerprint(&source);
        let target_fingerprint = if Arc::ptr_eq(&source, &target) {
            source_fingerprint
        } else {
            diagnostic_fingerprint(&target)
        };
        Self {
            source,
            target,
            source_fingerprint,
            target_fingerprint,
        }
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return whether `candidate` exactly matches the retained source.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether source and target are byte-for-byte identical.
    #[must_use]
    pub fn authorizes_source(&self, candidate: &Arc<[u8]>) -> bool {
        Arc::ptr_eq(candidate, &self.source) || candidate.as_ref() == self.source.as_ref()
    }

    /// Return the exact target-to-source inverse pair.
    #[must_use]
    pub fn is_byte_noop(&self) -> bool {
        Arc::ptr_eq(&self.source, &self.target) || self.source.as_ref() == self.target.as_ref()
    }

    /// Clone the retained source allocation without copying bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
        }
    }

    /// Clone the retained target allocation without copying bytes.
    #[must_use]
    pub fn source(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    #[must_use]
    pub fn target(&self) -> Arc<[u8]> {
        Arc::clone(&self.target)
    }
}

/// A process-local pair of exact immutable package byte owners.
///
/// Focused semantic patches can retain this pair instead of independently
/// reimplementing source-byte authorization and source/target ownership. The
/// complete source and target allocations are retained behind private,
/// immutable handles for the lifetime of the pair. Cloning and inversion are
/// `O(1)` [`Arc`] operations; constructing the pair reads each distinct
/// artifact allocation once to cache compact diagnostic fingerprints.
///
/// Fingerprints are diagnostics only. They never authorize patch application:
/// [`Self::authorizes_owner`] requires either allocation identity or one exact
/// byte comparison. This type is process-local in-memory state, not a compact
/// or durable patch serialization format.
#[derive(Clone, PartialEq, Eq)]
#[doc(hidden)]
pub struct OwnedExactArtifacts {
    source: SharedBytes,
    target: SharedBytes,
    source_fingerprint: u64,
    target_fingerprint: u64,
}

impl fmt::Debug for OwnedExactArtifacts {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("OwnedExactArtifacts")
            .field("source_fingerprint", &self.source_fingerprint)
            .field("target_fingerprint", &self.target_fingerprint)
            .finish_non_exhaustive()
    }
}

impl OwnedExactArtifacts {
    /// Retain one exact source/target artifact pair.
    ///
    /// The supplied allocations are owned immutably and are never exposed as
    /// mutable or borrowed raw slices. Callers remain responsible for
    /// constructing the target through their bounded, semantically verified
    /// focused transaction.
    ///
    /// # Costs
    ///
    /// Reads each distinct complete artifact once to cache diagnostics.
    /// Retain one exact source/target owner pair without copying payloads.
    #[doc(hidden)]
    #[must_use]
    pub fn new(source: SharedBytes, target: SharedBytes) -> Self {
        let source_fingerprint = diagnostic_fingerprint(&source);
        let target_fingerprint = if SharedBytes::ptr_eq(&source, &target) {
            source_fingerprint
        } else {
            diagnostic_fingerprint(&target)
        };
        Self {
            source,
            target,
            source_fingerprint,
            target_fingerprint,
        }
    }

    /// Return the source artifact's compact diagnostic fingerprint.
    ///
    /// This value is not collision-resistant and never authorizes source
    /// identity. Use [`Self::authorizes_owner`] for exact authorization.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    ///
    /// This value is not collision-resistant and is not an artifact identity.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether `candidate` is the exact retained source artifact.
    #[doc(hidden)]
    ///
    /// Allocation identity is an `O(1)` fast path. A different allocation is
    /// authorized only by one complete byte equality comparison. Cached
    /// fingerprints do not participate in this decision.
    #[must_use]
    pub fn authorizes_owner(&self, candidate: &SharedBytes) -> bool {
        SharedBytes::ptr_eq(candidate, &self.source) || candidate.as_ref() == self.source.as_ref()
    }

    /// Return whether source and target are byte-for-byte identical.
    ///
    /// Allocation identity is an `O(1)` fast path; otherwise this reads the
    /// artifacts for exact equality. Fingerprint equality is insufficient.
    #[must_use]
    pub fn is_byte_noop(&self) -> bool {
        SharedBytes::ptr_eq(&self.source, &self.target)
            || self.source.as_ref() == self.target.as_ref()
    }

    /// Return the exact target-to-source inverse pair.
    ///
    /// This only clones and swaps shared handles and cached diagnostics.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
        }
    }

    /// Clone the retained exact source owner without copying bytes.
    #[doc(hidden)]
    #[must_use]
    pub fn source_owner(&self) -> SharedBytes {
        self.source.clone()
    }

    /// Clone the retained exact target owner without copying bytes.
    #[doc(hidden)]
    #[must_use]
    pub fn target_owner(&self) -> SharedBytes {
        self.target.clone()
    }
}

/// One ordered package member with its physical ZIP provenance.
#[derive(Debug)]
pub struct Entry {
    name: Box<str>,
    data: Vec<u8>,
    raw_name: Box<[u8]>,
    metadata: EntryMetadata,
    raw_record: RawEntryRecord,
    opaque: bool,
}

impl Entry {
    fn new(
        name: &str,
        data: Vec<u8>,
        raw_name: Box<[u8]>,
        metadata: EntryMetadata,
        raw_record: RawEntryRecord,
        opaque: bool,
    ) -> Self {
        Self {
            name: name.into(),
            data,
            raw_name,
            metadata,
            raw_record,
            opaque,
        }
    }

    /// Borrow the physical member name in source order.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the decoded member payload.
    ///
    /// For an opaque entry this returns the raw compressed byte stream for
    /// compatibility with the original payload accessor. Use the structured
    /// payload or opaque flag before interpreting it.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Borrow the exact central-directory filename bytes.
    #[must_use]
    pub fn raw_name(&self) -> &[u8] {
        &self.raw_name
    }

    /// Borrow the preserved physical metadata.
    #[must_use]
    pub const fn metadata(&self) -> &EntryMetadata {
        &self.metadata
    }

    /// Borrow the preserved raw ZIP records.
    #[must_use]
    pub const fn raw_record(&self) -> &RawEntryRecord {
        &self.raw_record
    }

    /// Return whether the compression method was not decoded by this crate.
    #[must_use]
    pub const fn is_opaque(&self) -> bool {
        self.opaque
    }

    /// Return decoded content or the opaque raw-record provenance.
    #[must_use]
    pub fn payload(&self) -> EntryPayload<'_> {
        if self.opaque {
            EntryPayload::Opaque(&self.raw_record)
        } else {
            EntryPayload::Decoded(&self.data)
        }
    }

    /// Consume the entry without cloning its name or payload.
    ///
    /// For an opaque entry, the returned bytes are the raw compressed stream;
    /// callers must retain the raw record when they need preserve-mode
    /// serialization.
    #[must_use]
    pub fn into_parts(self) -> (String, Vec<u8>) {
        (self.name.into(), self.data)
    }
}

/// Physical origin retained by an immutable package snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceProvenance {
    /// Logical entries map directly to the authoritative ZIP source.
    ExactZip,
    /// Logical entries were normalized from a nested legacy `Index.zip`.
    LegacyZip,
    /// A modern ZIP was projected to a fixed semantic member set.
    SemanticZip,
    /// A legacy ZIP was projected to a fixed semantic member set.
    LegacySemanticZip,
}

/// A fixed logical-member admission profile applied before ZIP payload decode.
///
/// This unstable integration type lets a format owner select a narrow set of
/// exact normalized logical names whose declared physical shape must be
/// checked before any package entries are materialized. It is intentionally
/// not a caller-extensible path filter.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LogicalEntryLimits {
    profile: LogicalEntryLimitProfile,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum LogicalEntryLimitProfile {
    SemanticComponents,
    SemanticMetadata,
}

impl LogicalEntryLimits {
    /// Select only IWA components and materialize no package sidecars.
    pub const SEMANTIC_COMPONENTS: Self = Self {
        profile: LogicalEntryLimitProfile::SemanticComponents,
    };

    /// iWork's three canonical semantic metadata authorities, each capped at
    /// 64 KiB and required to use a supported ZIP compression method.
    pub const SEMANTIC_METADATA: Self = Self {
        profile: LogicalEntryLimitProfile::SemanticMetadata,
    };

    /// Compatibility alias for the original Pages-specific spelling.
    pub const PAGES_METADATA: Self = Self::SEMANTIC_METADATA;

    const MAX_SEMANTIC_METADATA_BYTES: u64 = 64 * 1024;

    const fn maximum_for(self, logical_name: &[u8]) -> Option<u64> {
        match self.profile {
            LogicalEntryLimitProfile::SemanticComponents => None,
            LogicalEntryLimitProfile::SemanticMetadata => match logical_name {
                b"Metadata/Properties.plist"
                | b"Metadata/BuildVersionHistory.plist"
                | b"Metadata/DocumentIdentifier" => Some(Self::MAX_SEMANTIC_METADATA_BYTES),
                _ => None,
            },
        }
    }

    const fn includes_metadata(self) -> bool {
        matches!(self.profile, LogicalEntryLimitProfile::SemanticMetadata)
    }
}

/// Ordered raw entries extracted from one physical iWork ZIP input.
#[derive(Debug)]
pub struct Catalog {
    entries: Vec<Entry>,
    source: SharedBytes,
    source_is_exact: bool,
    semantic_profile: Option<LogicalEntryLimitProfile>,
    legacy_outer_prefix: Option<Box<[u8]>>,
}

/// One exact canonical iWork metadata member selected from raw ZIP
/// provenance.
#[doc(hidden)]
#[derive(Debug, Clone, Copy)]
pub struct PackageMetadataEntry<'a> {
    data: &'a [u8],
    opaque: bool,
}

impl<'a> PackageMetadataEntry<'a> {
    /// Borrow the selected member payload.
    #[must_use]
    pub const fn data(self) -> &'a [u8] {
        self.data
    }

    /// Whether the selected authority uses unsupported compression.
    #[must_use]
    pub const fn is_opaque(self) -> bool {
        self.opaque
    }
}

/// The fixed three-member iWork metadata projection selected from exact raw
/// ZIP authority names.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, Default)]
pub struct PackageMetadataSidecars<'a> {
    properties: Option<PackageMetadataEntry<'a>>,
    build_version_history: Option<PackageMetadataEntry<'a>>,
    document_identifier: Option<PackageMetadataEntry<'a>>,
}

impl<'a> PackageMetadataSidecars<'a> {
    /// Borrow exact `Metadata/Properties.plist`, if present.
    #[must_use]
    pub const fn properties_plist(self) -> Option<PackageMetadataEntry<'a>> {
        self.properties
    }

    /// Borrow exact `Metadata/BuildVersionHistory.plist`, if present.
    #[must_use]
    pub const fn build_version_history_plist(self) -> Option<PackageMetadataEntry<'a>> {
        self.build_version_history
    }

    /// Borrow exact `Metadata/DocumentIdentifier`, if present.
    #[must_use]
    pub const fn document_identifier(self) -> Option<PackageMetadataEntry<'a>> {
        self.document_identifier
    }
}

impl Catalog {
    /// Select iWork's exact raw metadata authorities from the retained package.
    ///
    /// Flat packages require byte-for-byte central names. Legacy packages may
    /// remove only the one explicit raw outer prefix that led to `Index.zip`.
    #[doc(hidden)]
    pub fn __semantic_metadata_sidecars(&self) -> Result<PackageMetadataSidecars<'_>> {
        let mut selected = PackageMetadataSidecars::default();
        for entry in &self.entries {
            let raw_name = entry.raw_name();
            let logical_name = match self.legacy_outer_prefix.as_deref() {
                Some(prefix) => raw_name.strip_prefix(prefix),
                None => Some(raw_name),
            };
            let Some(logical_name) = logical_name else {
                continue;
            };
            let authority = match logical_name {
                b"Metadata/Properties.plist" => Some("Metadata/Properties.plist"),
                b"Metadata/BuildVersionHistory.plist" => Some("Metadata/BuildVersionHistory.plist"),
                b"Metadata/DocumentIdentifier" => Some("Metadata/DocumentIdentifier"),
                _ => None,
            };
            let Some(authority) = authority else {
                continue;
            };
            if entry.metadata().local().name() != entry.metadata().central().name() {
                return Err(Error::InvalidBundle(format!(
                    "canonical logical entry {authority} has mismatched local and central names"
                )));
            }
            if entry.metadata().local().compression_method()
                != entry.metadata().central().compression_method()
            {
                return Err(Error::InvalidBundle(format!(
                    "canonical logical entry {authority} has mismatched local and central compression methods"
                )));
            }
            let value = PackageMetadataEntry {
                data: entry.data(),
                opaque: entry.is_opaque(),
            };
            match logical_name {
                b"Metadata/Properties.plist" => selected.properties = Some(value),
                b"Metadata/BuildVersionHistory.plist" => {
                    selected.build_version_history = Some(value);
                },
                b"Metadata/DocumentIdentifier" => selected.document_identifier = Some(value),
                _ => {},
            }
        }
        Ok(selected)
    }

    /// Compatibility alias for the original Pages-specific spelling.
    #[doc(hidden)]
    pub fn __pages_metadata_sidecars(&self) -> Result<PackageMetadataSidecars<'_>> {
        self.__semantic_metadata_sidecars()
    }

    pub(crate) const fn has_semantic_profile(&self) -> bool {
        self.semantic_profile.is_some()
    }

    /// Parse a package with the default physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope or any configured physical
    /// limit is invalid.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse an immutable, already-owned package source with the default
    /// physical limits.
    ///
    /// The catalog retains this exact [`Arc`] allocation for preserve-mode
    /// writes and does not copy the source bytes. The shared byte slice is
    /// immutable, so entries and raw ZIP records can safely borrow from it
    /// for the lifetime of the catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope or any configured physical
    /// limit is invalid.
    pub fn from_shared_bytes(source: Arc<[u8]>) -> Result<Self> {
        Self::from_shared_bytes_with_limits(source, Limits::default())
    }

    /// Parse a package under caller-selected physical limits.
    ///
    /// A legacy package containing `.../Index.zip` is flattened into the
    /// modern entry order used by the mutable facade: nested IWA members are
    /// emitted first, followed by outer entries with the legacy prefix
    /// removed. The nested archive must contain only IWA members.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope, nested index, duplicate entry,
    /// or configured physical limit is invalid.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_bytes_with_optional_logical_entry_limits(bytes, limits, None)
    }

    /// Parse borrowed package bytes while applying a fixed logical-member
    /// profile before any ZIP payload is decoded.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes_with_limits`], plus a
    /// refusal when a selected logical authority has unsupported compression
    /// or exceeds its declared uncompressed-size ceiling.
    #[doc(hidden)]
    pub fn __from_bytes_with_logical_entry_limits(
        bytes: &[u8],
        limits: Limits,
        logical_entry_limits: LogicalEntryLimits,
    ) -> Result<Self> {
        Self::from_bytes_with_optional_logical_entry_limits(
            bytes,
            limits,
            Some(logical_entry_limits),
        )
    }

    fn from_bytes_with_optional_logical_entry_limits(
        bytes: &[u8],
        limits: Limits,
        logical_entry_limits: Option<LogicalEntryLimits>,
    ) -> Result<Self> {
        let checked_limits = limits.validate()?;
        let input_size = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("borrowed ZIP input length does not fit u64".to_owned())
        })?;
        checked_limits.check_input_size(input_size, "borrowed ZIP input")?;
        let mut source_bytes = Vec::new();
        source_bytes
            .try_reserve_exact(bytes.len())
            .map_err(|_error| Error::Allocation {
                resource: "borrowed ZIP source bytes",
                amount: bytes.len(),
            })?;
        source_bytes.extend_from_slice(bytes);
        let shared_source: Arc<[u8]> = source_bytes.into();
        Self::from_source_with_checked_limits(
            SharedBytes::from_shared_slice(shared_source),
            checked_limits,
            logical_entry_limits,
        )
    }

    /// Parse a package from an immutable positional source under the default
    /// physical limits.
    ///
    /// The source is snapshotted through [`ReadAt`] without using a shared
    /// cursor. Its [`litchi_core::SourceVersion`] must remain unchanged for
    /// the complete bounded read; otherwise the package is rejected rather
    /// than publishing a mixed-source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the source cannot be read, changes during the
    /// snapshot, exceeds the physical limits, or is not a valid package.
    pub fn from_read_at(source: &dyn ReadAt) -> Result<Self> {
        Self::from_read_at_with_limits(source, Limits::default())
    }

    /// Parse an immutable, already-owned package source under caller-selected
    /// physical limits.
    ///
    /// The catalog retains the supplied [`Arc`] allocation for preserve-mode
    /// writes and does not copy the source bytes. A legacy package containing
    /// `.../Index.zip` is flattened into the modern entry order used by the
    /// mutable facade: nested IWA members are emitted first, followed by outer
    /// entries with the legacy prefix removed. The nested archive must contain
    /// only IWA members.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope, nested index, duplicate entry,
    /// or configured physical limit is invalid.
    pub fn from_shared_bytes_with_limits(source: Arc<[u8]>, limits: Limits) -> Result<Self> {
        Self::from_shared_bytes_with_optional_logical_entry_limits(source, limits, None)
    }

    /// Parse an owned package vector without copying its payload allocation.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope or a configured physical limit
    /// is invalid.
    #[cfg(any(test, feature = "internal-iwork-source"))]
    pub(crate) fn from_owned_bytes_with_limits(source: Vec<u8>, limits: Limits) -> Result<Self> {
        let checked_limits = limits.validate()?;
        Self::from_source_with_checked_limits(
            SharedBytes::from_owned_vec(source),
            checked_limits,
            None,
        )
    }

    /// Parse an immutable source owner without copying package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope or a configured physical limit
    /// is invalid.
    #[cfg(feature = "internal-iwork-source")]
    pub(crate) fn from_source_owner_with_limits(
        source: SharedBytes,
        limits: Limits,
    ) -> Result<Self> {
        let checked_limits = limits.validate()?;
        Self::from_source_with_checked_limits(source, checked_limits, None)
    }

    /// Parse shared package bytes while applying a fixed logical-member
    /// profile before any ZIP payload is decoded.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_shared_bytes_with_limits`],
    /// plus a refusal when a selected logical authority has unsupported
    /// compression or exceeds its declared uncompressed-size ceiling.
    #[doc(hidden)]
    pub fn __from_shared_bytes_with_logical_entry_limits(
        source: Arc<[u8]>,
        limits: Limits,
        logical_entry_limits: LogicalEntryLimits,
    ) -> Result<Self> {
        Self::from_shared_bytes_with_optional_logical_entry_limits(
            source,
            limits,
            Some(logical_entry_limits),
        )
    }

    fn from_shared_bytes_with_optional_logical_entry_limits(
        source: Arc<[u8]>,
        limits: Limits,
        logical_entry_limits: Option<LogicalEntryLimits>,
    ) -> Result<Self> {
        let checked_limits = limits.validate()?;
        Self::from_source_with_checked_limits(
            SharedBytes::from_shared_slice(source),
            checked_limits,
            logical_entry_limits,
        )
    }

    /// Parse a package from an immutable positional source under caller-
    /// selected physical limits.
    ///
    /// The source is read only after its length has passed the input ceiling,
    /// so a rejected source cannot trigger an unbounded allocation. The
    /// source version is checked again before the snapshot is parsed.
    ///
    /// # Errors
    ///
    /// Returns an error when the source cannot be read, changes during the
    /// snapshot, exceeds the physical limits, or is not a valid package.
    pub fn from_read_at_with_limits(source: &dyn ReadAt, limits: Limits) -> Result<Self> {
        let checked_limits = limits.validate()?;
        let source_bytes = read_source(source, checked_limits)?;
        Self::from_source_with_checked_limits(
            SharedBytes::from_shared_slice(source_bytes),
            checked_limits,
            None,
        )
    }

    fn from_source_with_checked_limits(
        source: SharedBytes,
        checked_limits: Limits,
        logical_entry_limits: Option<LogicalEntryLimits>,
    ) -> Result<Self> {
        let archive = ZipArchive::new_with_limits(source.as_ref(), checked_limits)?;
        if crate::zip::is_encrypted(&archive) {
            return Err(Error::Encrypted);
        }

        let semantic_profile = logical_entry_limits.map(|limits| limits.profile);
        if semantic_profile.is_some() {
            reject_semantic_aliases(&archive)?;
        }
        let (has_direct_iwa, nested_name) = if semantic_profile.is_some() {
            let has_direct = archive
                .physical_entries()
                .any(|entry| semantic_iwa_name(entry).is_some());
            (has_direct, semantic_nested_index_name(&archive)?)
        } else {
            (
                archive.file_names().any(crate::zip::is_iwa_name),
                crate::zip::nested_index_name(&archive)?,
            )
        };
        if has_direct_iwa && nested_name.is_some() {
            return Err(Error::InvalidBundle(
                "iWork package mixes direct IWA members with a legacy Index.zip".to_owned(),
            ));
        }
        if let Some(logical_entry_limits) = logical_entry_limits {
            if let Some(index_name) = nested_name.as_deref() {
                preflight_legacy_outer_logical_entries(&archive, index_name, logical_entry_limits)?;
            } else {
                preflight_flat_logical_entries(&archive, logical_entry_limits)?;
            }
        }
        let include_semantic_metadata =
            logical_entry_limits.is_some_and(LogicalEntryLimits::includes_metadata);
        let (entries, source_is_exact, legacy_outer_prefix) = if has_direct_iwa {
            let entries = if semantic_profile.is_some() {
                collect_semantic_flat(&archive, &source, include_semantic_metadata)?
            } else {
                collect_flat(&archive, &source)?
            };
            (entries, semantic_profile.is_none(), None)
        } else if let Some(index_name) = nested_name {
            let prefix: Box<[u8]> = legacy_raw_outer_prefix(&archive, &index_name)?.into();
            let entries = if semantic_profile.is_some() {
                collect_semantic_legacy(
                    &archive,
                    &index_name,
                    checked_limits,
                    &source,
                    &prefix,
                    include_semantic_metadata,
                )?
            } else {
                collect_legacy(&archive, &index_name, checked_limits, &source)?
            };
            (entries, false, Some(prefix))
        } else {
            let entries = if semantic_profile.is_some() {
                collect_semantic_flat(&archive, &source, include_semantic_metadata)?
            } else {
                collect_flat(&archive, &source)?
            };
            (entries, semantic_profile.is_none(), None)
        };
        Ok(Catalog {
            entries,
            source,
            source_is_exact,
            semantic_profile,
            legacy_outer_prefix,
        })
    }

    /// Return the number of extracted entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether no entries were extracted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Clone the immutable source handle without copying its bytes.
    ///
    /// The handle is useful to a higher-level package owner that wants to
    /// retain an exact preserve-mode no-op path while it builds its own
    /// semantic entry index.
    #[must_use]
    pub fn shared_source(&self) -> Arc<[u8]> {
        self.source
            .shared_slice()
            .unwrap_or_else(|| Arc::<[u8]>::from(self.source.as_slice()))
    }

    /// Clone the exact immutable source owner without copying package bytes.
    #[must_use]
    #[cfg(any(test, feature = "internal-iwork-source"))]
    pub(crate) fn source_owner(&self) -> SharedBytes {
        self.source.clone()
    }

    /// Borrow the authoritative immutable source bytes.
    ///
    /// This view is allocation-free and remains valid for the lifetime of the
    /// catalog. Higher-level format facades should not expose it as semantic
    /// document state.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.source.as_slice()
    }

    /// Return whether the catalog's logical entries still describe the
    /// original ZIP envelope exactly.
    #[must_use]
    pub const fn source_is_exact(&self) -> bool {
        self.source_is_exact
    }

    /// Return the physical origin of this logical package snapshot.
    #[must_use]
    pub const fn source_provenance(&self) -> SourceProvenance {
        if self.semantic_profile.is_some() && self.legacy_outer_prefix.is_some() {
            SourceProvenance::LegacySemanticZip
        } else if self.semantic_profile.is_some() {
            SourceProvenance::SemanticZip
        } else if self.source_is_exact {
            SourceProvenance::ExactZip
        } else {
            SourceProvenance::LegacyZip
        }
    }

    /// Borrow entries in their preserved source order.
    pub fn iter(&self) -> impl Iterator<Item = &Entry> {
        self.entries.iter()
    }

    /// Write the original ZIP bytes without rebuilding or normalizing any
    /// member, central record, archive comment, or opaque entry.
    ///
    /// This is the preserve-mode no-op path. Catalog has no mutating
    /// operations in this bounded slice, so the source remains authoritative.
    ///
    /// # Errors
    ///
    /// Returns an error when the caller-provided sink rejects the source bytes.
    pub fn write_to<W: Write>(&self, mut sink: W) -> Result<()> {
        sink.write_all(&self.source)?;
        Ok(())
    }

    /// Return an exact byte-for-byte copy of the source ZIP.
    ///
    /// # Errors
    ///
    /// Returns an error if the source copy cannot be allocated.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.source.len())
            .map_err(|_error| Error::Allocation {
                resource: "catalog source bytes",
                amount: self.source.len(),
            })?;
        bytes.extend_from_slice(&self.source);
        Ok(bytes)
    }

    /// Reassemble a flat package after replacing existing decoded member
    /// payloads.
    ///
    /// This is a deliberately narrow physical-fidelity transaction:
    /// directory and member order, raw names, local and central extra fields,
    /// comments, attributes, archive comments, and untouched compressed data
    /// are retained. Only CRCs, sizes, local offsets, and the central-directory
    /// offset are changed when required by an edit. Store and Deflate members
    /// are supported; unsupported compression methods may remain opaque but
    /// cannot be selected for editing. Legacy nested `Index.zip` catalogs and
    /// ZIP64 layouts are rejected for edits rather than silently normalized.
    ///
    /// All validation and replacement compression completes before the output
    /// buffer is allocated. The complete output is bounded by
    /// [`Limits::max_input_bytes`], then returned as one committed artifact.
    /// An empty edit list uses the existing exact source path, including for
    /// legacy catalogs.
    ///
    /// # Errors
    ///
    /// Returns an error when an edit is ambiguous, exceeds a physical limit,
    /// selects an unsupported method, or the source layout cannot be safely
    /// patched without losing physical metadata.
    pub fn reassemble_to_bytes(&self, edits: &[EntryEdit<'_>], limits: Limits) -> Result<Vec<u8>> {
        self.reassemble_with_deletions_to_bytes(edits, &[], limits)
    }

    /// Reassemble a flat package after replacing and deleting existing
    /// normalized members.
    ///
    /// Deletion names use the same exact normalized-name lookup as
    /// [`EntryEdit`]. Every edited or deleted name must select exactly one
    /// existing non-directory member, names may not be repeated, and one
    /// member cannot be both edited and deleted. Deleting an opaque member is
    /// supported because its payload is never decoded; editing one remains an
    /// error.
    ///
    /// Retained members keep their source order, raw local and central names,
    /// header metadata, comments, and compressed bytes. Only edited payload
    /// fields, retained local-header offsets, and the central-directory and
    /// end-of-central-directory summaries are changed. Legacy normalized
    /// catalogs and ZIP64 layouts are rejected whenever either mutation list
    /// is non-empty. Empty edit and deletion lists return the exact source
    /// bytes, including for legacy catalogs.
    ///
    /// # Errors
    ///
    /// Returns an error when selection is missing or ambiguous, an edit
    /// selects an opaque member, a physical/resource limit is exceeded, or
    /// the source layout cannot be patched without losing metadata.
    pub fn reassemble_with_deletions_to_bytes(
        &self,
        edits: &[EntryEdit<'_>],
        deleted_names: &[&str],
        limits: Limits,
    ) -> Result<Vec<u8>> {
        let checked_limits = limits.validate()?;
        let source_size = u64::try_from(self.source.len()).map_err(|_error| {
            Error::InvalidBundle("catalog source length does not fit u64".to_owned())
        })?;
        checked_limits.check_input_size(source_size, "catalog source")?;

        if edits.is_empty() && deleted_names.is_empty() {
            checked_limits.check_output_size(source_size)?;
            return self.to_bytes();
        }
        if !self.source_is_exact {
            let reason = if self.semantic_profile.is_some() {
                "semantic catalogs have a projected logical entry set"
            } else {
                "legacy nested Index.zip catalogs have normalized logical entries"
            };
            return Err(Error::Reassembly(reason.to_owned()));
        }

        let archive = ZipArchive::new_with_limits(self.source.as_ref(), checked_limits)?;
        let shape = validate_reassembly_shape(&archive)?;
        let (prepared, deleted) =
            prepare_mutations(&archive, edits, deleted_names, checked_limits)?;
        let output_size = reassembled_output_size(&archive, &prepared, &deleted)?;
        checked_limits.check_output_size(output_size)?;
        let output_len = usize::try_from(output_size).map_err(|_error| {
            Error::InvalidBundle("reassembled ZIP length does not fit usize".to_owned())
        })?;

        let physical_count = archive.physical_entries().count();
        let mut local_offsets = Vec::new();
        local_offsets
            .try_reserve_exact(physical_count)
            .map_err(|_error| Error::Allocation {
                resource: "reassembled ZIP local offsets",
                amount: physical_count,
            })?;

        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_error| Error::Allocation {
                resource: "reassembled ZIP output",
                amount: output_len,
            })?;

        let prelude_end = archive
            .physical_entries()
            .next()
            .map_or(archive.directory_offset(), |entry| {
                entry.local_record().start
            });
        output.extend_from_slice(&self.source[..prelude_end]);

        for (index, physical) in archive.physical_entries().enumerate() {
            if deleted.contains(&index) {
                local_offsets.push(None);
                continue;
            }
            let local_offset = u64::try_from(output.len()).map_err(|_error| {
                Error::InvalidBundle("reassembled local offset does not fit u64".to_owned())
            })?;
            local_offsets.push(Some(local_offset));
            if let Some(edit) = prepared.get(&index) {
                append_edited_local(&mut output, self.source.as_ref(), physical, edit)?;
            } else {
                output.extend_from_slice(&self.source[physical.local_record()]);
            }
        }

        let new_directory_offset = u64::try_from(output.len()).map_err(|_error| {
            Error::InvalidBundle("reassembled central directory offset does not fit u64".to_owned())
        })?;
        for physical_index in archive.physical_indices_in_central_order() {
            if deleted.contains(&physical_index) {
                continue;
            }
            let physical = archive.physical_entry(physical_index).ok_or_else(|| {
                Error::Reassembly("central order references a missing ZIP entry".to_owned())
            })?;
            let local_offset = local_offsets
                .get(physical_index)
                .copied()
                .flatten()
                .ok_or_else(|| {
                    Error::Reassembly(
                        "retained central record has no reassembled local offset".to_owned(),
                    )
                })?;
            let start = output.len();
            output.extend_from_slice(&self.source[physical.central_record()]);
            patch_central_record(
                &mut output[start..],
                local_offset,
                prepared.get(&physical_index),
                shape.base_offset,
            )?;
        }
        let new_directory_size = output
            .len()
            .checked_sub(usize::try_from(new_directory_offset).map_err(|_error| {
                Error::Reassembly("reassembled central offset does not fit usize".to_owned())
            })?)
            .ok_or_else(|| {
                Error::Reassembly("reassembled central directory range is invalid".to_owned())
            })?;
        let retained_count = physical_count.checked_sub(deleted.len()).ok_or_else(|| {
            Error::Reassembly("deleted ZIP entry count exceeds physical entry count".to_owned())
        })?;

        let tail_start = output.len();
        output.extend_from_slice(&self.source[archive.eocd_offset()..]);
        patch_end_of_central_directory(
            &mut output[tail_start..],
            new_directory_offset,
            shape.base_offset,
            new_directory_size,
            retained_count,
        )?;

        if output.len() != output_len {
            return Err(Error::Reassembly(format!(
                "reassembled ZIP size changed during publication (planned {output_len}, wrote {})",
                output.len()
            )));
        }
        Ok(output)
    }

    /// Transactionally prepare an edited package and write its committed ZIP
    /// artifact to a caller-owned sink.
    ///
    /// The complete artifact is materialized and validated before the sink is
    /// touched. A sink can still fail while accepting bytes; callers that need
    /// filesystem atomicity should use their own atomic replacement boundary.
    ///
    /// # Errors
    ///
    /// Returns the same validation, allocation, and sink errors as
    /// [`Self::reassemble_to_bytes`].
    pub fn write_reassembled_to<W: Write>(
        &self,
        edits: &[EntryEdit<'_>],
        mut sink: W,
        limits: Limits,
    ) -> Result<()> {
        if edits.is_empty() {
            let checked_limits = limits.validate()?;
            let source_size = u64::try_from(self.source.len()).map_err(|_error| {
                Error::InvalidBundle("catalog source length does not fit u64".to_owned())
            })?;
            checked_limits.check_input_size(source_size, "catalog source")?;
            checked_limits.check_output_size(source_size)?;
            self.write_to(&mut sink)?;
            return Ok(());
        }
        let bytes = self.reassemble_to_bytes(edits, limits)?;
        sink.write_all(&bytes)?;
        Ok(())
    }
}

impl IntoIterator for Catalog {
    type Item = Entry;
    type IntoIter = std::vec::IntoIter<Entry>;

    fn into_iter(self) -> Self::IntoIter {
        self.entries.into_iter()
    }
}

fn read_source(source: &dyn ReadAt, limits: Limits) -> Result<Arc<[u8]>> {
    let expected = source.version()?;
    let source_length = source.len()?;
    limits.check_input_size(source_length, "ReadAt input")?;
    let length = usize::try_from(source_length).map_err(|_error| {
        Error::InvalidBundle("ReadAt input length does not fit usize".to_owned())
    })?;

    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length)
        .map_err(|_error| Error::Allocation {
            resource: "ReadAt source bytes",
            amount: length,
        })?;
    bytes.resize(length, 0);

    let read_error = source.read_exact_at(0, &mut bytes).err();
    let observed = source.version()?;
    if observed != expected {
        return Err(Error::SourceChanged { expected, observed });
    }
    if let Some(error) = read_error {
        return Err(error.into());
    }

    Ok(bytes.into())
}

fn header_metadata(header: &PhysicalHeader) -> HeaderMetadata {
    HeaderMetadata {
        version_needed: header.version_needed,
        flags: header.flags,
        compression_method: header.compression_method,
        last_modified: DosDateTime {
            time: header.last_mod_time,
            date: header.last_mod_date,
        },
        name: header.name.clone(),
        extra: header.extra.clone(),
        comment: header.comment.clone(),
    }
}

/// Write ordered, uncompressed package members to a physical iWork ZIP.
///
/// The input iterator must be cloneable because the complete member budget is
/// checked before the first byte reaches `sink`. This keeps a rejected
/// package transaction from leaving a partially written physical archive.
/// This is a new logical-package writer: it deliberately emits Store entries
/// and has no physical metadata input. Use `Catalog::write_to` for an untouched
/// preserve-mode round trip.
/// ZIP grammar and implementation details remain private to this crate.
///
/// # Errors
///
/// Returns an error when the entry budget is exceeded or the physical ZIP
/// writer rejects the sink or an entry.
pub fn write_to<'a, I, W>(entries: I, sink: W, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = (&'a str, &'a [u8])> + Clone,
    W: Write,
{
    let checked_limits = limits.validate()?;
    validate_output(entries.clone(), checked_limits)?;

    let mut writer = StreamingArchiveWriter::with_writer(sink);
    for (name, data) in entries {
        writer.write_stored(name, data)?;
    }
    writer.finish()?;
    Ok(())
}

/// Encode ordered, uncompressed package members as a physical iWork ZIP.
///
/// This is a new logical-package writer and intentionally does not preserve
/// metadata from an Entry. Use `Catalog::to_bytes` for an untouched preserve-mode
/// round trip.
///
/// # Errors
///
/// Returns an error when the entry budget is exceeded or the physical ZIP
/// writer rejects an entry.
pub fn to_bytes<'a, I>(entries: I, limits: Limits) -> Result<Vec<u8>>
where
    I: IntoIterator<Item = (&'a str, &'a [u8])> + Clone,
{
    let mut bytes = Vec::new();
    write_to(entries, &mut bytes, limits)?;
    Ok(bytes)
}

fn validate_output<'a, I>(entries: I, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = (&'a str, &'a [u8])>,
{
    let maximum_entries = u64::try_from(limits.max_entries()).map_err(|error| {
        Error::InvalidBundle(format!(
            "package output entry limit does not fit u64: {error}"
        ))
    })?;
    let mut count = 0usize;
    let mut total = 0u64;
    for (_name, data) in entries {
        count = count.checked_add(1).ok_or_else(|| {
            Error::InvalidBundle("package output entry count overflow".to_owned())
        })?;
        if count > limits.max_entries() {
            let observed = u64::try_from(count).map_err(|error| {
                Error::InvalidBundle(format!(
                    "package output entry count does not fit u64: {error}"
                ))
            })?;
            return Err(Error::Limit {
                kind: crate::LimitKind::Entries,
                observed,
                maximum: maximum_entries,
            });
        }
        let size = u64::try_from(data.len()).map_err(|error| {
            Error::InvalidBundle(format!(
                "package output member length does not fit u64: {error}"
            ))
        })?;
        if size > limits.max_entry_bytes() {
            return Err(Error::Limit {
                kind: crate::LimitKind::EntryBytes,
                observed: size,
                maximum: limits.max_entry_bytes(),
            });
        }
        total = total
            .checked_add(size)
            .ok_or_else(|| Error::InvalidBundle("package output total size overflow".to_owned()))?;
        if total > limits.max_total_bytes() {
            return Err(Error::Limit {
                kind: crate::LimitKind::TotalBytes,
                observed: total,
                maximum: limits.max_total_bytes(),
            });
        }
    }
    Ok(())
}

fn validate_reassembly_shape(archive: &ZipArchive<'_>) -> Result<ReassemblyShape> {
    let source = archive.source();
    let directory_offset = archive.directory_offset();
    let eocd_offset = archive.eocd_offset();
    if directory_offset > eocd_offset || eocd_offset > source.len() {
        return Err(Error::Reassembly(
            "ZIP central directory range is invalid".to_owned(),
        ));
    }
    let central_directory_size = eocd_offset - directory_offset;
    let tail = source
        .get(eocd_offset..)
        .ok_or_else(|| Error::Reassembly("ZIP end of central directory is truncated".to_owned()))?;
    if tail.len() < 22 || raw_u32(tail, 0) != Some(0x0605_4b50) {
        return Err(Error::Reassembly(
            "ZIP end of central directory is malformed".to_owned(),
        ));
    }
    let comment_len = usize::from(raw_u16(tail, 20).ok_or_else(|| {
        Error::Reassembly("ZIP end of central directory comment length is truncated".to_owned())
    })?);
    let tail_end = 22usize.checked_add(comment_len).ok_or_else(|| {
        Error::Reassembly("ZIP end of central directory comment length overflows".to_owned())
    })?;
    if tail.len() < tail_end {
        return Err(Error::Reassembly(
            "ZIP archive comment is truncated".to_owned(),
        ));
    }
    if raw_u16(tail, 4) != Some(0)
        || raw_u16(tail, 6) != Some(0)
        || raw_u16(tail, 8) != raw_u16(tail, 10)
        || raw_u32(tail, 12) != u32::try_from(central_directory_size).ok()
    {
        return Err(Error::Reassembly(
            "ZIP uses a multi-disk or ZIP64 central-directory layout".to_owned(),
        ));
    }

    let physical_count = archive.physical_entries().count();
    if physical_count > usize::from(u16::MAX)
        || raw_u16(tail, 10) != u16::try_from(physical_count).ok()
    {
        return Err(Error::Reassembly(
            "ZIP entry count requires ZIP64 or is inconsistent".to_owned(),
        ));
    }
    let raw_directory_offset = raw_u32(tail, 16)
        .ok_or_else(|| Error::Reassembly("ZIP central-directory offset is truncated".to_owned()))?;
    let actual_directory_offset = u64::try_from(directory_offset)
        .map_err(|_error| Error::Reassembly("ZIP central offset does not fit u64".to_owned()))?;
    if raw_directory_offset == u32::MAX
        || archive.base_offset() > actual_directory_offset
        || actual_directory_offset - archive.base_offset() != u64::from(raw_directory_offset)
    {
        return Err(Error::Reassembly(
            "ZIP central-directory offset has an unsupported base".to_owned(),
        ));
    }
    if eocd_offset >= 20 && raw_u32(&source[eocd_offset - 20..eocd_offset], 0) == Some(0x0706_4b50)
    {
        return Err(Error::Reassembly(
            "ZIP64 end-of-central-directory locator is not supported by this slice".to_owned(),
        ));
    }
    if eocd_offset >= 76
        && raw_u32(&source[eocd_offset - 76..eocd_offset - 20], 0) == Some(0x0606_4b50)
    {
        return Err(Error::Reassembly(
            "ZIP64 end-of-central-directory record is not supported by this slice".to_owned(),
        ));
    }

    for physical in archive.physical_entries() {
        let local_range = physical.local_record();
        let central_range = physical.central_record();
        let local = source
            .get(local_range.clone())
            .ok_or_else(|| Error::Reassembly("ZIP local record range is truncated".to_owned()))?;
        let central = source
            .get(central_range.clone())
            .ok_or_else(|| Error::Reassembly("ZIP central record range is truncated".to_owned()))?;
        if local.len() < 30
            || central.len() < 46
            || local_range.end > directory_offset
            || central_range.start < directory_offset
            || central_range.end > eocd_offset
        {
            return Err(Error::Reassembly(
                "ZIP member record crosses a structural boundary".to_owned(),
            ));
        }
        if raw_u32(local, 0) != Some(0x0403_4b50)
            || raw_u32(central, 0) != Some(0x0201_4b50)
            || [18usize, 22]
                .into_iter()
                .any(|offset| raw_u32(local, offset) == Some(u32::MAX))
            || [20usize, 24, 42]
                .into_iter()
                .any(|offset| raw_u32(central, offset) == Some(u32::MAX))
        {
            return Err(Error::Reassembly(
                "ZIP member uses ZIP64 fields or has malformed fixed headers".to_owned(),
            ));
        }

        let central_compressed = raw_u32(central, 20).ok_or_else(|| {
            Error::Reassembly("ZIP central compressed size is truncated".to_owned())
        })?;
        let central_uncompressed = raw_u32(central, 24).ok_or_else(|| {
            Error::Reassembly("ZIP central uncompressed size is truncated".to_owned())
        })?;
        if u64::from(central_compressed) != physical.compressed_size()
            || u64::from(central_uncompressed) != physical.uncompressed_size()
        {
            return Err(Error::Reassembly(
                "ZIP central size fields disagree with the parsed entry".to_owned(),
            ));
        }
        let local_offset = raw_u32(central, 42)
            .ok_or_else(|| Error::Reassembly("ZIP local-header offset is truncated".to_owned()))?;
        let expected_offset = u64::try_from(local_range.start)
            .map_err(|_error| Error::Reassembly("ZIP local offset does not fit u64".to_owned()))?
            .checked_sub(archive.base_offset())
            .ok_or_else(|| Error::Reassembly("ZIP local offset has an invalid base".to_owned()))?;
        if u64::from(local_offset) != expected_offset {
            return Err(Error::Reassembly(
                "ZIP central local-header offset is inconsistent".to_owned(),
            ));
        }
    }

    let mut central_cursor = directory_offset;
    for physical_index in archive.physical_indices_in_central_order() {
        let physical = archive.physical_entry(physical_index).ok_or_else(|| {
            Error::Reassembly("central order references a missing ZIP entry".to_owned())
        })?;
        let record = physical.central_record();
        if record.start != central_cursor {
            return Err(Error::Reassembly(
                "ZIP central directory contains an unsupported non-member record or gap".to_owned(),
            ));
        }
        central_cursor = record.end;
    }
    if central_cursor != eocd_offset {
        return Err(Error::Reassembly(
            "ZIP central directory contains an unsupported trailing record or gap".to_owned(),
        ));
    }

    Ok(ReassemblyShape {
        base_offset: archive.base_offset(),
    })
}

fn prepare_mutations(
    archive: &ZipArchive<'_>,
    edits: &[EntryEdit<'_>],
    deleted_names: &[&str],
    limits: Limits,
) -> Result<(HashMap<usize, PreparedEdit>, HashSet<usize>)> {
    let mut requested_edits = HashMap::new();
    requested_edits
        .try_reserve(edits.len())
        .map_err(|_error| Error::Allocation {
            resource: "reassembly edit index",
            amount: edits.len(),
        })?;
    for &edit in edits {
        if requested_edits.insert(edit.name(), edit).is_some() {
            return Err(Error::Reassembly(format!(
                "member is edited more than once: {}",
                edit.name()
            )));
        }
    }

    let mut requested_deletions = HashSet::new();
    requested_deletions
        .try_reserve(deleted_names.len())
        .map_err(|_error| Error::Allocation {
            resource: "reassembly deletion index",
            amount: deleted_names.len(),
        })?;
    for &name in deleted_names {
        if !requested_deletions.insert(name) {
            return Err(Error::Reassembly(format!(
                "member is deleted more than once: {name}"
            )));
        }
        if requested_edits.contains_key(name) {
            return Err(Error::Reassembly(format!(
                "member cannot be both edited and deleted: {name}"
            )));
        }
    }

    let mut matched = Vec::new();
    matched
        .try_reserve_exact(edits.len())
        .map_err(|_error| Error::Allocation {
            resource: "matched reassembly edits",
            amount: edits.len(),
        })?;
    let mut deleted = HashSet::new();
    deleted
        .try_reserve(deleted_names.len())
        .map_err(|_error| Error::Allocation {
            resource: "matched reassembly deletions",
            amount: deleted_names.len(),
        })?;
    let mut matched_names = HashSet::new();
    let selected_count = edits
        .len()
        .checked_add(deleted_names.len())
        .ok_or_else(|| Error::Reassembly("mutation selection count overflows usize".to_owned()))?;
    matched_names
        .try_reserve(selected_count)
        .map_err(|_error| Error::Allocation {
            resource: "matched reassembly member names",
            amount: selected_count,
        })?;
    let mut total_uncompressed = 0u64;
    for physical in archive.physical_entries() {
        total_uncompressed = total_uncompressed
            .checked_add(physical.uncompressed_size())
            .ok_or_else(|| Error::Reassembly("ZIP uncompressed size overflows u64".to_owned()))?;
    }

    for (index, physical) in archive.physical_entries().enumerate() {
        if physical.is_directory() {
            continue;
        }
        let name = physical.name();
        let selected_for_edit = requested_edits.get(name).copied();
        let selected_for_deletion = requested_deletions.contains(name);
        if selected_for_edit.is_none() && !selected_for_deletion {
            continue;
        }
        if !matched_names.insert(name) {
            return Err(Error::Reassembly(format!(
                "normalized member name selects more than one physical entry: {name}"
            )));
        }
        if selected_for_deletion {
            total_uncompressed = total_uncompressed
                .checked_sub(physical.uncompressed_size())
                .ok_or_else(|| {
                    Error::Reassembly("ZIP uncompressed size underflows during deletion".to_owned())
                })?;
            deleted.insert(index);
            continue;
        }

        let edit = selected_for_edit.ok_or_else(|| {
            Error::Reassembly(format!("selected member has no edit payload: {name}"))
        })?;
        let data_size = u64::try_from(edit.data().len()).map_err(|_error| {
            Error::InvalidBundle("edited member length does not fit u64".to_owned())
        })?;
        if data_size > limits.max_entry_bytes() {
            return Err(Error::Limit {
                kind: crate::LimitKind::EntryBytes,
                observed: data_size,
                maximum: limits.max_entry_bytes(),
            });
        }
        if physical.local_header().flags != physical.central_header().flags
            || physical.local_header().compression_method
                != physical.central_header().compression_method
        {
            return Err(Error::Reassembly(format!(
                "edited member has inconsistent local and central metadata: {}",
                physical.name()
            )));
        }
        let method = physical.central_header().compression_method;
        if !matches!(method, 0 | 8) {
            return Err(Error::Reassembly(format!(
                "edited member uses unsupported compression method {method}: {}",
                physical.name()
            )));
        }
        let compressed_range = physical.compressed_data_range();
        let local_range = physical.local_record();
        let suffix = &archive.source()[compressed_range.end..local_range.end];
        let descriptor_start = descriptor_start(physical, suffix)?;
        total_uncompressed = total_uncompressed
            .checked_sub(physical.uncompressed_size())
            .and_then(|value| value.checked_add(data_size))
            .ok_or_else(|| Error::Reassembly("ZIP uncompressed size overflows u64".to_owned()))?;
        matched.push(MatchedEdit {
            index,
            edit,
            descriptor_start,
        });
    }
    for edit in edits {
        if !matched_names.contains(edit.name()) {
            return Err(Error::Reassembly(format!(
                "edited member does not exist in the flat catalog: {}",
                edit.name()
            )));
        }
    }
    for &name in deleted_names {
        if !matched_names.contains(name) {
            return Err(Error::Reassembly(format!(
                "deleted member does not exist in the flat catalog: {name}"
            )));
        }
    }
    if total_uncompressed > limits.max_total_bytes() {
        return Err(Error::Limit {
            kind: crate::LimitKind::TotalBytes,
            observed: total_uncompressed,
            maximum: limits.max_total_bytes(),
        });
    }

    let mut prepared = HashMap::new();
    prepared
        .try_reserve(matched.len())
        .map_err(|_error| Error::Allocation {
            resource: "prepared reassembly edits",
            amount: matched.len(),
        })?;
    for matched_edit in matched {
        let physical = archive.physical_entry(matched_edit.index).ok_or_else(|| {
            Error::Reassembly("matched edit references a missing ZIP entry".to_owned())
        })?;
        let data = matched_edit.edit.data();
        let compressed =
            encode_replacement(physical.central_header().compression_method, data, limits)?;
        let compressed_size = u64::try_from(compressed.len()).map_err(|_error| {
            Error::InvalidBundle("edited compressed member length does not fit u64".to_owned())
        })?;
        if compressed_size > u64::from(u32::MAX) {
            return Err(Error::Reassembly(format!(
                "edited member is too large for non-ZIP64 reassembly: {}",
                physical.name()
            )));
        }
        prepared.insert(
            matched_edit.index,
            PreparedEdit {
                compressed,
                crc32: soapberry_zip::crc32(data),
                uncompressed_size: u64::try_from(data.len()).map_err(|_error| {
                    Error::InvalidBundle("edited member length does not fit u64".to_owned())
                })?,
                descriptor_start: matched_edit.descriptor_start,
            },
        );
    }
    Ok((prepared, deleted))
}

fn descriptor_start(physical: &PhysicalEntry, suffix: &[u8]) -> Result<Option<usize>> {
    if physical.central_header().flags & 0x0008 == 0 {
        return Ok(None);
    }
    let signed = raw_u32(suffix, 0) == Some(0x0807_4b50);
    let start = if signed { 4 } else { 0 };
    let Some(crc32) = raw_u32(suffix, start) else {
        return Err(Error::Reassembly(format!(
            "edited member has a truncated data descriptor: {}",
            physical.name()
        )));
    };
    let Some(compressed_size) = raw_u32(suffix, start + 4) else {
        return Err(Error::Reassembly(format!(
            "edited member has a truncated data descriptor size: {}",
            physical.name()
        )));
    };
    let Some(uncompressed_size) = raw_u32(suffix, start + 8) else {
        return Err(Error::Reassembly(format!(
            "edited member has a truncated data descriptor length: {}",
            physical.name()
        )));
    };
    if crc32 != physical.crc32()
        || u64::from(compressed_size) != physical.compressed_size()
        || u64::from(uncompressed_size) != physical.uncompressed_size()
    {
        return Err(Error::Reassembly(format!(
            "edited member data descriptor disagrees with its central record: {}",
            physical.name()
        )));
    }
    Ok(Some(start))
}

fn encode_replacement(method: u16, data: &[u8], limits: Limits) -> Result<Vec<u8>> {
    match method {
        0 => {
            let mut compressed = Vec::new();
            compressed
                .try_reserve_exact(data.len())
                .map_err(|_error| Error::Allocation {
                    resource: "edited stored member",
                    amount: data.len(),
                })?;
            compressed.extend_from_slice(data);
            Ok(compressed)
        },
        8 => {
            let max = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
                Error::InvalidBundle("reassembly output limit does not fit usize".to_owned())
            })?;
            let bounded = BoundedBuffer::new(max);
            let mut encoder = DeflateEncoder::new(bounded, Compression::default());
            encoder.write_all(data)?;
            Ok(encoder.finish()?.into_inner())
        },
        other => Err(Error::Reassembly(format!(
            "compression method {other} cannot be edited"
        ))),
    }
}

fn reassembled_output_size(
    archive: &ZipArchive<'_>,
    prepared: &HashMap<usize, PreparedEdit>,
    deleted: &HashSet<usize>,
) -> Result<u64> {
    let source = archive.source();
    let prelude_end = archive
        .physical_entries()
        .next()
        .map_or(archive.directory_offset(), |entry| {
            entry.local_record().start
        });
    let mut size = u64::try_from(prelude_end)
        .map_err(|_error| Error::Reassembly("ZIP prelude length does not fit u64".to_owned()))?;
    for (index, physical) in archive.physical_entries().enumerate() {
        if deleted.contains(&index) {
            continue;
        }
        let local_len = if let Some(edit) = prepared.get(&index) {
            let header_len = physical
                .compressed_data_range()
                .start
                .checked_sub(physical.local_record().start)
                .ok_or_else(|| Error::Reassembly("ZIP local header range is invalid".to_owned()))?;
            let suffix_len = physical
                .local_record()
                .end
                .checked_sub(physical.compressed_data_range().end)
                .ok_or_else(|| Error::Reassembly("ZIP descriptor range is invalid".to_owned()))?;
            header_len
                .checked_add(edit.compressed.len())
                .and_then(|value| value.checked_add(suffix_len))
                .ok_or_else(|| {
                    Error::Reassembly("reassembled local record overflows usize".to_owned())
                })?
        } else {
            source[physical.local_record()].len()
        };
        size = size
            .checked_add(u64::try_from(local_len).map_err(|_error| {
                Error::Reassembly("reassembled local record length does not fit u64".to_owned())
            })?)
            .ok_or_else(|| Error::Reassembly("reassembled ZIP length overflows u64".to_owned()))?;
    }
    let mut central_size = 0usize;
    for physical_index in archive.physical_indices_in_central_order() {
        if deleted.contains(&physical_index) {
            continue;
        }
        let physical = archive.physical_entry(physical_index).ok_or_else(|| {
            Error::Reassembly("central order references a missing ZIP entry".to_owned())
        })?;
        central_size = central_size
            .checked_add(physical.central_record().len())
            .ok_or_else(|| {
                Error::Reassembly("reassembled central directory length overflows usize".to_owned())
            })?;
    }
    let tail_size = source
        .len()
        .checked_sub(archive.eocd_offset())
        .ok_or_else(|| Error::Reassembly("ZIP tail range is invalid".to_owned()))?;
    size = size
        .checked_add(u64::try_from(central_size).map_err(|_error| {
            Error::Reassembly("central directory length does not fit u64".to_owned())
        })?)
        .and_then(|value| value.checked_add(u64::try_from(tail_size).ok()?))
        .ok_or_else(|| Error::Reassembly("reassembled ZIP length overflows u64".to_owned()))?;
    Ok(size)
}

fn append_edited_local(
    output: &mut Vec<u8>,
    source: &[u8],
    physical: &PhysicalEntry,
    edit: &PreparedEdit,
) -> Result<()> {
    let local = physical.local_record();
    let compressed = physical.compressed_data_range();
    let header_len = compressed
        .start
        .checked_sub(local.start)
        .ok_or_else(|| Error::Reassembly("ZIP local header range is invalid".to_owned()))?;
    let header_start = output.len();
    output.extend_from_slice(&source[local.start..compressed.start]);
    if edit.descriptor_start.is_none() {
        patch_u32_at(output, header_start + 14, edit.crc32, "local CRC-32")?;
        patch_u32_at(
            output,
            header_start + 18,
            u32::try_from(edit.compressed.len()).map_err(|_error| {
                Error::Reassembly("edited compressed size does not fit u32".to_owned())
            })?,
            "local compressed size",
        )?;
        patch_u32_at(
            output,
            header_start + 22,
            u32::try_from(edit.uncompressed_size).map_err(|_error| {
                Error::Reassembly("edited uncompressed size does not fit u32".to_owned())
            })?,
            "local uncompressed size",
        )?;
    }
    output.extend_from_slice(&edit.compressed);
    let suffix_start = output.len();
    output.extend_from_slice(&source[compressed.end..local.end]);
    if let Some(descriptor_start) = edit.descriptor_start {
        let descriptor = suffix_start.checked_add(descriptor_start).ok_or_else(|| {
            Error::Reassembly("data descriptor offset overflows usize".to_owned())
        })?;
        patch_u32_at(output, descriptor, edit.crc32, "descriptor CRC-32")?;
        patch_u32_at(
            output,
            descriptor + 4,
            u32::try_from(edit.compressed.len()).map_err(|_error| {
                Error::Reassembly("edited descriptor compressed size does not fit u32".to_owned())
            })?,
            "descriptor compressed size",
        )?;
        patch_u32_at(
            output,
            descriptor + 8,
            u32::try_from(edit.uncompressed_size).map_err(|_error| {
                Error::Reassembly("edited descriptor uncompressed size does not fit u32".to_owned())
            })?,
            "descriptor uncompressed size",
        )?;
    }
    debug_assert_eq!(
        header_len + edit.compressed.len() + (local.end - compressed.end),
        output.len() - header_start
    );
    Ok(())
}

fn patch_central_record(
    record: &mut [u8],
    local_offset: u64,
    prepared_edit: Option<&PreparedEdit>,
    base_offset: u64,
) -> Result<()> {
    let relative_offset = local_offset.checked_sub(base_offset).ok_or_else(|| {
        Error::Reassembly("reassembled local offset has an invalid base".to_owned())
    })?;
    patch_u32_at(
        record,
        42,
        u32::try_from(relative_offset).map_err(|_error| {
            Error::Reassembly("reassembled local offset does not fit u32".to_owned())
        })?,
        "central local-header offset",
    )?;
    if let Some(edit) = prepared_edit {
        patch_u32_at(record, 16, edit.crc32, "central CRC-32")?;
        patch_u32_at(
            record,
            20,
            u32::try_from(edit.compressed.len()).map_err(|_error| {
                Error::Reassembly("central compressed size does not fit u32".to_owned())
            })?,
            "central compressed size",
        )?;
        patch_u32_at(
            record,
            24,
            u32::try_from(edit.uncompressed_size).map_err(|_error| {
                Error::Reassembly("central uncompressed size does not fit u32".to_owned())
            })?,
            "central uncompressed size",
        )?;
    }
    Ok(())
}

fn patch_end_of_central_directory(
    tail: &mut [u8],
    directory_offset: u64,
    base_offset: u64,
    central_directory_size: usize,
    entry_count: usize,
) -> Result<()> {
    let relative_offset = directory_offset.checked_sub(base_offset).ok_or_else(|| {
        Error::Reassembly("reassembled central offset has an invalid base".to_owned())
    })?;
    let count = u16::try_from(entry_count).map_err(|_error| {
        Error::Reassembly("reassembled ZIP entry count requires ZIP64".to_owned())
    })?;
    patch_u16_at(tail, 8, count, "end-of-central-directory disk entry count")?;
    patch_u16_at(tail, 10, count, "end-of-central-directory entry count")?;
    patch_u32_at(
        tail,
        12,
        u32::try_from(central_directory_size).map_err(|_error| {
            Error::Reassembly("reassembled central directory size requires ZIP64".to_owned())
        })?,
        "end-of-central-directory central size",
    )?;
    patch_u32_at(
        tail,
        16,
        u32::try_from(relative_offset).map_err(|_error| {
            Error::Reassembly("reassembled central directory offset does not fit u32".to_owned())
        })?,
        "end-of-central-directory offset",
    )?;
    Ok(())
}

fn patch_u16_at(bytes: &mut [u8], start: usize, value: u16, label: &str) -> Result<()> {
    let end = start
        .checked_add(2)
        .ok_or_else(|| Error::Reassembly(format!("{label} offset overflows usize")))?;
    let target = bytes
        .get_mut(start..end)
        .ok_or_else(|| Error::Reassembly(format!("{label} is truncated")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn patch_u32_at(bytes: &mut [u8], start: usize, value: u32, label: &str) -> Result<()> {
    let end = start
        .checked_add(4)
        .ok_or_else(|| Error::Reassembly(format!("{label} offset overflows usize")))?;
    let target = bytes
        .get_mut(start..end)
        .ok_or_else(|| Error::Reassembly(format!("{label} is truncated")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn raw_u16(bytes: &[u8], start: usize) -> Option<u16> {
    bytes
        .get(start..start.checked_add(2)?)
        .map(|value| u16::from_le_bytes([value[0], value[1]]))
}

fn raw_u32(bytes: &[u8], start: usize) -> Option<u32> {
    bytes
        .get(start..start.checked_add(4)?)
        .map(|value| u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn diagnostic_fingerprint(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325_u64, |value, byte| {
        (value ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
    })
}

fn preflight_flat_logical_entries(
    archive: &ZipArchive<'_>,
    limits: LogicalEntryLimits,
) -> Result<()> {
    for physical in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        preflight_logical_entry(physical, physical.raw_name(), limits)?;
    }
    Ok(())
}

pub(crate) fn semantic_nested_index_name(archive: &ZipArchive<'_>) -> Result<Option<String>> {
    Ok(semantic_nested_index_entry(archive)?.map(|entry| entry.name().to_owned()))
}

pub(crate) fn semantic_nested_index_entry<'archive, 'data>(
    archive: &'archive ZipArchive<'data>,
) -> Result<Option<&'archive PhysicalEntry>> {
    let mut selected: Option<&PhysicalEntry> = None;
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let central_exact = is_exact_semantic_nested_index(entry.raw_name());
        let local_exact = is_exact_semantic_nested_index(&entry.local_header().name);
        let central_alias = !central_exact
            && (entry.name().rsplit('/').next() == Some("Index.zip")
                || raw_has_index_basename(entry.raw_name()));
        let local_collision =
            !central_exact && (local_exact || raw_has_index_basename(&entry.local_header().name));
        if central_alias || local_collision {
            return Err(Error::InvalidBundle(format!(
                "semantic package contains a non-canonical raw Index.zip authority: {:?}",
                String::from_utf8_lossy(entry.raw_name())
            )));
        }
        if !central_exact {
            continue;
        }
        if let Some(first) = selected {
            return Err(Error::InvalidBundle(format!(
                "iWork package contains ambiguous exact nested indexes: {} and {}",
                first.name(),
                entry.name()
            )));
        }
        selected = Some(entry);
    }
    Ok(selected)
}

fn raw_has_index_basename(raw_name: &[u8]) -> bool {
    raw_name
        .split(|byte| matches!(*byte, b'/' | b'\\'))
        .next_back()
        == Some(b"Index.zip".as_slice())
}

fn is_exact_semantic_nested_index(raw_name: &[u8]) -> bool {
    is_exact_portable_raw_name(raw_name)
        && (raw_name == b"Index.zip" || raw_name.ends_with(b"/Index.zip"))
}

pub(crate) fn semantic_detection_root_entry<'archive, 'data>(
    archive: &'archive ZipArchive<'data>,
    nested: bool,
) -> Result<Option<&'archive PhysicalEntry>> {
    let authority = if nested {
        b"Document.iwa".as_slice()
    } else {
        b"Index/Document.iwa".as_slice()
    };
    let normalized_authority = if nested {
        "Document.iwa"
    } else {
        "Index/Document.iwa"
    };
    let mut selected: Option<&PhysicalEntry> = None;
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let central_exact = entry.raw_name() == authority;
        let local_exact = entry.local_header().name.as_ref() == authority;
        let central_alias = entry.raw_name() != authority
            && (entry.name() == normalized_authority
                || raw_path_normalizes_to(entry.raw_name(), authority));
        let local_collision = !central_exact
            && (local_exact || raw_path_normalizes_to(&entry.local_header().name, authority));
        if central_alias || local_collision {
            return Err(Error::InvalidBundle(format!(
                "semantic package contains a non-canonical raw Document.iwa authority: {:?}",
                String::from_utf8_lossy(entry.raw_name())
            )));
        }
        if !central_exact {
            continue;
        }
        if selected.is_some() {
            return Err(Error::InvalidBundle(
                "iWork package contains multiple exact Document.iwa components".to_owned(),
            ));
        }
        selected = Some(entry);
    }
    Ok(selected)
}

pub(crate) fn raw_path_normalizes_to(raw_name: &[u8], authority: &[u8]) -> bool {
    let mut expected = authority.split(|byte| *byte == b'/');
    let first = expected.next();
    let second = expected.next();
    let expected_depth = usize::from(first.is_some()) + usize::from(second.is_some());
    if expected.next().is_some() {
        return false;
    }

    let mut retained = [None, None];
    let mut depth = 0_usize;
    for component in raw_name.split(|byte| matches!(*byte, b'/' | b'\\')) {
        if component.is_empty() || component == b"." {
            continue;
        }
        if component == b".." {
            if depth != 0 {
                depth -= 1;
                if depth < retained.len() {
                    retained[depth] = None;
                }
            }
            continue;
        }
        if depth < retained.len() {
            retained[depth] = Some(component);
        }
        depth = depth.saturating_add(1);
    }
    depth == expected_depth && retained[0] == first && retained[1] == second
}

pub(crate) fn reject_semantic_aliases(archive: &ZipArchive<'_>) -> Result<()> {
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let normalized_iwa = crate::zip::is_iwa_name(entry.name())
            && (entry.name().starts_with("Index/") || !entry.name().contains('/'));
        let normalized_index = entry.name().rsplit('/').next() == Some("Index.zip");
        if (normalized_iwa && semantic_iwa_name(entry).is_none())
            || (normalized_index
                && !(is_exact_portable_raw_name(entry.raw_name())
                    && (entry.raw_name() == b"Index.zip"
                        || entry.raw_name().ends_with(b"/Index.zip"))))
        {
            return Err(Error::InvalidBundle(format!(
                "semantic package contains a non-canonical raw authority: {:?}",
                String::from_utf8_lossy(entry.raw_name())
            )));
        }
    }
    Ok(())
}

pub(crate) fn semantic_iwa_name(physical: &PhysicalEntry) -> Option<&str> {
    semantic_iwa_name_in(physical, false)
}

pub(crate) fn semantic_nested_iwa_name(physical: &PhysicalEntry) -> Option<&str> {
    semantic_iwa_name_in(physical, true)
}

fn semantic_iwa_name_in(physical: &PhysicalEntry, allow_root: bool) -> Option<&str> {
    let raw_name = physical.raw_name();
    if !is_exact_portable_raw_name(raw_name)
        || !(raw_name.starts_with(b"Index/") || allow_root && !raw_name.contains(&b'/'))
        || !raw_name.ends_with(b".iwa")
    {
        return None;
    }
    std::str::from_utf8(raw_name).ok()
}

fn is_exact_portable_raw_name(raw_name: &[u8]) -> bool {
    let Ok(name) = std::str::from_utf8(raw_name) else {
        return false;
    };
    !name.is_empty()
        && !name.starts_with('/')
        && !name.ends_with('/')
        && !name.contains(['\0', '\\'])
        && !name.chars().any(char::is_control)
        && !name.split('/').any(|component| {
            component.is_empty() || matches!(component, "." | "..") || component.contains(':')
        })
}

pub(crate) fn preflight_semantic_iwa(physical: &PhysicalEntry, name: &str) -> Result<()> {
    preflight_semantic_supported_entry(physical, "canonical IWA entry", name)
}

pub(crate) fn preflight_semantic_container(physical: &PhysicalEntry, name: &str) -> Result<()> {
    preflight_semantic_supported_entry(physical, "canonical container", name)
}

pub(crate) fn preflight_semantic_iwa_entries(archive: &ZipArchive<'_>, nested: bool) -> Result<()> {
    let mut seen = HashSet::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let name = if nested {
            semantic_nested_iwa_name(entry)
        } else {
            semantic_iwa_name(entry)
        };
        let Some(name) = name else {
            if nested && !is_semantic_irrelevant_payload(entry.raw_name()) {
                return Err(Error::InvalidBundle(format!(
                    "legacy package index contains a non-canonical IWA member: {}",
                    entry.name()
                )));
            }
            continue;
        };
        preflight_semantic_iwa(entry, name)?;
        seen.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "semantic IWA authority names",
            amount: 1,
        })?;
        if !seen.insert(name) {
            return Err(Error::InvalidBundle(format!(
                "duplicate semantic IWA authority is ambiguous: {name}"
            )));
        }
    }
    Ok(())
}

fn preflight_semantic_supported_entry(
    physical: &PhysicalEntry,
    label: &str,
    name: &str,
) -> Result<()> {
    let local = physical.local_header();
    let central = physical.central_header();
    if local.name.as_ref() != central.name.as_ref() {
        return Err(Error::InvalidBundle(format!(
            "{label} {name} has mismatched local and central names"
        )));
    }
    if local.compression_method != central.compression_method {
        return Err(Error::InvalidBundle(format!(
            "{label} {name} has mismatched local and central compression methods"
        )));
    }
    if !matches!(central.compression_method, 0 | 8) {
        return Err(Error::InvalidBundle(format!(
            "{label} {name} uses unsupported ZIP compression"
        )));
    }
    Ok(())
}

fn preflight_legacy_outer_logical_entries(
    archive: &ZipArchive<'_>,
    index_name: &str,
    limits: LogicalEntryLimits,
) -> Result<()> {
    let raw_prefix = legacy_raw_outer_prefix(archive, index_name)?;
    for physical in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory() && entry.name() != index_name)
    {
        let raw_name = physical.raw_name();
        if let Some(logical_name) = raw_name.strip_prefix(raw_prefix) {
            preflight_logical_entry(physical, logical_name, limits)?;
        }
    }
    Ok(())
}

fn legacy_raw_outer_prefix<'a>(archive: &'a ZipArchive<'_>, index_name: &str) -> Result<&'a [u8]> {
    let raw_name = archive
        .physical_entries()
        .find(|entry| entry.name() == index_name)
        .ok_or_else(|| {
            Error::InvalidBundle(format!(
                "legacy package index has no physical member: {index_name}"
            ))
        })?
        .raw_name();
    std::str::from_utf8(raw_name).map_err(|_error| {
        Error::InvalidBundle("legacy package index raw name is not UTF-8".to_owned())
    })?;
    raw_name.strip_suffix(b"Index.zip").ok_or_else(|| {
        Error::InvalidBundle("legacy package index raw name has an invalid suffix".to_owned())
    })
}

fn preflight_logical_entry(
    physical: &PhysicalEntry,
    logical_name: &[u8],
    limits: LogicalEntryLimits,
) -> Result<()> {
    let Some(maximum) = limits.maximum_for(logical_name) else {
        return Ok(());
    };
    let local = physical.local_header();
    let central = physical.central_header();
    if local.name.as_ref() != central.name.as_ref() {
        return Err(Error::InvalidBundle(format!(
            "canonical logical entry {} has mismatched local and central names",
            String::from_utf8_lossy(logical_name)
        )));
    }
    if local.compression_method != central.compression_method {
        return Err(Error::InvalidBundle(format!(
            "canonical logical entry {} has mismatched local and central compression methods",
            String::from_utf8_lossy(logical_name)
        )));
    }
    if !matches!(central.compression_method, 0 | 8) {
        return Err(Error::InvalidBundle(format!(
            "canonical logical entry {} uses unsupported ZIP compression",
            String::from_utf8_lossy(logical_name)
        )));
    }
    if physical.uncompressed_size() > maximum {
        return Err(Error::Limit {
            kind: crate::LimitKind::EntryBytes,
            observed: physical.uncompressed_size(),
            maximum,
        });
    }
    Ok(())
}

fn collect_flat(archive: &ZipArchive<'_>, source: &SharedBytes) -> Result<Vec<Entry>> {
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        push_entry(
            archive,
            source.clone(),
            entry,
            entry.name(),
            &mut entries,
            &mut seen,
        )?;
    }
    Ok(entries)
}

fn collect_semantic_flat(
    archive: &ZipArchive<'_>,
    source: &SharedBytes,
    include_metadata: bool,
) -> Result<Vec<Entry>> {
    preflight_semantic_iwa_entries(archive, false)?;
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let semantic_iwa = semantic_iwa_name(entry);
        if semantic_iwa.is_none()
            && !(include_metadata && is_semantic_metadata_authority(entry.raw_name()))
        {
            continue;
        }
        let normalized_name = if let Some(name) = semantic_iwa {
            name
        } else {
            std::str::from_utf8(entry.raw_name()).map_err(|_error| {
                Error::InvalidBundle("semantic metadata authority is not UTF-8".to_owned())
            })?
        };
        push_entry(
            archive,
            source.clone(),
            entry,
            normalized_name,
            &mut entries,
            &mut seen,
        )?;
    }
    Ok(entries)
}

fn collect_legacy(
    archive: &ZipArchive<'_>,
    index_name: &str,
    limits: Limits,
    source: &SharedBytes,
) -> Result<Vec<Entry>> {
    let prefix = index_name.strip_suffix("Index.zip").ok_or_else(|| {
        Error::InvalidBundle(format!("invalid legacy package index name: {index_name}"))
    })?;
    let declared_index_size = archive
        .physical_entries()
        .find(|entry| entry.name() == index_name)
        .ok_or_else(|| {
            Error::InvalidBundle(format!(
                "legacy package index has no physical member: {index_name}"
            ))
        })?
        .uncompressed_size();
    limits.check_input_size(declared_index_size, "legacy iWork Index.zip")?;
    let index_data = archive.read(index_name)?;
    let index_size = u64::try_from(index_data.len()).map_err(|error| {
        Error::InvalidBundle(format!(
            "legacy iWork Index.zip length does not fit u64: {error}"
        ))
    })?;
    limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index_source = SharedBytes::from_owned_vec(index_data);
    let index = ZipArchive::new_with_limits(index_source.as_ref(), limits).map_err(|error| {
        if matches!(&error, Error::Limit { .. }) {
            error
        } else {
            Error::InvalidBundle(format!("legacy package index: {error}"))
        }
    })?;

    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in index
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        if !crate::zip::is_iwa_name(entry.name()) {
            return Err(Error::InvalidBundle(format!(
                "legacy package index contains a non-IWA member: {}",
                entry.name()
            )));
        }
        push_entry(
            &index,
            index_source.clone(),
            entry,
            entry.name(),
            &mut entries,
            &mut seen,
        )?;
    }
    if entries.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }

    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        if entry.name() == index_name {
            continue;
        }
        let normalized = entry.name().strip_prefix(prefix).unwrap_or(entry.name());
        push_entry(
            archive,
            source.clone(),
            entry,
            normalized,
            &mut entries,
            &mut seen,
        )?;
    }
    Ok(entries)
}

fn collect_semantic_legacy(
    archive: &ZipArchive<'_>,
    index_name: &str,
    limits: Limits,
    source: &SharedBytes,
    raw_prefix: &[u8],
    include_metadata: bool,
) -> Result<Vec<Entry>> {
    let index_entry = archive
        .physical_entries()
        .find(|entry| entry.name() == index_name)
        .ok_or_else(|| {
            Error::InvalidBundle(format!(
                "legacy package index has no physical member: {index_name}"
            ))
        })?;
    preflight_semantic_container(index_entry, index_name)?;
    let declared_index_size = index_entry.uncompressed_size();
    limits.check_input_size(declared_index_size, "legacy iWork Index.zip")?;
    let index_data = archive.read_entry(index_entry)?;
    let index_size = u64::try_from(index_data.len()).map_err(|error| {
        Error::InvalidBundle(format!(
            "legacy iWork Index.zip length does not fit u64: {error}"
        ))
    })?;
    limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index_source = SharedBytes::from_owned_vec(index_data);
    let index = ZipArchive::new_with_limits(index_source.as_ref(), limits).map_err(|error| {
        if matches!(&error, Error::Limit { .. }) {
            error
        } else {
            Error::InvalidBundle(format!("legacy package index: {error}"))
        }
    })?;

    preflight_semantic_iwa_entries(&index, true)?;

    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in index
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let Some(name) = semantic_nested_iwa_name(entry) else {
            continue;
        };
        push_entry(
            &index,
            index_source.clone(),
            entry,
            name,
            &mut entries,
            &mut seen,
        )?;
    }
    if entries.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }

    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory() && entry.name() != index_name)
    {
        if !include_metadata {
            continue;
        }
        let Some(logical_raw_name) = entry.raw_name().strip_prefix(raw_prefix) else {
            continue;
        };
        if !is_semantic_metadata_authority(logical_raw_name) {
            continue;
        }
        let normalized_name = std::str::from_utf8(logical_raw_name).map_err(|_error| {
            Error::InvalidBundle("semantic metadata authority is not UTF-8".to_owned())
        })?;
        push_entry(
            archive,
            source.clone(),
            entry,
            normalized_name,
            &mut entries,
            &mut seen,
        )?;
    }
    Ok(entries)
}

const fn is_semantic_metadata_authority(raw_name: &[u8]) -> bool {
    matches!(
        raw_name,
        b"Metadata/Properties.plist"
            | b"Metadata/BuildVersionHistory.plist"
            | b"Metadata/DocumentIdentifier"
    )
}

fn is_semantic_irrelevant_payload(raw_name: &[u8]) -> bool {
    is_exact_portable_raw_name(raw_name)
        && (raw_name.starts_with(b"Data/") || raw_name.starts_with(b"Preview/"))
}

fn push_entry(
    archive: &ZipArchive<'_>,
    source: SharedBytes,
    physical: &PhysicalEntry,
    normalized_name: &str,
    entries: &mut Vec<Entry>,
    seen: &mut HashSet<String>,
) -> Result<()> {
    seen.try_reserve(1).map_err(|_error| Error::Allocation {
        resource: "package entry names",
        amount: 1,
    })?;
    if !seen.insert(normalized_name.to_owned()) {
        return Err(Error::InvalidBundle(format!(
            "duplicate package entry is ambiguous: {normalized_name}"
        )));
    }
    entries.try_reserve(1).map_err(|_error| Error::Allocation {
        resource: "package entries",
        amount: 1,
    })?;
    let raw_record = RawEntryRecord::new(source, physical);
    let metadata = EntryMetadata::new(physical);
    let data = if physical.is_supported() {
        archive.read_entry(physical)?
    } else {
        physical.compressed_data(archive.source()).to_vec()
    };
    entries.push(Entry::new(
        normalized_name,
        data,
        physical.raw_name().to_vec().into_boxed_slice(),
        metadata,
        raw_record,
        !physical.is_supported(),
    ));
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::SourceVersion;
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    #[derive(Debug)]
    struct TestSource {
        bytes: Arc<[u8]>,
        max_read: usize,
        reads: AtomicUsize,
        revision: AtomicU64,
        change_on_read: bool,
    }

    impl TestSource {
        fn new(bytes: Vec<u8>, max_read: usize, change_on_read: bool) -> Self {
            Self {
                bytes: bytes.into(),
                max_read,
                reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
                change_on_read,
            }
        }
    }

    impl ReadAt for TestSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_error| io::Error::other("test source length overflow"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::Relaxed);
            if self.change_on_read {
                self.revision.store(1, Ordering::Relaxed);
            }
            let start = usize::try_from(offset)
                .map_err(|_error| io::Error::other("test source offset overflow"))?;
            let Some(input) = self.bytes.get(start..) else {
                return Ok(0);
            };
            let count = input.len().min(output.len()).min(self.max_read);
            output[..count].copy_from_slice(&input[..count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                41,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }

    fn zip(entries: &[(&str, &[u8])]) -> Result<Vec<u8>> {
        let mut writer = StreamingArchiveWriter::new();
        for (name, data) in entries {
            writer.write_stored(name, data)?;
        }
        Ok(writer.finish_to_bytes()?)
    }

    fn semantic_iwa(identifier: u64) -> Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 6_000,
                    data: vec![1, 2, 3],
                }],
            )?],
        };
        Ok(SnappyStream::compress(&archive.to_bytes()?)?)
    }

    fn push_u16(bytes: &mut Vec<u8>, value: u16) {
        bytes.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(bytes: &mut Vec<u8>, value: u32) {
        bytes.extend_from_slice(&value.to_le_bytes());
    }

    fn checked_u16(value: usize) -> u16 {
        match u16::try_from(value) {
            Ok(converted) => converted,
            Err(error) => panic!("test ZIP field does not fit u16: {error}"),
        }
    }

    fn checked_u32(value: usize) -> u32 {
        match u32::try_from(value) {
            Ok(converted) => converted,
            Err(error) => panic!("test ZIP field does not fit u32: {error}"),
        }
    }

    fn physical_zip(method: u16) -> (Vec<u8>, usize, usize) {
        let name = b"Opaque/entry.bin";
        let data = b"opaque payload";
        let local_extra = b"\xaa\xbb\x03\0xyz";
        let central_extra = b"\xcc\xdd\x02\0\xfe\xed";
        let file_comment = b"entry-comment\0\xff";
        let archive_comment = b"archive-comment\0\xfe";
        let crc32 = soapberry_zip::crc32(data);
        let mut bytes = Vec::new();

        push_u32(&mut bytes, 0x0403_4b50);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 0x0800);
        push_u16(&mut bytes, method);
        push_u16(&mut bytes, 0x1234);
        push_u16(&mut bytes, 0x5678);
        push_u32(&mut bytes, crc32);
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u16(&mut bytes, checked_u16(name.len()));
        push_u16(&mut bytes, checked_u16(local_extra.len()));
        bytes.extend_from_slice(name);
        bytes.extend_from_slice(local_extra);
        bytes.extend_from_slice(data);

        let central_offset = bytes.len();
        push_u32(&mut bytes, 0x0201_4b50);
        push_u16(&mut bytes, 0x0314);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 0x0800);
        push_u16(&mut bytes, method);
        push_u16(&mut bytes, 0x9abc);
        push_u16(&mut bytes, 0xdef0);
        push_u32(&mut bytes, crc32);
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u16(&mut bytes, checked_u16(name.len()));
        push_u16(&mut bytes, checked_u16(central_extra.len()));
        push_u16(&mut bytes, checked_u16(file_comment.len()));
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u32(&mut bytes, 0);
        push_u32(&mut bytes, 0);
        bytes.extend_from_slice(name);
        bytes.extend_from_slice(central_extra);
        bytes.extend_from_slice(file_comment);
        let central_end = bytes.len();

        push_u32(&mut bytes, 0x0605_4b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 1);
        push_u16(&mut bytes, 1);
        push_u32(&mut bytes, checked_u32(central_end - central_offset));
        push_u32(&mut bytes, checked_u32(central_offset));
        push_u16(&mut bytes, checked_u16(archive_comment.len()));
        bytes.extend_from_slice(archive_comment);
        (bytes, central_offset, central_end)
    }

    fn raw_named_zip(
        local_name: &[u8],
        central_name: &[u8],
        local_method: u16,
        central_method: u16,
        data: &[u8],
    ) -> Vec<u8> {
        let crc32 = soapberry_zip::crc32(data);
        let mut bytes = Vec::new();
        push_u32(&mut bytes, 0x0403_4b50);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 0x0800);
        push_u16(&mut bytes, local_method);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u32(&mut bytes, crc32);
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u16(&mut bytes, checked_u16(local_name.len()));
        push_u16(&mut bytes, 0);
        bytes.extend_from_slice(local_name);
        bytes.extend_from_slice(data);

        let central_offset = bytes.len();
        push_u32(&mut bytes, 0x0201_4b50);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 0x0800);
        push_u16(&mut bytes, central_method);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u32(&mut bytes, crc32);
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u16(&mut bytes, checked_u16(central_name.len()));
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u32(&mut bytes, 0);
        push_u32(&mut bytes, 0);
        bytes.extend_from_slice(central_name);

        let central_size = bytes.len() - central_offset;
        push_u32(&mut bytes, 0x0605_4b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 1);
        push_u16(&mut bytes, 1);
        push_u32(&mut bytes, checked_u32(central_size));
        push_u32(&mut bytes, checked_u32(central_offset));
        push_u16(&mut bytes, 0);
        bytes
    }

    fn raw_named_entries(entries: &[(&[u8], &[u8], u16, u16, &[u8])]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut offsets = Vec::new();
        for &(local_name, _central_name, local_method, _central_method, data) in entries {
            offsets.push(bytes.len());
            let crc32 = soapberry_zip::crc32(data);
            push_u32(&mut bytes, 0x0403_4b50);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 0x0800);
            push_u16(&mut bytes, local_method);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u32(&mut bytes, crc32);
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u16(&mut bytes, checked_u16(local_name.len()));
            push_u16(&mut bytes, 0);
            bytes.extend_from_slice(local_name);
            bytes.extend_from_slice(data);
        }

        let central_offset = bytes.len();
        for (index, &(_local_name, central_name, _local_method, central_method, data)) in
            entries.iter().enumerate()
        {
            let crc32 = soapberry_zip::crc32(data);
            push_u32(&mut bytes, 0x0201_4b50);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 0x0800);
            push_u16(&mut bytes, central_method);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u32(&mut bytes, crc32);
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u16(&mut bytes, checked_u16(central_name.len()));
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u32(&mut bytes, 0);
            push_u32(&mut bytes, checked_u32(offsets[index]));
            bytes.extend_from_slice(central_name);
        }

        let central_size = bytes.len() - central_offset;
        push_u32(&mut bytes, 0x0605_4b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, checked_u16(entries.len()));
        push_u16(&mut bytes, checked_u16(entries.len()));
        push_u32(&mut bytes, checked_u32(central_size));
        push_u32(&mut bytes, checked_u32(central_offset));
        push_u16(&mut bytes, 0);
        bytes
    }

    fn physical_two_entry_zip() -> Vec<u8> {
        let entries = [
            (
                b"Untouched/a".as_slice(),
                b"keep this payload".as_slice(),
                b"".as_slice(),
                b"".as_slice(),
                b"".as_slice(),
            ),
            (
                b"Opaque/entry.bin".as_slice(),
                b"old payload".as_slice(),
                b"\xaa\xbb\x03\0xyz".as_slice(),
                b"\xcc\xdd\x02\0\xfe\xed".as_slice(),
                b"entry-comment\0\xff".as_slice(),
            ),
        ];
        let archive_comment = b"archive-comment\0\xfe";
        let mut bytes = Vec::new();
        let mut local_offsets = [0usize; 2];
        for (index, (name, data, local_extra, _central_extra, _comment)) in
            entries.iter().enumerate()
        {
            local_offsets[index] = bytes.len();
            let crc32 = soapberry_zip::crc32(data);
            push_u32(&mut bytes, 0x0403_4b50);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 0x0800);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0x1234 + u16::try_from(index).unwrap_or(0));
            push_u16(&mut bytes, 0x5678);
            push_u32(&mut bytes, crc32);
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u16(&mut bytes, checked_u16(name.len()));
            push_u16(&mut bytes, checked_u16(local_extra.len()));
            bytes.extend_from_slice(name);
            bytes.extend_from_slice(local_extra);
            bytes.extend_from_slice(data);
        }

        let central_offset = bytes.len();
        for (index, (name, data, _local_extra, central_extra, comment)) in
            entries.iter().enumerate()
        {
            let crc32 = soapberry_zip::crc32(data);
            push_u32(&mut bytes, 0x0201_4b50);
            push_u16(&mut bytes, 0x0314);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 0x0800);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0x9abc);
            push_u16(&mut bytes, 0xdef0);
            push_u32(&mut bytes, crc32);
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u32(&mut bytes, checked_u32(data.len()));
            push_u16(&mut bytes, checked_u16(name.len()));
            push_u16(&mut bytes, checked_u16(central_extra.len()));
            push_u16(&mut bytes, checked_u16(comment.len()));
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u32(&mut bytes, 0);
            push_u32(&mut bytes, checked_u32(local_offsets[index]));
            bytes.extend_from_slice(name);
            bytes.extend_from_slice(central_extra);
            bytes.extend_from_slice(comment);
        }
        let central_size = bytes.len() - central_offset;
        push_u32(&mut bytes, 0x0605_4b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 2);
        push_u16(&mut bytes, 2);
        push_u32(&mut bytes, checked_u32(central_size));
        push_u32(&mut bytes, checked_u32(central_offset));
        push_u16(&mut bytes, checked_u16(archive_comment.len()));
        bytes.extend_from_slice(archive_comment);
        bytes
    }

    #[test]
    fn exact_artifacts_authorize_pointer_or_exact_bytes_but_not_fingerprints() {
        let source: Arc<[u8]> = b"exact package source".to_vec().into();
        let target: Arc<[u8]> = b"exact package target".to_vec().into();
        let mut artifacts = ExactArtifacts::new(Arc::clone(&source), target);

        assert!(artifacts.authorizes_source(&source));
        let equal_copy: Arc<[u8]> = source.as_ref().to_vec().into();
        assert!(!Arc::ptr_eq(&source, &equal_copy));
        assert!(artifacts.authorizes_source(&equal_copy));

        let tampered: Arc<[u8]> = b"tampered package source".to_vec().into();
        artifacts.source_fingerprint = diagnostic_fingerprint(&tampered);
        assert!(!artifacts.authorizes_source(&tampered));
    }

    #[test]
    fn exact_artifacts_clone_inverse_and_handoffs_share_allocations() {
        let source: Arc<[u8]> = b"source bytes".to_vec().into();
        let mut target: Arc<[u8]> = b"target bytes".to_vec().into();
        let artifacts = ExactArtifacts::new(Arc::clone(&source), Arc::clone(&target));

        assert!(Arc::get_mut(&mut target).is_none());
        let cloned = artifacts.clone();
        assert!(Arc::ptr_eq(&cloned.source(), &source));
        assert!(Arc::ptr_eq(&cloned.target(), &target));
        assert_eq!(cloned.source_fingerprint(), artifacts.source_fingerprint());
        assert_eq!(cloned.target_fingerprint(), artifacts.target_fingerprint());

        let inverse = artifacts.inverse();
        assert!(Arc::ptr_eq(&inverse.source(), &target));
        assert!(Arc::ptr_eq(&inverse.target(), &source));
        assert_eq!(inverse.source_fingerprint(), artifacts.target_fingerprint());
        assert_eq!(inverse.target_fingerprint(), artifacts.source_fingerprint());
    }

    #[test]
    fn exact_artifacts_retain_owned_vector_buffers_through_inverse() {
        let source = b"owned source bytes".to_vec();
        let target = b"owned target bytes".to_vec();
        let source_pointer = source.as_ptr();
        let target_pointer = target.as_ptr();
        let source = SharedBytes::from_owned_vec(source);
        let target = SharedBytes::from_owned_vec(target);
        let artifacts = OwnedExactArtifacts::new(source.clone(), target.clone());

        assert!(artifacts.authorizes_owner(&source));
        assert_eq!(artifacts.source_owner().as_ref().as_ptr(), source_pointer);
        assert_eq!(artifacts.target_owner().as_ref().as_ptr(), target_pointer);
        let inverse = artifacts.inverse();
        assert_eq!(inverse.source_owner().as_ref().as_ptr(), target_pointer);
        assert_eq!(inverse.target_owner().as_ref().as_ptr(), source_pointer);
    }

    #[test]
    fn exact_artifacts_byte_noop_requires_exact_equality() {
        let shared: Arc<[u8]> = b"same bytes".to_vec().into();
        let shared_pair = ExactArtifacts::new(Arc::clone(&shared), Arc::clone(&shared));
        assert!(shared_pair.is_byte_noop());

        let equal_copy: Arc<[u8]> = shared.as_ref().to_vec().into();
        let equal_pair = ExactArtifacts::new(Arc::clone(&shared), equal_copy);
        assert!(equal_pair.is_byte_noop());

        let changed: Arc<[u8]> = b"same bytez".to_vec().into();
        let changed_pair = ExactArtifacts::new(shared, changed);
        assert!(!changed_pair.is_byte_noop());
        let debug = format!("{changed_pair:?}");
        assert!(!debug.contains("same bytes"));
        assert!(!debug.contains("same bytez"));
    }

    #[test]
    fn preserves_flat_entry_order_and_payloads() -> Result<()> {
        let bytes = zip(&[("Metadata/a", b"a"), ("Index/Document.iwa", b"iwa")])?;
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(
            catalog.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Metadata/a", "Index/Document.iwa"]
        );
        assert_eq!(catalog.iter().next().map(Entry::data), Some(&b"a"[..]));
        Ok(())
    }

    #[test]
    fn logical_entry_profile_checks_declared_sizes_before_entry_read() -> Result<()> {
        for authority in [
            "Metadata/Properties.plist",
            "Metadata/BuildVersionHistory.plist",
            "Metadata/DocumentIdentifier",
        ] {
            let exact = vec![b'x'; LogicalEntryLimits::MAX_SEMANTIC_METADATA_BYTES as usize];
            let accepted = zip(&[(authority, &exact)])?;
            crate::zip::reset_test_entry_read_count();
            let catalog = Catalog::__from_bytes_with_logical_entry_limits(
                &accepted,
                Limits::default(),
                LogicalEntryLimits::PAGES_METADATA,
            )?;
            assert_eq!(
                catalog.iter().next().map(Entry::data),
                Some(exact.as_slice())
            );
            assert_eq!(crate::zip::test_entry_read_count(), 1);

            let one_over = vec![b'x'; exact.len() + 1];
            let rejected = zip(&[(authority, &one_over)])?;
            crate::zip::reset_test_entry_read_count();
            assert!(matches!(
                Catalog::__from_bytes_with_logical_entry_limits(
                    &rejected,
                    Limits::default(),
                    LogicalEntryLimits::PAGES_METADATA,
                ),
                Err(Error::Limit {
                    kind: crate::LimitKind::EntryBytes,
                    observed,
                    maximum,
                }) if observed == one_over.len() as u64
                    && maximum == LogicalEntryLimits::MAX_SEMANTIC_METADATA_BYTES
            ));
            assert_eq!(
                crate::zip::test_entry_read_count(),
                0,
                "oversized {authority} must fail from PhysicalEntry before read_entry"
            );
        }
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_does_not_promote_raw_normalization_near_names() -> Result<()> {
        for raw_name in [
            b"Metadata\\Properties.plist".as_slice(),
            b"/Metadata/Properties.plist",
            b"Metadata//Properties.plist",
            b"Metadata/./Properties.plist",
            b"Metadata/x/../Properties.plist",
        ] {
            let bytes = raw_named_zip(raw_name, raw_name, 0, 0, b"decoy");
            let archive = ZipArchive::new_with_limits(&bytes, Limits::default())?;
            preflight_flat_logical_entries(&archive, LogicalEntryLimits::PAGES_METADATA)?;
            crate::zip::reset_test_entry_read_count();
            let catalog = Catalog::__from_bytes_with_logical_entry_limits(
                &bytes,
                Limits::default(),
                LogicalEntryLimits::PAGES_METADATA,
            )?;
            assert!(
                catalog
                    .__pages_metadata_sidecars()?
                    .properties_plist()
                    .is_none()
            );
        }
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_refuses_selected_header_mismatches_before_read() -> Result<()> {
        let authority = b"Metadata/Properties.plist";
        let local_near_name = b"Metadata/Properties.plisX";
        let name_mismatch = raw_named_zip(local_near_name, authority, 0, 0, b"properties");
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            Catalog::__from_bytes_with_logical_entry_limits(
                &name_mismatch,
                Limits::default(),
                LogicalEntryLimits::PAGES_METADATA,
            ),
            Err(Error::InvalidBundle(message))
                if message.contains("mismatched local and central names")
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);

        let method_mismatch = raw_named_zip(authority, authority, 99, 0, b"properties");
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            Catalog::__from_bytes_with_logical_entry_limits(
                &method_mismatch,
                Limits::default(),
                LogicalEntryLimits::PAGES_METADATA,
            ),
            Err(Error::InvalidBundle(message))
                if message.contains("mismatched local and central compression methods")
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);
        Ok(())
    }

    #[test]
    fn semantic_profile_reads_only_exact_iwa_and_three_sidecars_in_modern_zip() -> Result<()> {
        let document = semantic_iwa(1)?;
        let metadata = semantic_iwa(2)?;
        let entries: [(&str, &[u8]); 8] = [
            ("Index/Document.iwa", &document),
            ("Index/Metadata.iwa", &metadata),
            ("Metadata/Properties.plist", b"properties"),
            ("Metadata/BuildVersionHistory.plist", b"history"),
            ("Metadata/DocumentIdentifier", b"identifier"),
            ("Data/asset.bin", b"asset"),
            ("Data/not-a-component.iwa", &metadata),
            ("Preview/preview.jpg", b"preview"),
        ];
        let bytes = zip(&entries)?;

        crate::zip::reset_test_entry_read_count();
        let semantic = Catalog::__from_bytes_with_logical_entry_limits(
            &bytes,
            Limits::default(),
            LogicalEntryLimits::SEMANTIC_METADATA,
        )?;
        assert_eq!(crate::zip::test_entry_read_count(), 5);
        assert_eq!(
            semantic.iter().map(Entry::name).collect::<Vec<_>>(),
            [
                "Index/Document.iwa",
                "Index/Metadata.iwa",
                "Metadata/Properties.plist",
                "Metadata/BuildVersionHistory.plist",
                "Metadata/DocumentIdentifier",
            ]
        );
        assert!(!semantic.source_is_exact());
        assert_eq!(semantic.source_provenance(), SourceProvenance::SemanticZip);

        crate::zip::reset_test_entry_read_count();
        let generic = Catalog::from_bytes(&bytes)?;
        assert_eq!(crate::zip::test_entry_read_count(), 8);
        assert_eq!(generic.len(), 8);
        assert!(generic.source_is_exact());
        Ok(())
    }

    #[test]
    fn semantic_components_profile_reads_no_sidecars() -> Result<()> {
        let document = semantic_iwa(1)?;
        let entries: [(&str, &[u8]); 5] = [
            ("Index/Document.iwa", &document),
            ("Metadata/Properties.plist", b"properties"),
            ("Metadata/BuildVersionHistory.plist", b"history"),
            ("Metadata/DocumentIdentifier", b"identifier"),
            ("Data/asset.bin", b"asset"),
        ];
        let bytes = zip(&entries)?;

        crate::zip::reset_test_entry_read_count();
        let catalog = Catalog::__from_bytes_with_logical_entry_limits(
            &bytes,
            Limits::default(),
            LogicalEntryLimits::SEMANTIC_COMPONENTS,
        )?;
        assert_eq!(crate::zip::test_entry_read_count(), 1);
        assert_eq!(
            catalog.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Index/Document.iwa"]
        );
        assert!(
            catalog
                .__semantic_metadata_sidecars()?
                .properties_plist()
                .is_none()
        );
        Ok(())
    }

    #[test]
    fn semantic_profile_reads_only_nested_iwa_and_prefixed_sidecars_in_legacy_zip() -> Result<()> {
        let document = semantic_iwa(1)?;
        let metadata = semantic_iwa(2)?;
        let index = zip(&[
            ("Index/Document.iwa", document.as_slice()),
            ("Index/Metadata.iwa", metadata.as_slice()),
            ("Data/not-a-component.iwa", metadata.as_slice()),
        ])?;
        let outer_entries: [(&str, &[u8]); 6] = [
            ("legacy.pages/Index.zip", &index),
            ("legacy.pages/Metadata/Properties.plist", b"properties"),
            (
                "legacy.pages/Metadata/BuildVersionHistory.plist",
                b"history",
            ),
            ("legacy.pages/Metadata/DocumentIdentifier", b"identifier"),
            ("legacy.pages/Data/asset.bin", b"asset"),
            ("legacy.pages/Preview/preview.jpg", b"preview"),
        ];
        let bytes = zip(&outer_entries)?;

        crate::zip::reset_test_entry_read_count();
        let semantic = Catalog::__from_bytes_with_logical_entry_limits(
            &bytes,
            Limits::default(),
            LogicalEntryLimits::SEMANTIC_METADATA,
        )?;
        assert_eq!(crate::zip::test_entry_read_count(), 6);
        assert_eq!(semantic.len(), 5);
        assert!(!semantic.source_is_exact());
        assert_eq!(
            semantic.source_provenance(),
            SourceProvenance::LegacySemanticZip
        );

        crate::zip::reset_test_entry_read_count();
        let generic = Catalog::from_bytes(&bytes)?;
        assert_eq!(crate::zip::test_entry_read_count(), 8);
        assert_eq!(generic.len(), 8);
        assert_eq!(generic.source_provenance(), SourceProvenance::LegacyZip);
        Ok(())
    }

    #[test]
    fn semantic_profile_skips_irrelevant_opaque_and_refuses_selected_opaque_before_read()
    -> Result<()> {
        for name in [b"Data/opaque.bin".as_slice(), b"Preview/opaque.jpg"] {
            let irrelevant = raw_named_zip(name, name, 99, 99, b"opaque");
            crate::zip::reset_test_entry_read_count();
            let catalog = Catalog::__from_bytes_with_logical_entry_limits(
                &irrelevant,
                Limits::default(),
                LogicalEntryLimits::SEMANTIC_METADATA,
            )?;
            assert!(catalog.is_empty());
            assert_eq!(crate::zip::test_entry_read_count(), 0);
        }

        let selected = raw_named_zip(
            b"Index/Document.iwa",
            b"Index/Document.iwa",
            99,
            99,
            b"opaque",
        );
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            Catalog::__from_bytes_with_logical_entry_limits(
                &selected,
                Limits::default(),
                LogicalEntryLimits::SEMANTIC_COMPONENTS,
            ),
            Err(Error::InvalidBundle(message)) if message.contains("unsupported ZIP compression")
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);
        Ok(())
    }

    #[test]
    fn semantic_profile_batch_preflights_late_iwa_before_any_component_read() -> Result<()> {
        let inner = raw_named_entries(&[
            (
                b"Index/Document.iwa",
                b"Index/Document.iwa",
                0,
                0,
                b"document",
            ),
            (
                b"Index/Metadata.iwa",
                b"Index/Metadata.iwa",
                99,
                99,
                b"metadata",
            ),
        ]);

        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            Catalog::__from_bytes_with_logical_entry_limits(
                &inner,
                Limits::default(),
                LogicalEntryLimits::SEMANTIC_COMPONENTS,
            ),
            Err(Error::InvalidBundle(message)) if message.contains("unsupported ZIP compression")
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);

        let outer = zip(&[("legacy.pages/Index.zip", inner.as_slice())])?;
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            Catalog::__from_bytes_with_logical_entry_limits(
                &outer,
                Limits::default(),
                LogicalEntryLimits::SEMANTIC_COMPONENTS,
            ),
            Err(Error::InvalidBundle(message)) if message.contains("unsupported ZIP compression")
        ));
        assert_eq!(
            crate::zip::test_entry_read_count(),
            1,
            "legacy detection must read only the selected Index.zip container"
        );
        Ok(())
    }

    #[test]
    fn semantic_detection_reads_only_the_exact_modern_root() -> Result<()> {
        let document = semantic_iwa(1)?;
        let bytes = raw_named_entries(&[
            (
                b"Index/Document.iwa",
                b"Index/Document.iwa",
                0,
                0,
                document.as_slice(),
            ),
            (
                b"Index/Foreign-local.iwa",
                b"Index/Foreign.iwa",
                99,
                0,
                b"corrupt foreign payload",
            ),
            (
                b"Index/Opaque.iwa",
                b"Index/Opaque.iwa",
                99,
                99,
                b"opaque foreign payload",
            ),
            (
                b"Index/nested/Slide.iwa",
                b"Index/nested/Slide.iwa",
                0,
                0,
                b"corrupt nested basename decoy",
            ),
        ]);

        crate::zip::reset_test_entry_read_count();
        let root = crate::inspect_semantic_detection_root(&bytes, Limits::default())?;
        assert!(root.has_iwa_components());
        assert!(!root.has_keynote_components());
        assert!(root.document().is_some());
        assert_eq!(crate::zip::test_entry_read_count(), 1);
        Ok(())
    }

    #[test]
    fn semantic_detection_reads_only_the_legacy_container_and_root() -> Result<()> {
        let document = semantic_iwa(1)?;
        let index = raw_named_entries(&[
            (b"Document.iwa", b"Document.iwa", 0, 0, document.as_slice()),
            (
                b"Foreign-local.iwa",
                b"Foreign.iwa",
                99,
                0,
                b"corrupt foreign payload",
            ),
            (
                b"Opaque.iwa",
                b"Opaque.iwa",
                99,
                99,
                b"opaque foreign payload",
            ),
            (
                b"nested/TemplateSlide.iwa",
                b"nested/TemplateSlide.iwa",
                0,
                0,
                b"corrupt nested basename decoy",
            ),
        ]);
        let bytes = zip(&[("legacy.pages/Index.zip", index.as_slice())])?;

        crate::zip::reset_test_entry_read_count();
        let root = crate::inspect_semantic_detection_root(&bytes, Limits::default())?;
        assert!(root.has_iwa_components());
        assert!(!root.has_keynote_components());
        assert!(root.document().is_some());
        assert_eq!(crate::zip::test_entry_read_count(), 2);
        Ok(())
    }

    #[test]
    fn semantic_detection_does_not_gather_high_count_non_root_names() -> Result<()> {
        let document = semantic_iwa(1)?;
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &document)?;
        for index in 0..1_024 {
            writer.write_stored(&format!("Index/Foreign-{index}.iwa"), b"corrupt")?;
        }
        let bytes = writer.finish_to_bytes()?;

        crate::zip::reset_test_entry_read_count();
        let root = crate::inspect_semantic_detection_root(&bytes, Limits::default())?;
        assert!(root.document().is_some());
        assert_eq!(crate::zip::test_entry_read_count(), 1);
        Ok(())
    }

    #[test]
    fn semantic_detection_refuses_root_and_nested_index_raw_aliases() -> Result<()> {
        let document = semantic_iwa(1)?;
        for alias in [
            b"Index//Document.iwa".as_slice(),
            b"Index/./Document.iwa",
            b"Index\\Document.iwa",
        ] {
            let bytes = raw_named_zip(alias, alias, 0, 0, &document);
            assert!(matches!(
                crate::inspect_semantic_detection_root(&bytes, Limits::default()),
                Err(Error::InvalidBundle(message))
                    if message.contains("non-canonical raw Document.iwa authority")
            ));
        }

        let index = zip(&[("Document.iwa", document.as_slice())])?;
        for alias in [
            b"legacy.pages/./Index.zip".as_slice(),
            b"legacy.pages\\Index.zip",
        ] {
            let bytes = raw_named_zip(alias, alias, 0, 0, &index);
            assert!(matches!(
                crate::inspect_semantic_detection_root(&bytes, Limits::default()),
                Err(Error::InvalidBundle(message))
                    if message.contains("non-canonical raw Index.zip authority")
            ));
        }

        let duplicate_root = raw_named_entries(&[
            (
                b"Index/Document.iwa",
                b"Index/Document.iwa",
                0,
                0,
                document.as_slice(),
            ),
            (
                b"Index/Document.iwa",
                b"Index/Document.iwa",
                0,
                0,
                document.as_slice(),
            ),
        ]);
        let duplicate_result =
            crate::inspect_semantic_detection_root(&duplicate_root, Limits::default());
        assert!(duplicate_result.is_err());

        let mixed = zip(&[
            ("Index/Document.iwa", document.as_slice()),
            ("legacy.pages/Index.zip", index.as_slice()),
        ])?;
        assert!(matches!(
            crate::inspect_semantic_detection_root(&mixed, Limits::default()),
            Err(Error::InvalidBundle(message)) if message.contains("mixes direct IWA")
        ));
        Ok(())
    }

    #[test]
    fn writes_flat_entry_order_and_payloads_without_exposing_zip_types() -> Result<()> {
        let entries: [(&str, &[u8]); 2] = [("Metadata/a", b"a"), ("Index/Document.iwa", b"opaque")];
        let bytes = to_bytes(entries.iter().copied(), Limits::default())?;
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(
            catalog
                .into_iter()
                .map(Entry::into_parts)
                .collect::<Vec<_>>(),
            [
                ("Metadata/a".to_owned(), b"a".to_vec()),
                ("Index/Document.iwa".to_owned(), b"opaque".to_vec()),
            ]
        );
        Ok(())
    }

    #[test]
    fn preserves_physical_metadata_and_exact_noop_bytes() -> Result<()> {
        let (bytes, central_offset, central_end) = physical_zip(0);
        let catalog = Catalog::from_bytes(&bytes)?;
        let entry = catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("physical metadata test produced no entry".to_owned())
        })?;

        assert!(!entry.is_opaque());
        assert_eq!(entry.raw_name(), b"Opaque/entry.bin");
        assert_eq!(entry.data(), b"opaque payload");
        assert_eq!(entry.metadata().local().flags(), 0x0800);
        assert_eq!(entry.metadata().central().flags(), 0x0800);
        assert_eq!(entry.metadata().local().compression_method(), 0);
        assert_eq!(entry.metadata().central().compression_method(), 0);
        assert_eq!(entry.metadata().local().last_modified().time(), 0x1234);
        assert_eq!(entry.metadata().local().last_modified().date(), 0x5678);
        assert_eq!(entry.metadata().central().last_modified().time(), 0x9abc);
        assert_eq!(entry.metadata().central().last_modified().date(), 0xdef0);
        assert_eq!(entry.metadata().local().extra(), b"\xaa\xbb\x03\0xyz");
        assert_eq!(
            entry.metadata().central().extra(),
            b"\xcc\xdd\x02\0\xfe\xed"
        );
        assert_eq!(entry.metadata().central().comment(), b"entry-comment\0\xff");
        assert_eq!(entry.raw_record().local_record(), &bytes[..central_offset]);
        assert_eq!(
            entry.raw_record().central_directory_record(),
            &bytes[central_offset..central_end]
        );
        assert_eq!(entry.raw_record().compressed_data(), b"opaque payload");
        assert_eq!(catalog.to_bytes()?, bytes);
        let mut streamed = Vec::new();
        catalog.write_to(&mut streamed)?;
        assert_eq!(streamed, bytes);
        Ok(())
    }

    #[test]
    fn accepts_shared_source_without_copying_input() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(0);
        let source: Arc<[u8]> = bytes.into();
        let catalog = Catalog::from_shared_bytes(source.clone())?;

        assert!(Arc::ptr_eq(
            &source,
            &catalog
                .source
                .shared_slice()
                .expect("shared ingress keeps its Arc<[u8]> owner")
        ));
        assert_eq!(catalog.len(), 1);
        Ok(())
    }

    #[test]
    fn accepts_owned_source_without_copying_payload_buffer() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(0);
        let original_pointer = bytes.as_ptr();
        let original_capacity = bytes.capacity();
        let catalog = Catalog::from_owned_bytes_with_limits(bytes, Limits::default())?;
        let owner = catalog.source_owner();

        assert_eq!(owner.as_ref().as_ptr(), original_pointer);
        assert!(original_capacity >= owner.as_ref().len());
        assert_eq!(catalog.source_bytes().as_ptr(), original_pointer);
        assert!(SharedBytes::ptr_eq(&owner, &catalog.source_owner()));
        assert_eq!(catalog.len(), 1);
        Ok(())
    }

    #[test]
    fn shared_source_respects_input_limits() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(0);
        let source: Arc<[u8]> = bytes.into();
        let limits = Limits::new(1, 10, 100, 100, 100)?;

        assert!(matches!(
            Catalog::from_shared_bytes_with_limits(source, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::InputBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn borrowed_source_respects_input_limits_before_copying() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(0);
        let limits = Limits::new(1, 10, 100, 100, 100)?;

        assert!(matches!(
            Catalog::from_bytes_with_limits(&bytes, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::InputBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn shared_source_preserves_opaque_zip_bytes_exactly() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(99);
        let source: Arc<[u8]> = bytes.into();
        let catalog = Catalog::from_shared_bytes(source.clone())?;
        let entry = catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("shared source test produced no entry".to_owned())
        })?;

        assert!(entry.is_opaque());
        assert_eq!(entry.raw_record().compressed_data(), b"opaque payload");
        assert_eq!(catalog.to_bytes()?.as_slice(), source.as_ref());
        let mut streamed = Vec::new();
        catalog.write_to(&mut streamed)?;
        assert_eq!(streamed.as_slice(), source.as_ref());
        Ok(())
    }

    #[test]
    fn snapshots_a_positional_source_in_bounded_reads() -> Result<()> {
        let bytes = zip(&[("Metadata/a", b"a"), ("Data/b", b"b")])?;
        let source = TestSource::new(bytes.clone(), 3, false);
        let catalog = Catalog::from_read_at(&source)?;

        assert!(source.reads.load(Ordering::Relaxed) > 1);
        assert_eq!(catalog.to_bytes()?, bytes);
        assert_eq!(
            catalog.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Metadata/a", "Data/b"]
        );
        Ok(())
    }

    #[test]
    fn checks_positional_source_limits_before_reading() -> Result<()> {
        let bytes = zip(&[("Data/a", b"a")])?;
        let source = TestSource::new(bytes, 4, false);
        let limits = Limits::new(1, 10, 100, 100, 100)?;

        assert!(matches!(
            Catalog::from_read_at_with_limits(&source, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::InputBytes,
                ..
            })
        ));
        assert_eq!(source.reads.load(Ordering::Relaxed), 0);
        Ok(())
    }

    #[test]
    fn rejects_a_positional_source_that_changes_during_snapshot() -> Result<()> {
        let bytes = zip(&[("Data/a", b"a")])?;
        let source = TestSource::new(bytes, 4, true);
        let error = Catalog::from_read_at(&source).err().ok_or_else(|| {
            Error::InvalidBundle("changed source unexpectedly parsed successfully".to_owned())
        })?;

        assert!(matches!(
            error,
            Error::SourceChanged { expected, observed }
                if expected.id() == 41
                    && expected.revision() == 0
                    && observed.id() == 41
                    && observed.revision() == 1
        ));
        Ok(())
    }

    #[test]
    fn retains_unsupported_compression_as_opaque_raw_record() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(99);
        let catalog = Catalog::from_bytes(&bytes)?;
        let entry = catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("opaque metadata test produced no entry".to_owned())
        })?;

        assert!(entry.is_opaque());
        assert_eq!(entry.metadata().central().compression_method(), 99);
        assert_eq!(entry.data(), b"opaque payload");
        assert!(matches!(
            entry.payload(),
            EntryPayload::Opaque(record) if record.compressed_data() == b"opaque payload"
        ));
        assert_eq!(catalog.to_bytes()?, bytes);
        Ok(())
    }

    #[test]
    fn rejects_truncated_local_metadata_before_materializing_payload() {
        let (mut bytes, _central_offset, _central_end) = physical_zip(0);
        bytes[28..30].copy_from_slice(&u16::MAX.to_le_bytes());
        assert!(matches!(
            Catalog::from_bytes(&bytes),
            Err(Error::InvalidBundle(message)) if message.contains("local file header")
        ));
    }

    #[test]
    fn rejects_output_limits_before_writing_any_bytes() -> Result<()> {
        let entries: [(&str, &[u8]); 2] = [("Data/a", b"a"), ("Data/b", b"b")];
        let limits = Limits::new(1024, 1, 1024, 1024, 1024)?;
        let mut sink = Vec::new();
        let result = write_to(entries.iter().copied(), &mut sink, limits);
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidBundle("output unexpectedly succeeded".to_owned()))?;
        assert!(matches!(
            error,
            Error::Limit {
                kind: crate::LimitKind::Entries,
                ..
            }
        ));
        assert!(sink.is_empty());
        Ok(())
    }

    #[test]
    fn flattens_legacy_index_before_outer_entries() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa")])?;
        let bytes = zip(&[
            ("legacy.pages/Index.zip", &index),
            ("legacy.pages/Data/a", b"a"),
        ])?;
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(
            catalog.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Index/Document.iwa", "Data/a"]
        );
        Ok(())
    }

    #[test]
    fn preserves_nested_index_limit_error_type() -> Result<()> {
        let index = zip(&[
            ("Index/Document.iwa", b"iwa"),
            ("Index/CalculationEngine.iwa", b"iwa"),
        ])?;
        let bytes = zip(&[("legacy.pages/Index.zip", &index)])?;
        let input = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("test legacy package length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(input, 1, input, input, 1024)?;

        assert!(matches!(
            Catalog::from_bytes_with_limits(&bytes, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::Entries,
                observed: 2,
                maximum: 1,
            })
        ));
        Ok(())
    }

    #[test]
    fn rejects_declared_nested_index_size_before_decompression() -> Result<()> {
        let repeated = vec![0u8; 64 * 1024];
        let index = zip(&[("Index/Document.iwa", repeated.as_slice())])?;
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("legacy.pages/Index.zip", &index)?;
        let bytes = writer.finish_to_bytes()?;
        assert!(bytes.len() < index.len());
        let input = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("test legacy package length does not fit u64".to_owned())
        })?;
        let nested = u64::try_from(index.len()).map_err(|_error| {
            Error::InvalidBundle("test nested index length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(input, 10, nested, nested, 1024)?;

        assert!(matches!(
            Catalog::from_bytes_with_limits(&bytes, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::InputBytes,
                observed,
                maximum,
            }) if observed == nested && maximum == input
        ));
        Ok(())
    }

    #[test]
    fn rejects_legacy_non_iwa_members() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa"), ("Metadata/a", b"bad")])?;
        let bytes = zip(&[("legacy.pages/Index.zip", &index)])?;
        assert!(matches!(
            Catalog::from_bytes(&bytes),
            Err(Error::InvalidBundle(message)) if message.contains("non-IWA")
        ));
        Ok(())
    }

    #[test]
    fn rejects_mixed_direct_and_legacy_representations() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa")])?;
        let bytes = zip(&[
            ("legacy.pages/Index.zip", &index),
            ("Index/CalculationEngine.iwa", b"iwa"),
        ])?;
        assert!(matches!(
            Catalog::from_bytes(&bytes),
            Err(Error::InvalidBundle(message)) if message.contains("mixes direct IWA")
        ));
        Ok(())
    }

    #[test]
    fn reassembles_stored_edit_while_preserving_physical_provenance() -> Result<()> {
        let bytes = physical_two_entry_zip();
        let catalog = Catalog::from_bytes(&bytes)?;
        let before = catalog.iter().collect::<Vec<_>>();
        let edited = catalog.reassemble_to_bytes(
            &[EntryEdit::new(
                "Opaque/entry.bin",
                b"new payload with a different size",
            )],
            Limits::default(),
        )?;
        let after_catalog = Catalog::from_bytes(&edited)?;
        let after = after_catalog.iter().collect::<Vec<_>>();

        assert_eq!(
            after.iter().map(|entry| entry.name()).collect::<Vec<_>>(),
            ["Untouched/a", "Opaque/entry.bin"]
        );
        assert_eq!(after[0].data(), before[0].data());
        assert_eq!(
            after[0].raw_record().local_record(),
            before[0].raw_record().local_record()
        );
        assert_eq!(
            after[0].raw_record().central_directory_record(),
            before[0].raw_record().central_directory_record()
        );
        assert_eq!(after[1].data(), b"new payload with a different size");
        assert_eq!(after[1].raw_name(), before[1].raw_name());
        assert_eq!(after[1].metadata().local(), before[1].metadata().local());
        assert_eq!(
            after[1].metadata().central().name(),
            before[1].metadata().central().name()
        );
        assert_eq!(
            after[1].metadata().central().extra(),
            before[1].metadata().central().extra()
        );
        assert_eq!(
            after[1].metadata().central().comment(),
            before[1].metadata().central().comment()
        );
        let archive = soapberry_zip::ZipArchive::from_slice(&edited)?;
        assert_eq!(archive.comment().as_bytes(), b"archive-comment\0\xfe");
        Ok(())
    }

    #[test]
    fn reassembles_deflate_edit_and_round_trips_metadata() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("Compressed/data", b"old deflate payload")?;
        writer.write_stored("Untouched/data", b"keep")?;
        let bytes = writer.finish_to_bytes()?;
        let catalog = Catalog::from_bytes(&bytes)?;
        let before = catalog
            .iter()
            .next()
            .ok_or_else(|| Error::InvalidBundle("deflate fixture has no first entry".to_owned()))?;
        assert_eq!(before.metadata().central().compression_method(), 8);

        let edited = catalog.reassemble_to_bytes(
            &[EntryEdit::new(
                "Compressed/data",
                b"replacement deflate payload with more bytes",
            )],
            Limits::default(),
        )?;
        let after_catalog = Catalog::from_bytes(&edited)?;
        let after = after_catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("reassembled deflate fixture has no first entry".to_owned())
        })?;

        assert_eq!(after.data(), b"replacement deflate payload with more bytes");
        assert_eq!(after.metadata().central().compression_method(), 8);
        assert_eq!(after.metadata().local(), before.metadata().local());
        assert_eq!(
            after.metadata().central().name(),
            before.metadata().central().name()
        );
        assert_eq!(
            after.metadata().central().extra(),
            before.metadata().central().extra()
        );
        assert_eq!(
            after.metadata().central().comment(),
            before.metadata().central().comment()
        );
        assert_eq!(
            after.metadata().crc32(),
            soapberry_zip::crc32(b"replacement deflate payload with more bytes")
        );
        Ok(())
    }

    #[test]
    fn deletes_a_member_and_preserves_the_remaining_physical_record() -> Result<()> {
        let bytes = physical_two_entry_zip();
        let catalog = Catalog::from_bytes(&bytes)?;
        let before = catalog.iter().collect::<Vec<_>>();
        let rebuilt =
            catalog.reassemble_with_deletions_to_bytes(&[], &["Untouched/a"], Limits::default())?;
        let after_catalog = Catalog::from_bytes(&rebuilt)?;
        let after = after_catalog.iter().collect::<Vec<_>>();

        assert_eq!(after.len(), 1);
        assert_eq!(after[0].name(), "Opaque/entry.bin");
        assert_eq!(after[0].data(), before[1].data());
        assert_eq!(
            after[0].raw_record().compressed_data(),
            before[1].raw_record().compressed_data()
        );
        assert_eq!(
            after[0].raw_record().local_record(),
            before[1].raw_record().local_record()
        );
        assert_eq!(after[0].raw_name(), before[1].raw_name());
        assert_eq!(after[0].metadata().local(), before[1].metadata().local());
        assert_eq!(
            after[0].metadata().central(),
            before[1].metadata().central()
        );

        let old_central = before[1].raw_record().central_directory_record();
        let new_central = after[0].raw_record().central_directory_record();
        assert_eq!(&new_central[..42], &old_central[..42]);
        assert_eq!(&new_central[46..], &old_central[46..]);
        assert_ne!(&new_central[42..46], &old_central[42..46]);

        let archive = soapberry_zip::ZipArchive::from_slice(&rebuilt)?;
        assert_eq!(archive.comment().as_bytes(), b"archive-comment\0\xfe");
        let eocd = rebuilt.len() - (22 + b"archive-comment\0\xfe".len());
        assert_eq!(raw_u16(&rebuilt[eocd..], 8), Some(1));
        assert_eq!(raw_u16(&rebuilt[eocd..], 10), Some(1));

        let ordered = zip(&[
            ("Data/first", b"1"),
            ("Data/middle", b"2"),
            ("Data/last", b"3"),
        ])?;
        let ordered_catalog = Catalog::from_bytes(&ordered)?;
        let ordered_rebuilt = ordered_catalog.reassemble_with_deletions_to_bytes(
            &[],
            &["Data/middle"],
            Limits::default(),
        )?;
        let retained = Catalog::from_bytes(&ordered_rebuilt)?;
        assert_eq!(
            retained.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Data/first", "Data/last"]
        );
        assert_eq!(
            retained.iter().map(Entry::data).collect::<Vec<_>>(),
            [&b"1"[..], &b"3"[..]]
        );
        Ok(())
    }

    #[test]
    fn combines_an_edit_and_deletion_in_one_reassembly() -> Result<()> {
        let bytes = physical_two_entry_zip();
        let catalog = Catalog::from_bytes(&bytes)?;
        let before = catalog.iter().collect::<Vec<_>>();
        let rebuilt = catalog.reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(
                "Opaque/entry.bin",
                b"edited after the preceding member was deleted",
            )],
            &["Untouched/a"],
            Limits::default(),
        )?;
        let after_catalog = Catalog::from_bytes(&rebuilt)?;
        let after = after_catalog.iter().collect::<Vec<_>>();

        assert_eq!(after.len(), 1);
        assert_eq!(after[0].name(), "Opaque/entry.bin");
        assert_eq!(
            after[0].data(),
            b"edited after the preceding member was deleted"
        );
        assert_eq!(after[0].raw_name(), before[1].raw_name());
        assert_eq!(after[0].metadata().local(), before[1].metadata().local());
        assert_eq!(
            after[0].metadata().central().extra(),
            before[1].metadata().central().extra()
        );
        assert_eq!(
            after[0].metadata().central().comment(),
            before[1].metadata().central().comment()
        );
        Ok(())
    }

    #[test]
    fn deletion_can_remove_an_opaque_member_without_decoding_it() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(99);
        let catalog = Catalog::from_bytes(&bytes)?;
        assert!(catalog.iter().next().is_some_and(Entry::is_opaque));

        let rebuilt = catalog.reassemble_with_deletions_to_bytes(
            &[],
            &["Opaque/entry.bin"],
            Limits::default(),
        )?;
        let after = Catalog::from_bytes(&rebuilt)?;
        assert!(after.is_empty());
        let eocd = rebuilt.len() - (22 + b"archive-comment\0\xfe".len());
        assert_eq!(raw_u16(&rebuilt[eocd..], 8), Some(0));
        assert_eq!(raw_u16(&rebuilt[eocd..], 10), Some(0));
        assert_eq!(raw_u32(&rebuilt[eocd..], 12), Some(0));
        Ok(())
    }

    #[test]
    fn deletion_selection_is_exact_unique_and_disjoint_from_edits() -> Result<()> {
        let bytes = physical_two_entry_zip();
        let catalog = Catalog::from_bytes(&bytes)?;

        let deletion_cases: [&[&str]; 3] = [
            &["Missing/entry"],
            &["Untouched"],
            &["Untouched/a", "Untouched/a"],
        ];
        for deleted_names in deletion_cases {
            assert!(matches!(
                catalog.reassemble_with_deletions_to_bytes(&[], deleted_names, Limits::default()),
                Err(Error::Reassembly(_))
            ));
        }
        assert!(matches!(
            catalog.reassemble_with_deletions_to_bytes(
                &[EntryEdit::new("Untouched/a", b"edit")],
                &["Untouched/a"],
                Limits::default()
            ),
            Err(Error::Reassembly(message)) if message.contains("both edited and deleted")
        ));
        Ok(())
    }

    #[test]
    fn deletion_reassembly_enforces_shape_and_output_limits() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa")])?;
        let legacy = zip(&[
            ("legacy.pages/Index.zip", index.as_slice()),
            ("legacy.pages/Data/a", b"a"),
        ])?;
        let legacy_catalog = Catalog::from_bytes(&legacy)?;
        assert_eq!(
            legacy_catalog.reassemble_with_deletions_to_bytes(&[], &[], Limits::default())?,
            legacy
        );
        assert!(matches!(
            legacy_catalog.reassemble_with_deletions_to_bytes(
                &[],
                &["Data/a"],
                Limits::default()
            ),
            Err(Error::Reassembly(message)) if message.contains("legacy nested Index.zip")
        ));

        let bytes = physical_two_entry_zip();
        let eocd = bytes
            .windows(4)
            .rposition(|window| window == 0x0605_4b50u32.to_le_bytes())
            .ok_or_else(|| Error::InvalidBundle("test ZIP has no EOCD".to_owned()))?;
        let tail = &bytes[eocd..];
        let central_size = raw_u32(tail, 12).ok_or_else(|| {
            Error::InvalidBundle("test ZIP has no central directory size".to_owned())
        })?;
        let central_offset = raw_u32(tail, 16).ok_or_else(|| {
            Error::InvalidBundle("test ZIP has no central directory offset".to_owned())
        })?;
        let mut with_zip64_locator = Vec::new();
        with_zip64_locator.extend_from_slice(&bytes[..eocd]);
        push_u32(&mut with_zip64_locator, 0x0606_4b50);
        with_zip64_locator.extend_from_slice(&44u64.to_le_bytes());
        push_u16(&mut with_zip64_locator, 45);
        push_u16(&mut with_zip64_locator, 45);
        push_u32(&mut with_zip64_locator, 0);
        push_u32(&mut with_zip64_locator, 0);
        with_zip64_locator.extend_from_slice(&2u64.to_le_bytes());
        with_zip64_locator.extend_from_slice(&2u64.to_le_bytes());
        with_zip64_locator.extend_from_slice(&u64::from(central_size).to_le_bytes());
        with_zip64_locator.extend_from_slice(&u64::from(central_offset).to_le_bytes());
        push_u32(&mut with_zip64_locator, 0x0706_4b50);
        push_u32(&mut with_zip64_locator, 0);
        with_zip64_locator.extend_from_slice(
            &u64::try_from(eocd)
                .map_err(|_error| {
                    Error::InvalidBundle("test ZIP64 offset does not fit u64".to_owned())
                })?
                .to_le_bytes(),
        );
        push_u32(&mut with_zip64_locator, 1);
        let standard_eocd = with_zip64_locator.len();
        with_zip64_locator.extend_from_slice(tail);
        with_zip64_locator[standard_eocd + 8..standard_eocd + 12].copy_from_slice(&[0xff; 4]);
        let zip64_catalog = Catalog::from_bytes(&with_zip64_locator)?;
        assert!(matches!(
            zip64_catalog.reassemble_with_deletions_to_bytes(
                &[],
                &["Untouched/a"],
                Limits::default()
            ),
            Err(Error::Reassembly(_))
        ));

        let source_size = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("test ZIP length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(source_size, 10, 4096, 4096, 1024)?;
        assert!(matches!(
            Catalog::from_bytes(&bytes)?.reassemble_with_deletions_to_bytes(
                &[EntryEdit::new("Opaque/entry.bin", &[b'x'; 1024])],
                &["Untouched/a"],
                limits
            ),
            Err(Error::Limit {
                kind: crate::LimitKind::OutputBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn empty_reassembly_uses_the_exact_noop_source_path() -> Result<()> {
        let bytes = physical_two_entry_zip();
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(catalog.reassemble_to_bytes(&[], Limits::default())?, bytes);
        assert_eq!(
            catalog.reassemble_with_deletions_to_bytes(&[], &[], Limits::default())?,
            bytes
        );
        let mut streamed = Vec::new();
        catalog.write_reassembled_to(&[], &mut streamed, Limits::default())?;
        assert_eq!(streamed, bytes);

        let source_size = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("test ZIP length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(source_size - 1, 10, 4096, 4096, 1024)?;
        let mut limited_sink = vec![0xde, 0xad, 0xbe, 0xef];
        assert!(matches!(
            catalog.write_reassembled_to(&[], &mut limited_sink, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::InputBytes,
                ..
            })
        ));
        assert_eq!(limited_sink, [0xde, 0xad, 0xbe, 0xef]);
        Ok(())
    }

    #[test]
    fn rejected_edits_leave_the_sink_untouched() -> Result<()> {
        let bytes = physical_two_entry_zip();
        let catalog = Catalog::from_bytes(&bytes)?;
        let original_sink = vec![0xde, 0xad, 0xbe, 0xef];

        let cases = [
            vec![
                EntryEdit::new("Opaque/entry.bin", b"one"),
                EntryEdit::new("Opaque/entry.bin", b"two"),
            ],
            vec![EntryEdit::new("Missing/entry", b"missing")],
        ];
        for edits in cases {
            let mut sink = original_sink.clone();
            assert!(matches!(
                catalog.write_reassembled_to(&edits, &mut sink, Limits::default()),
                Err(Error::Reassembly(_))
            ));
            assert_eq!(sink, original_sink);
        }

        let (opaque_bytes, _central_offset, _central_end) = physical_zip(99);
        let opaque_catalog = Catalog::from_bytes(&opaque_bytes)?;
        let mut opaque_sink = original_sink.clone();
        assert!(matches!(
            opaque_catalog.write_reassembled_to(
                &[EntryEdit::new("Opaque/entry.bin", b"cannot edit")],
                &mut opaque_sink,
                Limits::default()
            ),
            Err(Error::Reassembly(_))
        ));
        assert_eq!(opaque_sink, original_sink);

        let input_limit = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("test ZIP length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(input_limit, 10, 4096, 4096, 1024)?;
        let mut limited_sink = original_sink.clone();
        assert!(matches!(
            catalog.write_reassembled_to(
                &[EntryEdit::new("Opaque/entry.bin", &[b'x'; 1024])],
                &mut limited_sink,
                limits
            ),
            Err(Error::Limit {
                kind: crate::LimitKind::OutputBytes,
                ..
            })
        ));
        assert_eq!(limited_sink, original_sink);
        Ok(())
    }
}
