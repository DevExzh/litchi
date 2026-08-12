//! Immutable, source-backed access to an OPC package.
//!
//! This module intentionally exposes a smaller surface than [`OpcPackage`].
//! The latter owns mutable parts, while this type keeps ordinary payloads in a
//! positional source until a caller explicitly asks for one.

use crate::constants::{content_type, relationship_type};
use crate::error::{OpcError, Result};
use crate::limits::{ReadLimits, ReadResource};
use crate::members::NonPartMember;
use crate::package::OpcPackage;
use crate::packuri::{PACKAGE_URI, PackURI};
use crate::part::PartFactory;
use crate::pkgreader::{PackageReader, SerializedRelationship, SourceCatalog};
use crate::rel::Relationships;
use litchi_core::{ReadAt, SourceVersion};
use soapberry_zip::ReaderAt as ZipReaderAt;
use soapberry_zip::office::{EntryId, IndexedArchive};
use std::collections::HashMap;
use std::io::Write;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex};

const SOURCE_PUBLICATION_CHUNK_BYTES: usize = 64 * 1024;
const MAX_SOURCE_OVERLAY_PARTS: usize = 64;

struct PendingOverlay {
    target: usize,
    replacement: Vec<u8>,
}

struct ChangedOverlay {
    target: usize,
    replacement: Arc<Vec<u8>>,
}

/// Validation failure returned by [`SourceCacheLimits::new`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SourceCacheLimitError {
    /// A cache needs a positive byte capacity to retain any payload.
    ZeroMaximumBytes,
    /// A cache needs a positive entry capacity to retain any payload.
    ZeroMaximumEntries,
}

impl std::fmt::Display for SourceCacheLimitError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ZeroMaximumBytes => "source cache maximum bytes must be greater than zero",
            Self::ZeroMaximumEntries => "source cache maximum entries must be greater than zero",
        })
    }
}

impl std::error::Error for SourceCacheLimitError {}

/// Finite retention policy for source-backed part payloads.
///
/// Both limits are enforced: a part is retained only when it fits in
/// [`Self::max_bytes`] and the cache can make room below [`Self::max_entries`].
/// Larger requested parts are returned to the caller but are never cached.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceCacheLimits {
    max_bytes: usize,
    max_entries: usize,
}

impl SourceCacheLimits {
    /// Construct a validated finite cache policy.
    ///
    /// # Errors
    ///
    /// Returns an error when either bound is zero.
    pub const fn new(
        max_bytes: usize,
        max_entries: usize,
    ) -> std::result::Result<Self, SourceCacheLimitError> {
        if max_bytes == 0 {
            return Err(SourceCacheLimitError::ZeroMaximumBytes);
        }
        if max_entries == 0 {
            return Err(SourceCacheLimitError::ZeroMaximumEntries);
        }
        Ok(Self {
            max_bytes,
            max_entries,
        })
    }

    /// Maximum total retained payload bytes.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Maximum retained payload entries.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }
}

impl Default for SourceCacheLimits {
    fn default() -> Self {
        // Both values are literal non-zero constants, so this cannot fail.
        Self {
            max_bytes: 8 * 1024 * 1024,
            max_entries: 128,
        }
    }
}

/// Content-free point-in-time diagnostics for a source-backed payload cache.
///
/// Counters are monotonically increasing for the package lifetime and use
/// relaxed atomic updates, so a snapshot is observational rather than a
/// globally linearized transaction. `retained_entries`, `retained_bytes`, and
/// `in_flight_loads` are captured together while the cache takes its existing
/// short lock. No member names, part URIs, or ZIP entry IDs are exposed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SourceCacheDiagnostics {
    /// Requests satisfied directly from a retained payload entry.
    pub hits: u64,
    /// Requests that became the loader for a cold payload read.
    pub cold_loads: u64,
    /// Requests that found an existing same-part cold load and waited for it.
    pub waiter_joins: u64,
    /// Cold reads that completed successfully, whether retained or uncached.
    pub successful_loads: u64,
    /// Cold reads that failed before a payload could be published.
    pub failed_loads: u64,
    /// Retained entries removed to satisfy byte or entry capacity.
    pub evictions: u64,
    /// Successful cold reads returned without retention for any bypass reason.
    pub bypasses: u64,
    /// Successful cold reads returned without retention because their payload
    /// exceeded the configured byte limit.
    pub oversized_bypasses: u64,
    /// Requests or successful loads that could not be coordinated or retained
    /// because cache bookkeeping allocation failed.
    pub allocation_bypasses: u64,
    /// Payload entries currently retained by the cache.
    pub retained_entries: usize,
    /// Payload bytes currently retained by the cache.
    pub retained_bytes: usize,
    /// Same-part cold loads currently coordinated by a flight.
    pub in_flight_loads: usize,
}

#[derive(Debug, Default)]
struct CacheCounters {
    hits: AtomicU64,
    cold_loads: AtomicU64,
    waiter_joins: AtomicU64,
    successful_loads: AtomicU64,
    failed_loads: AtomicU64,
    evictions: AtomicU64,
    bypasses: AtomicU64,
    oversized_bypasses: AtomicU64,
    allocation_bypasses: AtomicU64,
}

#[derive(Clone)]
struct SourceReader {
    snapshot: SourceSnapshot,
}

impl ZipReaderAt for SourceReader {
    fn read_at(&self, output: &mut [u8], offset: u64) -> std::io::Result<usize> {
        self.snapshot.ensure_current_io_if_monitored()?;
        let read = self.snapshot.source.read_at(offset, output)?;
        self.snapshot.ensure_current_io_if_monitored()?;
        Ok(read)
    }
}

#[derive(Clone)]
struct SourceSnapshot {
    source: Arc<dyn ReadAt>,
    version: SourceVersion,
    length: u64,
    monitor_reads: Arc<std::sync::atomic::AtomicBool>,
}

impl SourceSnapshot {
    fn ensure_current(&self) -> Result<()> {
        let actual = self.source.version()?;
        if actual == self.version {
            Ok(())
        } else {
            Err(OpcError::SourceChanged {
                expected: self.version,
                actual,
            })
        }
    }

    fn ensure_current_io_if_monitored(&self) -> std::io::Result<()> {
        if !self.monitor_reads.load(Ordering::Acquire) {
            return Ok(());
        }
        let actual = self.source.version()?;
        if actual == self.version {
            Ok(())
        } else {
            Err(std::io::Error::other(
                "source-backed OPC source changed during publication",
            ))
        }
    }

    fn monitor_publication(&self) {
        self.monitor_reads.store(true, Ordering::Release);
    }
}

struct Counted<'count, W> {
    inner: W,
    written: &'count mut u64,
}

impl<W: Write> Write for Counted<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        *self.written = self.written.saturating_add(written as u64);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

struct SourceCheckedSink<W> {
    inner: W,
    snapshot: SourceSnapshot,
}

impl<W: Write> Write for SourceCheckedSink<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.snapshot.ensure_current_io_if_monitored()?;
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.snapshot.ensure_current_io_if_monitored()?;
        self.inner.flush()
    }
}

struct Chunked<W> {
    inner: W,
}

impl<W: Write> Write for Chunked<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.inner
            .write(&bytes[..bytes.len().min(SOURCE_PUBLICATION_CHUNK_BYTES)])
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

#[derive(Debug)]
struct CatalogPart {
    partname: PackURI,
    content_type: String,
    relationships: Relationships,
    entry_id: EntryId,
}

/// Immutable metadata and deferred payload access for one OPC package part.
pub struct PartView<'package> {
    package: &'package SourceBackedPackage,
    index: usize,
}

impl PartView<'_> {
    /// The part's absolute OPC URI.
    #[must_use]
    pub fn partname(&self) -> &PackURI {
        &self.package.parts[self.index].partname
    }

    /// The content type declared by `[Content_Types].xml`.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.package.parts[self.index].content_type
    }

    /// The part's already-validated relationships.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        &self.package.parts[self.index].relationships
    }

    /// Read this part's payload, using the package's bounded cache when able.
    pub fn data(&self) -> Result<PartData> {
        self.package.read_part(self.index)
    }
}

/// Pinned immutable bytes returned by [`PartView::data`].
#[derive(Clone, Debug)]
pub struct PartData {
    bytes: Arc<Vec<u8>>,
}

impl PartData {
    /// Borrow the part payload.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Share the pinned payload allocation with another owner.
    #[must_use]
    pub fn into_arc(self) -> Arc<Vec<u8>> {
        self.bytes
    }

    /// Return whether both values pin the same payload allocation.
    ///
    /// This compares allocation identity only; equal bytes loaded separately
    /// return `false`.
    #[must_use]
    pub fn shares_allocation_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.bytes, &other.bytes)
    }
}

#[derive(Debug)]
struct CacheEntry {
    bytes: Arc<Vec<u8>>,
    last_used: u64,
}

#[derive(Debug, Default)]
struct CacheState {
    entries: HashMap<EntryId, CacheEntry>,
    flights: HashMap<EntryId, Arc<LoadFlight>>,
    total_bytes: usize,
    clock: u64,
}

#[derive(Debug, Default)]
struct FlightState {
    complete: bool,
    bytes: Option<Arc<Vec<u8>>>,
}

#[derive(Debug, Default)]
struct LoadFlight {
    state: Mutex<FlightState>,
    completed: Condvar,
}

impl LoadFlight {
    fn wait(&self) -> Option<Arc<Vec<u8>>> {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        while !state.complete {
            state = self
                .completed
                .wait(state)
                .unwrap_or_else(|poisoned| poisoned.into_inner());
        }
        state.bytes.as_ref().map(Arc::clone)
    }

    fn finish_success(&self, bytes: Arc<Vec<u8>>) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.bytes = Some(bytes);
        state.complete = true;
        self.completed.notify_all();
    }

    fn finish_failure(&self) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.complete = true;
        self.completed.notify_all();
    }
}

enum CacheAccess {
    Hit(Arc<Vec<u8>>),
    Loader(Arc<LoadFlight>),
    Waiter(Arc<LoadFlight>),
    Bypass,
}

#[derive(Debug)]
struct PartCache {
    limits: SourceCacheLimits,
    state: Mutex<CacheState>,
    counters: CacheCounters,
}

impl PartCache {
    fn new(limits: SourceCacheLimits) -> Self {
        Self {
            limits,
            state: Mutex::new(CacheState::default()),
            counters: CacheCounters::default(),
        }
    }

    fn enter(&self, entry_id: EntryId) -> CacheAccess {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.clock = state.clock.wrapping_add(1);
        let clock = state.clock;
        if let Some(entry) = state.entries.get_mut(&entry_id) {
            entry.last_used = clock;
            self.counters.hits.fetch_add(1, Ordering::Relaxed);
            return CacheAccess::Hit(Arc::clone(&entry.bytes));
        }
        if let Some(flight) = state.flights.get(&entry_id) {
            self.counters.waiter_joins.fetch_add(1, Ordering::Relaxed);
            return CacheAccess::Waiter(Arc::clone(flight));
        }
        if state.flights.try_reserve(1).is_err() {
            self.counters.cold_loads.fetch_add(1, Ordering::Relaxed);
            self.counters
                .allocation_bypasses
                .fetch_add(1, Ordering::Relaxed);
            return CacheAccess::Bypass;
        }
        let flight = Arc::new(LoadFlight::default());
        state.flights.insert(entry_id, Arc::clone(&flight));
        self.counters.cold_loads.fetch_add(1, Ordering::Relaxed);
        CacheAccess::Loader(flight)
    }

    fn complete_success(&self, entry_id: EntryId, flight: &Arc<LoadFlight>, bytes: Arc<Vec<u8>>) {
        let delivered = Arc::clone(&bytes);
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        self.counters
            .successful_loads
            .fetch_add(1, Ordering::Relaxed);
        self.record_retention(self.insert_locked(&mut state, entry_id, bytes));
        // Complete before removing the flight so an oversized, deliberately
        // uncached value still has no gap in which a late peer can start a
        // duplicate load instead of joining this successful delivery.
        flight.finish_success(delivered);
        remove_flight(&mut state, entry_id, flight);
        drop(state);
    }

    fn complete_failure(&self, entry_id: EntryId, flight: &Arc<LoadFlight>) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        self.counters.failed_loads.fetch_add(1, Ordering::Relaxed);
        // Publish failure to current waiters before allowing a new retrying
        // loader to install a replacement flight.
        flight.finish_failure();
        remove_flight(&mut state, entry_id, flight);
        drop(state);
    }

    fn complete_bypass_success(&self, entry_id: EntryId, bytes: Arc<Vec<u8>>) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        self.counters
            .successful_loads
            .fetch_add(1, Ordering::Relaxed);
        self.record_retention(self.insert_locked(&mut state, entry_id, bytes));
    }

    fn complete_bypass_failure(&self) {
        self.counters.failed_loads.fetch_add(1, Ordering::Relaxed);
    }

    fn record_retention(&self, retention: CacheRetention) {
        match retention {
            CacheRetention::Retained => {},
            CacheRetention::Oversized => {
                self.counters.bypasses.fetch_add(1, Ordering::Relaxed);
                self.counters
                    .oversized_bypasses
                    .fetch_add(1, Ordering::Relaxed);
            },
            CacheRetention::AllocationFailure => {
                self.counters.bypasses.fetch_add(1, Ordering::Relaxed);
                self.counters
                    .allocation_bypasses
                    .fetch_add(1, Ordering::Relaxed);
            },
        }
    }

    fn insert_locked(
        &self,
        state: &mut CacheState,
        entry_id: EntryId,
        bytes: Arc<Vec<u8>>,
    ) -> CacheRetention {
        let weight = bytes.len();
        if weight > self.limits.max_bytes {
            return CacheRetention::Oversized;
        }
        state.clock = state.clock.wrapping_add(1);
        let clock = state.clock;
        while state.entries.len() >= self.limits.max_entries
            || state.total_bytes.saturating_add(weight) > self.limits.max_bytes
        {
            let Some((&oldest, _)) = state
                .entries
                .iter()
                .min_by_key(|(_, entry)| entry.last_used)
            else {
                break;
            };
            if let Some(removed) = state.entries.remove(&oldest) {
                state.total_bytes = state.total_bytes.saturating_sub(removed.bytes.len());
                self.counters.evictions.fetch_add(1, Ordering::Relaxed);
            }
        }
        if state.entries.try_reserve(1).is_err() {
            return CacheRetention::AllocationFailure;
        }
        if let Some(previous) = state.entries.insert(
            entry_id,
            CacheEntry {
                bytes,
                last_used: clock,
            },
        ) {
            state.total_bytes = state.total_bytes.saturating_sub(previous.bytes.len());
        }
        state.total_bytes = state.total_bytes.saturating_add(weight);
        CacheRetention::Retained
    }

    fn diagnostics(&self) -> SourceCacheDiagnostics {
        let state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        SourceCacheDiagnostics {
            hits: self.counters.hits.load(Ordering::Relaxed),
            cold_loads: self.counters.cold_loads.load(Ordering::Relaxed),
            waiter_joins: self.counters.waiter_joins.load(Ordering::Relaxed),
            successful_loads: self.counters.successful_loads.load(Ordering::Relaxed),
            failed_loads: self.counters.failed_loads.load(Ordering::Relaxed),
            evictions: self.counters.evictions.load(Ordering::Relaxed),
            bypasses: self.counters.bypasses.load(Ordering::Relaxed),
            oversized_bypasses: self.counters.oversized_bypasses.load(Ordering::Relaxed),
            allocation_bypasses: self.counters.allocation_bypasses.load(Ordering::Relaxed),
            retained_entries: state.entries.len(),
            retained_bytes: state.total_bytes,
            in_flight_loads: state.flights.len(),
        }
    }
}

#[derive(Clone, Copy)]
enum CacheRetention {
    Retained,
    Oversized,
    AllocationFailure,
}

fn remove_flight(state: &mut CacheState, entry_id: EntryId, flight: &Arc<LoadFlight>) {
    if state
        .flights
        .get(&entry_id)
        .is_some_and(|current| Arc::ptr_eq(current, flight))
    {
        state.flights.remove(&entry_id);
    }
}

/// A structurally validated OPC package backed by an immutable positional source.
///
/// Opening reads and validates ZIP metadata, content types, and relationship
/// XML, but never reads ordinary part payloads. The ordinary view is immutable.
/// [`Self::write_part_overlays_to_stream`] is a narrow, consuming publisher for
/// a bounded same-topology Part replacement set that raw-copies every other
/// ZIP member; call [`Self::into_opc_package`] when a general owning mutable
/// package is needed.
pub struct SourceBackedPackage {
    source: SourceSnapshot,
    archive: IndexedArchive<SourceReader>,
    limits: ReadLimits,
    package_relationships: Relationships,
    parts: Vec<CatalogPart>,
    parts_by_name: HashMap<PackURI, usize>,
    non_part_members: Vec<NonPartMember>,
    cache: PartCache,
}

impl SourceBackedPackage {
    /// Open a source-backed package with the standard bounded read policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(
            source,
            ReadLimits::default(),
            SourceCacheLimits::default(),
        )
    }

    /// Open a source-backed package with an explicit bounded read policy.
    ///
    /// The source version is captured before indexing and checked after every
    /// mandatory open read.  A changed source is never silently accepted.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(
            source,
            limits,
            SourceCacheLimits::default(),
        )
    }

    /// Open a source-backed package with an explicit payload-cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(source, ReadLimits::default(), cache_limits)
    }

    /// Open a source-backed package with explicit read and cache policies.
    ///
    /// The source version is captured before indexing and checked after every
    /// mandatory open read. A changed source is never silently accepted.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        let version = source.version()?;
        let length = source.len()?;
        limits.check(ReadResource::InputBytes, length, limits.max_input_bytes())?;
        let snapshot = SourceSnapshot {
            source: Arc::clone(&source),
            version,
            length,
            monitor_reads: Arc::new(std::sync::atomic::AtomicBool::new(false)),
        };
        snapshot.ensure_current()?;
        let archive = IndexedArchive::from_reader_with_limits(
            SourceReader {
                snapshot: snapshot.clone(),
            },
            length,
            limits.zip_limits(),
        )?;
        snapshot.ensure_current()?;
        let SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
        } = PackageReader::source_catalog(&archive, limits)?;
        snapshot.ensure_current()?;

        let package_relationships = relationships_for_package(pkg_srels)?;
        let mut catalog_parts = Vec::new();
        catalog_parts
            .try_reserve_exact(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC catalog parts",
                source,
            })?;
        let mut parts_by_name = HashMap::new();
        parts_by_name
            .try_reserve(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC part lookup",
                source,
            })?;
        for (index, part) in parts.into_iter().enumerate() {
            let relationships = relationships_for_part(&part.partname, part.srels)?;
            let entry_id = archive
                .entry_id(part.partname.membername())
                .ok_or_else(|| OpcError::PartNotFound(part.partname.to_string()))?;
            parts_by_name.insert(part.partname.clone(), index);
            catalog_parts.push(CatalogPart {
                partname: part.partname,
                content_type: part.content_type,
                relationships,
                entry_id,
            });
        }

        Ok(Self {
            source: snapshot,
            archive,
            limits,
            package_relationships,
            parts: catalog_parts,
            parts_by_name,
            non_part_members,
            cache: PartCache::new(cache_limits),
        })
    }

    /// Package-level relationships parsed during opening.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        &self.package_relationships
    }

    /// Return metadata-only views of every ordinary part.
    pub fn iter_parts(&self) -> impl Iterator<Item = PartView<'_>> {
        (0..self.parts.len()).map(|index| PartView {
            package: self,
            index,
        })
    }

    /// Look up one ordinary part without reading its payload.
    pub fn part(&self, partname: &PackURI) -> Result<PartView<'_>> {
        self.source.ensure_current()?;
        self.parts_by_name
            .get(partname)
            .copied()
            .map(|index| PartView {
                package: self,
                index,
            })
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Return the unique main document part without reading its payload.
    pub fn main_document_part(&self) -> Result<PartView<'_>> {
        self.source.ensure_current()?;
        let mut matching = self.package_relationships.iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::OFFICE_DOCUMENT | relationship_type::STRICT_OFFICE_DOCUMENT
            )
        });
        let relationship = matching.next().ok_or_else(|| {
            OpcError::InvalidRelationship("main-document relationship is missing".to_string())
        })?;
        if matching.next().is_some() {
            return Err(OpcError::InvalidRelationship(
                "package has multiple main-document relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(OpcError::InvalidRelationship(
                "main-document relationship cannot be external".to_string(),
            ));
        }
        let partname = relationship.target_partname()?;
        self.part(&partname)
    }

    /// ZIP items present in the source but not modelled as OPC parts.
    #[must_use]
    pub fn non_part_members(&self) -> &[NonPartMember] {
        &self.non_part_members
    }

    /// Return content-free payload-cache activity and current occupancy.
    ///
    /// See [`SourceCacheDiagnostics`] for the precise event definitions. This
    /// operation does not read part payloads or expose member identifiers.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        self.cache.diagnostics()
    }

    /// Fully materialize this immutable view into the existing mutable package type.
    pub fn into_opc_package(self) -> Result<OpcPackage> {
        self.source.ensure_current()?;
        let mut package = OpcPackage::new();
        copy_relationships(&self.package_relationships, package.rels_mut())?;
        for index in 0..self.parts.len() {
            let bytes = self.read_part(index)?;
            let catalog_part = &self.parts[index];
            let mut part = PartFactory::load(
                catalog_part.partname.clone(),
                catalog_part.content_type.clone(),
                bytes.as_bytes().to_vec(),
            )?;
            copy_relationships(&catalog_part.relationships, part.rels_mut())?;
            package.try_add_part(part)?;
        }
        package.set_non_part_members(self.non_part_members);
        Ok(package)
    }

    /// Replace one existing ordinary Part and publish to a sequential stream.
    ///
    /// This is an explicit low-level OPC operation. The Part URI, content
    /// type, relationships, package catalog, and physical member topology are
    /// immutable; only the selected payload may change. Every other ZIP member
    /// is raw-copied from the positional source. Unsupported physical layouts
    /// are refused before output instead of silently materializing the package.
    ///
    /// An exact payload no-op copies the complete source artifact byte for
    /// byte, including signatures and unsupported physical details. A real
    /// change to a signed package is refused because this operation accepts no
    /// signature-stripping or resigning policy.
    ///
    /// # Errors
    ///
    /// Returns a typed source, limit, Part, signature, XML-publication, ZIP, or
    /// sink error. If a non-atomic sink accepts bytes before failing, the error
    /// is [`OpcError::IncompleteOutput`].
    pub fn write_part_overlay_to_stream<W: Write>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: Vec<u8>,
    ) -> Result<()> {
        self.write_part_overlays_to_stream(writer, vec![(partname.clone(), replacement)])
    }

    /// Replace a bounded set of existing ordinary Parts and publish to a
    /// sequential stream.
    ///
    /// The replacement set is sorted and checked for duplicate Part URIs. Its
    /// maximum size is 64. Part URIs, content types, relationships, the package
    /// catalog, and physical member topology are immutable. Every unselected
    /// ZIP member and every selected exact no-op member is raw-copied.
    ///
    /// If every replacement is byte-identical to its source payload, the
    /// complete source artifact is copied byte for byte, including signatures
    /// and unsupported physical details. A real change to a signed package is
    /// refused because this operation accepts no signature policy.
    ///
    /// All selected payloads, aggregate limits, signatures, changed XML, and
    /// the preservation plan are validated before the first output byte.
    ///
    /// # Errors
    ///
    /// Returns a typed source, limit, duplicate-Part, Part, signature,
    /// XML-publication, ZIP, or sink error. If a non-atomic sink accepts bytes
    /// before failing, the error is [`OpcError::IncompleteOutput`].
    pub fn write_part_overlays_to_stream<W: Write>(
        self,
        writer: W,
        mut replacements: Vec<(PackURI, Vec<u8>)>,
    ) -> Result<()> {
        if replacements.len() > MAX_SOURCE_OVERLAY_PARTS {
            return Err(overlay_unavailable(format!(
                "replacement set exceeds the {MAX_SOURCE_OVERLAY_PARTS}-Part bound"
            )));
        }
        if replacements.is_empty() {
            return self.write_exact_source(writer);
        }
        replacements.sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
        if let Some(duplicate) = replacements.windows(2).find(|pair| pair[0].0 == pair[1].0) {
            return Err(OpcError::DuplicatePartName(duplicate[0].0.to_string()));
        }

        let mut overlays = Vec::new();
        overlays
            .try_reserve_exact(replacements.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC replacement plan",
                source,
            })?;
        for (partname, replacement) in replacements {
            let target = self
                .parts_by_name
                .get(&partname)
                .copied()
                .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
            overlays.push(PendingOverlay {
                target,
                replacement,
            });
        }
        self.validate_overlay_limits(&overlays)?;

        let mut changed = Vec::new();
        changed
            .try_reserve_exact(overlays.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC changed replacement plan",
                source,
            })?;
        for overlay in overlays {
            // Reading every selected closure proves its local framing,
            // compression, declared size, and CRC before output.
            let original = self.read_part(overlay.target)?;
            if original.as_bytes() != overlay.replacement.as_slice() {
                changed.push((overlay, original));
            }
        }
        if changed.is_empty() {
            return self.write_exact_source(writer);
        }
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(changed.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC changed payloads",
                source,
            })?;
        for (overlay, original) in changed {
            let target_part = &self.parts[overlay.target];
            if xml_minifier::audit::package::is_xml_part(
                target_part.partname.as_str(),
                &target_part.content_type,
            ) {
                validate_overlay_xml(target_part.partname.as_str(), original.as_bytes())?;
                validate_overlay_xml(target_part.partname.as_str(), &overlay.replacement)?;
            }
            replacements.push(ChangedOverlay {
                target: overlay.target,
                replacement: Arc::new(overlay.replacement),
            });
        }
        self.write_changed_overlays(writer, &replacements)
    }

    fn validate_overlay_limits(&self, overlays: &[PendingOverlay]) -> Result<()> {
        for overlay in overlays {
            let replacement_bytes =
                u64::try_from(overlay.replacement.len()).map_err(|_| OpcError::ReadLimit {
                    resource: ReadResource::PartBytes,
                    actual: u64::MAX,
                    maximum: self.limits.max_part_bytes(),
                })?;
            self.limits.check(
                ReadResource::PartBytes,
                replacement_bytes,
                self.limits.max_part_bytes(),
            )?;
            self.limits.check(
                ReadResource::ArchiveEntryBytes,
                replacement_bytes,
                self.limits.max_archive_entry_bytes(),
            )?;
        }

        let mut part_total = 0_u64;
        let mut archive_total = 0_u64;
        for part in &self.parts {
            let bytes = self
                .archive
                .metadata_for(part.entry_id)?
                .uncompressed_size();
            part_total = checked_overlay_total(
                part_total,
                bytes,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
        }
        for name in self.archive.file_names() {
            archive_total = checked_overlay_total(
                archive_total,
                self.archive.metadata(name)?.uncompressed_size(),
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        let mut adjusted_parts = part_total;
        let mut adjusted_archive = archive_total;
        for overlay in overlays {
            let target_bytes = self
                .archive
                .metadata_for(self.parts[overlay.target].entry_id)?
                .uncompressed_size();
            let replacement_bytes = overlay.replacement.len() as u64;
            adjusted_parts = adjusted_overlay_total(
                adjusted_parts,
                target_bytes,
                replacement_bytes,
                ReadResource::TotalPartBytes,
                self.limits.max_total_part_bytes(),
            )?;
            adjusted_archive = adjusted_overlay_total(
                adjusted_archive,
                target_bytes,
                replacement_bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
        }
        self.limits.check(
            ReadResource::TotalPartBytes,
            adjusted_parts,
            self.limits.max_total_part_bytes(),
        )?;
        self.limits.check(
            ReadResource::ArchiveTotalBytes,
            adjusted_archive,
            self.limits.max_archive_total_bytes(),
        )?;
        Ok(())
    }

    fn has_signature_infrastructure(&self) -> bool {
        self.package_relationships
            .iter()
            .any(|relationship| is_signature_relationship(relationship.reltype()))
            || self.parts.iter().any(|part| {
                is_signature_path(part.partname.as_str())
                    || is_signature_content_type(&part.content_type)
                    || part
                        .relationships
                        .iter()
                        .any(|relationship| is_signature_relationship(relationship.reltype()))
            })
    }

    fn write_exact_source<W: Write>(self, writer: W) -> Result<()> {
        self.source.monitor_publication();
        self.source.ensure_current()?;
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(SOURCE_PUBLICATION_CHUNK_BYTES)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC publication buffer",
                source,
            })?;
        buffer.resize(SOURCE_PUBLICATION_CHUNK_BYTES, 0);
        let mut written = 0_u64;
        let result = {
            let counted = Counted {
                inner: writer,
                written: &mut written,
            };
            let mut sink = SourceCheckedSink {
                inner: counted,
                snapshot: self.source.clone(),
            };
            let mut offset = 0_u64;
            (|| {
                while offset < self.source.length {
                    let remaining =
                        usize::try_from((self.source.length - offset).min(buffer.len() as u64))
                            .map_err(|_| {
                                overlay_unavailable("source range does not fit this platform")
                            })?;
                    let read = self
                        .source
                        .source
                        .read_at(offset, &mut buffer[..remaining])?;
                    if read == 0 {
                        return Err(OpcError::IoError(std::io::Error::new(
                            std::io::ErrorKind::UnexpectedEof,
                            "source-backed OPC source ended during publication",
                        )));
                    }
                    self.source.ensure_current()?;
                    sink.write_all(&buffer[..read])?;
                    offset = offset
                        .checked_add(read as u64)
                        .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
                }
                sink.flush()?;
                Ok(())
            })()
        };
        finish_source_publication(result, &self.source, written)
    }

    fn write_changed_overlays<W: Write>(
        self,
        writer: W,
        replacements: &[ChangedOverlay],
    ) -> Result<()> {
        self.source.monitor_publication();
        self.source.ensure_current()?;
        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC preservation index",
                source,
            })?;
        scratch.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let index = match self.archive.preservation_index(&mut scratch) {
            Ok(index) => index,
            Err(error) => {
                self.source.ensure_current()?;
                return Err(overlay_unavailable(error.to_string()));
            },
        };
        self.source.ensure_current()?;

        let mut replacement_bytes = 0_u64;
        for replacement in replacements {
            let target_name = self.parts[replacement.target].partname.membername();
            let target_entries = index
                .entries()
                .iter()
                .filter(|entry| entry.raw_name_bytes() == target_name.as_bytes())
                .count();
            if target_entries != 1 {
                return Err(overlay_unavailable(
                    "selected Part does not have one canonical UTF-8 source member",
                ));
            }
            replacement_bytes = replacement_bytes
                .checked_add(replacement.replacement.len() as u64)
                .ok_or_else(|| overlay_unavailable("replacement byte total overflows u64"))?;
        }
        let conservative_output_bound = self
            .source
            .length
            .checked_add(replacement_bytes.saturating_mul(2))
            .and_then(|bytes| bytes.checked_add(SOURCE_PUBLICATION_CHUNK_BYTES as u64));
        if conservative_output_bound.is_none_or(|bytes| bytes > u64::from(u32::MAX)) {
            return Err(overlay_unavailable(
                "selected Part replacement may require ZIP64 output",
            ));
        }

        let mut plan = soapberry_zip::PreservationPlan::new();
        for entry in index.entries() {
            if let Some(replacement) = replacements.iter().find(|replacement| {
                entry.raw_name_bytes()
                    == self.parts[replacement.target]
                        .partname
                        .membername()
                        .as_bytes()
            }) {
                let target_name = self.parts[replacement.target].partname.membername();
                plan.push(soapberry_zip::PreservationAction::Regenerate {
                    id: entry.id(),
                    entry: soapberry_zip::RegeneratedEntry::new_shared(
                        target_name,
                        Arc::clone(&replacement.replacement),
                    )
                    .compression_method(soapberry_zip::CompressionMethod::Deflate),
                });
            } else {
                plan.push(soapberry_zip::PreservationAction::Copy(entry.id()));
            }
        }

        self.source.ensure_current()?;
        let mut written = 0_u64;
        let result = {
            let counted = Counted {
                inner: writer,
                written: &mut written,
            };
            let checked = SourceCheckedSink {
                inner: counted,
                snapshot: self.source.clone(),
            };
            let result = index.write_to(&plan, Chunked { inner: checked });
            match result {
                Ok(mut sink) => sink.flush().map_err(OpcError::IoError),
                Err(error) => Err(OpcError::ZipError(error.to_string())),
            }
        };
        finish_source_publication(result, &self.source, written)
    }

    fn read_part(&self, index: usize) -> Result<PartData> {
        let entry_id = self
            .parts
            .get(index)
            .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?
            .entry_id;
        loop {
            self.source.ensure_current()?;
            match self.cache.enter(entry_id) {
                CacheAccess::Hit(bytes) => {
                    self.source.ensure_current()?;
                    return Ok(PartData { bytes });
                },
                CacheAccess::Waiter(flight) => {
                    if let Some(bytes) = flight.wait() {
                        self.source.ensure_current()?;
                        return Ok(PartData { bytes });
                    }
                    // The loader may have failed; in that case the flight is
                    // removed and this caller retries rather than observing a
                    // retained error. This also re-checks source freshness.
                },
                CacheAccess::Loader(flight) => {
                    return self.load_part(index, entry_id, Some(flight));
                },
                CacheAccess::Bypass => return self.load_part(index, entry_id, None),
            }
        }
    }

    fn load_part(
        &self,
        index: usize,
        entry_id: EntryId,
        flight: Option<Arc<LoadFlight>>,
    ) -> Result<PartData> {
        let result = (|| {
            let part = self
                .parts
                .get(index)
                .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?;
            let bytes = self.archive.read_entry(part.entry_id)?;
            self.source.ensure_current()?;
            self.limits.check(
                ReadResource::PartBytes,
                bytes.len() as u64,
                self.limits.max_part_bytes(),
            )?;
            // Check immediately before publishing. If the source changed
            // during the cold read, no stale payload enters the cache.
            self.source.ensure_current()?;
            Ok(Arc::new(bytes))
        })();
        match (flight, result) {
            (Some(flight), Ok(bytes)) => {
                self.cache
                    .complete_success(entry_id, &flight, Arc::clone(&bytes));
                Ok(PartData { bytes })
            },
            (Some(flight), Err(error)) => {
                self.cache.complete_failure(entry_id, &flight);
                Err(error)
            },
            (None, Ok(bytes)) => {
                self.cache
                    .complete_bypass_success(entry_id, Arc::clone(&bytes));
                Ok(PartData { bytes })
            },
            (None, Err(error)) => {
                self.cache.complete_bypass_failure();
                Err(error)
            },
        }
    }
}

fn relationships_for_package(
    serialized: impl IntoIterator<Item = SerializedRelationship>,
) -> Result<Relationships> {
    let mut relationships = Relationships::new(PACKAGE_URI.to_string());
    for relationship in serialized {
        relationships.try_add_relationship(
            relationship.reltype,
            relationship.target_ref,
            relationship.r_id,
            relationship.target_mode,
        )?;
    }
    Ok(relationships)
}

fn relationships_for_part(
    partname: &PackURI,
    serialized: impl IntoIterator<Item = SerializedRelationship>,
) -> Result<Relationships> {
    let mut relationships = Relationships::for_source(partname);
    for relationship in serialized {
        relationships.try_add_relationship(
            relationship.reltype,
            relationship.target_ref,
            relationship.r_id,
            relationship.target_mode,
        )?;
    }
    Ok(relationships)
}

fn copy_relationships(from: &Relationships, to: &mut Relationships) -> Result<()> {
    for relationship in from.iter() {
        to.try_add_relationship(
            relationship.reltype().to_string(),
            relationship.target_ref().to_string(),
            relationship.r_id().to_string(),
            relationship.target_mode(),
        )?;
    }
    Ok(())
}

fn validate_overlay_xml(part: &str, bytes: &[u8]) -> Result<()> {
    xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default())
        .map(|_report| ())
        .map_err(|source| OpcError::XmlPublication {
            part: part.to_string(),
            source,
        })
}

fn checked_overlay_total(
    current: u64,
    bytes: u64,
    resource: ReadResource,
    maximum: u64,
) -> Result<u64> {
    current.checked_add(bytes).ok_or(OpcError::ReadLimit {
        resource,
        actual: u64::MAX,
        maximum,
    })
}

fn adjusted_overlay_total(
    current: u64,
    removed: u64,
    added: u64,
    resource: ReadResource,
    maximum: u64,
) -> Result<u64> {
    current
        .checked_sub(removed)
        .and_then(|remaining| remaining.checked_add(added))
        .ok_or(OpcError::ReadLimit {
            resource,
            actual: u64::MAX,
            maximum,
        })
}

fn overlay_unavailable(reason: impl Into<String>) -> OpcError {
    OpcError::SourceBackedOverlayUnavailable {
        reason: reason.into(),
    }
}

fn finish_source_publication(
    result: Result<()>,
    source: &SourceSnapshot,
    written: u64,
) -> Result<()> {
    let freshness = source.ensure_current();
    let result = match (result, freshness) {
        (_, Err(error @ OpcError::SourceChanged { .. })) => Err(error),
        (Err(error), _) => Err(error),
        (Ok(()), Err(error)) => Err(error),
        (Ok(()), Ok(())) => Ok(()),
    };
    match result {
        Err(source) if written != 0 => Err(OpcError::IncompleteOutput {
            written,
            source: Box::new(source),
        }),
        other => other,
    }
}

fn is_signature_relationship(kind: &str) -> bool {
    matches!(
        kind,
        relationship_type::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_signature_path(path: &str) -> bool {
    const DIRECTORY: &[u8] = b"/_xmlsignatures/";
    path.as_bytes()
        .get(..DIRECTORY.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(DIRECTORY))
}

fn is_signature_content_type(value: &str) -> bool {
    matches!(
        value,
        content_type::OPC_DIGITAL_SIGNATURE_ORIGIN
            | content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
            | content_type::OPC_DIGITAL_SIGNATURE_CERTIFICATE
    )
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic on failure by design"
    )]

    use super::*;
    use std::collections::HashMap;
    use std::sync::Barrier;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};
    use std::time::Duration;

    struct CountingSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        reads: AtomicUsize,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                reads: AtomicUsize::new(0),
            }
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(42, self.revision.load(Ordering::SeqCst)))
        }
    }

    struct SlowPayloadSource {
        bytes: Vec<u8>,
        payload_offset: usize,
        payload_reads: AtomicUsize,
    }

    impl SlowPayloadSource {
        fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
            Self {
                bytes,
                payload_offset,
                payload_reads: AtomicUsize::new(0),
            }
        }
    }

    impl ReadAt for SlowPayloadSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset == self.payload_offset {
                self.payload_reads.fetch_add(1, Ordering::SeqCst);
                // Keep the cold load in flight long enough for the peer to
                // enter the same part concurrently.
                std::thread::sleep(Duration::from_millis(100));
            }
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(77, 0))
        }
    }

    fn archive_bytes(root_relationships: &[u8], document: &[u8], include_junk: bool) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships)
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<orphan/>")
            .unwrap();
        if include_junk {
            writer.write_stored("scratch.bin", b"not a part").unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn root_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
    }

    fn signed_root_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="signature/origin.xml"/></Relationships>"#
    }

    fn signed_archive(document: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", signed_root_relationships())
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer
            .write_stored("signature/origin.xml", b"<origin/>")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[derive(Debug)]
    struct RawRecord {
        local: Vec<u8>,
        central: Vec<u8>,
    }

    fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
        let archive = soapberry_zip::ZipArchive::from_slice(data)
            .unwrap()
            .into_zip_archive();
        let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
        let index = soapberry_zip::PreservationIndex::new(&archive, &mut scratch).unwrap();
        index
            .entries()
            .iter()
            .map(|entry| {
                let local = entry.local_span();
                let central = entry.central_record();
                (
                    entry.raw_name_bytes().to_vec(),
                    RawRecord {
                        local: data[local.start as usize..local.end as usize].to_vec(),
                        central: data[central.start as usize..central.end as usize].to_vec(),
                    },
                )
            })
            .collect()
    }

    fn central_without_local_offset(bytes: &[u8]) -> Vec<u8> {
        let mut bytes = bytes.to_vec();
        bytes[42..46].fill(0);
        bytes
    }

    struct MutatingSink {
        source: Arc<CountingSource>,
        bytes: Vec<u8>,
        change_after: usize,
        changed: bool,
    }

    impl Write for MutatingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            if !self.changed && self.bytes.len() >= self.change_after {
                self.source.changed();
                self.changed = true;
            }
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    struct BoundedFailingSink {
        accepted: usize,
        limit: usize,
        largest_write: usize,
    }

    impl Write for BoundedFailingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.largest_write = self.largest_write.max(bytes.len());
            let remaining = self.limit.saturating_sub(self.accepted);
            if remaining == 0 {
                return Err(std::io::Error::other("injected sink failure"));
            }
            let written = remaining.min(bytes.len());
            self.accepted += written;
            Ok(written)
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn mandatory_xml_is_opened_but_ordinary_payload_corruption_is_deferred() {
        let malformed = archive_bytes(b"<Relationships", b"document", false);
        let malformed_source = Arc::new(CountingSource::new(malformed));
        assert!(matches!(
            SourceBackedPackage::from_read_at(malformed_source),
            Err(OpcError::QuickXmlError(_))
        ));

        const DOCUMENT: &[u8] = b"source-backed deferred corruption";
        let mut corrupt = archive_bytes(root_relationships(), DOCUMENT, false);
        let position = corrupt
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        corrupt[position] ^= 0xff;
        let source = Arc::new(CountingSource::new(corrupt));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let main = package.main_document_part().unwrap();
        assert!(matches!(main.data(), Err(OpcError::ZipError(_))));
    }

    #[test]
    fn cache_hits_pin_payloads_and_failures_are_not_retained() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"cached payload",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let part = package.main_document_part().unwrap();
        let first = part.data().unwrap();
        let after_first = source.reads.load(Ordering::SeqCst);
        let second = part.data().unwrap();
        assert_eq!(source.reads.load(Ordering::SeqCst), after_first);
        assert!(first.shares_allocation_with(&second));
        assert_eq!(
            package.cache_diagnostics(),
            SourceCacheDiagnostics {
                hits: 1,
                cold_loads: 1,
                successful_loads: 1,
                retained_entries: 1,
                retained_bytes: b"cached payload".len(),
                ..SourceCacheDiagnostics::default()
            }
        );

        const DOCUMENT: &[u8] = b"never cache a failed read";
        let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let position = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        bytes[position] ^= 0xff;
        let corrupt_source = Arc::new(CountingSource::new(bytes));
        let corrupt_package = SourceBackedPackage::from_read_at(corrupt_source.clone()).unwrap();
        let corrupt_part = corrupt_package.main_document_part().unwrap();
        assert!(corrupt_part.data().is_err());
        let after_failure = corrupt_source.reads.load(Ordering::SeqCst);
        assert!(corrupt_part.data().is_err());
        assert!(corrupt_source.reads.load(Ordering::SeqCst) > after_failure);
        let diagnostics = corrupt_package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 2);
        assert_eq!(diagnostics.failed_loads, 2);
        assert_eq!(diagnostics.retained_entries, 0);
    }

    #[test]
    fn concurrent_cold_reads_share_one_archive_load_and_one_arc() {
        const DOCUMENT: &[u8] = b"single-flight source-backed payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let start = Arc::new(Barrier::new(3));
        let (first, second) = std::thread::scope(|scope| {
            let package = &package;
            let first_start = Arc::clone(&start);
            let first_task = scope.spawn(move || {
                first_start.wait();
                package
                    .main_document_part()
                    .unwrap()
                    .data()
                    .unwrap()
                    .into_arc()
            });
            let second_start = Arc::clone(&start);
            let second_task = scope.spawn(move || {
                second_start.wait();
                package
                    .main_document_part()
                    .unwrap()
                    .data()
                    .unwrap()
                    .into_arc()
            });
            start.wait();
            std::thread::sleep(Duration::from_millis(10));
            assert_eq!(package.cache_diagnostics().in_flight_loads, 1);
            (first_task.join().unwrap(), second_task.join().unwrap())
        });
        assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
        assert!(Arc::ptr_eq(&first, &second));
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 1);
        assert_eq!(diagnostics.waiter_joins, 1);
        assert_eq!(diagnostics.successful_loads, 1);
        assert_eq!(diagnostics.in_flight_loads, 0);
    }

    #[test]
    fn cache_evicts_by_byte_weight_and_entry_count_and_rejects_oversized_values() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"document", false),
        )))
        .unwrap();
        let first_id = package.parts[0].entry_id;
        let second_id = package.parts[1].entry_id;
        let cache = PartCache::new(SourceCacheLimits::new(3, 3).unwrap());
        let first = Arc::new(vec![1, 2]);
        cache.complete_bypass_success(first_id, Arc::clone(&first));
        assert!(Arc::ptr_eq(
            &cache.state.lock().unwrap().entries[&first_id].bytes,
            &first
        ));
        cache.complete_bypass_success(second_id, Arc::new(vec![3, 4]));
        assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
        assert!(cache.state.lock().unwrap().entries.contains_key(&second_id));

        let entry_limited = PartCache::new(SourceCacheLimits::new(10, 1).unwrap());
        entry_limited.complete_bypass_success(first_id, Arc::new(vec![1, 2]));
        entry_limited.complete_bypass_success(second_id, Arc::new(vec![3, 4]));
        assert!(
            !entry_limited
                .state
                .lock()
                .unwrap()
                .entries
                .contains_key(&first_id)
        );
        assert!(
            entry_limited
                .state
                .lock()
                .unwrap()
                .entries
                .contains_key(&second_id)
        );

        cache.complete_bypass_success(first_id, Arc::new(vec![0, 0, 0, 0]));
        assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
        assert_eq!(cache.diagnostics().evictions, 1);
        assert_eq!(cache.diagnostics().oversized_bypasses, 1);
    }

    #[test]
    fn cache_limits_reject_zero_bounds() {
        assert_eq!(
            SourceCacheLimits::new(0, 1),
            Err(SourceCacheLimitError::ZeroMaximumBytes)
        );
        assert_eq!(
            SourceCacheLimits::new(1, 0),
            Err(SourceCacheLimitError::ZeroMaximumEntries)
        );
    }

    #[test]
    fn catalog_entry_id_resolves_to_the_same_payload_as_its_member_name() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"entry identity", false),
        )))
        .unwrap();
        let part = package.main_document_part().unwrap();
        let catalog = package
            .parts
            .iter()
            .find(|catalog| catalog.partname == *part.partname())
            .unwrap();
        assert_eq!(
            package.archive.read_entry(catalog.entry_id).unwrap(),
            package.archive.read(catalog.partname.membername()).unwrap()
        );
    }

    #[test]
    fn source_changes_reject_metadata_cache_and_conversion_access() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"stable payload",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let part = package.main_document_part().unwrap();
        part.data().unwrap();
        source.changed();
        assert!(matches!(
            package.part(part.partname()),
            Err(OpcError::SourceChanged { .. })
        ));
        assert!(matches!(part.data(), Err(OpcError::SourceChanged { .. })));
        assert!(matches!(
            package.into_opc_package(),
            Err(OpcError::SourceChanged { .. })
        ));
    }

    #[test]
    fn catalog_reports_non_parts_and_conversion_matches_loaded_parts() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        assert_eq!(package.iter_parts().count(), 2);
        assert_eq!(package.non_part_members().len(), 1);
        assert_eq!(package.non_part_members()[0].name(), "scratch.bin");
        assert_eq!(
            package
                .main_document_part()
                .unwrap()
                .data()
                .unwrap()
                .as_bytes(),
            b"document"
        );
        let owned = package.into_opc_package().unwrap();
        assert_eq!(owned.part_count(), 2);
        assert_eq!(owned.non_part_members().len(), 1);
        assert_eq!(owned.main_document_part().unwrap().blob(), b"document");
    }

    #[test]
    fn one_part_overlay_raw_copies_every_unselected_member_and_reopens() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let source_raw = raw_records(&source_bytes);
        let source = Arc::new(CountingSource::new(source_bytes));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert_eq!(reopened.get_part(&target).unwrap().blob(), b"<after/>");
        assert_eq!(reopened.part_count(), 2);
        assert_eq!(reopened.non_part_members().len(), 1);
        assert_eq!(reopened.non_part_members()[0].name(), "scratch.bin");
        let output_raw = raw_records(&output);
        assert_eq!(output_raw.len(), source_raw.len());
        for (name, source_record) in source_raw {
            if name == b"word/document.xml" {
                assert_ne!(output_raw[&name].local, source_record.local);
            } else {
                assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
                assert_eq!(
                    central_without_local_offset(&output_raw[&name].central),
                    central_without_local_offset(&source_record.central),
                    "{name:?}"
                );
            }
        }
    }

    #[test]
    fn multi_part_overlay_changes_only_selected_raw_members_and_reopens() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let source_raw = raw_records(&source_bytes);
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes))).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let orphan = PackURI::new("/custom/orphan.xml").unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlays_to_stream(
                &mut output,
                vec![
                    (orphan.clone(), b"<orphan-after/>".to_vec()),
                    (document.clone(), b"<document-after/>".to_vec()),
                ],
            )
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert_eq!(
            reopened.get_part(&document).unwrap().blob(),
            b"<document-after/>"
        );
        assert_eq!(
            reopened.get_part(&orphan).unwrap().blob(),
            b"<orphan-after/>"
        );
        let output_raw = raw_records(&output);
        assert_eq!(output_raw.len(), source_raw.len());
        for (name, source_record) in source_raw {
            if matches!(name.as_slice(), b"word/document.xml" | b"custom/orphan.xml") {
                assert_ne!(output_raw[&name].local, source_record.local, "{name:?}");
            } else {
                assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
                assert_eq!(
                    central_without_local_offset(&output_raw[&name].central),
                    central_without_local_offset(&source_record.central),
                    "{name:?}"
                );
            }
        }
    }

    #[test]
    fn multi_part_overlay_checks_set_bounds_duplicates_and_aggregate_limits() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
        let document = PackURI::new("/word/document.xml").unwrap();
        let orphan = PackURI::new("/custom/orphan.xml").unwrap();

        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document.clone(), b"<first/>".to_vec()),
                    (document.clone(), b"<second/>".to_vec()),
                ],
            ),
            Err(OpcError::DuplicatePartName(_))
        ));
        assert!(output.is_empty());

        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let oversized_set = (0..=MAX_SOURCE_OVERLAY_PARTS)
            .map(|_| (document.clone(), b"<changed/>".to_vec()))
            .collect();
        assert!(matches!(
            package.write_part_overlays_to_stream(&mut output, oversized_set),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(output.is_empty());

        let limits = ReadLimits::builder()
            .max_part_bytes(20)
            .unwrap()
            .max_total_part_bytes(21)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes)),
            limits,
        )
        .unwrap();
        assert!(matches!(
            package.write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document, b"<document/>".to_vec()),
                    (orphan, b"<orphan-2/>".to_vec()),
                ],
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                ..
            })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn multi_part_overlay_all_noop_preserves_signed_source_identity() {
        let source_bytes = signed_archive(b"<signed/>");
        let document = PackURI::new("/word/document.xml").unwrap();
        let signature = PackURI::new("/signature/origin.xml").unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes.clone())))
                .unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document.clone(), b"<signed/>".to_vec()),
                    (signature.clone(), b"<origin/>".to_vec()),
                ],
            )
            .unwrap();
        assert_eq!(output, source_bytes);

        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            signed_archive(b"<signed/>"),
        )))
        .unwrap();
        output.clear();
        assert!(matches!(
            package.write_part_overlays_to_stream(
                &mut output,
                vec![
                    (document, b"<changed/>".to_vec()),
                    (signature, b"<origin/>".to_vec()),
                ],
            ),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_exact_noop_preserves_every_source_byte() {
        let source_bytes = archive_bytes(root_relationships(), b"malformed but unchanged", true);
        let source = Arc::new(CountingSource::new(source_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"malformed but unchanged".to_vec())
            .unwrap();
        assert_eq!(output, source_bytes);
    }

    #[test]
    fn one_part_overlay_refuses_unsupported_physical_layout_before_output() {
        let mut source_bytes = b"unsupported ZIP prelude".to_vec();
        source_bytes.extend_from_slice(&archive_bytes(root_relationships(), b"<before/>", true));
        let source = Arc::new(CountingSource::new(source_bytes));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec()),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_rejects_invalid_xml_and_signed_changes_before_output() {
        let target = PackURI::new("/word/document.xml").unwrap();
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<broken".to_vec()),
            Err(OpcError::XmlPublication { .. })
        ));
        assert!(output.is_empty());

        let signed_bytes = signed_archive(b"<signed/>");
        let signed = Arc::new(CountingSource::new(signed_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(signed).unwrap();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<changed/>".to_vec()),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(output.is_empty());

        let signed = Arc::new(CountingSource::new(signed_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(signed).unwrap();
        package
            .write_part_overlay_to_stream(&mut output, &target, b"<signed/>".to_vec())
            .unwrap();
        assert_eq!(output, signed_bytes);
    }

    #[test]
    fn one_part_overlay_enforces_replacement_limits_without_output() {
        let limits = ReadLimits::builder()
            .max_part_bytes(10)
            .unwrap()
            .max_total_part_bytes(19)
            .unwrap()
            .build()
            .unwrap();
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let package = SourceBackedPackage::from_read_at_with_limits(source, limits).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, vec![b'x'; 11]),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 11,
                maximum: 10
            })
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn one_part_overlay_reports_source_changes_before_and_during_output() {
        let target = PackURI::new("/word/document.xml").unwrap();
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        source.changed();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec()),
            Err(OpcError::SourceChanged { .. })
        ));
        assert!(output.is_empty());

        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
        let mut sink = MutatingSink {
            source,
            bytes: Vec::new(),
            change_after: 1,
            changed: false,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert!(written > 0);
                assert!(matches!(*source, OpcError::SourceChanged { .. }));
            },
            other => panic!("unexpected source-change error: {other:?}"),
        }
    }

    #[test]
    fn one_part_overlay_bounds_writes_and_reports_partial_sink_failure() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            true,
        )));
        let package = SourceBackedPackage::from_read_at(source).unwrap();
        let target = PackURI::new("/word/document.xml").unwrap();
        let mut sink = BoundedFailingSink {
            accepted: 0,
            limit: 100,
            largest_write: 0,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
            other => panic!("unexpected sink error: {other:?}"),
        }
        assert!(sink.largest_write <= SOURCE_PUBLICATION_CHUNK_BYTES);
    }
}
