//! Immutable, source-backed access to an OPC package.
//!
//! This module intentionally exposes a smaller surface than [`OpcPackage`].
//! The latter owns mutable parts, while this type keeps ordinary payloads in a
//! positional source until a caller explicitly asks for one.

use crate::constants::relationship_type;
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
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex};

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
    source: Arc<dyn ReadAt>,
}

impl ZipReaderAt for SourceReader {
    fn read_at(&self, output: &mut [u8], offset: u64) -> std::io::Result<usize> {
        self.source.read_at(offset, output)
    }
}

struct SourceSnapshot {
    source: Arc<dyn ReadAt>,
    version: SourceVersion,
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
/// XML, but never reads ordinary part payloads.  The type has no mutation,
/// writer, raw-copy, or format-migration APIs; call [`Self::into_opc_package`]
/// when an owning mutable package is needed.
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
        };
        snapshot.ensure_current()?;
        let archive = IndexedArchive::from_reader_with_limits(
            SourceReader { source },
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

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic on failure by design"
    )]

    use super::*;
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
        writer.write_stored("custom/orphan.xml", b"orphan").unwrap();
        if include_junk {
            writer.write_stored("scratch.bin", b"not a part").unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn root_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
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
}
