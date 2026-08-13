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
use crate::pkgreader::{
    PackageReader, SerializedRelationship, SourceCatalog, ValidationCatalogError,
    ValidationCatalogPhase,
};
use crate::rel::{Relationships, TargetMode};
use litchi_core::{ExecutionContext, ExecutionError, ReadAt, Reservation, Resource, SourceVersion};
use sha2::{Digest as _, Sha256};
use soapberry_zip::ReaderAt as ZipReaderAt;
use soapberry_zip::office::{EntryId, IndexedArchive};
use std::collections::HashMap;
use std::io::Write;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::time::Duration;

const SOURCE_PUBLICATION_CHUNK_BYTES: usize = 64 * 1024;
const MAX_SOURCE_OVERLAY_PARTS: usize = 64;

struct PendingOverlay {
    target: usize,
    replacement: Vec<u8>,
}

struct ChangedOverlay {
    target: ChangedOverlayTarget,
    replacement: Arc<Vec<u8>>,
}

enum ChangedOverlayTarget {
    Part(usize),
    Member(String),
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
    /// Whether this cache charges retained and in-flight payloads to a caller
    /// supplied hierarchical memory budget.
    pub budget_managed: bool,
    /// Managed budget reservations rejected by the hierarchical budget. A
    /// final `InputBytes` refusal is counted; a temporary oversized read
    /// window that is shrunk to the remaining capacity is not.
    pub budget_reservation_failures: u64,
    /// Current memory usage observed on the managed context's local budget.
    /// This is content-free and may include sibling operations sharing the
    /// same budget.
    pub budget_memory_used: u64,
    /// Bytes reserved by retained cache entries and active cold-load flights.
    /// This deliberately excludes ordinary caller-owned [`PartData`] handles
    /// that were returned after a cache entry was evicted or bypassed.
    pub budget_cache_reserved_bytes: u64,
    /// Local memory limit observed on the managed context's budget. `None`
    /// means that the compatibility, unmanaged cache path is active.
    pub budget_memory_limit: Option<u64>,
    /// Cumulative physical bytes accepted from positional reads charged to
    /// this context. Shared contexts may include sibling operations. This is
    /// never released when a package, cache entry, or payload handle is
    /// dropped.
    pub budget_input_bytes_used: u64,
    /// Local cumulative input-byte limit, or `None` for an unmanaged cache.
    pub budget_input_bytes_limit: Option<u64>,
    /// Cumulative declared cold-load work charged before payload I/O. A
    /// successful ZIP read proves that the declared uncompressed size is also
    /// the actual materialized size; hits and waiters add no work. Shared
    /// contexts may include sibling operations.
    pub budget_work_used: u64,
    /// Local cumulative work limit, or `None` for an unmanaged cache.
    pub budget_work_limit: Option<u64>,
    /// Current retained object usage observed on the managed context's local
    /// budget. Shared contexts may include sibling operations. Unlike input
    /// bytes and work, this usage is released when the package catalog or a
    /// payload object is dropped.
    pub budget_objects_used: u64,
    /// Local retained-object limit, or `None` for an unmanaged cache.
    pub budget_objects_limit: Option<u64>,
    /// Object units retained by the package catalog itself.
    pub budget_catalog_reserved_objects: u64,
    /// Object units retained by cache entries and active flights. Returned
    /// handles that outlive eviction are reflected in `budget_objects_used`
    /// but deliberately excluded here, matching the retained-byte diagnostic.
    /// Shared reservations are counted once.
    pub budget_cache_reserved_objects: u64,
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
    budget_reservation_failures: AtomicU64,
}

#[derive(Clone)]
struct SourceReader {
    snapshot: SourceSnapshot,
}

impl ZipReaderAt for SourceReader {
    fn read_at(&self, output: &mut [u8], offset: u64) -> std::io::Result<usize> {
        let result = read_source_at_with_context(
            &self.snapshot,
            self.snapshot.context.as_ref(),
            offset,
            output,
            "archive",
        );
        result.map_err(|error| {
            let execution = match &error {
                OpcError::Cancelled => Some(ExecutionError::Cancelled),
                OpcError::Execution(execution) => Some(execution.clone()),
                _ => None,
            };
            if let Some(execution) = execution {
                record_source_execution_failure(&self.snapshot, execution);
            }
            match error {
                OpcError::IoError(error) => error,
                OpcError::Execution(error) => execution_io_error(error),
                error => std::io::Error::other(error.to_string()),
            }
        })
    }
}

#[derive(Clone)]
struct SourceSnapshot {
    source: Arc<dyn ReadAt>,
    version: SourceVersion,
    length: u64,
    monitor_reads: Arc<std::sync::atomic::AtomicBool>,
    lineage: SourceLineage,
    context: Option<ExecutionContext>,
    execution_failure: Option<Arc<Mutex<Option<ExecutionError>>>>,
    input_reservation_failures: Option<Arc<AtomicU64>>,
}

/// Process-local identity for one opened source-backed package lineage.
///
/// A lineage is intentionally distinct from [`SourceVersion`]. Two source
/// adapters may report the same caller-chosen version token while still being
/// different package instances; patches must never cross that boundary.
#[derive(Clone, Debug)]
pub struct SourceLineage(Arc<()>);

impl PartialEq for SourceLineage {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl Eq for SourceLineage {}

/// Exact immutable source artifact retained for a later reversible restore.
#[derive(Clone)]
pub struct SourceArtifact {
    snapshot: SourceSnapshot,
}

/// SHA-256 identity of an exact source artifact.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceArtifactFingerprint([u8; 32]);

impl SourceArtifactFingerprint {
    /// Construct from a completed SHA-256 digest.
    #[must_use]
    pub const fn from_sha256(digest: [u8; 32]) -> Self {
        Self(digest)
    }
}

impl SourceArtifact {
    /// Hash the exact current artifact without materializing it.
    pub fn fingerprint(&self) -> Result<SourceArtifactFingerprint> {
        self.snapshot.ensure_current()?;
        let mut hasher = Sha256::new();
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(SOURCE_PUBLICATION_CHUNK_BYTES)
            .map_err(|source| OpcError::Allocation {
                resource: "source artifact fingerprint buffer",
                source,
            })?;
        buffer.resize(SOURCE_PUBLICATION_CHUNK_BYTES, 0);
        let mut offset = 0_u64;
        while offset < self.snapshot.length {
            let remaining =
                usize::try_from((self.snapshot.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
            let read = read_source_at_with_context(
                &self.snapshot,
                self.snapshot.context.as_ref(),
                offset,
                &mut buffer[..remaining],
                "fingerprinting",
            )?;
            if read == 0 {
                return Err(OpcError::IoError(std::io::Error::new(
                    std::io::ErrorKind::UnexpectedEof,
                    "source-backed OPC source ended during fingerprinting",
                )));
            }
            self.snapshot.ensure_current()?;
            hasher.update(&buffer[..read]);
            offset = offset
                .checked_add(read as u64)
                .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
        }
        self.snapshot.ensure_current()?;
        Ok(SourceArtifactFingerprint(hasher.finalize().into()))
    }

    /// Copy the retained source artifact exactly to a sequential sink.
    pub fn write_to_stream<W: Write>(&self, writer: W) -> Result<()> {
        write_exact_snapshot(&self.snapshot, writer, self.snapshot.context.as_ref())
    }
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
        if written > bytes.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "source-backed OPC sink reported {written} bytes for a {}-byte write",
                    bytes.len()
                ),
            ));
        }
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

struct ContextCheckedSink<W> {
    inner: W,
    context: Option<ExecutionContext>,
    failure: Arc<Mutex<Option<ExecutionError>>>,
}

impl<W: Write> ContextCheckedSink<W> {
    fn record_failure(&self, error: ExecutionError) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check_before_write(&self) -> std::io::Result<()> {
        let Some(context) = self.context.as_ref() else {
            return Ok(());
        };
        context.check().map_err(|error| {
            let message = error.to_string();
            self.record_failure(error);
            // `Write::write_all` retries `Interrupted` indefinitely. Use a
            // terminal I/O classification; the shared failure slot below
            // restores the typed execution error after ZIP writing returns.
            std::io::Error::other(message)
        })
    }
}

impl<W: Write> Write for ContextCheckedSink<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.check_before_write()?;
        let result = self.inner.write(bytes);
        if result.is_ok() {
            // Check after each bounded Chunked write as well. If a sink
            // cancels from inside its first write, the accepted byte count is
            // preserved and the next write/flush returns the typed failure.
            if let Some(context) = self.context.as_ref() {
                if let Err(error) = context.check() {
                    self.record_failure(error);
                }
            }
        }
        result
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.check_before_write()?;
        let result = self.inner.flush();
        if result.is_ok() {
            if let Some(context) = self.context.as_ref() {
                if let Err(error) = context.check() {
                    self.record_failure(error);
                }
            }
        }
        result
    }
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

/// Shared immutable bytes returned by [`PartView::data`].
#[derive(Clone, Debug)]
pub struct PartData {
    payload: CachedPayload,
}

impl PartData {
    /// Borrow the part payload.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.payload.bytes.as_slice()
    }

    /// Share an unmanaged payload allocation with another owner.
    ///
    /// Managed payloads retain a hierarchical memory reservation for the
    /// lifetime of this handle and cannot be detached as a bare `Arc`. Use
    /// [`Self::as_bytes`] or clone the [`PartData`] handle instead.
    pub fn into_arc(&self) -> Result<Arc<Vec<u8>>> {
        if self.payload.reservation.is_some() {
            return Err(OpcError::ManagedPartDataArcEscape);
        }
        Ok(Arc::clone(&self.payload.bytes))
    }

    /// Return whether both values pin the same payload allocation.
    ///
    /// This compares allocation identity only; equal bytes loaded separately
    /// return `false`.
    #[must_use]
    pub fn shares_allocation_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.payload.bytes, &other.payload.bytes)
    }
}

/// One payload allocation and, for managed packages, the reservation retained
/// by a cache entry or active same-Part flight.
#[derive(Clone, Debug)]
struct CachedPayload {
    bytes: Arc<Vec<u8>>,
    reservation: Option<Arc<Reservation>>,
    object_reservation: Option<Arc<Reservation>>,
}

impl CachedPayload {
    fn reserved_bytes(&self) -> u64 {
        self.reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount())
    }

    fn reserved_objects(&self) -> u64 {
        self.object_reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount())
    }
}

#[derive(Debug)]
struct CacheEntry {
    payload: CachedPayload,
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
    payload: Option<CachedPayload>,
}

#[derive(Debug)]
struct LoadFlight {
    state: Mutex<FlightState>,
    completed: Condvar,
    reservation: Option<Arc<Reservation>>,
    flight_object_reservation: Option<Arc<Reservation>>,
    payload_object_reservation: Option<Arc<Reservation>>,
}

impl LoadFlight {
    fn new(
        reservation: Option<Arc<Reservation>>,
        flight_object_reservation: Option<Arc<Reservation>>,
        payload_object_reservation: Option<Arc<Reservation>>,
    ) -> Self {
        Self {
            state: Mutex::new(FlightState::default()),
            completed: Condvar::new(),
            reservation,
            flight_object_reservation,
            payload_object_reservation,
        }
    }

    fn reservation(&self) -> Option<Arc<Reservation>> {
        self.reservation.as_ref().map(Arc::clone)
    }

    fn payload_object_reservation(&self) -> Option<Arc<Reservation>> {
        self.payload_object_reservation.as_ref().map(Arc::clone)
    }

    fn wait(
        &self,
        context: Option<&ExecutionContext>,
    ) -> std::result::Result<Option<CachedPayload>, ExecutionError> {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        while !state.complete {
            if let Some(context) = context {
                context.check()?;
            }
            state = match self
                .completed
                .wait_timeout(state, Duration::from_millis(10))
            {
                Ok((state, _)) => state,
                Err(poisoned) => poisoned.into_inner().0,
            };
        }
        if let Some(context) = context {
            context.check()?;
        }
        Ok(state.payload.clone())
    }

    fn finish_success(&self, payload: CachedPayload) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.payload = Some(payload);
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
    Hit(CachedPayload),
    Loader(Arc<LoadFlight>),
    Waiter(Arc<LoadFlight>),
    Bypass(LoadResources),
}

struct LoadResources {
    reservation: Option<Arc<Reservation>>,
    payload_object_reservation: Option<Arc<Reservation>>,
}

#[derive(Debug)]
struct PartCache {
    limits: SourceCacheLimits,
    state: Mutex<CacheState>,
    counters: CacheCounters,
    budget: Option<ExecutionContext>,
    input_reservation_failures: Option<Arc<AtomicU64>>,
}

impl PartCache {
    fn new(limits: SourceCacheLimits) -> Self {
        Self {
            limits,
            state: Mutex::new(CacheState::default()),
            counters: CacheCounters::default(),
            budget: None,
            input_reservation_failures: None,
        }
    }

    fn new_managed(
        limits: SourceCacheLimits,
        context: ExecutionContext,
        input_reservation_failures: Arc<AtomicU64>,
    ) -> Self {
        Self {
            limits,
            state: Mutex::new(CacheState::default()),
            counters: CacheCounters::default(),
            budget: Some(context),
            input_reservation_failures: Some(input_reservation_failures),
        }
    }

    fn is_managed(&self) -> bool {
        self.budget.is_some()
    }

    fn check_context(&self) -> std::result::Result<(), ExecutionError> {
        if let Some(context) = self.budget.as_ref() {
            context.check()?;
        }
        Ok(())
    }

    fn context(&self) -> Option<&ExecutionContext> {
        self.budget.as_ref()
    }

    fn enter(
        &self,
        entry_id: EntryId,
        declared_bytes: u64,
    ) -> std::result::Result<CacheAccess, ExecutionError> {
        self.check_context()?;
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.clock = state.clock.wrapping_add(1);
        let clock = state.clock;
        if let Some(entry) = state.entries.get_mut(&entry_id) {
            entry.last_used = clock;
            self.counters.hits.fetch_add(1, Ordering::Relaxed);
            return Ok(CacheAccess::Hit(entry.payload.clone()));
        }
        if let Some(flight) = state.flights.get(&entry_id) {
            self.counters.waiter_joins.fetch_add(1, Ordering::Relaxed);
            return Ok(CacheAccess::Waiter(Arc::clone(flight)));
        }

        let reservation = self.reserve_for_load(&mut state, declared_bytes)?;
        let payload_object_reservation = self.reserve_object_for_load()?;
        if state.flights.try_reserve(1).is_err() {
            self.charge_cold_work(declared_bytes)?;
            self.counters.cold_loads.fetch_add(1, Ordering::Relaxed);
            self.counters
                .allocation_bypasses
                .fetch_add(1, Ordering::Relaxed);
            return Ok(CacheAccess::Bypass(LoadResources {
                reservation,
                payload_object_reservation,
            }));
        }
        let flight_object_reservation = self.reserve_object_for_load()?;
        self.charge_cold_work(declared_bytes)?;
        let flight = Arc::new(LoadFlight::new(
            reservation,
            flight_object_reservation,
            payload_object_reservation,
        ));
        state.flights.insert(entry_id, Arc::clone(&flight));
        self.counters.cold_loads.fetch_add(1, Ordering::Relaxed);
        Ok(CacheAccess::Loader(flight))
    }

    fn reserve_object_for_load(
        &self,
    ) -> std::result::Result<Option<Arc<Reservation>>, ExecutionError> {
        let Some(context) = self.budget.as_ref() else {
            return Ok(None);
        };
        match context.reserve(Resource::Objects, 1) {
            Ok(reservation) => Ok(Some(Arc::new(reservation))),
            Err(error) => {
                if matches!(error, ExecutionError::ResourceLimit(_)) {
                    self.counters
                        .budget_reservation_failures
                        .fetch_add(1, Ordering::Relaxed);
                }
                Err(error)
            },
        }
    }

    fn charge_cold_work(&self, declared_bytes: u64) -> std::result::Result<(), ExecutionError> {
        let Some(context) = self.budget.as_ref() else {
            return Ok(());
        };
        // Work is cumulative. Charge the declared decompression output before
        // the archive reader starts; `load_part` later requires the actual
        // verified output length to equal this declaration, so no guessed or
        // second work charge is needed.
        context.consume(Resource::Work, declared_bytes)
    }

    fn reserve_for_load(
        &self,
        state: &mut CacheState,
        declared_bytes: u64,
    ) -> std::result::Result<Option<Arc<Reservation>>, ExecutionError> {
        let Some(context) = self.budget.as_ref() else {
            return Ok(None);
        };

        context.check()?;
        self.make_room_for_load(state, declared_bytes);
        match context.reserve(Resource::Memory, declared_bytes) {
            Ok(reservation) => Ok(Some(Arc::new(reservation))),
            Err(first_error) => {
                if matches!(first_error, ExecutionError::ResourceLimit(_)) {
                    self.counters
                        .budget_reservation_failures
                        .fetch_add(1, Ordering::Relaxed);
                }
                // Cache retention is best effort. If a shared ancestor is
                // currently full, dropping all clean entries can make room
                // without ever exceeding that ancestor's limit.
                self.evict_all(state);
                match context.reserve(Resource::Memory, declared_bytes) {
                    Ok(reservation) => Ok(Some(Arc::new(reservation))),
                    Err(error) => {
                        if matches!(error, ExecutionError::ResourceLimit(_)) {
                            self.counters
                                .budget_reservation_failures
                                .fetch_add(1, Ordering::Relaxed);
                        }
                        Err(error)
                    },
                }
            },
        }
    }

    fn make_room_for_load(&self, state: &mut CacheState, declared_bytes: u64) {
        let weight = usize::try_from(declared_bytes).unwrap_or(usize::MAX);
        if weight > self.limits.max_bytes {
            self.evict_all(state);
            return;
        }
        while state.entries.len() >= self.limits.max_entries
            || state.total_bytes.saturating_add(weight) > self.limits.max_bytes
        {
            if !self.evict_oldest(state) {
                break;
            }
        }
    }

    fn evict_all(&self, state: &mut CacheState) {
        while self.evict_oldest(state) {}
    }

    fn evict_oldest(&self, state: &mut CacheState) -> bool {
        let Some((&oldest, _)) = state
            .entries
            .iter()
            .filter(|(_, entry)| !payload_is_externally_pinned(&entry.payload))
            .min_by_key(|(_, entry)| entry.last_used)
        else {
            return false;
        };
        if let Some(removed) = state.entries.remove(&oldest) {
            state.total_bytes = state
                .total_bytes
                .saturating_sub(removed.payload.bytes.len());
            self.counters.evictions.fetch_add(1, Ordering::Relaxed);
            true
        } else {
            false
        }
    }

    fn complete_success(
        &self,
        entry_id: EntryId,
        flight: &Arc<LoadFlight>,
        payload: CachedPayload,
    ) -> std::result::Result<(), ExecutionError> {
        self.check_context()?;
        let delivered = payload.clone();
        {
            let mut state = self
                .state
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            // This is the final cooperative cancellation point immediately
            // before publishing a clean value into the shared cache.
            self.check_context()?;
            self.counters
                .successful_loads
                .fetch_add(1, Ordering::Relaxed);
            self.record_retention(self.insert_locked(&mut state, entry_id, payload));
        }
        // Complete before removing the flight so an oversized, deliberately
        // uncached value still has no gap in which a late peer can start a
        // duplicate load instead of joining this successful delivery.
        flight.finish_success(delivered);
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        remove_flight(&mut state, entry_id, flight);
        Ok(())
    }

    fn complete_failure(&self, entry_id: EntryId, flight: &Arc<LoadFlight>) {
        self.counters.failed_loads.fetch_add(1, Ordering::Relaxed);
        // Publish failure to current waiters before allowing a new retrying
        // loader to install a replacement flight.
        flight.finish_failure();
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        remove_flight(&mut state, entry_id, flight);
    }

    fn complete_bypass_success(
        &self,
        entry_id: EntryId,
        payload: CachedPayload,
    ) -> std::result::Result<(), ExecutionError> {
        self.check_context()?;
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        // Keep the no-flight allocation-fallback path under the same
        // pre-publication cancellation contract as the normal flight path.
        self.check_context()?;
        self.counters
            .successful_loads
            .fetch_add(1, Ordering::Relaxed);
        self.record_retention(self.insert_locked(&mut state, entry_id, payload));
        Ok(())
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
            CacheRetention::Pinned => {
                self.counters.bypasses.fetch_add(1, Ordering::Relaxed);
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
        payload: CachedPayload,
    ) -> CacheRetention {
        let weight = payload.bytes.len();
        if weight > self.limits.max_bytes {
            return CacheRetention::Oversized;
        }
        state.clock = state.clock.wrapping_add(1);
        let clock = state.clock;
        while state.entries.len() >= self.limits.max_entries
            || state.total_bytes.saturating_add(weight) > self.limits.max_bytes
        {
            if !self.evict_oldest(state) {
                return CacheRetention::Pinned;
            }
        }
        if state.entries.try_reserve(1).is_err() {
            return CacheRetention::AllocationFailure;
        }
        if let Some(previous) = state.entries.insert(
            entry_id,
            CacheEntry {
                payload,
                last_used: clock,
            },
        ) {
            state.total_bytes = state
                .total_bytes
                .saturating_sub(previous.payload.bytes.len());
        }
        state.total_bytes = state.total_bytes.saturating_add(weight);
        CacheRetention::Retained
    }

    fn diagnostics(&self) -> SourceCacheDiagnostics {
        let state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        // A successful loader briefly owns the same reservation through its
        // cache entry, completion payload, and returned handles. Count only
        // unique reservation identities so a diagnostic snapshot cannot
        // report more retained cache bytes than the hierarchical budget.
        let mut budget_cache_reserved_bytes = state
            .entries
            .values()
            .map(|entry| entry.payload.reserved_bytes())
            .sum::<u64>();
        let mut budget_cache_reserved_objects = state
            .entries
            .values()
            .map(|entry| entry.payload.reserved_objects())
            .sum::<u64>();
        for flight in state.flights.values() {
            if let Some(reservation) = flight.reservation.as_ref() {
                let already_counted = state.entries.values().any(|entry| {
                    entry
                        .payload
                        .reservation
                        .as_ref()
                        .is_some_and(|existing| Arc::ptr_eq(existing, reservation))
                });
                if !already_counted {
                    budget_cache_reserved_bytes =
                        budget_cache_reserved_bytes.saturating_add(reservation.amount());
                }
            }
            if let Some(object_reservation) = flight.payload_object_reservation.as_ref() {
                let already_counted = state.entries.values().any(|entry| {
                    entry
                        .payload
                        .object_reservation
                        .as_ref()
                        .is_some_and(|existing| Arc::ptr_eq(existing, object_reservation))
                });
                if !already_counted {
                    budget_cache_reserved_objects =
                        budget_cache_reserved_objects.saturating_add(object_reservation.amount());
                }
            }
            if let Some(object_reservation) = flight.flight_object_reservation.as_ref() {
                budget_cache_reserved_objects =
                    budget_cache_reserved_objects.saturating_add(object_reservation.amount());
            }
        }
        let (budget_input_bytes_used, budget_input_bytes_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::InputBytes),
                    Some(context.budget().limit(Resource::InputBytes)),
                )
            });
        let (budget_work_used, budget_work_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::Work),
                    Some(context.budget().limit(Resource::Work)),
                )
            });
        let (budget_objects_used, budget_objects_limit) =
            self.budget.as_ref().map_or((0, None), |context| {
                (
                    context.budget().used(Resource::Objects),
                    Some(context.budget().limit(Resource::Objects)),
                )
            });
        let budget_reservation_failures = self
            .counters
            .budget_reservation_failures
            .load(Ordering::Relaxed)
            .saturating_add(
                self.input_reservation_failures
                    .as_ref()
                    .map_or(0, |counter| counter.load(Ordering::Relaxed)),
            );
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
            budget_managed: self.budget.is_some(),
            budget_reservation_failures,
            budget_memory_used: self
                .budget
                .as_ref()
                .map_or(0, |context| context.budget().used(Resource::Memory)),
            budget_cache_reserved_bytes,
            budget_memory_limit: self
                .budget
                .as_ref()
                .map(|context| context.budget().limit(Resource::Memory)),
            budget_input_bytes_used,
            budget_input_bytes_limit,
            budget_work_used,
            budget_work_limit,
            budget_objects_used,
            budget_objects_limit,
            budget_catalog_reserved_objects: 0,
            budget_cache_reserved_objects,
        }
    }
}

#[derive(Clone, Copy)]
enum CacheRetention {
    Retained,
    Oversized,
    Pinned,
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

fn payload_is_externally_pinned(payload: &CachedPayload) -> bool {
    // Unmanaged `PartData::into_arc` can outlive its handle, while managed
    // handles retain a reservation. Check both identities before evicting an
    // entry so either form of caller ownership keeps the bytes pinned.
    Arc::strong_count(&payload.bytes) > 1
        || payload
            .reservation
            .as_ref()
            .is_some_and(|reservation| Arc::strong_count(reservation) > 1)
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
    catalog_object_reservation: Option<Arc<Reservation>>,
}

/// Validation-only open failure with exact ingress phase provenance.
pub(crate) struct ValidationOpenError {
    pub(crate) phase: ValidationCatalogPhase,
    pub(crate) error: OpcError,
}

impl SourceBackedPackage {
    /// Validation-only source open. Ordinary callers retain the existing open
    /// path; this variant adds phase provenance without changing its hot path.
    pub(crate) fn from_read_at_for_validation(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
    ) -> std::result::Result<Self, ValidationOpenError> {
        let phase = |phase, error| ValidationOpenError { phase, error };
        let version = source
            .version()
            .map_err(OpcError::from)
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let length = source
            .len()
            .map_err(OpcError::from)
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        limits
            .check(ReadResource::InputBytes, length, limits.max_input_bytes())
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let snapshot = SourceSnapshot {
            source: Arc::clone(&source),
            version,
            length,
            monitor_reads: Arc::new(std::sync::atomic::AtomicBool::new(false)),
            lineage: SourceLineage(Arc::new(())),
            context: None,
            execution_failure: None,
            input_reservation_failures: None,
        };
        snapshot
            .ensure_current()
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let archive = IndexedArchive::from_reader_with_limits(
            SourceReader {
                snapshot: snapshot.clone(),
            },
            length,
            limits.zip_limits(),
        )
        .map_err(OpcError::from)
        .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        snapshot
            .ensure_current()
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
        } = PackageReader::source_catalog_for_validation(&archive, limits).map_err(
            |ValidationCatalogError {
                 phase: stage,
                 error,
             }| phase(stage, error),
        )?;
        snapshot
            .ensure_current()
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;

        let package_relationships = relationships_for_package(pkg_srels)
            .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?;
        let mut catalog_parts = Vec::new();
        catalog_parts
            .try_reserve_exact(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC catalog parts",
                source,
            })
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let mut parts_by_name = HashMap::new();
        parts_by_name
            .try_reserve(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC part lookup",
                source,
            })
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        for (index, part) in parts.into_iter().enumerate() {
            let relationships = relationships_for_part(&part.partname, part.srels)
                .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?;
            let entry_id = archive
                .entry_id(part.partname.membername())
                .ok_or_else(|| OpcError::PartNotFound(part.partname.to_string()))
                .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
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
            cache: PartCache::new(SourceCacheLimits::default()),
            catalog_object_reservation: None,
        })
    }

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

    /// Open a source-backed package whose lazy payload cache is charged to an
    /// explicit hierarchical execution budget.
    ///
    /// Compatibility constructors remain unmanaged and keep their existing
    /// behavior. This opt-in path checks cancellation before opening and
    /// reserves each Part's declared uncompressed size before reading its
    /// payload. The reservation is retained with the clean cache entry and
    /// active same-Part flight. A returned [`PartData`] is a budgeted handle;
    /// use [`PartData::as_bytes`] or clone that handle while consuming the
    /// payload. Its [`PartData::into_arc`] escape is rejected on this managed
    /// path so the reservation cannot be silently detached.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            limits,
            SourceCacheLimits::default(),
            context,
        )
    }

    /// Open a managed source-backed package with explicit read, cache, and
    /// hierarchical execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_inner(source, limits, cache_limits, Some(context))
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
        Self::from_read_at_inner(source, limits, cache_limits, None)
    }

    fn from_read_at_inner(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: Option<ExecutionContext>,
    ) -> Result<Self> {
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let execution_failure = context
            .as_ref()
            .map(|_| Arc::new(Mutex::new(None::<ExecutionError>)));
        let input_reservation_failures = context.as_ref().map(|_| Arc::new(AtomicU64::new(0)));
        let version = source.version()?;
        let length = source.len()?;
        limits.check(ReadResource::InputBytes, length, limits.max_input_bytes())?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let snapshot = SourceSnapshot {
            source: Arc::clone(&source),
            version,
            length,
            monitor_reads: Arc::new(std::sync::atomic::AtomicBool::new(false)),
            lineage: SourceLineage(Arc::new(())),
            context: context.clone(),
            execution_failure: execution_failure.clone(),
            input_reservation_failures: input_reservation_failures.clone(),
        };
        snapshot.ensure_current()?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        let archive = match IndexedArchive::from_reader_with_limits(
            SourceReader {
                snapshot: snapshot.clone(),
            },
            length,
            limits.zip_limits(),
        ) {
            Ok(archive) => archive,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&snapshot) {
                    return Err(map_execution_error(execution));
                }
                return Err(error.into());
            },
        };
        snapshot.ensure_current()?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
        // The indexed archive and its source catalog remain owned by the
        // package for its whole lifetime. Reserve one object for the package
        // catalog owner and one for every physical member before parsing the
        // source catalog and projecting deferred part vectors; this is a
        // bounded retained charge, not a cumulative event. Payloads and
        // load flights use separate units.
        let catalog_object_reservation = if let Some(context) = context.as_ref() {
            let member_objects = u64::try_from(archive.len()).map_err(|_| {
                overlay_unavailable("source-backed OPC catalog member count overflows u64")
            })?;
            let object_count = member_objects.checked_add(1).ok_or_else(|| {
                overlay_unavailable("source-backed OPC catalog object count overflows u64")
            })?;
            Some(Arc::new(
                context
                    .reserve(Resource::Objects, object_count)
                    .map_err(map_execution_error)?,
            ))
        } else {
            None
        };
        let SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
        } = match PackageReader::source_catalog(&archive, limits) {
            Ok(catalog) => catalog,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&snapshot) {
                    return Err(map_execution_error(execution));
                }
                return Err(error);
            },
        };
        snapshot.ensure_current()?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }

        let package_relationships =
            relationships_for_package_with_context(pkg_srels, context.as_ref())?;
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }
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
            if let Some(context) = context.as_ref() {
                context.check().map_err(map_execution_error)?;
            }
            let relationships =
                relationships_for_part_with_context(&part.partname, part.srels, context.as_ref())?;
            if let Some(context) = context.as_ref() {
                context.check().map_err(map_execution_error)?;
            }
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
        if let Some(context) = context.as_ref() {
            context.check().map_err(map_execution_error)?;
        }

        let cache = if let Some(context) = context {
            let input_reservation_failures = input_reservation_failures.ok_or_else(|| {
                overlay_unavailable("managed source input reservation counter is unavailable")
            })?;
            PartCache::new_managed(cache_limits, context, input_reservation_failures)
        } else {
            PartCache::new(cache_limits)
        };
        Ok(Self {
            source: snapshot,
            archive,
            limits,
            package_relationships,
            parts: catalog_parts,
            parts_by_name,
            non_part_members,
            cache,
            catalog_object_reservation,
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
        let mut diagnostics = self.cache.diagnostics();
        diagnostics.budget_catalog_reserved_objects = self
            .catalog_object_reservation
            .as_ref()
            .map_or(0, |reservation| reservation.amount());
        diagnostics
    }

    /// Check the caller-supplied execution policy for a source-backed
    /// operation.
    ///
    /// Compatibility packages have no execution context and therefore always
    /// pass this check. Managed packages use the check to let a semantic
    /// facade honor cancellation even when its own parsed value is already
    /// retained and no new [`PartData`] read is necessary.
    pub fn check_execution(&self) -> Result<()> {
        self.cache.check_context().map_err(map_execution_error)
    }

    /// Return the exact source lineage captured by this package.
    ///
    /// The returned token is clone-cheap and cannot be constructed by a
    /// caller. It lets a semantic facade bind snapshots and patches to this
    /// opened package rather than to a merely equal [`SourceVersion`].
    #[must_use]
    pub fn source_lineage(&self) -> SourceLineage {
        self.source.lineage.clone()
    }

    /// Clone the caller-supplied execution context, if this package is on the
    /// managed path. Compatibility packages return `None`.
    #[must_use]
    pub fn execution_context(&self) -> Option<ExecutionContext> {
        self.cache.context().cloned()
    }

    /// Return the exact process-local source identity and revision captured at
    /// open after verifying that the positional source is still current.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.source.ensure_current()?;
        Ok(self.source.version)
    }

    /// Retain an O(1) handle to the exact immutable source artifact.
    #[must_use]
    pub fn source_artifact(&self) -> SourceArtifact {
        SourceArtifact {
            snapshot: self.source.clone(),
        }
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

    /// Replace one ordinary Part while removing a bounded set of external
    /// relationships owned by that Part.
    ///
    /// This is an explicit, forward-only topology-changing publisher intended
    /// for sanitizers which have already proved that every removed relationship
    /// reference is inside the replacement payload. Relationship IDs must be
    /// unique, must exist on `partname`, and must identify external targets.
    /// The selected Part and its relationships member are regenerated; every
    /// other physical ZIP member is raw-copied.
    ///
    /// # Errors
    ///
    /// Returns before output for a missing, duplicate, internal, oversized, or
    /// signed selection. Sink failures after output begins are reported as
    /// [`OpcError::IncompleteOutput`].
    pub fn write_part_overlay_with_external_relationship_removals_to_stream<W: Write>(
        self,
        writer: W,
        partname: &PackURI,
        replacement: Vec<u8>,
        mut removed_relationship_ids: Vec<String>,
    ) -> Result<()> {
        if removed_relationship_ids.len() > MAX_SOURCE_OVERLAY_PARTS {
            return Err(overlay_unavailable(format!(
                "relationship removal set exceeds the {MAX_SOURCE_OVERLAY_PARTS}-relationship bound"
            )));
        }
        if removed_relationship_ids.is_empty() {
            return self.write_part_overlay_to_stream(writer, partname, replacement);
        }
        removed_relationship_ids.sort_unstable();
        if let Some(duplicate) = removed_relationship_ids
            .windows(2)
            .find(|pair| pair[0] == pair[1])
        {
            return Err(OpcError::DuplicateRelationshipId(duplicate[0].clone()));
        }

        let target = self
            .parts_by_name
            .get(partname)
            .copied()
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
        let target_part = &self.parts[target];
        for id in &removed_relationship_ids {
            let relationship = target_part.relationships.get(id).ok_or_else(|| {
                OpcError::RelationshipNotFound(format!("relationship '{id}' was not found"))
            })?;
            if !relationship.is_external() {
                return Err(OpcError::InvalidRelationship(format!(
                    "relationship '{id}' is not external"
                )));
            }
        }
        let relationship_xml =
            relationship_xml_without(&target_part.relationships, &removed_relationship_ids)?;
        self.limits.check(
            ReadResource::RelationshipXmlBytes,
            relationship_xml.len() as u64,
            self.limits.max_relationship_xml_bytes() as u64,
        )?;

        let relationship_uri = partname.rels_uri().map_err(OpcError::InvalidPackUri)?;
        let relationship_member = relationship_uri.membername().to_owned();
        let relationship_entry = self.archive.entry_id(&relationship_member).ok_or_else(|| {
            OpcError::RelationshipNotFound(format!(
                "relationships member '{}' was not found",
                relationship_uri.as_str()
            ))
        })?;
        self.validate_overlay_limits(std::iter::once((target, replacement.len())))?;
        self.validate_part_and_relationship_overlay_limits(
            target,
            replacement.len(),
            relationship_entry,
            relationship_xml.len(),
        )?;
        let original_part = self.read_part(target)?;
        let original_relationships = match self.archive.read_entry(relationship_entry) {
            Ok(bytes) => bytes,
            Err(error) => {
                if let Some(execution) = take_source_execution_failure(&self.source) {
                    return Err(map_execution_error(execution));
                }
                return Err(error.into());
            },
        };
        self.source.ensure_current()?;
        if self.has_signature_infrastructure() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        if xml_minifier::audit::package::is_xml_part(
            target_part.partname.as_str(),
            &target_part.content_type,
        ) {
            validate_overlay_xml(target_part.partname.as_str(), original_part.as_bytes())?;
            validate_overlay_xml(target_part.partname.as_str(), &replacement)?;
        }
        validate_overlay_xml(relationship_uri.as_str(), &original_relationships)?;
        validate_overlay_xml(relationship_uri.as_str(), &relationship_xml)?;

        let changed = [
            ChangedOverlay {
                target: ChangedOverlayTarget::Part(target),
                replacement: Arc::new(replacement),
            },
            ChangedOverlay {
                target: ChangedOverlayTarget::Member(relationship_member),
                replacement: Arc::new(relationship_xml),
            },
        ];
        self.write_changed_overlays(writer, &changed)
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
        self.validate_overlay_limits(
            overlays
                .iter()
                .map(|overlay| (overlay.target, overlay.replacement.len())),
        )?;

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
                target: ChangedOverlayTarget::Part(overlay.target),
                replacement: Arc::new(overlay.replacement),
            });
        }
        self.write_changed_overlays(writer, &replacements)
    }

    fn validate_overlay_limits<I>(&self, overlays: I) -> Result<()>
    where
        I: Iterator<Item = (usize, usize)> + Clone,
    {
        for (_, replacement_len) in overlays.clone() {
            let replacement_bytes =
                u64::try_from(replacement_len).map_err(|_| OpcError::ReadLimit {
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
        for (target, replacement_len) in overlays {
            let target_bytes = self
                .archive
                .metadata_for(self.parts[target].entry_id)?
                .uncompressed_size();
            let replacement_bytes = replacement_len as u64;
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

    fn validate_part_and_relationship_overlay_limits(
        &self,
        target: usize,
        replacement_len: usize,
        relationship_entry: EntryId,
        relationship_len: usize,
    ) -> Result<()> {
        let relationship_bytes = relationship_len as u64;
        self.limits.check(
            ReadResource::ArchiveEntryBytes,
            relationship_bytes,
            self.limits.max_archive_entry_bytes(),
        )?;
        self.limits.check(
            ReadResource::RelationshipXmlBytes,
            relationship_bytes,
            self.limits.max_relationship_xml_bytes() as u64,
        )?;

        let mut archive_total = 0_u64;
        let mut relationship_total = 0_u64;
        for name in self.archive.file_names() {
            let bytes = self.archive.metadata(name)?.uncompressed_size();
            archive_total = checked_overlay_total(
                archive_total,
                bytes,
                ReadResource::ArchiveTotalBytes,
                self.limits.max_archive_total_bytes(),
            )?;
            if is_relationship_member_name(name) {
                relationship_total = checked_overlay_total(
                    relationship_total,
                    bytes,
                    ReadResource::TotalRelationshipXmlBytes,
                    self.limits.max_total_relationship_xml_bytes() as u64,
                )?;
            }
        }
        let original_part = self
            .archive
            .metadata_for(self.parts[target].entry_id)?
            .uncompressed_size();
        let original_relationship = self
            .archive
            .metadata_for(relationship_entry)?
            .uncompressed_size();
        let adjusted_archive = adjusted_overlay_total(
            archive_total,
            original_part,
            replacement_len as u64,
            ReadResource::ArchiveTotalBytes,
            self.limits.max_archive_total_bytes(),
        )?;
        let adjusted_archive = adjusted_overlay_total(
            adjusted_archive,
            original_relationship,
            relationship_bytes,
            ReadResource::ArchiveTotalBytes,
            self.limits.max_archive_total_bytes(),
        )?;
        self.limits.check(
            ReadResource::ArchiveTotalBytes,
            adjusted_archive,
            self.limits.max_archive_total_bytes(),
        )?;
        let adjusted_relationships = adjusted_overlay_total(
            relationship_total,
            original_relationship,
            relationship_bytes,
            ReadResource::TotalRelationshipXmlBytes,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )?;
        self.limits.check(
            ReadResource::TotalRelationshipXmlBytes,
            adjusted_relationships,
            self.limits.max_total_relationship_xml_bytes() as u64,
        )
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
        write_exact_snapshot(&self.source, writer, self.cache.context())
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
                if let Some(execution) = take_source_execution_failure(&self.source) {
                    return Err(map_execution_error(execution));
                }
                self.source.ensure_current()?;
                return Err(overlay_unavailable(error.to_string()));
            },
        };
        self.source.ensure_current()?;

        let mut replacement_bytes = 0_u64;
        for replacement in replacements {
            let target_name = replacement_member_name(replacement, &self.parts);
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
                    == replacement_member_name(replacement, &self.parts).as_bytes()
            }) {
                let target_name = replacement_member_name(replacement, &self.parts);
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
        self.cache.check_context().map_err(map_execution_error)?;
        let mut written = 0_u64;
        let execution_failure = Arc::new(Mutex::new(None));
        let result = {
            let counted = Counted {
                inner: writer,
                written: &mut written,
            };
            let checked = SourceCheckedSink {
                inner: counted,
                snapshot: self.source.clone(),
            };
            let cooperative = ContextCheckedSink {
                inner: checked,
                context: self.cache.context().cloned(),
                failure: Arc::clone(&execution_failure),
            };
            let result = index.write_to(&plan, Chunked { inner: cooperative });
            match result {
                Ok(mut sink) => sink.flush().map_err(OpcError::IoError),
                Err(error) => Err(OpcError::ZipError(error.to_string())),
            }
        };
        let result = execution_failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
            .map_or(result, |error| Err(map_execution_error(error)));
        let result = take_source_execution_failure(&self.source)
            .map_or(result, |error| Err(map_execution_error(error)));
        finish_source_publication(result, &self.source, written)
    }

    fn read_part(&self, index: usize) -> Result<PartData> {
        self.cache.check_context().map_err(map_execution_error)?;
        let entry_id = self
            .parts
            .get(index)
            .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?
            .entry_id;
        let declared_bytes = if self.cache.is_managed() {
            let declared = self.archive.metadata_for(entry_id)?.uncompressed_size();
            self.limits.check(
                ReadResource::PartBytes,
                declared,
                self.limits.max_part_bytes(),
            )?;
            Some(declared)
        } else {
            None
        };
        loop {
            self.source.ensure_current()?;
            match self
                .cache
                .enter(entry_id, declared_bytes.unwrap_or_default())
                .map_err(map_execution_error)?
            {
                CacheAccess::Hit(bytes) => {
                    self.source.ensure_current()?;
                    self.cache.check_context().map_err(map_execution_error)?;
                    return Ok(PartData { payload: bytes });
                },
                CacheAccess::Waiter(flight) => {
                    if let Some(payload) = flight
                        .wait(self.cache.context())
                        .map_err(map_execution_error)?
                    {
                        self.source.ensure_current()?;
                        self.cache.check_context().map_err(map_execution_error)?;
                        return Ok(PartData { payload });
                    }
                    // The loader may have failed; in that case the flight is
                    // removed and this caller retries rather than observing a
                    // retained error. This also re-checks source freshness.
                },
                CacheAccess::Loader(flight) => {
                    return self.load_part(index, entry_id, declared_bytes, Some(flight), None);
                },
                CacheAccess::Bypass(reservation) => {
                    return self.load_part(
                        index,
                        entry_id,
                        declared_bytes,
                        None,
                        Some(reservation),
                    );
                },
            }
        }
    }

    fn load_part(
        &self,
        index: usize,
        entry_id: EntryId,
        declared_bytes: Option<u64>,
        flight: Option<Arc<LoadFlight>>,
        bypass_resources: Option<LoadResources>,
    ) -> Result<PartData> {
        let result = (|| {
            let part = self
                .parts
                .get(index)
                .ok_or_else(|| OpcError::PartNotFound(index.to_string()))?;
            let bytes = match self.archive.read_entry(part.entry_id) {
                Ok(bytes) => bytes,
                Err(error) => {
                    if let Some(execution) = take_source_execution_failure(&self.source) {
                        return Err(map_execution_error(execution));
                    }
                    return Err(error.into());
                },
            };
            // The decompressor has finished and no payload has been
            // published yet. Cancellation here discards the cold result.
            self.cache.check_context().map_err(map_execution_error)?;
            if let Some(declared) = declared_bytes {
                if bytes.len() as u64 != declared {
                    return Err(OpcError::ZipError(format!(
                        "source-backed OPC Part declared {declared} uncompressed bytes but read {}",
                        bytes.len()
                    )));
                }
            }
            self.source.ensure_current()?;
            self.limits.check(
                ReadResource::PartBytes,
                bytes.len() as u64,
                self.limits.max_part_bytes(),
            )?;
            // Check immediately before publishing. If the source changed
            // during the cold read, no stale payload enters the cache.
            self.source.ensure_current()?;
            self.cache.check_context().map_err(map_execution_error)?;
            Ok(Arc::new(bytes))
        })();
        let reservation = flight
            .as_ref()
            .and_then(|flight| flight.reservation())
            .or_else(|| {
                bypass_resources
                    .as_ref()
                    .and_then(|resources| resources.reservation.as_ref().map(Arc::clone))
            });
        let object_reservation = flight
            .as_ref()
            .and_then(|flight| flight.payload_object_reservation())
            .or_else(|| {
                bypass_resources.as_ref().and_then(|resources| {
                    resources
                        .payload_object_reservation
                        .as_ref()
                        .map(Arc::clone)
                })
            });
        match (flight, result) {
            (Some(flight), Ok(bytes)) => {
                let payload = CachedPayload {
                    bytes,
                    reservation,
                    object_reservation,
                };
                if let Err(error) = self
                    .cache
                    .complete_success(entry_id, &flight, payload.clone())
                {
                    self.cache.complete_failure(entry_id, &flight);
                    return Err(map_execution_error(error));
                }
                Ok(PartData { payload })
            },
            (Some(flight), Err(error)) => {
                self.cache.complete_failure(entry_id, &flight);
                Err(error)
            },
            (None, Ok(bytes)) => {
                let payload = CachedPayload {
                    bytes,
                    reservation,
                    object_reservation,
                };
                if let Err(error) = self
                    .cache
                    .complete_bypass_success(entry_id, payload.clone())
                {
                    self.cache.complete_bypass_failure();
                    return Err(map_execution_error(error));
                }
                Ok(PartData { payload })
            },
            (None, Err(error)) => {
                self.cache.complete_bypass_failure();
                Err(error)
            },
        }
    }
}

fn replacement_member_name<'a>(
    replacement: &'a ChangedOverlay,
    parts: &'a [CatalogPart],
) -> &'a str {
    match &replacement.target {
        ChangedOverlayTarget::Part(index) => parts[*index].partname.membername(),
        ChangedOverlayTarget::Member(name) => name,
    }
}

fn is_relationship_member_name(member_name: &str) -> bool {
    let Some((directory, filename)) = member_name.rsplit_once('/') else {
        return false;
    };
    let has_rels_extension = filename
        .rsplit_once('.')
        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("rels"));
    has_rels_extension && (directory == "_rels" || directory.ends_with("/_rels"))
}

fn relationship_xml_without(
    relationships: &Relationships,
    removed_ids: &[String],
) -> Result<Vec<u8>> {
    const HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#;
    const FOOTER: &str = "</Relationships>";
    const ELEMENT_OVERHEAD: usize = 80;
    const MAX_ESCAPE_EXPANSION: usize = 6;

    let retained_count = relationships
        .iter()
        .filter(|relationship| {
            removed_ids
                .binary_search_by(|id| id.as_str().cmp(relationship.r_id()))
                .is_err()
        })
        .count();
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(retained_count)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed relationship removal order",
            source,
        })?;
    for relationship in relationships.iter().filter(|relationship| {
        removed_ids
            .binary_search_by(|id| id.as_str().cmp(relationship.r_id()))
            .is_err()
    }) {
        retained.push(relationship);
    }
    retained.sort_unstable_by_key(|relationship| relationship.r_id());

    let capacity = retained
        .iter()
        .try_fold(HEADER.len() + FOOTER.len(), |total, relationship| {
            let value_bytes = relationship
                .r_id()
                .len()
                .checked_add(relationship.reltype().len())?
                .checked_add(relationship.target_ref().len())?;
            total
                .checked_add(ELEMENT_OVERHEAD)?
                .checked_add(value_bytes.checked_mul(MAX_ESCAPE_EXPANSION)?)
        })
        .ok_or_else(|| {
            overlay_unavailable("relationship removal serialization capacity overflows usize")
        })?;
    let mut xml = String::new();
    xml.try_reserve_exact(capacity)
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed relationship removal XML",
            source,
        })?;
    xml.push_str(HEADER);
    for relationship in retained {
        xml.push_str(r#"<Relationship Id=""#);
        push_xml_escaped(&mut xml, relationship.r_id());
        xml.push_str(r#"" Type=""#);
        push_xml_escaped(&mut xml, relationship.reltype());
        xml.push_str(r#"" Target=""#);
        push_xml_escaped(&mut xml, relationship.target_ref());
        if relationship.target_mode() == TargetMode::External {
            xml.push_str(r#"" TargetMode="External"/>"#);
        } else {
            xml.push_str(r#""/>"#);
        }
    }
    xml.push_str(FOOTER);
    Ok(xml.into_bytes())
}

fn push_xml_escaped(output: &mut String, value: &str) {
    for character in value.chars() {
        output.push_str(match character {
            '&' => "&amp;",
            '<' => "&lt;",
            '>' => "&gt;",
            '"' => "&quot;",
            '\'' => "&apos;",
            _ => {
                output.push(character);
                continue;
            },
        });
    }
}

fn relationships_for_package(
    serialized: impl IntoIterator<Item = SerializedRelationship>,
) -> Result<Relationships> {
    relationships_for_package_with_context(serialized, None)
}

fn relationships_for_package_with_context(
    serialized: impl IntoIterator<Item = SerializedRelationship>,
    context: Option<&ExecutionContext>,
) -> Result<Relationships> {
    let mut relationships = Relationships::new(PACKAGE_URI.to_string());
    for relationship in serialized {
        if let Some(context) = context {
            context.check().map_err(map_execution_error)?;
        }
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
    relationships_for_part_with_context(partname, serialized, None)
}

fn relationships_for_part_with_context(
    partname: &PackURI,
    serialized: impl IntoIterator<Item = SerializedRelationship>,
    context: Option<&ExecutionContext>,
) -> Result<Relationships> {
    let mut relationships = Relationships::for_source(partname);
    for relationship in serialized {
        if let Some(context) = context {
            context.check().map_err(map_execution_error)?;
        }
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

fn map_execution_error(error: ExecutionError) -> OpcError {
    match error {
        ExecutionError::Cancelled => OpcError::Cancelled,
        error => OpcError::Execution(error),
    }
}

fn execution_io_error(error: ExecutionError) -> std::io::Error {
    std::io::Error::other(error.to_string())
}

fn record_execution_failure(failure: &Arc<Mutex<Option<ExecutionError>>>, error: ExecutionError) {
    let mut slot = failure
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn record_source_execution_failure(snapshot: &SourceSnapshot, error: ExecutionError) {
    if let Some(failure) = snapshot.execution_failure.as_ref() {
        record_execution_failure(failure, error);
    }
}

fn record_input_reservation_failure(snapshot: &SourceSnapshot) {
    if let Some(counter) = snapshot.input_reservation_failures.as_ref() {
        counter.fetch_add(1, Ordering::Relaxed);
    }
}

fn take_source_execution_failure(snapshot: &SourceSnapshot) -> Option<ExecutionError> {
    snapshot.execution_failure.as_ref().and_then(|failure| {
        failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
    })
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

fn write_exact_snapshot<W: Write>(
    source: &SourceSnapshot,
    writer: W,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    source.monitor_publication();
    source.ensure_current()?;
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
            snapshot: source.clone(),
        };
        let mut offset = 0_u64;
        (|| {
            while offset < source.length {
                if let Some(context) = context {
                    context.check().map_err(map_execution_error)?;
                }
                let remaining = usize::try_from((source.length - offset).min(buffer.len() as u64))
                    .map_err(|_| overlay_unavailable("source range does not fit this platform"))?;
                let read = read_source_at_with_context(
                    source,
                    context,
                    offset,
                    &mut buffer[..remaining],
                    "publication",
                )?;
                if read == 0 {
                    return Err(OpcError::IoError(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "source-backed OPC source ended during publication",
                    )));
                }
                source.ensure_current()?;
                if let Some(context) = context {
                    context.check().map_err(map_execution_error)?;
                }
                sink.write_all(&buffer[..read])?;
                offset = offset
                    .checked_add(read as u64)
                    .ok_or_else(|| overlay_unavailable("source offset overflow"))?;
            }
            if let Some(context) = context {
                context.check().map_err(map_execution_error)?;
            }
            sink.flush()?;
            Ok(())
        })()
    };
    finish_source_publication(result, source, written)
}

/// Reserve a bounded positional-read window, perform exactly one source read,
/// and commit only the bytes actually accepted. The retry loop is important
/// for short-read adapters: a caller with one input byte remaining must still
/// be allowed to read one byte even when the adapter initially receives a
/// larger output buffer. A reservation is held only across the physical read,
/// so cumulative [`Resource::InputBytes`] usage is exact under retries and
/// never leaks on an I/O failure.
fn read_source_at_with_context(
    snapshot: &SourceSnapshot,
    context: Option<&ExecutionContext>,
    offset: u64,
    output: &mut [u8],
    operation: &str,
) -> Result<usize> {
    if let Some(context) = context {
        context.check().map_err(map_execution_error)?;
    }
    snapshot.ensure_current_io_if_monitored()?;
    let requested = output.len();
    let (read_output, reservation) = if let Some(context) = context {
        let mut attempt = u64::try_from(requested)
            .map_err(|_| overlay_unavailable("source read length overflows u64"))?;
        loop {
            match context.reserve(Resource::InputBytes, attempt) {
                Ok(reservation) => {
                    let length = usize::try_from(attempt)
                        .map_err(|_| overlay_unavailable("source read length overflows usize"))?;
                    break (&mut output[..length], Some(reservation));
                },
                Err(error) => {
                    let Some(limit) = (match &error {
                        ExecutionError::ResourceLimit(limit) => Some(limit),
                        ExecutionError::Cancelled
                        | ExecutionError::WorkersExceedInFlightTasks { .. }
                        | ExecutionError::ParallelThresholdExceedsInFlightBytes { .. } => None,
                        _ => None,
                    }) else {
                        return Err(map_execution_error(error));
                    };
                    let previous = limit.observed.saturating_sub(attempt);
                    let available = limit.limit.saturating_sub(previous);
                    let next = available.min(attempt.saturating_sub(1));
                    if next == 0 {
                        record_input_reservation_failure(snapshot);
                        return Err(map_execution_error(error));
                    }
                    attempt = next;
                },
            }
        }
    } else {
        (output, None)
    };
    let read = match snapshot.source.read_at(offset, read_output) {
        Ok(read) => read,
        Err(error) => return Err(OpcError::IoError(error)),
    };
    validate_source_read_count(read, read_output.len(), operation)?;
    if let Some(reservation) = reservation {
        if !reservation.commit(read as u64) {
            return Err(overlay_unavailable(
                "source-backed OPC input reservation underflow",
            ));
        }
    }
    if let Some(context) = context {
        context.check().map_err(|error| {
            record_source_execution_failure(snapshot, error.clone());
            map_execution_error(error)
        })?;
    }
    snapshot
        .ensure_current_io_if_monitored()
        .map_err(OpcError::IoError)?;
    Ok(read)
}

fn validate_source_read_count(read: usize, requested: usize, operation: &str) -> Result<()> {
    if read <= requested {
        return Ok(());
    }
    Err(OpcError::IoError(std::io::Error::new(
        std::io::ErrorKind::InvalidData,
        format!(
            "source-backed OPC source reported {read} bytes for a {requested}-byte {operation} read"
        ),
    )))
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
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Barrier;
    use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
    use std::time::Duration;

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
    };

    struct CountingSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        reads: AtomicUsize,
        read_bytes: AtomicU64,
        max_read: usize,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                reads: AtomicUsize::new(0),
                read_bytes: AtomicU64::new(0),
                max_read: usize::MAX,
            }
        }

        fn chunked(bytes: Vec<u8>, max_read: usize) -> Self {
            let mut source = Self::new(bytes);
            source.max_read = max_read.max(1);
            source
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
            let count = count.min(self.max_read);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            self.read_bytes.fetch_add(count as u64, Ordering::SeqCst);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(42, self.revision.load(Ordering::SeqCst)))
        }
    }

    struct CancelOnHitVersionSource {
        bytes: Vec<u8>,
        cancellation_source: CancellationSource,
        skip_versions: AtomicUsize,
        armed: AtomicBool,
    }

    impl CancelOnHitVersionSource {
        fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
            Self {
                bytes,
                cancellation_source,
                skip_versions: AtomicUsize::new(0),
                armed: AtomicBool::new(false),
            }
        }

        fn arm_after_cache_enter(&self) {
            // The part lookup and `read_part` perform three freshness checks
            // before a hit's post-entry check can run.
            self.skip_versions.store(3, Ordering::SeqCst);
            self.armed.store(true, Ordering::SeqCst);
        }
    }

    impl ReadAt for CancelOnHitVersionSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
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
            if self.armed.load(Ordering::SeqCst)
                && self
                    .skip_versions
                    .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |remaining| {
                        remaining.checked_sub(1)
                    })
                    .is_err()
            {
                self.armed.store(false, Ordering::SeqCst);
                self.cancellation_source.cancel();
            }
            Ok(SourceVersion::new(43, 0))
        }
    }

    struct OverReportingSource {
        bytes: Vec<u8>,
        overreport: AtomicBool,
    }

    impl OverReportingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                overreport: AtomicBool::new(false),
            }
        }
    }

    impl ReadAt for OverReportingSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            if self.overreport.load(Ordering::SeqCst) {
                return Ok(output.len().saturating_add(1));
            }
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
            Ok(SourceVersion::new(93, 0))
        }
    }

    struct OverReportingSink {
        calls: usize,
        accepted: usize,
    }

    impl Write for OverReportingSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.calls = self.calls.saturating_add(1);
            Ok(bytes.len().saturating_add(1))
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
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

    struct CancelDuringPayloadSource {
        bytes: Vec<u8>,
        payload_offset: usize,
        cancellation_source: CancellationSource,
        armed: AtomicBool,
    }

    struct CancelDuringOpenSource {
        bytes: Vec<u8>,
        cancellation_source: CancellationSource,
        reads: AtomicUsize,
        cancel_after: usize,
    }

    impl CancelDuringOpenSource {
        fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
            Self {
                bytes,
                cancellation_source,
                reads: AtomicUsize::new(0),
                cancel_after: 1,
            }
        }
    }

    impl ReadAt for CancelDuringOpenSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            if self.reads.fetch_add(1, Ordering::SeqCst) + 1 >= self.cancel_after {
                self.cancellation_source.cancel();
            }
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(95, 0))
        }
    }

    impl CancelDuringPayloadSource {
        fn new(
            bytes: Vec<u8>,
            payload_offset: usize,
            cancellation_source: CancellationSource,
        ) -> Self {
            Self {
                bytes,
                payload_offset,
                cancellation_source,
                armed: AtomicBool::new(true),
            }
        }
    }

    impl ReadAt for CancelDuringPayloadSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            if offset == self.payload_offset && self.armed.swap(false, Ordering::SeqCst) {
                // The bytes have been decompressed into the loader's private
                // allocation, but the publication checks must reject them.
                self.cancellation_source.cancel();
            }
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(79, 0))
        }
    }

    struct ChangeDuringPayloadSource {
        bytes: Vec<u8>,
        payload_offset: usize,
        revision: AtomicU64,
        armed: AtomicBool,
    }

    impl ChangeDuringPayloadSource {
        fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
            Self {
                bytes,
                payload_offset,
                revision: AtomicU64::new(0),
                armed: AtomicBool::new(false),
            }
        }
    }

    impl ReadAt for ChangeDuringPayloadSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            if offset == self.payload_offset && self.armed.swap(false, Ordering::SeqCst) {
                self.revision.fetch_add(1, Ordering::SeqCst);
            }
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(88, self.revision.load(Ordering::SeqCst)))
        }
    }

    fn managed_context_with_cancellation(
        memory: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let (budget, cancellation_source, context) =
            managed_context_with_resources(memory, u64::MAX, u64::MAX, u64::MAX);
        (budget, cancellation_source, context)
    }

    fn managed_context_with_resources(
        memory: u64,
        input_bytes: u64,
        objects: u64,
        work: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "opc-source-cache-test",
            Limits::new(memory, input_bytes, u64::MAX, objects, u64::MAX, work),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancellation_source, context)
    }

    fn managed_context(memory: u64) -> (Budget, ExecutionContext) {
        let (budget, cancellation_source, context) = managed_context_with_cancellation(memory);
        drop(cancellation_source);
        (budget, context)
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

    fn archive_with_document_relationships(document_relationships: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships())
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<before/>")
            .unwrap();
        writer
            .write_stored("word/_rels/document.xml.rels", document_relationships)
            .unwrap();
        writer
            .write_stored("custom/orphan.xml", b"<orphan/>")
            .unwrap();
        writer.write_stored("scratch.bin", b"untouched").unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn document_relationships() -> &'static [u8] {
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://remove.invalid/" TargetMode="External"/><Relationship Id="rInternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../custom/orphan.xml"/></Relationships>"#
    }

    #[test]
    fn source_artifact_paths_reject_overreported_read_counts_without_output() {
        let source = Arc::new(OverReportingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        )));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let package = SourceBackedPackage::from_read_at(read_at).unwrap();
        let artifact = package.source_artifact();
        source.overreport.store(true, Ordering::SeqCst);

        assert!(matches!(
            artifact.fingerprint(),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));

        let mut output = Vec::new();
        assert!(matches!(
            artifact.write_to_stream(&mut output),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn source_artifact_copy_rejects_overreporting_sink_without_false_progress() {
        let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
            archive_bytes(root_relationships(), b"<before/>", false),
        )))
        .unwrap();
        let artifact = package.source_artifact();
        let mut sink = OverReportingSink {
            calls: 0,
            accepted: 0,
        };

        assert!(matches!(
            artifact.write_to_stream(&mut sink),
            Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
        ));
        assert_eq!(sink.calls, 1);
        assert_eq!(sink.accepted, 0);
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

    struct CancelAfterWriteSink {
        cancellation_source: CancellationSource,
        bytes: Vec<u8>,
        cancelled: bool,
    }

    impl Write for CancelAfterWriteSink {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.bytes.extend_from_slice(bytes);
            if !self.cancelled {
                self.cancelled = true;
                self.cancellation_source.cancel();
            }
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
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
    fn managed_cache_reserves_declared_payload_and_releases_on_package_drop() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"managed payload",
            false,
        )));
        let (budget, context) = managed_context(1024);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();
        assert_eq!(budget.used(Resource::Memory), 0);

        let data = package.main_document_part().unwrap().data().unwrap();
        assert_eq!(data.as_bytes(), b"managed payload");
        assert_eq!(
            budget.used(Resource::Memory),
            b"managed payload".len() as u64
        );
        let diagnostics = package.cache_diagnostics();
        assert!(diagnostics.budget_managed);
        assert_eq!(diagnostics.budget_reservation_failures, 0);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            b"managed payload".len() as u64
        );
        assert_eq!(
            diagnostics.budget_memory_used,
            b"managed payload".len() as u64
        );

        drop(data);
        // The clean cache entry owns the reservation until the package drops.
        assert_eq!(
            budget.used(Resource::Memory),
            b"managed payload".len() as u64
        );
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_input_bytes_are_exact_for_chunked_reads_and_cache_hits_are_free() {
        const DOCUMENT: &[u8] = b"short-read physical input accounting";
        let source = Arc::new(CountingSource::chunked(
            archive_bytes(root_relationships(), DOCUMENT, false),
            3,
        ));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        assert_eq!(
            budget.used(Resource::InputBytes),
            source.read_bytes.load(Ordering::SeqCst)
        );
        let first = package.main_document_part().unwrap().data().unwrap();
        let after_cold = source.read_bytes.load(Ordering::SeqCst);
        assert_eq!(budget.used(Resource::InputBytes), after_cold);
        let second = package.main_document_part().unwrap().data().unwrap();
        assert!(first.shares_allocation_with(&second));
        assert_eq!(source.read_bytes.load(Ordering::SeqCst), after_cold);
        assert_eq!(
            package.cache_diagnostics().budget_input_bytes_used,
            after_cold
        );
        drop(second);
        drop(first);
        drop(package);
        // InputBytes and Work are cumulative; only retained resources release
        // when the package and its handles are dropped.
        assert_eq!(budget.used(Resource::InputBytes), after_cold);
    }

    #[test]
    fn managed_input_budget_refusal_counts_only_the_terminal_read_reservation() {
        const DOCUMENT: &[u8] = b"input reservation refusal accounting";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context.clone(),
        )
        .unwrap();
        let input_before = budget.used(Resource::InputBytes);
        context
            .consume(Resource::InputBytes, u64::MAX - input_before)
            .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::InputBytes
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::InputBytes), u64::MAX);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_work_one_under_refuses_payload_before_physical_io() {
        const DOCUMENT: &[u8] = b"work preflight one under";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, (DOCUMENT.len() - 1) as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Work
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Work), 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_object_one_under_refuses_payload_before_physical_io() {
        const DOCUMENT: &[u8] = b"object preflight one under";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        // archive_bytes(false) contains four non-directory members and one
        // package-level catalog owner is retained by SourceBackedPackage.
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, 5, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Objects
        ));
        // The catalog reservation is retained by the package; the one-under
        // payload-object preflight happens before any ordinary payload read.
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Objects), 5);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_failed_cold_load_consumes_work_and_input_but_releases_retained_objects() {
        const DOCUMENT: &[u8] = b"managed failed cold-load accounting";
        let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let position = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        bytes[position] ^= 0xff;
        let source = Arc::new(CountingSource::new(bytes));
        let (budget, _cancellation_source, context) =
            managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.read_bytes.load(Ordering::SeqCst);
        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::ZipError(_))
        ));
        let reads_after = source.read_bytes.load(Ordering::SeqCst);
        assert!(reads_after > reads_before);
        assert_eq!(budget.used(Resource::InputBytes), reads_after);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(package.cache_diagnostics().budget_cache_reserved_objects, 0);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_cancellation_between_cache_admission_and_source_read_is_typed() {
        const DOCUMENT: &[u8] = b"cancel before managed source read";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, cancellation_source, context) = managed_context_with_cancellation(4096);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let part = package.main_document_part().unwrap();
        let index = part.index;
        let entry_id = package.parts[index].entry_id;
        let declared = package
            .archive
            .metadata_for(entry_id)
            .unwrap()
            .uncompressed_size();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let flight = match package.cache.enter(entry_id, declared).unwrap() {
            CacheAccess::Loader(flight) => flight,
            CacheAccess::Hit(_) | CacheAccess::Waiter(_) | CacheAccess::Bypass(_) => {
                panic!("fresh managed Part must become the loader")
            },
        };
        // The loader has charged its bounded cold-load resources, but has not
        // entered SourceReader yet. This is the exact race boundary where a
        // cancellation must survive ZIP's std::io::Error conversion.
        cancellation_source.cancel();
        let error = package
            .load_part(index, entry_id, Some(declared), Some(flight), None)
            .unwrap_err();
        assert!(matches!(error, OpcError::Cancelled));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(package.cache_diagnostics().budget_cache_reserved_objects, 0);
        drop(package);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_cancellation_before_preservation_source_read_is_typed() {
        const DOCUMENT: &[u8] = b"<before/>";
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let source = Arc::new(CancelOnHitVersionSource::new(
            archive_bytes(root_relationships(), DOCUMENT, false),
            cancellation_source,
        ));
        let budget = Budget::root(
            "opc-source-cache-preservation-cancel-test",
            Limits::new(4096, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(4096).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let target = package.main_document_part().unwrap().partname().clone();
        source.arm_after_cache_enter();
        let mut output = Vec::new();
        let error = package
            .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
            .unwrap_err();
        assert!(matches!(error, OpcError::Cancelled));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_constructor_honors_pre_cancellation_without_source_reads() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"cancelled before open",
            false,
        )));
        let reads_before = source.reads.load(Ordering::SeqCst);
        let budget = Budget::root(
            "opc-source-cache-cancel-test",
            Limits::new(1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        cancellation_source.cancel();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(1024).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget, cancellation, execution_limits);

        assert!(matches!(
            SourceBackedPackage::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            ),
            Err(OpcError::Cancelled)
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    }

    #[test]
    fn managed_cache_hit_honors_cancellation_without_releasing_cached_budget() {
        const DOCUMENT: &[u8] = b"managed cancellation hit";
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let source = Arc::new(CancelOnHitVersionSource::new(
            archive_bytes(root_relationships(), DOCUMENT, false),
            cancellation_source,
        ));
        let budget = Budget::root(
            "opc-source-cache-hit-cancel-test",
            Limits::new(
                DOCUMENT.len() as u64,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let first = package.main_document_part().unwrap().data().unwrap();
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        source.arm_after_cache_enter();

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::Cancelled)
        ));
        assert_eq!(package.cache_diagnostics().hits, 1);
        // Cancellation rejects the handle request, but never steals the
        // clean entry's reservation from the package-owned cache.
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        drop(first);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_cache_eviction_releases_unpinned_reservation_before_next_read() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            false,
        )));
        let (budget, context) = managed_context(64);
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                ReadLimits::default(),
                SourceCacheLimits::new(9, 2).unwrap(),
                context,
            )
            .unwrap();
        let first_name = package.parts[0].partname.clone();
        let second_name = package.parts[1].partname.clone();
        let first = package.part(&first_name).unwrap().data().unwrap();
        assert_eq!(budget.used(Resource::Memory), b"document".len() as u64);
        drop(first);

        let second = package.part(&second_name).unwrap().data().unwrap();
        assert_eq!(second.as_bytes(), b"<orphan/>");
        assert_eq!(budget.used(Resource::Memory), b"<orphan/>".len() as u64);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.evictions, 1);
        assert_eq!(diagnostics.retained_bytes, b"<orphan/>".len());
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            b"<orphan/>".len() as u64
        );
        drop(second);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_cache_does_not_evict_externally_pinned_entry_and_bypasses_retention() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            false,
        )));
        let (budget, context) = managed_context(64);
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                ReadLimits::default(),
                SourceCacheLimits::new(b"document".len(), 1).unwrap(),
                context,
            )
            .unwrap();
        let first_name = package.parts[0].partname.clone();
        let second_name = package.parts[1].partname.clone();
        let first = package.part(&first_name).unwrap().data().unwrap();
        let first_error = first.into_arc().expect_err("managed Arc escape must fail");
        assert_eq!(budget.used(Resource::Memory), b"document".len() as u64);
        assert_eq!(
            package.cache_diagnostics().retained_bytes,
            b"document".len()
        );
        let second = package.part(&second_name).unwrap().data().unwrap();

        assert!(matches!(first_error, OpcError::ManagedPartDataArcEscape));
        assert_eq!(first.as_bytes(), b"document");
        assert_eq!(second.as_bytes(), b"<orphan/>");
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.retained_bytes, b"document".len());
        assert_eq!(diagnostics.bypasses, 1);
        assert_eq!(
            budget.used(Resource::Memory),
            (b"document".len() + b"<orphan/>".len()) as u64
        );
        assert_eq!(budget.used(Resource::Objects), 7);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 1);
        assert!(
            package
                .cache
                .state
                .lock()
                .unwrap()
                .entries
                .contains_key(&package.parts[0].entry_id)
        );

        drop(second);
        drop(first);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }

    #[test]
    fn managed_budget_rejects_before_payload_io_and_reports_content_free_failure() {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"payload too large",
            false,
        )));
        let (budget, context) = managed_context((b"payload too large".len() - 1) as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);
        let error = package.main_document_part().unwrap().data().unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(budget.used(Resource::Memory), 0);
        let diagnostics = package.cache_diagnostics();
        assert!(diagnostics.budget_managed);
        assert_eq!(diagnostics.budget_reservation_failures, 2);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.retained_bytes, 0);
    }

    #[test]
    fn managed_cache_respects_hierarchical_parent_memory_limit() {
        const DOCUMENT: &[u8] = b"hierarchical budget payload";
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let root = Budget::root(
            "opc-source-cache-root",
            Limits::new(
                (DOCUMENT.len() - 1) as u64,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let child = root.child(
            "opc-source-cache-child",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        drop(cancellation_source);
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(child, cancellation, execution_limits);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let reads_before = source.reads.load(Ordering::SeqCst);

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(root.used(Resource::Memory), 0);
        assert_eq!(package.cache_diagnostics().budget_reservation_failures, 2);
    }

    #[test]
    fn managed_sibling_caches_compete_for_parent_memory_before_payload_io() {
        const DOCUMENT: &[u8] = b"sibling parent budget payload";
        let root = Budget::root(
            "opc-source-cache-sibling-root",
            Limits::new(
                DOCUMENT.len() as u64,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let first_budget = root.child(
            "opc-source-cache-sibling-first",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let second_budget = root.child(
            "opc-source-cache-sibling-second",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
            0,
        )
        .unwrap();
        let (first_cancellation_source, first_cancellation) = CancellationSource::pair();
        let (second_cancellation_source, second_cancellation) = CancellationSource::pair();
        let first_context =
            ExecutionContext::new(first_budget, first_cancellation, execution_limits);
        let second_context =
            ExecutionContext::new(second_budget, second_cancellation, execution_limits);
        drop(first_cancellation_source);
        drop(second_cancellation_source);
        let first_source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let second_source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let first_package = SourceBackedPackage::from_read_at_with_execution_context(
            first_source,
            ReadLimits::default(),
            first_context,
        )
        .unwrap();
        let second_package = SourceBackedPackage::from_read_at_with_execution_context(
            second_source.clone(),
            ReadLimits::default(),
            second_context,
        )
        .unwrap();
        let reads_before = second_source.reads.load(Ordering::SeqCst);
        let first = first_package.main_document_part().unwrap().data().unwrap();
        assert_eq!(root.used(Resource::Memory), DOCUMENT.len() as u64);
        drop(first);

        assert!(matches!(
            second_package.main_document_part().unwrap().data(),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Memory
        ));
        assert_eq!(second_source.reads.load(Ordering::SeqCst), reads_before);
        assert_eq!(root.used(Resource::Memory), DOCUMENT.len() as u64);
        assert_eq!(
            second_package
                .cache_diagnostics()
                .budget_reservation_failures,
            2
        );
        drop(second_package);
        drop(first_package);
        assert_eq!(root.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_same_part_waiters_share_one_reservation_and_flight() {
        const DOCUMENT: &[u8] = b"managed single-flight payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
        let (budget, context) = managed_context(DOCUMENT.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let start = Arc::new(Barrier::new(3));
        let (first, second) = std::thread::scope(|scope| {
            let package = &package;
            let first_start = Arc::clone(&start);
            let first_task = scope.spawn(move || {
                first_start.wait();
                package.main_document_part().unwrap().data().unwrap()
            });
            let second_start = Arc::clone(&start);
            let second_task = scope.spawn(move || {
                second_start.wait();
                package.main_document_part().unwrap().data().unwrap()
            });
            start.wait();
            std::thread::sleep(Duration::from_millis(10));
            let diagnostics = package.cache_diagnostics();
            assert_eq!(diagnostics.in_flight_loads, 1);
            assert_eq!(
                diagnostics.budget_cache_reserved_bytes,
                DOCUMENT.len() as u64
            );
            assert_eq!(diagnostics.budget_cache_reserved_objects, 2);
            assert_eq!(budget.used(Resource::Objects), 7);
            (first_task.join().unwrap(), second_task.join().unwrap())
        });
        assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
        assert!(first.shares_allocation_with(&second));
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 1);
        assert_eq!(diagnostics.waiter_joins, 1);
        assert_eq!(diagnostics.successful_loads, 1);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            DOCUMENT.len() as u64
        );
        assert_eq!(diagnostics.budget_cache_reserved_objects, 1);
        drop(first);
        drop(second);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_waiter_cancellation_does_not_block_or_publish_loader_payload() {
        const DOCUMENT: &[u8] = b"managed waiter cancellation payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(DOCUMENT.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let start = Arc::new(Barrier::new(3));
        let (first, second) = std::thread::scope(|scope| {
            let package = &package;
            let first_start = Arc::clone(&start);
            let first_task = scope.spawn(move || {
                first_start.wait();
                package.main_document_part().unwrap().data()
            });
            let second_start = Arc::clone(&start);
            let second_task = scope.spawn(move || {
                second_start.wait();
                package.main_document_part().unwrap().data()
            });
            start.wait();
            std::thread::sleep(Duration::from_millis(10));
            assert_eq!(package.cache_diagnostics().in_flight_loads, 1);
            cancellation_source.cancel();
            (first_task.join().unwrap(), second_task.join().unwrap())
        });

        assert!(matches!(first, Err(OpcError::Cancelled)));
        assert!(matches!(second, Err(OpcError::Cancelled)));
        assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.in_flight_loads, 0);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert!(budget.used(Resource::InputBytes) > 0);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_source_change_drops_reservation_and_does_not_retain_payload() {
        const DOCUMENT: &[u8] = b"source changes during managed read";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
        let (budget, context) = managed_context(1024);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        source.armed.store(true, Ordering::SeqCst);

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::SourceChanged { .. })
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
        assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
        assert!(budget.used(Resource::InputBytes) > 0);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    }

    #[test]
    fn managed_cancellation_after_decompression_prevents_publication() {
        const DOCUMENT: &[u8] = b"managed prepublication cancellation payload";
        let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
        let payload_offset = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(DOCUMENT.len() as u64);
        let source = Arc::new(CancelDuringPayloadSource::new(
            bytes,
            payload_offset,
            cancellation_source,
        ));
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            ReadLimits::default(),
            context,
        )
        .unwrap();

        assert!(matches!(
            package.main_document_part().unwrap().data(),
            Err(OpcError::Cancelled)
        ));
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.in_flight_loads, 0);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_open_checks_cancellation_after_indexing_before_catalog_publication() {
        let (budget, cancellation_source, context) = managed_context_with_cancellation(u64::MAX);
        let source = Arc::new(CancelDuringOpenSource::new(
            archive_bytes(root_relationships(), b"open cancellation", false),
            cancellation_source,
        ));
        assert!(matches!(
            SourceBackedPackage::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            ),
            Err(OpcError::Cancelled)
        ));
        assert!(source.reads.load(Ordering::SeqCst) > 0);
        assert_eq!(budget.used(Resource::Memory), 0);
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
                    .unwrap()
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
                    .unwrap()
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
        cache
            .complete_bypass_success(
                first_id,
                CachedPayload {
                    bytes: Arc::clone(&first),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        assert!(Arc::ptr_eq(
            &cache.state.lock().unwrap().entries[&first_id].payload.bytes,
            &first
        ));
        drop(first);
        cache
            .complete_bypass_success(
                second_id,
                CachedPayload {
                    bytes: Arc::new(vec![3, 4]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
        assert!(cache.state.lock().unwrap().entries.contains_key(&second_id));

        let entry_limited = PartCache::new(SourceCacheLimits::new(10, 1).unwrap());
        entry_limited
            .complete_bypass_success(
                first_id,
                CachedPayload {
                    bytes: Arc::new(vec![1, 2]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
        entry_limited
            .complete_bypass_success(
                second_id,
                CachedPayload {
                    bytes: Arc::new(vec![3, 4]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
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

        cache
            .complete_bypass_success(
                first_id,
                CachedPayload {
                    bytes: Arc::new(vec![0, 0, 0, 0]),
                    reservation: None,
                    object_reservation: None,
                },
            )
            .unwrap();
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
    fn external_relationship_removal_overlay_changes_only_owner_and_rels_member() {
        let source_bytes = archive_with_document_relationships(document_relationships());
        let source_raw = raw_records(&source_bytes);
        let document = PackURI::new("/word/document.xml").unwrap();
        let package =
            SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(source_bytes))).unwrap();
        let mut output = Vec::new();
        package
            .write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                b"<after/>".to_vec(),
                vec!["rExternal".to_owned()],
            )
            .unwrap();

        let reopened = OpcPackage::from_bytes(&output).unwrap();
        let main = reopened.get_part(&document).unwrap();
        assert_eq!(main.blob(), b"<after/>");
        assert!(main.rels().get("rExternal").is_none());
        assert_eq!(
            main.rels()
                .get("rInternal")
                .unwrap()
                .target_partname()
                .unwrap(),
            PackURI::new("/custom/orphan.xml").unwrap()
        );
        let output_raw = raw_records(&output);
        assert_eq!(output_raw.len(), source_raw.len());
        for (name, source_record) in source_raw {
            if matches!(
                name.as_slice(),
                b"word/document.xml" | b"word/_rels/document.xml.rels"
            ) {
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
    fn external_relationship_removal_overlay_refuses_ids_and_limits_before_output() {
        let source_bytes = archive_with_document_relationships(document_relationships());
        let document = PackURI::new("/word/document.xml").unwrap();
        for (ids, expected) in [
            (vec!["missing".to_owned()], "missing external relationship"),
            (
                vec!["rExternal".to_owned(), "rExternal".to_owned()],
                "duplicate relationship",
            ),
            (vec!["rInternal".to_owned()], "internal relationship"),
        ] {
            let package = SourceBackedPackage::from_read_at(Arc::new(CountingSource::new(
                source_bytes.clone(),
            )))
            .unwrap();
            let mut output = Vec::new();
            let error = package
                .write_part_overlay_with_external_relationship_removals_to_stream(
                    &mut output,
                    &document,
                    b"<after/>".to_vec(),
                    ids,
                )
                .unwrap_err();
            match expected {
                "missing external relationship" => {
                    assert!(matches!(error, OpcError::RelationshipNotFound(_)));
                },
                "duplicate relationship" => {
                    assert!(matches!(error, OpcError::DuplicateRelationshipId(_)));
                },
                "internal relationship" => {
                    assert!(matches!(error, OpcError::InvalidRelationship(_)));
                },
                _ => unreachable!(),
            }
            assert!(output.is_empty());
        }

        let limits = ReadLimits::builder()
            .max_part_bytes(10)
            .unwrap()
            .build()
            .unwrap();
        let package = SourceBackedPackage::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes)),
            limits,
        )
        .unwrap();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                vec![b'x'; 11],
                vec!["rExternal".to_owned()],
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual: 11,
                maximum: 10,
            })
        ));
        assert!(output.is_empty());
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

    #[test]
    fn managed_publication_checks_before_output_and_between_copy_chunks() {
        let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
        let target = PackURI::new("/word/document.xml").unwrap();

        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(source_bytes.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        cancellation_source.cancel();
        let mut output = Vec::new();
        assert!(matches!(
            package.write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec()),
            Err(OpcError::Cancelled)
        ));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);

        let (budget, cancellation_source, context) =
            managed_context_with_cancellation(source_bytes.len() as u64);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(CountingSource::new(source_bytes)),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut sink = CancelAfterWriteSink {
            cancellation_source,
            bytes: Vec::new(),
            cancelled: false,
        };
        let error = package
            .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
            .unwrap_err();
        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert!(written > 0);
                assert!(matches!(*source, OpcError::Cancelled));
            },
            other => panic!("unexpected managed cancellation error: {other:?}"),
        }
        assert!(!sink.bytes.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}
