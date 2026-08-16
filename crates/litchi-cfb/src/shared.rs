//! Immutable, positional CFB reads backed by [`litchi_core::ReadAt`].

use crate::shared_bulk::SharedOleBulkRead;
use crate::{
    consts::{ENDOFCHAIN, MAXREGSECT, STGTY_STREAM},
    directory_name::directory_name_data,
    file::{DirectoryEntry, OleError, OleFile, ParsedOleIndex},
};
use litchi_core::{ExecutionContext, ReadAt, SourceVersion};
use std::{
    cmp::Ordering,
    io::{self, Read, Seek, SeekFrom},
    sync::{
        Arc, Condvar, Mutex,
        atomic::{AtomicU64, Ordering as AtomicOrdering},
    },
};

/// Bounded ingress settings for [`SharedOleFile`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SharedOleFileLimits {
    max_input_bytes: u64,
}

impl SharedOleFileLimits {
    /// Largest CFB input accepted by the default shared reader.
    pub const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

    /// Creates a finite input ceiling for one positional CFB source.
    ///
    /// # Errors
    ///
    /// Returns an error if the ceiling is zero or exceeds the CFB shared
    /// reader's hard ingress ceiling.
    pub fn new(max_input_bytes: u64) -> Result<Self, OleError> {
        if max_input_bytes == 0 || max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(OleError::InvalidData(format!(
                "shared CFB input limit must be between 1 and {} bytes",
                Self::MAX_INPUT_BYTES
            )));
        }
        Ok(Self { max_input_bytes })
    }

    /// Maximum source length accepted before parsing begins.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }
}

impl Default for SharedOleFileLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct PendingPhysicalRange {
    physical: u64,
    output_start: usize,
    length: usize,
    sector: u32,
}

// MS-CFB routes streams strictly smaller than the 4096-byte cutoff through
// MiniFAT. Keep this bound local to the positional reader rather than adding
// another public limit: it is an invariant of the format, not caller policy.
const MINIFAT_DIRECT_READ_MAX_BYTES: u64 = 4095;
const MINIFAT_DIRECT_ROOT_SIZE_RATIO: u64 = 2;
// Keep the direct target SID and the state in one atomic word. A direct read
// may be repeated only after the exact same directory entry has completed;
// observing another SID, or an in-flight direct read, takes over permanently
// through the root Mini Stream cache. The high 32 bits carry the target SID;
// the low byte carries the state. The remaining bits stay zero so this can be
// extended without changing the synchronization layout.
const MINIFAT_STATE_MASK: u64 = 0xFF;
const MINIFAT_SID_SHIFT: u32 = 32;
const MINIFAT_CACHE_SID: u32 = u32::MAX;
const MINIFAT_DIRECT_UNCLAIMED: u8 = 0;
const MINIFAT_DIRECT_IN_FLIGHT: u8 = 1;
const MINIFAT_DIRECT_DONE: u8 = 2;
const MINIFAT_CACHE_IN_FLIGHT: u8 = 3;
const MINIFAT_CACHE_READY: u8 = 4;
const MINIFAT_CACHE_RETRY: u8 = 5;
const MINIFAT_CACHE_REQUESTED: u8 = 6;

const fn minifat_state(sid: u32, state: u8) -> u64 {
    ((sid as u64) << MINIFAT_SID_SHIFT) | state as u64
}

const fn minifat_state_sid(value: u64) -> u32 {
    (value >> MINIFAT_SID_SHIFT) as u32
}

const fn minifat_state_kind(value: u64) -> u8 {
    (value & MINIFAT_STATE_MASK) as u8
}

/// The result of one bounded direct MiniFAT read which may be observed by
/// callers that were already waiting for that read. The owner keeps the
/// source I/O outside this state lock; only the publication and waiter
/// bookkeeping are serialized here.
enum MiniFATSingleFlightStatus {
    InFlight,
    Succeeded(Vec<u8>),
    /// The owner succeeded, but there was no registered waiter (or the
    /// bounded handoff copy could not be reserved). Waiters that race into
    /// this marker retry after the owner leaves; sequential callers therefore
    /// never retain a payload copy on the common path.
    CompletedNoHandoff,
    Failed,
}

struct MiniFATSingleFlightSlot {
    sid: u32,
    owner_active: bool,
    waiters: usize,
    status: MiniFATSingleFlightStatus,
}

/// A bounded, one-operation rendezvous for an eligible MiniFAT SID.
///
/// This is deliberately a mutex/condvar rather than a queue or executor. It
/// has no worker thread and retains at most one direct payload (the format's
/// <=4095-byte MiniFAT direct-read bound) while overlapping callers consume
/// it. The payload is discarded as soon as the owner and all waiters leave.
struct MiniFATSingleFlight {
    slot: Mutex<Option<MiniFATSingleFlightSlot>>,
    wake: Condvar,
}

impl MiniFATSingleFlight {
    fn new() -> Self {
        Self {
            slot: Mutex::new(None),
            wake: Condvar::new(),
        }
    }
}

enum MiniFATDirectClaim<'file> {
    Owner(MiniFATDirectOwner<'file>),
    Waiter(MiniFATDirectWaiter<'file>),
    Cache,
}

struct MiniFATDirectOwner<'file> {
    file: &'file SharedOleFile,
    sid: u32,
    published: bool,
    success: bool,
}

impl MiniFATDirectOwner<'_> {
    fn publish_success(&mut self, payload: &[u8]) -> Result<(), OleError> {
        self.file.publish_minifat_direct(self.sid, payload, true)?;
        self.published = true;
        self.success = true;
        Ok(())
    }

    fn publish_failure(&mut self) -> Result<(), OleError> {
        self.file.publish_minifat_direct(self.sid, &[], false)?;
        self.published = true;
        self.success = false;
        Ok(())
    }
}

impl Drop for MiniFATDirectOwner<'_> {
    fn drop(&mut self) {
        self.file
            .release_minifat_direct(self.sid, self.published, self.success);
    }
}

struct MiniFATDirectWaiter<'file> {
    file: &'file SharedOleFile,
    sid: u32,
    registered: bool,
}

impl MiniFATDirectWaiter<'_> {
    fn wait(&mut self) -> Result<Option<Result<Vec<u8>, OleError>>, OleError> {
        let mut slot = self
            .file
            .minifat_singleflight
            .slot
            .lock()
            .map_err(|_error| minifat_singleflight_poisoned())?;
        loop {
            let Some(current) = slot.as_ref() else {
                return Ok(None);
            };
            if current.sid != self.sid {
                return Ok(None);
            }
            match &current.status {
                MiniFATSingleFlightStatus::InFlight => {
                    slot = self
                        .file
                        .minifat_singleflight
                        .wake
                        .wait(slot)
                        .map_err(|_error| minifat_singleflight_poisoned())?;
                },
                MiniFATSingleFlightStatus::Succeeded(payload) => {
                    return Ok(Some(clone_minifat_waiter_payload(payload)));
                },
                MiniFATSingleFlightStatus::CompletedNoHandoff
                | MiniFATSingleFlightStatus::Failed => {
                    if current.owner_active {
                        slot = self
                            .file
                            .minifat_singleflight
                            .wake
                            .wait(slot)
                            .map_err(|_error| minifat_singleflight_poisoned())?;
                    } else {
                        return Ok(None);
                    }
                },
            }
        }
    }
}

impl Drop for MiniFATDirectWaiter<'_> {
    fn drop(&mut self) {
        if self.registered {
            self.file.release_minifat_waiter(self.sid);
            self.registered = false;
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MiniFATOpenMode {
    DirectIfEligible,
    ForceCache,
}

/// One parsed, immutable CFB view over a thread-safe positional source.
///
/// Opening this type runs the existing [`OleFile`] validation pipeline once.
/// The validation cursor is then discarded; regular stream reads address the
/// [`ReadAt`] source directly and can run concurrently without a shared seek
/// cursor or reader lock. Mini-stream bytes remain lazy and are initialized at
/// most once, with failures left retryable.
pub struct SharedOleFile {
    pub(crate) source: Arc<dyn ReadAt>,
    pub(crate) expected_version: SourceVersion,
    pub(crate) index: Arc<ParsedOleIndex>,
    /// Serializes only lazy mini-stream initialization. Regular streams never
    /// acquire this lock or any shared cursor lock.
    ministream: Mutex<Option<Arc<[u8]>>>,
    /// Coordinates target-aware direct MiniFAT opens with concurrent/root-
    /// cache opens. The cache path takes over permanently once it is selected.
    minifat_direct_state: AtomicU64,
    /// Coordinates one bounded direct read with overlapping same-SID callers.
    /// This lock is never held during positional source I/O and is independent
    /// of the root Mini Stream cache mutex, avoiding lock-order cycles.
    minifat_singleflight: MiniFATSingleFlight,
}

impl std::fmt::Debug for SharedOleFile {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SharedOleFile")
            .field("file_size", &self.index.file_size)
            .field("sector_size", &self.index.sector_size)
            .finish_non_exhaustive()
    }
}

impl SharedOleFile {
    #[cfg(test)]
    pub(crate) fn mini_stream_is_materialized(&self) -> bool {
        self.ministream
            .lock()
            .map(|cached| cached.is_some())
            .unwrap_or(true)
    }

    /// Starts an explicit, bounded bulk-read session.
    ///
    /// This is the only shared-reader API that may schedule work. Normal
    /// [`Self::open_stream`] calls remain serial and create no runtime.
    #[must_use]
    pub fn bulk_read(&self, context: ExecutionContext) -> SharedOleBulkRead<'_> {
        SharedOleBulkRead::new(self, context)
    }

    /// Opens and validates a positional source under the default finite input
    /// ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds the default limit, changes
    /// during parsing, or is not a valid CFB file.
    pub fn open(source: Arc<dyn ReadAt>) -> Result<Self, OleError> {
        Self::open_with_limits(source, SharedOleFileLimits::default())
    }

    /// Opens and validates a positional source under caller-selected limits.
    ///
    /// The source version is captured before parsing and compared after the
    /// existing cursor-based parser has completely validated allocation chains,
    /// overlaps, directory trees, and final-sector truncation rules.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds the limit, changes during
    /// parsing, or fails the normal [`OleFile`] validation rules.
    pub fn open_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SharedOleFileLimits,
    ) -> Result<Self, OleError> {
        let expected_version = source.version()?;
        let source_length = source.len();
        let observed = source.version()?;
        if observed != expected_version {
            return Err(OleError::SourceChanged {
                expected: expected_version,
                observed,
            });
        }
        let source_length = source_length?;
        if source_length > limits.max_input_bytes() {
            return Err(OleError::InvalidData(format!(
                "shared CFB input length {source_length} exceeds configured limit {}",
                limits.max_input_bytes()
            )));
        }

        let parsed = OleFile::open(ReadAtCursor::new(source.clone(), source_length));
        let observed = source.version()?;
        if observed != expected_version {
            return Err(OleError::SourceChanged {
                expected: expected_version,
                observed,
            });
        }
        let index = parsed?.into_parsed_index();

        Ok(Self {
            source,
            expected_version,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        })
    }

    /// Physical length captured while parsing this CFB file.
    #[must_use]
    pub fn file_size(&self) -> u64 {
        self.index.file_size
    }

    /// Returns the captured positional source identity after checking that it
    /// is still current.
    pub fn source_version(&self) -> Result<SourceVersion, OleError> {
        self.check_source_version()?;
        Ok(self.expected_version)
    }

    /// Returns the declared length of a stream without reading its contents.
    ///
    /// # Errors
    ///
    /// Returns an error if the path does not name a stream.
    pub fn stream_len(&self, path: &[&str]) -> Result<u64, OleError> {
        let entry = self.find_entry(path)?;
        if entry.entry_type != STGTY_STREAM {
            return Err(OleError::InvalidFormat("Not a stream".to_string()));
        }
        Ok(entry.size)
    }

    /// Returns whether an entry is present at `path`.
    #[must_use]
    pub fn exists(&self, path: &[&str]) -> bool {
        self.find_entry(path).is_ok()
    }

    /// Iterates the already-validated directory index without allocating or
    /// reading stream payloads.
    pub fn directory_entries(&self) -> impl Iterator<Item = &DirectoryEntry> {
        self.index.dir_entries.iter().filter_map(Option::as_ref)
    }

    /// Materializes one stream through immutable positional reads.
    ///
    /// Regular streams never acquire a shared lock or cursor. A selected small
    /// MiniFAT stream uses a bounded positional range when the complete root
    /// mini stream would be at least twice as large; this avoids retaining
    /// unrelated mini-stream bytes. The existing root mini-stream cache remains
    /// the path after a different target, concurrent caller, multi-target bulk
    /// read, or whenever materialization is justified. A successful eligible
    /// direct open can be repeated for the exact same directory entry while it
    /// remains the active target; a direct failure is retryable, but cache
    /// takeover is permanent.
    /// Initialization errors are not retained, so a later cached read can
    /// retry.
    ///
    /// # Errors
    ///
    /// Returns an error if the path does not name a stream, source I/O fails,
    /// or the source version changes before or after the payload read.
    pub fn open_stream(&self, path: &[&str]) -> Result<Vec<u8>, OleError> {
        self.open_stream_with_mode(path, MiniFATOpenMode::DirectIfEligible)
    }

    /// Opens a stream while forcing MiniFAT streams through the shared root
    /// Mini Stream cache. This is crate-private because bulk sessions need a
    /// stable cache policy across their scheduled requests, while the public
    /// single-stream API retains target-aware direct reads.
    pub(crate) fn open_stream_force_cache(&self, path: &[&str]) -> Result<Vec<u8>, OleError> {
        self.open_stream_with_mode(path, MiniFATOpenMode::ForceCache)
    }

    fn open_stream_with_mode(
        &self,
        path: &[&str],
        mode: MiniFATOpenMode,
    ) -> Result<Vec<u8>, OleError> {
        let (is_minifat, entry_sid, start_sector, size) = {
            let entry = self.find_entry(path)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()));
            }
            (entry.is_minifat, entry.sid, entry.start_sector, entry.size)
        };

        loop {
            self.check_source_version()?;
            let claim = if is_minifat && mode == MiniFATOpenMode::DirectIfEligible {
                self.claim_minifat_direct_mode(entry_sid, size)?
            } else {
                MiniFATDirectClaim::Cache
            };
            match claim {
                MiniFATDirectClaim::Owner(mut owner) => {
                    let result = self.read_minifat_stream_range(start_sector, size);
                    let version = self.check_source_version();
                    let result = match version {
                        Err(error) => Err(error),
                        Ok(()) => result,
                    };
                    match &result {
                        Ok(payload) => {
                            if let Err(error) = owner.publish_success(payload) {
                                drop(owner);
                                return Err(error);
                            }
                        },
                        Err(_error) => {
                            // Preserve the owner's original non-cloneable
                            // source/format failure even if the private
                            // waiter marker itself is poisoned.
                            let _ = owner.publish_failure();
                        },
                    }
                    drop(owner);
                    return result;
                },
                MiniFATDirectClaim::Waiter(mut waiter) => {
                    let result = match waiter.wait()? {
                        Some(result) => result,
                        None => {
                            drop(waiter);
                            continue;
                        },
                    };
                    drop(waiter);
                    self.check_source_version()?;
                    return result;
                },
                MiniFATDirectClaim::Cache => {
                    let result = if is_minifat {
                        self.read_minifat_stream(start_sector, size)
                    } else {
                        self.read_fat_stream(start_sector, size)
                    };
                    self.check_source_version()?;
                    return result;
                },
            }
        }
    }

    /// Reads one bounded logical range into caller-provided storage.
    ///
    /// The range must be contained by the selected stream's declared length.
    /// The destination is filled in logical stream order, even when the CFB
    /// allocation chain is fragmented. The operation follows only the
    /// sectors needed for this range and never materializes the complete FAT
    /// stream or the root mini-stream. In particular, a MiniFAT range read
    /// leaves [`Self::open_stream`]'s lazy mini-stream cache untouched.
    ///
    /// An empty destination performs stream lookup, bounds validation, and a
    /// source-version check but performs no payload read. The source version is
    /// checked both before and after a non-empty read. If a later range read or
    /// source check fails, callers must discard the destination; bytes written
    /// before the failure are not rolled back.
    ///
    /// # Errors
    ///
    /// Returns [`OleError::StreamNotFound`] when `path` does not identify an
    /// entry, an invalid-format error when it identifies a storage, an invalid
    /// data error when the requested range is outside the stream, a typed
    /// source I/O or source-version error when the positional source cannot be
    /// read consistently, or a corruption error when a validated allocation
    /// chain cannot be traversed safely.
    pub fn read_stream_range(
        &self,
        path: &[&str],
        offset: u64,
        output: &mut [u8],
    ) -> Result<(), OleError> {
        let (is_minifat, start_sector, size) = {
            let entry = self.find_entry(path)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()));
            }
            (entry.is_minifat, entry.start_sector, entry.size)
        };
        let end = offset
            .checked_add(output.len() as u64)
            .ok_or_else(|| OleError::InvalidData("stream range end overflow".to_string()))?;
        if end > size {
            return Err(OleError::InvalidData(format!(
                "stream range {offset}..{end} exceeds length {size}"
            )));
        }
        if output.is_empty() {
            return self.check_source_version();
        }

        self.check_source_version()?;
        if is_minifat {
            self.read_minifat_range(start_sector, offset, output, false, false)?;
        } else {
            let sector_size = self.index.sector_size;
            if offset == 0 && size == output.len() as u64 {
                // A full logical stream already has the exact caller-owned
                // destination required by the validated chain reader. Reuse
                // its contiguous-run batching instead of rediscovering the
                // same runs through the partial-range path.
                let required = output.len().div_ceil(sector_size);
                self.read_chain_into(
                    &self.index.fat,
                    start_sector,
                    required,
                    sector_size,
                    "FAT",
                    output,
                )?;
                return self.check_source_version();
            }
            let mut sector = start_sector;
            let first_ordinal = usize::try_from(offset / sector_size as u64).map_err(|_error| {
                OleError::InvalidData("FAT range sector does not fit usize".to_string())
            })?;
            for _ in 0..first_ordinal {
                sector = next_chain_sector(&self.index.fat, sector, "FAT")?;
                if sector == ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(
                        "FAT chain ends before stream range".to_string(),
                    ));
                }
            }
            let mut within = usize::try_from(offset % sector_size as u64).map_err(|_error| {
                OleError::InvalidData("FAT range offset does not fit usize".to_string())
            })?;
            let mut written = 0_usize;
            while written < output.len() {
                let run_start = sector;
                let run_within = within;
                let mut run_bytes = (sector_size - run_within).min(output.len() - written);
                let mut remaining = output.len() - written - run_bytes;
                let mut next_run = None;
                let mut last_sector = sector;

                while remaining > 0 {
                    let next = next_chain_sector(&self.index.fat, last_sector, "FAT")?;
                    if next == ENDOFCHAIN {
                        return Err(OleError::CorruptedFile(
                            "FAT chain ends within stream range".to_string(),
                        ));
                    }
                    let contiguous = last_sector.checked_add(1).ok_or_else(|| {
                        OleError::CorruptedFile("FAT sector index overflow".to_string())
                    })?;
                    if next != contiguous {
                        next_run = Some(next);
                        break;
                    }
                    last_sector = next;
                    let count = sector_size.min(remaining);
                    run_bytes = run_bytes.checked_add(count).ok_or_else(|| {
                        OleError::CorruptedFile("FAT range run size overflow".to_string())
                    })?;
                    remaining -= count;
                }

                let physical = (u64::from(run_start) + 1)
                    .checked_mul(sector_size as u64)
                    .and_then(|value| value.checked_add(run_within as u64))
                    .ok_or_else(|| {
                        OleError::CorruptedFile("FAT range physical offset overflow".to_string())
                    })?;
                self.source
                    .read_exact_at(physical, &mut output[written..written + run_bytes])?;
                written += run_bytes;
                if let Some(next) = next_run {
                    sector = next;
                    within = 0;
                }
            }
        }
        self.check_source_version()
    }

    fn should_read_minifat_range(&self, size: u64) -> bool {
        if size == 0 || size > MINIFAT_DIRECT_READ_MAX_BYTES {
            return false;
        }
        let Some(root) = self.index.root.as_ref() else {
            // Keep the existing cached path's error and source-check order
            // when a malformed synthetic index has no root entry.
            return false;
        };
        let Some(direct_threshold) = size.checked_mul(MINIFAT_DIRECT_ROOT_SIZE_RATIO) else {
            return false;
        };
        root.size >= direct_threshold
    }

    fn claim_minifat_direct_mode(
        &self,
        sid: u32,
        size: u64,
    ) -> Result<MiniFATDirectClaim<'_>, OleError> {
        if !self.should_read_minifat_range(size) {
            self.request_ministream_cache();
            return Ok(MiniFATDirectClaim::Cache);
        }
        loop {
            let mut slot = self
                .minifat_singleflight
                .slot
                .lock()
                .map_err(|_error| minifat_singleflight_poisoned())?;
            let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            let observed_sid = minifat_state_sid(observed);
            let observed_kind = minifat_state_kind(observed);
            match observed_kind {
                MINIFAT_CACHE_REQUESTED
                | MINIFAT_CACHE_IN_FLIGHT
                | MINIFAT_CACHE_READY
                | MINIFAT_CACHE_RETRY => {
                    drop(slot);
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_IN_FLIGHT if observed_sid != sid => {
                    drop(slot);
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_DONE if observed_sid != sid => {
                    drop(slot);
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_IN_FLIGHT => {
                    if let Some(current) = slot.as_mut()
                        && current.sid == sid
                    {
                        if let Some(next) = current.waiters.checked_add(1) {
                            current.waiters = next;
                            drop(slot);
                            return Ok(MiniFATDirectClaim::Waiter(MiniFATDirectWaiter {
                                file: self,
                                sid,
                                registered: true,
                            }));
                        }
                        drop(slot);
                        self.request_ministream_cache();
                        return Ok(MiniFATDirectClaim::Cache);
                    }
                    // The owner publishes its flight slot before exposing
                    // the atomic in-flight state. A transient empty slot can
                    // only be the release edge; retry after dropping the lock.
                    drop(slot);
                    std::thread::yield_now();
                },
                MINIFAT_DIRECT_UNCLAIMED | MINIFAT_DIRECT_DONE => {
                    if let Some(current) = slot.as_ref()
                        && current.sid != sid
                        && (current.owner_active || current.waiters > 0)
                    {
                        drop(slot);
                        self.request_ministream_cache();
                        return Ok(MiniFATDirectClaim::Cache);
                    }
                    if let Some(current) = slot.as_mut()
                        && current.sid == sid
                    {
                        if !current.owner_active
                            && matches!(
                                current.status,
                                MiniFATSingleFlightStatus::Failed
                                    | MiniFATSingleFlightStatus::CompletedNoHandoff
                            )
                        {
                            let desired = minifat_state(sid, MINIFAT_DIRECT_IN_FLIGHT);
                            if self
                                .minifat_direct_state
                                .compare_exchange(
                                    observed,
                                    desired,
                                    AtomicOrdering::AcqRel,
                                    AtomicOrdering::Acquire,
                                )
                                .is_ok()
                            {
                                // The first caller leaving a terminal marker
                                // becomes the designated retry owner. Keep
                                // existing waiters registered so they can
                                // share the retry result instead of repeatedly
                                // re-registering against the failed marker.
                                current.owner_active = true;
                                current.status = MiniFATSingleFlightStatus::InFlight;
                                drop(slot);
                                return Ok(MiniFATDirectClaim::Owner(MiniFATDirectOwner {
                                    file: self,
                                    sid,
                                    published: false,
                                    success: false,
                                }));
                            }
                            drop(slot);
                            continue;
                        }
                        if current.owner_active || current.waiters > 0 {
                            if let Some(next) = current.waiters.checked_add(1) {
                                current.waiters = next;
                                drop(slot);
                                return Ok(MiniFATDirectClaim::Waiter(MiniFATDirectWaiter {
                                    file: self,
                                    sid,
                                    registered: true,
                                }));
                            }
                            drop(slot);
                            self.request_ministream_cache();
                            return Ok(MiniFATDirectClaim::Cache);
                        }
                    }

                    let desired = minifat_state(sid, MINIFAT_DIRECT_IN_FLIGHT);
                    if self
                        .minifat_direct_state
                        .compare_exchange(
                            observed,
                            desired,
                            AtomicOrdering::AcqRel,
                            AtomicOrdering::Acquire,
                        )
                        .is_ok()
                    {
                        *slot = Some(MiniFATSingleFlightSlot {
                            sid,
                            owner_active: true,
                            waiters: 0,
                            status: MiniFATSingleFlightStatus::InFlight,
                        });
                        drop(slot);
                        return Ok(MiniFATDirectClaim::Owner(MiniFATDirectOwner {
                            file: self,
                            sid,
                            published: false,
                            success: false,
                        }));
                    }
                    drop(slot);
                },
                _ => {
                    drop(slot);
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                },
            }
        }
    }

    /// Publishes either the bounded immutable payload or a retryable failure
    /// marker to same-SID waiters. `OleError` is intentionally not retained:
    /// it is not cloneable and the owner must return its original value.
    fn publish_minifat_direct(
        &self,
        sid: u32,
        payload: &[u8],
        success: bool,
    ) -> Result<(), OleError> {
        let mut slot = self
            .minifat_singleflight
            .slot
            .lock()
            .map_err(|_error| minifat_singleflight_poisoned())?;
        if let Some(current) = slot.as_mut()
            && current.sid == sid
        {
            current.status = if success {
                if current.waiters > 0 {
                    match clone_minifat_waiter_payload(payload) {
                        Ok(payload) => MiniFATSingleFlightStatus::Succeeded(payload),
                        Err(_error) => MiniFATSingleFlightStatus::CompletedNoHandoff,
                    }
                } else {
                    MiniFATSingleFlightStatus::CompletedNoHandoff
                }
            } else {
                MiniFATSingleFlightStatus::Failed
            };
            self.minifat_singleflight.wake.notify_all();
        }
        Ok(())
    }

    fn release_minifat_direct(&self, sid: u32, published: bool, success: bool) {
        // This is an unwind-safe RAII path. If another thread poisoned the
        // private marker while unwinding, recover its guard so waiters can be
        // notified and the atomic cache policy can still settle; user-facing
        // claim/wait/publish paths return a typed poison error instead.
        let mut slot = self
            .minifat_singleflight
            .slot
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(current) = slot.as_mut()
            && current.sid == sid
        {
            if !published {
                current.status = MiniFATSingleFlightStatus::Failed;
            }
            current.owner_active = false;
            if current.waiters == 0 {
                slot.take();
            }
            self.minifat_singleflight.wake.notify_all();
        }
        drop(slot);

        let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        if minifat_state_kind(observed) == MINIFAT_DIRECT_IN_FLIGHT
            && minifat_state_sid(observed) == sid
        {
            let target = if published && success {
                minifat_state(sid, MINIFAT_DIRECT_DONE)
            } else {
                minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)
            };
            // A concurrent cached read may have taken ownership while this
            // direct read was in flight. Leave that cache state untouched.
            let _ = self.minifat_direct_state.compare_exchange(
                observed,
                target,
                AtomicOrdering::AcqRel,
                AtomicOrdering::Acquire,
            );
        }
    }

    fn release_minifat_waiter(&self, sid: u32) {
        // Waiter Drop must remain cleanup-only even after a peer panic. The
        // explicit claim/wait operations report poison as OleError; this
        // recovery is solely to prevent a leaked waiter from stranding the
        // bounded flight marker.
        let mut slot = self
            .minifat_singleflight
            .slot
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(current) = slot.as_mut()
            && current.sid == sid
        {
            current.waiters = current.waiters.saturating_sub(1);
            if current.waiters == 0 && !current.owner_active {
                slot.take();
            }
        }
    }

    fn request_ministream_cache(&self) {
        loop {
            let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            let kind = minifat_state_kind(observed);
            if matches!(
                kind,
                MINIFAT_CACHE_REQUESTED | MINIFAT_CACHE_IN_FLIGHT | MINIFAT_CACHE_READY
            ) {
                return;
            }
            let desired = minifat_state(MINIFAT_CACHE_SID, MINIFAT_CACHE_REQUESTED);
            if self
                .minifat_direct_state
                .compare_exchange(
                    observed,
                    desired,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                return;
            }
        }
    }

    fn begin_ministream_cache(&self) -> u64 {
        loop {
            let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            let kind = minifat_state_kind(state);
            if kind == MINIFAT_CACHE_READY {
                return state;
            }
            if kind == MINIFAT_CACHE_IN_FLIGHT {
                // Cache initialization is serialized by `ministream`; a
                // second caller cannot observe this state after acquiring the
                // mutex unless an earlier initializer left it stranded. Do
                // not spin while holding the mutex.
                return state;
            }
            let desired = minifat_state(MINIFAT_CACHE_SID, MINIFAT_CACHE_IN_FLIGHT);
            if self
                .minifat_direct_state
                .compare_exchange(
                    state,
                    desired,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                return state;
            }
        }
    }

    fn finish_ministream_cache(&self, _previous: u64, success: bool) {
        let in_flight = minifat_state(MINIFAT_CACHE_SID, MINIFAT_CACHE_IN_FLIGHT);
        let target = if success {
            minifat_state(MINIFAT_CACHE_SID, MINIFAT_CACHE_READY)
        } else {
            // Once a caller has selected the root cache, a failed cache load
            // must remain retryable through that cache. Returning to the
            // direct state could reintroduce a second bounded read after a
            // concurrent/different target already observed this path.
            minifat_state(MINIFAT_CACHE_SID, MINIFAT_CACHE_RETRY)
        };
        let _ = self.minifat_direct_state.compare_exchange(
            in_flight,
            target,
            AtomicOrdering::AcqRel,
            AtomicOrdering::Acquire,
        );
    }

    fn read_minifat_stream_range(&self, start_sector: u32, size: u64) -> Result<Vec<u8>, OleError> {
        let size = usize::try_from(size)
            .map_err(|_error| OleError::CorruptedFile("MiniFAT stream is too large".to_string()))?;
        let mut data = try_zeroed_vec(size, "MiniFAT stream data")?;
        // `open_stream` has already performed lookup and bounds validation.
        // Require the exact chain terminator here so a malformed synthetic
        // index cannot turn the bounded reader into a weaker full-read check.
        self.read_minifat_range(start_sector, 0, &mut data, true, true)?;
        Ok(data)
    }

    fn read_minifat_range(
        &self,
        start_sector: u32,
        offset: u64,
        output: &mut [u8],
        require_chain_end: bool,
        zero_fill_truncated: bool,
    ) -> Result<(), OleError> {
        let root =
            self.index.root.as_ref().ok_or_else(|| {
                OleError::CorruptedFile("mini stream has no root entry".to_string())
            })?;
        let mini_sector_size = self.index.mini_sector_size;
        let first_ordinal =
            usize::try_from(offset / mini_sector_size as u64).map_err(|_error| {
                OleError::InvalidData("MiniFAT range sector does not fit usize".to_string())
            })?;
        let mut mini_sector = start_sector;
        for _ in 0..first_ordinal {
            mini_sector = next_chain_sector(&self.index.minifat, mini_sector, "MiniFAT")?;
            if mini_sector == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "MiniFAT chain ends before stream range".to_string(),
                ));
            }
        }
        let mut within = usize::try_from(offset % mini_sector_size as u64).map_err(|_error| {
            OleError::InvalidData("MiniFAT range offset does not fit usize".to_string())
        })?;
        let mut written = 0_usize;
        // A MiniFAT chain describes logical mini-sectors, while the root
        // chain describes their physical sectors. Keep one pending span when
        // those mappings happen to be physically adjacent. This preserves the
        // logical destination order without materializing the root mini-stream
        // or requiring a temporary staging buffer.
        let mut pending: Option<PendingPhysicalRange> = None;
        while written < output.len() {
            let count = (mini_sector_size - within).min(output.len() - written);
            let current =
                match self.minifat_physical_range(root.size, mini_sector, within, count, written) {
                    Ok(current) => current,
                    Err(error) => {
                        // The previous implementation had already read the
                        // preceding chunk before discovering this mapping error.
                        // Preserve that source/error precedence.
                        self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
                        return Err(error);
                    },
                };
            if let Some(previous) = pending {
                let contiguous = match previous.physical.checked_add(previous.length as u64) {
                    Some(contiguous) => contiguous,
                    None => {
                        let error = OleError::CorruptedFile(
                            "mini-stream physical run end overflow".to_string(),
                        );
                        self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
                        return Err(error);
                    },
                };
                if current.physical == contiguous {
                    let length = match previous.length.checked_add(count) {
                        Some(length) => length,
                        None => {
                            let error = OleError::CorruptedFile(
                                "mini-stream run size overflow".to_string(),
                            );
                            self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
                            return Err(error);
                        },
                    };
                    pending = Some(PendingPhysicalRange {
                        physical: previous.physical,
                        output_start: previous.output_start,
                        length,
                        sector: previous.sector,
                    });
                } else {
                    self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
                    pending = Some(current);
                }
            } else {
                pending = Some(current);
            }
            written += count;
            within = 0;
            if written < output.len() {
                let next = match next_chain_sector(&self.index.minifat, mini_sector, "MiniFAT") {
                    Ok(next) => next,
                    Err(error) => {
                        // The previous implementation had already read the
                        // current chunk before following the next MiniFAT
                        // link. Flush the pending span before returning a
                        // structural error so source failures and error
                        // precedence remain unchanged.
                        self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
                        return Err(error);
                    },
                };
                mini_sector = next;
                if mini_sector == ENDOFCHAIN {
                    self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
                    return Err(OleError::CorruptedFile(
                        "MiniFAT chain ends within stream range".to_string(),
                    ));
                }
            }
        }
        self.flush_pending_range(&mut pending, output, zero_fill_truncated)?;
        if require_chain_end && !output.is_empty() {
            let next = next_chain_sector(&self.index.minifat, mini_sector, "MiniFAT")?;
            if next != ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "MiniFAT chain exceeds its declared length".to_string(),
                ));
            }
        }
        Ok(())
    }

    fn flush_pending_range(
        &self,
        pending: &mut Option<PendingPhysicalRange>,
        output: &mut [u8],
        zero_fill_truncated: bool,
    ) -> Result<(), OleError> {
        let Some(range) = pending.take() else {
            return Ok(());
        };
        let output_end = range
            .output_start
            .checked_add(range.length)
            .ok_or_else(|| {
                OleError::CorruptedFile("mini-stream output range overflow".to_string())
            })?;
        if range.physical >= self.index.file_size {
            return Err(OleError::CorruptedFile(format!(
                "Sector {} is outside the file",
                range.sector
            )));
        }
        let present = usize::try_from(
            self.index
                .file_size
                .saturating_sub(range.physical)
                .min(range.length as u64),
        )
        .map_err(|_error| OleError::CorruptedFile("sector read is too large".to_string()))?;
        if zero_fill_truncated {
            if present > 0 {
                self.source.read_exact_at(
                    range.physical,
                    &mut output[range.output_start..range.output_start + present],
                )?;
            }
        } else {
            self.source
                .read_exact_at(range.physical, &mut output[range.output_start..output_end])?;
        }
        Ok(())
    }

    fn minifat_physical_range(
        &self,
        root_size: u64,
        mini_sector: u32,
        within: usize,
        count: usize,
        output_start: usize,
    ) -> Result<PendingPhysicalRange, OleError> {
        let mini_offset = usize::try_from(mini_sector)
            .map_err(|_error| {
                OleError::CorruptedFile("mini-sector index does not fit usize".to_string())
            })?
            .checked_mul(self.index.mini_sector_size)
            .and_then(|value| value.checked_add(within))
            .ok_or_else(|| OleError::CorruptedFile("mini-sector offset overflow".to_string()))?;
        let mini_end = mini_offset
            .checked_add(count)
            .ok_or_else(|| OleError::CorruptedFile("mini-sector range end overflow".to_string()))?;
        let root_size = usize::try_from(root_size).map_err(|_error| {
            OleError::CorruptedFile("root mini-stream length does not fit usize".to_string())
        })?;
        if mini_end > root_size {
            return Err(OleError::CorruptedFile(
                "mini-sector range exceeds root mini stream".to_string(),
            ));
        }
        let root_ordinal = mini_offset / self.index.sector_size;
        let root_within = mini_offset % self.index.sector_size;
        let root_sector = *self.index.root_chain.get(root_ordinal).ok_or_else(|| {
            OleError::CorruptedFile("mini-sector is outside the root FAT chain".to_string())
        })?;
        let physical = (u64::from(root_sector) + 1)
            .checked_mul(self.index.sector_size as u64)
            .and_then(|value| value.checked_add(root_within as u64))
            .ok_or_else(|| {
                OleError::CorruptedFile("mini-stream range physical offset overflow".to_string())
            })?;
        Ok(PendingPhysicalRange {
            physical,
            output_start,
            length: count,
            sector: root_sector,
        })
    }

    pub(crate) fn check_source_version(&self) -> Result<(), OleError> {
        let observed = self.source.version()?;
        if observed == self.expected_version {
            Ok(())
        } else {
            Err(OleError::SourceChanged {
                expected: self.expected_version,
                observed,
            })
        }
    }

    fn read_fat_stream(&self, start_sector: u32, size: u64) -> Result<Vec<u8>, OleError> {
        let size = usize::try_from(size)
            .map_err(|_error| OleError::CorruptedFile("FAT stream is too large".to_string()))?;
        let required = size.div_ceil(self.index.sector_size);
        let mut data = try_zeroed_vec(size, "FAT stream data")?;
        self.read_chain_into(
            &self.index.fat,
            start_sector,
            required,
            self.index.sector_size,
            "FAT",
            &mut data,
        )?;
        Ok(data)
    }

    fn read_minifat_stream(&self, start_sector: u32, size: u64) -> Result<Vec<u8>, OleError> {
        // Request cache ownership before acquiring the initialization mutex.
        // Automatic noneligible paths, forced bulk paths, and internal cache
        // callers therefore all publish the takeover before another direct
        // claimant can observe the old state.
        self.request_ministream_cache();
        let ministream = {
            let mut cached = self.ministream.lock().map_err(|_error| {
                OleError::InvalidData("shared mini-stream cache is poisoned".to_string())
            })?;
            // Enter cache mode even when a previous caller already populated
            // the bytes. This is important for a forced bulk read that follows
            // a successful direct read: after that point future automatic
            // opens must not return to direct mode.
            let previous = self.begin_ministream_cache();
            if cached.is_none() {
                // Do not publish failed initialization: a transient source
                // I/O error leaves a subsequent cache read free to retry.
                let loaded: Result<Arc<[u8]>, OleError> = (|| {
                    self.check_source_version()?;
                    let loaded = self.load_ministream()?;
                    self.check_source_version()?;
                    Ok(loaded)
                })();
                let loaded = match loaded {
                    Ok(loaded) => loaded,
                    Err(error) => {
                        self.finish_ministream_cache(previous, false);
                        return Err(error);
                    },
                };
                *cached = Some(loaded);
            }
            self.finish_ministream_cache(previous, true);
            cached.as_ref().cloned().ok_or_else(|| {
                OleError::CorruptedFile("shared mini-stream cache is missing".to_string())
            })?
        };
        let size = usize::try_from(size)
            .map_err(|_error| OleError::CorruptedFile("MiniFAT stream is too large".to_string()))?;
        let required = size.div_ceil(self.index.mini_sector_size);
        let mut data = try_zeroed_vec(size, "MiniFAT stream data")?;
        let mut sector = start_sector;

        for index in 0..required {
            let sector_index = usize::try_from(sector).map_err(|_error| {
                OleError::CorruptedFile("Mini sector index does not fit usize".to_string())
            })?;
            let position = sector_index
                .checked_mul(self.index.mini_sector_size)
                .ok_or_else(|| {
                    OleError::CorruptedFile("Mini sector offset overflow".to_string())
                })?;
            let end = position
                .checked_add(self.index.mini_sector_size)
                .ok_or_else(|| OleError::CorruptedFile("Mini sector end overflow".to_string()))?;
            if end > ministream.len() {
                return Err(OleError::CorruptedFile(
                    "Mini sector out of bounds".to_string(),
                ));
            }
            let output_start = index
                .checked_mul(self.index.mini_sector_size)
                .ok_or_else(|| {
                    OleError::CorruptedFile("MiniFAT stream offset overflow".to_string())
                })?;
            let output_end = output_start
                .checked_add(self.index.mini_sector_size)
                .unwrap_or(usize::MAX)
                .min(data.len());
            data[output_start..output_end].copy_from_slice(
                &ministream[position..position + output_end.saturating_sub(output_start)],
            );
            sector = next_chain_sector(&self.index.minifat, sector, "MiniFAT")?;
            if index + 1 == required {
                if sector != ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT chain exceeds its declared length".to_string(),
                    ));
                }
            } else if sector == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "MiniFAT chain ends before its declared length".to_string(),
                ));
            }
        }
        Ok(data)
    }

    fn load_ministream(&self) -> Result<Arc<[u8]>, OleError> {
        let root = self
            .index
            .root
            .as_ref()
            .ok_or_else(|| OleError::CorruptedFile("No root entry".to_string()))?;
        self.read_fat_stream(root.start_sector, root.size)
            .map(Vec::into)
    }

    fn read_chain_into(
        &self,
        table: &[u32],
        start_sector: u32,
        required: usize,
        sector_size: usize,
        table_name: &str,
        output: &mut [u8],
    ) -> Result<(), OleError> {
        if required == 0 {
            return Ok(());
        }
        if start_sector >= MAXREGSECT || required > table.len() {
            return Err(OleError::CorruptedFile(format!(
                "invalid {table_name} stream chain"
            )));
        }

        let mut sector = start_sector;
        let mut completed = 0usize;
        while completed < required {
            let run_start = sector;
            let mut run_count = 0usize;
            let next_after_run;
            loop {
                run_count = run_count.checked_add(1).ok_or_else(|| {
                    OleError::CorruptedFile("CFB sector run count overflow".to_string())
                })?;
                let next = next_chain_sector(table, sector, table_name)?;
                let total = completed.checked_add(run_count).ok_or_else(|| {
                    OleError::CorruptedFile("CFB sector count overflow".to_string())
                })?;
                if total == required {
                    next_after_run = next;
                    break;
                }
                if next == ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(format!(
                        "{table_name} chain ends before its declared length"
                    )));
                }
                if next
                    != sector.checked_add(1).ok_or_else(|| {
                        OleError::CorruptedFile("CFB sector index overflow".to_string())
                    })?
                {
                    next_after_run = next;
                    break;
                }
                sector = next;
            }
            if completed + run_count == required && next_after_run != ENDOFCHAIN {
                return Err(OleError::CorruptedFile(format!(
                    "{table_name} chain exceeds its declared length"
                )));
            }

            let output_start = completed
                .checked_mul(sector_size)
                .ok_or_else(|| OleError::CorruptedFile("CFB stream offset overflow".to_string()))?;
            let run_bytes = run_count.checked_mul(sector_size).ok_or_else(|| {
                OleError::CorruptedFile("CFB sector run size overflow".to_string())
            })?;
            let output_end = output_start
                .checked_add(run_bytes)
                .unwrap_or(usize::MAX)
                .min(output.len());
            self.read_sector_run(run_start, &mut output[output_start..output_end])?;

            completed += run_count;
            sector = next_after_run;
        }
        Ok(())
    }

    fn read_sector_run(&self, start_sector: u32, output: &mut [u8]) -> Result<(), OleError> {
        if output.is_empty() {
            return Ok(());
        }
        let position = (u64::from(start_sector) + 1)
            .checked_mul(self.index.sector_size as u64)
            .ok_or_else(|| OleError::CorruptedFile("Sector offset overflow".to_string()))?;
        if position >= self.index.file_size {
            return Err(OleError::CorruptedFile(format!(
                "Sector {start_sector} is outside the file"
            )));
        }
        let present = usize::try_from(
            self.index
                .file_size
                .saturating_sub(position)
                .min(output.len() as u64),
        )
        .map_err(|_error| OleError::CorruptedFile("sector read is too large".to_string()))?;
        self.source
            .read_exact_at(position, &mut output[..present])?;
        Ok(())
    }

    pub(crate) fn find_entry(&self, path: &[&str]) -> Result<&DirectoryEntry, OleError> {
        if path.is_empty() {
            return self.index.root.as_ref().ok_or(OleError::StreamNotFound);
        }
        let root = self.index.root.as_ref().ok_or(OleError::StreamNotFound)?;
        let mut current_sid = root.sid_child;
        for (index, name) in path.iter().enumerate() {
            let entry = self.find_child_by_name(current_sid, name)?;
            if index + 1 == path.len() {
                return Ok(entry);
            }
            current_sid = entry.sid_child;
        }
        Err(OleError::StreamNotFound)
    }

    fn find_child_by_name(&self, sid: u32, name: &str) -> Result<&DirectoryEntry, OleError> {
        let target = directory_name_data(name).map_err(|_error| OleError::StreamNotFound)?;
        let mut current_sid = sid;
        for _ in 0..self.index.dir_entries.len() {
            if current_sid == ENDOFCHAIN {
                return Err(OleError::StreamNotFound);
            }
            let index = usize::try_from(current_sid).map_err(|_error| OleError::StreamNotFound)?;
            let entry = self
                .index
                .dir_entries
                .get(index)
                .and_then(Option::as_ref)
                .ok_or(OleError::StreamNotFound)?;
            let entry_name = self
                .index
                .dir_name_data
                .get(index)
                .and_then(Option::as_ref)
                .ok_or(OleError::StreamNotFound)?;
            current_sid = match target.compare(entry_name) {
                Ordering::Less => entry.sid_left,
                Ordering::Equal => return Ok(entry),
                Ordering::Greater => entry.sid_right,
            };
        }
        Err(OleError::StreamNotFound)
    }
}

fn next_chain_sector(table: &[u32], sector: u32, table_name: &str) -> Result<u32, OleError> {
    if sector >= MAXREGSECT {
        return Err(OleError::CorruptedFile(format!(
            "invalid sector marker 0x{sector:08X} in {table_name} chain"
        )));
    }
    let index = usize::try_from(sector).map_err(|_error| {
        OleError::CorruptedFile(format!("invalid sector index {sector} in {table_name}"))
    })?;
    let next = *table.get(index).ok_or_else(|| {
        OleError::CorruptedFile(format!("invalid sector index {sector} in {table_name}"))
    })?;
    if next != ENDOFCHAIN && next >= MAXREGSECT {
        return Err(OleError::CorruptedFile(format!(
            "invalid sector marker 0x{next:08X} in {table_name} chain"
        )));
    }
    Ok(next)
}

fn try_zeroed_vec(length: usize, resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| OleError::allocation(resource, source))?;
    output.resize(length, 0);
    Ok(output)
}

fn clone_minifat_waiter_payload(payload: &[u8]) -> Result<Vec<u8>, OleError> {
    if payload.len() > MINIFAT_DIRECT_READ_MAX_BYTES as usize {
        return Err(OleError::CorruptedFile(
            "MiniFAT single-flight payload exceeds its bounded size".to_string(),
        ));
    }
    let mut copy = try_zeroed_vec(payload.len(), "MiniFAT direct waiter data")?;
    copy.copy_from_slice(payload);
    Ok(copy)
}

fn minifat_singleflight_poisoned() -> OleError {
    OleError::InvalidData("shared MiniFAT single-flight state is poisoned".to_string())
}

/// A private local cursor used only to feed the existing validated parser.
struct ReadAtCursor {
    source: Arc<dyn ReadAt>,
    length: u64,
    position: u64,
}

impl ReadAtCursor {
    fn new(source: Arc<dyn ReadAt>, length: u64) -> Self {
        Self {
            source,
            length,
            position: 0,
        }
    }
}

impl Read for ReadAtCursor {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() || self.position >= self.length {
            return Ok(0);
        }
        let available = self.length - self.position;
        let count = usize::try_from(available.min(output.len() as u64))
            .map_err(|_error| io::Error::other("ReadAt cursor range is too large"))?;
        let read = self.source.read_at(self.position, &mut output[..count])?;
        if read > count {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional source reported more bytes than requested",
            ));
        }
        self.position = self
            .position
            .checked_add(read as u64)
            .ok_or_else(|| io::Error::other("ReadAt cursor position overflow"))?;
        Ok(read)
    }
}

impl Seek for ReadAtCursor {
    fn seek(&mut self, from: SeekFrom) -> io::Result<u64> {
        let base = match from {
            SeekFrom::Start(position) => {
                self.position = position;
                return Ok(position);
            },
            SeekFrom::End(offset) => i128::from(self.length) + i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
        };
        self.position = u64::try_from(base).map_err(|_error| {
            io::Error::new(io::ErrorKind::InvalidInput, "invalid ReadAt cursor seek")
        })?;
        Ok(self.position)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::{OleWriter, SharedOleBulkError};
    use litchi_core::{
        Budget, CancellationSource, CancellationToken, ExecutionContext, ExecutionError,
        ExecutionLimits, Limits, Resource,
    };
    use std::{
        io::Cursor,
        num::{NonZeroU64, NonZeroUsize},
        sync::{
            Barrier, Mutex,
            atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering as AtomicOrdering},
        },
        thread,
    };

    #[derive(Debug)]
    struct TestSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        reads: AtomicUsize,
        read_ranges: Mutex<Vec<(u64, usize)>>,
        active_reads: AtomicUsize,
        max_active_reads: AtomicUsize,
        change_on_read: AtomicBool,
        fail_next_read: AtomicBool,
        fail_all_reads: AtomicBool,
        fail_read_length: AtomicUsize,
        short_next_read: AtomicBool,
        interrupt_next_read: AtomicBool,
        panic_next_read: AtomicBool,
        cancel_on_read: AtomicBool,
        cancellation: Mutex<Option<CancellationSource>>,
        barrier: Mutex<Option<Arc<Barrier>>>,
        barrier_reads: AtomicUsize,
        next_read_barrier: Mutex<Option<Arc<Barrier>>>,
    }

    impl TestSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                reads: AtomicUsize::new(0),
                read_ranges: Mutex::new(Vec::new()),
                active_reads: AtomicUsize::new(0),
                max_active_reads: AtomicUsize::new(0),
                change_on_read: AtomicBool::new(false),
                fail_next_read: AtomicBool::new(false),
                fail_all_reads: AtomicBool::new(false),
                fail_read_length: AtomicUsize::new(0),
                short_next_read: AtomicBool::new(false),
                interrupt_next_read: AtomicBool::new(false),
                panic_next_read: AtomicBool::new(false),
                cancel_on_read: AtomicBool::new(false),
                cancellation: Mutex::new(None),
                barrier: Mutex::new(None),
                barrier_reads: AtomicUsize::new(0),
                next_read_barrier: Mutex::new(None),
            }
        }

        fn reset_read_count(&self) {
            self.reads.store(0, AtomicOrdering::SeqCst);
            self.read_ranges.lock().unwrap().clear();
        }

        fn read_ranges(&self) -> Vec<(u64, usize)> {
            self.read_ranges.lock().unwrap().clone()
        }

        fn synchronize_next_two_reads(&self) {
            *self.barrier.lock().unwrap() = Some(Arc::new(Barrier::new(2)));
            self.barrier_reads.store(0, AtomicOrdering::SeqCst);
            self.max_active_reads.store(0, AtomicOrdering::SeqCst);
        }

        fn fail_next_read_of_length(&self, length: usize) {
            self.fail_read_length.store(length, AtomicOrdering::SeqCst);
        }

        fn cancel_on_next_read(&self, source: CancellationSource) {
            *self.cancellation.lock().unwrap() = Some(source);
            self.cancel_on_read.store(true, AtomicOrdering::SeqCst);
        }

        fn block_next_read(&self) -> Arc<Barrier> {
            let barrier = Arc::new(Barrier::new(2));
            *self.next_read_barrier.lock().unwrap() = Some(Arc::clone(&barrier));
            barrier
        }

        fn panic_on_next_read(&self) {
            self.panic_next_read.store(true, AtomicOrdering::SeqCst);
        }

        fn fail_all_reads(&self) {
            self.fail_all_reads.store(true, AtomicOrdering::SeqCst);
        }
    }

    impl ReadAt for TestSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, AtomicOrdering::SeqCst);
            self.read_ranges
                .lock()
                .unwrap()
                .push((offset, output.len()));
            if let Some(barrier) = self.next_read_barrier.lock().unwrap().take() {
                barrier.wait();
            }
            if self.panic_next_read.swap(false, AtomicOrdering::SeqCst) {
                panic!("injected positional read panic");
            }
            let active = self.active_reads.fetch_add(1, AtomicOrdering::SeqCst) + 1;
            self.max_active_reads
                .fetch_max(active, AtomicOrdering::SeqCst);
            let ticket = self.barrier_reads.fetch_add(1, AtomicOrdering::SeqCst);
            let barrier = self.barrier.lock().unwrap().clone();
            if ticket < 2 {
                if let Some(barrier) = barrier {
                    barrier.wait();
                }
            }
            if self.change_on_read.swap(false, AtomicOrdering::SeqCst) {
                self.revision.store(1, AtomicOrdering::SeqCst);
            }
            if self.cancel_on_read.swap(false, AtomicOrdering::SeqCst) {
                if let Some(source) = self.cancellation.lock().unwrap().as_ref() {
                    source.cancel();
                }
            }
            if self.fail_all_reads.load(AtomicOrdering::SeqCst) {
                self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
                return Err(io::Error::other(
                    "injected persistent positional read failure",
                ));
            }
            let fail_length = self.fail_read_length.load(AtomicOrdering::SeqCst);
            if fail_length == output.len()
                && fail_length != 0
                && self
                    .fail_read_length
                    .compare_exchange(
                        fail_length,
                        0,
                        AtomicOrdering::SeqCst,
                        AtomicOrdering::SeqCst,
                    )
                    .is_ok()
            {
                self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
                return Err(io::Error::other("injected length-specific read failure"));
            }
            if self.fail_next_read.swap(false, AtomicOrdering::SeqCst) {
                self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
                return Err(io::Error::other("injected positional read failure"));
            }
            if self.interrupt_next_read.swap(false, AtomicOrdering::SeqCst) {
                self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
                return Err(io::Error::from(io::ErrorKind::Interrupted));
            }
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            let count = self.bytes.len().saturating_sub(start).min(output.len());
            let count = if self.short_next_read.swap(false, AtomicOrdering::SeqCst) {
                count.saturating_sub(1)
            } else {
                count
            };
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.active_reads.fetch_sub(1, AtomicOrdering::SeqCst);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                7,
                self.revision.load(AtomicOrdering::SeqCst),
            ))
        }
    }

    fn sample_bytes() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Small"], b"mini stream").unwrap();
        writer
            .create_stream_owned(&["Large"], vec![0xA5; 8192])
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn multi_mini_bytes() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer
            .create_stream_owned(&["Mini"], (0..200).map(|index| index as u8).collect())
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn two_mini_bytes(length: usize) -> (Vec<u8>, Vec<u8>) {
        assert!(length < 4096);
        let mut writer = OleWriter::new();
        let selected: Vec<u8> = (0..length)
            .map(|index| u8::try_from((index * 31 + 7) % 251).unwrap())
            .collect();
        writer
            .create_stream_owned(&["Selected"], selected.clone())
            .unwrap();
        writer
            .create_stream_owned(&["Other"], vec![0xD3; length])
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        (output.into_inner(), selected)
    }

    fn truncated_two_mini_file() -> (SharedOleFile, Arc<TestSource>, u64, usize) {
        let (bytes, _selected) = two_mini_bytes(4095);
        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let other = parsed.find_entry(&["Other"]).unwrap();
        let other_start = other.start_sector;
        let mini_sector_size = parsed.index.mini_sector_size;
        let sector_size = parsed.index.sector_size;
        let root_chain = parsed.index.root_chain.clone();
        let root_last = *root_chain.last().unwrap();
        let truncated_len =
            usize::try_from((u64::from(root_last) + 1) * sector_size as u64 + 3).unwrap();
        assert!(truncated_len < bytes.len());

        let mini_offset = usize::try_from(other_start).unwrap() * mini_sector_size;
        let root_sector = root_chain[mini_offset / sector_size];
        let physical_start =
            (u64::from(root_sector) + 1) * sector_size as u64 + (mini_offset % sector_size) as u64;
        let present = truncated_len.saturating_sub(usize::try_from(physical_start).unwrap());
        assert!(present < 4095);

        let mut index = match Arc::try_unwrap(parsed.index) {
            Ok(index) => index,
            Err(_index) => panic!("test owns the parsed index"),
        };
        index.file_size = u64::try_from(truncated_len).unwrap();
        let source = Arc::new(TestSource::new(bytes[..truncated_len].to_vec()));
        let expected_version = source.version().unwrap();
        let file = SharedOleFile {
            source: source.clone(),
            expected_version,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        };
        (file, source, physical_start, present)
    }

    fn large_mini_bytes(sector_size: usize, length: usize) -> (Vec<u8>, Vec<u8>) {
        let mut writer = if sector_size == 512 {
            OleWriter::new()
        } else {
            OleWriter::with_sector_size(sector_size).unwrap()
        };
        let expected: Vec<u8> = (0..length)
            .map(|index| u8::try_from((index * 31 + 7) % 251).unwrap())
            .collect();
        writer
            .create_stream_owned(&["Mini"], expected.clone())
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        (output.into_inner(), expected)
    }

    fn shared(source: Arc<TestSource>) -> SharedOleFile {
        SharedOleFile::open(source).unwrap()
    }

    fn context(
        workers: usize,
        max_tasks: usize,
        max_bytes: u64,
        min_parallel_bytes: u64,
        cancellation: CancellationToken,
        memory: u64,
    ) -> ExecutionContext {
        context_with_work(
            workers,
            max_tasks,
            max_bytes,
            min_parallel_bytes,
            cancellation,
            memory,
            1 << 20,
        )
    }

    fn context_with_work(
        workers: usize,
        max_tasks: usize,
        max_bytes: u64,
        min_parallel_bytes: u64,
        cancellation: CancellationToken,
        memory: u64,
        work: u64,
    ) -> ExecutionContext {
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(workers).unwrap(),
            NonZeroUsize::new(max_tasks).unwrap(),
            NonZeroU64::new(max_bytes).unwrap(),
            min_parallel_bytes,
        )
        .unwrap();
        ExecutionContext::new(
            Budget::root(
                "shared-cfb-test",
                Limits::new(memory, 1 << 20, 1 << 20, 100, 100, work),
            ),
            cancellation,
            limits,
        )
    }

    fn bulk_bytes() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer
            .create_stream_owned(&["First"], vec![0x11; 8192])
            .unwrap();
        writer
            .create_stream_owned(&["Second"], vec![0x22; 8192])
            .unwrap();
        writer.create_stream(&["Small"], b"mini stream").unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn preserves_existing_parser_corruption_checks() {
        let mut bytes = sample_bytes();
        bytes[0] ^= 0xFF;
        let source: Arc<dyn ReadAt> = Arc::new(TestSource::new(bytes));

        assert!(matches!(
            SharedOleFile::open(source),
            Err(OleError::NotOleFile)
        ));
    }

    #[test]
    fn rejects_input_before_parsing_when_over_the_finite_limit() {
        let source: Arc<dyn ReadAt> = Arc::new(TestSource::new(sample_bytes()));
        let limits = SharedOleFileLimits::new(512).unwrap();

        assert!(matches!(
            SharedOleFile::open_with_limits(source, limits),
            Err(OleError::InvalidData(message)) if message.contains("exceeds configured limit")
        ));
    }

    #[test]
    fn rejects_a_source_that_changes_during_structural_open() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(
            SharedOleFile::open(source),
            Err(OleError::SourceChanged { .. })
        ));
    }

    #[test]
    fn rejects_a_source_that_changes_during_payload_read() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(
            file.open_stream(&["Large"]),
            Err(OleError::SourceChanged { .. })
        ));
    }

    #[test]
    fn regular_stream_reads_are_concurrent_without_a_shared_cursor() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = Arc::new(shared(source.clone()));
        source.synchronize_next_two_reads();

        thread::scope(|scope| {
            let first = file.clone();
            let second = file.clone();
            let first_read = scope.spawn(move || first.open_stream(&["Large"]));
            let second_read = scope.spawn(move || second.open_stream(&["Large"]));
            assert_eq!(first_read.join().unwrap().unwrap().len(), 8192);
            assert_eq!(second_read.join().unwrap().unwrap().len(), 8192);
        });

        assert!(source.max_active_reads.load(AtomicOrdering::SeqCst) >= 2);
    }

    #[test]
    fn selected_ministream_reads_are_bounded_and_same_target_repeats_stay_direct() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.reset_read_count();

        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
        let after_first_small = source.reads.load(AtomicOrdering::SeqCst);
        assert_eq!(source.read_ranges().len(), 1);
        assert_eq!(source.read_ranges()[0].1, b"mini stream".len());
        assert!(!file.mini_stream_is_materialized());

        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
        assert!(source.reads.load(AtomicOrdering::SeqCst) > after_first_small);
        assert!(!file.mini_stream_is_materialized());
        let after_second_small = source.reads.load(AtomicOrdering::SeqCst);
        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
        assert!(source.reads.load(AtomicOrdering::SeqCst) > after_second_small);

        // A root containing only the selected 4095-byte stream is not twice
        // as large as that stream, so repeated reads retain the existing root
        // mini-stream cache behavior.
        let (bytes, expected) = large_mini_bytes(512, 4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        source.reset_read_count();
        assert_eq!(file.open_stream(&["Mini"]).unwrap(), expected);
        assert!(file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_READY
        );
        let after_first_cached = source.reads.load(AtomicOrdering::SeqCst);
        assert_eq!(file.open_stream(&["Mini"]).unwrap().len(), 4095);
        assert_eq!(
            source.reads.load(AtomicOrdering::SeqCst),
            after_first_cached
        );
    }

    #[test]
    fn open_stream_selected_minifat_reads_one_exact_cross_sector_range() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        let entry = file.find_entry(&["Selected"]).unwrap();
        let mini_offset = usize::try_from(entry.start_sector)
            .unwrap()
            .checked_mul(file.index.mini_sector_size)
            .unwrap();
        let root_sector = file.index.root_chain[mini_offset / file.index.sector_size];
        let physical = (u64::from(root_sector) + 1) * file.index.sector_size as u64
            + (mini_offset % file.index.sector_size) as u64;

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert_eq!(source.read_ranges(), vec![(physical, 4095)]);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn minifat_range_reads_do_not_consume_the_direct_repeat_state() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        let mut range = vec![0u8; 17];

        source.reset_read_count();
        file.read_stream_range(&["Selected"], 31, &mut range)
            .unwrap();
        assert_eq!(range, expected[31..48]);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_DIRECT_UNCLAIMED
        );

        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .into_iter()
                .filter(|(_, length)| *length == expected.len())
                .count(),
            1
        );
    }

    #[test]
    fn eligible_minifat_open_repeats_same_target_without_root_cache() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .iter()
                .filter(|(_, length)| *length == expected.len())
                .count(),
            1
        );

        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
        let direct_ranges = source
            .read_ranges()
            .into_iter()
            .filter(|(_, length)| *length == expected.len())
            .count();
        assert_eq!(direct_ranges, 2);
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .into_iter()
                .filter(|(_, length)| *length == expected.len())
                .count(),
            3
        );
    }

    #[test]
    fn same_target_repeat3_and_repeat8_keep_exact_direct_ranges() {
        for repeat in [3_usize, 8] {
            let (bytes, expected) = two_mini_bytes(4095);
            let source = Arc::new(TestSource::new(bytes));
            let file = shared(source.clone());
            source.reset_read_count();

            for _ in 0..repeat {
                assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
            }

            assert!(!file.mini_stream_is_materialized(), "repeat {repeat}");
            assert_eq!(
                source
                    .read_ranges()
                    .into_iter()
                    .filter(|(_, length)| *length == expected.len())
                    .count(),
                repeat,
                "repeat {repeat}"
            );
        }
    }

    #[test]
    fn case_aliases_resolve_to_the_same_direct_target_sid() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        let selected_sid = file.find_entry(&["Selected"]).unwrap().sid;
        let alias_sid = file.find_entry(&["sElEcTeD"]).unwrap().sid;
        assert_eq!(selected_sid, alias_sid);

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert_eq!(file.open_stream(&["sElEcTeD"]).unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .into_iter()
                .filter(|(_, length)| *length == expected.len())
                .count(),
            2
        );
    }

    #[test]
    fn different_minifat_target_takes_cache_and_a_b_a_stays_cached() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), selected);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(file.open_stream(&["Other"]).unwrap(), vec![0xD3; 4095]);
        assert!(file.mini_stream_is_materialized());
        let after_other = source.reads.load(AtomicOrdering::SeqCst);
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), selected);
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), after_other);

        let ranges = source.read_ranges();
        assert_eq!(
            ranges
                .iter()
                .filter(|(_, length)| *length == selected.len())
                .count(),
            1
        );
        assert!(ranges.iter().any(|(_, length)| *length > selected.len()));
    }

    #[test]
    fn fat_interleave_does_not_change_the_active_minifat_direct_target() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
        assert_eq!(file.open_stream(&["Large"]).unwrap(), vec![0xA5; 8192]);
        assert_eq!(file.open_stream(&["small"]).unwrap(), b"mini stream");
        assert!(!file.mini_stream_is_materialized());

        let ranges = source.read_ranges();
        assert_eq!(
            ranges
                .iter()
                .filter(|(_, length)| *length == b"mini stream".len())
                .count(),
            2
        );
        assert!(ranges.iter().any(|(_, length)| *length == 8192));
    }

    #[test]
    fn cache_failure_retries_cache_without_returning_to_direct() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), selected);
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        assert!(matches!(file.open_stream(&["Other"]), Err(OleError::Io(_))));
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_RETRY
        );

        assert_eq!(file.open_stream(&["Other"]).unwrap(), vec![0xD3; 4095]);
        assert!(file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_READY
        );
        let after_retry = source.reads.load(AtomicOrdering::SeqCst);
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), selected);
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), after_retry);

        let ranges = source.read_ranges();
        assert_eq!(
            ranges
                .iter()
                .filter(|(_, length)| *length == selected.len())
                .count(),
            1,
            "cache retry must not admit a second direct range"
        );
        assert!(
            ranges
                .iter()
                .filter(|(_, length)| *length > selected.len())
                .count()
                >= 2
        );
    }

    #[test]
    fn bulk_single_minifat_keeps_direct_width_one_and_two() {
        for workers in [1, 2] {
            let (bytes, selected) = two_mini_bytes(4095);
            let source = Arc::new(TestSource::new(bytes));
            let file = shared(source.clone());
            let (_cancel, token) = CancellationSource::pair();
            source.reset_read_count();

            let outputs = file
                .bulk_read(context(workers, 2, 8190, 1, token, 16_384))
                .read_streams(&[&["Selected"]])
                .unwrap();
            assert_eq!(outputs, vec![selected.clone()]);
            assert!(!file.mini_stream_is_materialized(), "workers {workers}");
            assert_eq!(
                source
                    .read_ranges()
                    .into_iter()
                    .filter(|(_, length)| *length == selected.len())
                    .count(),
                1,
                "workers {workers}"
            );

            let (_cancel, token) = CancellationSource::pair();
            let repeated = file
                .bulk_read(context(workers, 2, 8190, 1, token, 8_192))
                .read_streams(&[&["Selected"]])
                .unwrap();
            assert_eq!(repeated, vec![selected.clone()]);
            assert!(!file.mini_stream_is_materialized(), "workers {workers}");
            assert_eq!(
                source
                    .read_ranges()
                    .into_iter()
                    .filter(|(_, length)| *length == selected.len())
                    .count(),
                2,
                "workers {workers}"
            );
        }
    }

    #[test]
    fn bulk_multi_minifat_batch_converges_on_one_root_cache() {
        for workers in [1, 2] {
            let (bytes, selected) = two_mini_bytes(4095);
            let source = Arc::new(TestSource::new(bytes));
            let file = shared(source.clone());
            let (_cancel, token) = CancellationSource::pair();
            source.reset_read_count();

            let outputs = file
                .bulk_read(context(workers, 2, 8190, 1, token, 16_384))
                .read_streams(&[&["Selected"], &["Other"]])
                .unwrap();
            assert_eq!(outputs, vec![selected, vec![0xD3; 4095]]);
            assert!(file.mini_stream_is_materialized(), "workers {workers}");
            let after_bulk = source.reads.load(AtomicOrdering::SeqCst);
            let (_cancel, token) = CancellationSource::pair();
            assert_eq!(
                file.bulk_read(context(workers, 2, 8190, 1, token, 8_192))
                    .read_streams(&[&["Selected"]])
                    .unwrap()[0]
                    .len(),
                4095
            );
            assert_eq!(source.reads.load(AtomicOrdering::SeqCst), after_bulk);
        }
    }

    #[test]
    fn minifat_direct_thresholds_cover_zero_ratio_and_size_cutoffs() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let mut file = shared(source);

        assert!(!file.should_read_minifat_range(0));
        assert!(!file.should_read_minifat_range(MINIFAT_DIRECT_READ_MAX_BYTES + 1));

        Arc::get_mut(&mut file.index)
            .unwrap()
            .root
            .as_mut()
            .unwrap()
            .size = MINIFAT_DIRECT_READ_MAX_BYTES * MINIFAT_DIRECT_ROOT_SIZE_RATIO;
        assert!(file.should_read_minifat_range(MINIFAT_DIRECT_READ_MAX_BYTES));
        Arc::get_mut(&mut file.index)
            .unwrap()
            .root
            .as_mut()
            .unwrap()
            .size -= 1;
        assert!(!file.should_read_minifat_range(MINIFAT_DIRECT_READ_MAX_BYTES));
    }

    #[test]
    fn delayed_same_sid_waiter_and_different_sid_takeover_converge() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.synchronize_next_two_reads();
        let (first, repeat, other) = thread::scope(|scope| {
            let first_file = Arc::clone(&file);
            let first = scope.spawn(move || first_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let repeat_file = Arc::clone(&file);
            let repeat = scope.spawn(move || repeat_file.open_stream(&["sElEcTeD"]));
            let other_file = Arc::clone(&file);
            let other = scope.spawn(move || other_file.open_stream(&["Other"]));
            (
                first.join().unwrap(),
                repeat.join().unwrap(),
                other.join().unwrap(),
            )
        });

        assert_eq!(first.unwrap(), selected);
        assert_eq!(repeat.unwrap(), selected);
        assert_eq!(other.unwrap(), vec![0xD3; 4095]);
        assert!(file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .iter()
                .filter(|(_, length)| *length == 4095)
                .count(),
            1
        );
    }

    #[test]
    fn failed_direct_marker_wakes_same_sid_waiter_for_a_retry() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.fail_next_read_of_length(expected.len());
        let read_gate = source.block_next_read();

        let (first, second) = thread::scope(|scope| {
            let first_file = Arc::clone(&file);
            let first = scope.spawn(move || first_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let second_file = Arc::clone(&file);
            let second = scope.spawn(move || second_file.open_stream(&["Selected"]));
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters == 0)
            {
                thread::yield_now();
            }
            read_gate.wait();
            (first.join().unwrap(), second.join().unwrap())
        });

        let mut saw_failure = false;
        let mut saw_success = false;
        for result in [first, second] {
            match result {
                Ok(value) => {
                    assert_eq!(value, expected);
                    saw_success = true;
                },
                Err(OleError::Io(_)) => saw_failure = true,
                Err(error) => panic!("unexpected direct failure: {error:?}"),
            }
        }
        assert!(saw_failure);
        assert!(saw_success);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .iter()
                .filter(|(_, length)| *length == 4095)
                .count(),
            2
        );
    }

    #[test]
    fn persistent_io_failure_with_multiple_waiters_designates_retries() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.fail_all_reads();
        let read_gate = source.block_next_read();

        let results = thread::scope(|scope| {
            let owner_file = Arc::clone(&file);
            let owner = scope.spawn(move || owner_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let mut waiters = Vec::new();
            for _ in 0..2 {
                let waiter_file = Arc::clone(&file);
                waiters.push(scope.spawn(move || waiter_file.open_stream(&["Selected"])));
            }
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters < 2)
            {
                thread::yield_now();
            }
            read_gate.wait();
            let mut results = vec![owner.join().unwrap()];
            results.extend(waiters.into_iter().map(|waiter| waiter.join().unwrap()));
            results
        });

        assert_eq!(results.len(), 3);
        assert!(
            results
                .iter()
                .all(|result| matches!(result, Err(OleError::Io(_))))
        );
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .iter()
                .filter(|(_, length)| *length == 4095)
                .count(),
            3
        );
    }

    #[test]
    fn persistent_structural_failure_with_multiple_waiters_designates_retries() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let mut parsed = shared(source.clone());
        let start =
            usize::try_from(parsed.find_entry(&["Selected"]).unwrap().start_sector).unwrap();
        Arc::get_mut(&mut parsed.index).unwrap().minifat[start] = ENDOFCHAIN;
        let file = Arc::new(parsed);
        let read_gate = source.block_next_read();

        let results = thread::scope(|scope| {
            let owner_file = Arc::clone(&file);
            let owner = scope.spawn(move || owner_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let mut waiters = Vec::new();
            for _ in 0..2 {
                let waiter_file = Arc::clone(&file);
                waiters.push(scope.spawn(move || waiter_file.open_stream(&["Selected"])));
            }
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters < 2)
            {
                thread::yield_now();
            }
            read_gate.wait();
            let mut results = vec![owner.join().unwrap()];
            results.extend(waiters.into_iter().map(|waiter| waiter.join().unwrap()));
            results
        });

        assert_eq!(results.len(), 3);
        assert!(
            results
                .iter()
                .all(|result| matches!(result, Err(OleError::CorruptedFile(_))))
        );
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn completed_without_handoff_designates_a_retry_owner() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source);
        let entry = file.find_entry(&["Selected"]).unwrap();
        *file.minifat_singleflight.slot.lock().unwrap() = Some(MiniFATSingleFlightSlot {
            sid: entry.sid,
            owner_active: false,
            waiters: 2,
            status: MiniFATSingleFlightStatus::CompletedNoHandoff,
        });
        file.minifat_direct_state.store(
            minifat_state(entry.sid, MINIFAT_DIRECT_DONE),
            AtomicOrdering::SeqCst,
        );

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("terminal no-handoff marker should elect a retry owner")
        };
        owner.publish_failure().unwrap();
        drop(owner);
    }

    #[test]
    fn same_sid_source_change_wakes_waiter_without_cache_or_retry_read() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.change_on_read.store(true, AtomicOrdering::SeqCst);
        let read_gate = source.block_next_read();

        let (owner, waiter) = thread::scope(|scope| {
            let owner_file = Arc::clone(&file);
            let owner = scope.spawn(move || owner_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let waiter_file = Arc::clone(&file);
            let waiter = scope.spawn(move || waiter_file.open_stream(&["Selected"]));
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters == 0)
            {
                thread::yield_now();
            }
            read_gate.wait();
            (owner.join().unwrap(), waiter.join().unwrap())
        });

        assert!(matches!(owner, Err(OleError::SourceChanged { .. })));
        assert!(matches!(waiter, Err(OleError::SourceChanged { .. })));
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            source
                .read_ranges()
                .iter()
                .filter(|(_, length)| *length == expected.len())
                .count(),
            1
        );
    }

    #[test]
    fn same_sid_structural_failure_wakes_waiter_without_cache() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let mut parsed = shared(source.clone());
        let start =
            usize::try_from(parsed.find_entry(&["Selected"]).unwrap().start_sector).unwrap();
        Arc::get_mut(&mut parsed.index).unwrap().minifat[start] = ENDOFCHAIN;
        let file = Arc::new(parsed);
        let read_gate = source.block_next_read();

        let (owner, waiter) = thread::scope(|scope| {
            let owner_file = Arc::clone(&file);
            let owner = scope.spawn(move || owner_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let waiter_file = Arc::clone(&file);
            let waiter = scope.spawn(move || waiter_file.open_stream(&["Selected"]));
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters == 0)
            {
                thread::yield_now();
            }
            read_gate.wait();
            (owner.join().unwrap(), waiter.join().unwrap())
        });

        assert!(matches!(owner, Err(OleError::CorruptedFile(_))));
        assert!(matches!(waiter, Err(OleError::CorruptedFile(_))));
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn direct_owner_unwind_wakes_same_sid_waiter_via_raii() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.panic_on_next_read();
        let read_gate = source.block_next_read();

        let (owner, waiter) = thread::scope(|scope| {
            let owner_file = Arc::clone(&file);
            let owner = scope.spawn(move || owner_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let waiter_file = Arc::clone(&file);
            let waiter = scope.spawn(move || waiter_file.open_stream(&["Selected"]));
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters == 0)
            {
                thread::yield_now();
            }
            read_gate.wait();
            (owner.join(), waiter.join().unwrap())
        });

        assert!(owner.is_err());
        assert_eq!(waiter.unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn poisoned_singleflight_state_returns_a_typed_error() {
        let source = Arc::new(TestSource::new(two_mini_bytes(36).0));
        let file = Arc::new(shared(source));
        let poison_file = Arc::clone(&file);
        assert!(
            thread::spawn(move || {
                let _guard = poison_file.minifat_singleflight.slot.lock().unwrap();
                panic!("poison the private single-flight state");
            })
            .join()
            .is_err()
        );

        assert!(matches!(
            file.open_stream(&["Selected"]),
            Err(OleError::InvalidData(message))
                if message == "shared MiniFAT single-flight state is poisoned"
        ));
    }

    #[test]
    fn cache_request_published_before_direct_failure_remains_permanent() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source);
        let selected = file.find_entry(&["Selected"]).unwrap();
        let other = file.find_entry(&["Other"]).unwrap();

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(selected.sid, selected.size)
            .unwrap()
        else {
            panic!("selected stream should claim direct mode")
        };
        assert!(matches!(
            file.claim_minifat_direct_mode(other.sid, other.size)
                .unwrap(),
            MiniFATDirectClaim::Cache
        ));
        owner.publish_failure().unwrap();
        drop(owner);
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_REQUESTED
        );
        assert!(matches!(
            file.claim_minifat_direct_mode(selected.sid, selected.size)
                .unwrap(),
            MiniFATDirectClaim::Cache
        ));
    }

    #[test]
    fn concurrent_eligible_opens_share_one_direct_range_without_cache() {
        let (bytes, expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.reset_read_count();
        let read_gate = source.block_next_read();

        thread::scope(|scope| {
            let first = Arc::clone(&file);
            let second = Arc::clone(&file);
            let first = scope.spawn(move || first.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let second = scope.spawn(move || second.open_stream(&["Selected"]));
            while file
                .minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .is_none_or(|slot| slot.waiters == 0)
            {
                thread::yield_now();
            }
            read_gate.wait();
            assert_eq!(first.join().unwrap().unwrap(), expected);
            assert_eq!(second.join().unwrap().unwrap(), expected);
        });

        let ranges = source.read_ranges();
        assert_eq!(
            ranges
                .iter()
                .filter(|(_, length)| *length == expected.len())
                .count(),
            1
        );
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn same_sid_singleflight_covers_36_and_4095_with_aliases_and_isolation() {
        for length in [36_usize, 4095] {
            let (bytes, expected) = two_mini_bytes(length);
            let source = Arc::new(TestSource::new(bytes));
            let file = Arc::new(shared(source.clone()));
            source.reset_read_count();
            let read_gate = source.block_next_read();

            let (mut owner, waiters) = thread::scope(|scope| {
                let owner_file = Arc::clone(&file);
                let owner = scope.spawn(move || owner_file.open_stream(&["Selected"]));
                while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                    thread::yield_now();
                }
                let mut handles = Vec::new();
                for index in 0..7 {
                    let waiter_file = Arc::clone(&file);
                    handles.push(scope.spawn(move || {
                        if index % 2 == 0 {
                            waiter_file.open_stream(&["Selected"])
                        } else {
                            waiter_file.open_stream(&["sElEcTeD"])
                        }
                    }));
                }
                while file
                    .minifat_singleflight
                    .slot
                    .lock()
                    .unwrap()
                    .as_ref()
                    .is_none_or(|slot| slot.waiters < 7)
                {
                    thread::yield_now();
                }
                read_gate.wait();
                let owner = owner.join().unwrap().unwrap();
                let waiters = handles
                    .into_iter()
                    .map(|handle| handle.join().unwrap().unwrap())
                    .collect::<Vec<_>>();
                (owner, waiters)
            });

            assert_eq!(owner, expected);
            assert!(waiters.iter().all(|value| value == &expected));
            owner.fill(0);
            assert!(waiters.iter().all(|value| value == &expected));
            assert!(!file.mini_stream_is_materialized());
            assert_eq!(
                source
                    .read_ranges()
                    .iter()
                    .filter(|(_, read_length)| *read_length == length)
                    .count(),
                1,
                "length {length}"
            );
        }
    }

    #[test]
    fn force_cache_requests_before_a_delayed_direct_read_finishes() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        source.reset_read_count();
        source.synchronize_next_two_reads();

        let (direct, cached) = thread::scope(|scope| {
            let direct_file = Arc::clone(&file);
            let direct = scope.spawn(move || direct_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let cache_file = Arc::clone(&file);
            let cached = scope.spawn(move || cache_file.open_stream_force_cache(&["Other"]));
            (direct.join().unwrap(), cached.join().unwrap())
        });
        assert_eq!(direct.unwrap(), selected);
        assert_eq!(cached.unwrap(), vec![0xD3; 4095]);
        assert!(file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_READY
        );
        let ranges = source.read_ranges();
        assert_eq!(
            ranges
                .iter()
                .filter(|(_, length)| *length == selected.len())
                .count(),
            1
        );
        assert!(ranges.iter().any(|(_, length)| *length > selected.len()));
    }

    #[test]
    fn failed_cache_takeover_during_direct_read_leaves_retryable_cache_state() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = Arc::new(shared(source.clone()));
        let root_size = usize::try_from(file.index.root.as_ref().unwrap().size).unwrap();
        source.reset_read_count();
        source.fail_next_read_of_length(root_size);
        source.synchronize_next_two_reads();

        let (selected_result, other_result) = thread::scope(|scope| {
            let selected_file = Arc::clone(&file);
            let selected = scope.spawn(move || selected_file.open_stream(&["Selected"]));
            while source.reads.load(AtomicOrdering::SeqCst) < 1 {
                thread::yield_now();
            }
            let other_file = Arc::clone(&file);
            let other = scope.spawn(move || other_file.open_stream(&["Other"]));
            (selected.join().unwrap(), other.join().unwrap())
        });
        assert_eq!(selected_result.unwrap(), selected);
        assert!(matches!(other_result, Err(OleError::Io(_))));
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_RETRY
        );

        assert_eq!(file.open_stream(&["Other"]).unwrap(), vec![0xD3; 4095]);
        assert!(file.mini_stream_is_materialized());
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_READY
        );
        let ranges = source.read_ranges();
        assert_eq!(
            ranges
                .iter()
                .filter(|(_, length)| *length == selected.len())
                .count(),
            1,
            "cache failure must not re-admit a second direct read"
        );
        assert!(ranges.iter().any(|(_, length)| *length == root_size));
    }

    #[test]
    fn bulk_preload_materializes_once_and_repeated_eligible_open_is_cached() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        let (_cancel, token) = CancellationSource::pair();

        source.reset_read_count();
        let outputs = file
            .bulk_read(context(1, 2, 8190, 0, token, 16_384))
            .read_streams(&[&["Selected"], &["Other"]])
            .unwrap();
        assert_eq!(outputs[0], selected);
        assert_eq!(outputs[1], vec![0xD3; 4095]);
        assert!(file.mini_stream_is_materialized());
        let after_bulk = source.reads.load(AtomicOrdering::SeqCst);

        assert_eq!(file.open_stream(&["Selected"]).unwrap(), outputs[0]);
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), after_bulk);
    }

    #[test]
    fn direct_open_zero_fills_a_truncated_final_root_sector_like_the_cache() {
        let (file, source, physical_start, present) = truncated_two_mini_file();
        let direct = file.open_stream(&["Other"]).unwrap();
        assert_eq!(&direct[..present], &vec![0xD3; present]);
        assert!(direct[present..].iter().all(|&byte| byte == 0));
        assert_eq!(source.read_ranges(), vec![(physical_start, present)]);

        source.reset_read_count();
        let repeated = file.open_stream(&["Other"]).unwrap();
        assert_eq!(repeated, direct);
        assert!(!file.mini_stream_is_materialized());
        assert_eq!(source.read_ranges(), vec![(physical_start, present)]);

        source.reset_read_count();
        let cached = file
            .read_minifat_stream(file.find_entry(&["Other"]).unwrap().start_sector, 4095)
            .unwrap();
        assert_eq!(cached, direct);
        assert_eq!(&cached[..present], &vec![0xD3; present]);
        assert!(cached[present..].iter().all(|&byte| byte == 0));
    }

    #[test]
    fn public_minifat_range_keeps_truncated_eof_error_and_destination_tail() {
        let (file, source, physical_start, present) = truncated_two_mini_file();
        let mut output = vec![0xA5; 4095];

        assert!(matches!(
            file.read_stream_range(&["Other"], 0, &mut output),
            Err(OleError::Io(error)) if error.kind() == io::ErrorKind::UnexpectedEof
        ));
        assert_eq!(&output[..present], &vec![0xD3; present]);
        assert!(output[present..].iter().all(|&byte| byte == 0xA5));
        assert_eq!(
            source.read_ranges(),
            vec![
                (physical_start, output.len()),
                (
                    physical_start + u64::try_from(present).unwrap(),
                    output.len() - present,
                ),
            ]
        );
    }

    #[test]
    fn direct_open_rejects_a_mini_sector_start_outside_the_captured_file() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let mut file = shared(source.clone());
        let sector_size = file.index.sector_size as u64;
        let file_size = file.index.file_size;
        let outside = u32::try_from(file_size.div_ceil(sector_size)).unwrap();
        Arc::get_mut(&mut file.index).unwrap().root_chain[0] = outside;

        source.reset_read_count();
        assert!(matches!(
            file.open_stream(&["Selected"]),
            Err(OleError::CorruptedFile(message)) if message.contains("Sector") && message.contains("outside the file")
        ));
        assert!(source.read_ranges().is_empty());
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn open_stream_selected_minifat_preserves_fragmented_logical_order() {
        let (mut bytes, expected) = two_mini_bytes(4095);
        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let entry = parsed.find_entry(&["Selected"]).unwrap();
        let start_sector = entry.start_sector;
        let mini_sector_size = parsed.index.mini_sector_size;
        let sector_size = parsed.index.sector_size;
        let root_chain = parsed.index.root_chain.clone();
        let mini_count = expected.len().div_ceil(mini_sector_size);
        let physical_mini = |mini_sector: usize| {
            let mini_offset = mini_sector * mini_sector_size;
            let root_sector = root_chain[mini_offset / sector_size];
            (u64::from(root_sector) + 1) * sector_size as u64 + (mini_offset % sector_size) as u64
        };

        // Make the logical chain visit mini-sector 2 before mini-sector 1.
        // Swap their physical payloads too, so the logical stream remains the
        // same while the source ranges prove that physical ordering is not
        // assumed from MiniFAT ordering.
        let first = physical_mini(1) as usize;
        let second = physical_mini(2) as usize;
        for index in 0..mini_sector_size {
            bytes.swap(first + index, second + index);
        }

        let mut index = match Arc::try_unwrap(parsed.index) {
            Ok(index) => index,
            Err(_index) => panic!("test owns the parsed index"),
        };
        let mut chain: Vec<u32> = (0..mini_count)
            .map(|mini_sector| u32::try_from(start_sector as usize + mini_sector).unwrap())
            .collect();
        chain.swap(1, 2);
        for window in chain.windows(2) {
            index.minifat[window[0] as usize] = window[1];
        }
        index.minifat[*chain.last().unwrap() as usize] = ENDOFCHAIN;
        let source = Arc::new(TestSource::new(bytes));
        let expected_version = source.version().unwrap();
        let file = SharedOleFile {
            source: source.clone(),
            expected_version,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        };

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Selected"]).unwrap(), expected);
        assert_eq!(
            source.read_ranges(),
            vec![
                (physical_mini(0), mini_sector_size),
                (physical_mini(2), mini_sector_size),
                (physical_mini(1), mini_sector_size),
                (physical_mini(3), expected.len() - mini_sector_size * 3,),
            ]
        );
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn open_stream_selected_minifat_refuses_source_change_during_payload_read() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(
            file.open_stream(&["Selected"]),
            Err(OleError::SourceChanged { .. })
        ));
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn open_stream_fat_control_keeps_full_contiguous_chain_read() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let entry = file.find_entry(&["Large"]).unwrap();
        let physical = (u64::from(entry.start_sector) + 1) * file.index.sector_size as u64;

        source.reset_read_count();
        assert_eq!(file.open_stream(&["Large"]).unwrap(), vec![0xA5; 8192]);
        assert_eq!(source.read_ranges(), vec![(physical, 8192)]);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn selected_minifat_open_refuses_short_or_excess_chain_without_materialization() {
        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let mut file = shared(source.clone());
        let start = file.find_entry(&["Selected"]).unwrap().start_sector as usize;
        Arc::get_mut(&mut file.index).unwrap().minifat[start] = ENDOFCHAIN;
        assert!(matches!(
            file.open_stream(&["Selected"]),
            Err(OleError::CorruptedFile(message)) if message.contains("ends within stream range")
        ));
        assert!(!file.mini_stream_is_materialized());

        let (bytes, _expected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let mut file = shared(source);
        let start = file.find_entry(&["Selected"]).unwrap().start_sector as usize;
        let mut last = start;
        for _ in 1..64 {
            last = usize::try_from(Arc::get_mut(&mut file.index).unwrap().minifat[last]).unwrap();
        }
        Arc::get_mut(&mut file.index).unwrap().minifat[last] = last as u32;
        assert!(matches!(
            file.open_stream(&["Selected"]),
            Err(OleError::CorruptedFile(message)) if message.contains("exceeds its declared length")
        ));
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn failed_ministream_initialization_can_retry() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);

        assert!(matches!(file.open_stream(&["Small"]), Err(OleError::Io(_))));
        assert_eq!(file.open_stream(&["Small"]).unwrap(), b"mini stream");
    }

    #[test]
    fn bounded_ranges_read_only_the_requested_mini_and_fat_bytes() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mini = file.find_entry(&["Small"]).unwrap();
        let root = file.index.root.as_ref().unwrap();
        let mini_offset = usize::try_from(mini.start_sector)
            .unwrap()
            .checked_mul(file.index.mini_sector_size)
            .unwrap();
        let root_ordinal = mini_offset / file.index.sector_size;
        let mut root_sector = root.start_sector;
        for _ in 0..root_ordinal {
            root_sector = file.index.fat[root_sector as usize];
        }
        let mini_physical = (u64::from(root_sector) + 1) * file.index.sector_size as u64
            + (mini_offset % file.index.sector_size) as u64;

        source.reset_read_count();
        let mut mini_output = [0u8; 11];
        file.read_stream_range(&["Small"], 0, &mut mini_output)
            .unwrap();
        assert_eq!(&mini_output, b"mini stream");
        assert_eq!(
            source.read_ranges(),
            vec![(mini_physical, mini_output.len())]
        );
        assert!(!file.mini_stream_is_materialized());

        let large = file.find_entry(&["Large"]).unwrap();
        let first_fat_sector = large.start_sector;
        source.reset_read_count();
        let mut fat_output = [0u8; 4];
        file.read_stream_range(&["Large"], 510, &mut fat_output)
            .unwrap();
        assert_eq!(fat_output, [0xA5; 4]);
        assert_eq!(
            source.read_ranges(),
            vec![(
                (u64::from(first_fat_sector) + 1) * file.index.sector_size as u64 + 510,
                fat_output.len()
            )]
        );
    }

    #[test]
    fn contiguous_minifat_ranges_coalesce_to_one_exact_read_for_common_sector_sizes() {
        for sector_size in [512, 4_096] {
            let (bytes, expected) = large_mini_bytes(sector_size, 4_095);
            let source = Arc::new(TestSource::new(bytes));
            let file = shared(source.clone());
            let entry = file.find_entry(&["Mini"]).unwrap();
            let mini_offset = usize::try_from(entry.start_sector)
                .unwrap()
                .checked_mul(file.index.mini_sector_size)
                .unwrap();
            let root_ordinal = mini_offset / file.index.sector_size;
            let root_sector = file.index.root_chain[root_ordinal];
            let physical_start = (u64::from(root_sector) + 1) * file.index.sector_size as u64
                + (mini_offset % file.index.sector_size) as u64;

            source.reset_read_count();
            let mut output = vec![0u8; expected.len()];
            file.read_stream_range(&["Mini"], 0, &mut output).unwrap();

            assert_eq!(output, expected);
            assert_eq!(
                source.read_ranges(),
                vec![(physical_start, output.len())],
                "sector size {sector_size}"
            );
            assert!(!file.mini_stream_is_materialized());

            source.reset_read_count();
            let offset = 1;
            let length = expected.len() - 2;
            let mut partial = vec![0u8; length];
            file.read_stream_range(&["Mini"], offset as u64, &mut partial)
                .unwrap();
            assert_eq!(partial, expected[offset..offset + length]);
            assert_eq!(
                source.read_ranges(),
                vec![(physical_start + offset as u64, length)],
                "partial sector boundaries for sector size {sector_size}"
            );
        }
    }

    #[test]
    fn fragmented_minifat_chain_keeps_logical_order_and_separates_physical_runs() {
        let (mut bytes, expected) = large_mini_bytes(512, 4_095);
        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let entry = parsed.find_entry(&["Mini"]).unwrap();
        let entry_sid = usize::try_from(entry.sid).unwrap();
        let start_sector = entry.start_sector;
        let mini_sector_size = parsed.index.mini_sector_size;
        let sector_size = parsed.index.sector_size;
        let root_chain = parsed.index.root_chain.clone();
        let mini_count = expected.len().div_ceil(mini_sector_size);

        let physical_mini = |mini_sector: usize| {
            let mini_offset = mini_sector * mini_sector_size;
            let root_sector = root_chain[mini_offset / sector_size];
            (u64::from(root_sector) + 1) * sector_size as u64 + (mini_offset % sector_size) as u64
        };
        // Make the logical chain visit mini-sector 2 before mini-sector 1.
        // Swap their physical payloads too, so the logical stream remains the
        // same while the source ranges prove that physical ordering is not
        // assumed from MiniFAT ordering.
        let first = physical_mini(1) as usize;
        let second = physical_mini(2) as usize;
        for index in 0..mini_sector_size {
            bytes.swap(first + index, second + index);
        }

        let mut index = match Arc::try_unwrap(parsed.index) {
            Ok(index) => index,
            Err(_index) => panic!("test owns the parsed index"),
        };
        let mut chain: Vec<u32> = (0..mini_count)
            .map(|mini_sector| u32::try_from(start_sector as usize + mini_sector).unwrap())
            .collect();
        chain.swap(1, 2);
        for window in chain.windows(2) {
            index.minifat[window[0] as usize] = window[1];
        }
        index.minifat[*chain.last().unwrap() as usize] = ENDOFCHAIN;
        let expected_version = SourceVersion::new(7, 0);
        let source = Arc::new(TestSource::new(bytes));
        let file = SharedOleFile {
            source: source.clone(),
            expected_version,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        };

        source.reset_read_count();
        let mut output = vec![0u8; expected.len()];
        file.read_stream_range(&["Mini"], 0, &mut output).unwrap();

        assert_eq!(output, expected);
        let first_physical = physical_mini(0);
        let second_physical = physical_mini(2);
        let third_physical = physical_mini(1);
        let fourth_physical = physical_mini(3);
        assert_eq!(
            source.read_ranges(),
            vec![
                (first_physical, mini_sector_size),
                (second_physical, mini_sector_size),
                (third_physical, mini_sector_size),
                (fourth_physical, expected.len() - mini_sector_size * 3,),
            ]
        );
        assert!(!file.mini_stream_is_materialized());
        // Keep the synthetic index mutation tied to the stream we selected;
        // this also catches accidental path lookup against a different SID.
        assert_eq!(file.find_entry(&["Mini"]).unwrap().sid as usize, entry_sid);
    }

    #[test]
    fn deep_minifat_ranges_use_the_prevalidated_root_chain_index() {
        const ROOT_SECTORS: usize = 4_096;
        const OUTPUT_SECTORS: usize = 8;

        let base = shared(Arc::new(TestSource::new(sample_bytes())));
        let small_sid = usize::try_from(base.find_entry(&["Small"]).unwrap().sid).unwrap();
        let SharedOleFile { index, .. } = base;
        let mut index = match Arc::try_unwrap(index) {
            Ok(index) => index,
            Err(_index) => panic!("test owns the parsed index"),
        };

        index.root_chain = (0..ROOT_SECTORS)
            .map(|ordinal| u32::try_from(ordinal).unwrap())
            .collect();
        // The parsed root index is the only valid lookup route in this
        // synthetic deep-root shape. Any per-chunk FAT walk fails at its first
        // link, detecting the CPU-amplifying implementation without timing.
        index.fat = vec![ENDOFCHAIN; ROOT_SECTORS];
        let first_mini = (ROOT_SECTORS - OUTPUT_SECTORS)
            .checked_mul(index.sector_size / index.mini_sector_size)
            .unwrap();
        let mini_count = OUTPUT_SECTORS
            .checked_mul(index.sector_size / index.mini_sector_size)
            .unwrap();
        index.minifat = vec![ENDOFCHAIN; first_mini + mini_count];
        for ordinal in 0..mini_count - 1 {
            index.minifat[first_mini + ordinal] = u32::try_from(first_mini + ordinal + 1).unwrap();
        }
        let root = index.root.as_mut().unwrap();
        root.start_sector = 0;
        root.size = u64::try_from(ROOT_SECTORS * index.sector_size).unwrap();
        let small = index.dir_entries[small_sid].as_mut().unwrap();
        small.start_sector = u32::try_from(first_mini).unwrap();
        small.size = u64::try_from(OUTPUT_SECTORS * index.sector_size).unwrap();

        let mut bytes = vec![0_u8; (ROOT_SECTORS + 1) * index.sector_size];
        let expected: Vec<u8> = (0..usize::try_from(small.size).unwrap())
            .map(|offset| u8::try_from(offset % 251).unwrap())
            .collect();
        let physical_start = (ROOT_SECTORS - OUTPUT_SECTORS + 1) * index.sector_size;
        bytes[physical_start..physical_start + expected.len()].copy_from_slice(&expected);
        index.file_size = u64::try_from(bytes.len()).unwrap();

        let source = Arc::new(TestSource::new(bytes));
        let expected_version = source.version().unwrap();
        let file = SharedOleFile {
            source: source.clone(),
            expected_version,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        };

        let mut output = vec![0_u8; expected.len()];
        file.read_stream_range(&["Small"], 0, &mut output).unwrap();
        assert_eq!(output, expected);
        assert_eq!(file.index.root_chain.len(), ROOT_SECTORS);
        assert_eq!(file.index.fat[0], ENDOFCHAIN);

        let mut repeated = [0_u8; 64];
        file.read_stream_range(&["Small"], 0, &mut repeated)
            .unwrap();
        assert_eq!(&repeated, &expected[..repeated.len()]);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn full_fat_range_uses_one_exact_contiguous_chain_read() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let entry = file.find_entry(&["Large"]).unwrap();
        let sector_size = file.index.sector_size;
        let start = (u64::from(entry.start_sector) + 1) * sector_size as u64;

        source.reset_read_count();
        let mut output = vec![0u8; usize::try_from(entry.size).unwrap()];
        file.read_stream_range(&["Large"], 0, &mut output).unwrap();
        assert_eq!(output, vec![0xA5; 8192]);
        assert_eq!(source.read_ranges(), vec![(start, output.len())]);
    }

    #[test]
    fn bounded_ranges_cover_boundaries_and_reject_out_of_range_without_io() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());

        source.reset_read_count();
        let mut boundary = [0u8; 3];
        file.read_stream_range(&["Small"], 8, &mut boundary)
            .unwrap();
        assert_eq!(&boundary, b"eam");
        assert_eq!(source.read_ranges().len(), 1);
        assert_eq!(source.read_ranges()[0].1, boundary.len());

        source.reset_read_count();
        let mut empty = [];
        file.read_stream_range(&["Small"], 11, &mut empty).unwrap();
        assert!(source.read_ranges().is_empty());

        let mut too_far = [0u8; 1];
        assert!(matches!(
            file.read_stream_range(&["Small"], 11, &mut too_far),
            Err(OleError::InvalidData(message)) if message.contains("exceeds length")
        ));
        assert!(source.read_ranges().is_empty());
    }

    #[test]
    fn fragmented_fat_ranges_follow_logical_chain_order() {
        let mut bytes = sample_bytes();
        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let entry = parsed.find_entry(&["Large"]).unwrap();
        let sector_size = parsed.index.sector_size;
        let mut chain = Vec::new();
        let mut sector = entry.start_sector;
        for _ in 0..8_192usize.div_ceil(sector_size) {
            chain.push(sector);
            sector = parsed.index.fat[sector as usize];
        }
        let fat_sector = u32::from_le_bytes(bytes[0x4c..0x50].try_into().unwrap());
        let fat_offset = (fat_sector as usize + 1) * sector_size;
        let [first, second, third, fourth] = [chain[0], chain[1], chain[2], chain[3]];
        for (current, next) in [(first, third), (third, second), (second, fourth)] {
            let offset = fat_offset + current as usize * 4;
            bytes[offset..offset + 4].copy_from_slice(&next.to_le_bytes());
        }
        let second_offset = (second as usize + 1) * sector_size;
        let third_offset = (third as usize + 1) * sector_size;
        for index in 0..sector_size {
            bytes.swap(second_offset + index, third_offset + index);
        }

        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        source.reset_read_count();
        let mut output = vec![0u8; sector_size + 17];
        file.read_stream_range(&["Large"], 17, &mut output).unwrap();
        assert_eq!(output, vec![0xA5; sector_size + 17]);
        assert_eq!(
            source.read_ranges(),
            vec![
                (
                    (u64::from(first) + 1) * sector_size as u64 + 17,
                    sector_size - 17
                ),
                ((u64::from(third) + 1) * sector_size as u64, 34),
            ]
        );
    }

    #[test]
    fn range_source_failures_are_retryable_and_do_not_materialize_mini_stream() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        let mut output = [0u8; 4];
        assert!(matches!(
            file.read_stream_range(&["Small"], 0, &mut output),
            Err(OleError::Io(_))
        ));
        assert!(!file.mini_stream_is_materialized());
        file.read_stream_range(&["Small"], 0, &mut output).unwrap();
        assert_eq!(&output, b"mini");
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn minifat_ranges_handle_short_and_interrupted_positional_reads() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let entry = file.find_entry(&["Small"]).unwrap();
        let mini_offset = usize::try_from(entry.start_sector)
            .unwrap()
            .checked_mul(file.index.mini_sector_size)
            .unwrap();
        let root_sector = file.index.root_chain[mini_offset / file.index.sector_size];
        let physical = (u64::from(root_sector) + 1) * file.index.sector_size as u64
            + (mini_offset % file.index.sector_size) as u64;

        source.reset_read_count();
        source.short_next_read.store(true, AtomicOrdering::SeqCst);
        let mut short = [0u8; 4];
        file.read_stream_range(&["Small"], 0, &mut short).unwrap();
        assert_eq!(&short, b"mini");
        assert_eq!(
            source.read_ranges(),
            vec![(physical, short.len()), (physical + 3, 1)]
        );

        source.reset_read_count();
        source
            .interrupt_next_read
            .store(true, AtomicOrdering::SeqCst);
        let mut interrupted = [0u8; 4];
        file.read_stream_range(&["Small"], 0, &mut interrupted)
            .unwrap();
        assert_eq!(&interrupted, b"mini");
        assert_eq!(
            source.read_ranges(),
            vec![(physical, interrupted.len()), (physical, interrupted.len())]
        );
    }

    #[test]
    fn range_reads_refuse_source_changes_before_publication() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.change_on_read.store(true, AtomicOrdering::SeqCst);
        let mut output = [0u8; 4];
        assert!(matches!(
            file.read_stream_range(&["Small"], 0, &mut output),
            Err(OleError::SourceChanged { .. })
        ));
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn concurrent_range_reads_are_independent_and_source_positional() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = Arc::new(shared(source.clone()));
        source.synchronize_next_two_reads();

        thread::scope(|scope| {
            let first = Arc::clone(&file);
            let second = Arc::clone(&file);
            let first_read = scope.spawn(move || {
                let mut output = [0u8; 8];
                first
                    .read_stream_range(&["Large"], 0, &mut output)
                    .map(|()| output)
            });
            let second_read = scope.spawn(move || {
                let mut output = [0u8; 8];
                second
                    .read_stream_range(&["Large"], 0, &mut output)
                    .map(|()| output)
            });
            assert_eq!(first_read.join().unwrap().unwrap(), [0xA5; 8]);
            assert_eq!(second_read.join().unwrap().unwrap(), [0xA5; 8]);
        });
        assert!(source.max_active_reads.load(AtomicOrdering::SeqCst) >= 2);
    }

    #[test]
    fn malformed_minifat_chain_is_refused_without_root_materialization() {
        let source = Arc::new(TestSource::new(multi_mini_bytes()));
        let mut file = shared(source);
        let entry = file.find_entry(&["Mini"]).unwrap();
        let start = entry.start_sector as usize;
        Arc::get_mut(&mut file.index).unwrap().minifat[start] = ENDOFCHAIN;
        let mut output = [0u8; 1];
        assert!(matches!(
            file.read_stream_range(&["Mini"], 64, &mut output),
            Err(OleError::CorruptedFile(message)) if message.contains("ends before stream range")
        ));
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn bulk_reads_preserve_input_order_and_match_serial_reads() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let paths: &[&[&str]] = &[&["Second"], &["Small"], &["First"]];
        let (_cancel, token) = CancellationSource::pair();
        let context = context(2, 2, 16_384, 1, token, 32_768);

        let session = file.bulk_read(context);
        let bulk = session.read_streams(paths).unwrap();
        let repeated = session.read_streams(paths).unwrap();
        assert_eq!(repeated, bulk);
        assert_eq!(session.pool_build_count(), 1);
        let serial: Vec<_> = paths
            .iter()
            .map(|path| file.open_stream(path).unwrap())
            .collect();
        assert_eq!(bulk, serial);
        assert_eq!(bulk[0][0], 0x22);
        assert_eq!(bulk[2][0], 0x11);
    }

    #[test]
    fn bulk_reads_observe_preexisting_cancellation_atomically() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let (cancel, token) = CancellationSource::pair();
        cancel.cancel();

        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 8192))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::Cancelled))
        ));
    }

    #[test]
    fn bulk_preflight_failures_do_not_consume_minifat_direct_state() {
        let (bytes, selected) = two_mini_bytes(4095);
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(source.clone());
        source.reset_read_count();

        let (_cancel, token) = CancellationSource::pair();
        assert!(matches!(
            file.bulk_read(context(1, 1, 4095, 0, token, 8_192))
                .read_streams(&[&["Missing"]]),
            Err(SharedOleBulkError::Ole(OleError::StreamNotFound))
        ));
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), 0);
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_DIRECT_UNCLAIMED
        );

        let (cancel, token) = CancellationSource::pair();
        cancel.cancel();
        assert!(matches!(
            file.bulk_read(context(1, 1, 4095, 0, token, 8_192))
                .read_streams(&[&["Selected"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::Cancelled))
        ));
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), 0);
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_DIRECT_UNCLAIMED
        );

        let (_cancel, token) = CancellationSource::pair();
        let output = file
            .bulk_read(context(1, 1, 4095, 0, token, 8_192))
            .read_streams(&[&["Selected"]])
            .unwrap();
        assert_eq!(output, vec![selected]);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn bulk_reads_observe_mid_read_cancellation_atomically() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        let (cancel, token) = CancellationSource::pair();
        source.cancel_on_next_read(cancel);

        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 8192))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::Cancelled))
        ));
    }

    #[test]
    fn bulk_reads_enforce_byte_and_context_budget_caps() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let (_cancel, token) = CancellationSource::pair();
        assert!(matches!(
            file.bulk_read(context(1, 1, 4096, 0, token, 4096))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::StreamExceedsInFlightBytes { .. })
        ));

        let (_cancel, token) = CancellationSource::pair();
        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 4096))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(
                ExecutionError::ResourceLimit(_)
            ))
        ));
    }

    #[test]
    fn bulk_reads_charge_work_before_starting_payload_reads() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        source.reset_read_count();
        let (_cancel, token) = CancellationSource::pair();

        assert!(matches!(
            file.bulk_read(context_with_work(1, 1, 8192, 0, token, 8192, 4096))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::Work
        ));
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), 0);
    }

    #[test]
    fn bulk_reads_honor_task_cap_and_use_a_local_pool_when_eligible() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        let (_cancel, token) = CancellationSource::pair();
        file.bulk_read(context(1, 1, 8192, 0, token, 16_384))
            .read_streams(&[&["First"], &["Second"]])
            .unwrap();
        assert_eq!(source.max_active_reads.load(AtomicOrdering::SeqCst), 1);

        source.synchronize_next_two_reads();
        let (_cancel, token) = CancellationSource::pair();
        file.bulk_read(context(2, 2, 16_384, 1, token, 16_384))
            .read_streams(&[&["First"], &["Second"]])
            .unwrap();
        assert!(source.max_active_reads.load(AtomicOrdering::SeqCst) >= 2);
    }

    #[test]
    fn bulk_session_skips_pool_when_only_bounded_batches_are_ineligible() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let (_cancel, token) = CancellationSource::pair();
        // Aggregate bytes exceed the threshold, but the 12 KiB batch cap
        // separates the two 8 KiB streams and neither one-stream batch can
        // qualify for parallel scheduling.
        let session = file.bulk_read(context(2, 2, 12_288, 12_288, token, 16_384));

        session.read_streams(&[&["First"], &["Second"]]).unwrap();
        assert_eq!(session.pool_build_count(), 0);
    }

    #[test]
    fn bulk_reads_reject_source_change_without_returning_partial_results() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source.clone());
        source.change_on_read.store(true, AtomicOrdering::SeqCst);
        let (_cancel, token) = CancellationSource::pair();

        assert!(matches!(
            file.bulk_read(context(1, 1, 8192, 0, token, 8192))
                .read_streams(&[&["First"]]),
            Err(SharedOleBulkError::Ole(OleError::SourceChanged { .. }))
        ));
    }
}
