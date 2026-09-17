//! Immutable, positional CFB reads backed by [`litchi_core::ReadAt`].

use crate::shared_bulk::SharedOleBulkRead;
use crate::{
    consts::{ENDOFCHAIN, MAXREGSECT, STGTY_STREAM},
    directory_name::directory_name_data,
    file::{DirectoryEntry, OleError, OleFile, OleFileLimits, ParsedOleIndex},
};
use litchi_core::{ExecutionContext, ReadAt, SourceVersion};
use std::{
    cmp::Ordering,
    io::{self, Read, Seek, SeekFrom},
    sync::{
        Arc, Condvar, Mutex,
        atomic::{AtomicU64, AtomicUsize, Ordering as AtomicOrdering},
    },
};

/// Bounded ingress settings for [`SharedOleFile`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SharedOleFileLimits {
    max_input_bytes: u64,
    max_directory_bytes: u64,
}

impl SharedOleFileLimits {
    /// Largest CFB input accepted by the default shared reader.
    pub const MAX_INPUT_BYTES: u64 = OleFileLimits::MAX_INPUT_BYTES;
    /// Default directory stream ceiling used by the default shared reader.
    pub const DEFAULT_MAX_DIRECTORY_BYTES: u64 = OleFileLimits::DEFAULT_MAX_DIRECTORY_BYTES;
    /// Largest directory stream ceiling accepted by the shared reader.
    pub const MAX_DIRECTORY_BYTES: u64 = OleFileLimits::MAX_DIRECTORY_BYTES;

    /// Creates a finite input ceiling for one positional CFB source.
    ///
    /// # Errors
    ///
    /// Returns an error if the ceiling is zero or exceeds the CFB shared
    /// reader's hard ingress ceiling.
    pub fn new(max_input_bytes: u64) -> Result<Self, OleError> {
        let limits = OleFileLimits::new(max_input_bytes)?;
        Ok(Self {
            max_input_bytes: limits.max_input_bytes(),
            max_directory_bytes: limits.max_directory_bytes(),
        })
    }

    /// Maximum source length accepted before parsing begins.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Selects the maximum padded directory-stream size accepted while
    /// parsing one positional CFB source.
    pub fn with_max_directory_bytes(mut self, max_directory_bytes: u64) -> Result<Self, OleError> {
        let limits = OleFileLimits::new(self.max_input_bytes)?
            .with_max_directory_bytes(max_directory_bytes)?;
        self.max_directory_bytes = limits.max_directory_bytes();
        Ok(self)
    }

    /// Maximum padded directory-stream size accepted during parsing.
    #[must_use]
    pub const fn max_directory_bytes(self) -> u64 {
        self.max_directory_bytes
    }
}

impl Default for SharedOleFileLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
            max_directory_bytes: Self::DEFAULT_MAX_DIRECTORY_BYTES,
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

#[derive(Debug, Clone, Copy)]
enum SharedOleStreamCursorState {
    End,
    Fat { sector: u32, within: usize },
    MiniFAT { sector: u32, within: usize },
}

/// A resumable position inside one validated allocation chain of one reader.
///
/// [`SharedOleFile::read_stream_range`] and [`SharedOleFile::stream_cursor_at`]
/// walk the chain from the stream's **first** sector on every call, so a scan
/// that reaches `n` ascending offsets in one stream walks the prefix `n` times.
/// Carrying one hint across those calls through
/// [`SharedOleFile::read_stream_range_hinted`] or
/// [`SharedOleFile::stream_cursor_at_hinted`] walks each chain link once
/// instead. Both record ordinals on the same scale -- sectors for a FAT stream,
/// mini-sectors for a MiniFAT one -- so one hint may be carried across both.
///
/// A hint is three integers plus a stream identity and a borrow of the reader
/// it came from. It is **not** an index of the chain, so retaining one costs
/// nothing proportional to the stream, and it is `Copy`, so it allocates
/// nothing.
///
/// # Why a hint cannot serve the wrong bytes
///
/// Three independent properties, in order:
///
/// 1. **It cannot cross readers.** A hint is obtained only from
///    [`SharedOleFile::chain_hint`] and borrows that reader for its whole life,
///    so the reader cannot be dropped or moved while a hint to it exists and no
///    other live reader can occupy its address. A hint offered to a different
///    reader is discarded on an address comparison.
/// 2. **It cannot cross streams.** It records the directory entry SID, the
///    allocation table it indexes and the stream's first sector, and is
///    discarded unless all three match the stream being read.
/// 3. **It cannot move a read backwards.** A chain is singly linked, so a hint
///    sitting past the requested offset is discarded and the walk restarts at
///    the stream's first sector.
///
/// A discarded hint costs exactly what an unhinted read costs: the walk starts
/// at the first sector. A hint therefore only ever removes work it has already
/// done; it can never redirect a read.
///
/// It also caches no source bytes and holds nothing a source mutation can
/// invalidate. The FAT and MiniFAT tables it indexes are captured and fully
/// validated while the file is opened and are immutable for the life of the
/// reader; every read taken through a hint still fences the retained source
/// exactly as an unhinted read does.
#[derive(Debug, Clone, Copy)]
pub struct StreamChainHint<'a> {
    file: &'a SharedOleFile,
    resume: Option<ChainResume>,
}

#[derive(Debug, Clone, Copy)]
struct ChainResume {
    /// Directory entry identity. All three fields must match before a retained
    /// position is used, so a hint taken from one stream can never serve
    /// another even if the two happened to share a first sector.
    sid: u32,
    is_minifat: bool,
    start_sector: u32,
    /// Sector ordinal within the chain, and the sector found at it.
    ordinal: usize,
    sector: u32,
}

impl<'a> StreamChainHint<'a> {
    /// Where to begin walking for `ordinal`: the retained position when it
    /// belongs to this reader and this stream and is at or before `ordinal`,
    /// and the stream's first sector in every other case.
    fn resume_from(
        &self,
        file: &SharedOleFile,
        sid: u32,
        is_minifat: bool,
        start_sector: u32,
        ordinal: usize,
    ) -> (u32, usize) {
        if !std::ptr::eq(self.file, file) {
            return (start_sector, 0);
        }
        match self.resume {
            Some(resume)
                if resume.sid == sid
                    && resume.is_minifat == is_minifat
                    && resume.start_sector == start_sector
                    && resume.ordinal <= ordinal =>
            {
                (resume.sector, resume.ordinal)
            },
            _ => (start_sector, 0),
        }
    }

    /// Records a position the caller has just walked to and validated.
    fn record(
        &mut self,
        file: &SharedOleFile,
        sid: u32,
        is_minifat: bool,
        start_sector: u32,
        ordinal: usize,
        sector: u32,
    ) {
        if !std::ptr::eq(self.file, file) {
            return;
        }
        self.resume = Some(ChainResume {
            sid,
            is_minifat,
            start_sector,
            ordinal,
            sector,
        });
    }
}

/// A non-owning, forward-only cursor over one validated CFB stream.
///
/// The cursor resolves the stream directory entry once and keeps only the
/// current logical chain position. FAT and MiniFAT payloads are read directly
/// from the retained positional source; creating or using a cursor never
/// materializes the shared Mini Stream cache.
///
/// A cursor is intentionally not `Clone`. Callers that need independent
/// traversals should create independent cursors from [`SharedOleFile`].
pub struct SharedOleStreamCursor<'a> {
    file: &'a SharedOleFile,
    length: u64,
    position: u64,
    state: SharedOleStreamCursorState,
}

impl std::fmt::Debug for SharedOleStreamCursor<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SharedOleStreamCursor")
            .field("length", &self.length)
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

// MS-CFB routes streams strictly smaller than the 4096-byte cutoff through
// MiniFAT. Keep this bound local to the positional reader rather than adding
// another public limit: it is an invariant of the format, not caller policy.
/// Stands for "no directory entry" in a chain hint. A CFB directory can
/// never contain `u32::MAX` entries, so no real SID can collide with it and
/// a hint recorded under it can never be reused by a hinted read.
const MINIFAT_UNHINTED_SID: u32 = u32::MAX;
const MINIFAT_DIRECT_READ_MAX_BYTES: u64 = 4095;
const MINIFAT_DIRECT_ROOT_SIZE_RATIO: u64 = 2;
// Keep the direct target SID, flight epoch, waiter intent, and state in one
// atomic word. The high 32 bits carry the target SID; bits 10..31 carry a
// bounded 22-bit epoch; bit 8 announces waiter intent; bit 9 records slot
// presence; the low byte carries the state. The epoch prevents a delayed
// owner from releasing a later same-SID flight after an ABA transition.
const MINIFAT_STATE_MASK: u64 = 0xFF;
const MINIFAT_INTENT_BIT: u64 = 1 << 8;
const MINIFAT_SLOT_PRESENT_BIT: u64 = 1 << 9;
const MINIFAT_EPOCH_SHIFT: u32 = 10;
const MINIFAT_EPOCH_MASK: u64 = ((1 << 22) - 1) << MINIFAT_EPOCH_SHIFT;
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

const fn minifat_state_with_meta(sid: u32, state: u8, epoch: u32, intent: bool) -> u64 {
    minifat_state_with_slot(sid, state, epoch, intent, false)
}

const fn minifat_state_with_slot(
    sid: u32,
    state: u8,
    epoch: u32,
    intent: bool,
    slot_present: bool,
) -> u64 {
    ((sid as u64) << MINIFAT_SID_SHIFT)
        | (((epoch as u64) << MINIFAT_EPOCH_SHIFT) & MINIFAT_EPOCH_MASK)
        | if intent { MINIFAT_INTENT_BIT } else { 0 }
        | if slot_present {
            MINIFAT_SLOT_PRESENT_BIT
        } else {
            0
        }
        | state as u64
}

const fn minifat_state_sid(value: u64) -> u32 {
    (value >> MINIFAT_SID_SHIFT) as u32
}

const fn minifat_state_kind(value: u64) -> u8 {
    (value & MINIFAT_STATE_MASK) as u8
}

const fn minifat_state_epoch(value: u64) -> u32 {
    ((value & MINIFAT_EPOCH_MASK) >> MINIFAT_EPOCH_SHIFT) as u32
}

const fn minifat_state_intent(value: u64) -> bool {
    value & MINIFAT_INTENT_BIT != 0
}

const fn minifat_state_slot_present(value: u64) -> bool {
    value & MINIFAT_SLOT_PRESENT_BIT != 0
}

const fn next_minifat_epoch(value: u64) -> Option<u32> {
    let epoch = minifat_state_epoch(value);
    if epoch == (1 << 22) - 1 {
        None
    } else {
        Some(epoch + 1)
    }
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
    epoch: u32,
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
    /// Announces waiter intent before a caller takes the slot mutex. Owners
    /// use this atomic to decide whether the slow handoff path is necessary;
    /// the uncontended sequential direct path never locks `slot`.
    waiter_intent: AtomicUsize,
    #[cfg(test)]
    /// Test-only gates make the owner-release linearization race deterministic
    /// without adding synchronization or branches to production builds.
    release_pause: AtomicUsize,
    #[cfg(test)]
    release_pause_ready: AtomicUsize,
    #[cfg(test)]
    release_pause_continue: AtomicUsize,
    #[cfg(test)]
    claim_slow_entered: AtomicUsize,
}

impl MiniFATSingleFlight {
    fn new() -> Self {
        Self {
            slot: Mutex::new(None),
            wake: Condvar::new(),
            waiter_intent: AtomicUsize::new(0),
            #[cfg(test)]
            release_pause: AtomicUsize::new(0),
            #[cfg(test)]
            release_pause_ready: AtomicUsize::new(0),
            #[cfg(test)]
            release_pause_continue: AtomicUsize::new(0),
            #[cfg(test)]
            claim_slow_entered: AtomicUsize::new(0),
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
    epoch: u32,
    published: bool,
    success: bool,
}

impl MiniFATDirectOwner<'_> {
    fn publish_success(&mut self, payload: &[u8]) -> Result<(), OleError> {
        self.file
            .publish_minifat_direct(self.sid, self.epoch, payload, true)?;
        self.published = true;
        self.success = true;
        Ok(())
    }

    fn publish_failure(&mut self) -> Result<(), OleError> {
        self.file
            .publish_minifat_direct(self.sid, self.epoch, &[], false)?;
        self.published = true;
        self.success = false;
        Ok(())
    }
}

impl Drop for MiniFATDirectOwner<'_> {
    fn drop(&mut self) {
        self.file
            .release_minifat_direct(self.sid, self.epoch, self.published, self.success);
    }
}

struct MiniFATDirectWaiter<'file> {
    file: &'file SharedOleFile,
    sid: u32,
    epoch: u32,
    registered: bool,
    intent: bool,
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
            if current.epoch != self.epoch {
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
        if self.intent {
            self.file
                .release_minifat_waiter(self.sid, self.epoch, self.registered);
            self.intent = false;
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
    pub(crate) source_is_owned_immutable: bool,
    pub(crate) limits: SharedOleFileLimits,
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

    /// Opens bytes whose immutable ownership is retained by the CFB reader.
    ///
    /// Unlike a type-erased [`ReadAt`] adapter, an `Arc<[u8]>` cannot change
    /// while this reader and its derived plans retain a clone. Publication may
    /// use that sealed provenance internally; callers cannot mark an arbitrary
    /// positional source as immutable through a flag or version token.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds the default limit or is not a
    /// valid CFB file.
    pub fn open_owned(source: Arc<[u8]>, version: SourceVersion) -> Result<Self, OleError> {
        let source: Arc<dyn ReadAt> = Arc::new(OwnedArcSource {
            source,
            version,
            #[cfg(test)]
            read_count: None,
        });
        Self::open_source_with_limits(source, SharedOleFileLimits::default(), true)
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
        Self::open_source_with_limits(source, limits, false)
    }

    fn open_source_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SharedOleFileLimits,
        source_is_owned_immutable: bool,
    ) -> Result<Self, OleError> {
        let expected_version = source.version()?;
        // The length observation is bracketed by this capture and the
        // post-parse comparison below, exactly as the parse itself is. A third
        // observation between them would only re-prove what the post-parse one
        // proves, so it survives on the two branches that return *before* the
        // post-parse comparison can run and would otherwise report a
        // length-derived failure while the source was already stale.
        let source_length = match source.len() {
            Ok(length) => length,
            Err(error) => {
                Self::refuse_if_changed(source.as_ref(), expected_version)?;
                return Err(error.into());
            },
        };
        if source_length > limits.max_input_bytes() {
            Self::refuse_if_changed(source.as_ref(), expected_version)?;
            return Err(OleError::LimitExceeded {
                resource: "input bytes",
                observed: source_length,
                maximum: limits.max_input_bytes(),
            });
        }

        let parser_limits = OleFileLimits::new(limits.max_input_bytes())?
            .with_max_directory_bytes(limits.max_directory_bytes())?;
        let parsed = OleFile::open_with_limits(
            ReadAtCursor::new(source.clone(), source_length),
            parser_limits,
        );
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
            source_is_owned_immutable,
            limits,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        })
    }

    #[cfg(test)]
    pub(crate) fn open_owned_arc_source_for_test(
        source: Arc<[u8]>,
        read_count: Arc<AtomicUsize>,
    ) -> Result<Self, OleError> {
        let source: Arc<dyn ReadAt> = Arc::new(OwnedArcSource {
            source,
            version: SourceVersion::new(0xcafe_0181, 0),
            read_count: Some(read_count),
        });
        Self::open_source_with_limits(source, SharedOleFileLimits::default(), true)
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

    /// Returns the source identity captured when this view was opened, without
    /// observing the source.
    ///
    /// Use this only to compare against an observation the caller has already
    /// taken of the *same* source object, which [`Self::source_arc`] hands
    /// out. It proves nothing about the source on its own: it is a stored
    /// value, and a caller that needs to know the source is still current must
    /// use [`Self::source_version`] or read through this view.
    #[must_use]
    pub const fn captured_source_version(&self) -> SourceVersion {
        self.expected_version
    }

    /// Clones the exact positional source retained by this shared CFB view.
    ///
    /// This is a low-level handoff for container owners that need to compose
    /// another source-backed parser without pairing the catalog with an
    /// unrelated [`ReadAt`] implementation. The returned `Arc` points to the
    /// same source object used to validate this CFB view.
    #[must_use]
    pub fn source_arc(&self) -> Arc<dyn ReadAt> {
        Arc::clone(&self.source)
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

    /// Creates a non-owning, forward-only cursor at a logical stream offset.
    ///
    /// The directory entry is resolved once. Subsequent cursor operations use
    /// the same validated FAT or MiniFAT chain and never restart traversal
    /// from the stream's first sector. MiniFAT cursors read the selected
    /// logical bytes directly and leave the shared Mini Stream cache untouched.
    /// Construction resolves only the immutable validated catalog and chain;
    /// it publishes no source bytes and performs no source-version fence.
    /// The first and every subsequent [`SharedOleStreamCursor::read_exact`]
    /// call fences the retained source after its payload read and before
    /// publishing those bytes to the caller.
    ///
    /// Construction itself still walks the chain from the stream's first
    /// sector to reach `offset`. A caller that builds many cursors over one
    /// stream at ascending offsets should use
    /// [`Self::stream_cursor_at_hinted`], which is this function with a
    /// resumable chain position.
    ///
    /// # Errors
    ///
    /// Returns an error when `path` is not a stream, `offset` is past the
    /// declared stream length, or the selected chain cannot reach `offset`.
    pub fn stream_cursor_at(
        &self,
        path: &[&str],
        offset: u64,
    ) -> Result<SharedOleStreamCursor<'_>, OleError> {
        self.stream_cursor_at_hinted(path, offset, &mut self.chain_hint())
    }

    /// Creates a cursor, resuming the allocation-chain walk from `hint`.
    ///
    /// This is [`Self::stream_cursor_at`] with one difference: instead of
    /// walking the chain from the stream's first sector to reach `offset`, it
    /// resumes from the position `hint` retains and records where it finished.
    /// A caller that builds one cursor per item while moving forward through a
    /// stream -- resolving shared strings out of a workbook's string table, for
    /// instance -- therefore walks each chain link once for the whole scan
    /// rather than once per cursor. `stream_cursor_at` is exactly this function
    /// with a fresh hint, so there is one implementation of the lookup, the
    /// bounds check and the walk.
    ///
    /// The hint is the same value [`Self::read_stream_range_hinted`] uses and
    /// records ordinals on the same scale, so one hint may be carried across
    /// both. Every validation keeps its position and its identity: the links
    /// the hint lets the walk skip are links this same hint already walked and
    /// validated through `next_chain_sector`, and links ahead of it are walked
    /// and checked here. A hint naming a different reader, a different
    /// directory entry, a different first sector, a different allocation table,
    /// or a **later** ordinal than the one requested is discarded, and the walk
    /// restarts at the stream's first sector.
    ///
    /// # Errors
    ///
    /// Identical to [`Self::stream_cursor_at`].
    pub fn stream_cursor_at_hinted<'a>(
        &'a self,
        path: &[&str],
        offset: u64,
        hint: &mut StreamChainHint<'_>,
    ) -> Result<SharedOleStreamCursor<'a>, OleError> {
        let (sid, is_minifat, start_sector, length) = {
            let entry = self.find_entry(path)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()));
            }
            (entry.sid, entry.is_minifat, entry.start_sector, entry.size)
        };
        if offset > length {
            return Err(OleError::InvalidData(format!(
                "stream cursor offset {offset} exceeds length {length}"
            )));
        }

        let state = if offset == length {
            SharedOleStreamCursorState::End
        } else if is_minifat {
            let unit = self.index.mini_sector_size;
            let ordinal = usize::try_from(offset / unit as u64).map_err(|_error| {
                OleError::InvalidData("MiniFAT cursor sector does not fit usize".to_string())
            })?;
            let sector = cursor_chain_sector(
                &self.index.minifat,
                hint.resume_from(self, sid, true, start_sector, ordinal),
                ordinal,
                "MiniFAT",
            )?;
            hint.record(self, sid, true, start_sector, ordinal, sector);
            SharedOleStreamCursorState::MiniFAT {
                sector,
                within: usize::try_from(offset % unit as u64).map_err(|_error| {
                    OleError::InvalidData("MiniFAT cursor offset does not fit usize".to_string())
                })?,
            }
        } else {
            let unit = self.index.sector_size;
            let ordinal = usize::try_from(offset / unit as u64).map_err(|_error| {
                OleError::InvalidData("FAT cursor sector does not fit usize".to_string())
            })?;
            let sector = cursor_chain_sector(
                &self.index.fat,
                hint.resume_from(self, sid, false, start_sector, ordinal),
                ordinal,
                "FAT",
            )?;
            hint.record(self, sid, false, start_sector, ordinal, sector);
            SharedOleStreamCursorState::Fat {
                sector,
                within: usize::try_from(offset % unit as u64).map_err(|_error| {
                    OleError::InvalidData("FAT cursor offset does not fit usize".to_string())
                })?,
            }
        };
        Ok(SharedOleStreamCursor {
            file: self,
            length,
            position: offset,
            state,
        })
    }

    /// Creates an empty resumable chain position bound to this reader.
    ///
    /// Pass the same hint to every [`Self::read_stream_range_hinted`] and
    /// [`Self::stream_cursor_at_hinted`] call of one forward scan over one
    /// stream.
    ///
    /// The returned value borrows this reader for its whole life, so the
    /// reader cannot be dropped while the hint exists and no second live
    /// reader can occupy its address. That is what makes the address
    /// comparison inside the hinted readers decisive rather than merely
    /// probable. It is **not** a compile-time bar on offering a hint to
    /// another reader: the hinted readers accept a hint of any lifetime, and a
    /// foreign one is discarded at run time, costing exactly what an unhinted
    /// call costs.
    #[must_use]
    pub const fn chain_hint(&self) -> StreamChainHint<'_> {
        StreamChainHint {
            file: self,
            resume: None,
        }
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
    /// source-version check but performs no payload read. A non-empty read is
    /// bracketed by two successful source-version observations: the most
    /// recent earlier one taken by this reader, which for the first read is
    /// the capture taken while opening it, and one taken after this read
    /// completes. Concurrent callers share one reader, so the earlier
    /// observation is not necessarily one this thread took. If a later range
    /// read or source check fails, callers must discard the destination;
    /// bytes written before the failure are not rolled back.
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
        self.read_stream_range_hinted(path, offset, output, &mut self.chain_hint())
    }

    /// Reads a stream range, resuming the allocation-chain walk from `hint`.
    ///
    /// This is [`Self::read_stream_range`] with one difference: instead of
    /// walking the chain from the stream's first sector to reach `offset`, it
    /// resumes from the position `hint` retains and records where it finished.
    /// A caller that reads one stream forward in several calls therefore walks
    /// each chain link once for the whole scan rather than once per call.
    /// `read_stream_range` is exactly this function with a fresh hint, so there
    /// is one implementation of the read and of every check inside it.
    ///
    /// Every validation keeps its position and its identity. The links the hint
    /// lets the walk skip are links this same hint already walked and validated
    /// on an earlier call through the same chain; links ahead of it are walked
    /// and checked here. A hint naming a different directory entry, a different
    /// first sector, a different allocation table, or a **later** ordinal than
    /// the one requested is discarded, and the walk restarts at the stream's
    /// first sector.
    ///
    /// The hint holds no source bytes and nothing a source mutation can
    /// invalidate: the allocation tables are captured and validated at open and
    /// are immutable thereafter, and this call fences the retained source
    /// exactly as [`Self::read_stream_range`] does.
    ///
    /// # Errors
    ///
    /// Identical to [`Self::read_stream_range`].
    pub fn read_stream_range_hinted(
        &self,
        path: &[&str],
        offset: u64,
        output: &mut [u8],
        hint: &mut StreamChainHint<'_>,
    ) -> Result<(), OleError> {
        let (sid, is_minifat, start_sector, size) = {
            let entry = self.find_entry(path)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()));
            }
            (entry.sid, entry.is_minifat, entry.start_sector, entry.size)
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

        let result = (|| -> Result<(), OleError> {
            if is_minifat {
                self.read_minifat_range_hinted(
                    sid,
                    start_sector,
                    offset,
                    output,
                    false,
                    false,
                    hint,
                )?;
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
                    return Ok(());
                }
                let first_ordinal =
                    usize::try_from(offset / sector_size as u64).map_err(|_error| {
                        OleError::InvalidData("FAT range sector does not fit usize".to_string())
                    })?;
                // Resume the prefix walk where an earlier call through this
                // same hint left it. `walked` is zero and `sector` the
                // stream's first sector whenever the hint does not apply, so
                // the loop below is the original walk in that case. The links
                // it skips are links this hint already traversed and checked.
                let (mut sector, walked) =
                    hint.resume_from(self, sid, is_minifat, start_sector, first_ordinal);
                for _ in walked..first_ordinal {
                    sector = next_chain_sector(&self.index.fat, sector, "FAT")?;
                    if sector == ENDOFCHAIN {
                        return Err(OleError::CorruptedFile(
                            "FAT chain ends before stream range".to_string(),
                        ));
                    }
                }
                hint.record(self, sid, is_minifat, start_sector, first_ordinal, sector);
                let mut within =
                    usize::try_from(offset % sector_size as u64).map_err(|_error| {
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
                            OleError::CorruptedFile(
                                "FAT range physical offset overflow".to_string(),
                            )
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
            Ok(())
        })();
        self.finish_stream_range(result)
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

    fn announce_singleflight_waiter(&self) -> bool {
        let intent = &self.minifat_singleflight.waiter_intent;
        let mut observed = intent.load(AtomicOrdering::Acquire);
        let _ = loop {
            let Some(next) = observed.checked_add(1) else {
                return false;
            };
            match intent.compare_exchange_weak(
                observed,
                next,
                AtomicOrdering::AcqRel,
                AtomicOrdering::Acquire,
            ) {
                Ok(previous) => break previous,
                Err(next_observed) => observed = next_observed,
            }
        };
        // Publish the intent in the same atomic word that an owner CASes.
        // If the state changes while setting the bit, retry against the new
        // generation; the count remains live until the waiter is registered
        // or its guard is dropped.
        loop {
            let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            let desired = state | MINIFAT_INTENT_BIT;
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
                return true;
            }
        }
    }

    fn release_singleflight_intent(&self) {
        let previous = self
            .minifat_singleflight
            .waiter_intent
            .fetch_sub(1, AtomicOrdering::AcqRel);
        debug_assert!(previous > 0, "single-flight waiter intent underflow");
        if previous != 1 {
            return;
        }
        loop {
            let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            if self
                .minifat_singleflight
                .waiter_intent
                .load(AtomicOrdering::Acquire)
                != 0
            {
                // A new announcer may have incremented the lifetime count
                // while this last old waiter was clearing the bit. Reassert
                // the same-word intent before returning so count>0 cannot be
                // left with an unmarked state.
                let desired = state | MINIFAT_INTENT_BIT;
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
                    if self
                        .minifat_singleflight
                        .waiter_intent
                        .load(AtomicOrdering::Acquire)
                        != 0
                    {
                        return;
                    }
                    continue;
                }
                continue;
            }
            if !minifat_state_intent(state) {
                return;
            }
            if self
                .minifat_direct_state
                .compare_exchange(
                    state,
                    state & !MINIFAT_INTENT_BIT,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                if self
                    .minifat_singleflight
                    .waiter_intent
                    .load(AtomicOrdering::Acquire)
                    == 0
                {
                    return;
                }
            }
        }
    }

    fn reassert_singleflight_intent(&self) -> bool {
        if self
            .minifat_singleflight
            .waiter_intent
            .load(AtomicOrdering::Acquire)
            == 0
        {
            return false;
        }
        loop {
            let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            if minifat_state_intent(state) {
                return true;
            }
            if self
                .minifat_direct_state
                .compare_exchange(
                    state,
                    state | MINIFAT_INTENT_BIT,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                return true;
            }
        }
    }

    /// Marks the handoff slot in the same policy word while the slot mutex is
    /// held. This is the owner-release linearization point for a waiter that
    /// may consume and drop before the owner itself unwinds.
    fn mark_singleflight_slot_present(&self, sid: u32, epoch: u32) -> bool {
        loop {
            let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            if minifat_state_kind(state) != MINIFAT_DIRECT_IN_FLIGHT
                || minifat_state_sid(state) != sid
                || minifat_state_epoch(state) != epoch
            {
                return false;
            }
            if minifat_state_slot_present(state) {
                return true;
            }
            if self
                .minifat_direct_state
                .compare_exchange(
                    state,
                    state | MINIFAT_SLOT_PRESENT_BIT,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                return true;
            }
        }
    }

    /// Clears a slot bit while the slot mutex is held. Cache takeover keeps
    /// the bit until this cleanup runs, so stale owner release cannot skip the
    /// mutex merely because waiter intent has already reached zero.
    fn clear_singleflight_slot_present(&self) {
        loop {
            let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            if !minifat_state_slot_present(state) {
                return;
            }
            if self
                .minifat_direct_state
                .compare_exchange(
                    state,
                    state & !MINIFAT_SLOT_PRESENT_BIT,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                return;
            }
        }
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
        let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        let observed_sid = minifat_state_sid(observed);
        let observed_kind = minifat_state_kind(observed);
        match observed_kind {
            MINIFAT_CACHE_REQUESTED
            | MINIFAT_CACHE_IN_FLIGHT
            | MINIFAT_CACHE_READY
            | MINIFAT_CACHE_RETRY => {
                return Ok(MiniFATDirectClaim::Cache);
            },
            MINIFAT_DIRECT_IN_FLIGHT if observed_sid != sid => {
                self.request_ministream_cache();
                return Ok(MiniFATDirectClaim::Cache);
            },
            MINIFAT_DIRECT_DONE if observed_sid != sid => {
                self.request_ministream_cache();
                return Ok(MiniFATDirectClaim::Cache);
            },
            MINIFAT_DIRECT_IN_FLIGHT => {
                if self.minifat_singleflight.slot.is_poisoned() {
                    return Err(minifat_singleflight_poisoned());
                }
                if !self.announce_singleflight_waiter() {
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                }
                return self.claim_minifat_direct_slow(sid, true);
            },
            MINIFAT_DIRECT_UNCLAIMED | MINIFAT_DIRECT_DONE
                if observed_kind == MINIFAT_DIRECT_UNCLAIMED || observed_sid == sid =>
            {
                if self.minifat_singleflight.slot.is_poisoned() {
                    return Err(minifat_singleflight_poisoned());
                }
                let Some(epoch) = next_minifat_epoch(observed) else {
                    // The bounded epoch is deliberately fail-closed. A
                    // wrapped generation could let a delayed owner mutate
                    // a later same-SID flight, so select the permanent
                    // root-cache policy at exhaustion.
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                };
                if self
                    .minifat_singleflight
                    .waiter_intent
                    .load(AtomicOrdering::Acquire)
                    == 0
                    && !minifat_state_intent(observed)
                    && !minifat_state_slot_present(observed)
                {
                    let desired =
                        minifat_state_with_meta(sid, MINIFAT_DIRECT_IN_FLIGHT, epoch, false);
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
                        // No slot is needed for the uncontended owner. A
                        // waiter that announces after this CAS can create
                        // the slot under the slow path before publication.
                        return Ok(MiniFATDirectClaim::Owner(MiniFATDirectOwner {
                            file: self,
                            sid,
                            epoch,
                            published: false,
                            success: false,
                        }));
                    }
                }
                return self.claim_minifat_direct_slow(sid, false);
            },
            _ => {
                self.request_ministream_cache();
                return Ok(MiniFATDirectClaim::Cache);
            },
        }
    }

    fn claim_minifat_direct_slow(
        &self,
        sid: u32,
        mut announced: bool,
    ) -> Result<MiniFATDirectClaim<'_>, OleError> {
        #[cfg(test)]
        self.minifat_singleflight
            .claim_slow_entered
            .store(1, AtomicOrdering::SeqCst);
        let mut slot = match self.minifat_singleflight.slot.lock() {
            Ok(slot) => slot,
            Err(_error) => {
                if announced {
                    self.release_singleflight_intent();
                }
                return Err(minifat_singleflight_poisoned());
            },
        };
        loop {
            let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
            let observed_sid = minifat_state_sid(observed);
            let observed_kind = minifat_state_kind(observed);
            match observed_kind {
                MINIFAT_CACHE_REQUESTED
                | MINIFAT_CACHE_IN_FLIGHT
                | MINIFAT_CACHE_READY
                | MINIFAT_CACHE_RETRY => {
                    drop(slot);
                    if announced {
                        self.release_singleflight_intent();
                    }
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_IN_FLIGHT if observed_sid != sid => {
                    drop(slot);
                    if announced {
                        self.release_singleflight_intent();
                    }
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_DONE if observed_sid != sid => {
                    drop(slot);
                    if announced {
                        self.release_singleflight_intent();
                    }
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_IN_FLIGHT => {
                    let epoch = minifat_state_epoch(observed);
                    if let Some(current) = slot.as_ref()
                        && (current.sid != sid || current.epoch != epoch)
                    {
                        drop(slot);
                        if announced {
                            self.release_singleflight_intent();
                        }
                        self.request_ministream_cache();
                        return Ok(MiniFATDirectClaim::Cache);
                    }
                    if slot.is_none() {
                        if !announced {
                            if !self.announce_singleflight_waiter() {
                                drop(slot);
                                self.request_ministream_cache();
                                return Ok(MiniFATDirectClaim::Cache);
                            }
                            announced = true;
                        }
                        if !self.reassert_singleflight_intent() {
                            drop(slot);
                            return Ok(MiniFATDirectClaim::Cache);
                        }
                        // A cache takeover may race with the intent CAS while
                        // this caller still owns the slot mutex. Re-evaluate
                        // before creating a handoff marker for a cache flight.
                        let current_state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
                        if minifat_state_kind(current_state) != MINIFAT_DIRECT_IN_FLIGHT
                            || minifat_state_sid(current_state) != sid
                            || minifat_state_epoch(current_state) != epoch
                        {
                            continue;
                        }
                        if !self.mark_singleflight_slot_present(sid, epoch) {
                            continue;
                        }
                        let current_state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
                        if minifat_state_kind(current_state) != MINIFAT_DIRECT_IN_FLIGHT
                            || minifat_state_sid(current_state) != sid
                            || minifat_state_epoch(current_state) != epoch
                        {
                            self.clear_singleflight_slot_present();
                            continue;
                        }
                        *slot = Some(MiniFATSingleFlightSlot {
                            sid,
                            epoch,
                            owner_active: true,
                            waiters: 0,
                            status: MiniFATSingleFlightStatus::InFlight,
                        });
                    }
                    if let Some(current) = slot.as_mut()
                        && current.sid == sid
                        && current.epoch == epoch
                    {
                        if !self.mark_singleflight_slot_present(sid, epoch) {
                            drop(slot);
                            if announced {
                                self.release_singleflight_intent();
                            }
                            return Ok(MiniFATDirectClaim::Cache);
                        }
                        if !self.reassert_singleflight_intent() {
                            drop(slot);
                            return Ok(MiniFATDirectClaim::Cache);
                        }
                        let current_state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
                        if minifat_state_kind(current_state) != MINIFAT_DIRECT_IN_FLIGHT
                            || minifat_state_sid(current_state) != sid
                            || minifat_state_epoch(current_state) != epoch
                        {
                            drop(slot);
                            if announced {
                                self.release_singleflight_intent();
                            }
                            return Ok(MiniFATDirectClaim::Cache);
                        }
                        if let Some(next) = current.waiters.checked_add(1) {
                            current.waiters = next;
                            drop(slot);
                            return Ok(MiniFATDirectClaim::Waiter(MiniFATDirectWaiter {
                                file: self,
                                sid,
                                epoch,
                                registered: true,
                                intent: true,
                            }));
                        }
                        drop(slot);
                        if announced {
                            self.release_singleflight_intent();
                        }
                        self.request_ministream_cache();
                        return Ok(MiniFATDirectClaim::Cache);
                    }
                    drop(slot);
                    if announced {
                        self.release_singleflight_intent();
                    }
                    self.request_ministream_cache();
                    return Ok(MiniFATDirectClaim::Cache);
                },
                MINIFAT_DIRECT_UNCLAIMED | MINIFAT_DIRECT_DONE => {
                    if let Some(current) = slot.as_ref()
                        && current.sid != sid
                        && (current.owner_active || current.waiters > 0)
                    {
                        drop(slot);
                        if announced {
                            self.release_singleflight_intent();
                        }
                        self.request_ministream_cache();
                        return Ok(MiniFATDirectClaim::Cache);
                    }
                    if let Some(current) = slot.as_mut()
                        && current.sid != sid
                    {
                        slot.take();
                        self.clear_singleflight_slot_present();
                    }
                    if let Some(current) = slot.as_mut()
                        && current.sid == sid
                    {
                        if !current.owner_active
                            && matches!(
                                &current.status,
                                MiniFATSingleFlightStatus::Failed
                                    | MiniFATSingleFlightStatus::CompletedNoHandoff
                            )
                        {
                            let Some(epoch) = next_minifat_epoch(observed) else {
                                drop(slot);
                                if announced {
                                    self.release_singleflight_intent();
                                }
                                self.request_ministream_cache();
                                return Ok(MiniFATDirectClaim::Cache);
                            };
                            if self
                                .minifat_direct_state
                                .compare_exchange(
                                    observed,
                                    minifat_state_with_slot(
                                        sid,
                                        MINIFAT_DIRECT_IN_FLIGHT,
                                        epoch,
                                        minifat_state_intent(observed),
                                        true,
                                    ),
                                    AtomicOrdering::AcqRel,
                                    AtomicOrdering::Acquire,
                                )
                                .is_ok()
                            {
                                if announced {
                                    self.release_singleflight_intent();
                                }
                                current.epoch = epoch;
                                // Existing waiters belong to the failed
                                // epoch and will drop/re-register after
                                // observing the terminal marker. Do not let a
                                // delayed old guard decrement the new epoch's
                                // handoff count.
                                current.waiters = 0;
                                current.owner_active = true;
                                current.status = MiniFATSingleFlightStatus::InFlight;
                                drop(slot);
                                return Ok(MiniFATDirectClaim::Owner(MiniFATDirectOwner {
                                    file: self,
                                    sid,
                                    epoch,
                                    published: false,
                                    success: false,
                                }));
                            }
                            continue;
                        }
                        if current.owner_active || current.waiters > 0 {
                            if !announced {
                                if !self.announce_singleflight_waiter() {
                                    drop(slot);
                                    self.request_ministream_cache();
                                    return Ok(MiniFATDirectClaim::Cache);
                                }
                                announced = true;
                            }
                            if !self.reassert_singleflight_intent() {
                                drop(slot);
                                return Ok(MiniFATDirectClaim::Cache);
                            }
                            let current_state =
                                self.minifat_direct_state.load(AtomicOrdering::Acquire);
                            if minifat_state_kind(current_state) != MINIFAT_DIRECT_DONE
                                && minifat_state_kind(current_state) != MINIFAT_DIRECT_UNCLAIMED
                            {
                                drop(slot);
                                if announced {
                                    self.release_singleflight_intent();
                                }
                                return Ok(MiniFATDirectClaim::Cache);
                            }
                            if let Some(next) = current.waiters.checked_add(1) {
                                let current_epoch = current.epoch;
                                current.waiters = next;
                                drop(slot);
                                return Ok(MiniFATDirectClaim::Waiter(MiniFATDirectWaiter {
                                    file: self,
                                    sid,
                                    epoch: current_epoch,
                                    registered: true,
                                    intent: true,
                                }));
                            }
                            drop(slot);
                            if announced {
                                self.release_singleflight_intent();
                            }
                            self.request_ministream_cache();
                            return Ok(MiniFATDirectClaim::Cache);
                        }
                    }

                    if let Some(current) = slot.as_mut()
                        && current.sid == sid
                        && current.waiters == 0
                        && !current.owner_active
                    {
                        slot.take();
                        self.clear_singleflight_slot_present();
                    }
                    let Some(epoch) = next_minifat_epoch(observed) else {
                        drop(slot);
                        if announced {
                            self.release_singleflight_intent();
                        }
                        self.request_ministream_cache();
                        return Ok(MiniFATDirectClaim::Cache);
                    };
                    let intent = minifat_state_intent(observed)
                        || self
                            .minifat_singleflight
                            .waiter_intent
                            .load(AtomicOrdering::Acquire)
                            != 0;
                    if self
                        .minifat_direct_state
                        .compare_exchange(
                            observed,
                            minifat_state_with_meta(sid, MINIFAT_DIRECT_IN_FLIGHT, epoch, intent),
                            AtomicOrdering::AcqRel,
                            AtomicOrdering::Acquire,
                        )
                        .is_ok()
                    {
                        drop(slot);
                        if announced {
                            self.release_singleflight_intent();
                        }
                        return Ok(MiniFATDirectClaim::Owner(MiniFATDirectOwner {
                            file: self,
                            sid,
                            epoch,
                            published: false,
                            success: false,
                        }));
                    }
                },
                _ => {
                    drop(slot);
                    if announced {
                        self.release_singleflight_intent();
                    }
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
        epoch: u32,
        payload: &[u8],
        success: bool,
    ) -> Result<(), OleError> {
        let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        if minifat_state_kind(observed) != MINIFAT_DIRECT_IN_FLIGHT
            || minifat_state_sid(observed) != sid
            || minifat_state_epoch(observed) != epoch
        {
            return Ok(());
        }
        // A sequential direct owner deliberately has no slot and performs no
        // allocation here. A waiter may have announced intent but not yet
        // registered; that caller will retry if it loses the handoff race.
        if self
            .minifat_singleflight
            .waiter_intent
            .load(AtomicOrdering::Acquire)
            == 0
            && !minifat_state_intent(observed)
        {
            return Ok(());
        }
        if self.minifat_singleflight.slot.is_poisoned() {
            return Err(minifat_singleflight_poisoned());
        }
        let mut slot = self
            .minifat_singleflight
            .slot
            .lock()
            .map_err(|_error| minifat_singleflight_poisoned())?;
        let current_state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        if minifat_state_kind(current_state) != MINIFAT_DIRECT_IN_FLIGHT
            || minifat_state_sid(current_state) != sid
            || minifat_state_epoch(current_state) != epoch
        {
            return Ok(());
        }
        if let Some(current) = slot.as_mut()
            && current.sid == sid
            && current.epoch == epoch
            && current.waiters > 0
        {
            current.status = if success {
                match clone_minifat_waiter_payload(payload) {
                    Ok(payload) => MiniFATSingleFlightStatus::Succeeded(payload),
                    Err(_error) => MiniFATSingleFlightStatus::CompletedNoHandoff,
                }
            } else {
                MiniFATSingleFlightStatus::Failed
            };
            self.minifat_singleflight.wake.notify_all();
        }
        Ok(())
    }

    fn release_minifat_direct(&self, sid: u32, epoch: u32, published: bool, success: bool) {
        let state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        if minifat_state_kind(state) == MINIFAT_DIRECT_IN_FLIGHT
            && minifat_state_sid(state) == sid
            && minifat_state_epoch(state) == epoch
            && self
                .minifat_singleflight
                .waiter_intent
                .load(AtomicOrdering::Acquire)
                == 0
            && !minifat_state_intent(state)
            && !minifat_state_slot_present(state)
            && !self.minifat_singleflight.slot.is_poisoned()
        {
            let target_kind = if published && success {
                MINIFAT_DIRECT_DONE
            } else {
                MINIFAT_DIRECT_UNCLAIMED
            };
            let target_sid = if published && success { sid } else { 0 };
            let target = minifat_state_with_meta(target_sid, target_kind, epoch, false);
            if self
                .minifat_direct_state
                .compare_exchange(
                    state,
                    target,
                    AtomicOrdering::AcqRel,
                    AtomicOrdering::Acquire,
                )
                .is_ok()
            {
                return;
            }
        }
        // This is an unwind-safe RAII path. If another thread poisoned the
        // private marker while unwinding, recover its guard so waiters can be
        // notified and the atomic cache policy can still settle; user-facing
        // claim/wait/publish paths return a typed poison error instead.
        let mut slot = self
            .minifat_singleflight
            .slot
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let observed = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        let state_matches = minifat_state_kind(observed) == MINIFAT_DIRECT_IN_FLIGHT
            && minifat_state_sid(observed) == sid
            && minifat_state_epoch(observed) == epoch;
        let cache_state = matches!(
            minifat_state_kind(observed),
            MINIFAT_CACHE_REQUESTED
                | MINIFAT_CACHE_IN_FLIGHT
                | MINIFAT_CACHE_READY
                | MINIFAT_CACHE_RETRY
        );
        let slot_matches = slot
            .as_ref()
            .is_some_and(|current| current.sid == sid && current.epoch == epoch);

        if slot_matches && (state_matches || cache_state) {
            if !published || !success {
                if let Some(current) = slot.as_mut() {
                    current.status = MiniFATSingleFlightStatus::Failed;
                }
            } else if let Some(current) = slot.as_mut()
                && matches!(current.status, MiniFATSingleFlightStatus::InFlight)
            {
                // A waiter registered after publication could not have
                // received a payload handoff. Wake it to retry instead of
                // leaving the slot permanently InFlight.
                current.status = MiniFATSingleFlightStatus::CompletedNoHandoff;
            }
            if let Some(current) = slot.as_mut() {
                current.owner_active = false;
            }
        }

        // Keep the slot mutex and SLOT_PRESENT bit held through the terminal
        // state CAS. A claimant cannot observe IN_FLIGHT + an empty slot and
        // then register into a flight whose owner has already published its
        // terminal state.
        #[cfg(test)]
        if self
            .minifat_singleflight
            .release_pause
            .load(AtomicOrdering::SeqCst)
            != 0
        {
            self.minifat_singleflight
                .release_pause_ready
                .store(1, AtomicOrdering::SeqCst);
            while self
                .minifat_singleflight
                .release_pause_continue
                .load(AtomicOrdering::SeqCst)
                == 0
            {
                std::thread::yield_now();
            }
        }
        let mut state_terminal = false;
        if state_matches {
            let target_kind = if published && success {
                MINIFAT_DIRECT_DONE
            } else {
                MINIFAT_DIRECT_UNCLAIMED
            };
            let target_sid = if published && success { sid } else { 0 };
            // A late waiter announces intent without taking the slot mutex.
            // Reload after every failed CAS so that this owner cannot unlock
            // with an ownerless IN_FLIGHT marker when that bit changes after
            // the first state load. Cache takeover is allowed to win; its
            // state is handled by the cleanup below.
            loop {
                let current_state = self.minifat_direct_state.load(AtomicOrdering::Acquire);
                if minifat_state_kind(current_state) != MINIFAT_DIRECT_IN_FLIGHT
                    || minifat_state_sid(current_state) != sid
                    || minifat_state_epoch(current_state) != epoch
                {
                    break;
                }
                let intent = minifat_state_intent(current_state)
                    || self
                        .minifat_singleflight
                        .waiter_intent
                        .load(AtomicOrdering::Acquire)
                        != 0;
                let slot_present = minifat_state_slot_present(current_state) || slot.is_some();
                let target =
                    minifat_state_with_slot(target_sid, target_kind, epoch, intent, slot_present);
                if self
                    .minifat_direct_state
                    .compare_exchange(
                        current_state,
                        target,
                        AtomicOrdering::AcqRel,
                        AtomicOrdering::Acquire,
                    )
                    .is_ok()
                {
                    state_terminal = true;
                    break;
                }
            }
        }

        let state_after = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        let cache_after = matches!(
            minifat_state_kind(state_after),
            MINIFAT_CACHE_REQUESTED
                | MINIFAT_CACHE_IN_FLIGHT
                | MINIFAT_CACHE_READY
                | MINIFAT_CACHE_RETRY
        );
        if slot_matches {
            let mut remove_slot = false;
            if let Some(current) = slot.as_ref() {
                remove_slot = current.waiters == 0
                    && (state_terminal || cache_after)
                    && !current.owner_active;
            }
            if remove_slot {
                slot.take();
                self.clear_singleflight_slot_present();
            }
            self.minifat_singleflight.wake.notify_all();
        } else if slot.is_none() && (state_terminal || cache_after) {
            self.clear_singleflight_slot_present();
        }
        drop(slot);
    }

    fn release_minifat_waiter(&self, sid: u32, epoch: u32, registered: bool) {
        // Waiter Drop must remain cleanup-only even after a peer panic. The
        // explicit claim/wait operations report poison as OleError; this
        // recovery is solely to prevent a leaked waiter from stranding the
        // bounded flight marker.
        if registered {
            let mut slot = self
                .minifat_singleflight
                .slot
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            if let Some(current) = slot.as_mut()
                && current.sid == sid
                && current.epoch == epoch
            {
                current.waiters = current.waiters.saturating_sub(1);
                if current.waiters == 0 && !current.owner_active {
                    slot.take();
                    self.clear_singleflight_slot_present();
                }
            }
            if slot.is_none() {
                self.clear_singleflight_slot_present();
            }
        }
        self.release_singleflight_intent();
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
            let desired = minifat_state_with_slot(
                MINIFAT_CACHE_SID,
                MINIFAT_CACHE_REQUESTED,
                minifat_state_epoch(observed),
                minifat_state_intent(observed),
                minifat_state_slot_present(observed),
            );
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
            let desired = minifat_state_with_slot(
                MINIFAT_CACHE_SID,
                MINIFAT_CACHE_IN_FLIGHT,
                minifat_state_epoch(state),
                minifat_state_intent(state),
                minifat_state_slot_present(state),
            );
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
        let in_flight = self.minifat_direct_state.load(AtomicOrdering::Acquire);
        if minifat_state_kind(in_flight) != MINIFAT_CACHE_IN_FLIGHT {
            return;
        }
        let target_kind = if success {
            MINIFAT_CACHE_READY
        } else {
            // Once a caller has selected the root cache, a failed cache load
            // must remain retryable through that cache. Returning to the
            // direct state could reintroduce a second bounded read after a
            // concurrent/different target already observed this path.
            MINIFAT_CACHE_RETRY
        };
        let target = minifat_state_with_slot(
            MINIFAT_CACHE_SID,
            target_kind,
            minifat_state_epoch(in_flight),
            minifat_state_intent(in_flight),
            minifat_state_slot_present(in_flight),
        );
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
        // `MINIFAT_UNHINTED_SID` cannot equal any directory entry's SID, so a
        // hint recorded here can never be reused; this call therefore behaves
        // exactly as it did before hints existed.
        self.read_minifat_range_hinted(
            MINIFAT_UNHINTED_SID,
            start_sector,
            offset,
            output,
            require_chain_end,
            zero_fill_truncated,
            &mut self.chain_hint(),
        )
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "one reader with an explicit resumable chain position"
    )]
    fn read_minifat_range_hinted(
        &self,
        sid: u32,
        start_sector: u32,
        offset: u64,
        output: &mut [u8],
        require_chain_end: bool,
        zero_fill_truncated: bool,
        hint: &mut StreamChainHint<'_>,
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
        let (mut mini_sector, walked) =
            hint.resume_from(self, sid, true, start_sector, first_ordinal);
        for _ in walked..first_ordinal {
            mini_sector = next_chain_sector(&self.index.minifat, mini_sector, "MiniFAT")?;
            if mini_sector == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "MiniFAT chain ends before stream range".to_string(),
                ));
            }
        }
        hint.record(self, sid, true, start_sector, first_ordinal, mini_sector);
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

    /// Observes `source` and refuses when it is no longer `expected`.
    ///
    /// This is [`Self::check_source_version`] for the open path, where no
    /// `SharedOleFile` exists yet to hold the captured version. It is called
    /// only on branches that leave `open_source_with_limits` before the
    /// post-parse comparison, so that a failure derived from the source's
    /// length is never reported in place of the mutation that caused it.
    fn refuse_if_changed(source: &dyn ReadAt, expected: SourceVersion) -> Result<(), OleError> {
        let observed = source.version()?;
        if observed == expected {
            Ok(())
        } else {
            Err(OleError::SourceChanged { expected, observed })
        }
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

    fn finish_stream_range(&self, result: Result<(), OleError>) -> Result<(), OleError> {
        match result {
            Ok(()) => self.check_source_version(),
            Err(original) => match self.check_source_version() {
                Err(error @ OleError::SourceChanged { .. }) => Err(error),
                _ => Err(original),
            },
        }
    }

    fn finish_cursor_operation<T>(&self, result: Result<T, OleError>) -> Result<T, OleError> {
        match result {
            Ok(value) => self.check_source_version().map(|()| value),
            Err(original) => match self.check_source_version() {
                Err(error @ OleError::SourceChanged { .. }) => Err(error),
                _ => Err(original),
            },
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

struct OwnedArcSource {
    source: Arc<[u8]>,
    version: SourceVersion,
    #[cfg(test)]
    read_count: Option<Arc<AtomicUsize>>,
}

impl ReadAt for OwnedArcSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.source.len())
            .map_err(|_error| io::Error::other("owned CFB source length exceeds u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        #[cfg(test)]
        if let Some(read_count) = &self.read_count {
            read_count.fetch_add(1, AtomicOrdering::SeqCst);
        }
        if output.is_empty() {
            return Ok(0);
        }
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(available) = self.source.get(start..) else {
            return Ok(0);
        };
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

impl SharedOleStreamCursor<'_> {
    /// Declared logical stream length.
    #[must_use]
    pub const fn len(&self) -> u64 {
        self.length
    }

    /// Whether the selected stream is logically empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.length == 0
    }

    /// Current logical stream position.
    #[must_use]
    pub const fn position(&self) -> u64 {
        self.position
    }

    /// Advances the cursor by `bytes` without reading payload bytes.
    ///
    /// Movement is strictly forward. A failed movement leaves the logical
    /// position unchanged. Skipping traverses only the immutable, validated
    /// allocation tables and publishes no source bytes, so it takes no
    /// source-version observation of its own. The next read uses the moved
    /// position and then fences the retained source before publishing the
    /// bytes it read.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-data error when the movement exceeds the
    /// stream, or a corruption error when the chain cannot be traversed
    /// consistently.
    pub fn skip_forward(&mut self, bytes: u64) -> Result<(), OleError> {
        self.move_forward(bytes)
    }

    /// Moves to a later logical position without reading payload bytes.
    ///
    /// Passing a position before the current position is rejected; callers
    /// must create a new cursor for backward movement. Skipping publishes no
    /// source bytes and does not perform a source-version fence; the next
    /// read detects a persistent source change.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-data error for backward or out-of-bounds
    /// movement, or a corruption error when the chain cannot be traversed
    /// consistently.
    pub fn skip_to(&mut self, position: u64) -> Result<(), OleError> {
        if position < self.position {
            return Err(OleError::InvalidData(
                "stream cursor cannot move backwards".to_string(),
            ));
        }
        let bytes = position
            .checked_sub(self.position)
            .ok_or_else(|| OleError::InvalidData("stream cursor position underflow".to_string()))?;
        self.move_forward(bytes)
    }

    /// Reads exactly `output.len()` logical bytes and advances the cursor.
    ///
    /// Every read fences the retained source after payload I/O. Together with
    /// the most recent earlier observation taken by the shared reader, which
    /// for the first read is the capture taken while opening it, that later
    /// fence brackets the bytes this call publishes: both observations must
    /// equal the opened source version. A concurrent caller may have taken
    /// that earlier observation. Because no fence runs before the payload
    /// read, a source mutation that is reverted before this read's own fence
    /// is not observed, exactly as
    /// `litchi_core::FileVersionPolicy` already documents for reverted
    /// transitions. Reads are grouped into physically contiguous FAT or
    /// MiniFAT runs when doing so is safe. The destination may contain a
    /// prefix when a later source, chain, or version check fails; callers must
    /// discard it on any error. A failed read does not commit the cursor
    /// position, so a stable source may be retried. If a read/chain error and a
    /// source mutation race, the post-read `SourceChanged` error wins. If the
    /// fence itself cannot observe the source while the payload read also
    /// failed, the payload error is reported; a fence failure after a
    /// successful read is reported instead of the bytes.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-data error when the requested bytes exceed the
    /// remaining logical stream, a corruption error when allocation traversal
    /// fails, or a source I/O/source-version error.
    pub fn read_exact(&mut self, output: &mut [u8]) -> Result<(), OleError> {
        let state = self.state;
        let position = self.position;
        let result = (|| {
            let requested = output.len() as u64;
            let end = position.checked_add(requested).ok_or_else(|| {
                OleError::InvalidData("stream cursor read end overflow".to_string())
            })?;
            if end > self.length {
                return Err(OleError::InvalidData(format!(
                    "stream cursor read {position}..{end} exceeds length {}",
                    self.length
                )));
            }
            let state = self.read_from_state(state, output)?;
            let state = if end == self.length {
                SharedOleStreamCursorState::End
            } else {
                state
            };
            Ok((state, end))
        })();
        match self.file.finish_cursor_operation(result) {
            Ok((state, position)) => {
                self.state = state;
                self.position = position;
                Ok(())
            },
            Err(error) => Err(error),
        }
    }

    fn move_forward(&mut self, bytes: u64) -> Result<(), OleError> {
        let state = self.state;
        let position = self.position;
        let end = position
            .checked_add(bytes)
            .ok_or_else(|| OleError::InvalidData("stream cursor position overflow".to_string()))?;
        if end > self.length {
            return Err(OleError::InvalidData(format!(
                "stream cursor position {end} exceeds length {}",
                self.length
            )));
        }
        let state = self.advance_state(state, bytes)?;
        self.state = if end == self.length {
            SharedOleStreamCursorState::End
        } else {
            state
        };
        self.position = end;
        Ok(())
    }

    fn read_from_state(
        &self,
        mut state: SharedOleStreamCursorState,
        output: &mut [u8],
    ) -> Result<SharedOleStreamCursorState, OleError> {
        let mut written = 0_usize;
        while written < output.len() {
            state = self.normalize_state(state)?;
            let (physical, available) = self.physical_span(state)?;
            let mut run_bytes = available.min(output.len() - written);
            let mut probe = self.advance_within(state, run_bytes)?;

            while written + run_bytes < output.len() {
                let next = self.normalize_state(probe)?;
                let (next_physical, next_available) = self.physical_span(next)?;
                let contiguous = physical.checked_add(run_bytes as u64).ok_or_else(|| {
                    OleError::CorruptedFile(
                        "stream cursor physical run offset overflow".to_string(),
                    )
                })?;
                if next_physical != contiguous {
                    probe = next;
                    break;
                }
                let count = next_available.min(output.len() - written - run_bytes);
                run_bytes = run_bytes.checked_add(count).ok_or_else(|| {
                    OleError::CorruptedFile("stream cursor run size overflow".to_string())
                })?;
                probe = self.advance_within(next, count)?;
            }

            self.file
                .source
                .read_exact_at(physical, &mut output[written..written + run_bytes])?;
            written += run_bytes;
            state = probe;
        }
        Ok(state)
    }

    fn advance_state(
        &self,
        mut state: SharedOleStreamCursorState,
        mut bytes: u64,
    ) -> Result<SharedOleStreamCursorState, OleError> {
        while bytes > 0 {
            state = self.normalize_state(state)?;
            let available = self.physical_span(state)?.1;
            let count = bytes.min(available as u64);
            state = self.advance_within(
                state,
                usize::try_from(count).map_err(|_error| {
                    OleError::InvalidData("stream cursor movement does not fit usize".to_string())
                })?,
            )?;
            bytes -= count;
        }
        Ok(state)
    }

    fn advance_within(
        &self,
        state: SharedOleStreamCursorState,
        bytes: usize,
    ) -> Result<SharedOleStreamCursorState, OleError> {
        match state {
            SharedOleStreamCursorState::End => Err(OleError::CorruptedFile(
                "stream cursor passed EOF".to_string(),
            )),
            SharedOleStreamCursorState::Fat { sector, within } => {
                let unit = self.file.index.sector_size;
                let next = within.checked_add(bytes).ok_or_else(|| {
                    OleError::CorruptedFile("FAT cursor offset overflow".to_string())
                })?;
                if next > unit {
                    return Err(OleError::CorruptedFile(
                        "FAT cursor crossed a sector boundary".to_string(),
                    ));
                }
                Ok(SharedOleStreamCursorState::Fat {
                    sector,
                    within: next,
                })
            },
            SharedOleStreamCursorState::MiniFAT { sector, within } => {
                let unit = self.file.index.mini_sector_size;
                let next = within.checked_add(bytes).ok_or_else(|| {
                    OleError::CorruptedFile("MiniFAT cursor offset overflow".to_string())
                })?;
                if next > unit {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT cursor crossed a sector boundary".to_string(),
                    ));
                }
                Ok(SharedOleStreamCursorState::MiniFAT {
                    sector,
                    within: next,
                })
            },
        }
    }

    fn normalize_state(
        &self,
        state: SharedOleStreamCursorState,
    ) -> Result<SharedOleStreamCursorState, OleError> {
        match state {
            SharedOleStreamCursorState::End => Ok(state),
            SharedOleStreamCursorState::Fat { sector, within } => {
                let unit = self.file.index.sector_size;
                if within < unit {
                    return Ok(state);
                }
                if within > unit {
                    return Err(OleError::CorruptedFile(
                        "FAT cursor offset exceeds sector size".to_string(),
                    ));
                }
                let next = next_chain_sector(&self.file.index.fat, sector, "FAT")?;
                if next == ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(
                        "FAT chain ends before cursor range".to_string(),
                    ));
                }
                Ok(SharedOleStreamCursorState::Fat {
                    sector: next,
                    within: 0,
                })
            },
            SharedOleStreamCursorState::MiniFAT { sector, within } => {
                let unit = self.file.index.mini_sector_size;
                if within < unit {
                    return Ok(state);
                }
                if within > unit {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT cursor offset exceeds sector size".to_string(),
                    ));
                }
                let next = next_chain_sector(&self.file.index.minifat, sector, "MiniFAT")?;
                if next == ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT chain ends before cursor range".to_string(),
                    ));
                }
                Ok(SharedOleStreamCursorState::MiniFAT {
                    sector: next,
                    within: 0,
                })
            },
        }
    }

    fn physical_span(&self, state: SharedOleStreamCursorState) -> Result<(u64, usize), OleError> {
        match state {
            SharedOleStreamCursorState::End => Err(OleError::CorruptedFile(
                "stream cursor passed EOF".to_string(),
            )),
            SharedOleStreamCursorState::Fat { sector, within } => {
                let unit = self.file.index.sector_size;
                if within >= unit {
                    return Err(OleError::CorruptedFile(
                        "FAT cursor offset exceeds sector size".to_string(),
                    ));
                }
                let physical = (u64::from(sector) + 1)
                    .checked_mul(unit as u64)
                    .and_then(|value| value.checked_add(within as u64))
                    .ok_or_else(|| {
                        OleError::CorruptedFile("FAT cursor physical offset overflow".to_string())
                    })?;
                Ok((physical, unit - within))
            },
            SharedOleStreamCursorState::MiniFAT { sector, within } => {
                let root = self.file.index.root.as_ref().ok_or_else(|| {
                    OleError::CorruptedFile("mini stream has no root entry".to_string())
                })?;
                let unit = self.file.index.mini_sector_size;
                if within >= unit {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT cursor offset exceeds sector size".to_string(),
                    ));
                }
                let mini_offset = usize::try_from(sector)
                    .map_err(|_error| {
                        OleError::CorruptedFile(
                            "MiniFAT cursor sector does not fit usize".to_string(),
                        )
                    })?
                    .checked_mul(unit)
                    .and_then(|value| value.checked_add(within))
                    .ok_or_else(|| {
                        OleError::CorruptedFile("MiniFAT cursor offset overflow".to_string())
                    })?;
                let root_remaining =
                    root.size.checked_sub(mini_offset as u64).ok_or_else(|| {
                        OleError::CorruptedFile(
                            "MiniFAT cursor exceeds root mini-stream length".to_string(),
                        )
                    })?;
                let count =
                    (unit - within).min(usize::try_from(root_remaining).unwrap_or(usize::MAX));
                if count == 0 {
                    return Err(OleError::CorruptedFile(
                        "MiniFAT cursor exceeds root mini-stream length".to_string(),
                    ));
                }
                let range = self
                    .file
                    .minifat_physical_range(root.size, sector, within, count, 0)?;
                Ok((range.physical, range.length))
            },
        }
    }
}

/// A bounded forward read-ahead over one [`SharedOleStreamCursor`].
///
/// A plain cursor issues one positional request, and takes one source-version
/// observation, per [`SharedOleStreamCursor::read_exact`] call. A caller that
/// frames a record stream as a four-byte header followed by a payload
/// therefore pays two requests and two observations per record, whatever the
/// records weigh: on a caller-supplied range source that is two round trips
/// per record. This wrapper reads a bounded span of the stream once -- through
/// the cursor's own validated chain walk, which coalesces physically
/// contiguous sectors, so a fill costs one request per contiguous run -- and
/// serves the reads it covers out of memory.
///
/// It is a wrapper rather than a mode of the cursor on purpose. Every byte it
/// moves goes through the public cursor API, so a caller that does not build
/// one cannot pay for it: the cursor keeps its size, its layout, its drop
/// glue and its read path exactly as they were. The XLS worksheet scan and the
/// shared-string resolver build cursors in their inner loops, and measurement
/// showed a field added to the cursor itself costing them about 6% of a
/// whole-sheet walk.
///
/// # What a caller trades
///
/// **Bytes.** A fill may read bytes the caller never consumes -- at most what
/// it has already consumed, because fills start at
/// [`Self::FIRST_FILL_BYTES`] and double to the declared ceiling, and never
/// more than that ceiling. A caller that walks a stream to its end reads
/// exactly the bytes it would have read; a caller that stops early reads at
/// most one growth step further. Fills never cross the declared stream length,
/// so no byte outside the selected stream is ever read.
///
/// **When a mutation is reported.** The source observation that follows every
/// cursor read still happens, once per fill. A read the window already covers
/// publishes bytes that fill bracketed, and takes no observation of its own,
/// so a mutation between two served reads is reported by the next fill rather
/// than by the read that follows it. The operation-level bracket is unchanged:
/// a caller that closes with its own observation -- as
/// `SharedOleFile::source_version` does -- still refuses a source that moved
/// at any point. A caller that needs a per-read mutation fence must leave the
/// ceiling at zero, which makes every read a plain cursor read again.
///
/// # Memory
///
/// One window of at most the declared ceiling, allocated fallibly and grown
/// only to what a fill needs.
pub struct BufferedOleStreamCursor<'a> {
    cursor: SharedOleStreamCursor<'a>,
    /// Largest number of stream bytes one fill may leave resident. Zero
    /// disables read-ahead, making every read a plain cursor read.
    ceiling: usize,
    /// Size of the next fill before it is clamped by the ceiling, by the bytes
    /// the current read needs and by the remaining stream.
    target: usize,
    /// Stream bytes `[position, cursor.position())`.
    window: Vec<u8>,
    /// Bytes of `window` the caller has already been served.
    consumed: usize,
}

impl std::fmt::Debug for BufferedOleStreamCursor<'_> {
    /// Reports the shape of the window, never the stream bytes in it, which is
    /// what [`SharedOleStreamCursor`]'s own `Debug` does for the same reason.
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("BufferedOleStreamCursor")
            .field("position", &self.position())
            .field("length", &self.cursor.len())
            .field("ceiling", &self.ceiling)
            .field("buffered", &self.buffered())
            .finish_non_exhaustive()
    }
}

impl<'a> BufferedOleStreamCursor<'a> {
    /// Largest window one wrapper may be asked to hold.
    ///
    /// A caller may declare less. This is the 64 KiB the XLS worksheet window
    /// took in change 0568, and it bounds the memory one wrapper retains as
    /// well as the bytes one fill may read beyond the current request.
    pub const MAX_READ_AHEAD_BYTES: usize = 64 * 1024;

    /// Size of the first fill that reads beyond what the current read needs,
    /// which then doubles to the declared ceiling.
    ///
    /// This is change 0568's worksheet-window schedule, for the same reason:
    /// it bounds what a caller that stops after a few records reads beyond
    /// them by what it has already consumed, instead of paying one
    /// ceiling-sized read for a stream it abandons.
    pub const FIRST_FILL_BYTES: usize = 512;

    /// Wraps `cursor` with a window of at most `ceiling` bytes.
    ///
    /// `ceiling` is clamped to [`Self::MAX_READ_AHEAD_BYTES`]. Zero produces a
    /// wrapper that reads exactly as the bare cursor does, request for request
    /// and observation for observation, which is what a caller wants until it
    /// has established that reading ahead is safe.
    #[must_use]
    pub fn new(cursor: SharedOleStreamCursor<'a>, ceiling: usize) -> Self {
        Self {
            cursor,
            ceiling: ceiling.min(Self::MAX_READ_AHEAD_BYTES),
            target: 0,
            window: Vec::new(),
            consumed: 0,
        }
    }

    /// Changes the ceiling of a wrapper already in use.
    ///
    /// This is [`Self::new`] for a caller that can only decide how far ahead it
    /// may read after it has read something -- a record walk that must not read
    /// past the point where the stream declares itself encrypted, for instance,
    /// opens its window once that declaration can no longer appear. Bytes
    /// already resident stay resident and stay serveable; the new ceiling
    /// governs the fills that follow.
    pub fn set_read_ahead(&mut self, ceiling: usize) {
        self.ceiling = ceiling.min(Self::MAX_READ_AHEAD_BYTES);
        self.target = 0;
    }

    /// Declared logical stream length.
    #[must_use]
    pub fn len(&self) -> u64 {
        self.cursor.len()
    }

    /// Whether the selected stream is logically empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.cursor.is_empty()
    }

    /// Logical stream position of the next byte the caller will be served.
    ///
    /// This is the wrapped cursor's position less the bytes a fill has made
    /// resident but the caller has not consumed.
    #[must_use]
    pub fn position(&self) -> u64 {
        self.cursor.position() - self.buffered() as u64
    }

    /// Resident bytes the caller has not been served yet.
    const fn buffered(&self) -> usize {
        self.window.len() - self.consumed
    }

    /// Reads exactly `output.len()` logical bytes and advances the position.
    ///
    /// Nothing is committed before the step that could fail has succeeded, so
    /// a failed read leaves the position and the window as it found them and a
    /// stable source may be retried. As with
    /// [`SharedOleStreamCursor::read_exact`], the destination may hold a prefix
    /// when a later step fails and callers must discard it on any error.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-data error when the requested bytes exceed the
    /// remaining logical stream, a bounded allocation failure when the window
    /// cannot be grown, or any error the wrapped cursor reports.
    pub fn read_exact(&mut self, output: &mut [u8]) -> Result<(), OleError> {
        let position = self.position();
        let end = position
            .checked_add(output.len() as u64)
            .ok_or_else(|| OleError::InvalidData("stream cursor read end overflow".to_string()))?;
        if end > self.cursor.len() {
            return Err(OleError::InvalidData(format!(
                "stream cursor read {position}..{end} exceeds length {}",
                self.cursor.len()
            )));
        }

        if self.ceiling != 0 && output.len() <= self.buffered() {
            self.serve(output);
            return Ok(());
        }
        if self.ceiling != 0 && output.len() <= self.ceiling {
            self.fill(output.len())?;
            if output.len() > self.buffered() {
                return Err(OleError::CorruptedFile(
                    "stream cursor window is shorter than the read it serves".to_string(),
                ));
            }
            self.serve(output);
            return Ok(());
        }

        // A read the window could never hold, and every read of a wrapper with
        // no ceiling: drain what is resident and let the cursor read the
        // remainder straight into the destination. With no ceiling nothing is
        // ever resident, so this is exactly the request and exactly the
        // observation the bare cursor would have taken, for an empty read as
        // much as for a full one.
        let served = self.buffered().min(output.len());
        let (resident, rest) = output.split_at_mut(served);
        if served != 0 {
            resident.copy_from_slice(&self.window[self.consumed..self.consumed + served]);
        }
        self.cursor.read_exact(rest)?;
        self.consumed += served;
        if self.consumed == self.window.len() {
            self.window.clear();
            self.consumed = 0;
        }
        Ok(())
    }

    /// Advances by `bytes` without publishing them.
    ///
    /// Bytes a fill already made resident are passed in memory; the rest is
    /// passed to [`SharedOleStreamCursor::skip_forward`], which reads nothing
    /// and takes no observation.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-data error when the movement exceeds the stream,
    /// or a corruption error when the chain cannot be traversed consistently.
    pub fn skip_forward(&mut self, bytes: u64) -> Result<(), OleError> {
        if bytes <= self.buffered() as u64 {
            // The chain links under these bytes were walked and checked by the
            // fill that made them resident.
            self.consumed += bytes as usize;
            return Ok(());
        }
        let beyond = bytes - self.buffered() as u64;
        self.cursor.skip_forward(beyond)?;
        self.window.clear();
        self.consumed = 0;
        Ok(())
    }

    /// Moves to a later logical position without publishing bytes.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-data error for backward or out-of-bounds
    /// movement, or a corruption error when the chain cannot be traversed
    /// consistently.
    pub fn skip_to(&mut self, position: u64) -> Result<(), OleError> {
        let current = self.position();
        if position < current {
            return Err(OleError::InvalidData(
                "stream cursor cannot move backwards".to_string(),
            ));
        }
        self.skip_forward(position - current)
    }

    /// Copies `output.len()` resident bytes out of the window.
    fn serve(&mut self, output: &mut [u8]) {
        let from = self.consumed;
        output.copy_from_slice(&self.window[from..from + output.len()]);
        self.consumed = from + output.len();
    }

    /// Makes at least `need` bytes resident with one cursor read.
    ///
    /// A failed fill leaves the window holding exactly what it held before:
    /// the cursor does not commit a failed read, and the appended tail is
    /// dropped here.
    fn fill(&mut self, need: usize) -> Result<(), OleError> {
        let resident = self.buffered();
        let remaining = self.cursor.len() - self.cursor.position();
        let headroom = self.ceiling.saturating_sub(resident);
        let reach = usize::try_from(remaining).unwrap_or(usize::MAX);
        let wanted = need
            .saturating_sub(resident)
            .max(self.target.min(reach))
            .min(headroom)
            .min(reach);

        self.window.drain(..self.consumed);
        self.consumed = 0;
        self.window
            .try_reserve_exact(wanted)
            .map_err(|source| OleError::allocation("CFB stream cursor window", source))?;
        self.window.resize(resident + wanted, 0);
        if let Err(error) = self.cursor.read_exact(&mut self.window[resident..]) {
            // A failed cursor read leaves its destination undefined, so the
            // window keeps only the bytes earlier fills published.
            self.window.truncate(resident);
            return Err(error);
        }
        // `max` then `min`, never `clamp`: a caller may declare a ceiling below
        // the first growth step, and `clamp` panics when its bounds cross.
        self.target = self
            .target
            .saturating_mul(2)
            .max(Self::FIRST_FILL_BYTES)
            .min(self.ceiling);
        Ok(())
    }
}

/// Walks a validated allocation chain to `ordinal`, starting from `resume`.
///
/// `resume` is the `(sector, walked)` pair a [`StreamChainHint`] retains: the
/// stream's first sector and `0` whenever no hint applies, which is the walk
/// this function performed before hints existed. Every link from `walked` to
/// `ordinal` is traversed and checked here regardless of where the walk began.
fn cursor_chain_sector(
    table: &[u32],
    resume: (u32, usize),
    ordinal: usize,
    table_name: &str,
) -> Result<u32, OleError> {
    let (mut sector, walked) = resume;
    for _ in walked..ordinal {
        sector = next_chain_sector(table, sector, table_name)?;
        if sector == ENDOFCHAIN {
            return Err(OleError::CorruptedFile(format!(
                "{table_name} chain ends before cursor offset"
            )));
        }
    }
    if sector == ENDOFCHAIN {
        return Err(OleError::CorruptedFile(format!(
            "{table_name} chain ends before cursor offset"
        )));
    }
    Ok(sector)
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
        ExecutionLimits, Limits, Resource, ScopedWorkers,
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
        versions: AtomicUsize,
        fail_version: AtomicBool,
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
        change_on_len: AtomicBool,
        fail_next_len: AtomicBool,
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
                versions: AtomicUsize::new(0),
                fail_version: AtomicBool::new(false),
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
                change_on_len: AtomicBool::new(false),
                fail_next_len: AtomicBool::new(false),
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

        fn reset_observation_counts(&self) {
            self.reset_read_count();
            self.versions.store(0, AtomicOrdering::SeqCst);
        }

        fn observation_counts(&self) -> (usize, usize) {
            (
                self.reads.load(AtomicOrdering::SeqCst),
                self.versions.load(AtomicOrdering::SeqCst),
            )
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

        /// Mutates the source while its length is being observed.
        ///
        /// This is the window the open path's second observation used to
        /// cover: the capture has been taken, the length has not yet been
        /// consumed, and the branches that reject the length run before the
        /// post-parse comparison can.
        fn change_while_observing_length(&self) {
            self.change_on_len.store(true, AtomicOrdering::SeqCst);
        }

        fn fail_next_len(&self) {
            self.fail_next_len.store(true, AtomicOrdering::SeqCst);
        }
    }

    impl ReadAt for TestSource {
        fn len(&self) -> io::Result<u64> {
            if self.change_on_len.swap(false, AtomicOrdering::SeqCst) {
                self.revision.fetch_add(1, AtomicOrdering::SeqCst);
            }
            if self.fail_next_len.swap(false, AtomicOrdering::SeqCst) {
                return Err(io::Error::other("injected source length failure"));
            }
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
            self.versions.fetch_add(1, AtomicOrdering::SeqCst);
            if self.fail_version.load(AtomicOrdering::SeqCst) {
                return Err(io::Error::other("injected source-version failure"));
            }
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

    /// `sample_bytes` with a patterned `Large` payload and its first four
    /// sectors rewired into a non-monotonic chain, so a physical run ends
    /// where the chain says rather than where the payload does. The pattern
    /// matters: a constant payload cannot tell a leg that read the sectors in
    /// the wrong order from one that did not.
    fn fragmented_payload() -> Vec<u8> {
        (0..8192usize)
            .map(|index| u8::try_from((index * 37 + 11) % 251).unwrap())
            .collect()
    }

    fn fragmented_large_bytes() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Small"], b"mini stream").unwrap();
        writer
            .create_stream_owned(&["Large"], fragmented_payload())
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let mut bytes = output.into_inner();

        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let entry = parsed.find_entry(&["Large"]).unwrap();
        let sector_size = parsed.index.sector_size;
        let mut chain = Vec::new();
        let mut sector = entry.start_sector;
        for _ in 0..8192usize.div_ceil(sector_size) {
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
        bytes
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
            source_is_owned_immutable: false,
            limits: SharedOleFileLimits::default(),
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

    #[derive(Debug, Default)]
    struct CountingWorkers {
        tasks: AtomicUsize,
        max_wave: AtomicUsize,
    }

    impl ScopedWorkers for CountingWorkers {
        fn run_all<'task>(&self, tasks: &mut [&mut (dyn FnMut() + Send + 'task)]) {
            self.tasks.fetch_add(tasks.len(), AtomicOrdering::Relaxed);
            self.max_wave
                .fetch_max(tasks.len(), AtomicOrdering::Relaxed);
            for task in tasks.iter_mut() {
                task();
            }
        }
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
            Err(OleError::LimitExceeded {
                resource: "input bytes",
                observed,
                maximum: 512,
            }) if observed > 512
        ));
    }

    #[test]
    fn shared_limits_reuse_typed_low_level_limit_validation() {
        assert!(matches!(
            SharedOleFileLimits::new(0),
            Err(OleError::InvalidLimit {
                resource: "CFB input bytes",
                value: 0,
                maximum: OleFileLimits::MAX_INPUT_BYTES,
            })
        ));
        assert!(matches!(
            SharedOleFileLimits::new(SharedOleFileLimits::MAX_INPUT_BYTES + 1),
            Err(OleError::InvalidLimit {
                resource: "CFB input bytes",
                value,
                maximum: OleFileLimits::MAX_INPUT_BYTES,
            }) if value == OleFileLimits::MAX_INPUT_BYTES + 1
        ));
        assert_eq!(
            SharedOleFileLimits::default().max_directory_bytes(),
            SharedOleFileLimits::DEFAULT_MAX_DIRECTORY_BYTES
        );
        assert!(matches!(
            SharedOleFileLimits::default().with_max_directory_bytes(0),
            Err(OleError::InvalidLimit {
                resource: "CFB directory bytes",
                value: 0,
                maximum: OleFileLimits::MAX_DIRECTORY_BYTES,
            })
        ));
        assert!(matches!(
            SharedOleFileLimits::default()
                .with_max_directory_bytes(SharedOleFileLimits::MAX_DIRECTORY_BYTES + 1),
            Err(OleError::InvalidLimit {
                resource: "CFB directory bytes",
                value,
                maximum,
            }) if value == maximum + 1
        ));
    }

    #[test]
    fn shared_open_propagates_directory_byte_limit() {
        let bytes = sample_bytes();
        let source: Arc<dyn ReadAt> = Arc::new(TestSource::new(bytes));
        let limits = SharedOleFileLimits::new(SharedOleFileLimits::MAX_INPUT_BYTES)
            .unwrap()
            .with_max_directory_bytes(511)
            .unwrap();

        assert!(matches!(
            SharedOleFile::open_with_limits(source, limits),
            Err(OleError::LimitExceeded {
                resource: "directory bytes",
                observed,
                maximum,
            }) if maximum == 511 && observed > maximum
        ));
    }

    #[test]
    fn shared_validation_reuses_selected_directory_byte_limit() {
        let mut writer = OleWriter::new();
        for name in ["Workbook", "One", "Two", "Three", "Four"] {
            writer.create_stream(&[name], b"").unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let source = Arc::new(TestSource::new(output.into_inner()));
        let limits = SharedOleFileLimits::new(SharedOleFileLimits::MAX_INPUT_BYTES)
            .unwrap()
            .with_max_directory_bytes(1024)
            .unwrap();
        let mut file = SharedOleFile::open_with_limits(source, limits).unwrap();
        file.limits = file.limits.with_max_directory_bytes(512).unwrap();

        let report = file
            .validate(litchi_core::ValidationLimits::default())
            .unwrap();
        assert!(!report.is_complete());
        assert!(!report.has_errors());
        assert!(matches!(
            report.checks()[0].status(),
            litchi_core::CheckStatus::Blocked { .. }
        ));
    }

    #[test]
    fn shared_limit_rejects_oversized_source_before_positional_reads() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let limits = SharedOleFileLimits::new(512).unwrap();

        assert!(matches!(
            SharedOleFile::open_with_limits(source.clone(), limits),
            Err(OleError::LimitExceeded {
                resource: "input bytes",
                observed,
                maximum: 512,
            }) if observed == source.bytes.len() as u64
        ));
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), 0);
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
            epoch: 0,
            owner_active: false,
            waiters: 2,
            status: MiniFATSingleFlightStatus::CompletedNoHandoff,
        });
        file.minifat_direct_state.store(
            minifat_state_with_slot(entry.sid, MINIFAT_DIRECT_DONE, 0, false, true),
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
    fn delayed_old_epoch_waiter_drop_does_not_decrement_new_handoff() {
        let (bytes, expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();
        *file.minifat_singleflight.slot.lock().unwrap() = Some(MiniFATSingleFlightSlot {
            sid: entry.sid,
            epoch: 1,
            owner_active: false,
            waiters: 1,
            status: MiniFATSingleFlightStatus::Failed,
        });
        file.minifat_singleflight
            .waiter_intent
            .store(1, AtomicOrdering::SeqCst);
        file.minifat_direct_state.store(
            minifat_state_with_slot(entry.sid, MINIFAT_DIRECT_DONE, 1, true, true),
            AtomicOrdering::SeqCst,
        );

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("terminal marker should elect a retry owner");
        };
        let old_waiter = MiniFATDirectWaiter {
            file: &file,
            sid: entry.sid,
            epoch: 1,
            registered: true,
            intent: true,
        };
        let MiniFATDirectClaim::Waiter(mut new_waiter) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("new-epoch caller should register with the retry owner");
        };

        drop(old_waiter);
        assert_eq!(
            file.minifat_singleflight
                .slot
                .lock()
                .unwrap()
                .as_ref()
                .unwrap()
                .waiters,
            1
        );
        owner.publish_success(&expected).unwrap();
        drop(owner);
        assert_eq!(new_waiter.wait().unwrap().unwrap().unwrap(), expected);
        drop(new_waiter);
    }

    #[test]
    fn direct_epoch_exhaustion_fails_closed_to_root_cache() {
        let (bytes, _expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();
        let max_epoch = (1_u32 << 22) - 1;
        file.minifat_direct_state.store(
            minifat_state_with_meta(entry.sid, MINIFAT_DIRECT_DONE, max_epoch, false),
            AtomicOrdering::SeqCst,
        );

        assert!(matches!(
            file.claim_minifat_direct_mode(entry.sid, entry.size)
                .unwrap(),
            MiniFATDirectClaim::Cache
        ));
        assert_eq!(
            minifat_state_kind(file.minifat_direct_state.load(AtomicOrdering::SeqCst)),
            MINIFAT_CACHE_REQUESTED
        );
    }

    #[test]
    fn cache_takeover_preserves_epoch_and_intent_until_waiters_leave() {
        let (bytes, _expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();
        let epoch = 37;
        file.minifat_singleflight
            .waiter_intent
            .store(1, AtomicOrdering::SeqCst);
        file.minifat_direct_state.store(
            minifat_state_with_meta(entry.sid, MINIFAT_DIRECT_DONE, epoch, true),
            AtomicOrdering::SeqCst,
        );

        file.request_ministream_cache();
        let requested = file.minifat_direct_state.load(AtomicOrdering::SeqCst);
        assert_eq!(minifat_state_kind(requested), MINIFAT_CACHE_REQUESTED);
        assert_eq!(minifat_state_epoch(requested), epoch);
        assert!(minifat_state_intent(requested));

        let previous = file.begin_ministream_cache();
        assert_eq!(minifat_state_kind(previous), MINIFAT_CACHE_REQUESTED);
        let in_flight = file.minifat_direct_state.load(AtomicOrdering::SeqCst);
        assert_eq!(minifat_state_kind(in_flight), MINIFAT_CACHE_IN_FLIGHT);
        assert_eq!(minifat_state_epoch(in_flight), epoch);
        assert!(minifat_state_intent(in_flight));

        file.finish_ministream_cache(previous, true);
        let ready = file.minifat_direct_state.load(AtomicOrdering::SeqCst);
        assert_eq!(minifat_state_kind(ready), MINIFAT_CACHE_READY);
        assert_eq!(minifat_state_epoch(ready), epoch);
        assert!(minifat_state_intent(ready));

        file.release_singleflight_intent();
        assert!(!minifat_state_intent(
            file.minifat_direct_state.load(AtomicOrdering::SeqCst)
        ));
    }

    #[test]
    fn intent_count_transition_reasserts_the_same_word_bit() {
        let (bytes, _expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();
        file.minifat_singleflight
            .waiter_intent
            .store(1, AtomicOrdering::SeqCst);
        file.minifat_direct_state.store(
            minifat_state_with_meta(entry.sid, MINIFAT_DIRECT_DONE, 9, false),
            AtomicOrdering::SeqCst,
        );

        assert!(file.reassert_singleflight_intent());
        assert!(minifat_state_intent(
            file.minifat_direct_state.load(AtomicOrdering::SeqCst)
        ));
        file.release_singleflight_intent();
        assert_eq!(
            file.minifat_singleflight
                .waiter_intent
                .load(AtomicOrdering::SeqCst),
            0
        );
        assert!(!minifat_state_intent(
            file.minifat_direct_state.load(AtomicOrdering::SeqCst)
        ));

        assert!(file.announce_singleflight_waiter());
        assert!(minifat_state_intent(
            file.minifat_direct_state.load(AtomicOrdering::SeqCst)
        ));
        file.release_singleflight_intent();
        assert!(!minifat_state_intent(
            file.minifat_direct_state.load(AtomicOrdering::SeqCst)
        ));
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
        source.reset_read_count();
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
    fn waiter_drop_before_owner_release_clears_slot_before_repeat_overlap() {
        let (bytes, expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("first direct caller should own the flight");
        };
        assert!(file.announce_singleflight_waiter());
        let MiniFATDirectClaim::Waiter(mut waiter) =
            file.claim_minifat_direct_slow(entry.sid, true).unwrap()
        else {
            panic!("overlapping caller should register as a waiter");
        };
        owner.publish_success(&expected).unwrap();
        assert_eq!(waiter.wait().unwrap().unwrap().unwrap(), expected);
        drop(waiter);
        assert!(minifat_state_slot_present(
            file.minifat_direct_state.load(AtomicOrdering::Acquire)
        ));
        drop(owner);
        assert!(!minifat_state_slot_present(
            file.minifat_direct_state.load(AtomicOrdering::Acquire)
        ));

        let MiniFATDirectClaim::Owner(mut repeat_owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("repeat should claim a fresh direct flight");
        };
        assert!(file.announce_singleflight_waiter());
        let MiniFATDirectClaim::Waiter(mut repeat_waiter) =
            file.claim_minifat_direct_slow(entry.sid, true).unwrap()
        else {
            panic!("new overlap should register with the fresh flight");
        };
        repeat_owner.publish_success(&expected).unwrap();
        drop(repeat_owner);
        assert_eq!(repeat_waiter.wait().unwrap().unwrap().unwrap(), expected);
        drop(repeat_waiter);
        assert!(!file.minifat_singleflight.slot.lock().unwrap().is_some());
    }

    #[test]
    fn claimant_is_serialized_through_owner_terminal_release() {
        let (bytes, expected) = two_mini_bytes(36);
        let file = Arc::new(shared(Arc::new(TestSource::new(bytes))));
        let entry = file.find_entry(&["Selected"]).unwrap();

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("first direct caller should own the flight");
        };
        assert!(file.announce_singleflight_waiter());
        let MiniFATDirectClaim::Waiter(mut waiter) =
            file.claim_minifat_direct_slow(entry.sid, true).unwrap()
        else {
            panic!("overlapping caller should register as a waiter");
        };
        owner.publish_success(&expected).unwrap();
        assert_eq!(waiter.wait().unwrap().unwrap().unwrap(), expected);
        drop(waiter);

        // With no registered waiters left, the slot is still present until
        // the owner releases. Pause precisely while the owner holds that
        // mutex before its terminal state CAS, then force a claimant into the
        // slow path. The claimant must block on the mutex and observe the
        // terminal state only after the owner has linearized its release.
        file.minifat_singleflight
            .release_pause_ready
            .store(0, AtomicOrdering::SeqCst);
        file.minifat_singleflight
            .release_pause_continue
            .store(0, AtomicOrdering::SeqCst);
        file.minifat_singleflight
            .claim_slow_entered
            .store(0, AtomicOrdering::SeqCst);
        file.minifat_singleflight
            .release_pause
            .store(1, AtomicOrdering::SeqCst);

        let (release_result, claimant_result, claimant_entered, intent_seen) =
            thread::scope(|scope| {
                let release = scope.spawn(move || drop(owner));
                for _ in 0..100_000 {
                    if file
                        .minifat_singleflight
                        .release_pause_ready
                        .load(AtomicOrdering::SeqCst)
                        != 0
                    {
                        break;
                    }
                    thread::yield_now();
                }
                let claimant_file = Arc::clone(&file);
                let claimant = scope.spawn(move || claimant_file.open_stream(&["Selected"]));
                let mut claimant_entered = false;
                for _ in 0..100_000 {
                    if file
                        .minifat_singleflight
                        .claim_slow_entered
                        .load(AtomicOrdering::SeqCst)
                        != 0
                    {
                        claimant_entered = true;
                        break;
                    }
                    thread::yield_now();
                }
                let intent_seen =
                    minifat_state_intent(file.minifat_direct_state.load(AtomicOrdering::Acquire));
                file.minifat_singleflight
                    .release_pause_continue
                    .store(1, AtomicOrdering::SeqCst);
                (
                    release.join(),
                    claimant.join(),
                    claimant_entered,
                    intent_seen,
                )
            });

        assert!(release_result.is_ok());
        assert!(claimant_entered);
        assert!(intent_seen);
        assert_eq!(claimant_result.unwrap().unwrap(), expected);
        assert!(!file.minifat_singleflight.slot.lock().unwrap().is_some());
        let terminal = file.minifat_direct_state.load(AtomicOrdering::Acquire);
        assert_eq!(minifat_state_kind(terminal), MINIFAT_DIRECT_DONE);
        assert_eq!(minifat_state_sid(terminal), entry.sid);
        assert!(!minifat_state_slot_present(terminal));

        // A repeat after the forced intent race must claim a fresh direct
        // owner, proving that the release did not strand an ownerless flight.
        assert_eq!(file.open_stream(&["sElEcTeD"]).unwrap(), expected);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn late_intent_terminalizes_failure_and_completed_without_handoff() {
        let (bytes, expected) = two_mini_bytes(36);
        for completed_without_handoff in [false, true] {
            let file = Arc::new(shared(Arc::new(TestSource::new(bytes.clone()))));
            let entry = file.find_entry(&["Selected"]).unwrap();
            let MiniFATDirectClaim::Owner(mut owner) = file
                .claim_minifat_direct_mode(entry.sid, entry.size)
                .unwrap()
            else {
                panic!("first direct caller should own the flight");
            };
            let owner_epoch = owner.epoch;
            assert!(file.announce_singleflight_waiter());
            let MiniFATDirectClaim::Waiter(waiter) =
                file.claim_minifat_direct_slow(entry.sid, true).unwrap()
            else {
                panic!("overlapping caller should register as a waiter");
            };
            if completed_without_handoff {
                owner.publish_success(&expected).unwrap();
                file.minifat_singleflight
                    .slot
                    .lock()
                    .unwrap()
                    .as_mut()
                    .unwrap()
                    .status = MiniFATSingleFlightStatus::CompletedNoHandoff;
            } else {
                owner.publish_failure().unwrap();
            }
            drop(waiter);

            file.minifat_singleflight
                .release_pause_ready
                .store(0, AtomicOrdering::SeqCst);
            file.minifat_singleflight
                .release_pause_continue
                .store(0, AtomicOrdering::SeqCst);
            file.minifat_singleflight
                .release_pause
                .store(1, AtomicOrdering::SeqCst);
            let late_release = Arc::new(AtomicUsize::new(0));

            let (
                release_result,
                late_result,
                intent_seen,
                terminal,
                terminal_intent,
                terminal_count,
                terminal_epoch,
                terminal_sid,
            ) = thread::scope(|scope| {
                let release = scope.spawn(move || drop(owner));
                for _ in 0..100_000 {
                    if file
                        .minifat_singleflight
                        .release_pause_ready
                        .load(AtomicOrdering::SeqCst)
                        != 0
                    {
                        break;
                    }
                    thread::yield_now();
                }
                let late_file = Arc::clone(&file);
                let late_release_signal = Arc::clone(&late_release);
                let late = scope.spawn(move || {
                    assert!(late_file.announce_singleflight_waiter());
                    while late_release_signal.load(AtomicOrdering::SeqCst) == 0 {
                        thread::yield_now();
                    }
                    late_file.release_singleflight_intent();
                });
                let mut intent_seen = false;
                for _ in 0..100_000 {
                    let state = file.minifat_direct_state.load(AtomicOrdering::Acquire);
                    if minifat_state_intent(state)
                        && file
                            .minifat_singleflight
                            .waiter_intent
                            .load(AtomicOrdering::Acquire)
                            != 0
                    {
                        intent_seen = true;
                        break;
                    }
                    thread::yield_now();
                }
                file.minifat_singleflight
                    .release_pause_continue
                    .store(1, AtomicOrdering::SeqCst);
                let release_result = release.join();
                // Capture the terminal metadata before allowing the late
                // guard to release its intent count.
                let terminal = file.minifat_direct_state.load(AtomicOrdering::Acquire);
                let terminal_intent = minifat_state_intent(terminal);
                let terminal_count = file
                    .minifat_singleflight
                    .waiter_intent
                    .load(AtomicOrdering::Acquire);
                let terminal_epoch = minifat_state_epoch(terminal);
                let terminal_sid = minifat_state_sid(terminal);
                late_release.store(1, AtomicOrdering::SeqCst);
                (
                    release_result,
                    late.join(),
                    intent_seen,
                    terminal,
                    terminal_intent,
                    terminal_count,
                    terminal_epoch,
                    terminal_sid,
                )
            });

            assert!(release_result.is_ok());
            assert!(late_result.is_ok());
            assert!(intent_seen);
            assert_eq!(
                minifat_state_kind(terminal),
                if completed_without_handoff {
                    MINIFAT_DIRECT_DONE
                } else {
                    MINIFAT_DIRECT_UNCLAIMED
                }
            );
            assert!(!minifat_state_slot_present(terminal));
            assert!(terminal_intent);
            assert_eq!(terminal_count, 1);
            assert_eq!(terminal_epoch, owner_epoch);
            assert_eq!(
                terminal_sid,
                if completed_without_handoff {
                    entry.sid
                } else {
                    0
                }
            );
            let after_late = file.minifat_direct_state.load(AtomicOrdering::Acquire);
            assert_eq!(
                file.minifat_singleflight
                    .waiter_intent
                    .load(AtomicOrdering::Acquire),
                0
            );
            assert!(!minifat_state_intent(after_late));
            file.minifat_singleflight
                .release_pause
                .store(0, AtomicOrdering::SeqCst);
            assert_eq!(file.open_stream(&["sElEcTeD"]).unwrap(), expected);
            assert!(!file.mini_stream_is_materialized());
        }
    }

    #[test]
    fn intent_before_registration_gets_a_bounded_handoff() {
        let (bytes, expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();
        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("first direct caller should own the flight");
        };

        assert!(file.announce_singleflight_waiter());
        let MiniFATDirectClaim::Waiter(mut waiter) =
            file.claim_minifat_direct_slow(entry.sid, true).unwrap()
        else {
            panic!("announced caller should register before publication");
        };
        owner.publish_success(&expected).unwrap();
        drop(owner);
        assert_eq!(waiter.wait().unwrap().unwrap().unwrap(), expected);
        drop(waiter);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn owner_completion_before_and_after_intent_never_strands_a_late_caller() {
        let (bytes, expected) = two_mini_bytes(36);
        let file = shared(Arc::new(TestSource::new(bytes)));
        let entry = file.find_entry(&["Selected"]).unwrap();

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("first direct caller should own the flight");
        };
        owner.publish_success(&expected).unwrap();
        drop(owner);

        assert!(file.announce_singleflight_waiter());
        let MiniFATDirectClaim::Owner(mut retry) =
            file.claim_minifat_direct_slow(entry.sid, true).unwrap()
        else {
            panic!("late intent should retry after an already-completed owner");
        };
        retry.publish_success(&expected).unwrap();
        drop(retry);

        let MiniFATDirectClaim::Owner(mut owner) = file
            .claim_minifat_direct_mode(entry.sid, entry.size)
            .unwrap()
        else {
            panic!("same-SID sequential caller should remain a direct owner");
        };
        assert!(file.announce_singleflight_waiter());
        owner.publish_success(&expected).unwrap();
        drop(owner);
        let MiniFATDirectClaim::Owner(mut retry) =
            file.claim_minifat_direct_slow(entry.sid, true).unwrap()
        else {
            panic!("intent racing after publication should retry, not wait");
        };
        retry.publish_success(&expected).unwrap();
        drop(retry);
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
            source_is_owned_immutable: false,
            limits: SharedOleFileLimits::default(),
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
            source_is_owned_immutable: false,
            limits: SharedOleFileLimits::default(),
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
            source_is_owned_immutable: false,
            limits: SharedOleFileLimits::default(),
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
    fn fat_range_read_errors_preserve_stable_source_error_and_retry_successfully() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut output = [0u8; 8];

        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        assert!(matches!(
            file.read_stream_range(&["Large"], 0, &mut output),
            Err(OleError::Io(_))
        ));

        file.read_stream_range(&["Large"], 0, &mut output).unwrap();
        assert_eq!(output, [0xA5; 8]);
    }

    #[test]
    fn fat_range_read_error_prefers_source_change_from_post_read_fence() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        let mut output = [0u8; 8];
        assert!(matches!(
            file.read_stream_range(&["Large"], 0, &mut output),
            Err(OleError::SourceChanged { .. })
        ));
    }

    #[test]
    fn minifat_range_read_error_prefers_source_change_from_post_read_fence() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        source.change_on_read.store(true, AtomicOrdering::SeqCst);

        let mut output = [0u8; 4];
        assert!(matches!(
            file.read_stream_range(&["Small"], 0, &mut output),
            Err(OleError::SourceChanged { .. })
        ));
    }

    #[test]
    fn stream_cursor_fat_sequential_reads_match_ranges_without_restarting_chain() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let expected = vec![0xA5; 8192];
        source.reset_read_count();

        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        assert_eq!(cursor.len(), expected.len() as u64);
        assert_eq!(cursor.position(), 0);
        let mut observed = vec![0u8; expected.len()];
        for chunk in observed.chunks_mut(512) {
            cursor.read_exact(chunk).unwrap();
        }

        assert_eq!(observed, expected);
        assert_eq!(cursor.position(), expected.len() as u64);
        let ranges = source.read_ranges();
        assert_eq!(ranges.len(), expected.len() / 512);
        for (index, (offset, length)) in ranges.into_iter().enumerate() {
            assert_eq!(offset, 512 + index as u64 * 512);
            assert_eq!(length, 512);
        }
    }

    /// Reads one stream twice over the same ascending, non-overlapping
    /// partition: once with a plain `read_stream_range` per piece, and once
    /// with `read_stream_range_hinted` carrying one retained chain position
    /// across all of them. Asserts the two legs are indistinguishable from
    /// outside the reader -- same bytes, same positional reads in the same
    /// order, same number of source-version observations.
    ///
    /// This is the property the `litchi-xls` workbook-globals scan relies on.
    /// The saving is CPU only: the hint must not move a byte of I/O, must not
    /// change how often the retained source is fenced, and must not change
    /// which bytes land in the destination.
    fn assert_hinted_matches_unhinted(
        bytes: &[u8],
        path: &[&str],
        length: usize,
        pieces: &[usize],
    ) {
        assert_eq!(
            pieces.iter().sum::<usize>(),
            length,
            "the partition must cover the stream exactly"
        );

        let plain_source = Arc::new(TestSource::new(bytes.to_vec()));
        let plain_file = shared(Arc::clone(&plain_source));
        let mut plain_output = vec![0u8; length];
        plain_source.reset_observation_counts();
        let mut offset = 0u64;
        for &piece in pieces {
            let start = usize::try_from(offset).unwrap();
            plain_file
                .read_stream_range(path, offset, &mut plain_output[start..start + piece])
                .unwrap();
            offset += piece as u64;
        }
        let plain_ranges = plain_source.read_ranges();
        let plain_counts = plain_source.observation_counts();

        let hinted_source = Arc::new(TestSource::new(bytes.to_vec()));
        let hinted_file = shared(Arc::clone(&hinted_source));
        let mut hinted_output = vec![0u8; length];
        hinted_source.reset_observation_counts();
        let mut hint = hinted_file.chain_hint();
        let mut offset = 0u64;
        for &piece in pieces {
            let start = usize::try_from(offset).unwrap();
            hinted_file
                .read_stream_range_hinted(
                    path,
                    offset,
                    &mut hinted_output[start..start + piece],
                    &mut hint,
                )
                .unwrap();
            offset += piece as u64;
        }
        let hinted_ranges = hinted_source.read_ranges();
        let hinted_counts = hinted_source.observation_counts();

        assert_eq!(hinted_output, plain_output, "payload bytes differ");
        assert_eq!(
            hinted_ranges, plain_ranges,
            "positional reads differ for {path:?} over {pieces:?}"
        );
        assert_eq!(
            hinted_counts, plain_counts,
            "read or source-version observation counts differ for {path:?}"
        );
    }

    /// The fill schedule the workbook-globals scan actually issues: a small
    /// exact prologue followed by a doubling window, ending on a short final
    /// piece. It is the shape that makes the prefix walk quadratic, so it is
    /// the one the differential must cover.
    fn doubling_window_partition(length: usize) -> Vec<usize> {
        let mut pieces = Vec::new();
        let mut filled = 0usize;
        let mut window = 4usize;
        while filled < length {
            let piece = window.min(length - filled);
            pieces.push(piece);
            filled += piece;
            window = window.saturating_mul(2).min(64 * 1024);
        }
        pieces
    }

    fn uniform_partition(length: usize, piece: usize) -> Vec<usize> {
        let mut pieces = vec![piece; length / piece];
        if length % piece != 0 {
            pieces.push(length % piece);
        }
        pieces
    }

    #[test]
    fn a_hinted_fat_read_matches_an_unhinted_one_over_every_partition() {
        let bytes = sample_bytes();
        for piece in [1usize, 7, 511, 512, 513, 1000, 4096, 8192] {
            assert_hinted_matches_unhinted(
                &bytes,
                &["Large"],
                8192,
                &uniform_partition(8192, piece),
            );
        }
        assert_hinted_matches_unhinted(&bytes, &["Large"], 8192, &doubling_window_partition(8192));
    }

    #[test]
    fn a_hinted_read_of_a_fragmented_fat_chain_matches_an_unhinted_one() {
        let bytes = fragmented_large_bytes();
        for piece in [1usize, 512, 700, 1024] {
            assert_hinted_matches_unhinted(
                &bytes,
                &["Large"],
                8192,
                &uniform_partition(8192, piece),
            );
        }
        assert_hinted_matches_unhinted(&bytes, &["Large"], 8192, &doubling_window_partition(8192));
    }

    #[test]
    fn a_hinted_minifat_read_matches_an_unhinted_one() {
        // Change 0574 named the mini-stream-resident case as the way this
        // change could be falsified: a MiniFAT range read reaches the root
        // chain by different code from a FAT one.
        let (bytes, expected) = large_mini_bytes(512, 4095);
        for piece in [1usize, 63, 64, 65, 512, 1000] {
            assert_hinted_matches_unhinted(
                &bytes,
                &["Mini"],
                expected.len(),
                &uniform_partition(expected.len(), piece),
            );
        }
        assert_hinted_matches_unhinted(
            &bytes,
            &["Mini"],
            expected.len(),
            &doubling_window_partition(expected.len()),
        );
    }

    #[test]
    fn a_hinted_read_matches_an_unhinted_one_at_a_4096_byte_sector_size() {
        let (bytes, expected) = large_mini_bytes(4096, 4000);
        for piece in [1usize, 64, 999, 4000] {
            assert_hinted_matches_unhinted(
                &bytes,
                &["Mini"],
                expected.len(),
                &uniform_partition(expected.len(), piece),
            );
        }
    }

    #[test]
    fn a_hint_taken_from_another_stream_is_ignored() {
        // The failure this design must make impossible: a hint that names a
        // different directory entry must never position a read. `Small` is
        // MiniFAT-resident and `Large` is FAT-resident, so the hint carried
        // between them differs in every identity field it has.
        let bytes = sample_bytes();
        let source = Arc::new(TestSource::new(bytes.clone()));
        let file = shared(Arc::clone(&source));
        let mut hint = file.chain_hint();

        let mut small = [0u8; 11];
        file.read_stream_range_hinted(&["Small"], 0, &mut small, &mut hint)
            .unwrap();
        assert_eq!(&small, b"mini stream");

        source.reset_observation_counts();
        let mut hinted = vec![0u8; 4096];
        file.read_stream_range_hinted(&["Large"], 4096, &mut hinted, &mut hint)
            .unwrap();
        let hinted_ranges = source.read_ranges();
        let hinted_counts = source.observation_counts();

        source.reset_observation_counts();
        let mut plain = vec![0u8; 4096];
        file.read_stream_range(&["Large"], 4096, &mut plain)
            .unwrap();

        assert_eq!(hinted, plain);
        assert_eq!(hinted, vec![0xA5; 4096]);
        assert_eq!(hinted_ranges, source.read_ranges());
        assert_eq!(hinted_counts, source.observation_counts());
    }

    #[test]
    fn a_hint_bound_to_another_reader_is_ignored() {
        // A hint borrows the reader it came from, so it cannot outlive it and
        // no second live reader can occupy its address. What remains reachable
        // is two readers alive at once whose stream identities agree in every
        // field a hint records -- same SID, same allocation table, same first
        // sector -- and whose chains then diverge. `fragmented_large_bytes`
        // is `sample_bytes` with `Large`'s chain rewired and the matching
        // sector payloads swapped, so the two files lay the same logical
        // stream out on different physical sectors from ordinal 1 onward.
        let contiguous = sample_bytes();
        let fragmented = fragmented_large_bytes();
        let plain_source = Arc::new(TestSource::new(contiguous.clone()));
        let plain = shared(Arc::clone(&plain_source));
        let other_source = Arc::new(TestSource::new(fragmented.clone()));
        let other = shared(Arc::clone(&other_source));
        assert!(!std::ptr::eq(&plain, &other));

        let plain_entry = plain.find_entry(&["Large"]).unwrap();
        let other_entry = other.find_entry(&["Large"]).unwrap();
        assert_eq!(
            (
                plain_entry.sid,
                plain_entry.is_minifat,
                plain_entry.start_sector
            ),
            (
                other_entry.sid,
                other_entry.is_minifat,
                other_entry.start_sector
            ),
            "the two readers must be separated by reader identity alone"
        );

        let mut foreign = plain.chain_hint();
        let mut warm = vec![0u8; 512];
        plain
            .read_stream_range_hinted(&["Large"], 1_024, &mut warm, &mut foreign)
            .unwrap();

        other_source.reset_observation_counts();
        let mut hinted = vec![0u8; 1_024];
        other
            .read_stream_range_hinted(&["Large"], 1_024, &mut hinted, &mut foreign)
            .unwrap();
        let hinted_ranges = other_source.read_ranges();
        let hinted_counts = other_source.observation_counts();

        other_source.reset_observation_counts();
        let mut unhinted = vec![0u8; 1_024];
        other
            .read_stream_range(&["Large"], 1_024, &mut unhinted)
            .unwrap();

        assert_eq!(
            unhinted[..],
            fragmented_payload()[1_024..2_048],
            "the control leg must read the fragmented file's own bytes"
        );
        assert_eq!(hinted, unhinted, "a foreign hint changed the bytes read");
        assert_eq!(hinted_ranges, other_source.read_ranges());
        assert_eq!(hinted_counts, other_source.observation_counts());
    }

    #[test]
    fn a_hint_past_the_requested_offset_is_ignored() {
        // A chain is singly linked, so a position past the offset cannot serve
        // it. Such a hint must be discarded and the walk restarted at the
        // stream's first sector, never trusted in place.
        let bytes = fragmented_large_bytes();
        let source = Arc::new(TestSource::new(bytes.clone()));
        let file = shared(Arc::clone(&source));
        let mut hint = file.chain_hint();

        let mut forward = vec![0u8; 512];
        file.read_stream_range_hinted(&["Large"], 7_168, &mut forward, &mut hint)
            .unwrap();

        source.reset_observation_counts();
        let mut hinted = vec![0u8; 1_024];
        file.read_stream_range_hinted(&["Large"], 512, &mut hinted, &mut hint)
            .unwrap();
        let hinted_ranges = source.read_ranges();
        let hinted_counts = source.observation_counts();

        source.reset_observation_counts();
        let mut plain = vec![0u8; 1_024];
        file.read_stream_range(&["Large"], 512, &mut plain).unwrap();

        assert_eq!(hinted, plain, "a backward hint changed the bytes read");
        assert_eq!(hinted_ranges, source.read_ranges());
        assert_eq!(hinted_counts, source.observation_counts());
    }

    #[test]
    fn a_hint_does_not_change_a_refusal_or_its_text() {
        // Every refusal the range reader raises must survive a hint unchanged,
        // including one raised past a prefix the hint let the walk skip.
        let bytes = sample_bytes();
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(Arc::clone(&source));
        let mut hint = file.chain_hint();
        let mut scratch = vec![0u8; 512];
        file.read_stream_range_hinted(&["Large"], 4_096, &mut scratch, &mut hint)
            .unwrap();

        let cases: Vec<(Vec<&str>, u64, usize)> = vec![
            (vec!["Large"], 8_192, 1),
            (vec!["Large"], u64::MAX, 1),
            (vec!["Missing"], 0, 1),
            (vec![], 0, 1),
        ];
        for (path, offset, length) in cases {
            let mut hinted_output = vec![0u8; length];
            let hinted = file
                .read_stream_range_hinted(&path, offset, &mut hinted_output, &mut hint)
                .unwrap_err();
            let mut plain_output = vec![0u8; length];
            let plain = file
                .read_stream_range(&path, offset, &mut plain_output)
                .unwrap_err();
            assert_eq!(
                hinted.to_string(),
                plain.to_string(),
                "refusal text changed for {path:?} at {offset}"
            );
        }
    }

    #[test]
    fn a_hint_costs_one_chain_walk_where_an_unhinted_read_costs_one_per_call() {
        // The mechanism, as an arithmetic statement rather than a profile: an
        // unhinted scan walks `0 + 1 + ... + (n-1)` links to read a stream in
        // `n` sector-sized calls; a hinted one walks `n - 1`.
        let bytes = sample_bytes();
        let file = shared(Arc::new(TestSource::new(bytes.clone())));
        let sector_size = file.index.sector_size;
        let sectors = 8192 / sector_size;
        let unhinted_steps = sectors * (sectors - 1) / 2;
        let hinted_steps = sectors - 1;
        assert!(
            unhinted_steps > hinted_steps,
            "the prefix walk must be quadratic in the number of calls"
        );
        assert_eq!((unhinted_steps, hinted_steps), (120, 15));

        assert_hinted_matches_unhinted(
            &bytes,
            &["Large"],
            8192,
            &uniform_partition(8192, sector_size),
        );
    }

    /// Builds a cursor over one stream at every offset in an ascending
    /// sequence, twice: once with a plain `stream_cursor_at` per offset, and
    /// once with `stream_cursor_at_hinted` carrying one retained chain
    /// position across all of them. Asserts the two legs are indistinguishable
    /// from outside the reader -- same bytes, same positional reads in the same
    /// order, same source-version observations, same reported length and
    /// position.
    ///
    /// This is the property the `litchi-xls` shared-string resolver relies on:
    /// it builds one cursor per string cell. The saving is CPU only.
    fn assert_hinted_cursor_matches_unhinted(
        bytes: &[u8],
        path: &[&str],
        offsets: &[u64],
        read_len: usize,
    ) {
        let plain_source = Arc::new(TestSource::new(bytes.to_vec()));
        let plain_file = shared(Arc::clone(&plain_source));
        plain_source.reset_observation_counts();
        let mut plain_payloads = Vec::new();
        let mut plain_states = Vec::new();
        for &offset in offsets {
            let mut cursor = plain_file.stream_cursor_at(path, offset).unwrap();
            plain_states.push((cursor.len(), cursor.position()));
            let mut payload = vec![0u8; read_len];
            cursor.read_exact(&mut payload).unwrap();
            plain_payloads.push(payload);
        }
        let plain_ranges = plain_source.read_ranges();
        let plain_counts = plain_source.observation_counts();

        let hinted_source = Arc::new(TestSource::new(bytes.to_vec()));
        let hinted_file = shared(Arc::clone(&hinted_source));
        hinted_source.reset_observation_counts();
        let mut hint = hinted_file.chain_hint();
        let mut hinted_payloads = Vec::new();
        let mut hinted_states = Vec::new();
        for &offset in offsets {
            let mut cursor = hinted_file
                .stream_cursor_at_hinted(path, offset, &mut hint)
                .unwrap();
            hinted_states.push((cursor.len(), cursor.position()));
            let mut payload = vec![0u8; read_len];
            cursor.read_exact(&mut payload).unwrap();
            hinted_payloads.push(payload);
        }

        assert_eq!(
            hinted_payloads, plain_payloads,
            "payload bytes differ for {path:?} over {offsets:?}"
        );
        assert_eq!(hinted_states, plain_states, "cursor state differs");
        assert_eq!(
            hinted_source.read_ranges(),
            plain_ranges,
            "positional reads differ for {path:?} over {offsets:?}"
        );
        assert_eq!(
            hinted_source.observation_counts(),
            plain_counts,
            "read or source-version observation counts differ for {path:?}"
        );
    }

    /// Every sector boundary of the 8 KiB `Large` stream, plus a sequence that
    /// lands mid-sector, plus one that revisits the same ordinal twice.
    fn ascending_cursor_offsets(length: u64, step: u64) -> Vec<u64> {
        (0..length / step).map(|index| index * step).collect()
    }

    #[test]
    fn a_hinted_fat_cursor_matches_an_unhinted_one_across_a_stream() {
        let bytes = sample_bytes();
        for step in [512u64, 1_024, 7, 513, 1_000] {
            assert_hinted_cursor_matches_unhinted(
                &bytes,
                &["Large"],
                &ascending_cursor_offsets(8_192 - 64, step),
                64,
            );
        }
        // Repeated identical offsets exercise the `resume.ordinal == ordinal`
        // edge, where the walk must skip every link and still produce the
        // sector a cold walk would.
        assert_hinted_cursor_matches_unhinted(
            &bytes,
            &["Large"],
            &[0, 0, 512, 512, 512, 4_096, 4_096, 8_128],
            64,
        );
    }

    #[test]
    fn a_hinted_cursor_on_a_fragmented_fat_chain_matches_an_unhinted_one() {
        // A non-monotonic physical chain: a leg that resumed at the wrong
        // sector reads different bytes rather than merely reading them in a
        // different order.
        let bytes = fragmented_large_bytes();
        for step in [512u64, 1_024, 700] {
            assert_hinted_cursor_matches_unhinted(
                &bytes,
                &["Large"],
                &ascending_cursor_offsets(8_192 - 64, step),
                64,
            );
        }
    }

    #[test]
    fn a_hinted_minifat_cursor_matches_an_unhinted_one_across_a_stream() {
        // The MiniFAT constructor reaches its chain through different code
        // from the FAT one, so it is driven separately rather than assumed.
        let (bytes, expected) = large_mini_bytes(512, 4_095);
        let length = expected.len() as u64;
        for step in [64u64, 128, 63, 500] {
            assert_hinted_cursor_matches_unhinted(
                &bytes,
                &["Mini"],
                &ascending_cursor_offsets(length - 32, step),
                32,
            );
        }
        let (bytes, expected) = large_mini_bytes(4_096, 4_000);
        let length = expected.len() as u64;
        for step in [64u64, 999] {
            assert_hinted_cursor_matches_unhinted(
                &bytes,
                &["Mini"],
                &ascending_cursor_offsets(length - 32, step),
                32,
            );
        }
    }

    #[test]
    fn one_hint_serves_both_a_hinted_read_and_a_hinted_cursor() {
        // The two hinted entry points record ordinals on the same scale, so a
        // caller may carry one hint across both. This asserts that claim
        // rather than leaving it to the doc comment: the interleaved leg must
        // be byte-for-byte and read-for-read identical to the unhinted one.
        let bytes = fragmented_large_bytes();
        let payload = fragmented_payload();

        let plain_source = Arc::new(TestSource::new(bytes.clone()));
        let plain = shared(Arc::clone(&plain_source));
        plain_source.reset_observation_counts();
        let mut plain_bytes = Vec::new();
        for offset in [0u64, 1_024, 2_048, 4_096, 7_680] {
            let mut range = vec![0u8; 256];
            plain
                .read_stream_range(&["Large"], offset, &mut range)
                .unwrap();
            plain_bytes.push(range);
            let mut cursor = plain.stream_cursor_at(&["Large"], offset + 256).unwrap();
            let mut read = vec![0u8; 256];
            cursor.read_exact(&mut read).unwrap();
            plain_bytes.push(read);
        }
        let plain_ranges = plain_source.read_ranges();
        let plain_counts = plain_source.observation_counts();

        let hinted_source = Arc::new(TestSource::new(bytes));
        let hinted = shared(Arc::clone(&hinted_source));
        hinted_source.reset_observation_counts();
        let mut hint = hinted.chain_hint();
        let mut hinted_bytes = Vec::new();
        for offset in [0u64, 1_024, 2_048, 4_096, 7_680] {
            let mut range = vec![0u8; 256];
            hinted
                .read_stream_range_hinted(&["Large"], offset, &mut range, &mut hint)
                .unwrap();
            hinted_bytes.push(range);
            let mut cursor = hinted
                .stream_cursor_at_hinted(&["Large"], offset + 256, &mut hint)
                .unwrap();
            let mut read = vec![0u8; 256];
            cursor.read_exact(&mut read).unwrap();
            hinted_bytes.push(read);
        }

        assert_eq!(hinted_bytes, plain_bytes);
        assert_eq!(
            hinted_bytes[0],
            payload[0..256],
            "the control must read the file's own bytes"
        );
        assert_eq!(hinted_source.read_ranges(), plain_ranges);
        assert_eq!(hinted_source.observation_counts(), plain_counts);
    }

    #[test]
    fn a_cursor_hint_taken_from_another_stream_is_ignored() {
        // `Small` is MiniFAT-resident and `Large` is FAT-resident, so a hint
        // carried between them differs in every identity field it has.
        let bytes = sample_bytes();
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(Arc::clone(&source));
        let mut hint = file.chain_hint();

        let mut small = file
            .stream_cursor_at_hinted(&["Small"], 0, &mut hint)
            .unwrap();
        let mut small_bytes = [0u8; 11];
        small.read_exact(&mut small_bytes).unwrap();
        assert_eq!(&small_bytes, b"mini stream");

        source.reset_observation_counts();
        let mut hinted = file
            .stream_cursor_at_hinted(&["Large"], 4_096, &mut hint)
            .unwrap();
        let mut hinted_bytes = vec![0u8; 4_096];
        hinted.read_exact(&mut hinted_bytes).unwrap();
        let hinted_ranges = source.read_ranges();
        let hinted_counts = source.observation_counts();

        source.reset_observation_counts();
        let mut plain = file.stream_cursor_at(&["Large"], 4_096).unwrap();
        let mut plain_bytes = vec![0u8; 4_096];
        plain.read_exact(&mut plain_bytes).unwrap();

        assert_eq!(hinted_bytes, plain_bytes);
        assert_eq!(hinted_bytes, vec![0xA5; 4_096]);
        assert_eq!(hinted_ranges, source.read_ranges());
        assert_eq!(hinted_counts, source.observation_counts());
    }

    #[test]
    fn a_cursor_hint_bound_to_another_reader_is_ignored() {
        // Two readers whose stream identities agree in every field a hint
        // records -- same SID, same allocation table, same first sector -- and
        // whose chains then diverge, so reader identity is the only thing that
        // can separate them. `fragmented_large_bytes` is `sample_bytes` with
        // `Large`'s chain rewired and the matching sector payloads swapped.
        let plain_source = Arc::new(TestSource::new(sample_bytes()));
        let plain = shared(Arc::clone(&plain_source));
        let other_source = Arc::new(TestSource::new(fragmented_large_bytes()));
        let other = shared(Arc::clone(&other_source));
        assert!(!std::ptr::eq(&plain, &other));

        let plain_entry = plain.find_entry(&["Large"]).unwrap();
        let other_entry = other.find_entry(&["Large"]).unwrap();
        assert_eq!(
            (
                plain_entry.sid,
                plain_entry.is_minifat,
                plain_entry.start_sector
            ),
            (
                other_entry.sid,
                other_entry.is_minifat,
                other_entry.start_sector
            ),
            "the two readers must be separated by reader identity alone"
        );

        let mut foreign = plain.chain_hint();
        let mut warm = plain
            .stream_cursor_at_hinted(&["Large"], 1_024, &mut foreign)
            .unwrap();
        let mut discard = vec![0u8; 512];
        warm.read_exact(&mut discard).unwrap();

        other_source.reset_observation_counts();
        let mut hinted_cursor = other
            .stream_cursor_at_hinted(&["Large"], 1_024, &mut foreign)
            .unwrap();
        let mut hinted = vec![0u8; 1_024];
        hinted_cursor.read_exact(&mut hinted).unwrap();
        let hinted_ranges = other_source.read_ranges();
        let hinted_counts = other_source.observation_counts();

        other_source.reset_observation_counts();
        let mut plain_cursor = other.stream_cursor_at(&["Large"], 1_024).unwrap();
        let mut unhinted = vec![0u8; 1_024];
        plain_cursor.read_exact(&mut unhinted).unwrap();

        assert_eq!(
            unhinted[..],
            fragmented_payload()[1_024..2_048],
            "the control leg must read the fragmented file's own bytes"
        );
        assert_eq!(hinted, unhinted, "a foreign hint changed the bytes read");
        assert_eq!(hinted_ranges, other_source.read_ranges());
        assert_eq!(hinted_counts, other_source.observation_counts());
    }

    #[test]
    fn a_cursor_hint_past_the_requested_offset_is_ignored() {
        // A chain is singly linked, so a position past the offset cannot serve
        // it. Such a hint must be discarded and the walk restarted at the
        // stream's first sector, never trusted in place.
        let bytes = fragmented_large_bytes();
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(Arc::clone(&source));
        let mut hint = file.chain_hint();

        let mut forward = file
            .stream_cursor_at_hinted(&["Large"], 7_168, &mut hint)
            .unwrap();
        let mut discard = vec![0u8; 512];
        forward.read_exact(&mut discard).unwrap();

        source.reset_observation_counts();
        let mut hinted_cursor = file
            .stream_cursor_at_hinted(&["Large"], 512, &mut hint)
            .unwrap();
        let mut hinted = vec![0u8; 1_024];
        hinted_cursor.read_exact(&mut hinted).unwrap();
        let hinted_ranges = source.read_ranges();
        let hinted_counts = source.observation_counts();

        source.reset_observation_counts();
        let mut plain_cursor = file.stream_cursor_at(&["Large"], 512).unwrap();
        let mut plain = vec![0u8; 1_024];
        plain_cursor.read_exact(&mut plain).unwrap();

        assert_eq!(
            plain[..],
            fragmented_payload()[512..1_536],
            "the control leg must read the file's own bytes"
        );
        assert_eq!(hinted, plain, "a backward hint changed the bytes read");
        assert_eq!(hinted_ranges, source.read_ranges());
        assert_eq!(hinted_counts, source.observation_counts());
    }

    #[test]
    fn a_cursor_hint_does_not_change_a_refusal_or_its_text() {
        // Every refusal cursor construction raises must survive a hint
        // unchanged, including one raised after a prefix the hint let the walk
        // skip. The hint is warm, so the comparison is not trivially the
        // unhinted path.
        let bytes = sample_bytes();
        let source = Arc::new(TestSource::new(bytes));
        let file = shared(Arc::clone(&source));
        let mut hint = file.chain_hint();
        file.stream_cursor_at_hinted(&["Large"], 4_096, &mut hint)
            .unwrap();

        let cases: Vec<(Vec<&str>, u64)> = vec![
            (vec!["Large"], 8_193),
            (vec!["Large"], u64::MAX),
            (vec!["Missing"], 0),
            (vec![], 0),
            (vec!["Storage"], 0),
        ];
        for (path, offset) in cases {
            let hinted = file
                .stream_cursor_at_hinted(&path, offset, &mut hint)
                .err()
                .map(|error| (format!("{error:?}"), error.to_string()));
            let plain = file
                .stream_cursor_at(&path, offset)
                .err()
                .map(|error| (format!("{error:?}"), error.to_string()));
            assert!(plain.is_some(), "case {path:?} at {offset} must be refused");
            assert_eq!(
                hinted, plain,
                "refusal type or text changed for {path:?} at {offset}"
            );
        }

        // `offset == length` is the End state rather than a refusal, and must
        // agree too.
        let hinted_end = file
            .stream_cursor_at_hinted(&["Large"], 8_192, &mut hint)
            .unwrap();
        let plain_end = file.stream_cursor_at(&["Large"], 8_192).unwrap();
        assert_eq!(
            (hinted_end.len(), hinted_end.position()),
            (plain_end.len(), plain_end.position())
        );
    }

    /// `sample_bytes` with a patterned `Large` payload, so a leg that resumed
    /// at the wrong sector reads different bytes rather than the same constant.
    fn patterned_large_bytes() -> (Vec<u8>, Vec<u8>) {
        let payload = fragmented_payload();
        let mut writer = OleWriter::new();
        writer.create_stream(&["Small"], b"mini stream").unwrap();
        writer
            .create_stream_owned(&["Large"], payload.clone())
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        (output.into_inner(), payload)
    }

    /// A reader whose `Large` FAT chain is well formed for the first
    /// `poison_ordinal` links and then carries `marker`.
    ///
    /// Open-time validation walks every declared chain exactly
    /// (`validate_stream_allocations`), so such a file cannot be produced by
    /// corrupting bytes and opening them: it would be refused at open. The
    /// index is therefore corrupted after a clean open, which is the same
    /// synthetic-index construction the mini-stream tests in this module use.
    /// Returns the reader, the true chain and the payload.
    fn file_with_poisoned_large_chain(
        poison_ordinal: usize,
        marker: u32,
    ) -> (SharedOleFile, Vec<u32>, Vec<u8>) {
        let (bytes, payload) = patterned_large_bytes();
        let base = shared(Arc::new(TestSource::new(bytes.clone())));
        let (start_sector, sid) = {
            let entry = base.find_entry(&["Large"]).unwrap();
            (entry.start_sector, entry.sid)
        };
        let sector_size = base.index.sector_size;
        let sectors = payload.len().div_ceil(sector_size);
        let mut chain = Vec::with_capacity(sectors);
        let mut sector = start_sector;
        for _ in 0..sectors {
            chain.push(sector);
            sector = base.index.fat[sector as usize];
        }
        assert!(poison_ordinal < sectors);

        let SharedOleFile {
            source,
            expected_version,
            source_is_owned_immutable,
            limits,
            index,
            ..
        } = base;
        let mut index = match Arc::try_unwrap(index) {
            Ok(index) => index,
            Err(_index) => panic!("test owns the parsed index"),
        };
        index.fat[chain[poison_ordinal] as usize] = marker;
        assert_eq!(index.dir_entries[sid as usize].as_ref().unwrap().sid, sid);

        let file = SharedOleFile {
            source,
            expected_version,
            source_is_owned_immutable,
            limits,
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        };
        (file, chain, payload)
    }

    /// A cursor construction reduced to a comparable value: the resolved
    /// length and position on success, the error's type and text on refusal.
    fn cursor_outcome(
        result: Result<SharedOleStreamCursor<'_>, OleError>,
    ) -> Result<(u64, u64), (String, String)> {
        match result {
            Ok(cursor) => Ok((cursor.len(), cursor.position())),
            Err(error) => Err((format!("{error:?}"), error.to_string())),
        }
    }

    #[test]
    fn a_cursor_hint_does_not_change_a_refusal_on_a_malformed_chain() {
        // A chain walk is the deterministic recurrence `s[k+1] = fat[s[k]]`,
        // so resuming at a position that same walk produced must reach the
        // same sector -- and fail at the same link -- as a cold walk. That has
        // to hold on a chain that is malformed past the resumed position, not
        // only on well-formed fixtures, because a malformed chain is the case
        // where a divergence would be a silent contract break rather than a
        // wrong byte.
        const POISON: usize = 5;
        // The four markers a corrupt FAT entry can carry, covering every
        // branch of `next_chain_sector`: a terminator, the first reserved
        // marker, a reserved marker above it, and an in-range-looking index
        // that is outside the table. The last one is deliberately included
        // even though it is *not* refused at the first ordinal past the
        // corruption -- it is refused one link later -- because an outcome
        // that differs between the legs would be just as much a contract
        // break as a refusal that differs.
        for marker in [ENDOFCHAIN, MAXREGSECT, 0xFFFF_FFFB, 0x00FF_FFFF] {
            let (file, _chain, _payload) = file_with_poisoned_large_chain(POISON, marker);
            let mut refusals = 0_usize;

            for offset in [3_072u64, 4_096, 6_144, 8_191] {
                // Leg 1: a fresh hint. Nothing to resume from; this is the
                // pre-change walk.
                let mut fresh = file.chain_hint();
                let hinted_fresh =
                    cursor_outcome(file.stream_cursor_at_hinted(&["Large"], offset, &mut fresh));

                // Leg 2: a hint warmed by a *successful* construction at an
                // ordinal before the corruption, so the resumed walk really
                // does skip links before reaching the bad one.
                let mut warm = file.chain_hint();
                file.stream_cursor_at_hinted(&["Large"], 1_536, &mut warm)
                    .expect("ordinals before the corruption are reachable");
                let hinted_warm =
                    cursor_outcome(file.stream_cursor_at_hinted(&["Large"], offset, &mut warm));

                // The control.
                let plain = cursor_outcome(file.stream_cursor_at(&["Large"], offset));
                if plain.is_err() {
                    refusals += 1;
                }
                assert_eq!(
                    hinted_fresh, plain,
                    "a fresh hint changed the outcome for marker 0x{marker:08X} at {offset}"
                );
                assert_eq!(
                    hinted_warm, plain,
                    "a warm hint changed the outcome for marker 0x{marker:08X} at {offset}"
                );
            }
            assert!(
                refusals > 0,
                "marker 0x{marker:08X} at ordinal {POISON} must refuse at least one offset"
            );

            // Ordinals before the corruption stay readable, and the hinted and
            // unhinted forms agree there too.
            let mut warm = file.chain_hint();
            for offset in [0u64, 512, 1_024, 2_048, 2_560] {
                let hinted = file
                    .stream_cursor_at_hinted(&["Large"], offset, &mut warm)
                    .unwrap();
                let plain = file.stream_cursor_at(&["Large"], offset).unwrap();
                assert_eq!(
                    (hinted.len(), hinted.position()),
                    (plain.len(), plain.position())
                );
            }
        }
    }

    #[test]
    fn a_cursor_hint_actually_resumes_the_walk() {
        // The differential tests above cannot distinguish a working hint from
        // one that is silently discarded every time: both produce identical
        // bytes. This one can. The chain is poisoned at ordinal 2, so a walk
        // that starts at the stream's first sector cannot reach ordinal 8 --
        // and a walk that resumes from a position at ordinal 6 never touches
        // the poisoned link and reaches it exactly.
        //
        // The recorded position is one this reader's own walk produced, so
        // this is the same skip a warm hint performs in production; only the
        // way the position was obtained differs, and that is what makes the
        // skip observable.
        const POISON: usize = 2;
        let (file, chain, payload) = file_with_poisoned_large_chain(POISON, ENDOFCHAIN);
        let (sid, start_sector) = {
            let entry = file.find_entry(&["Large"]).unwrap();
            (entry.sid, entry.start_sector)
        };

        assert!(
            file.stream_cursor_at(&["Large"], 4_096).is_err(),
            "a cold walk must not reach past the corruption"
        );

        let mut hint = file.chain_hint();
        hint.record(&file, sid, false, start_sector, 6, chain[6]);
        let mut cursor = file
            .stream_cursor_at_hinted(&["Large"], 4_096, &mut hint)
            .expect("a resumed walk skips the links it already validated");
        let mut observed = vec![0u8; 512];
        cursor.read_exact(&mut observed).unwrap();
        assert_eq!(observed[..], payload[4_096..4_608]);

        // The same position offered to a *different* reader must be discarded,
        // and the cold walk it falls back to must then fail.
        let (other, _chain, _payload) = file_with_poisoned_large_chain(POISON, ENDOFCHAIN);
        assert!(
            other
                .stream_cursor_at_hinted(&["Large"], 4_096, &mut hint)
                .is_err(),
            "a foreign position must be discarded, not trusted"
        );
    }

    #[test]
    fn a_cursor_hint_costs_one_chain_walk_where_unhinted_construction_costs_one_per_cursor() {
        // The mechanism as an arithmetic statement rather than a profile:
        // building `n` cursors at ascending sector boundaries walks
        // `0 + 1 + ... + (n-1)` links unhinted and `n - 1` hinted.
        let bytes = sample_bytes();
        let file = shared(Arc::new(TestSource::new(bytes.clone())));
        let sector_size = file.index.sector_size as u64;
        let cursors = 8_192 / sector_size;
        let unhinted_steps = cursors * (cursors - 1) / 2;
        let hinted_steps = cursors - 1;
        assert_eq!((unhinted_steps, hinted_steps), (120, 15));

        assert_hinted_cursor_matches_unhinted(
            &bytes,
            &["Large"],
            &ascending_cursor_offsets(8_192 - 64, sector_size),
            64,
        );
    }

    #[test]
    fn stream_cursor_fragmented_fat_reads_match_logical_order() {
        let mut bytes = sample_bytes();
        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let entry = parsed.find_entry(&["Large"]).unwrap();
        let sector_size = parsed.index.sector_size;
        let mut chain = Vec::new();
        let mut sector = entry.start_sector;
        for _ in 0..8192usize.div_ceil(sector_size) {
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
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut observed = vec![0u8; 8192];
        cursor.read_exact(&mut observed).unwrap();

        assert_eq!(observed, vec![0xA5; 8192]);
        assert_eq!(
            source.read_ranges()[..3],
            [
                ((u64::from(first) + 1) * sector_size as u64, sector_size),
                ((u64::from(third) + 1) * sector_size as u64, sector_size),
                ((u64::from(second) + 1) * sector_size as u64, sector_size),
            ]
        );
    }

    #[test]
    fn stream_cursor_minifat_parity_leaves_cache_unmaterialized() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Small"], 0).unwrap();
        let mut observed = [0u8; 11];
        cursor.read_exact(&mut observed).unwrap();

        assert_eq!(&observed, b"mini stream");
        assert_eq!(cursor.position(), 11);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn stream_cursor_minifat_handles_partial_root_tail_like_range_reader() {
        let (bytes, _selected) = two_mini_bytes(4095);
        let parsed = SharedOleFile::open(Arc::new(TestSource::new(bytes.clone()))).unwrap();
        let mut index = match Arc::try_unwrap(parsed.index) {
            Ok(index) => index,
            Err(_index) => panic!("test owns the parsed index"),
        };
        let root = index.root.as_mut().unwrap();
        assert!(root.size >= 8191);
        root.size = 8191;
        let source = Arc::new(TestSource::new(bytes));
        let expected_version = source.version().unwrap();
        let file = SharedOleFile {
            source: source.clone(),
            expected_version,
            source_is_owned_immutable: false,
            limits: SharedOleFileLimits::default(),
            index: Arc::new(index),
            ministream: Mutex::new(None),
            minifat_direct_state: AtomicU64::new(minifat_state(0, MINIFAT_DIRECT_UNCLAIMED)),
            minifat_singleflight: MiniFATSingleFlight::new(),
        };

        let mut expected = vec![0u8; 4095];
        file.read_stream_range(&["Other"], 0, &mut expected)
            .unwrap();
        source.reset_read_count();
        let mut cursor = file.stream_cursor_at(&["Other"], 0).unwrap();
        let mut observed = vec![0u8; expected.len()];
        cursor.read_exact(&mut observed).unwrap();

        assert_eq!(observed, expected);
        assert_eq!(cursor.position(), expected.len() as u64);
        assert!(!file.mini_stream_is_materialized());
    }

    #[test]
    fn stream_cursor_forward_skip_rejects_backward_and_out_of_bounds_moves() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source);
        let mut cursor = file.stream_cursor_at(&["Small"], 0).unwrap();

        cursor.skip_forward(4).unwrap();
        assert_eq!(cursor.position(), 4);
        let mut output = [0u8; 4];
        cursor.read_exact(&mut output).unwrap();
        assert_eq!(&output, b" str");
        assert!(matches!(
            cursor.skip_to(3),
            Err(OleError::InvalidData(message)) if message.contains("backwards")
        ));
        assert!(matches!(
            cursor.skip_forward(4),
            Err(OleError::InvalidData(message)) if message.contains("exceeds length")
        ));
        assert_eq!(cursor.position(), 8);
    }

    #[test]
    fn stream_cursor_skip_does_not_fence_but_next_read_rejects_mutation() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.revision.store(1, AtomicOrdering::SeqCst);

        cursor.skip_forward(512).unwrap();
        assert_eq!(cursor.position(), 512);
        let mut output = [0u8; 4];
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::SourceChanged { .. })
        ));
    }

    /// The frozen baseline the read-ahead is measured against: without it a
    /// cursor issues one positional request, and takes one source observation,
    /// per `read_exact` call, however small the call is.
    #[test]
    fn a_cursor_without_read_ahead_reads_and_fences_once_per_call() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.reset_observation_counts();

        let mut frame = [0u8; 4];
        for _ in 0..64 {
            cursor.read_exact(&mut frame).unwrap();
        }

        let (reads, versions) = source.observation_counts();
        assert_eq!(reads, 64, "one positional request per read");
        assert_eq!(versions, 64, "one source observation per read");
        assert_eq!(cursor.position(), 256);
        assert_eq!(frame, [0xA5; 4]);
    }

    /// A wrapper with no ceiling is the bare cursor: same requests, same
    /// sizes, same observations. This is what a caller holds until it has
    /// established that reading ahead is safe, so it has to cost nothing.
    #[test]
    fn a_window_with_no_ceiling_reads_exactly_as_the_bare_cursor_does() {
        let plain_source = Arc::new(TestSource::new(sample_bytes()));
        let plain_file = shared(Arc::clone(&plain_source));
        let mut plain = plain_file.stream_cursor_at(&["Large"], 0).unwrap();
        plain_source.reset_observation_counts();
        let mut frame = [0u8; 4];
        for _ in 0..64 {
            plain.read_exact(&mut frame).unwrap();
        }
        plain.skip_forward(16).unwrap();
        plain.read_exact(&mut frame).unwrap();
        plain.read_exact(&mut []).unwrap();

        let window_source = Arc::new(TestSource::new(sample_bytes()));
        let window_file = shared(Arc::clone(&window_source));
        let mut windowed =
            BufferedOleStreamCursor::new(window_file.stream_cursor_at(&["Large"], 0).unwrap(), 0);
        window_source.reset_observation_counts();
        for _ in 0..64 {
            windowed.read_exact(&mut frame).unwrap();
        }
        windowed.skip_forward(16).unwrap();
        windowed.read_exact(&mut frame).unwrap();
        // An empty read is a read: without a ceiling it must still take the
        // observation the bare cursor takes for it.
        windowed.read_exact(&mut []).unwrap();

        assert_eq!(plain_source.read_ranges(), window_source.read_ranges());
        assert_eq!(
            plain_source.observation_counts(),
            window_source.observation_counts()
        );
        assert_eq!(plain.position(), windowed.position());
    }

    /// One fill per contiguous run, on change 0568's doubling schedule, and one
    /// observation per fill rather than one per served read. The request list
    /// is asserted exactly: it is the whole claim.
    #[test]
    fn read_ahead_serves_a_framed_walk_from_one_fill_per_run() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        source.reset_observation_counts();

        let mut observed = Vec::new();
        let mut frame = [0u8; 4];
        for _ in 0..2048 {
            cursor.read_exact(&mut frame).unwrap();
            observed.extend_from_slice(&frame);
        }

        assert_eq!(observed, vec![0xA5; 8192]);
        assert_eq!(cursor.position(), 8192);
        let (reads, versions) = source.observation_counts();
        assert_eq!(reads, 6, "2,048 four-byte reads cost six fills");
        assert_eq!(versions, 6, "one observation per fill, not one per read");
        assert_eq!(
            source.read_ranges(),
            vec![
                (512, 4),
                (516, 512),
                (1028, 1024),
                (2052, 2048),
                (4100, 4096),
                (8196, 508),
            ],
            "the first fill is exact and later fills double to the ceiling"
        );
    }

    /// A fill never crosses the declared stream length, so read-ahead moves
    /// when bytes are read, never which bytes exist to be read.
    #[test]
    fn read_ahead_never_reads_past_the_declared_stream() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 8192 - 6).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        source.reset_observation_counts();

        let mut frame = [0u8; 3];
        cursor.read_exact(&mut frame).unwrap();
        cursor.read_exact(&mut frame).unwrap();

        assert_eq!(cursor.position(), 8192);
        let ranges = source.read_ranges();
        let read_bytes: usize = ranges.iter().map(|(_offset, length)| *length).sum();
        assert_eq!(read_bytes, 6, "the tail fill is clamped by the stream end");
        assert!(matches!(
            cursor.read_exact(&mut frame),
            Err(OleError::InvalidData(message)) if message.contains("exceeds length")
        ));
    }

    /// A caller that abandons a stream early reads at most one growth step
    /// beyond what it consumed, which is what change 0568's schedule buys.
    #[test]
    fn read_ahead_over_reads_at_most_one_growth_step() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        source.reset_observation_counts();

        let mut frame = [0u8; 4];
        cursor.read_exact(&mut frame).unwrap();
        cursor.read_exact(&mut frame).unwrap();
        drop(cursor);

        let read_bytes: usize = source
            .read_ranges()
            .iter()
            .map(|(_offset, length)| *length)
            .sum();
        assert_eq!(
            read_bytes,
            4 + BufferedOleStreamCursor::FIRST_FILL_BYTES,
            "eight consumed bytes cost 516 read bytes, not a ceiling-sized fill"
        );
    }

    /// The change-under-read sweep of change 0621, extended to fill boundaries:
    /// a read the window covers does not observe the source, so a mutation
    /// between two served reads is reported by the next fill.
    #[test]
    fn read_ahead_reports_a_mutation_at_the_next_fill_not_at_a_served_read() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::FIRST_FILL_BYTES);
        let mut frame = [0u8; 4];
        // Two reads: the first fills four bytes, the second fills 512 and
        // leaves 508 of them resident.
        cursor.read_exact(&mut frame).unwrap();
        cursor.read_exact(&mut frame).unwrap();

        source.revision.store(1, AtomicOrdering::SeqCst);
        for _ in 0..127 {
            cursor
                .read_exact(&mut frame)
                .expect("a read the window covers publishes bytes the last fill fenced");
        }
        assert_eq!(cursor.position(), 516);
        assert!(
            matches!(
                cursor.read_exact(&mut frame),
                Err(OleError::SourceChanged { .. })
            ),
            "the first read that has to fill refuses the changed source"
        );
        assert_eq!(cursor.position(), 516, "a failed read does not commit");
    }

    /// The other half of the sweep: a mutation before the first fill is
    /// refused by that fill, exactly as it was refused by the first read
    /// before read-ahead existed.
    #[test]
    fn read_ahead_reports_a_mutation_taken_before_its_first_fill() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        source.revision.store(1, AtomicOrdering::SeqCst);

        let mut frame = [0u8; 4];
        assert!(matches!(
            cursor.read_exact(&mut frame),
            Err(OleError::SourceChanged { .. })
        ));
        assert_eq!(cursor.position(), 0);
    }

    /// A failed fill leaves the logical position, the frontier and the window
    /// exactly as it found them, so a stable source may be retried and serves
    /// the same bytes.
    #[test]
    fn a_failed_fill_does_not_commit_the_cursor() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);

        let mut frame = [0u8; 4];
        assert!(matches!(
            cursor.read_exact(&mut frame),
            Err(OleError::Io(_))
        ));
        assert_eq!(cursor.position(), 0);

        cursor.read_exact(&mut frame).unwrap();
        assert_eq!(frame, [0xA5; 4]);
        assert_eq!(cursor.position(), 4);
    }

    /// Read-ahead moves requests, never bytes: over a chain whose sectors are
    /// stored out of logical order the window serves exactly what a plain
    /// cursor serves.
    #[test]
    fn read_ahead_matches_a_plain_cursor_on_a_fragmented_chain() {
        let bytes = fragmented_large_bytes();
        let payload = fragmented_payload();

        let plain_source = Arc::new(TestSource::new(bytes.clone()));
        let plain_file = shared(Arc::clone(&plain_source));
        let mut plain_cursor = plain_file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut plain = vec![0u8; payload.len()];
        for chunk in plain.chunks_mut(7) {
            plain_cursor.read_exact(chunk).unwrap();
        }

        let windowed_source = Arc::new(TestSource::new(bytes));
        let windowed_file = shared(Arc::clone(&windowed_source));
        let mut windowed_cursor = BufferedOleStreamCursor::new(
            windowed_file.stream_cursor_at(&["Large"], 0).unwrap(),
            BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES,
        );
        let mut windowed = vec![0u8; payload.len()];
        for chunk in windowed.chunks_mut(7) {
            windowed_cursor.read_exact(chunk).unwrap();
        }

        assert_eq!(plain, payload);
        assert_eq!(windowed, payload);
        assert!(
            windowed_source.read_ranges().len() < plain_source.read_ranges().len(),
            "the window still costs fewer requests on a fragmented chain"
        );
    }

    /// The same parity over a MiniFAT stream, which resolves its physical
    /// ranges through the root entry rather than the FAT.
    #[test]
    fn read_ahead_matches_a_plain_cursor_on_a_minifat_stream() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Small"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        let mut observed = [0u8; 11];
        for chunk in observed.chunks_mut(3) {
            cursor.read_exact(chunk).unwrap();
        }

        assert_eq!(&observed, b"mini stream");
        assert_eq!(cursor.position(), 11);
        assert!(!file.mini_stream_is_materialized());
    }

    /// A read the window could never hold drains what is resident and takes
    /// the rest straight into the destination.
    #[test]
    fn a_read_larger_than_the_ceiling_bypasses_the_window() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor = BufferedOleStreamCursor::new(cursor, 64);
        let mut frame = [0u8; 4];
        cursor.read_exact(&mut frame).unwrap();
        let mut bulk = vec![0u8; 4096];
        cursor.read_exact(&mut bulk).unwrap();

        assert_eq!(bulk, vec![0xA5; 4096]);
        assert_eq!(cursor.position(), 4100);
        let mut tail = vec![0u8; 8192 - 4100];
        cursor.read_exact(&mut tail).unwrap();
        assert_eq!(tail, vec![0xA5; 8192 - 4100]);
    }

    /// A skip inside the window passes bytes the last fill already walked and
    /// fenced; a skip past it drops the window and walks the chain.
    #[test]
    fn read_ahead_skips_inside_the_window_without_reading() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut cursor =
            BufferedOleStreamCursor::new(cursor, BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES);
        let mut frame = [0u8; 4];
        cursor.read_exact(&mut frame).unwrap();
        cursor.read_exact(&mut frame).unwrap();
        source.reset_observation_counts();

        cursor.skip_forward(64).unwrap();
        assert_eq!(cursor.position(), 72);
        let (reads, versions) = source.observation_counts();
        assert_eq!(reads, 0, "a skip inside the window reads nothing");
        assert_eq!(versions, 0, "and takes no observation, as skips never did");

        cursor.skip_forward(4096).unwrap();
        assert_eq!(cursor.position(), 4168);
        assert_eq!(
            source.observation_counts().0,
            0,
            "a skip past it reads nothing either"
        );
        cursor.read_exact(&mut frame).unwrap();
        assert_eq!(frame, [0xA5; 4]);
    }

    /// The declared ceiling is clamped to the bound the type publishes, so no
    /// caller can ask one cursor to hold an unbounded window.
    #[test]
    fn the_read_ahead_ceiling_is_clamped_to_the_published_bound() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(Arc::clone(&source));
        let unbounded = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut unbounded = BufferedOleStreamCursor::new(unbounded, usize::MAX);
        let mut bounded = BufferedOleStreamCursor::new(
            file.stream_cursor_at(&["Large"], 0).unwrap(),
            BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES,
        );
        let mut left = [0u8; 4];
        let mut right = [0u8; 4];
        for _ in 0..2048 {
            unbounded.read_exact(&mut left).unwrap();
            bounded.read_exact(&mut right).unwrap();
            assert_eq!(left, right);
        }
        assert_eq!(unbounded.position(), bounded.position());
        assert_eq!(
            BufferedOleStreamCursor::MAX_READ_AHEAD_BYTES,
            64 * 1024,
            "the published bound is the worksheet window's 64 KiB"
        );
    }

    #[test]
    fn read_exact_reports_a_fence_failure_after_a_successful_read() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.fail_version.store(true, AtomicOrdering::SeqCst);
        let mut output = [0u8; 4];
        // The bytes were read, but they are never published when the
        // trailing fence cannot observe the source.
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::Io(_))
        ));
        assert_eq!(cursor.position(), 0, "a failed read does not commit");
    }

    #[test]
    fn read_exact_reports_the_payload_error_when_the_fence_also_fails() {
        // When the payload read and the fence fail together the payload error
        // is reported: the fence can only override an earlier error with
        // `SourceChanged`, and an unobservable source is not a proven change.
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        source.fail_version.store(true, AtomicOrdering::SeqCst);
        let mut output = [0u8; 4];
        let error = cursor.read_exact(&mut output).unwrap_err();
        assert!(
            matches!(error, OleError::Io(_)),
            "unexpected error: {error:?}"
        );
        assert!(
            format!("{error}").contains("injected positional read failure"),
            "expected the payload failure, found: {error}"
        );
    }

    #[test]
    fn an_open_observes_the_source_twice_around_its_parse() {
        // The open path captures the source version, consumes the length,
        // parses, and compares. That is one opening observation and one
        // closing observation; the parse itself takes none, because
        // `ReadAtCursor` reads through the source directly. A third
        // observation between the capture and the parse would only re-prove
        // what the closing one proves.
        let source = Arc::new(TestSource::new(sample_bytes()));
        source.reset_observation_counts();
        let file = shared(Arc::clone(&source));
        let (_reads, versions) = source.observation_counts();
        assert_eq!(
            versions, 2,
            "an open observes the source once before and once after parsing"
        );
        assert_eq!(file.file_size(), source.bytes.len() as u64);
    }

    #[test]
    fn an_open_refuses_a_changed_source_over_its_input_limit() {
        // A mutation that lands while the length is observed must be reported
        // as `SourceChanged`, not as the limit the mutated length violates.
        // The relocated fence on this branch is what keeps that precedence.
        let bytes = sample_bytes();
        let source = Arc::new(TestSource::new(bytes.clone()));
        source.change_while_observing_length();
        let limits = SharedOleFileLimits::new(64).unwrap();
        let error = SharedOleFile::open_with_limits(source, limits).unwrap_err();
        assert!(
            matches!(error, OleError::SourceChanged { .. }),
            "expected SourceChanged, found: {error:?}"
        );

        // Without the mutation the same call still reports the limit.
        let stable = Arc::new(TestSource::new(bytes));
        let limits = SharedOleFileLimits::new(64).unwrap();
        let error = SharedOleFile::open_with_limits(stable, limits).unwrap_err();
        assert!(
            matches!(
                error,
                OleError::LimitExceeded {
                    resource: "input bytes",
                    ..
                }
            ),
            "expected LimitExceeded, found: {error:?}"
        );
    }

    #[test]
    fn an_open_refuses_a_changed_source_over_a_failed_length() {
        // Same precedence on the other branch that leaves the open before the
        // post-parse comparison: the length could not be observed at all.
        let bytes = sample_bytes();
        let source = Arc::new(TestSource::new(bytes.clone()));
        source.change_while_observing_length();
        source.fail_next_len();
        let error = SharedOleFile::open(source).unwrap_err();
        assert!(
            matches!(error, OleError::SourceChanged { .. }),
            "expected SourceChanged, found: {error:?}"
        );

        let stable = Arc::new(TestSource::new(bytes));
        stable.fail_next_len();
        let error = SharedOleFile::open(stable).unwrap_err();
        assert!(
            format!("{error}").contains("injected source length failure"),
            "expected the length failure, found: {error}"
        );
    }

    #[test]
    fn checked_reads_observe_the_source_version_exactly_once() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());

        // One cursor read publishes bytes and takes one trailing observation.
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.reset_observation_counts();
        let mut output = [0u8; 4];
        cursor.read_exact(&mut output).unwrap();
        assert_eq!(output, [0xA5; 4]);
        assert_eq!(source.observation_counts(), (1, 1));

        // Skipping publishes no bytes and takes no observation, so the
        // preceding read's trailing observation still opens the next bracket.
        source.reset_observation_counts();
        cursor.skip_forward(4).unwrap();
        assert_eq!(source.observation_counts(), (0, 0));

        source.reset_observation_counts();
        cursor.read_exact(&mut output).unwrap();
        assert_eq!(source.observation_counts(), (1, 1));

        // One bounded range read behaves the same way.
        source.reset_observation_counts();
        let mut range = [0u8; 16];
        file.read_stream_range(&["Large"], 0, &mut range).unwrap();
        assert_eq!(range, [0xA5; 16]);
        assert_eq!(source.observation_counts(), (1, 1));

        // An empty range read still observes the version once, and reads
        // nothing.
        source.reset_observation_counts();
        file.read_stream_range(&["Large"], 0, &mut []).unwrap();
        assert_eq!(source.observation_counts(), (0, 1));
    }

    #[test]
    fn first_cursor_read_after_open_reports_a_change_from_its_trailing_fence() {
        // The opening observation of the first read's bracket is the one taken
        // while opening the reader. A mutation that lands after that point is
        // reported by the read's own trailing fence, and the destination is
        // not authoritative on that error.
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.revision.store(1, AtomicOrdering::SeqCst);
        source.reset_observation_counts();
        let mut output = [0u8; 4];
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::SourceChanged { .. })
        ));
        let (reads, versions) = source.observation_counts();
        assert_eq!(reads, 1, "the payload read is attempted before the fence");
        assert_eq!(versions, 1, "one trailing observation reports the change");
        assert_eq!(cursor.position(), 0, "a failed read does not commit");
    }

    #[test]
    fn stream_cursor_reports_source_changes_around_hostile_reads() {
        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        source.revision.store(1, AtomicOrdering::SeqCst);
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        let mut output = [0u8; 4];
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::SourceChanged { .. })
        ));

        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.change_on_read.store(true, AtomicOrdering::SeqCst);
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::SourceChanged { .. })
        ));

        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::Io(_))
        ));
        cursor.read_exact(&mut output).unwrap();
        assert_eq!(output, [0xA5; 4]);

        let source = Arc::new(TestSource::new(sample_bytes()));
        let file = shared(source.clone());
        let mut cursor = file.stream_cursor_at(&["Large"], 0).unwrap();
        source.fail_next_read.store(true, AtomicOrdering::SeqCst);
        source.change_on_read.store(true, AtomicOrdering::SeqCst);
        assert!(matches!(
            cursor.read_exact(&mut output),
            Err(OleError::SourceChanged { .. })
        ));
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
    fn bulk_reads_use_caller_facility_and_release_operation_worker_width() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(Arc::clone(&source));
        let paths: &[&[&str]] = &[&["Second"], &["Small"], &["First"]];
        let (_cancel, token) = CancellationSource::pair();
        let budget = Budget::root(
            "shared-cfb-caller-workers",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(2, 2, u64::MAX),
        );
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(2).unwrap(),
            NonZeroUsize::new(2).unwrap(),
            NonZeroU64::new(16_384).unwrap(),
            1,
        )
        .unwrap();
        let workers = Arc::new(CountingWorkers::default());
        let context = ExecutionContext::new(budget.clone(), token, limits)
            .with_scoped_workers(workers.clone());
        let session = file.bulk_read(context);
        let result = session.read_streams(paths).unwrap();

        assert_eq!(result[0][0], 0x22);
        assert_eq!(result[2][0], 0x11);
        assert!(workers.tasks.load(AtomicOrdering::Relaxed) >= 2);
        assert!(workers.max_wave.load(AtomicOrdering::Relaxed) <= 2);
        assert_eq!(session.pool_build_count(), 0);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
        drop(session);
        assert_eq!(budget.used(Resource::Workers), 0);
    }

    #[test]
    fn bulk_reads_refuse_zero_io_before_the_first_payload_read() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(Arc::clone(&source));
        let baseline = source.reads.load(AtomicOrdering::SeqCst);
        let paths: &[&[&str]] = &[&["Second"], &["First"]];
        let (_cancel, token) = CancellationSource::pair();
        let budget = Budget::root(
            "shared-cfb-zero-io",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(2, 0, u64::MAX),
        );
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(2).unwrap(),
            NonZeroUsize::new(2).unwrap(),
            NonZeroU64::new(16_384).unwrap(),
            1,
        )
        .unwrap();
        let session = file.bulk_read(ExecutionContext::new(budget.clone(), token, limits));
        assert!(matches!(
            session.read_streams(paths),
            Err(SharedOleBulkError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::IoConcurrency
        ));
        assert_eq!(source.reads.load(AtomicOrdering::SeqCst), baseline);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
        assert_eq!(budget.used(Resource::CpuTasks), 0);
    }

    #[test]
    fn serial_bulk_failure_releases_worker_and_io_admissions() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(Arc::clone(&source));
        source.fail_all_reads();
        let paths: &[&[&str]] = &[&["Second"]];
        let (_cancel, token) = CancellationSource::pair();
        let budget = Budget::root(
            "shared-cfb-serial-failure",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(1, 1, u64::MAX),
        );
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(16_384).unwrap(),
            0,
        )
        .unwrap();
        let session = file.bulk_read(ExecutionContext::new(budget.clone(), token, limits));
        assert!(matches!(
            session.read_streams(paths),
            Err(SharedOleBulkError::Ole(_))
        ));
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
    }

    #[test]
    fn cancelled_bulk_read_releases_worker_and_io_admissions() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(Arc::clone(&source));
        let (cancel, token) = CancellationSource::pair();
        source.cancel_on_next_read(cancel);
        let budget = Budget::root(
            "shared-cfb-cancelled-read",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(1, 1, u64::MAX),
        );
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(16_384).unwrap(),
            0,
        )
        .unwrap();
        let session = file.bulk_read(ExecutionContext::new(budget.clone(), token, limits));

        assert!(matches!(
            session.read_streams(&[&["Second"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::Cancelled))
        ));
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
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
    fn cached_bulk_pool_caps_later_io_narrowing() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(Arc::clone(&source));
        let budget = Budget::root(
            "shared-cfb-cached-pool-narrowing",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(4, 4, u64::MAX),
        );
        let (_cancel, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(4).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(64 * 1024).unwrap(),
            1,
        )
        .unwrap();
        let session = file.bulk_read(ExecutionContext::new(budget.clone(), token, limits));
        let first_paths: &[&[&str]] = &[&["First"], &["Second"], &["First"], &["Second"]];
        session.read_streams(first_paths).unwrap();
        assert_eq!(session.pool_build_count(), 1);

        // Keep two of the four root I/O permits occupied while reusing the
        // cached four-thread pool. The operation must submit two-wide waves,
        // rather than handing the whole batch to that pool.
        let io_hold = budget.reserve(Resource::IoConcurrency, 2).unwrap();
        let second_paths: &[&[&str]] = &[&["First"], &["Second"]];
        source.synchronize_next_two_reads();
        session.read_streams(second_paths).unwrap();
        let max_active = source.max_active_reads.load(AtomicOrdering::SeqCst);
        assert!(
            max_active >= 2,
            "the narrowed operation did not overlap reads"
        );
        assert!(
            max_active <= 2,
            "cached pool exceeded the admitted I/O width"
        );
        assert_eq!(budget.used(Resource::Workers), 4);
        drop(io_hold);
        drop(session);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
    }

    #[test]
    fn zero_cpu_budget_refuses_before_building_a_bulk_pool() {
        let source = Arc::new(TestSource::new(bulk_bytes()));
        let file = shared(source);
        let budget = Budget::root(
            "shared-cfb-zero-cpu",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
                .with_execution_io(4, 4, 0),
        );
        let (_cancel, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(4).unwrap(),
            NonZeroUsize::new(4).unwrap(),
            NonZeroU64::new(64 * 1024).unwrap(),
            1,
        )
        .unwrap();
        let session = file.bulk_read(ExecutionContext::new(budget.clone(), token, limits));

        assert!(matches!(
            session.read_streams(&[&["First"], &["Second"]]),
            Err(SharedOleBulkError::Execution(ExecutionError::ResourceLimit(limit)))
                if limit.resource == Resource::CpuTasks
        ));
        assert_eq!(session.pool_build_count(), 0);
        assert_eq!(budget.used(Resource::Workers), 0);
        assert_eq!(budget.used(Resource::IoConcurrency), 0);
        assert_eq!(budget.used(Resource::CpuTasks), 0);
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
