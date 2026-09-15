//! A bounded, archive-owned forward-start read-ahead window.
//!
//! The window is deliberately kept below the `SourceReader` adapter.  It is
//! therefore shared by ZIP metadata and lazy member reads, while source
//! artifacts and the exact publication helpers continue to use the snapshot
//! directly.  A small admission gate serializes forward operations and lets
//! exact-publication transitions drain them without holding a cache lock over
//! a provider call.

use crate::error::{OpcError, Result};
use litchi_core::{Reservation, Resource};
use std::fmt;
use std::sync::{
    Arc, Condvar, Mutex, MutexGuard,
    atomic::{AtomicBool, Ordering},
};
use std::time::Duration;

/// Hard upper bound for one forward-start source read-ahead window.
pub(super) const MAX_SOURCE_READ_AHEAD_BYTES: usize = 64 * 1024;

/// Explicit source-read policy for an archive-owned source adapter.
///
/// The window size is private so callers cannot construct an unvalidated
/// policy value.  The zero-sized policy is the default and leaves the source
/// adapter's existing positional-read behavior unchanged.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceReadPolicy {
    window_bytes: usize,
}

impl Default for SourceReadPolicy {
    fn default() -> Self {
        Self::exact()
    }
}

impl SourceReadPolicy {
    /// Return the exact positional-read policy.
    #[must_use]
    pub const fn exact() -> Self {
        Self { window_bytes: 0 }
    }

    /// Validate and construct a forward-start policy.
    pub fn forward_start(window_bytes: usize) -> std::result::Result<Self, SourceReadPolicyError> {
        if window_bytes == 0 || window_bytes > MAX_SOURCE_READ_AHEAD_BYTES {
            return Err(SourceReadPolicyError {
                requested: window_bytes,
                maximum: MAX_SOURCE_READ_AHEAD_BYTES,
            });
        }
        Ok(Self { window_bytes })
    }

    /// Return the configured window size, or zero for exact mode.
    #[must_use]
    pub const fn window_bytes(self) -> usize {
        self.window_bytes
    }

    fn validate(self) -> std::result::Result<(), SourceReadPolicyError> {
        // Zero is the deliberately supported exact policy.  A zero-sized
        // forward request is rejected by `forward_start`; once represented by
        // the policy value, zero means that the archive adapter stays exact.
        if self.window_bytes > MAX_SOURCE_READ_AHEAD_BYTES {
            Err(SourceReadPolicyError {
                requested: self.window_bytes,
                maximum: MAX_SOURCE_READ_AHEAD_BYTES,
            })
        } else {
            Ok(())
        }
    }
}

/// Error returned when a source-read window is outside the finite policy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceReadPolicyError {
    /// Requested window size in bytes.
    pub requested: usize,
    /// Maximum permitted window size in bytes.
    pub maximum: usize,
}

impl SourceReadPolicyError {
    /// Requested window size in bytes.
    #[must_use]
    pub const fn requested(self) -> usize {
        self.requested
    }

    /// Maximum permitted window size in bytes.
    #[must_use]
    pub const fn maximum(self) -> usize {
        self.maximum
    }
}

impl fmt::Display for SourceReadPolicyError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.requested == 0 {
            write!(
                formatter,
                "source read-ahead window must be positive (maximum {} bytes)",
                self.maximum
            )
        } else {
            write!(
                formatter,
                "source read-ahead window {} bytes exceeds maximum {} bytes",
                self.requested, self.maximum
            )
        }
    }
}

impl std::error::Error for SourceReadPolicyError {}

/// Counters and retained-capacity state for one source-read policy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceReadDiagnostics {
    /// Whether forward-start reads are currently enabled.
    pub enabled: bool,
    /// Configured forward window size.
    pub configured_window_bytes: usize,
    /// Allocated window capacity still retained by the adapter, including a
    /// buffer temporarily owned by an in-flight physical fill.
    pub retained_window_bytes: usize,
    /// Non-empty forward-mode calls admitted to the adapter.
    pub requests: u64,
    /// Requests served from the retained window.
    pub hits: u64,
    /// Requests that did not fit the retained window.
    pub misses: u64,
    /// Successful physical fill operations, including zero-byte returns.
    pub fills: u64,
    /// Logical bytes requested by non-empty forward-mode calls.
    pub requested_bytes: u64,
    /// Accepted physical bytes returned by successful fills.
    pub returned_bytes: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ReadMode {
    Forward,
    Exact,
}

struct ReadAheadState {
    bytes: Vec<u8>,
    retained_capacity: usize,
    start: u64,
    valid_len: usize,
    memory_reservation: Option<Arc<Reservation>>,
    requests: u64,
    hits: u64,
    misses: u64,
    fills: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

/// Archive-owned bounded read-ahead state.
pub(super) struct ArchiveReadAhead {
    mode: Mutex<ModeState>,
    mode_done: Condvar,
    transition_requested: AtomicBool,
    state: Mutex<ReadAheadState>,
    configured_window_bytes: usize,
    source_lineage: super::SourceLineage,
    source_version: litchi_core::SourceVersion,
    source_length: u64,
}

struct ModeState {
    mode: ReadMode,
    forward_owner: Option<std::thread::ThreadId>,
    transition_owner: Option<std::thread::ThreadId>,
}

/// A forward operation's admission token.  It owns no cache lock, so an
/// underlying provider may call back into diagnostics or publication safely.
struct ForwardReadLease<'owner> {
    owner: &'owner ArchiveReadAhead,
    thread: std::thread::ThreadId,
}

impl Drop for ForwardReadLease<'_> {
    fn drop(&mut self) {
        let mut mode = self.owner.lock_mode();
        if mode.forward_owner == Some(self.thread) {
            mode.forward_owner = None;
            self.owner.mode_done.notify_all();
        }
    }
}

/// Marks an exact-publication transition while it performs source/context
/// fences and drains the cache.  Keeping the owner identity in the mode
/// state lets callbacks from the provider receive a typed refusal before
/// they could recursively enter a transition or read operation.
struct TransitionLease<'owner> {
    owner: &'owner ArchiveReadAhead,
    thread: std::thread::ThreadId,
}

impl Drop for TransitionLease<'_> {
    fn drop(&mut self) {
        // A panic in a transition fence must still wait for an already
        // admitted forward operation before releasing its window.  This
        // keeps a fill lease from restoring a buffer after the exact mode has
        // otherwise been abandoned.
        let mut mode = self.owner.lock_mode();
        let mut same_thread_forward = false;
        while let Some(owner) = mode.forward_owner {
            if owner == self.thread {
                // This violates the normal admission invariant, but can be
                // observed after recovering a deliberately poisoned mode
                // value.  Never wait on a lease that belongs to this thread.
                same_thread_forward = true;
                break;
            }
            mode = self
                .owner
                .mode_done
                .wait(mode)
                .unwrap_or_else(|poisoned| poisoned.into_inner());
        }
        drop(mode);
        if !same_thread_forward {
            let (bytes, reservation) = self.owner.take_window();
            drop(reservation);
            drop(bytes);
        }

        let mut mode = self.owner.lock_mode();
        if mode.transition_owner == Some(self.thread) {
            mode.transition_owner = None;
            self.owner.mode_done.notify_all();
        }
    }
}

/// Temporary ownership of the retained buffer during a provider call.
///
/// Moving the vector out of the state means neither `read_at` nor `version`
/// callbacks can re-enter a held cache mutex.  The guard restores the buffer
/// and reservation if a provider panics while it owns them.
struct FillLease<'owner> {
    owner: &'owner ArchiveReadAhead,
    buffer: Option<Vec<u8>>,
    reservation: Option<Arc<Reservation>>,
    active: bool,
}

impl Drop for FillLease<'_> {
    fn drop(&mut self) {
        if !self.active {
            return;
        }
        let mut state = self.owner.lock_state();
        if state.bytes.is_empty() {
            if let Some(buffer) = self.buffer.take() {
                state.bytes = buffer;
                state.memory_reservation = self.reservation.take();
                state.retained_capacity = state.bytes.capacity();
            }
        }
        self.active = false;
    }
}

impl ArchiveReadAhead {
    /// Construct a window after the source snapshot has captured its length
    /// and version.  Memory is admitted before the vector allocation and the
    /// reservation remains live until the window is disabled or dropped.
    pub(super) fn new(snapshot: &super::SourceSnapshot, policy: SourceReadPolicy) -> Result<Self> {
        policy.validate().map_err(|error| {
            super::overlay_unavailable(format!("invalid source-read policy: {error}"))
        })?;
        snapshot.ensure_current()?;
        if let Some(context) = snapshot.context_ref() {
            context.check().map_err(super::map_execution_error)?;
        }

        let configured_window_bytes = policy.window_bytes();
        let retained_capacity = if configured_window_bytes == 0 {
            0
        } else {
            let configured = u64::try_from(configured_window_bytes)
                .map_err(|_| super::overlay_unavailable("source read-ahead window exceeds u64"))?;
            usize::try_from(snapshot.length.min(configured)).map_err(|_| {
                super::overlay_unavailable("source length does not fit a read-ahead window")
            })?
        };

        let memory_reservation = if retained_capacity == 0 {
            None
        } else if let Some(context) = snapshot.context_ref() {
            let bytes = u64::try_from(retained_capacity)
                .map_err(|_| super::overlay_unavailable("read-ahead memory exceeds u64"))?;
            Some(Arc::new(
                context
                    .reserve(Resource::Memory, bytes)
                    .map_err(super::map_execution_error)?,
            ))
        } else {
            None
        };

        let mut bytes = Vec::new();
        if retained_capacity != 0 {
            if let Err(source) = bytes.try_reserve_exact(retained_capacity) {
                return Err(OpcError::Allocation {
                    resource: "source read-ahead window",
                    source,
                });
            }
            if bytes.capacity() != retained_capacity {
                return Err(super::overlay_unavailable(
                    "source read-ahead allocation exceeded its memory reservation",
                ));
            }
            bytes.resize(retained_capacity, 0);
        }

        // The provider and execution context can both change while the
        // bounded allocation is being admitted.  Do not publish a window
        // whose source fence was already invalid by the time allocation
        // completed.  Dropping this constructor's locals releases the
        // reservation on every failure path.
        snapshot.ensure_current()?;
        if let Some(context) = snapshot.context_ref() {
            context.check().map_err(super::map_execution_error)?;
        }

        Ok(Self {
            mode: Mutex::new(ModeState {
                mode: if configured_window_bytes == 0 {
                    ReadMode::Exact
                } else {
                    ReadMode::Forward
                },
                forward_owner: None,
                transition_owner: None,
            }),
            mode_done: Condvar::new(),
            transition_requested: AtomicBool::new(false),
            state: Mutex::new(ReadAheadState {
                retained_capacity: bytes.capacity(),
                bytes,
                start: 0,
                valid_len: 0,
                memory_reservation,
                requests: 0,
                hits: 0,
                misses: 0,
                fills: 0,
                requested_bytes: 0,
                returned_bytes: 0,
            }),
            configured_window_bytes,
            source_lineage: snapshot.snapshot_lineage(),
            source_version: snapshot.version(),
            source_length: snapshot.length,
        })
    }

    /// Read one positional range through the selected policy.
    pub(super) fn read_at(
        &self,
        snapshot: &super::SourceSnapshot,
        offset: u64,
        output: &mut [u8],
    ) -> Result<usize> {
        if let Err(error) = self.check_snapshot_identity(snapshot) {
            if matches!(error, OpcError::SourceChanged { .. }) {
                self.invalidate_best_effort();
            }
            return Err(error);
        }
        if self.current_thread_is_busy() {
            return Err(Self::reentrant_refusal());
        }
        let Some(lease) = self.admit_forward(snapshot.context_ref())? else {
            // Exact mode intentionally has no adapter lease: the helper does
            // not hold read-ahead synchronization across the provider call,
            // so recursive exact-provider behavior cannot deadlock this
            // adapter.  Forward mode is admitted before its first source
            // fence below, which covers version-callback reentrancy.
            // Preserve the existing exact helper's zero-length provider
            // trace as well as its context/source error ordering.
            return super::read_source_at_with_context(
                snapshot,
                snapshot.context_ref(),
                offset,
                output,
                "archive",
            );
        };
        if output.is_empty() {
            let result = self
                .check_before_read(snapshot)
                .and_then(|()| self.check_after_read(snapshot));
            if let Err(error) = &result {
                if matches!(error, OpcError::SourceChanged { .. }) {
                    self.invalidate_best_effort();
                }
            }
            drop(lease);
            return result.map(|()| 0);
        }
        self.forward_read(snapshot, offset, output, lease)
    }

    /// Monotonically switch the shared archive adapter to exact reads.
    ///
    /// The admission flag closes the gate before this method waits for a
    /// forward operation already in progress.  The window and reservation are
    /// then removed under the short state lock and dropped after that lock is
    /// released, before any caller-owned exact ZIP traversal can begin.
    pub(super) fn disable_for_exact_publication(
        &self,
        snapshot: &super::SourceSnapshot,
    ) -> Result<()> {
        if let Err(error) = self.check_snapshot_identity(snapshot) {
            if matches!(error, OpcError::SourceChanged { .. }) {
                self.invalidate_best_effort();
            }
            return Err(error);
        }
        if self.current_thread_is_busy() {
            return Err(Self::reentrant_refusal());
        }

        let transition = self.begin_transition()?;
        // Set the permanent admission fence before invoking source/context
        // callbacks.  A callback on this same thread sees the transition
        // owner and receives a typed refusal instead of recursively waiting
        // on this transition.
        let check = self.check_before_read(snapshot);

        let mut mode = self.lock_mode();
        while let Some(owner) = mode.forward_owner {
            if owner == transition.thread {
                // `begin_transition` rejects this state, but keep the check
                // here so a recovered/poisoned mode value cannot deadlock the
                // caller.
                drop(mode);
                return Err(Self::reentrant_refusal());
            }
            mode = self
                .mode_done
                .wait(mode)
                .unwrap_or_else(|poisoned| poisoned.into_inner());
        }
        drop(mode);

        let (bytes, reservation) = self.take_window();
        drop(reservation);
        drop(bytes);

        let result = match check {
            Err(error) => {
                if matches!(error, OpcError::SourceChanged { .. }) {
                    self.invalidate_best_effort();
                }
                Err(error)
            },
            Ok(()) => self.check_after_read(snapshot).inspect_err(|_| {
                self.invalidate_best_effort();
            }),
        };
        drop(transition);
        result
    }

    /// Return a point-in-time diagnostic snapshot.
    pub(super) fn diagnostics(&self) -> Result<SourceReadDiagnostics> {
        let mode = self.lock_mode();
        let state = self.lock_state();
        Ok(SourceReadDiagnostics {
            enabled: mode.mode == ReadMode::Forward
                && !self.transition_requested.load(Ordering::Acquire),
            configured_window_bytes: self.configured_window_bytes,
            retained_window_bytes: state.retained_capacity,
            requests: state.requests,
            hits: state.hits,
            misses: state.misses,
            fills: state.fills,
            requested_bytes: state.requested_bytes,
            returned_bytes: state.returned_bytes,
        })
    }

    fn forward_read(
        &self,
        snapshot: &super::SourceSnapshot,
        offset: u64,
        output: &mut [u8],
        lease: ForwardReadLease<'_>,
    ) -> Result<usize> {
        // The admission lease serializes forward operations, but deliberately
        // owns no mutex.  Every source/version callback below therefore runs
        // outside cache synchronization.
        if let Err(error) = self.check_before_read(snapshot) {
            if matches!(error, OpcError::SourceChanged { .. }) {
                self.invalidate_best_effort();
            }
            return Err(error);
        }
        if self.transition_requested.load(Ordering::Acquire) {
            drop(lease);
            return super::read_source_at_with_context(
                snapshot,
                snapshot.context_ref(),
                offset,
                output,
                "archive",
            );
        }

        let requested = u64::try_from(output.len())
            .map_err(|_| super::overlay_unavailable("read-ahead request exceeds u64"))?;
        let mut state = self.lock_state();

        state.requests = checked_increment(state.requests, "read-ahead request count")?;
        state.requested_bytes = checked_add(
            state.requested_bytes,
            requested,
            "read-ahead requested bytes",
        )?;
        if state.valid_len != 0 {
            let end = checked_end(state.start, state.valid_len)?;
            if offset >= state.start && offset < end {
                state.hits = checked_increment(state.hits, "read-ahead hit count")?;
                let relative = usize::try_from(offset - state.start).map_err(|_| {
                    super::overlay_unavailable("read-ahead hit offset does not fit usize")
                })?;
                let available = state.valid_len.checked_sub(relative).ok_or_else(|| {
                    super::overlay_unavailable("read-ahead hit offset exceeds retained bytes")
                })?;
                let count = output.len().min(available);
                let end_index = relative.checked_add(count).ok_or_else(|| {
                    super::overlay_unavailable("read-ahead hit copy range overflows usize")
                })?;
                output[..count].copy_from_slice(&state.bytes[relative..end_index]);
                drop(state);
                return self.finish_forward_read(snapshot, count, lease);
            }
        }

        state.misses = checked_increment(state.misses, "read-ahead miss count")?;
        let available = self.source_length.saturating_sub(offset);
        let configured = u64::try_from(self.configured_window_bytes)
            .map_err(|_| super::overlay_unavailable("read-ahead window exceeds u64"))?;
        let fill_len = usize::try_from(available.min(configured))
            .map_err(|_| super::overlay_unavailable("read-ahead fill length does not fit usize"))?;
        if fill_len == 0 {
            state.valid_len = 0;
            state.start = 0;
            drop(state);
            return self.finish_forward_read(snapshot, 0, lease);
        }

        // Invalidate before the provider call.  A transport error, count
        // violation, cancellation, or source change can never leave the
        // preceding window visible.
        state.valid_len = 0;
        state.start = 0;
        let buffer = std::mem::take(&mut state.bytes);
        let reservation = state.memory_reservation.take();
        drop(state);
        let mut fill = FillLease {
            owner: self,
            buffer: Some(buffer),
            reservation,
            active: true,
        };

        let read = {
            let buffer = fill
                .buffer
                .as_mut()
                .ok_or_else(|| super::overlay_unavailable("read-ahead fill buffer is missing"))?;
            if buffer.len() < fill_len {
                return Err(super::overlay_unavailable(
                    "read-ahead fill exceeds its retained buffer",
                ));
            }
            super::read_source_at_with_context(
                snapshot,
                snapshot.context_ref(),
                offset,
                &mut buffer[..fill_len],
                "read-ahead",
            )?
        };

        // The shared helper fences monitored sources, but ordinary
        // source-backed opens leave that monitor disabled.  These explicit
        // checks make a positive window safe for unmonitored reads as well.
        self.check_after_read(snapshot)?;
        checked_end(offset, read)?;

        let count = output.len().min(read);
        if count != 0 {
            let buffer = fill
                .buffer
                .as_ref()
                .ok_or_else(|| super::overlay_unavailable("read-ahead fill buffer is missing"))?;
            output[..count].copy_from_slice(&buffer[..count]);
        }

        let mut state = self.lock_state();
        state.fills = checked_increment(state.fills, "read-ahead fill count")?;
        state.returned_bytes = checked_add(
            state.returned_bytes,
            u64::try_from(read)
                .map_err(|_| super::overlay_unavailable("read-ahead return exceeds u64"))?,
            "read-ahead returned bytes",
        )?;
        state.valid_len = 0;
        state.start = 0;
        if state.bytes.is_empty() {
            if let Some(buffer) = fill.buffer.take() {
                state.bytes = buffer;
                state.memory_reservation = fill.reservation.take();
                state.retained_capacity = state.bytes.capacity();
                state.start = offset;
                state.valid_len = read;
            }
        }
        drop(state);
        fill.active = false;
        self.finish_forward_read(snapshot, count, lease)
    }

    fn finish_forward_read(
        &self,
        snapshot: &super::SourceSnapshot,
        count: usize,
        lease: ForwardReadLease<'_>,
    ) -> Result<usize> {
        match self.check_after_read(snapshot) {
            Ok(()) => {
                drop(lease);
                Ok(count)
            },
            Err(error) => {
                self.invalidate_best_effort();
                drop(lease);
                Err(error)
            },
        }
    }

    fn check_snapshot_identity(&self, snapshot: &super::SourceSnapshot) -> Result<()> {
        if snapshot.lineage() != &self.source_lineage {
            return Err(super::overlay_unavailable(
                "source read-ahead snapshot lineage does not match its owner",
            ));
        }
        let version = snapshot.version();
        if version != self.source_version || snapshot.length != self.source_length {
            return Err(OpcError::SourceChanged {
                expected: self.source_version,
                actual: version,
            });
        }
        Ok(())
    }

    fn check_before_read(&self, snapshot: &super::SourceSnapshot) -> Result<()> {
        snapshot.ensure_current()?;
        if let Some(context) = snapshot.context_ref() {
            context.check().map_err(super::map_execution_error)?;
        }
        Ok(())
    }

    fn check_after_read(&self, snapshot: &super::SourceSnapshot) -> Result<()> {
        snapshot.ensure_current()?;
        if let Some(context) = snapshot.context_ref() {
            context.check().map_err(super::map_execution_error)?;
        }
        Ok(())
    }

    fn invalidate_best_effort(&self) {
        let mut state = self.lock_state();
        state.valid_len = 0;
        state.start = 0;
    }

    fn lock_mode(&self) -> MutexGuard<'_, ModeState> {
        self.mode
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn lock_state(&self) -> MutexGuard<'_, ReadAheadState> {
        self.state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn current_thread_is_busy(&self) -> bool {
        let current = std::thread::current().id();
        let mode = self.lock_mode();
        mode.forward_owner == Some(current) || mode.transition_owner == Some(current)
    }

    fn reentrant_refusal() -> OpcError {
        super::overlay_unavailable(
            "reentrant source read-ahead operation refused while its provider callback is active",
        )
    }

    fn admit_forward(
        &self,
        context: Option<&litchi_core::ExecutionContext>,
    ) -> Result<Option<ForwardReadLease<'_>>> {
        let current = std::thread::current().id();
        if self.transition_requested.load(Ordering::Acquire) {
            return Ok(None);
        }
        let mut mode = self.lock_mode();
        loop {
            if mode.forward_owner == Some(current) || mode.transition_owner == Some(current) {
                return Err(Self::reentrant_refusal());
            }
            if mode.mode == ReadMode::Exact || self.transition_requested.load(Ordering::Acquire) {
                return Ok(None);
            }
            if mode.transition_owner.is_some() {
                if let Some(context) = context {
                    // Managed semantic waiters poll cancellation while the
                    // transition owner completes its short state drain.
                    // Publication itself intentionally uses an untimed wait
                    // so it cannot abandon an admitted provider lease.
                    context.check().map_err(super::map_execution_error)?;
                    mode = self
                        .mode_done
                        .wait_timeout(mode, Duration::from_millis(10))
                        .unwrap_or_else(|poisoned| poisoned.into_inner())
                        .0;
                } else {
                    mode = self
                        .mode_done
                        .wait(mode)
                        .unwrap_or_else(|poisoned| poisoned.into_inner());
                }
                continue;
            }
            if mode.forward_owner.is_none() {
                mode.forward_owner = Some(current);
                drop(mode);
                return Ok(Some(ForwardReadLease {
                    owner: self,
                    thread: current,
                }));
            }
            // A forward owner may be inside an arbitrary provider callback;
            // managed readers must remain cancellable while they wait for
            // that owner to release its lease.
            if let Some(context) = context {
                context.check().map_err(super::map_execution_error)?;
                mode = self
                    .mode_done
                    .wait_timeout(mode, Duration::from_millis(10))
                    .unwrap_or_else(|poisoned| poisoned.into_inner())
                    .0;
            } else {
                mode = self
                    .mode_done
                    .wait(mode)
                    .unwrap_or_else(|poisoned| poisoned.into_inner());
            }
        }
    }

    fn begin_transition(&self) -> Result<TransitionLease<'_>> {
        let current = std::thread::current().id();
        let mut mode = self.lock_mode();
        loop {
            if mode.forward_owner == Some(current) || mode.transition_owner == Some(current) {
                return Err(Self::reentrant_refusal());
            }
            if mode.transition_owner.is_none() {
                break;
            }
            mode = self
                .mode_done
                .wait(mode)
                .unwrap_or_else(|poisoned| poisoned.into_inner());
        }
        mode.transition_owner = Some(current);
        mode.mode = ReadMode::Exact;
        // This flag is monotonic.  It closes admission before the mode guard
        // is released, so a steady stream of callers cannot starve the
        // transition while an older forward operation drains.
        self.transition_requested.store(true, Ordering::SeqCst);
        drop(mode);
        Ok(TransitionLease {
            owner: self,
            thread: current,
        })
    }

    fn take_window(&self) -> (Vec<u8>, Option<Arc<Reservation>>) {
        let mut state = self.lock_state();
        state.valid_len = 0;
        state.start = 0;
        state.retained_capacity = 0;
        (
            std::mem::take(&mut state.bytes),
            state.memory_reservation.take(),
        )
    }
}

/// Hard upper bound for one coalesced structural run read.
///
/// A run is fetched by one positional read, so this is the same ceiling the
/// forward window already puts on a single speculative request.
pub(super) const MAX_STRUCTURAL_PREFETCH_RUN_BYTES: u64 = MAX_SOURCE_READ_AHEAD_BYTES as u64;

/// Hard upper bound for the bytes one structural prefetch retains at once.
///
/// Unlike the forward window, which holds one range, the prefetch holds every
/// run it admitted until the catalog has been read, because the relationship
/// walk is a LIFO traversal that revisits runs in no particular order. The
/// ceiling is measured: across the 533 ZIP containers under `test-data` the
/// largest run set one package retains is 14,755 bytes and the longest single
/// run is 9,298 bytes, so neither clamp binds on that corpus and 256 KiB still
/// bounds the transient allocation at a fraction of one member payload.
pub(super) const MAX_STRUCTURAL_PREFETCH_BYTES: usize = 256 * 1024;

/// The fewest members a run must hold before one read over it is admitted.
///
/// A one-member run trades one request for one request and can only lose: the
/// member might never be read. Two is the smallest size at which coalescing
/// can win.
pub(super) const MIN_STRUCTURAL_PREFETCH_RUN_MEMBERS: usize = 2;

/// One physically contiguous run of structural members.
///
/// `start` is the first member's local-header offset and `end` is the furthest
/// offset any member of the run would read to, so the half-open range is
/// exactly the union of those members' own first-read spans.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct StructuralRun {
    start: u64,
    end: u64,
    members: usize,
}

impl StructuralRun {
    /// The run's first byte.
    pub(super) const fn start(self) -> u64 {
        self.start
    }

    /// The run's length in bytes.
    pub(super) const fn length(self) -> u64 {
        self.end.saturating_sub(self.start)
    }
}

/// Group per-member first-read spans into the runs one read each can deliver.
///
/// `spans` are `(offset, length)` pairs in increasing local-header order, one
/// per member the caller intends to read. Two neighbours join a run only when
/// the earlier member's own span already reaches the later member's local
/// header: that is what makes their union a range with no hole, so every byte
/// of the run read lies inside some member's own first-read span and no byte
/// outside those spans is fetched. A member whose span stops short — because a
/// record the caller did not list sits between them, because the archive has a
/// gap there, or because the member's declared size clamps its window — starts
/// a new run instead.
///
/// Runs shorter than [`MIN_STRUCTURAL_PREFETCH_RUN_MEMBERS`] and runs longer
/// than [`MAX_STRUCTURAL_PREFETCH_RUN_BYTES`] are dropped, and admission stops
/// once [`MAX_STRUCTURAL_PREFETCH_BYTES`] would be exceeded. Dropping a run is
/// never a refusal: its members keep the grammar they have today.
pub(super) fn structural_runs<S>(spans: S) -> Result<Vec<StructuralRun>>
where
    S: IntoIterator<Item = (u64, u64)>,
{
    // The vector is grown fallibly and only when a run is actually admitted,
    // so a package with no coalescable run allocates nothing at all. The
    // caller streams its member list in, so nothing is buffered either.
    let mut runs: Vec<StructuralRun> = Vec::new();
    let mut retained = 0_usize;
    let mut current: Option<StructuralRun> = None;
    let mut previous_offset: Option<u64> = None;
    for (offset, length) in spans {
        let Some(end) = offset.checked_add(length) else {
            // A span that does not fit the address space is not a fact this
            // layer can act on. Abandon the run and keep the ordinary grammar.
            current = None;
            previous_offset = None;
            continue;
        };
        if previous_offset.is_some_and(|previous| offset <= previous) {
            // Not in increasing local-header order. Refuse to guess.
            return Ok(Vec::new());
        }
        previous_offset = Some(offset);
        current = match current {
            Some(run) if run.end >= offset => Some(StructuralRun {
                start: run.start,
                end: run.end.max(end),
                members: run.members.saturating_add(1),
            }),
            Some(run) => {
                push_admitted_run(&mut runs, &mut retained, run)?;
                Some(StructuralRun {
                    start: offset,
                    end,
                    members: 1,
                })
            },
            None => Some(StructuralRun {
                start: offset,
                end,
                members: 1,
            }),
        };
    }
    if let Some(run) = current {
        push_admitted_run(&mut runs, &mut retained, run)?;
    }
    Ok(runs)
}

/// Admit one run, or drop it silently.
///
/// Dropping is never a refusal: a run the ceilings exclude, or one too short
/// to pay for itself, simply leaves its members on the grammar they have
/// today.
fn push_admitted_run(
    runs: &mut Vec<StructuralRun>,
    retained: &mut usize,
    run: StructuralRun,
) -> Result<()> {
    let Ok(length) = usize::try_from(run.length()) else {
        return Ok(());
    };
    if run.members < MIN_STRUCTURAL_PREFETCH_RUN_MEMBERS
        || run.length() > MAX_STRUCTURAL_PREFETCH_RUN_BYTES
        || length == 0
    {
        return Ok(());
    }
    let Some(total) = retained.checked_add(length) else {
        return Ok(());
    };
    if total > MAX_STRUCTURAL_PREFETCH_BYTES {
        return Ok(());
    }
    runs.try_reserve(1).map_err(|source| OpcError::Allocation {
        resource: "source-backed OPC structural prefetch runs",
        source,
    })?;
    runs.push(run);
    *retained = total;
    Ok(())
}

/// One run's bytes, as one read delivered them.
#[derive(Debug)]
struct PrefetchedRun {
    base: u64,
    bytes: Vec<u8>,
}

/// A bounded, open-scoped prefetch of the structural members' local records.
///
/// This is the same kind of object as [`ArchiveReadAhead`] — a buffer of source
/// bytes below the ZIP substrate — and it is deliberately not the same
/// mechanism. The forward window guesses forward from wherever the traversal
/// happens to be; this one is aimed at ranges the caller already knows it will
/// read, and it never fetches anything else.
///
/// It is a cache of source bytes and nothing else. No value is trusted because
/// it came from here: every byte is parsed, bounded and verified by exactly the
/// code that parses, bounds and verifies it when each read goes to the source
/// separately, in the same order, with the same limits charged.
pub(super) struct StructuralPrefetch {
    /// Whether any run is retained. Checked before the lock so a package that
    /// primed nothing, or whose prefetch has been released, pays one relaxed
    /// load per read and no synchronization.
    active: AtomicBool,
    runs: Mutex<Vec<PrefetchedRun>>,
}

impl StructuralPrefetch {
    pub(super) fn new() -> Self {
        Self {
            active: AtomicBool::new(false),
            runs: Mutex::new(Vec::new()),
        }
    }

    /// Fetch every admitted run, one positional read each.
    ///
    /// This is best effort in the precise sense that a failure leaves the
    /// package exactly as it would have been without it: nothing is retained,
    /// every member read goes to the source as it does today, and the failure
    /// is observed again by whichever read actually needs those bytes. That is
    /// invariant 5 of change 0577's frozen design, and it is right here for the
    /// reason change 0611 gave for refusing it one member wide — this read
    /// spans several members and belongs to none of them, so abandoning it
    /// cannot give any member read a second refusal.
    ///
    /// A short read is retained as far as it reached. The members it covers in
    /// full are served from it; every other read reaches the source at exactly
    /// the offset and length it reaches today.
    pub(super) fn prime(&self, snapshot: &super::SourceSnapshot, runs: &[StructuralRun]) {
        if runs.is_empty() {
            return;
        }
        if snapshot.context_ref().is_some() {
            // A managed open reserves and commits `Resource::InputBytes` per
            // read and observes cancellation around it. A speculative read
            // spanning several members would move where those observations
            // land and how soon a finite input budget is exhausted, which is
            // not a trade this change makes on a caller's behalf. Managed
            // opens keep the exact grammar; the caller already gates on this,
            // and the check is repeated here so the property is local.
            return;
        }
        let mut retained: Vec<PrefetchedRun> = Vec::new();
        if retained.try_reserve_exact(runs.len()).is_err() {
            return;
        }
        for run in runs {
            let Ok(length) = usize::try_from(run.length()) else {
                return;
            };
            let mut bytes = Vec::new();
            if bytes.try_reserve_exact(length).is_err() {
                return;
            }
            bytes.resize(length, 0);
            let Ok(read) = super::read_source_at_with_context(
                snapshot,
                snapshot.context_ref(),
                run.start(),
                &mut bytes,
                "structural prefetch",
            ) else {
                return;
            };
            if read == 0 {
                return;
            }
            bytes.truncate(read);
            retained.push(PrefetchedRun {
                base: run.start(),
                bytes,
            });
        }
        let mut held = self.lock_runs();
        *held = retained;
        drop(held);
        self.active.store(true, Ordering::Release);
    }

    /// Answer `output` when one retained run holds every byte of it.
    ///
    /// Only a wholly covered, non-empty read is answered. A partly covered one
    /// would be served in different-sized pieces than the source would have
    /// delivered, which change 0611 measured moving which of two refusals a
    /// malformed member reaches; and a zero-length read has nothing to answer
    /// with, so it keeps the provider trace it has today.
    pub(super) fn serve(&self, offset: u64, output: &mut [u8]) -> Option<usize> {
        if output.is_empty() || !self.active.load(Ordering::Acquire) {
            return None;
        }
        let end = offset.checked_add(u64::try_from(output.len()).ok()?)?;
        let runs = self.lock_runs();
        let run = runs.iter().find(|run| {
            let run_end = run
                .base
                .saturating_add(u64::try_from(run.bytes.len()).unwrap_or(u64::MAX));
            run.base <= offset && end <= run_end
        })?;
        let start = usize::try_from(offset - run.base).ok()?;
        let slice = run.bytes.get(start..start.checked_add(output.len())?)?;
        output.copy_from_slice(slice);
        Some(output.len())
    }

    /// Release every retained run.
    ///
    /// The prefetch exists for the catalog phase only. Once that phase is over
    /// the package keeps no speculative bytes, and every later read behaves
    /// exactly as it does without this mechanism.
    pub(super) fn clear(&self) {
        self.active.store(false, Ordering::Release);
        let taken = std::mem::take(&mut *self.lock_runs());
        drop(taken);
    }

    fn lock_runs(&self) -> MutexGuard<'_, Vec<PrefetchedRun>> {
        self.runs
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }
}

fn checked_end(start: u64, length: usize) -> Result<u64> {
    start
        .checked_add(
            u64::try_from(length)
                .map_err(|_| super::overlay_unavailable("read-ahead length exceeds u64"))?,
        )
        .ok_or_else(|| super::overlay_unavailable("read-ahead range end overflows u64"))
}

fn checked_increment(value: u64, what: &'static str) -> Result<u64> {
    value
        .checked_add(1)
        .ok_or_else(|| super::overlay_unavailable(what))
}

fn checked_add(left: u64, right: u64, what: &'static str) -> Result<u64> {
    left.checked_add(right)
        .ok_or_else(|| super::overlay_unavailable(what))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
        SourceVersion,
    };
    use std::io;
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Mutex as StdMutex;
    use std::sync::atomic::{AtomicU64, Ordering};
    use std::sync::mpsc::{self, Receiver, Sender};
    use std::thread;

    fn snapshot(
        source: Arc<dyn ReadAt>,
        context: Option<ExecutionContext>,
    ) -> super::super::SourceSnapshot {
        let version = source.version().unwrap();
        let length = source.len().unwrap();
        super::super::SourceSnapshot {
            source,
            version,
            length,
            monitor_reads: Arc::new(crate::source_backed::MonitoredReadDepth::new(0)),
            lineage: super::super::SourceLineage(Arc::new(())),
            context,
            input_reservation_failures: None,
            output_reservation_failures: None,
        }
    }

    fn context_with_work(
        memory: u64,
        input: u64,
        work: u64,
    ) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "read-ahead-test",
            Limits::new(memory, input, u64::MAX, u64::MAX, u64::MAX, work),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let execution = ExecutionContext::new(budget.clone(), cancellation, limits);
        (budget, cancellation_source, execution)
    }

    fn context(memory: u64, input: u64) -> (Budget, CancellationSource, ExecutionContext) {
        context_with_work(memory, input, u64::MAX)
    }

    #[derive(Debug)]
    struct CountingSource {
        bytes: Vec<u8>,
        calls: AtomicU64,
        version: AtomicU64,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                calls: AtomicU64::new(0),
                version: AtomicU64::new(0),
            }
        }

        fn bump_version(&self) {
            self.version.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(91, self.version.load(Ordering::SeqCst)))
        }
    }

    #[derive(Debug)]
    struct CappedSource {
        bytes: Vec<u8>,
        cap: usize,
        calls: AtomicU64,
    }

    impl CappedSource {
        fn new(bytes: Vec<u8>, cap: usize) -> Self {
            Self {
                bytes,
                cap,
                calls: AtomicU64::new(0),
            }
        }
    }

    impl ReadAt for CappedSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture too large"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.cap).min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(92, 0))
        }
    }

    #[derive(Debug)]
    struct InterruptedOnceSource {
        bytes: Vec<u8>,
        calls: AtomicU64,
        interrupted: AtomicU64,
    }

    impl InterruptedOnceSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                calls: AtomicU64::new(0),
                interrupted: AtomicU64::new(0),
            }
        }
    }

    impl ReadAt for InterruptedOnceSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            if self
                .interrupted
                .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |value| {
                    (value == 0).then_some(1)
                })
                .is_ok()
            {
                return Err(io::Error::new(io::ErrorKind::Interrupted, "try again"));
            }
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(93, 0))
        }
    }

    #[derive(Debug)]
    struct CancelAfterReadSource {
        bytes: Vec<u8>,
        cancellation: CancellationSource,
        calls: AtomicU64,
    }

    impl ReadAt for CancelAfterReadSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.cancellation.cancel();
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(94, 0))
        }
    }

    #[derive(Debug)]
    struct ZeroSource {
        length: u64,
        calls: AtomicU64,
    }

    impl ReadAt for ZeroSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.length)
        }

        fn read_at(&self, _offset: u64, _output: &mut [u8]) -> io::Result<usize> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            Ok(0)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(95, 0))
        }
    }

    #[derive(Debug)]
    struct GatedSource {
        bytes: Vec<u8>,
        calls: AtomicU64,
        entered: StdMutex<Option<Sender<()>>>,
        release: StdMutex<Receiver<()>>,
    }

    impl GatedSource {
        fn new(bytes: Vec<u8>) -> (Self, Receiver<()>, Sender<()>) {
            let (entered_tx, entered_rx) = mpsc::channel();
            let (release_tx, release_rx) = mpsc::channel();
            (
                Self {
                    bytes,
                    calls: AtomicU64::new(0),
                    entered: StdMutex::new(Some(entered_tx)),
                    release: StdMutex::new(release_rx),
                },
                entered_rx,
                release_tx,
            )
        }
    }

    impl ReadAt for GatedSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let call = self.calls.fetch_add(1, Ordering::SeqCst);
            if call == 0 {
                if let Some(sender) = self.entered.lock().unwrap().take() {
                    sender.send(()).unwrap();
                }
                self.release.lock().unwrap().recv().unwrap();
            }
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(96, 0))
        }
    }

    #[derive(Debug)]
    struct HugeSource;

    impl ReadAt for HugeSource {
        fn len(&self) -> io::Result<u64> {
            Ok(u64::MAX)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            if offset == u64::MAX - 1 && !output.is_empty() {
                output[0] = 7;
                return Ok(1);
            }
            Ok(0)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(97, 0))
        }
    }

    struct CallbackSource {
        bytes: Vec<u8>,
        read_callback: StdMutex<Option<Arc<dyn Fn() + Send + Sync>>>,
        version_callback: StdMutex<Option<Arc<dyn Fn() + Send + Sync>>>,
    }

    impl CallbackSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                read_callback: StdMutex::new(None),
                version_callback: StdMutex::new(None),
            }
        }

        fn set_read_callback(&self, callback: Arc<dyn Fn() + Send + Sync>) {
            *self.read_callback.lock().unwrap() = Some(callback);
        }

        fn set_version_callback(&self, callback: Arc<dyn Fn() + Send + Sync>) {
            *self.version_callback.lock().unwrap() = Some(callback);
        }
    }

    impl ReadAt for CallbackSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let callback = self.read_callback.lock().unwrap().as_ref().cloned();
            if let Some(callback) = callback {
                callback();
            }
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            let callback = self.version_callback.lock().unwrap().as_ref().cloned();
            if let Some(callback) = callback {
                callback();
            }
            Ok(SourceVersion::new(98, 0))
        }
    }

    #[derive(Debug)]
    struct PanicSource;

    impl ReadAt for PanicSource {
        fn len(&self) -> io::Result<u64> {
            Ok(8)
        }

        fn read_at(&self, _offset: u64, _output: &mut [u8]) -> io::Result<usize> {
            panic!("test source panic during read-ahead fill");
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(99, 0))
        }
    }

    #[test]
    fn policy_is_exact_by_default_and_rejects_unbounded_values() {
        assert_eq!(SourceReadPolicy::default(), SourceReadPolicy::exact());
        assert_eq!(SourceReadPolicy::exact().window_bytes(), 0);
        assert_eq!(
            SourceReadPolicy::forward_start(1).unwrap().window_bytes(),
            1
        );
        assert_eq!(
            SourceReadPolicy::forward_start(MAX_SOURCE_READ_AHEAD_BYTES)
                .unwrap()
                .window_bytes(),
            MAX_SOURCE_READ_AHEAD_BYTES
        );
        let zero = SourceReadPolicy::forward_start(0).unwrap_err();
        assert_eq!(zero.requested(), 0);
        assert_eq!(zero.maximum(), MAX_SOURCE_READ_AHEAD_BYTES);
        let over = SourceReadPolicy::forward_start(MAX_SOURCE_READ_AHEAD_BYTES + 1).unwrap_err();
        assert!(over.to_string().contains("exceeds maximum"));

        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(vec![1, 2, 3]));
        let snap = snapshot(source, None);
        let exact = ArchiveReadAhead::new(&snap, SourceReadPolicy::exact()).unwrap();
        assert_eq!(exact.diagnostics().unwrap().configured_window_bytes, 0);
        assert!(!exact.diagnostics().unwrap().enabled);
    }

    #[test]
    fn forward_fills_once_and_returns_a_positive_crossing_prefix() {
        let source = Arc::new(CountingSource::new((0..32).collect()));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, None);
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();

        let mut first = [0_u8; 3];
        assert_eq!(read_ahead.read_at(&snap, 4, &mut first).unwrap(), 3);
        assert_eq!(first, [4, 5, 6]);
        let mut second = [0_u8; 8];
        assert_eq!(read_ahead.read_at(&snap, 6, &mut second).unwrap(), 6);
        assert_eq!(&second[..6], &[6, 7, 8, 9, 10, 11]);
        assert_eq!(source.calls.load(Ordering::SeqCst), 1);
        let diagnostics = read_ahead.diagnostics().unwrap();
        assert_eq!(diagnostics.requests, 2);
        assert_eq!(diagnostics.hits, 1);
        assert_eq!(diagnostics.misses, 1);
        assert_eq!(diagnostics.fills, 1);
        assert_eq!(diagnostics.requested_bytes, 11);
        assert_eq!(diagnostics.returned_bytes, 8);
    }

    #[test]
    fn managed_fill_charges_accepted_bytes_and_releases_memory_on_transition() {
        let (budget, _cancellation, execution) = context(16, 1024);
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new((0..64).collect()));
        let snap = snapshot(source, Some(execution));
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(16).unwrap()).unwrap();
        assert_eq!(budget.used(Resource::Memory), 16);
        let mut output = [0_u8; 4];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut output).unwrap(), 4);
        assert_eq!(budget.used(Resource::InputBytes), 16);
        assert_eq!(budget.used(Resource::Memory), 16);
        read_ahead.disable_for_exact_publication(&snap).unwrap();
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(read_ahead.diagnostics().unwrap().retained_window_bytes, 0);
        assert!(!read_ahead.diagnostics().unwrap().enabled);
    }

    #[test]
    fn managed_input_budget_shrinks_the_physical_fill_and_stops_at_zero() {
        let (budget, _cancellation, execution) = context(8, 2);
        let source = Arc::new(CountingSource::new((0..32).collect()));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, Some(execution));
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();

        let mut first = [0_u8; 1];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut first).unwrap(), 1);
        assert_eq!(first, [0]);
        assert_eq!(budget.used(Resource::InputBytes), 2);

        let calls = source.calls.load(Ordering::SeqCst);
        let mut second = [0_u8; 1];
        assert!(matches!(
            read_ahead.read_at(&snap, 16, &mut second),
            Err(OpcError::Execution(
                litchi_core::ExecutionError::ResourceLimit(_)
            ))
        ));
        assert_eq!(source.calls.load(Ordering::SeqCst), calls);
        let diagnostics = read_ahead.diagnostics().unwrap();
        assert_eq!(diagnostics.fills, 1);
        assert_eq!(diagnostics.returned_bytes, 2);
        assert_eq!(diagnostics.requested_bytes, 2);
    }

    #[test]
    fn short_and_zero_physical_returns_are_bounded_and_not_reused_past_end() {
        let source = Arc::new(CappedSource::new((0..16).collect(), 2));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, None);
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();
        let mut output = [0xa5_u8; 4];
        assert_eq!(read_ahead.read_at(&snap, 3, &mut output).unwrap(), 2);
        assert_eq!(&output[..2], &[3, 4]);
        assert_eq!(output[2], 0xa5);
        assert_eq!(read_ahead.read_at(&snap, 4, &mut output).unwrap(), 1);
        assert_eq!(output[0], 4);

        let zero = Arc::new(ZeroSource {
            length: 8,
            calls: AtomicU64::new(0),
        });
        let zero_read_at: Arc<dyn ReadAt> = zero.clone();
        let zero_snap = snapshot(zero_read_at, None);
        let zero_window =
            ArchiveReadAhead::new(&zero_snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();
        assert_eq!(
            zero_window.read_at(&zero_snap, 0, &mut [0_u8; 1]).unwrap(),
            0
        );
        let diagnostics = zero_window.diagnostics().unwrap();
        assert_eq!(diagnostics.fills, 1);
        assert_eq!(diagnostics.returned_bytes, 0);
        assert_eq!(diagnostics.retained_window_bytes, 8);
    }

    #[test]
    fn interrupted_read_consumes_work_and_cancellation_drops_publication() {
        let (budget, _cancellation, execution) = context_with_work(16, 64, 4);
        let source = Arc::new(InterruptedOnceSource::new((0..32).collect()));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, Some(execution));
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();
        let mut output = [0_u8; 2];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut output).unwrap(), 2);
        assert_eq!(output, [0, 1]);
        assert_eq!(source.calls.load(Ordering::SeqCst), 2);
        assert_eq!(budget.used(Resource::Work), 1);

        let (cancel_budget, cancellation, cancel_execution) = context(8, 64);
        let cancelling_source = Arc::new(CancelAfterReadSource {
            bytes: (0..16).collect(),
            cancellation,
            calls: AtomicU64::new(0),
        });
        let cancelling_read_at: Arc<dyn ReadAt> = cancelling_source.clone();
        let cancelling_snap = snapshot(cancelling_read_at, Some(cancel_execution));
        let cancelling_window = ArchiveReadAhead::new(
            &cancelling_snap,
            SourceReadPolicy::forward_start(8).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            cancelling_window.read_at(&cancelling_snap, 0, &mut [0_u8; 2]),
            Err(OpcError::Cancelled)
        ));
        assert_eq!(cancel_budget.used(Resource::InputBytes), 8);
        assert_eq!(cancelling_window.diagnostics().unwrap().fills, 0);
        assert_eq!(
            cancelling_window
                .diagnostics()
                .unwrap()
                .retained_window_bytes,
            8
        );
    }

    #[test]
    fn memory_admission_failure_releases_reservation_and_empty_source_has_no_buffer() {
        let (budget, _cancellation, execution) = context(7, 64);
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(vec![0; 16]));
        let snap = snapshot(source, Some(execution));
        assert!(matches!(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()),
            Err(OpcError::Execution(
                litchi_core::ExecutionError::ResourceLimit(_)
            ))
        ));
        assert_eq!(budget.used(Resource::Memory), 0);

        let (empty_budget, _cancellation, empty_execution) = context(0, 64);
        let empty_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(Vec::new()));
        let empty_snap = snapshot(empty_source, Some(empty_execution));
        let empty_window =
            ArchiveReadAhead::new(&empty_snap, SourceReadPolicy::forward_start(8).unwrap())
                .unwrap();
        let diagnostics = empty_window.diagnostics().unwrap();
        assert_eq!(diagnostics.retained_window_bytes, 0);
        assert_eq!(empty_budget.used(Resource::Memory), 0);
    }

    #[test]
    fn near_u64_maximum_ranges_use_checked_endpoints_and_preserve_sentinels() {
        let source: Arc<dyn ReadAt> = Arc::new(HugeSource);
        let snap = snapshot(source, None);
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();
        let mut output = [0xa5_u8; 2];
        assert_eq!(
            read_ahead
                .read_at(&snap, u64::MAX - 1, &mut output)
                .unwrap(),
            1
        );
        assert_eq!(output, [7, 0xa5]);
        assert_eq!(
            read_ahead
                .read_at(&snap, u64::MAX - 1, &mut output)
                .unwrap(),
            1
        );
        assert_eq!(read_ahead.read_at(&snap, u64::MAX, &mut output).unwrap(), 0);
    }

    #[test]
    fn exact_transition_waits_for_a_gated_fill_and_future_reads_are_exact() {
        let (source, entered, release) = GatedSource::new((0..64).collect());
        let source = Arc::new(source);
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = Arc::new(snapshot(source_read_at, None));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );

        let reader = {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            thread::spawn(move || {
                let mut output = [0_u8; 2];
                read_ahead.read_at(&snap, 0, &mut output).unwrap();
                output
            })
        };
        entered.recv().unwrap();

        let (done_tx, done_rx) = mpsc::channel();
        let transition = {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            thread::spawn(move || {
                read_ahead.disable_for_exact_publication(&snap).unwrap();
                done_tx.send(()).unwrap();
            })
        };
        assert!(matches!(
            done_rx.recv_timeout(Duration::from_millis(20)),
            Err(mpsc::RecvTimeoutError::Timeout)
        ));

        // Once the monotonic admission flag is visible, readers racing with
        // the blocked transition must take the exact path and leave the
        // forward counter unchanged.  They must not extend the transition's
        // wait set.
        for _ in 0..10_000 {
            if read_ahead.transition_requested.load(Ordering::Acquire) {
                break;
            }
            thread::yield_now();
        }
        assert!(read_ahead.transition_requested.load(Ordering::Acquire));
        let mut racing_readers = Vec::new();
        for _ in 0..4 {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            racing_readers.push(thread::spawn(move || {
                let mut output = [0_u8; 2];
                read_ahead.read_at(&snap, 16, &mut output).unwrap();
                output
            }));
        }
        for reader in racing_readers {
            assert_eq!(reader.join().unwrap(), [16, 17]);
        }
        release.send(()).unwrap();
        assert_eq!(reader.join().unwrap(), [0, 1]);
        transition.join().unwrap();

        let calls_after_fill = source.calls.load(Ordering::SeqCst);
        let mut output = [0_u8; 2];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut output).unwrap(), 2);
        assert_eq!(output, [0, 1]);
        assert_eq!(source.calls.load(Ordering::SeqCst), calls_after_fill + 1);
        let diagnostics = read_ahead.diagnostics().unwrap();
        assert!(!diagnostics.enabled);
        assert_eq!(diagnostics.retained_window_bytes, 0);
        assert_eq!(diagnostics.requests, 1);
    }

    #[test]
    fn managed_waiter_polls_cancellation_while_fill_remains_blocked() {
        let (_budget, cancellation, execution) = context(8, 64);
        let (source, entered, release) = GatedSource::new((0..64).collect());
        let source = Arc::new(source);
        let source_read_at: Arc<dyn ReadAt> = source;
        let snap = Arc::new(snapshot(source_read_at, Some(execution)));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );

        let first = {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            thread::spawn(move || {
                let mut output = [0_u8; 2];
                read_ahead.read_at(&snap, 0, &mut output)
            })
        };
        entered.recv().unwrap();

        let (started_tx, started_rx) = mpsc::channel();
        let (result_tx, result_rx) = mpsc::channel();
        let queued = {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            thread::spawn(move || {
                started_tx.send(()).unwrap();
                let mut output = [0_u8; 2];
                result_tx
                    .send(read_ahead.read_at(&snap, 16, &mut output))
                    .unwrap();
            })
        };
        started_rx.recv().unwrap();
        thread::sleep(Duration::from_millis(25));
        cancellation.cancel();

        let queued_before_release = result_rx.recv_timeout(Duration::from_millis(200));
        let queued_completed_before_release = queued_before_release.is_ok();
        // Always release the provider before asserting or joining, so a
        // failed cancellation check cannot strand this test's fill thread.
        release.send(()).unwrap();
        let first_result = first.join().unwrap();
        let queued_result = match queued_before_release {
            Ok(result) => result,
            Err(mpsc::RecvTimeoutError::Timeout) => {
                result_rx.recv_timeout(Duration::from_millis(200)).unwrap()
            },
            Err(mpsc::RecvTimeoutError::Disconnected) => {
                panic!("queued reader terminated without reporting its result")
            },
        };
        queued.join().unwrap();

        assert!(queued_completed_before_release);
        assert!(matches!(first_result, Err(OpcError::Cancelled)));
        assert!(matches!(queued_result, Err(OpcError::Cancelled)));
    }

    #[test]
    fn already_cancelled_managed_waiter_does_not_wait_for_a_blocked_fill() {
        let (_budget, cancellation, execution) = context(8, 64);
        let (source, entered, release) = GatedSource::new((0..64).collect());
        let source = Arc::new(source);
        let source_read_at: Arc<dyn ReadAt> = source;
        let snap = Arc::new(snapshot(source_read_at, Some(execution)));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );

        let first = {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            thread::spawn(move || {
                let mut output = [0_u8; 2];
                read_ahead.read_at(&snap, 0, &mut output)
            })
        };
        entered.recv().unwrap();
        cancellation.cancel();

        let (result_tx, result_rx) = mpsc::channel();
        let queued = {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            thread::spawn(move || {
                let mut output = [0_u8; 2];
                result_tx
                    .send(read_ahead.read_at(&snap, 16, &mut output))
                    .unwrap();
            })
        };
        let queued_before_release = result_rx.recv_timeout(Duration::from_millis(200));
        let queued_completed_before_release = queued_before_release.is_ok();
        // Release the blocked provider before validating the queued result or
        // joining either worker, even if the waiter failed to poll promptly.
        release.send(()).unwrap();
        let first_result = first.join().unwrap();
        let queued_result = match queued_before_release {
            Ok(result) => result,
            Err(mpsc::RecvTimeoutError::Timeout) => {
                result_rx.recv_timeout(Duration::from_millis(200)).unwrap()
            },
            Err(mpsc::RecvTimeoutError::Disconnected) => {
                panic!("queued reader terminated without reporting its result")
            },
        };
        queued.join().unwrap();

        assert!(queued_completed_before_release);
        assert!(matches!(first_result, Err(OpcError::Cancelled)));
        assert!(matches!(queued_result, Err(OpcError::Cancelled)));
    }

    #[test]
    fn provider_diagnostics_callback_observes_retained_capacity_without_deadlock() {
        let source = Arc::new(CallbackSource::new((0..16).collect()));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, None);
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );
        let (diagnostics_tx, diagnostics_rx) = mpsc::channel();
        let callback_read_ahead = Arc::clone(&read_ahead);
        source.set_read_callback(Arc::new(move || {
            diagnostics_tx
                .send(callback_read_ahead.diagnostics().unwrap())
                .unwrap();
        }));

        let mut output = [0_u8; 2];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut output).unwrap(), 2);
        let during_fill = diagnostics_rx
            .recv_timeout(Duration::from_millis(100))
            .unwrap();
        assert!(during_fill.enabled);
        assert_eq!(during_fill.retained_window_bytes, 8);
        assert_eq!(read_ahead.diagnostics().unwrap().retained_window_bytes, 8);
    }

    #[test]
    fn version_callback_reentrancy_is_refused_before_recursive_source_callbacks() {
        let source = Arc::new(CallbackSource::new((0..16).collect()));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = Arc::new(snapshot(source_read_at, None));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );
        let weak_read_ahead = Arc::downgrade(&read_ahead);
        let weak_snapshot = Arc::downgrade(&snap);
        source.set_version_callback(Arc::new(move || {
            let read_ahead = weak_read_ahead.upgrade().unwrap();
            let snapshot = weak_snapshot.upgrade().unwrap();
            let mut output = [0_u8; 1];
            assert!(matches!(
                read_ahead.read_at(&snapshot, 0, &mut output),
                Err(OpcError::SourceBackedOverlayUnavailable { .. })
            ));
            assert!(matches!(
                read_ahead.disable_for_exact_publication(&snapshot),
                Err(OpcError::SourceBackedOverlayUnavailable { .. })
            ));
        }));

        let mut output = [0_u8; 2];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut output).unwrap(), 2);
        assert_eq!(output, [0, 1]);
    }

    #[test]
    fn provider_panic_repairs_fill_state_and_transition_releases_memory() {
        let (budget, _cancellation, execution) = context(8, 64);
        let source: Arc<dyn ReadAt> = Arc::new(PanicSource);
        let snap = Arc::new(snapshot(source, Some(execution)));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );
        let read_ahead_for_thread = Arc::clone(&read_ahead);
        let snap_for_thread = Arc::clone(&snap);
        let panic = thread::spawn(move || {
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut output = [0_u8; 1];
                let _ = read_ahead_for_thread.read_at(&snap_for_thread, 0, &mut output);
            }))
        })
        .join()
        .unwrap();
        assert!(panic.is_err());
        assert_eq!(budget.used(Resource::Memory), 8);

        read_ahead.disable_for_exact_publication(&snap).unwrap();
        assert_eq!(budget.used(Resource::Memory), 0);
        let diagnostics = read_ahead.diagnostics().unwrap();
        assert!(!diagnostics.enabled);
        assert_eq!(diagnostics.retained_window_bytes, 0);
        assert_eq!(diagnostics.fills, 0);
    }

    #[test]
    fn poisoned_state_is_recovered_when_exact_transition_drains_memory() {
        let (budget, _cancellation, execution) = context(8, 64);
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new((0..16).collect()));
        let snap = Arc::new(snapshot(source, Some(execution)));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap(),
        );
        let poisoned = Arc::clone(&read_ahead);
        assert!(
            thread::spawn(move || {
                let _state = poisoned.state.lock().unwrap();
                panic!("poison read-ahead state for recovery test");
            })
            .join()
            .is_err()
        );

        read_ahead.disable_for_exact_publication(&snap).unwrap();
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(read_ahead.diagnostics().unwrap().retained_window_bytes, 0);
    }

    #[test]
    fn empty_and_eof_reads_do_not_call_provider_or_retain_zero_returns() {
        let source = Arc::new(CountingSource::new(vec![1, 2, 3]));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, None);
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();
        let mut empty = [9_u8; 0];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut empty).unwrap(), 0);
        let mut eof = [0_u8; 1];
        assert_eq!(read_ahead.read_at(&snap, 99, &mut eof).unwrap(), 0);
        assert_eq!(source.calls.load(Ordering::SeqCst), 0);
        let diagnostics = read_ahead.diagnostics().unwrap();
        assert_eq!(diagnostics.requests, 1);
        assert_eq!(diagnostics.misses, 1);
        assert_eq!(diagnostics.fills, 0);
        assert_eq!(diagnostics.retained_window_bytes, 3);
    }

    #[test]
    fn source_change_invalidates_an_unmonitored_positive_window() {
        let source = Arc::new(CountingSource::new((0..32).collect()));
        let source_read_at: Arc<dyn ReadAt> = source.clone();
        let snap = snapshot(source_read_at, None);
        let read_ahead =
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(8).unwrap()).unwrap();
        let mut output = [0_u8; 2];
        read_ahead.read_at(&snap, 0, &mut output).unwrap();
        source.bump_version();
        assert!(matches!(
            read_ahead.read_at(&snap, 1, &mut output),
            Err(OpcError::SourceChanged { .. })
        ));
        assert_eq!(read_ahead.diagnostics().unwrap().retained_window_bytes, 8);
        assert_eq!(read_ahead.diagnostics().unwrap().hits, 0);
    }

    #[test]
    fn concurrent_forward_readers_share_one_fill_and_transition_does_not_deadlock() {
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new((0..128).collect()));
        let snap = Arc::new(snapshot(source, None));
        let read_ahead = Arc::new(
            ArchiveReadAhead::new(&snap, SourceReadPolicy::forward_start(32).unwrap()).unwrap(),
        );
        let mut joins = Vec::new();
        for _ in 0..8 {
            let read_ahead = Arc::clone(&read_ahead);
            let snap = Arc::clone(&snap);
            joins.push(thread::spawn(move || {
                let mut output = [0_u8; 4];
                read_ahead.read_at(&snap, 0, &mut output).unwrap();
                output
            }));
        }
        for join in joins {
            assert_eq!(join.join().unwrap(), [0, 1, 2, 3]);
        }
        read_ahead.disable_for_exact_publication(&snap).unwrap();
        let mut output = [0_u8; 4];
        assert_eq!(read_ahead.read_at(&snap, 0, &mut output).unwrap(), 4);
        assert_eq!(output, [0, 1, 2, 3]);
        assert!(!read_ahead.diagnostics().unwrap().enabled);
    }
}

#[cfg(test)]
mod structural_run_tests {
    use super::{
        MAX_STRUCTURAL_PREFETCH_BYTES, MAX_STRUCTURAL_PREFETCH_RUN_BYTES, StructuralRun,
        structural_runs,
    };

    fn ranges(runs: &[StructuralRun]) -> Vec<(u64, u64)> {
        runs.iter().map(|run| (run.start, run.end)).collect()
    }

    #[test]
    fn adjoining_spans_join_one_run_and_a_gap_starts_another() {
        // Members 0-2 adjoin; member 3 starts after a hole; members 3-4 adjoin.
        let spans = [(0, 100), (100, 100), (200, 100), (400, 100), (500, 100)];
        let runs = structural_runs(spans).expect("runs");
        assert_eq!(ranges(&runs), vec![(0, 300), (400, 600)]);
    }

    #[test]
    fn a_run_covers_the_union_of_its_members_spans_with_no_hole() {
        // A span that overruns its neighbour's header — the descriptor room a
        // member's own first read already fetches — must not extend the run
        // past the last member's own end.
        let spans = [(0, 124), (100, 124), (200, 100)];
        let runs = structural_runs(spans).expect("runs");
        assert_eq!(ranges(&runs), vec![(0, 300)]);
    }

    #[test]
    fn a_single_member_run_is_not_admitted() {
        let spans = [(0, 100), (500, 100), (1000, 100)];
        assert!(structural_runs(spans).expect("runs").is_empty());
    }

    #[test]
    fn runs_are_disjoint_and_strictly_increasing() {
        let spans = [
            (0, 50),
            (50, 50),
            (300, 40),
            (340, 40),
            (380, 40),
            (900, 10),
            (910, 10),
        ];
        let runs = structural_runs(spans).expect("runs");
        assert_eq!(ranges(&runs), vec![(0, 100), (300, 420), (900, 920)]);
        let mut previous_end = 0;
        for run in &runs {
            assert!(run.start >= previous_end, "runs overlap: {runs:?}");
            previous_end = run.end;
        }
    }

    #[test]
    fn a_run_above_the_single_read_ceiling_is_dropped() {
        let long = MAX_STRUCTURAL_PREFETCH_RUN_BYTES;
        let spans = [(0, long), (long, 100)];
        assert!(structural_runs(spans).expect("runs").is_empty());
        let spans = [(0, long - 100), (long - 100, 100)];
        let runs = structural_runs(spans).expect("runs");
        assert_eq!(ranges(&runs), vec![(0, long)]);
    }

    #[test]
    fn admission_stops_at_the_retained_ceiling() {
        let run_bytes = MAX_STRUCTURAL_PREFETCH_RUN_BYTES;
        let mut spans = Vec::new();
        let mut offset = 0u64;
        // Ten maximum-size runs, separated by holes, is four times the ceiling.
        for _ in 0..10 {
            spans.push((offset, run_bytes / 2));
            spans.push((offset + run_bytes / 2, run_bytes / 2));
            offset += run_bytes * 2;
        }
        let runs = structural_runs(spans).expect("runs");
        let retained: u64 = runs.iter().map(|run| run.end - run.start).sum();
        assert!(
            retained <= MAX_STRUCTURAL_PREFETCH_BYTES as u64,
            "retained {retained} exceeds the ceiling"
        );
        assert_eq!(runs.len(), 4);
    }

    #[test]
    fn spans_out_of_order_admit_nothing() {
        let spans = [(100, 100), (0, 100)];
        assert!(structural_runs(spans).expect("runs").is_empty());
    }

    #[test]
    fn an_empty_member_list_admits_nothing() {
        assert!(structural_runs([]).expect("runs").is_empty());
    }
}
