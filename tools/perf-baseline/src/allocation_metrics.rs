//! Forbid-safe operation-scoped allocation counters for the benchmark targets.
//!
//! The normal binary never installs a global allocator wrapper and never calls
//! [`enable`]. Isolated allocator binaries install the shared `GlobalAlloc`
//! wrapper and call the record functions below after each successful or failed system
//! allocation. Keeping this state and all arithmetic safe lets the shared
//! harness library retain `#![forbid(unsafe_code)]`.
//!
//! Counters are absolute process counters. A region records boundary snapshots
//! and publishes checked differences; it never resets a counter. The region guard
//! is deliberately non-reentrant so a nested operation cannot publish a
//! misleading partial interval. One observer mutex linearizes each callback,
//! boundary, and snapshot after the system allocator has returned, so worker
//! callbacks between the boundaries are included in the region high-water mark.
//! Same-thread callback reentry fails closed through a const TLS guard because
//! a callback must never wait on the observer mutex it already owns.
//! The resulting peak is callback-order evidence: it includes other process
//! threads, excludes allocator-internal realloc overlap and physical RSS, and
//! can perturb allocator scheduling while instrumentation is enabled.

use std::cell::Cell;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Mutex, MutexGuard};

use serde::{Deserialize, Serialize};

const SCOPE: Scope = Scope::OperationGlobalSystemAllocator;

static ENABLED: AtomicBool = AtomicBool::new(false);

thread_local! {
    static CALLBACK_ENTRY_ACTIVE: Cell<bool> = const { Cell::new(false) };
}

#[cfg(test)]
pub(crate) static TEST_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub(crate) enum Scope {
    OperationGlobalSystemAllocator,
}

/// Returns the report identity selected by the executable wrapper.
pub(crate) fn instrumentation_identity() -> &'static str {
    if ENABLED.load(Ordering::Acquire) {
        "system_allocator_operation_scoped"
    } else {
        "none"
    }
}

/// Distinguishes observer-ordered region peaks from older allocator reports.
/// Normal reports omit this allocator-only compatibility identity.
pub(crate) fn counter_revision() -> Option<&'static str> {
    ENABLED
        .load(Ordering::Acquire)
        .then_some("serialized_region_peak_v3")
}

/// Returns the allocator identity selected by the executable wrapper.
pub(crate) fn allocator_identity() -> &'static str {
    if ENABLED.load(Ordering::Acquire) {
        "CountingSystemAllocator(std::alloc::System)"
    } else {
        "Rust system allocator"
    }
}

/// Returns the executable identity included in every report.
pub(crate) fn binary_identity() -> &'static str {
    if ENABLED.load(Ordering::Acquire) {
        "litchi-perf-baseline-alloc"
    } else {
        "litchi-perf-baseline"
    }
}

/// Enables publication of allocation regions for the allocator-only target.
///
/// The normal binary has no call site for this function and no global wrapper,
/// so compiling all package targets/features cannot instrument it.
pub fn enable() {
    ENABLED.store(true, Ordering::Release);
}

/// Records a successful `alloc` or `alloc_zeroed` call.
pub fn record_allocation(size: usize) {
    COUNTERS.allocation(size);
}

/// Records a successful `dealloc` call.
pub fn record_deallocation(size: usize) {
    COUNTERS.deallocation(size);
}

/// Records a successful `realloc` call.
pub fn record_reallocation(old_size: usize, new_size: usize) {
    COUNTERS.reallocation(old_size, new_size);
}

/// Records a failed allocation or reallocation call.
pub fn record_failed_allocation() {
    COUNTERS.failed_allocation();
}

/// A raw absolute counter snapshot used by the allocator target's unit tests.
#[doc(hidden)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Snapshot {
    pub allocation_calls: u64,
    pub deallocation_calls: u64,
    pub reallocation_calls: u64,
    pub failed_allocation_calls: u64,
    pub allocated_bytes: u64,
    pub deallocated_bytes: u64,
    pub live_bytes: u64,
    pub peak_live_bytes: u64,
    pub overflowed: bool,
    /// Sticky observer failure state. Numeric fields may be incomplete when set.
    pub observer_invalid: bool,
}

/// Returns absolute counters without resetting them.
#[doc(hidden)]
pub fn snapshot() -> Snapshot {
    let snapshot = COUNTERS.snapshot();
    Snapshot {
        allocation_calls: snapshot.allocation_calls,
        deallocation_calls: snapshot.deallocation_calls,
        reallocation_calls: snapshot.reallocation_calls,
        failed_allocation_calls: snapshot.failed_allocation_calls,
        allocated_bytes: snapshot.allocated_bytes,
        deallocated_bytes: snapshot.deallocated_bytes,
        live_bytes: snapshot.live_bytes,
        peak_live_bytes: snapshot.peak_live_bytes,
        overflowed: snapshot.overflowed,
        observer_invalid: snapshot.observer_invalid,
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub(crate) enum Status {
    /// Every absolute counter difference was checked successfully.
    Measured,
    /// The region could not be acquired.
    Unavailable,
    /// An absolute counter or checked difference overflowed.
    Overflow,
}

/// One operation's allocation observation. Numeric fields are omitted unless
/// the complete observation is measured; callers must not mix partial vectors.
#[derive(Clone, Debug, Deserialize, Serialize, PartialEq, Eq)]
pub(crate) struct Sample {
    pub status: Status,
    pub scope: Scope,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub allocation_calls: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub deallocation_calls: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reallocation_calls: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub failed_allocation_calls: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub allocated_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub deallocated_bytes: Option<u64>,
    /// Absolute process live bytes immediately before the operation.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub live_bytes_before: Option<u64>,
    /// Absolute process live bytes immediately after the operation.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub live_bytes_after: Option<u64>,
    /// Absolute process high-water live bytes before the operation.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub peak_live_bytes_before: Option<u64>,
    /// Absolute process high-water live bytes after the operation.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub peak_live_bytes_after: Option<u64>,
    /// Observer-ordered absolute live-byte high-water mark during this region.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub region_peak_live_bytes: Option<u64>,
}

/// Return an explicit unavailable sample for binaries that do not install the
/// benchmark allocator wrapper.  Keeping this state in the sample envelope
/// lets operation-scoped reports distinguish an uninstrumented run from a
/// measured zero without changing the normal binary's allocator behavior.
pub(crate) fn unavailable_sample() -> Sample {
    Sample::unavailable()
}

impl Sample {
    fn unavailable() -> Self {
        Self {
            status: Status::Unavailable,
            scope: SCOPE,
            allocation_calls: None,
            deallocation_calls: None,
            reallocation_calls: None,
            failed_allocation_calls: None,
            allocated_bytes: None,
            deallocated_bytes: None,
            live_bytes_before: None,
            live_bytes_after: None,
            peak_live_bytes_before: None,
            peak_live_bytes_after: None,
            region_peak_live_bytes: None,
        }
    }

    fn overflow() -> Self {
        Self {
            status: Status::Overflow,
            ..Self::unavailable()
        }
    }

    fn measured(
        before: CountersSnapshot,
        after: CountersSnapshot,
        region_peak_live_bytes: Option<u64>,
    ) -> Self {
        let values = [
            difference(before.allocation_calls, after.allocation_calls),
            difference(before.deallocation_calls, after.deallocation_calls),
            difference(before.reallocation_calls, after.reallocation_calls),
            difference(
                before.failed_allocation_calls,
                after.failed_allocation_calls,
            ),
            difference(before.allocated_bytes, after.allocated_bytes),
            difference(before.deallocated_bytes, after.deallocated_bytes),
        ];
        if before.observer_invalid || after.observer_invalid {
            return Self::unavailable();
        }
        let overflow = before.overflowed
            || after.overflowed
            || values.iter().any(Option::is_none)
            || after.peak_live_bytes < before.peak_live_bytes
            || region_peak_live_bytes
                .map(|peak| {
                    peak < before.live_bytes
                        || peak < after.live_bytes
                        || peak > after.peak_live_bytes
                })
                .unwrap_or(true);
        if overflow {
            return Self::overflow();
        }
        Self {
            status: Status::Measured,
            scope: SCOPE,
            allocation_calls: values[0],
            deallocation_calls: values[1],
            reallocation_calls: values[2],
            failed_allocation_calls: values[3],
            allocated_bytes: values[4],
            deallocated_bytes: values[5],
            live_bytes_before: Some(before.live_bytes),
            live_bytes_after: Some(after.live_bytes),
            peak_live_bytes_before: Some(before.peak_live_bytes),
            peak_live_bytes_after: Some(after.peak_live_bytes),
            region_peak_live_bytes,
        }
    }
}

enum RegionState {
    Disabled,
    Unavailable,
    Active(ActiveRegion),
}

struct ActiveRegion {
    /// The snapshot at the beginning of the complete operation.  This is
    /// intentionally retained across segment boundaries so `finish` keeps
    /// the historical combined-operation contract.
    before: CountersSnapshot,
    /// The snapshot at the beginning of the current non-overlapping segment.
    segment_before: CountersSnapshot,
    /// The maximum of every segment peak observed so far.  The observer's
    /// live peak is rebased at each split, so this field preserves the exact
    /// peak that one unsplit region would have published.
    combined_peak_live_bytes: Option<u64>,
    /// Sticky failure for a split boundary whose evidence was not complete.
    /// This is separate from observer invalidity, which reports unavailable.
    split_overflowed: bool,
}

/// A region that owns the non-overlap token until it is finished or dropped.
pub(crate) struct Region {
    state: RegionState,
}

impl Region {
    /// Publishes the current segment and rebases the same region for the next
    /// segment without releasing its non-overlap token.  The observer lock is
    /// held while the snapshot and rebase occur, so callbacks cannot fall into
    /// an unaccounted gap between the two segments.
    ///
    /// The complete region remains active until [`Region::finish`].  Its
    /// returned sample therefore still covers the original start through the
    /// eventual finish, including the maximum peak from every split segment.
    pub(crate) fn split(&mut self) -> Option<Sample> {
        match &mut self.state {
            RegionState::Disabled => None,
            RegionState::Unavailable => Some(Sample::unavailable()),
            RegionState::Active(active) => {
                let (after, segment_peak_live_bytes, observer_valid) = COUNTERS.split_region();
                if let Some(segment_peak) = segment_peak_live_bytes {
                    active.combined_peak_live_bytes = Some(
                        active
                            .combined_peak_live_bytes
                            .map_or(segment_peak, |combined_peak| {
                                combined_peak.max(segment_peak)
                            }),
                    );
                }
                let sample = if !observer_valid {
                    Sample::unavailable()
                } else if active.split_overflowed || segment_peak_live_bytes.is_none() {
                    active.split_overflowed = true;
                    Sample::overflow()
                } else {
                    Sample::measured(active.segment_before, after, segment_peak_live_bytes)
                };
                if sample.status == Status::Overflow {
                    active.split_overflowed = true;
                }
                active.segment_before = after;
                Some(sample)
            },
        }
    }

    /// Finishes the current segment and the complete region at one observer
    /// boundary.  Returning both samples together prevents callbacks from
    /// landing between the commit-core endpoint and the combined-operation
    /// endpoint.
    pub(crate) fn finish_split(mut self) -> (Option<Sample>, Option<Sample>) {
        let state = std::mem::replace(&mut self.state, RegionState::Disabled);
        match state {
            RegionState::Disabled => (None, None),
            RegionState::Unavailable => {
                let sample = Sample::unavailable();
                (Some(sample.clone()), Some(sample))
            },
            RegionState::Active(active) => {
                let (after, final_peak_live_bytes, observer_valid) = COUNTERS.finish_region();
                if !observer_valid {
                    return (Some(Sample::unavailable()), Some(Sample::unavailable()));
                }
                if active.split_overflowed {
                    let sample = Sample::overflow();
                    return (Some(sample.clone()), Some(sample));
                }
                let segment = Sample::measured(active.segment_before, after, final_peak_live_bytes);
                if segment.status == Status::Overflow {
                    let sample = Sample::overflow();
                    return (Some(sample.clone()), Some(sample));
                }
                let combined_peak_live_bytes =
                    match (active.combined_peak_live_bytes, final_peak_live_bytes) {
                        (Some(combined_peak), Some(final_peak)) => {
                            Some(combined_peak.max(final_peak))
                        },
                        (None, final_peak) => final_peak,
                        (Some(_), None) => None,
                    };
                let combined = Sample::measured(active.before, after, combined_peak_live_bytes);
                (Some(segment), Some(combined))
            },
        }
    }

    /// Ends the region and releases its token. No allocator counters are
    /// reset, and this method performs no heap allocation itself.
    pub(crate) fn finish(mut self) -> Option<Sample> {
        let state = std::mem::replace(&mut self.state, RegionState::Disabled);
        match state {
            RegionState::Disabled => None,
            RegionState::Unavailable => Some(Sample::unavailable()),
            RegionState::Active(active) => {
                let sample = if active.split_overflowed {
                    // The sticky split failure still owns the observer token.
                    // Consume its boundary before returning the fail-closed
                    // sample so a normal `finish` cannot strand the marker.
                    let (_, _, observer_valid) = COUNTERS.finish_region();
                    if observer_valid {
                        Sample::overflow()
                    } else {
                        Sample::unavailable()
                    }
                } else if active.combined_peak_live_bytes.is_none() {
                    COUNTERS.finish_sample(active.before)
                } else {
                    COUNTERS.finish_sample_with_combined_peak(
                        active.before,
                        active.combined_peak_live_bytes,
                    )
                };
                Some(sample)
            },
        }
    }
}

impl Drop for Region {
    fn drop(&mut self) {
        if matches!(self.state, RegionState::Active(_)) {
            COUNTERS.release_region();
            self.state = RegionState::Disabled;
        }
    }
}

/// Starts a region when operation instrumentation is enabled. A failed
/// acquisition returns a region which publishes `unavailable` and leaves the
/// operation itself untouched.
pub(crate) fn begin() -> Region {
    if !ENABLED.load(Ordering::Acquire) {
        return Region {
            state: RegionState::Disabled,
        };
    }
    match COUNTERS.begin_region() {
        Some(before) => Region {
            state: RegionState::Active(ActiveRegion {
                before,
                segment_before: before,
                combined_peak_live_bytes: None,
                split_overflowed: false,
            }),
        },
        None => Region {
            state: RegionState::Unavailable,
        },
    }
}

#[derive(Default)]
struct ObserverState {
    region_peak_live_bytes: Option<u64>,
}

#[derive(Default)]
struct Counters {
    allocation_calls: AtomicU64,
    deallocation_calls: AtomicU64,
    reallocation_calls: AtomicU64,
    failed_allocation_calls: AtomicU64,
    allocated_bytes: AtomicU64,
    deallocated_bytes: AtomicU64,
    live_bytes: AtomicU64,
    peak_live_bytes: AtomicU64,
    overflowed: AtomicBool,
    /// Poison/reentry invalidity is distinct from checked arithmetic overflow.
    observer_invalid: AtomicBool,
    observer: Mutex<ObserverState>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct CountersSnapshot {
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_allocation_calls: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    overflowed: bool,
    observer_invalid: bool,
}

struct CallbackEntryGuard;

impl CallbackEntryGuard {
    fn enter(observer_invalid: &AtomicBool) -> Option<Self> {
        let already_active = match CALLBACK_ENTRY_ACTIVE.try_with(|active| active.replace(true)) {
            Ok(already_active) => already_active,
            Err(_) => {
                observer_invalid.store(true, Ordering::Release);
                return None;
            },
        };
        if already_active {
            observer_invalid.store(true, Ordering::Release);
            None
        } else {
            Some(Self)
        }
    }
}

impl Drop for CallbackEntryGuard {
    fn drop(&mut self) {
        let _ = CALLBACK_ENTRY_ACTIVE.try_with(|active| active.set(false));
    }
}

impl Counters {
    fn snapshot(&self) -> CountersSnapshot {
        let _observer = self.lock_observer();
        self.snapshot_locked()
    }

    fn snapshot_locked(&self) -> CountersSnapshot {
        CountersSnapshot {
            allocation_calls: self.allocation_calls.load(Ordering::Acquire),
            deallocation_calls: self.deallocation_calls.load(Ordering::Acquire),
            reallocation_calls: self.reallocation_calls.load(Ordering::Acquire),
            failed_allocation_calls: self.failed_allocation_calls.load(Ordering::Acquire),
            allocated_bytes: self.allocated_bytes.load(Ordering::Acquire),
            deallocated_bytes: self.deallocated_bytes.load(Ordering::Acquire),
            live_bytes: self.live_bytes.load(Ordering::Acquire),
            peak_live_bytes: self.peak_live_bytes.load(Ordering::Acquire),
            overflowed: self.overflowed.load(Ordering::Acquire),
            observer_invalid: self.observer_invalid.load(Ordering::Acquire),
        }
    }

    fn lock_observer(&self) -> MutexGuard<'_, ObserverState> {
        match self.observer.lock() {
            Ok(observer) => observer,
            Err(poisoned) => {
                let mut observer = poisoned.into_inner();
                self.observer_invalid.store(true, Ordering::Release);
                observer.region_peak_live_bytes = None;
                observer
            },
        }
    }

    fn begin_region(&self) -> Option<CountersSnapshot> {
        if self.observer_invalid.load(Ordering::Acquire) {
            return None;
        }
        let mut observer = self.lock_observer();
        if self.observer_invalid.load(Ordering::Acquire)
            || observer.region_peak_live_bytes.is_some()
        {
            return None;
        }
        let before = self.snapshot_locked();
        observer.region_peak_live_bytes = Some(before.live_bytes);
        Some(before)
    }

    fn split_region(&self) -> (CountersSnapshot, Option<u64>, bool) {
        let mut observer = self.lock_observer();
        let after = self.snapshot_locked();
        let region_peak_live_bytes = observer.region_peak_live_bytes;
        let observer_valid = !self.observer_invalid.load(Ordering::Acquire);
        if observer_valid {
            if region_peak_live_bytes.is_none() {
                // An active region must always own an observer peak. Keep
                // this accounting failure sticky so a later valid boundary
                // cannot repair an earlier incomplete sample.
                self.overflowed.store(true, Ordering::Release);
            }
            observer.region_peak_live_bytes = Some(after.live_bytes);
        } else {
            observer.region_peak_live_bytes = None;
        }
        (after, region_peak_live_bytes, observer_valid)
    }

    fn finish_region(&self) -> (CountersSnapshot, Option<u64>, bool) {
        let mut observer = self.lock_observer();
        let after = self.snapshot_locked();
        let region_peak_live_bytes = observer.region_peak_live_bytes.take();
        let observer_valid = !self.observer_invalid.load(Ordering::Acquire);
        (after, region_peak_live_bytes, observer_valid)
    }

    fn finish_sample(&self, before: CountersSnapshot) -> Sample {
        self.finish_sample_with_combined_peak(before, None)
    }

    fn finish_sample_with_combined_peak(
        &self,
        before: CountersSnapshot,
        combined_peak_live_bytes: Option<u64>,
    ) -> Sample {
        let (after, final_peak_live_bytes, observer_valid) = self.finish_region();
        if observer_valid {
            let region_peak_live_bytes = match (combined_peak_live_bytes, final_peak_live_bytes) {
                (Some(combined_peak), Some(final_peak)) => Some(combined_peak.max(final_peak)),
                (None, Some(final_peak)) => Some(final_peak),
                // A split region must still have a live final segment peak.
                // Do not conceal a broken observer state with an earlier peak.
                (Some(_), None) => None,
                (None, None) => None,
            };
            Sample::measured(before, after, region_peak_live_bytes)
        } else {
            Sample::unavailable()
        }
    }

    fn release_region(&self) {
        let mut observer = self.lock_observer();
        observer.region_peak_live_bytes = None;
    }

    fn update_region_peak(&self, observer: &mut ObserverState, live: u64) {
        if let Some(peak) = observer.region_peak_live_bytes.as_mut()
            && live > *peak
        {
            *peak = live;
        }
    }

    fn enter_callback(&self) -> Option<CallbackEntryGuard> {
        CallbackEntryGuard::enter(&self.observer_invalid)
    }

    fn add(&self, counter: &AtomicU64, value: u64) {
        if checked_add(counter, value).is_none() {
            self.overflowed.store(true, Ordering::Release);
        }
    }

    fn live_add(&self, value: u64) -> Option<u64> {
        if let Some(live) = checked_add(&self.live_bytes, value) {
            self.update_peak(live);
            Some(live)
        } else {
            self.overflowed.store(true, Ordering::Release);
            None
        }
    }

    fn live_sub(&self, value: u64) {
        if checked_sub(&self.live_bytes, value).is_none() {
            self.overflowed.store(true, Ordering::Release);
        }
    }

    fn update_peak(&self, live: u64) {
        let mut peak = self.peak_live_bytes.load(Ordering::Acquire);
        while live > peak {
            match self.peak_live_bytes.compare_exchange(
                peak,
                live,
                Ordering::AcqRel,
                Ordering::Acquire,
            ) {
                Ok(_) => break,
                Err(observed) => peak = observed,
            }
        }
    }

    fn allocation(&self, size: usize) {
        let Some(_entry) = self.enter_callback() else {
            return;
        };
        let mut observer = self.lock_observer();
        self.allocation_locked(size, &mut observer);
    }

    fn allocation_locked(&self, size: usize, observer: &mut ObserverState) {
        let Some(size) = u64::try_from(size).ok() else {
            self.overflowed.store(true, Ordering::Release);
            return;
        };
        self.add(&self.allocation_calls, 1);
        self.add(&self.allocated_bytes, size);
        if let Some(live) = self.live_add(size) {
            self.update_region_peak(observer, live);
        }
    }

    fn deallocation(&self, size: usize) {
        let Some(_entry) = self.enter_callback() else {
            return;
        };
        let _observer = self.lock_observer();
        self.deallocation_locked(size);
    }

    fn deallocation_locked(&self, size: usize) {
        let Some(size) = u64::try_from(size).ok() else {
            self.overflowed.store(true, Ordering::Release);
            return;
        };
        self.add(&self.deallocation_calls, 1);
        self.add(&self.deallocated_bytes, size);
        self.live_sub(size);
    }

    fn reallocation(&self, old_size: usize, new_size: usize) {
        let Some(_entry) = self.enter_callback() else {
            return;
        };
        let mut observer = self.lock_observer();
        self.reallocation_locked(old_size, new_size, &mut observer);
    }

    fn reallocation_locked(&self, old_size: usize, new_size: usize, observer: &mut ObserverState) {
        let (Some(old_size), Some(new_size)) =
            (u64::try_from(old_size).ok(), u64::try_from(new_size).ok())
        else {
            self.overflowed.store(true, Ordering::Release);
            return;
        };
        self.add(&self.allocation_calls, 1);
        self.add(&self.reallocation_calls, 1);
        self.add(&self.allocated_bytes, new_size);
        self.add(&self.deallocated_bytes, old_size);
        if new_size >= old_size {
            if let Some(live) = self.live_add(new_size - old_size) {
                self.update_region_peak(observer, live);
            }
        } else {
            self.live_sub(old_size - new_size);
        }
    }

    fn failed_allocation(&self) {
        let Some(_entry) = self.enter_callback() else {
            return;
        };
        let _observer = self.lock_observer();
        self.add(&self.failed_allocation_calls, 1);
    }
}

fn checked_add(counter: &AtomicU64, value: u64) -> Option<u64> {
    let current = counter.load(Ordering::Acquire);
    let next = current.checked_add(value)?;
    counter.store(next, Ordering::Release);
    Some(next)
}

fn checked_sub(counter: &AtomicU64, value: u64) -> Option<u64> {
    let current = counter.load(Ordering::Acquire);
    let next = current.checked_sub(value)?;
    counter.store(next, Ordering::Release);
    Some(next)
}

static COUNTERS: Counters = Counters {
    allocation_calls: AtomicU64::new(0),
    deallocation_calls: AtomicU64::new(0),
    reallocation_calls: AtomicU64::new(0),
    failed_allocation_calls: AtomicU64::new(0),
    allocated_bytes: AtomicU64::new(0),
    deallocated_bytes: AtomicU64::new(0),
    live_bytes: AtomicU64::new(0),
    peak_live_bytes: AtomicU64::new(0),
    overflowed: AtomicBool::new(false),
    observer_invalid: AtomicBool::new(false),
    observer: Mutex::new(ObserverState {
        region_peak_live_bytes: None,
    }),
};

fn difference(before: u64, after: u64) -> Option<u64> {
    after.checked_sub(before)
}

#[cfg(test)]
mod tests {
    use std::sync::{Arc, Barrier};

    use super::{Counters, Sample, Scope, Status};

    use super::TEST_LOCK;

    #[test]
    fn disabled_region_publishes_no_sample() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(false, std::sync::atomic::Ordering::SeqCst);
        let mut region = super::begin();
        assert!(region.split().is_none());
        assert!(region.finish().is_none());
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn sample_status_serialization_is_explicit() {
        let sample = Sample {
            status: Status::Overflow,
            scope: Scope::OperationGlobalSystemAllocator,
            allocation_calls: None,
            deallocation_calls: None,
            reallocation_calls: None,
            failed_allocation_calls: None,
            allocated_bytes: None,
            deallocated_bytes: None,
            live_bytes_before: None,
            live_bytes_after: None,
            peak_live_bytes_before: None,
            peak_live_bytes_after: None,
            region_peak_live_bytes: None,
        };
        let value = serde_json::to_value(sample).unwrap();
        assert_eq!(value["status"], "overflow");
        assert_eq!(value["scope"], "operation_global_system_allocator");
        assert!(value.get("allocation_calls").is_none());
    }

    #[test]
    fn checked_difference_turns_counter_wrap_into_overflow_status() {
        let _lock = TEST_LOCK.lock().unwrap();
        let before = super::CountersSnapshot {
            allocated_bytes: u64::MAX,
            ..super::CountersSnapshot::default()
        };
        let after = super::CountersSnapshot::default();
        let sample = Sample::measured(before, after, None);
        assert_eq!(sample.status, Status::Overflow);
        assert!(sample.allocated_bytes.is_none());
        assert_eq!(super::difference(u64::MAX, 0), None);
    }

    #[test]
    fn counter_arithmetic_overflow_and_live_underflow_are_sticky() {
        let overflow = Counters::default();
        overflow
            .allocation_calls
            .store(u64::MAX, std::sync::atomic::Ordering::Relaxed);
        overflow.allocation(1);
        assert!(overflow.snapshot().overflowed);

        let underflow = Counters::default();
        underflow.live_sub(1);
        assert!(underflow.snapshot().overflowed);
        assert_eq!(underflow.snapshot().live_bytes, 0);
    }

    #[test]
    fn poisoned_observer_is_recovered_as_unavailable() {
        let counters = Arc::new(Counters::default());
        let before = counters.begin_region().unwrap();
        let poisoned = Arc::clone(&counters);
        let handle = std::thread::spawn(move || {
            let _observer = poisoned.observer.lock().unwrap();
            panic!("deliberately poison the observer mutex");
        });
        assert!(handle.join().is_err());

        let sample = counters.finish_sample(before);
        assert_eq!(sample.status, Status::Unavailable);
        assert!(sample.allocation_calls.is_none());
        let snapshot = counters.snapshot();
        assert!(!snapshot.overflowed);
        assert!(snapshot.observer_invalid);
        assert!(counters.begin_region().is_none());
        counters.failed_allocation();
        assert!(counters.snapshot().observer_invalid);
    }

    #[test]
    fn nested_callback_entry_is_suppressed_and_invalidates_future_regions() {
        let counters = Counters::default();
        let outer = counters.enter_callback().unwrap();

        counters.allocation(8);

        drop(outer);
        let snapshot = counters.snapshot();
        assert_eq!(snapshot.allocation_calls, 0);
        assert!(!snapshot.overflowed);
        assert!(snapshot.observer_invalid);
        assert!(counters.begin_region().is_none());
    }

    #[test]
    fn one_allocation_then_deallocation_retains_peak_live_bytes() {
        let counters = Counters::default();
        counters.allocation(64);
        counters.deallocation(64);

        let snapshot = counters.snapshot();
        assert_eq!(snapshot.live_bytes, 0);
        assert_eq!(snapshot.peak_live_bytes, 64);
        assert_eq!(snapshot.allocation_calls, 1);
        assert_eq!(snapshot.deallocation_calls, 1);
    }

    #[test]
    fn reallocation_growth_and_shrink_update_live_bytes_without_losing_peak() {
        let counters = Counters::default();
        counters.allocation(16);
        counters.reallocation(16, 40);

        let grown = counters.snapshot();
        assert_eq!(grown.live_bytes, 40);
        assert_eq!(grown.peak_live_bytes, 40);
        assert_eq!(grown.reallocation_calls, 1);

        counters.reallocation(40, 8);
        let shrunk = counters.snapshot();
        assert_eq!(shrunk.live_bytes, 8);
        assert_eq!(shrunk.peak_live_bytes, 40);
        assert_eq!(shrunk.reallocation_calls, 2);
    }

    #[test]
    fn failed_allocation_does_not_change_live_or_peak_bytes() {
        let counters = Counters::default();
        counters.allocation(48);
        let before = counters.snapshot();

        counters.failed_allocation();

        let after = counters.snapshot();
        assert_eq!(before.live_bytes, 48);
        assert_eq!(before.peak_live_bytes, 48);
        assert_eq!(after.live_bytes, before.live_bytes);
        assert_eq!(after.peak_live_bytes, before.peak_live_bytes);
        assert_eq!(after.failed_allocation_calls, 1);
        assert_eq!(after.allocation_calls, before.allocation_calls);
    }

    #[test]
    fn multi_step_counter_state_matches_an_independent_live_reference() {
        let counters = Counters::default();
        let operations = [
            (true, 10_usize),
            (true, 20),
            (false, 10),
            (true, 5),
            (false, 20),
            (false, 5),
        ];
        let mut reference_live = 0_u64;
        let mut reference_peak = 0_u64;

        for (allocate, size) in operations {
            if allocate {
                counters.allocation(size);
                reference_live += size as u64;
                reference_peak = reference_peak.max(reference_live);
            } else {
                counters.deallocation(size);
                reference_live -= size as u64;
            }
            let snapshot = counters.snapshot();
            assert_eq!(snapshot.live_bytes, reference_live);
            assert_eq!(snapshot.peak_live_bytes, reference_peak);
        }
    }

    #[test]
    fn concurrent_allocations_retain_the_exact_overlapping_peak() {
        const THREADS: usize = 6;
        const SIZE: usize = 17;
        let counters = Arc::new(Counters::default());
        let allocated = Arc::new(Barrier::new(THREADS + 1));
        let release = Arc::new(Barrier::new(THREADS + 1));
        let handles = (0..THREADS)
            .map(|_| {
                let counters = Arc::clone(&counters);
                let allocated = Arc::clone(&allocated);
                let release = Arc::clone(&release);
                std::thread::spawn(move || {
                    counters.allocation(SIZE);
                    allocated.wait();
                    release.wait();
                    counters.deallocation(SIZE);
                })
            })
            .collect::<Vec<_>>();

        allocated.wait();
        let during_overlap = counters.snapshot();
        let expected_peak = (THREADS * SIZE) as u64;
        release.wait();
        for handle in handles {
            handle.join().unwrap();
        }
        let after = counters.snapshot();
        assert_eq!(during_overlap.live_bytes, expected_peak);
        assert_eq!(during_overlap.peak_live_bytes, expected_peak);
        assert_eq!(after.live_bytes, 0);
        assert_eq!(after.peak_live_bytes, expected_peak);
    }

    #[test]
    fn region_peak_is_separate_from_historical_process_peak() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let baseline = super::COUNTERS.snapshot();
        super::COUNTERS.allocation(32);
        super::COUNTERS.deallocation(32);

        let sample = super::begin().finish().unwrap();

        assert_eq!(sample.status, Status::Measured);
        assert_eq!(sample.region_peak_live_bytes, Some(baseline.live_bytes));
        assert!(sample.peak_live_bytes_before.unwrap() > baseline.live_bytes);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn region_peak_retains_an_entry_allocation_after_it_is_freed() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let baseline = super::COUNTERS.snapshot();
        super::COUNTERS.allocation(24);
        let entry_live = baseline.live_bytes + 24;
        let region = super::begin();
        super::COUNTERS.deallocation(24);

        let sample = region.finish().unwrap();

        assert_eq!(sample.status, Status::Measured);
        assert_eq!(sample.live_bytes_before, Some(entry_live));
        assert_eq!(sample.live_bytes_after, Some(baseline.live_bytes));
        assert_eq!(sample.region_peak_live_bytes, Some(entry_live));
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn region_reallocation_growth_shrink_and_failed_allocations_are_ordered() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let baseline = super::COUNTERS.snapshot();
        let region = super::begin();
        super::COUNTERS.allocation(8);
        super::COUNTERS.reallocation(8, 40);
        super::COUNTERS.failed_allocation();
        super::COUNTERS.reallocation(40, 12);
        super::COUNTERS.deallocation(12);

        let sample = region.finish().unwrap();

        assert_eq!(sample.status, Status::Measured);
        assert_eq!(sample.live_bytes_before, Some(baseline.live_bytes));
        assert_eq!(sample.live_bytes_after, Some(baseline.live_bytes));
        assert_eq!(
            sample.region_peak_live_bytes,
            Some(baseline.live_bytes + 40)
        );
        assert_eq!(sample.reallocation_calls, Some(2));
        assert_eq!(sample.failed_allocation_calls, Some(1));
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn reentrant_callback_makes_active_and_future_regions_unavailable() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let was_invalid = super::COUNTERS
            .observer_invalid
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let region = super::begin();
        let outer = super::COUNTERS.enter_callback().unwrap();
        super::COUNTERS.allocation(8);
        drop(outer);

        let sample = region.finish().unwrap();
        let future = super::begin().finish().unwrap();

        assert_eq!(sample.status, Status::Unavailable);
        assert!(sample.allocation_calls.is_none());
        assert!(sample.region_peak_live_bytes.is_none());
        assert_eq!(future.status, Status::Unavailable);
        super::COUNTERS
            .observer_invalid
            .store(was_invalid, std::sync::atomic::Ordering::SeqCst);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn repeated_boundaries_and_dropped_regions_release_only_the_owner() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);

        let first = super::begin();
        let nested = super::begin();
        assert_eq!(nested.finish().unwrap().status, Status::Unavailable);
        assert_eq!(first.finish().unwrap().status, Status::Measured);

        let dropped = super::begin();
        drop(dropped);
        assert_eq!(super::begin().finish().unwrap().status, Status::Measured);

        let active_after = {
            let observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes.is_some()
        };
        assert!(!active_after);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn region_peak_includes_concurrent_callbacks_before_finish() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let baseline = super::COUNTERS.snapshot();
        let region = super::begin();
        const THREADS: usize = 2;
        const SIZE: usize = 19;
        let ready = Arc::new(Barrier::new(THREADS + 1));
        let release = Arc::new(Barrier::new(THREADS + 1));
        let handles = (0..THREADS)
            .map(|_| {
                let ready = Arc::clone(&ready);
                let release = Arc::clone(&release);
                std::thread::spawn(move || {
                    super::COUNTERS.allocation(SIZE);
                    ready.wait();
                    release.wait();
                    super::COUNTERS.deallocation(SIZE);
                })
            })
            .collect::<Vec<_>>();

        ready.wait();
        release.wait();
        for handle in handles {
            handle.join().unwrap();
        }
        let sample = region.finish().unwrap();

        let expected_peak = baseline.live_bytes + (THREADS * SIZE) as u64;
        assert_eq!(sample.status, Status::Measured);
        assert_eq!(sample.region_peak_live_bytes, Some(expected_peak));
        assert_eq!(sample.live_bytes_after, Some(baseline.live_bytes));
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn region_counts_cross_thread_totals_without_resetting_absolute_counters() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let before = super::COUNTERS.snapshot();
        let region = super::begin();
        let handles = (0..2)
            .map(|_| std::thread::spawn(|| super::record_allocation(128)))
            .collect::<Vec<_>>();
        for handle in handles {
            handle.join().unwrap();
        }
        let sample = region.finish().unwrap();
        let after = super::COUNTERS.snapshot();
        assert_eq!(sample.status, Status::Measured);
        assert!(sample.allocation_calls.unwrap() >= 2);
        assert_eq!(sample.live_bytes_before, Some(before.live_bytes));
        assert!(after.allocation_calls >= before.allocation_calls);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn region_guard_rejects_overlap_without_heap_scope_state() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let outer = super::begin();
        let before = super::COUNTERS.snapshot();
        let inner = super::begin();
        assert_eq!(inner.finish().unwrap().status, Status::Unavailable);
        assert_eq!(super::COUNTERS.snapshot(), before);
        assert_eq!(outer.finish().unwrap().status, Status::Measured);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn split_regions_are_sequential_and_preserve_the_combined_interval() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let was_invalid = super::COUNTERS
            .observer_invalid
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let was_overflowed = super::COUNTERS
            .overflowed
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let baseline = super::COUNTERS.snapshot();
        let mut region = super::begin();

        super::COUNTERS.allocation(16);
        super::COUNTERS.deallocation(4);
        let staging = region.split().expect("active staging segment");
        assert_eq!(staging.status, Status::Measured);

        // The same region retains the non-overlap token across the boundary;
        // a concurrent/nested region cannot create an unobserved callback gap.
        let nested = super::begin();
        assert_eq!(nested.finish().unwrap().status, Status::Unavailable);

        super::COUNTERS.allocation(40);
        super::COUNTERS.reallocation(40, 8);
        let (commit_core, combined) = region.finish_split();
        let commit_core = commit_core.expect("active commit segment");
        assert_eq!(commit_core.status, Status::Measured);
        let combined = combined.expect("active combined region");
        assert_eq!(combined.status, Status::Measured);
        assert_eq!(staging.live_bytes_before, Some(baseline.live_bytes));
        assert_eq!(staging.live_bytes_after, commit_core.live_bytes_before);
        assert_eq!(combined.live_bytes_before, staging.live_bytes_before);
        assert_eq!(combined.live_bytes_after, commit_core.live_bytes_after);
        assert_eq!(
            combined.peak_live_bytes_before,
            staging.peak_live_bytes_before
        );
        assert_eq!(
            combined.peak_live_bytes_after,
            commit_core.peak_live_bytes_after
        );
        assert_eq!(
            combined.region_peak_live_bytes,
            Some(
                staging
                    .region_peak_live_bytes
                    .unwrap()
                    .max(commit_core.region_peak_live_bytes.unwrap()),
            )
        );
        for (staging_value, core_value, combined_value) in [
            (
                staging.allocation_calls,
                commit_core.allocation_calls,
                combined.allocation_calls,
            ),
            (
                staging.deallocation_calls,
                commit_core.deallocation_calls,
                combined.deallocation_calls,
            ),
            (
                staging.reallocation_calls,
                commit_core.reallocation_calls,
                combined.reallocation_calls,
            ),
            (
                staging.failed_allocation_calls,
                commit_core.failed_allocation_calls,
                combined.failed_allocation_calls,
            ),
            (
                staging.allocated_bytes,
                commit_core.allocated_bytes,
                combined.allocated_bytes,
            ),
            (
                staging.deallocated_bytes,
                commit_core.deallocated_bytes,
                combined.deallocated_bytes,
            ),
        ] {
            assert_eq!(
                Some(staging_value.unwrap() + core_value.unwrap()),
                combined_value
            );
        }
        super::COUNTERS.deallocation(20);
        super::COUNTERS
            .overflowed
            .store(was_overflowed, std::sync::atomic::Ordering::SeqCst);
        super::COUNTERS
            .observer_invalid
            .store(was_invalid, std::sync::atomic::Ordering::SeqCst);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn split_regions_publish_measured_zero_and_repeated_boundaries() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let was_invalid = super::COUNTERS
            .observer_invalid
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let was_overflowed = super::COUNTERS
            .overflowed
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let baseline = super::COUNTERS.snapshot();
        let mut region = super::begin();

        let staging = region.split().expect("zero staging segment");
        assert_eq!(staging.status, Status::Measured);
        assert_eq!(staging.allocation_calls, Some(0));
        assert_eq!(staging.deallocation_calls, Some(0));
        assert_eq!(staging.live_bytes_before, Some(baseline.live_bytes));
        assert_eq!(staging.live_bytes_after, Some(baseline.live_bytes));
        assert_eq!(staging.region_peak_live_bytes, Some(baseline.live_bytes));

        let (commit_core, combined) = region.finish_split();
        let commit_core = commit_core.expect("zero commit-core segment");
        let combined = combined.expect("zero combined segment");
        assert_eq!(commit_core.status, Status::Measured);
        assert_eq!(commit_core.allocation_calls, Some(0));
        assert_eq!(commit_core.live_bytes_before, Some(baseline.live_bytes));
        assert_eq!(commit_core.live_bytes_after, Some(baseline.live_bytes));
        assert_eq!(
            commit_core.region_peak_live_bytes,
            Some(baseline.live_bytes)
        );
        assert_eq!(combined.status, Status::Measured);
        assert_eq!(combined.allocation_calls, Some(0));
        assert_eq!(combined.live_bytes_before, Some(baseline.live_bytes));
        assert_eq!(combined.live_bytes_after, Some(baseline.live_bytes));
        assert_eq!(combined.region_peak_live_bytes, Some(baseline.live_bytes));

        let mut repeated = super::begin();
        assert_eq!(repeated.split().unwrap().status, Status::Measured);
        assert_eq!(repeated.split().unwrap().status, Status::Measured);
        assert_eq!(repeated.finish().unwrap().status, Status::Measured);
        let mut dropped = super::begin();
        assert_eq!(dropped.split().unwrap().status, Status::Measured);
        drop(dropped);
        assert_eq!(super::begin().finish().unwrap().status, Status::Measured);

        super::COUNTERS
            .overflowed
            .store(was_overflowed, std::sync::atomic::Ordering::SeqCst);
        super::COUNTERS
            .observer_invalid
            .store(was_invalid, std::sync::atomic::Ordering::SeqCst);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn split_phase_counter_faults_remain_overflow_without_partial_values() {
        fn assert_staging_fault<F>(inject: F)
        where
            F: FnOnce(&Counters),
        {
            let counters = Counters::default();
            let before = counters.begin_region().expect("fault test region");
            inject(&counters);
            let (middle, middle_peak, middle_valid) = counters.split_region();
            assert!(middle_valid);
            let staging = Sample::measured(before, middle, middle_peak);
            assert_eq!(staging.status, Status::Overflow);
            assert!(staging.allocation_calls.is_none());

            let (after, final_peak, final_valid) = counters.finish_region();
            assert!(final_valid);
            let commit_core = Sample::measured(middle, after, final_peak);
            let combined = Sample::measured(before, after, final_peak);
            assert_eq!(commit_core.status, Status::Overflow);
            assert_eq!(combined.status, Status::Overflow);
            assert!(commit_core.allocated_bytes.is_none());
            assert!(combined.region_peak_live_bytes.is_none());
        }

        fn assert_core_fault<F>(inject: F)
        where
            F: FnOnce(&Counters),
        {
            let counters = Counters::default();
            let before = counters.begin_region().expect("core fault test region");
            let (middle, middle_peak, middle_valid) = counters.split_region();
            assert!(middle_valid);
            let staging = Sample::measured(before, middle, middle_peak);
            assert_eq!(staging.status, Status::Measured);

            inject(&counters);
            let (after, final_peak, final_valid) = counters.finish_region();
            assert!(final_valid);
            let commit_core = Sample::measured(middle, after, final_peak);
            let combined = Sample::measured(before, after, final_peak);
            assert_eq!(commit_core.status, Status::Overflow);
            assert_eq!(combined.status, Status::Overflow);
            assert!(commit_core.allocation_calls.is_none());
            assert!(combined.allocated_bytes.is_none());
        }

        assert_staging_fault(|counters| {
            counters
                .allocation_calls
                .store(u64::MAX, std::sync::atomic::Ordering::Relaxed);
            counters.allocation(1);
        });
        assert_staging_fault(|counters| counters.live_sub(1));

        assert_core_fault(|counters| {
            counters
                .allocation_calls
                .store(u64::MAX, std::sync::atomic::Ordering::Relaxed);
            counters.allocation(1);
        });
        assert_core_fault(|counters| counters.live_sub(1));

        let before = super::CountersSnapshot {
            allocation_calls: 1,
            ..super::CountersSnapshot::default()
        };
        let after = super::CountersSnapshot::default();
        let staging_reversed = Sample::measured(before, after, Some(0));
        assert_eq!(staging_reversed.status, Status::Overflow);
        assert!(staging_reversed.deallocation_calls.is_none());

        let before = super::CountersSnapshot::default();
        let middle = super::CountersSnapshot {
            allocation_calls: 1,
            ..super::CountersSnapshot::default()
        };
        let core_reversed = Sample::measured(middle, before, Some(0));
        assert_eq!(core_reversed.status, Status::Overflow);
        assert!(core_reversed.deallocation_calls.is_none());
    }

    #[test]
    fn split_boundary_observer_poison_and_reentry_remain_unavailable() {
        fn assert_unavailable(sample: Sample) {
            assert_eq!(sample.status, Status::Unavailable);
            assert!(sample.allocation_calls.is_none());
            assert!(sample.region_peak_live_bytes.is_none());
        }

        let counters = std::sync::Arc::new(Counters::default());
        let before = counters.begin_region().expect("poison test region");
        let poisoned = std::sync::Arc::clone(&counters);
        let handle = std::thread::spawn(move || {
            let _observer = poisoned.observer.lock().unwrap();
            panic!("deliberately poison the split observer mutex");
        });
        assert!(handle.join().is_err());
        let (middle, middle_peak, middle_valid) = counters.split_region();
        assert!(!middle_valid);
        assert_unavailable(Sample::measured(before, middle, middle_peak));
        let (after, final_peak, final_valid) = counters.finish_region();
        assert!(!final_valid);
        assert_unavailable(Sample::measured(middle, after, final_peak));
        assert_unavailable(Sample::measured(before, after, final_peak));

        let counters = Counters::default();
        let before = counters.begin_region().expect("reentry test region");
        let entry = counters.enter_callback().expect("outer callback entry");
        counters.allocation(8);
        drop(entry);
        let (middle, middle_peak, middle_valid) = counters.split_region();
        assert!(!middle_valid);
        assert_unavailable(Sample::measured(before, middle, middle_peak));
        let (after, final_peak, final_valid) = counters.finish_region();
        assert!(!final_valid);
        assert_unavailable(Sample::measured(middle, after, final_peak));
    }

    #[test]
    fn missing_split_peak_is_sticky_overflow_for_both_finish_paths() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let was_invalid = super::COUNTERS
            .observer_invalid
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let was_overflowed = super::COUNTERS
            .overflowed
            .swap(false, std::sync::atomic::Ordering::SeqCst);

        let mut split_finish = super::begin();
        {
            let mut observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes = None;
        }
        assert_eq!(split_finish.split().unwrap().status, Status::Overflow);
        let (core, combined) = split_finish.finish_split();
        assert_eq!(core.unwrap().status, Status::Overflow);
        assert_eq!(combined.unwrap().status, Status::Overflow);

        super::COUNTERS
            .overflowed
            .store(false, std::sync::atomic::Ordering::SeqCst);
        let mut ordinary_finish = super::begin();
        {
            let mut observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes = None;
        }
        assert_eq!(ordinary_finish.split().unwrap().status, Status::Overflow);
        assert_eq!(ordinary_finish.finish().unwrap().status, Status::Overflow);

        super::COUNTERS
            .overflowed
            .store(false, std::sync::atomic::Ordering::SeqCst);
        let mut invalid_after_overflow = super::begin();
        {
            let mut observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes = None;
        }
        assert_eq!(
            invalid_after_overflow.split().unwrap().status,
            Status::Overflow
        );
        super::COUNTERS
            .observer_invalid
            .store(true, std::sync::atomic::Ordering::SeqCst);
        let invalid_sample = invalid_after_overflow.finish().unwrap();
        assert_eq!(invalid_sample.status, Status::Unavailable);
        assert!(invalid_sample.region_peak_live_bytes.is_none());
        super::COUNTERS
            .observer_invalid
            .store(false, std::sync::atomic::Ordering::SeqCst);
        super::COUNTERS
            .overflowed
            .store(false, std::sync::atomic::Ordering::SeqCst);

        let mut missing_final = super::begin();
        assert_eq!(missing_final.split().unwrap().status, Status::Measured);
        {
            let mut observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes = None;
        }
        let (core, combined) = missing_final.finish_split();
        assert_eq!(core.unwrap().status, Status::Overflow);
        assert_eq!(combined.unwrap().status, Status::Overflow);

        super::COUNTERS
            .overflowed
            .store(was_overflowed, std::sync::atomic::Ordering::SeqCst);
        super::COUNTERS
            .observer_invalid
            .store(was_invalid, std::sync::atomic::Ordering::SeqCst);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn dropping_a_split_region_after_an_error_releases_the_token() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let was_invalid = super::COUNTERS
            .observer_invalid
            .swap(false, std::sync::atomic::Ordering::SeqCst);
        let was_overflowed = super::COUNTERS
            .overflowed
            .swap(false, std::sync::atomic::Ordering::SeqCst);

        fn fail_after_staging() -> Result<(), &'static str> {
            let mut region = super::begin();
            assert_eq!(region.split().unwrap().status, Status::Measured);
            Err("simulated edit.set or edit.commit failure")
        }

        assert!(fail_after_staging().is_err());
        assert_eq!(super::begin().finish().unwrap().status, Status::Measured);

        super::COUNTERS
            .overflowed
            .store(was_overflowed, std::sync::atomic::Ordering::SeqCst);
        super::COUNTERS
            .observer_invalid
            .store(was_invalid, std::sync::atomic::Ordering::SeqCst);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }

    #[test]
    fn concurrent_begin_race_grants_one_region_and_marks_others_unavailable() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let initially_active = {
            let observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes.is_some()
        };
        assert!(!initially_active);
        let barrier = Arc::new(Barrier::new(8));
        let handles = (0..8)
            .map(|_| {
                let barrier = Arc::clone(&barrier);
                std::thread::spawn(move || {
                    barrier.wait();
                    let region = super::begin();
                    barrier.wait();
                    region.finish().map(|sample| sample.status)
                })
            })
            .collect::<Vec<_>>();
        let statuses = handles
            .into_iter()
            .map(|handle| handle.join().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(
            statuses
                .iter()
                .filter(|status| **status == Some(Status::Measured))
                .count(),
            1
        );
        assert_eq!(
            statuses
                .iter()
                .filter(|status| **status == Some(Status::Unavailable))
                .count(),
            7
        );
        let active_after = {
            let observer = super::COUNTERS.lock_observer();
            observer.region_peak_live_bytes.is_some()
        };
        assert!(!active_after);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }
}
