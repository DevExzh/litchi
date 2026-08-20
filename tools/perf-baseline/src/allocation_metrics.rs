//! Forbid-safe operation-scoped allocation counters shared by the two targets.
//!
//! The normal binary never installs a global allocator wrapper and never calls
//! [`enable`]. The allocator binary owns the only `GlobalAlloc` implementation
//! and calls the record functions below after each successful or failed system
//! allocation. Keeping this state and all arithmetic safe lets the shared
//! harness library retain `#![forbid(unsafe_code)]`.
//!
//! Counters are absolute process counters. A region records two snapshots and
//! publishes checked differences; it never resets a counter. The region guard
//! is deliberately non-reentrant so a nested operation cannot publish a
//! misleading partial interval. Atomics make totals include allocations made
//! by worker threads while the operation is active.

use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use serde::{Deserialize, Serialize};

const SCOPE: Scope = Scope::OperationGlobalSystemAllocator;

static ENABLED: AtomicBool = AtomicBool::new(false);
static REGION_ACTIVE: AtomicBool = AtomicBool::new(false);

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
        }
    }

    fn measured(before: CountersSnapshot, after: CountersSnapshot) -> Self {
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
        let overflow = before.overflowed
            || after.overflowed
            || values.iter().any(Option::is_none)
            || after.peak_live_bytes < before.peak_live_bytes;
        if overflow {
            return Self {
                status: Status::Overflow,
                ..Self::unavailable()
            };
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
        }
    }
}

enum RegionState {
    Disabled,
    Unavailable,
    Active(CountersSnapshot),
}

/// A region that owns the non-overlap token until it is finished or dropped.
pub(crate) struct Region {
    state: RegionState,
}

impl Region {
    /// Ends the region and releases its token. No allocator counters are
    /// reset, and this method performs no heap allocation itself.
    pub(crate) fn finish(mut self) -> Option<Sample> {
        match self.state {
            RegionState::Disabled => None,
            RegionState::Unavailable => Some(Sample::unavailable()),
            RegionState::Active(before) => {
                let after = COUNTERS.snapshot();
                self.state = RegionState::Disabled;
                REGION_ACTIVE.store(false, Ordering::Release);
                Some(Sample::measured(before, after))
            },
        }
    }
}

impl Drop for Region {
    fn drop(&mut self) {
        if matches!(self.state, RegionState::Active(_)) {
            REGION_ACTIVE.store(false, Ordering::Release);
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
    if REGION_ACTIVE
        .compare_exchange(false, true, Ordering::Acquire, Ordering::Relaxed)
        .is_err()
    {
        return Region {
            state: RegionState::Unavailable,
        };
    }
    Region {
        state: RegionState::Active(COUNTERS.snapshot()),
    }
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
}

impl Counters {
    fn snapshot(&self) -> CountersSnapshot {
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
        }
    }

    fn add(&self, counter: &AtomicU64, value: u64) {
        if checked_add(counter, value).is_none() {
            self.overflowed.store(true, Ordering::Release);
        }
    }

    fn live_add(&self, value: u64) {
        if let Some(live) = checked_add(&self.live_bytes, value) {
            self.update_peak(live);
        } else {
            self.overflowed.store(true, Ordering::Release);
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
        let Some(size) = u64::try_from(size).ok() else {
            self.overflowed.store(true, Ordering::Release);
            return;
        };
        self.add(&self.allocation_calls, 1);
        self.add(&self.allocated_bytes, size);
        self.live_add(size);
    }

    fn deallocation(&self, size: usize) {
        let Some(size) = u64::try_from(size).ok() else {
            self.overflowed.store(true, Ordering::Release);
            return;
        };
        self.add(&self.deallocation_calls, 1);
        self.add(&self.deallocated_bytes, size);
        self.live_sub(size);
    }

    fn reallocation(&self, old_size: usize, new_size: usize) {
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
            self.live_add(new_size - old_size);
        } else {
            self.live_sub(old_size - new_size);
        }
    }

    fn failed_allocation(&self) {
        self.add(&self.failed_allocation_calls, 1);
    }
}

fn checked_add(counter: &AtomicU64, value: u64) -> Option<u64> {
    counter
        .fetch_update(Ordering::AcqRel, Ordering::Acquire, |current| {
            current.checked_add(value)
        })
        .ok()
}

fn checked_sub(counter: &AtomicU64, value: u64) -> Option<u64> {
    counter
        .fetch_update(Ordering::AcqRel, Ordering::Acquire, |current| {
            current.checked_sub(value)
        })
        .ok()
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
};

fn difference(before: u64, after: u64) -> Option<u64> {
    after.checked_sub(before)
}

#[cfg(test)]
mod tests {
    use std::sync::{Arc, Barrier};

    use super::{Counters, Sample, Scope, Status};

    static TEST_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn disabled_region_publishes_no_sample() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(false, std::sync::atomic::Ordering::SeqCst);
        assert!(super::begin().finish().is_none());
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
        let sample = Sample::measured(before, after);
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
    fn concurrent_begin_race_grants_one_region_and_marks_others_unavailable() {
        let _lock = TEST_LOCK.lock().unwrap();
        let was_enabled = super::ENABLED.swap(true, std::sync::atomic::Ordering::SeqCst);
        let was_active = super::REGION_ACTIVE.swap(false, std::sync::atomic::Ordering::SeqCst);
        assert!(!was_active, "test region must not already be active");
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
        assert!(!super::REGION_ACTIVE.load(std::sync::atomic::Ordering::SeqCst));
        super::REGION_ACTIVE.store(was_active, std::sync::atomic::Ordering::SeqCst);
        super::ENABLED.store(was_enabled, std::sync::atomic::Ordering::SeqCst);
    }
}
