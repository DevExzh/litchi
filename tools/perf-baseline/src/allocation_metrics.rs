//! Opt-in operation-scoped allocation counters for the benchmark binary.
//!
//! The normal `litchi-perf-baseline` target does not install the wrapper and
//! therefore pays no allocator instrumentation cost.  The companion
//! `litchi-perf-baseline-alloc` target enables the feature that installs one
//! process-global wrapper around `std::alloc::System`.
//!
//! Counters are absolute process counters.  A region records two snapshots and
//! publishes checked differences; it never resets a counter.  The region
//! guard is deliberately non-reentrant so a nested operation cannot publish a
//! misleading partial interval.  Atomics make totals include allocations made
//! by worker threads while the operation is active.

use serde::{Deserialize, Serialize};

pub(crate) const INSTRUMENTATION_IDENTITY: &str = if cfg!(feature = "allocator-metrics") {
    "system_allocator_operation_scoped"
} else {
    "none"
};

pub(crate) const ALLOCATOR_IDENTITY: &str = if cfg!(feature = "allocator-metrics") {
    "CountingSystemAllocator(std::alloc::System)"
} else {
    "Rust system allocator"
};

#[cfg(feature = "allocator-metrics")]
const SCOPE: &str = "operation_global_system_allocator";

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub(crate) enum Status {
    /// Every absolute counter difference was checked successfully.
    Measured,
    /// The region could not be acquired, or the feature is not active.
    Unavailable,
    /// An absolute counter or checked difference overflowed.
    Overflow,
}

/// One operation's allocation observation.  Numeric fields are omitted unless
/// the complete observation is measured; callers must not mix partial vectors.
#[derive(Clone, Debug, Deserialize, Serialize, PartialEq, Eq)]
pub(crate) struct Sample {
    pub status: Status,
    pub scope: String,
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
    #[cfg(feature = "allocator-metrics")]
    fn unavailable() -> Self {
        Self {
            status: Status::Unavailable,
            scope: SCOPE.to_owned(),
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

    #[cfg(feature = "allocator-metrics")]
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
            scope: SCOPE.to_owned(),
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

/// A region that owns the non-overlap token until it is finished or dropped.
pub(crate) struct Region {
    #[cfg(feature = "allocator-metrics")]
    before: CountersSnapshot,
    #[cfg(feature = "allocator-metrics")]
    active: bool,
    status: Status,
}

impl Region {
    /// Ends the region and releases its token.  No allocator counters are
    /// reset, and this method performs no heap allocation itself.
    pub(crate) fn finish(mut self) -> Option<Sample> {
        let _ = self.status;
        #[cfg(feature = "allocator-metrics")]
        {
            if !self.active {
                return Some(Sample::unavailable());
            }
            let after = COUNTERS.snapshot();
            self.active = false;
            REGION_ACTIVE.store(false, Ordering::SeqCst);
            return Some(Sample::measured(self.before, after));
        }
        None
    }
}

impl Drop for Region {
    fn drop(&mut self) {
        #[cfg(feature = "allocator-metrics")]
        if self.active {
            REGION_ACTIVE.store(false, Ordering::SeqCst);
            self.active = false;
        }
    }
}

/// Starts a region when operation instrumentation is enabled.  A failed
/// acquisition returns a region which publishes `unavailable` and leaves the
/// operation itself untouched.
pub(crate) fn begin() -> Region {
    #[cfg(feature = "allocator-metrics")]
    {
        if REGION_ACTIVE
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .is_err()
        {
            return Region {
                before: CountersSnapshot::default(),
                active: false,
                status: Status::Unavailable,
            };
        }
        return Region {
            before: COUNTERS.snapshot(),
            active: true,
            status: Status::Measured,
        };
    }
    Region {
        status: Status::Unavailable,
    }
}

#[cfg(feature = "allocator-metrics")]
use std::{
    alloc::{GlobalAlloc, Layout, System},
    sync::atomic::{AtomicBool, AtomicU64, Ordering},
};

#[cfg(feature = "allocator-metrics")]
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

#[cfg(feature = "allocator-metrics")]
#[derive(Clone, Copy, Default)]
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

#[cfg(feature = "allocator-metrics")]
impl Counters {
    fn snapshot(&self) -> CountersSnapshot {
        CountersSnapshot {
            allocation_calls: self.allocation_calls.load(Ordering::SeqCst),
            deallocation_calls: self.deallocation_calls.load(Ordering::SeqCst),
            reallocation_calls: self.reallocation_calls.load(Ordering::SeqCst),
            failed_allocation_calls: self.failed_allocation_calls.load(Ordering::SeqCst),
            allocated_bytes: self.allocated_bytes.load(Ordering::SeqCst),
            deallocated_bytes: self.deallocated_bytes.load(Ordering::SeqCst),
            live_bytes: self.live_bytes.load(Ordering::SeqCst),
            peak_live_bytes: self.peak_live_bytes.load(Ordering::SeqCst),
            overflowed: self.overflowed.load(Ordering::SeqCst),
        }
    }

    fn add(&self, counter: &AtomicU64, value: u64) {
        if checked_add(counter, value).is_none() {
            self.overflowed.store(true, Ordering::SeqCst);
        }
    }

    fn live_add(&self, value: u64) {
        if let Some(live) = checked_add(&self.live_bytes, value) {
            self.update_peak(live);
        } else {
            self.overflowed.store(true, Ordering::SeqCst);
        }
    }

    fn live_sub(&self, value: u64) {
        if checked_sub(&self.live_bytes, value).is_none() {
            self.overflowed.store(true, Ordering::SeqCst);
        }
    }

    fn update_peak(&self, live: u64) {
        let mut peak = self.peak_live_bytes.load(Ordering::SeqCst);
        while live > peak {
            match self.peak_live_bytes.compare_exchange(
                peak,
                live,
                Ordering::SeqCst,
                Ordering::SeqCst,
            ) {
                Ok(_) => break,
                Err(observed) => peak = observed,
            }
        }
    }

    fn allocation(&self, size: usize, zeroed: bool) {
        let Some(size) = u64::try_from(size).ok() else {
            self.overflowed.store(true, Ordering::SeqCst);
            return;
        };
        self.add(&self.allocation_calls, 1);
        self.add(&self.allocated_bytes, size);
        self.live_add(size);
        let _ = zeroed;
    }

    fn deallocation(&self, size: usize) {
        let Some(size) = u64::try_from(size).ok() else {
            self.overflowed.store(true, Ordering::SeqCst);
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
            self.overflowed.store(true, Ordering::SeqCst);
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

#[cfg(feature = "allocator-metrics")]
fn checked_add(counter: &AtomicU64, value: u64) -> Option<u64> {
    counter
        .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |current| {
            current.checked_add(value)
        })
        .ok()
}

#[cfg(feature = "allocator-metrics")]
fn checked_sub(counter: &AtomicU64, value: u64) -> Option<u64> {
    counter
        .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |current| {
            current.checked_sub(value)
        })
        .ok()
}

#[cfg(feature = "allocator-metrics")]
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

#[cfg(feature = "allocator-metrics")]
static REGION_ACTIVE: AtomicBool = AtomicBool::new(false);

#[cfg(feature = "allocator-metrics")]
struct CountingSystemAllocator;

#[cfg(feature = "allocator-metrics")]
// SAFETY: Every method delegates to `std::alloc::System`; the atomic updates
// are side-effect-only and do not change the pointer, layout, or ownership
// contract required by `GlobalAlloc`.
unsafe impl GlobalAlloc for CountingSystemAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: The caller upholds the `GlobalAlloc` layout contract.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            COUNTERS.failed_allocation();
        } else {
            COUNTERS.allocation(layout.size(), false);
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: The caller upholds the `GlobalAlloc` layout contract.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            COUNTERS.failed_allocation();
        } else {
            COUNTERS.allocation(layout.size(), true);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: The caller upholds the `GlobalAlloc` pointer/layout contract.
        unsafe { System.dealloc(pointer, layout) };
        COUNTERS.deallocation(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: The caller upholds the `GlobalAlloc` pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            COUNTERS.failed_allocation();
        } else {
            COUNTERS.reallocation(layout.size(), new_size);
        }
        result
    }
}

#[cfg(feature = "allocator-metrics")]
#[global_allocator]
static GLOBAL_ALLOCATOR: CountingSystemAllocator = CountingSystemAllocator;

#[cfg(feature = "allocator-metrics")]
fn difference(before: u64, after: u64) -> Option<u64> {
    after.checked_sub(before)
}

#[cfg(test)]
mod tests {
    use super::{Sample, Status};

    #[cfg(feature = "allocator-metrics")]
    static TEST_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn disabled_build_publishes_no_sample() {
        #[cfg(not(feature = "allocator-metrics"))]
        assert!(super::begin().finish().is_none());
    }

    #[test]
    fn sample_status_serialization_is_explicit() {
        let sample = Sample {
            status: Status::Overflow,
            scope: "operation_global_system_allocator".to_owned(),
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

    #[cfg(feature = "allocator-metrics")]
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

    #[cfg(feature = "allocator-metrics")]
    #[test]
    fn region_counts_cross_thread_totals_without_resetting_absolute_counters() {
        let _lock = TEST_LOCK.lock().unwrap();
        let before = super::COUNTERS.snapshot();
        let region = super::begin();
        let handles = (0..2)
            .map(|_| std::thread::spawn(|| vec![0_u8; 128]))
            .collect::<Vec<_>>();
        for handle in handles {
            drop(handle.join().unwrap());
        }
        let sample = region.finish().unwrap();
        let after = super::COUNTERS.snapshot();
        assert_eq!(sample.status, Status::Measured);
        assert!(sample.allocation_calls.unwrap() >= 2);
        assert_eq!(sample.live_bytes_before, Some(before.live_bytes));
        assert!(after.allocation_calls >= before.allocation_calls);
    }

    #[cfg(feature = "allocator-metrics")]
    #[test]
    fn region_guard_rejects_overlap() {
        let _lock = TEST_LOCK.lock().unwrap();
        let outer = super::begin();
        let inner = super::begin();
        assert_eq!(inner.finish().unwrap().status, Status::Unavailable);
        assert_eq!(outer.finish().unwrap().status, Status::Measured);
    }
}
