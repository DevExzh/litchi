//! Operation-scoped allocator evidence for the ODG joined qualification.
//!
//! This file is copied into the temporary probe overlay and is compiled only
//! with `--cfg odg_joined_allocator`.  It is deliberately independent of the
//! performance workspace's serde-based allocator helper: the ODG probe has a
//! smaller dependency surface, and the counters are rendered through the
//! already-present `serde_json` dependency.  No normal probe binary includes
//! this module or installs this allocator.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

static ENABLED: AtomicBool = AtomicBool::new(false);
static ALLOCATION_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOCATION_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOCATION_CALLS: AtomicU64 = AtomicU64::new(0);
static FAILED_ALLOCATION_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_UNDERFLOW: AtomicBool = AtomicBool::new(false);
static REGION_ACTIVE: AtomicBool = AtomicBool::new(false);
static REGION_PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

struct CountingSystemAllocator;

// SAFETY: every pointer and layout is passed unchanged to the system
// allocator; metric updates are atomic side effects and do not alter the
// allocator contract.
unsafe impl GlobalAlloc for CountingSystemAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            if ENABLED.load(Ordering::Relaxed) {
                FAILED_ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            }
        } else if ENABLED.load(Ordering::Relaxed) {
            record_allocation(layout.size() as u64);
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            if ENABLED.load(Ordering::Relaxed) {
                FAILED_ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            }
        } else if ENABLED.load(Ordering::Relaxed) {
            record_allocation(layout.size() as u64);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the caller owns `pointer` with the supplied layout.
        unsafe { System.dealloc(pointer, layout) };
        if ENABLED.load(Ordering::Relaxed) {
            DEALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            let observed = decrement_live(layout.size() as u64).unwrap_or(0);
            DEALLOCATED_BYTES.fetch_add(observed, Ordering::Relaxed);
        }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies a valid pointer/layout pair.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            if ENABLED.load(Ordering::Relaxed) {
                FAILED_ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            }
        } else if ENABLED.load(Ordering::Relaxed) {
            REALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            // A successful realloc replaces one allocation with another.  It
            // is therefore also an allocation call for conservation and the
            // reallocation count is a subset of allocation calls.
            ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOCATED_BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
            let old_size = decrement_live(layout.size() as u64).unwrap_or(0);
            DEALLOCATED_BYTES.fetch_add(old_size, Ordering::Relaxed);
            let new_size = new_size as u64;
            // The old observed allocation was removed above; account for the
            // replacement allocation independently for both growth and shrink.
            record_live_delta(new_size);
        }
        result
    }
}

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingSystemAllocator = CountingSystemAllocator;

#[derive(Clone, Copy)]
struct Snapshot {
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_allocation_calls: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    live_underflow: bool,
}

fn snapshot() -> Snapshot {
    Snapshot {
        allocation_calls: ALLOCATION_CALLS.load(Ordering::Acquire),
        deallocation_calls: DEALLOCATION_CALLS.load(Ordering::Acquire),
        reallocation_calls: REALLOCATION_CALLS.load(Ordering::Acquire),
        failed_allocation_calls: FAILED_ALLOCATION_CALLS.load(Ordering::Acquire),
        allocated_bytes: ALLOCATED_BYTES.load(Ordering::Acquire),
        deallocated_bytes: DEALLOCATED_BYTES.load(Ordering::Acquire),
        live_bytes: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_bytes: PEAK_LIVE_BYTES.load(Ordering::Acquire),
        live_underflow: LIVE_UNDERFLOW.load(Ordering::Acquire),
    }
}

fn record_live_delta(delta: u64) {
    let live = LIVE_BYTES
        .fetch_add(delta, Ordering::AcqRel)
        .saturating_add(delta);
    let mut peak = PEAK_LIVE_BYTES.load(Ordering::Acquire);
    while live > peak {
        match PEAK_LIVE_BYTES.compare_exchange_weak(peak, live, Ordering::AcqRel, Ordering::Acquire)
        {
            Ok(_) => break,
            Err(observed) => peak = observed,
        }
    }
    if REGION_ACTIVE.load(Ordering::Acquire) {
        let mut region_peak = REGION_PEAK_LIVE_BYTES.load(Ordering::Acquire);
        while live > region_peak {
            match REGION_PEAK_LIVE_BYTES.compare_exchange_weak(
                region_peak,
                live,
                Ordering::AcqRel,
                Ordering::Acquire,
            ) {
                Ok(_) => break,
                Err(observed) => region_peak = observed,
            }
        }
    }
}

fn decrement_live(delta: u64) -> Option<u64> {
    let mut current = LIVE_BYTES.load(Ordering::Acquire);
    loop {
        if current < delta {
            LIVE_UNDERFLOW.store(true, Ordering::Release);
            return None;
        }
        let next = current - delta;
        match LIVE_BYTES.compare_exchange_weak(current, next, Ordering::AcqRel, Ordering::Acquire) {
            Ok(_) => return Some(delta),
            Err(observed) => current = observed,
        }
    }
}

fn record_allocation(bytes: u64) {
    ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
    ALLOCATED_BYTES.fetch_add(bytes, Ordering::Relaxed);
    record_live_delta(bytes);
}

pub fn enable() {
    ENABLED.store(true, Ordering::Release);
}

pub struct Region {
    before: Snapshot,
}

pub fn begin() -> Region {
    REGION_ACTIVE.store(true, Ordering::Release);
    let before = snapshot();
    REGION_PEAK_LIVE_BYTES.store(before.live_bytes, Ordering::Release);
    Region { before }
}

impl Region {
    pub fn finish(self) -> serde_json::Value {
        let after = snapshot();
        let region_peak_live_bytes = REGION_PEAK_LIVE_BYTES.load(Ordering::Acquire);
        REGION_ACTIVE.store(false, Ordering::Release);
        let measured = |before: u64, after: u64| after.checked_sub(before);
        let status = if measured(self.before.allocation_calls, after.allocation_calls).is_some()
            && measured(self.before.deallocation_calls, after.deallocation_calls).is_some()
            && measured(self.before.reallocation_calls, after.reallocation_calls).is_some()
            && measured(
                self.before.failed_allocation_calls,
                after.failed_allocation_calls,
            )
            .is_some()
            && measured(self.before.allocated_bytes, after.allocated_bytes).is_some()
            && measured(self.before.deallocated_bytes, after.deallocated_bytes).is_some()
            && after
                .reallocation_calls
                .saturating_sub(self.before.reallocation_calls)
                <= after
                    .allocation_calls
                    .saturating_sub(self.before.allocation_calls)
            && after.peak_live_bytes >= self.before.peak_live_bytes
            && region_peak_live_bytes <= after.peak_live_bytes
            && region_peak_live_bytes >= self.before.live_bytes
            && region_peak_live_bytes >= after.live_bytes
            && !after.live_underflow
        {
            "measured"
        } else {
            "overflow"
        };
        let allocation = serde_json::json!({
            "status": status,
            "scope": "operation_global_system_allocator",
            "allocation_calls": measured(self.before.allocation_calls, after.allocation_calls),
            "deallocation_calls": measured(self.before.deallocation_calls, after.deallocation_calls),
            "reallocation_calls": measured(self.before.reallocation_calls, after.reallocation_calls),
            "failed_allocation_calls": measured(
                self.before.failed_allocation_calls,
                after.failed_allocation_calls,
            ),
            "allocated_bytes": measured(self.before.allocated_bytes, after.allocated_bytes),
            "deallocated_bytes": measured(self.before.deallocated_bytes, after.deallocated_bytes),
            "live_bytes_before": self.before.live_bytes,
            "live_bytes_after": after.live_bytes,
            "peak_live_bytes_before": self.before.peak_live_bytes,
            "peak_live_bytes_after": after.peak_live_bytes,
            "region_peak_live_bytes": region_peak_live_bytes,
            "live_underflow": after.live_underflow,
        });
        serde_json::json!({
            "tool": {
                "binary": "litchi-perf-odg-joined-qualification-alloc",
                "instrumentation": "system_allocator_operation_scoped",
                "allocator_counter_revision": "serialized_region_peak_v3",
            },
            "allocation": allocation,
            "latency_claim": "allocator_instrumented_elapsed_not_latency_claim",
        })
    }
}
