//! Isolated allocator observer for the paired DrawingML raw endpoint.
//!
//! The probe library stays `forbid(unsafe_code)` and receives only a safe
//! observer trait object.  This binary owns the one process-global allocator
//! wrapper and reports direct allocation, deallocation, and realloc old/new
//! byte deltas independently of the normal timing endpoint.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_perf_baseline::drawingml_raw_probe::{
    AllocationMetrics, AllocationObserver, AllocationToken,
};

#[derive(Default)]
struct Counters {
    allocation_calls: AtomicU64,
    deallocation_calls: AtomicU64,
    reallocation_calls: AtomicU64,
    failed_calls: AtomicU64,
    direct_bytes: AtomicU64,
    deallocated_bytes: AtomicU64,
    realloc_old_bytes: AtomicU64,
    realloc_new_bytes: AtomicU64,
    live_bytes: AtomicU64,
    peak_live_bytes: AtomicU64,
}

static COUNTERS: Counters = Counters {
    allocation_calls: AtomicU64::new(0),
    deallocation_calls: AtomicU64::new(0),
    reallocation_calls: AtomicU64::new(0),
    failed_calls: AtomicU64::new(0),
    direct_bytes: AtomicU64::new(0),
    deallocated_bytes: AtomicU64::new(0),
    realloc_old_bytes: AtomicU64::new(0),
    realloc_new_bytes: AtomicU64::new(0),
    live_bytes: AtomicU64::new(0),
    peak_live_bytes: AtomicU64::new(0),
};

pub static OBSERVER: DrawingMlObserver = DrawingMlObserver;

pub struct DrawingMlObserver;

impl AllocationObserver for DrawingMlObserver {
    fn begin(&self) -> AllocationToken {
        AllocationToken {
            allocation_calls: COUNTERS.allocation_calls.load(Ordering::Acquire),
            deallocation_calls: COUNTERS.deallocation_calls.load(Ordering::Acquire),
            reallocation_calls: COUNTERS.reallocation_calls.load(Ordering::Acquire),
            failed_calls: COUNTERS.failed_calls.load(Ordering::Acquire),
            direct_bytes: COUNTERS.direct_bytes.load(Ordering::Acquire),
            deallocated_bytes: COUNTERS.deallocated_bytes.load(Ordering::Acquire),
            realloc_old_bytes: COUNTERS.realloc_old_bytes.load(Ordering::Acquire),
            realloc_new_bytes: COUNTERS.realloc_new_bytes.load(Ordering::Acquire),
            live_bytes: COUNTERS.live_bytes.load(Ordering::Acquire),
            peak_live_bytes: COUNTERS.peak_live_bytes.load(Ordering::Acquire),
        }
    }

    fn finish(&self, before: AllocationToken) -> AllocationMetrics {
        let after = self.begin();
        AllocationMetrics {
            allocation_calls: delta(after.allocation_calls, before.allocation_calls),
            deallocation_calls: delta(after.deallocation_calls, before.deallocation_calls),
            reallocation_calls: delta(after.reallocation_calls, before.reallocation_calls),
            failed_calls: delta(after.failed_calls, before.failed_calls),
            direct_bytes: delta(after.direct_bytes, before.direct_bytes),
            deallocated_bytes: delta(after.deallocated_bytes, before.deallocated_bytes),
            realloc_old_bytes: delta(after.realloc_old_bytes, before.realloc_old_bytes),
            realloc_new_bytes: delta(after.realloc_new_bytes, before.realloc_new_bytes),
            live_before: before.live_bytes,
            live_after: after.live_bytes,
            peak_live_delta: after.peak_live_bytes.saturating_sub(before.peak_live_bytes),
            conservation_ok: conservation(before, after),
            elapsed_claimable: false,
        }
    }
}

fn delta(after: u64, before: u64) -> u64 {
    after.saturating_sub(before)
}

fn conservation(before: AllocationToken, after: AllocationToken) -> bool {
    let direct = delta(after.direct_bytes, before.direct_bytes) as i128;
    let deallocated = delta(after.deallocated_bytes, before.deallocated_bytes) as i128;
    let old = delta(after.realloc_old_bytes, before.realloc_old_bytes) as i128;
    let new = delta(after.realloc_new_bytes, before.realloc_new_bytes) as i128;
    let live = after.live_bytes as i128 - before.live_bytes as i128;
    live == direct + new - deallocated - old
}

unsafe impl GlobalAlloc for DrawingMlAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            COUNTERS.failed_calls.fetch_add(1, Ordering::AcqRel);
        } else {
            COUNTERS.allocation_calls.fetch_add(1, Ordering::AcqRel);
            COUNTERS
                .direct_bytes
                .fetch_add(layout.size() as u64, Ordering::AcqRel);
            add_live(layout.size() as u64);
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            COUNTERS.failed_calls.fetch_add(1, Ordering::AcqRel);
        } else {
            COUNTERS.allocation_calls.fetch_add(1, Ordering::AcqRel);
            COUNTERS
                .direct_bytes
                .fetch_add(layout.size() as u64, Ordering::AcqRel);
            add_live(layout.size() as u64);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        COUNTERS.deallocation_calls.fetch_add(1, Ordering::AcqRel);
        COUNTERS
            .deallocated_bytes
            .fetch_add(layout.size() as u64, Ordering::AcqRel);
        COUNTERS
            .live_bytes
            .fetch_sub(layout.size() as u64, Ordering::AcqRel);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            COUNTERS.failed_calls.fetch_add(1, Ordering::AcqRel);
        } else {
            COUNTERS.allocation_calls.fetch_add(1, Ordering::AcqRel);
            COUNTERS.reallocation_calls.fetch_add(1, Ordering::AcqRel);
            COUNTERS
                .realloc_old_bytes
                .fetch_add(layout.size() as u64, Ordering::AcqRel);
            COUNTERS
                .realloc_new_bytes
                .fetch_add(new_size as u64, Ordering::AcqRel);
            if new_size >= layout.size() {
                add_live((new_size - layout.size()) as u64);
            } else {
                COUNTERS
                    .live_bytes
                    .fetch_sub((layout.size() - new_size) as u64, Ordering::AcqRel);
            }
        }
        result
    }
}

fn add_live(size: u64) {
    let live = COUNTERS.live_bytes.fetch_add(size, Ordering::AcqRel) + size;
    let mut peak = COUNTERS.peak_live_bytes.load(Ordering::Acquire);
    while live > peak {
        match COUNTERS.peak_live_bytes.compare_exchange(
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

struct DrawingMlAllocator;

#[global_allocator]
static GLOBAL_ALLOCATOR: DrawingMlAllocator = DrawingMlAllocator;
