//! Process-local allocation accounting for the matched performance profile.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    reason = "This standalone evidence binary owns a process-local allocator observer."
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

pub struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: each method forwards the caller's allocator contract to `System`;
// the atomics only observe successful operations and do not alter ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair is supplied by the allocator caller.
        unsafe { System.dealloc(pointer, layout) };
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the pointer/layout pair and new size satisfy GlobalAlloc's contract.
        let replacement = unsafe { System.realloc(pointer, layout, new_size) };
        if replacement.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            REALLOC_OLD_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            REALLOC_NEW_BYTES.fetch_add(as_u64(new_size), Ordering::Relaxed);
            if new_size >= layout.size() {
                observe_growth(new_size - layout.size());
            } else {
                subtract_live(layout.size() - new_size);
            }
        }
        replacement
    }
}

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn observe_alloc(size: usize) {
    ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
    DIRECT_BYTES.fetch_add(as_u64(size), Ordering::Relaxed);
    observe_growth(size);
}

fn observe_growth(size: usize) {
    let size = as_u64(size);
    let live = LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    let mut old = PEAK_BYTES.load(Ordering::Relaxed);
    while live > old {
        match PEAK_BYTES.compare_exchange_weak(old, live, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => break,
            Err(observed) => old = observed,
        }
    }
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        INVALID.store(true, Ordering::Release);
    }
}

#[derive(Clone, Copy, Debug)]
pub struct Snapshot {
    pub allocation_calls: u64,
    pub reallocation_calls: u64,
    pub deallocation_calls: u64,
    pub direct_bytes: u64,
    pub realloc_old_bytes: u64,
    pub realloc_new_bytes: u64,
    pub deallocated_bytes: u64,
    pub live_bytes: u64,
    pub peak_bytes: u64,
    pub allocation_failed: u64,
    pub invalid: bool,
}

impl Snapshot {
    #[must_use]
    pub fn now() -> Self {
        Self {
            allocation_calls: ALLOC_CALLS.load(Ordering::Acquire),
            reallocation_calls: REALLOC_CALLS.load(Ordering::Acquire),
            deallocation_calls: DEALLOC_CALLS.load(Ordering::Acquire),
            direct_bytes: DIRECT_BYTES.load(Ordering::Acquire),
            realloc_old_bytes: REALLOC_OLD_BYTES.load(Ordering::Acquire),
            realloc_new_bytes: REALLOC_NEW_BYTES.load(Ordering::Acquire),
            deallocated_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
            live_bytes: LIVE_BYTES.load(Ordering::Acquire),
            peak_bytes: PEAK_BYTES.load(Ordering::Acquire),
            allocation_failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: INVALID.load(Ordering::Acquire),
        }
    }

    #[must_use]
    pub fn delta(self, after: Self) -> Delta {
        Delta {
            allocation_calls: after.allocation_calls.saturating_sub(self.allocation_calls),
            reallocation_calls: after
                .reallocation_calls
                .saturating_sub(self.reallocation_calls),
            deallocation_calls: after
                .deallocation_calls
                .saturating_sub(self.deallocation_calls),
            direct_bytes: after.direct_bytes.saturating_sub(self.direct_bytes),
            realloc_old_bytes: after
                .realloc_old_bytes
                .saturating_sub(self.realloc_old_bytes),
            realloc_new_bytes: after
                .realloc_new_bytes
                .saturating_sub(self.realloc_new_bytes),
            deallocated_bytes: after
                .deallocated_bytes
                .saturating_sub(self.deallocated_bytes),
            live_before: self.live_bytes,
            live_after: after.live_bytes,
            peak_live_absolute: after.peak_bytes,
            peak_live_delta: after.peak_bytes.saturating_sub(self.live_bytes),
            allocation_failed: after
                .allocation_failed
                .saturating_sub(self.allocation_failed),
            invalid: self.invalid || after.invalid,
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub struct Delta {
    pub allocation_calls: u64,
    pub reallocation_calls: u64,
    pub deallocation_calls: u64,
    pub direct_bytes: u64,
    pub realloc_old_bytes: u64,
    pub realloc_new_bytes: u64,
    pub deallocated_bytes: u64,
    pub live_before: u64,
    pub live_after: u64,
    pub peak_live_absolute: u64,
    pub peak_live_delta: u64,
    pub allocation_failed: u64,
    pub invalid: bool,
}

impl Delta {
    #[must_use]
    pub fn requested_bytes(self) -> u64 {
        self.direct_bytes.saturating_add(self.realloc_new_bytes)
    }

    #[must_use]
    pub fn balanced(self) -> bool {
        self.live_before
            .checked_add(self.direct_bytes)
            .and_then(|value| value.checked_add(self.realloc_new_bytes))
            .and_then(|value| value.checked_sub(self.realloc_old_bytes))
            .and_then(|value| value.checked_sub(self.deallocated_bytes))
            == Some(self.live_after)
    }
}

pub fn reset() {
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
    ALLOC_CALLS.store(0, Ordering::Release);
    REALLOC_CALLS.store(0, Ordering::Release);
    DEALLOC_CALLS.store(0, Ordering::Release);
    DIRECT_BYTES.store(0, Ordering::Release);
    REALLOC_OLD_BYTES.store(0, Ordering::Release);
    REALLOC_NEW_BYTES.store(0, Ordering::Release);
    DEALLOC_BYTES.store(0, Ordering::Release);
    ALLOC_FAILED.store(0, Ordering::Release);
    INVALID.store(false, Ordering::Release);
}
