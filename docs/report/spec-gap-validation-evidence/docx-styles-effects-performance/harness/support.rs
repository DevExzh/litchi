//! Process-local allocation accounting for the DOCX SVG lifecycle profile.
//!
//! The observer deliberately reports allocation traffic, peak live bytes, and
//! process RSS as separate quantities.  It does not attempt to replace an
//! allocator profiler or to infer resident memory from Rust allocation data.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    reason = "the opt-in evidence harness owns a process-local GlobalAlloc observer"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

pub struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
// This peak is reset only at a subphase boundary.  The outer peak remains in
// PEAK_BYTES for the complete measured operation, so subphase attribution
// cannot erase the operation-level peak receipt.
static PHASE_PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: each method forwards the allocator contract to `System`; the
// atomics observe successful operations without changing pointer ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair belongs to the allocator caller.
        unsafe { System.dealloc(pointer, layout) };
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            REALLOC_OLD.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            REALLOC_NEW.fetch_add(as_u64(new_size), Ordering::Relaxed);
            if new_size >= layout.size() {
                observe_growth(new_size - layout.size());
            } else {
                subtract_live(layout.size() - new_size);
            }
        }
        result
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
    observe_peak(&PEAK_BYTES, live);
    observe_peak(&PHASE_PEAK_BYTES, live);
}

fn observe_peak(peak: &AtomicU64, live: u64) {
    let mut old = peak.load(Ordering::Relaxed);
    while live > old {
        match peak.compare_exchange_weak(old, live, Ordering::Relaxed, Ordering::Relaxed) {
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
pub struct AllocSnapshot {
    pub calls: u64,
    pub realloc_calls: u64,
    pub dealloc_calls: u64,
    pub direct: u64,
    pub realloc_old: u64,
    pub realloc_new: u64,
    pub deallocated: u64,
    pub live: u64,
    pub peak: u64,
    pub phase_peak: u64,
    pub failed: u64,
    pub invalid: bool,
}

impl AllocSnapshot {
    #[must_use]
    pub fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: REALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
            direct: DIRECT_BYTES.load(Ordering::Acquire),
            realloc_old: REALLOC_OLD.load(Ordering::Acquire),
            realloc_new: REALLOC_NEW.load(Ordering::Acquire),
            deallocated: DEALLOC_BYTES.load(Ordering::Acquire),
            live: LIVE_BYTES.load(Ordering::Acquire),
            peak: PEAK_BYTES.load(Ordering::Acquire),
            phase_peak: PHASE_PEAK_BYTES.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: INVALID.load(Ordering::Acquire),
        }
    }

    #[must_use]
    pub fn delta(self, after: Self) -> AllocDelta {
        AllocDelta {
            calls: after.calls.saturating_sub(self.calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            direct: after.direct.saturating_sub(self.direct),
            realloc_old: after.realloc_old.saturating_sub(self.realloc_old),
            realloc_new: after.realloc_new.saturating_sub(self.realloc_new),
            deallocated: after.deallocated.saturating_sub(self.deallocated),
            live_before: self.live,
            live_after: after.live,
            peak_delta: after.peak.saturating_sub(self.peak),
            failed: after.failed.saturating_sub(self.failed),
            invalid: self.invalid || after.invalid,
        }
    }

    /// Return a counter delta whose peak is bounded to the subphase window
    /// marked by [`begin_phase_window`].  All allocation traffic counters and
    /// the live-byte equation still use the same process-global snapshots as
    /// the outer operation window.
    #[must_use]
    pub fn phase_delta(self, after: Self) -> AllocDelta {
        let mut delta = self.delta(after);
        delta.peak_delta = after.phase_peak.saturating_sub(self.phase_peak);
        delta
    }
}

#[derive(Clone, Copy, Debug)]
pub struct AllocDelta {
    pub calls: u64,
    pub realloc_calls: u64,
    pub dealloc_calls: u64,
    pub direct: u64,
    pub realloc_old: u64,
    pub realloc_new: u64,
    pub deallocated: u64,
    pub live_before: u64,
    pub live_after: u64,
    pub peak_delta: u64,
    pub failed: u64,
    pub invalid: bool,
}

impl AllocDelta {
    #[must_use]
    pub fn requested(self) -> u64 {
        self.direct.saturating_add(self.realloc_new)
    }

    #[must_use]
    pub fn balanced(self) -> bool {
        self.live_before
            .checked_add(self.direct)
            .and_then(|value| value.checked_add(self.realloc_new))
            .and_then(|value| value.checked_sub(self.realloc_old))
            .and_then(|value| value.checked_sub(self.deallocated))
            == Some(self.live_after)
    }
}

/// Start an allocation window while retaining the process's live-byte base.
/// The baseline is necessary because fixture and runtime allocations exist
/// outside the timed operation.
pub fn begin_window() {
    let live = LIVE_BYTES.load(Ordering::Acquire);
    PEAK_BYTES.store(live, Ordering::Release);
    PHASE_PEAK_BYTES.store(live, Ordering::Release);
    ALLOC_CALLS.store(0, Ordering::Release);
    REALLOC_CALLS.store(0, Ordering::Release);
    DEALLOC_CALLS.store(0, Ordering::Release);
    DIRECT_BYTES.store(0, Ordering::Release);
    REALLOC_OLD.store(0, Ordering::Release);
    REALLOC_NEW.store(0, Ordering::Release);
    DEALLOC_BYTES.store(0, Ordering::Release);
    ALLOC_FAILED.store(0, Ordering::Release);
    INVALID.store(false, Ordering::Release);
}

/// Start an allocation subphase without resetting the outer operation's
/// counters or peak.  The live-byte baseline is retained so the subphase peak
/// reports only growth above the bytes live at its own start.
pub fn begin_phase_window() {
    PHASE_PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
}

/// Reset counters before any warm-up or measured sample.
pub fn reset_process_counters() {
    begin_window();
}
