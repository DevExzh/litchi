//! Process-local allocation and provenance helpers for the PPTX profile.
//!
//! The allocator is deliberately boring: it forwards every operation to the
//! system allocator and records successful byte movement in atomics. The
//! runner invokes one binary process per process sample, so these counters do
//! not cross process boundaries.

#![allow(dead_code)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Duration;

use serde::Serialize;
use sha2::{Digest, Sha256};

pub struct CountingAllocator;

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: each method forwards the caller's valid allocator contract to the
// system allocator and only observes successful operations.
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
    loop {
        let before = LIVE_BYTES.load(Ordering::Acquire);
        if before < size {
            INVALID.store(true, Ordering::Release);
            LIVE_BYTES.store(0, Ordering::Release);
            return;
        }
        if LIVE_BYTES
            .compare_exchange_weak(before, before - size, Ordering::AcqRel, Ordering::Acquire)
            .is_ok()
        {
            return;
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub struct AllocSnapshot {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live: u64,
    peak: u64,
    failed: u64,
    invalid: bool,
}

impl AllocSnapshot {
    /// Start a phase with a local peak baseline at the currently retained live
    /// set. This keeps the phase peak independent of earlier warmups and
    /// phases while preserving the absolute live equation.
    pub fn phase_start() -> Self {
        let live = LIVE_BYTES.load(Ordering::Acquire);
        PEAK_BYTES.store(live, Ordering::Release);
        Self::now()
    }

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
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: INVALID.load(Ordering::Acquire),
        }
    }

    pub fn live_bytes(self) -> u64 {
        self.live
    }

    pub fn delta(self, after: Self, elapsed: Duration) -> PhaseRecord {
        let direct = after.direct.saturating_sub(self.direct);
        let realloc_old = after.realloc_old.saturating_sub(self.realloc_old);
        let realloc_new = after.realloc_new.saturating_sub(self.realloc_new);
        let deallocated = after.deallocated.saturating_sub(self.deallocated);
        let live_delta = signed_delta(after.live, self.live);
        PhaseRecord {
            elapsed_ns: elapsed.as_nanos().min(u128::from(u64::MAX)) as u64,
            requested_alloc_bytes: direct.saturating_add(realloc_new),
            direct_allocated_bytes: direct,
            realloc_new_bytes: realloc_new,
            realloc_old_bytes: realloc_old,
            deallocated_bytes: deallocated,
            live_before_bytes: self.live,
            live_after_bytes: after.live,
            live_delta_bytes: live_delta,
            retained_live_bytes_after: after.live,
            peak_live_during_bytes: after.peak,
            peak_live_delta_bytes: after.peak.saturating_sub(self.peak),
            alloc_balance_ok: balance(
                self.live,
                direct,
                realloc_new,
                realloc_old,
                deallocated,
                after.live,
            ),
            alloc_invalid: self.invalid || after.invalid,
            alloc_failed: after.failed.saturating_sub(self.failed),
            alloc_calls: after.calls.saturating_sub(self.calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
        }
    }
}

fn signed_delta(after: u64, before: u64) -> i64 {
    if after >= before {
        i64::try_from(after - before).unwrap_or(i64::MAX)
    } else {
        -i64::try_from(before - after).unwrap_or(i64::MAX)
    }
}

fn balance(
    before: u64,
    direct: u64,
    realloc_new: u64,
    realloc_old: u64,
    deallocated: u64,
    after: u64,
) -> bool {
    before
        .checked_add(direct)
        .and_then(|value| value.checked_add(realloc_new))
        .and_then(|value| value.checked_sub(realloc_old))
        .and_then(|value| value.checked_sub(deallocated))
        == Some(after)
}

#[derive(Clone, Copy, Debug, Serialize)]
pub struct PhaseRecord {
    pub elapsed_ns: u64,
    pub requested_alloc_bytes: u64,
    pub direct_allocated_bytes: u64,
    pub realloc_new_bytes: u64,
    pub realloc_old_bytes: u64,
    pub deallocated_bytes: u64,
    pub live_before_bytes: u64,
    pub live_after_bytes: u64,
    pub live_delta_bytes: i64,
    pub retained_live_bytes_after: u64,
    pub peak_live_during_bytes: u64,
    pub peak_live_delta_bytes: u64,
    pub alloc_balance_ok: bool,
    pub alloc_invalid: bool,
    pub alloc_failed: u64,
    pub alloc_calls: u64,
    pub realloc_calls: u64,
    pub dealloc_calls: u64,
}

pub fn reset_counters() {
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
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

pub fn fnv1a64(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

pub fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}
