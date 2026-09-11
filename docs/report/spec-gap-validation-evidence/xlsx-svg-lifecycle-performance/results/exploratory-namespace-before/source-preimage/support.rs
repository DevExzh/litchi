//! Reusable process-local profile support for the future XLSX API adapter.
//!
//! The adapter uses these helpers around the current public XLSX API. The
//! shell runner keeps collection closed until production semantic checks and
//! the profile freeze have completed.

#![allow(dead_code)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

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

// SAFETY: each method forwards the valid allocator contract to `System`; the
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
    pub failed: u64,
    pub invalid: bool,
}

impl AllocSnapshot {
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
    pub fn requested(self) -> u64 {
        self.direct.saturating_add(self.realloc_new)
    }

    pub fn balanced(self) -> bool {
        self.live_before
            .checked_add(self.direct)
            .and_then(|value| value.checked_add(self.realloc_new))
            .and_then(|value| value.checked_sub(self.realloc_old))
            .and_then(|value| value.checked_sub(self.deallocated))
            == Some(self.live_after)
    }
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
    let mut hash = 0xcbf29ce484222325;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

pub fn json_escape(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '"' => escaped.push_str("\\\""),
            '\\' => escaped.push_str("\\\\"),
            '\n' => escaped.push_str("\\n"),
            '\r' => escaped.push_str("\\r"),
            '\t' => escaped.push_str("\\t"),
            character if character.is_control() => {
                use std::fmt::Write as _;
                let _ = write!(escaped, "\\u{:04x}", u32::from(character));
            },
            character => escaped.push(character),
        }
    }
    escaped
}
