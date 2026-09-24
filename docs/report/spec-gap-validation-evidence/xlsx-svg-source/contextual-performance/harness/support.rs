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

/// Return a deterministic lowercase SHA-256 digest for corpus identity.
///
/// The profile intentionally keeps this small and dependency-free so the
/// receipt carries a stable corpus hash even when the workspace dependency
/// graph changes.  It is not part of the production API.
pub fn sha256_hex(bytes: &[u8]) -> String {
    let mut state = [
        0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab,
        0x5be0cd19,
    ];
    let full_len = bytes.len() / 64 * 64;
    for chunk in bytes[..full_len].chunks_exact(64) {
        let block = chunk.try_into().expect("chunks_exact returned 64 bytes");
        sha256_compress(&mut state, block);
    }
    let remainder = &bytes[full_len..];
    let mut tail = [0_u8; 128];
    tail[..remainder.len()].copy_from_slice(remainder);
    tail[remainder.len()] = 0x80;
    let bit_len = u64::try_from(bytes.len())
        .unwrap_or(u64::MAX)
        .saturating_mul(8);
    let tail_len = if remainder.len() < 56 { 64 } else { 128 };
    tail[tail_len - 8..tail_len].copy_from_slice(&bit_len.to_be_bytes());
    for block in tail[..tail_len].chunks_exact(64) {
        let block = block.try_into().expect("chunks_exact returned 64 bytes");
        sha256_compress(&mut state, block);
    }
    let mut output = String::with_capacity(64);
    for word in state {
        use std::fmt::Write as _;
        write!(output, "{word:08x}").expect("String cannot fail");
    }
    output
}

fn sha256_compress(state: &mut [u32; 8], block: &[u8; 64]) {
    const ROUND: [u32; 64] = [
        0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4,
        0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe,
        0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f,
        0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7,
        0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc,
        0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b,
        0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116,
        0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
        0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7,
        0xc67178f2,
    ];
    let mut schedule = [0_u32; 64];
    for (index, word) in schedule.iter_mut().take(16).enumerate() {
        let offset = index * 4;
        *word = u32::from_be_bytes([
            block[offset],
            block[offset + 1],
            block[offset + 2],
            block[offset + 3],
        ]);
    }
    for index in 16..64 {
        let small_sigma0 = schedule[index - 15].rotate_right(7)
            ^ schedule[index - 15].rotate_right(18)
            ^ (schedule[index - 15] >> 3);
        let small_sigma1 = schedule[index - 2].rotate_right(17)
            ^ schedule[index - 2].rotate_right(19)
            ^ (schedule[index - 2] >> 10);
        schedule[index] = schedule[index - 16]
            .wrapping_add(small_sigma0)
            .wrapping_add(schedule[index - 7])
            .wrapping_add(small_sigma1);
    }
    let mut working = *state;
    for index in 0..64 {
        let big_sigma1 =
            working[4].rotate_right(6) ^ working[4].rotate_right(11) ^ working[4].rotate_right(25);
        let choose = (working[4] & working[5]) ^ ((!working[4]) & working[6]);
        let temporary1 = working[7]
            .wrapping_add(big_sigma1)
            .wrapping_add(choose)
            .wrapping_add(ROUND[index])
            .wrapping_add(schedule[index]);
        let big_sigma0 =
            working[0].rotate_right(2) ^ working[0].rotate_right(13) ^ working[0].rotate_right(22);
        let majority =
            (working[0] & working[1]) ^ (working[0] & working[2]) ^ (working[1] & working[2]);
        let temporary2 = big_sigma0.wrapping_add(majority);
        working[7] = working[6];
        working[6] = working[5];
        working[5] = working[4];
        working[4] = working[3].wrapping_add(temporary1);
        working[3] = working[2];
        working[2] = working[1];
        working[1] = working[0];
        working[0] = temporary1.wrapping_add(temporary2);
    }
    for (word, value) in state.iter_mut().zip(working) {
        *word = word.wrapping_add(value);
    }
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
