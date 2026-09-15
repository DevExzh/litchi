//! Absolute process allocation counters with an operation-scoped region.
//!
//! The counters are process-global and only ever incremented; a region records
//! boundary snapshots and publishes checked differences. They are inert until
//! [`enable`] is called, which only the `peak` binary does, so the timing and
//! callgrind binary carries no allocator wrapper at all.

use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

static ENABLED: AtomicBool = AtomicBool::new(false);
static ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOCATION_CALLS: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

/// Publishes allocation regions. Called once by the instrumented binary.
pub fn enable() {
    ENABLED.store(true, Ordering::Release);
}

/// Whether the counting allocator is installed in this process.
#[must_use]
pub fn instrumented() -> bool {
    ENABLED.load(Ordering::Acquire)
}

fn bump_peak(live: u64) {
    let mut peak = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while live > peak {
        match PEAK_LIVE_BYTES.compare_exchange_weak(
            peak,
            live,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => peak = observed,
        }
    }
}

/// Records one successful allocation of `size` bytes.
pub fn record_allocation(size: usize) {
    let size = size as u64;
    ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
    ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
    let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed) + size;
    bump_peak(live);
}

/// Records one deallocation of `size` bytes.
pub fn record_deallocation(size: usize) {
    let size = size as u64;
    DEALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
    LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
}

/// Records one successful reallocation.
pub fn record_reallocation(old_size: usize, new_size: usize) {
    record_deallocation(old_size);
    record_allocation(new_size);
}

/// One boundary snapshot of the absolute counters.
#[derive(Clone, Copy, Debug, Default)]
pub struct Snapshot {
    allocated_bytes: u64,
    deallocated_bytes: u64,
    allocation_calls: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
}

fn snapshot() -> Snapshot {
    Snapshot {
        allocated_bytes: ALLOCATED_BYTES.load(Ordering::Relaxed),
        deallocated_bytes: DEALLOCATED_BYTES.load(Ordering::Relaxed),
        allocation_calls: ALLOCATION_CALLS.load(Ordering::Relaxed),
        live_bytes: LIVE_BYTES.load(Ordering::Relaxed),
        peak_live_bytes: PEAK_LIVE_BYTES.load(Ordering::Relaxed),
    }
}

/// One operation's allocation interval.
#[derive(Clone, Copy, Debug, Default)]
pub struct Region {
    /// Total bytes requested from the system allocator during the operation.
    pub allocated_bytes: u64,
    /// Total bytes returned to the system allocator during the operation.
    pub deallocated_bytes: u64,
    /// Successful allocation calls during the operation.
    pub allocation_calls: u64,
    /// Highest live byte count observed during the operation, relative to the
    /// live bytes already held when the operation began.
    pub peak_live_bytes: u64,
    /// Live bytes still held when the operation returned, relative to entry.
    pub retained_bytes: i64,
}

impl Region {
    /// Renders the region as one JSON object.
    #[must_use]
    pub fn json(&self) -> String {
        format!(
            "{{\"allocated_bytes\":{},\"deallocated_bytes\":{},\"allocation_calls\":{},\"peak_live_bytes\":{},\"retained_bytes\":{}}}",
            self.allocated_bytes,
            self.deallocated_bytes,
            self.allocation_calls,
            self.peak_live_bytes,
            self.retained_bytes
        )
    }
}

/// Runs `body` and publishes its allocation region.
///
/// # Errors
///
/// Propagates `body`'s error unchanged.
pub fn region<T, E>(body: impl FnOnce() -> Result<T, E>) -> Result<Region, E> {
    PEAK_LIVE_BYTES.store(LIVE_BYTES.load(Ordering::Relaxed), Ordering::Relaxed);
    let before = snapshot();
    let value = body()?;
    let after = snapshot();
    drop(value);
    Ok(Region {
        allocated_bytes: after.allocated_bytes.saturating_sub(before.allocated_bytes),
        deallocated_bytes: after
            .deallocated_bytes
            .saturating_sub(before.deallocated_bytes),
        allocation_calls: after.allocation_calls.saturating_sub(before.allocation_calls),
        peak_live_bytes: after.peak_live_bytes.saturating_sub(before.live_bytes),
        retained_bytes: after.live_bytes as i64 - before.live_bytes as i64,
    })
}
