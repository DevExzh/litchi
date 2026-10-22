//! Weighted clean-value cache of parsed source-backed worksheet values.
//!
//! A [`super::SourceBackedWorkbook`] keeps two kinds of parsed per-sheet
//! value between reads: a worksheet's parsed grid (its cell store) and the
//! projection of its blocks outside `sheetData`. Both are built in one pass
//! over the worksheet part the first time a read asks for them and are then
//! served from this cache, so later reads cost a lookup instead of another
//! scan of the part.
//!
//! The cache follows the weighted clean-value rule of ADR 0005 and of
//! performance change 0680:
//!
//! - Every entry carries a logical weight, and the cache never references
//!   more than [`MAX_RETAINED_WEIGHT`] bytes or [`MAX_RETAINED_ENTRIES`]
//!   values. A value that alone exceeds the weight ceiling is returned to the
//!   reader that built it but is not retained; a later reader rebuilds it.
//! - An entry is *clean* when only the cache holds it. A worksheet handle
//!   that has read a value pins it, so evicting the entry removes only the
//!   cache's reference and the handle keeps a valid value. Room is made by
//!   evicting the least-recently-used clean entries first; pinned entries
//!   are evicted only when no clean entry is left.
//! - Only complete, successfully built values are published. A failed or
//!   cancelled build publishes nothing, so a refusal is always recomputed
//!   from the source.
//! - Cache behaviour is semantically invisible: a miss rebuilds the same
//!   value from the same immutable source snapshot. The owning workbook is
//!   bound to one captured source version, and every read still runs the
//!   source-version and execution fences before it returns a cached value.
//!
//! The cache is not charged to an execution budget. The parsed worksheet
//! store was never charged before this cache existed; it was retained for the
//! workbook's lifetime instead. Charging it is a separate budget change.

use std::sync::{Arc, Mutex, MutexGuard, PoisonError};

use super::blocks::Blocks;
use super::cursor::Prefix;
use crate::cell::Store;

/// Logical ceiling on the parsed worksheet values one workbook retains.
///
/// A parsed store weighs a few hundred bytes per stored cell, so this keeps
/// roughly a third of a million cells resident across the workbook's sheets.
pub(super) const MAX_RETAINED_WEIGHT: usize = 64 * 1024 * 1024;

/// Ceiling on the number of retained values, whatever their weight.
pub(super) const MAX_RETAINED_ENTRIES: usize = 128;

/// Modeled bookkeeping cost of one retained entry.
const ENTRY_OVERHEAD_BYTES: usize = 64;

#[cfg(test)]
thread_local! {
    /// Test-only override of [`MAX_RETAINED_WEIGHT`] for the current thread.
    static MAX_WEIGHT_OVERRIDE: std::cell::Cell<Option<usize>> =
        const { std::cell::Cell::new(None) };
}

fn max_retained_weight() -> usize {
    #[cfg(test)]
    if let Some(bytes) = MAX_WEIGHT_OVERRIDE.with(std::cell::Cell::get) {
        return bytes;
    }
    MAX_RETAINED_WEIGHT
}

/// Lower the retained-weight ceiling of caches used on the current thread
/// while the returned guard is alive.
#[cfg(test)]
pub(super) fn limit_retained_weight(bytes: usize) -> WeightLimitGuard {
    MAX_WEIGHT_OVERRIDE.with(|value| value.set(Some(bytes)));
    WeightLimitGuard
}

#[cfg(test)]
pub(super) struct WeightLimitGuard;

#[cfg(test)]
impl Drop for WeightLimitGuard {
    fn drop(&mut self) {
        MAX_WEIGHT_OVERRIDE.with(|value| value.set(None));
    }
}

/// Which per-sheet value an entry holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Axis {
    /// The parsed cell store of one worksheet.
    Grid,
    /// The worksheet XML outside `sheetData`.
    Blocks,
    /// The parsed leading rows of one worksheet, which a cell cursor starts
    /// from (record 0267).
    Prefix,
}

/// One retained value.
#[derive(Clone)]
pub(super) enum Cached {
    Grid(Arc<Store>),
    Blocks(Arc<Blocks>),
    Prefix(Arc<Prefix>),
}

impl Cached {
    const fn axis(&self) -> Axis {
        match self {
            Self::Grid(_) => Axis::Grid,
            Self::Blocks(_) => Axis::Blocks,
            Self::Prefix(_) => Axis::Prefix,
        }
    }

    /// Whether only the cache holds this value.
    fn is_clean(&self) -> bool {
        match self {
            Self::Grid(value) => Arc::strong_count(value) == 1,
            Self::Blocks(value) => Arc::strong_count(value) == 1,
            Self::Prefix(value) => Arc::strong_count(value) == 1 && value.is_clean(),
        }
    }
}

struct Entry {
    position: usize,
    value: Cached,
    weight: usize,
    last_used: u64,
}

#[derive(Default)]
struct State {
    entries: Vec<Entry>,
    weight: usize,
    tick: u64,
    hits: u64,
    misses: u64,
    builds: u64,
    clean_evictions: u64,
    pinned_evictions: u64,
    oversized_bypasses: u64,
}

impl State {
    fn touch(&mut self) -> u64 {
        self.tick = self.tick.wrapping_add(1);
        self.tick
    }

    fn find(&mut self, position: usize, axis: Axis) -> Option<&mut Entry> {
        self.entries
            .iter_mut()
            .find(|entry| entry.position == position && entry.value.axis() == axis)
    }

    /// Remove the entry to evict next: the least-recently-used clean entry,
    /// or, when every entry is pinned, the least-recently-used entry.
    fn evict_one(&mut self) -> Option<Cached> {
        let clean = self
            .entries
            .iter()
            .enumerate()
            .filter(|(_, entry)| entry.value.is_clean())
            .min_by_key(|(_, entry)| entry.last_used)
            .map(|(index, _)| index);
        let index = clean.or_else(|| {
            self.entries
                .iter()
                .enumerate()
                .min_by_key(|(_, entry)| entry.last_used)
                .map(|(index, _)| index)
        })?;
        let entry = self.entries.swap_remove(index);
        self.weight = self.weight.saturating_sub(entry.weight);
        if entry.value.is_clean() {
            self.clean_evictions = self.clean_evictions.saturating_add(1);
        } else {
            self.pinned_evictions = self.pinned_evictions.saturating_add(1);
        }
        Some(entry.value)
    }
}

/// Workbook-scoped weighted clean-value cache; see the module documentation.
pub(super) struct ValueCache {
    state: Mutex<State>,
}

impl ValueCache {
    pub(super) fn new() -> Self {
        Self {
            state: Mutex::new(State::default()),
        }
    }

    fn lock(&self) -> MutexGuard<'_, State> {
        // A panic while the lock was held cannot leave a half-published
        // entry: every mutation completes before the guard is released.
        self.state.lock().unwrap_or_else(PoisonError::into_inner)
    }

    /// Look up one retained value and refresh its recency.
    ///
    /// `record_miss` is false for a lookup whose read goes on to look up
    /// again (the cell-route probe, and the re-check a builder makes after it
    /// has waited for a concurrent build of the same value), so one logical
    /// read counts at most one miss.
    pub(super) fn get(&self, position: usize, axis: Axis, record_miss: bool) -> Option<Cached> {
        let mut state = self.lock();
        let tick = state.touch();
        let found = state.find(position, axis).map(|entry| {
            entry.last_used = tick;
            entry.value.clone()
        });
        if found.is_some() {
            state.hits = state.hits.saturating_add(1);
        } else if record_miss {
            state.misses = state.misses.saturating_add(1);
        }
        found
    }

    /// Whether a value is retained, without counting a lookup or refreshing
    /// its recency.
    #[cfg(test)]
    pub(super) fn holds(&self, position: usize, axis: Axis) -> bool {
        self.lock().find(position, axis).is_some()
    }

    /// Offer one freshly built value and return the canonical one.
    ///
    /// If another reader published the same value while this one was being
    /// built, the published value is returned and the new one is dropped, so
    /// concurrent builds never retain two copies. A value heavier than the
    /// whole ceiling is returned without being retained. A worksheet prefix
    /// replaces a retained prefix with fewer rows.
    pub(super) fn publish(&self, position: usize, value: Cached, weight: usize) -> Cached {
        let weight = weight.saturating_add(ENTRY_OVERHEAD_BYTES);
        let max_weight = max_retained_weight();
        let mut evicted = Vec::new();
        let canonical = {
            let mut state = self.lock();
            state.builds = state.builds.saturating_add(1);
            let tick = state.touch();
            let longer = match (state.find(position, value.axis()), &value) {
                (
                    Some(Entry {
                        value: Cached::Prefix(held),
                        ..
                    }),
                    Cached::Prefix(offered),
                ) => offered.rows() > held.rows(),
                _ => false,
            };
            if longer && weight <= max_weight {
                // Releasing the shorter prefix makes room the longer one
                // then claims through the ordinary path below.
                if let Some(index) = state.entries.iter().position(|entry| {
                    entry.position == position && entry.value.axis() == Axis::Prefix
                }) {
                    let entry = state.entries.swap_remove(index);
                    state.weight = state.weight.saturating_sub(entry.weight);
                    if evicted.try_reserve(1).is_ok() {
                        evicted.push(entry.value);
                    }
                }
            }
            if let Some(existing) = state.find(position, value.axis()) {
                existing.last_used = tick;
                existing.value.clone()
            } else if weight > max_weight || state.entries.try_reserve(1).is_err() {
                // Not retaining a value is always correct: a later reader
                // rebuilds it from the same snapshot.
                state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
                value
            } else {
                while state.entries.len() >= MAX_RETAINED_ENTRIES
                    || state.weight.saturating_add(weight) > max_weight
                {
                    let Some(victim) = state.evict_one() else {
                        break;
                    };
                    // Large values are released after the lock is dropped
                    // whenever the reservation succeeds.
                    if evicted.try_reserve(1).is_ok() {
                        evicted.push(victim);
                    }
                }
                state.weight = state.weight.saturating_add(weight);
                state.entries.push(Entry {
                    position,
                    value: value.clone(),
                    weight,
                    last_used: tick,
                });
                value
            }
        };
        drop(evicted);
        canonical
    }

    /// Release every retained value no handle pins and return the logical
    /// weight released. Pinned values stay retained.
    pub(super) fn trim_clean(&self) -> usize {
        let mut released = Vec::new();
        let weight = {
            let mut state = self.lock();
            let mut weight = 0usize;
            let mut index = 0;
            while index < state.entries.len() {
                if state.entries[index].value.is_clean() {
                    let entry = state.entries.swap_remove(index);
                    state.weight = state.weight.saturating_sub(entry.weight);
                    state.clean_evictions = state.clean_evictions.saturating_add(1);
                    weight = weight.saturating_add(entry.weight);
                    if released.try_reserve(1).is_ok() {
                        released.push(entry.value);
                    }
                } else {
                    index += 1;
                }
            }
            weight
        };
        drop(released);
        weight
    }

    /// Point-in-time content-free diagnostics.
    pub(super) fn diagnostics(&self) -> CacheDiagnostics {
        let state = self.lock();
        let (clean_entries, clean_weight) = state
            .entries
            .iter()
            .filter(|entry| entry.value.is_clean())
            .fold((0usize, 0usize), |(count, weight), entry| {
                (count + 1, weight.saturating_add(entry.weight))
            });
        CacheDiagnostics {
            hits: state.hits,
            misses: state.misses,
            builds: state.builds,
            clean_evictions: state.clean_evictions,
            pinned_evictions: state.pinned_evictions,
            oversized_bypasses: state.oversized_bypasses,
            retained_entries: state.entries.len(),
            retained_weight: state.weight,
            clean_entries,
            clean_weight,
            max_weight: max_retained_weight(),
            max_entries: MAX_RETAINED_ENTRIES,
        }
    }
}

/// Content-free diagnostics of a source-backed workbook's parsed-value cache.
///
/// The cache holds parsed worksheet grids and the projections of their
/// blocks outside `sheetData` (see
/// [`SourceBackedWorkbook::semantic_cache_diagnostics`]). Weights are logical
/// byte estimates of the retained values, not allocator or resident-memory
/// measurements. Counters increase monotonically for the workbook's
/// lifetime; the gauges describe the moment the snapshot was taken.
///
/// [`SourceBackedWorkbook::semantic_cache_diagnostics`]: super::SourceBackedWorkbook::semantic_cache_diagnostics
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct CacheDiagnostics {
    /// Lookups answered by a retained value.
    pub hits: u64,
    /// Grid or block reads that found no retained value. A cell read that
    /// takes the streaming route makes no lookup and is not counted.
    pub misses: u64,
    /// Values built from the source and offered to the cache.
    pub builds: u64,
    /// Values released while only the cache held them.
    pub clean_evictions: u64,
    /// Values released from the cache while a worksheet handle still pinned
    /// them; the handle keeps its value.
    pub pinned_evictions: u64,
    /// Built values the cache did not retain because they alone exceed
    /// [`Self::max_weight`].
    pub oversized_bypasses: u64,
    /// Values currently retained.
    pub retained_entries: usize,
    /// Logical weight of every retained value, pinned or clean.
    pub retained_weight: usize,
    /// Retained values that no worksheet handle pins.
    pub clean_entries: usize,
    /// Logical weight of the retained values that no handle pins.
    pub clean_weight: usize,
    /// The ceiling on [`Self::retained_weight`].
    pub max_weight: usize,
    /// The ceiling on [`Self::retained_entries`].
    pub max_entries: usize,
}
