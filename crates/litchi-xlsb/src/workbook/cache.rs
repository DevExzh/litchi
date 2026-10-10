//! Bounded retention for parsed XLSB worksheet snapshots.
//!
//! The cache stores complete, successfully parsed worksheet values behind
//! `Arc`.  A returned value therefore remains usable after the cache evicts
//! it, while the owner retains at most the configured entry and logical
//! weight ceilings.  Parsing is deliberately performed by the callers,
//! outside this cache's mutex; the mutex only protects LRU bookkeeping and
//! publication.

use std::sync::{Arc, Mutex, MutexGuard, PoisonError};

use crate::sheet::Worksheet;

/// Default maximum number of parsed worksheets retained by one owner.
pub(crate) const DEFAULT_MAX_ENTRIES: usize = 128;

/// Default logical byte ceiling for parsed worksheets retained by one owner.
pub(crate) const DEFAULT_MAX_WEIGHT: usize = 64 * 1024 * 1024;

/// Cache policy for one eager or source-backed workbook owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Limits {
    pub(crate) max_entries: usize,
    pub(crate) max_weight: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_entries: DEFAULT_MAX_ENTRIES,
            max_weight: DEFAULT_MAX_WEIGHT,
        }
    }
}

/// Content-free counters and gauges for one worksheet cache.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Stats {
    pub hits: u64,
    pub misses: u64,
    pub builds: u64,
    pub clean_evictions: u64,
    pub pinned_evictions: u64,
    pub oversized_bypasses: u64,
    pub retained_entries: usize,
    pub retained_weight: usize,
    pub clean_entries: usize,
    pub clean_weight: usize,
    pub max_entries: usize,
    pub max_weight: usize,
}

struct Entry {
    position: usize,
    value: Arc<Worksheet>,
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

    fn find(&mut self, position: usize) -> Option<&mut Entry> {
        self.entries
            .iter_mut()
            .find(|entry| entry.position == position)
    }

    /// Remove the least-recently-used value, preferring an unpinned value.
    fn evict_one(&mut self) -> Option<(Arc<Worksheet>, bool)> {
        let clean = self
            .entries
            .iter()
            .enumerate()
            .filter(|(_, entry)| Arc::strong_count(&entry.value) == 1)
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
        let was_clean = Arc::strong_count(&entry.value) == 1;
        if was_clean {
            self.clean_evictions = self.clean_evictions.saturating_add(1);
        } else {
            self.pinned_evictions = self.pinned_evictions.saturating_add(1);
        }
        Some((entry.value, was_clean))
    }
}

/// Owner-scoped weighted LRU of parsed worksheet snapshots.
pub(crate) struct WorksheetCache {
    limits: Limits,
    state: Mutex<State>,
}

impl WorksheetCache {
    pub(crate) fn new() -> Self {
        Self::with_limits(Limits::default())
    }

    pub(crate) fn with_limits(limits: Limits) -> Self {
        Self {
            limits,
            state: Mutex::new(State {
                entries: Vec::new(),
                weight: 0,
                tick: 0,
                hits: 0,
                misses: 0,
                builds: 0,
                clean_evictions: 0,
                pinned_evictions: 0,
                oversized_bypasses: 0,
            }),
        }
    }

    fn lock(&self) -> MutexGuard<'_, State> {
        self.state.lock().unwrap_or_else(PoisonError::into_inner)
    }

    /// Return a retained worksheet and refresh its recency.
    pub(crate) fn get(&self, position: usize, record_miss: bool) -> Option<Arc<Worksheet>> {
        let mut state = self.lock();
        let tick = state.touch();
        let value = state.find(position).map(|entry| {
            entry.last_used = tick;
            Arc::clone(&entry.value)
        });
        if value.is_some() {
            state.hits = state.hits.saturating_add(1);
        } else if record_miss {
            state.misses = state.misses.saturating_add(1);
        }
        value
    }

    /// Publish a complete worksheet and return the canonical retained value.
    ///
    /// A concurrent builder may have published the same position while this
    /// builder was parsing. In that case its value wins and the duplicate is
    /// dropped after the bookkeeping lock is released. Values too large for
    /// this cache are returned to their caller without retention.
    pub(crate) fn publish(
        &self,
        position: usize,
        value: Arc<Worksheet>,
        logical_weight: usize,
    ) -> Arc<Worksheet> {
        const ENTRY_OVERHEAD_BYTES: usize = 64;
        let weight = logical_weight.saturating_add(ENTRY_OVERHEAD_BYTES);
        let mut dropped = Vec::new();
        let canonical = {
            let mut state = self.lock();
            state.builds = state.builds.saturating_add(1);
            let tick = state.touch();
            if let Some(existing) = state.find(position) {
                existing.last_used = tick;
                Arc::clone(&existing.value)
            } else if self.limits.max_entries == 0 || weight > self.limits.max_weight {
                state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
                value.clone()
            } else {
                let mut can_retain = true;
                while state.entries.len() >= self.limits.max_entries
                    || state.weight.saturating_add(weight) > self.limits.max_weight
                {
                    let Some((victim, _was_clean)) = state.evict_one() else {
                        can_retain = false;
                        break;
                    };
                    if dropped.try_reserve(1).is_ok() {
                        dropped.push(victim);
                    }
                }
                if can_retain {
                    state.weight = state.weight.saturating_add(weight);
                    state.entries.push(Entry {
                        position,
                        value: value.clone(),
                        weight,
                        last_used: tick,
                    });
                    value.clone()
                } else {
                    state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
                    value.clone()
                }
            }
        };
        drop(dropped);
        canonical
    }

    /// Remove `value` only when it is still the exact entry published for
    /// `position`.
    ///
    /// Source-backed publication uses this as a rollback seam when the final
    /// source or execution fence fails after a candidate entered the cache.
    /// A later canonical value, or an entry owned by another builder, is
    /// never removed by a stale rollback attempt.
    pub(crate) fn discard_if(&self, position: usize, value: &Arc<Worksheet>) -> bool {
        let (removed, dropped) = {
            let mut state = self.lock();
            let Some(index) = state
                .entries
                .iter()
                .position(|entry| entry.position == position)
            else {
                return false;
            };
            if !Arc::ptr_eq(&state.entries[index].value, value) {
                return false;
            }
            let entry = state.entries.swap_remove(index);
            state.weight = state.weight.saturating_sub(entry.weight);
            (true, entry.value)
        };
        drop(dropped);
        removed
    }

    /// Release all entries that no returned handle pins.
    pub(crate) fn trim_clean(&self) -> usize {
        let mut dropped = Vec::new();
        let mut released = 0usize;
        {
            let mut state = self.lock();
            let mut index = 0;
            while index < state.entries.len() {
                if Arc::strong_count(&state.entries[index].value) == 1 {
                    let entry = state.entries.swap_remove(index);
                    state.weight = state.weight.saturating_sub(entry.weight);
                    state.clean_evictions = state.clean_evictions.saturating_add(1);
                    released = released.saturating_add(entry.weight);
                    if dropped.try_reserve(1).is_ok() {
                        dropped.push(entry.value);
                    }
                } else {
                    index += 1;
                }
            }
        }
        drop(dropped);
        released
    }

    pub(crate) fn stats(&self) -> Stats {
        let state = self.lock();
        let (clean_entries, clean_weight) = state
            .entries
            .iter()
            .filter(|entry| Arc::strong_count(&entry.value) == 1)
            .fold((0usize, 0usize), |(entries, weight), entry| {
                (
                    entries.saturating_add(1),
                    weight.saturating_add(entry.weight),
                )
            });
        Stats {
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
            max_entries: self.limits.max_entries,
            max_weight: self.limits.max_weight,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{Limits, WorksheetCache};
    use crate::sheet::Worksheet;
    use std::sync::Arc;

    fn worksheet(name: &str) -> Arc<Worksheet> {
        Arc::new(Worksheet::new(name.to_string()))
    }

    #[test]
    fn publish_returns_canonical_value_and_tracks_hits() {
        let cache = WorksheetCache::with_limits(Limits {
            max_entries: 2,
            max_weight: 1024,
        });
        let first = worksheet("one");
        let canonical = cache.publish(0, Arc::clone(&first), 1);
        let duplicate = cache.publish(0, worksheet("duplicate"), 1);
        assert!(Arc::ptr_eq(&canonical, &duplicate));
        assert!(cache.get(0, true).is_some());
        assert_eq!(cache.stats().builds, 2);
        assert_eq!(cache.stats().hits, 1);
    }

    #[test]
    fn lru_evicts_clean_values_but_keeps_pinned_values_usable() {
        let cache = WorksheetCache::with_limits(Limits {
            max_entries: 1,
            max_weight: 1024,
        });
        let first = cache.publish(0, worksheet("one"), 1);
        let second = cache.publish(1, worksheet("two"), 1);
        assert!(cache.get(0, false).is_none());
        assert_eq!(Arc::strong_count(&second), 2);
        assert_eq!(Arc::strong_count(&first), 1);
        assert_eq!(cache.stats().pinned_evictions, 1);
    }

    #[test]
    fn concurrent_builders_publish_one_canonical_snapshot() {
        let cache = WorksheetCache::with_limits(Limits {
            max_entries: 8,
            max_weight: 4096,
        });
        let barrier = std::sync::Barrier::new(8);
        let values = std::thread::scope(|scope| {
            let workers = (0..8)
                .map(|_| {
                    scope.spawn(|| {
                        let built = worksheet("concurrent");
                        barrier.wait();
                        cache.publish(0, built, 1)
                    })
                })
                .collect::<Vec<_>>();
            workers
                .into_iter()
                .map(|worker| worker.join().unwrap())
                .collect::<Vec<_>>()
        });
        assert!(values.iter().all(|value| Arc::ptr_eq(&values[0], value)));
        assert_eq!(cache.stats().retained_entries, 1);
        assert_eq!(cache.stats().builds, 8);
    }

    #[test]
    fn weight_ceiling_evicts_before_entry_ceiling() {
        let cache = WorksheetCache::with_limits(Limits {
            max_entries: 8,
            max_weight: 300,
        });
        drop(cache.publish(0, worksheet("old"), 100));
        let current = cache.publish(1, worksheet("current"), 100);
        assert!(cache.get(0, false).is_none());
        assert!(Arc::ptr_eq(&current, &cache.get(1, false).unwrap()));
        let stats = cache.stats();
        assert_eq!(stats.retained_entries, 1);
        assert_eq!(stats.clean_evictions, 1);
        assert!(stats.retained_weight <= stats.max_weight);
        assert_eq!(cache.trim_clean(), 0);
        drop(current);
        assert!(cache.trim_clean() > 0);
        assert_eq!(cache.stats().retained_entries, 0);
    }

    #[test]
    fn oversized_values_are_returned_without_retention() {
        let cache = WorksheetCache::with_limits(Limits {
            max_entries: 1,
            max_weight: 64,
        });
        let value = cache.publish(0, worksheet("one"), 64);
        assert_eq!(Arc::strong_count(&value), 1);
        assert_eq!(cache.stats().retained_entries, 0);
        assert_eq!(cache.stats().oversized_bypasses, 1);
    }
}
