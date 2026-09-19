//! Private weighted cache for source-backed worksheet occurrence indexes.
//!
//! The cache deliberately owns only immutable locators.  It does not retain
//! source handles or decoded cell values.  A cloned index `Arc` pins its
//! entry, and the memory reservations carried by the index remain charged for
//! exactly the same lifetime.

use litchi_core::{ExecutionContext, Reservation, Resource, SourceVersion};
use std::collections::VecDeque;
use std::mem::size_of;
use std::sync::{Arc, Mutex};

/// Logical charge for one retained worksheet occurrence slot.
pub(crate) const SLOT_WEIGHT: u64 = 24;
/// Logical charge for one worksheet index entry and its LRU metadata.
pub(crate) const INDEX_OVERHEAD: u64 = 128;
const INITIAL_SLOT_CAPACITY: usize = 8;

/// One occurrence of a stored BIFF cell in worksheet stream order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellSlot {
    pub(crate) stream_offset: u64,
    pub(crate) row: u16,
    pub(crate) column: u16,
    pub(crate) kind: u16,
    pub(crate) xf: u16,
    pub(crate) ordinal: u16,
}

/// An immutable worksheet occurrence index.
///
/// The reservation list is intentionally part of the pinned value rather
/// than the cache map entry.  A query which clones this `Arc` therefore keeps
/// the originating execution budget charged even if the cache evicts its own
/// reference while that query is still replaying the index.
#[derive(Debug)]
pub(crate) struct WorksheetCellIndex {
    pub(crate) expected_version: SourceVersion,
    pub(crate) slots: Vec<CellSlot>,
    _reservations: Vec<Reservation>,
}

impl WorksheetCellIndex {
    pub(crate) fn slots_for(&self, row: u16, column: u16) -> &[CellSlot] {
        let start = self
            .slots
            .partition_point(|slot| (slot.row, slot.column) < (row, column));
        let end =
            self.slots[start..].partition_point(|slot| (slot.row, slot.column) == (row, column));
        &self.slots[start..start + end]
    }
}

#[derive(Debug)]
struct CacheEntry {
    index: Arc<WorksheetCellIndex>,
    weight: u64,
}

#[derive(Debug)]
struct CacheState {
    max_bytes: u64,
    resident_weight: u64,
    reserved_weight: u64,
    entries: Vec<Option<CacheEntry>>,
    lru: VecDeque<usize>,
    hotness: Vec<u8>,
}

/// Snapshot-local weighted LRU state.
#[derive(Debug)]
pub(crate) struct QueryIndexCache {
    state: Mutex<CacheState>,
}

fn cache_metadata_weight(worksheet_count: usize) -> Option<u64> {
    cache_metadata_weight_from_capacities(worksheet_count, worksheet_count, worksheet_count)
}

fn cache_metadata_weight_from_capacities(
    entries: usize,
    lru: usize,
    hotness: usize,
) -> Option<u64> {
    let entries = u64::try_from(entries)
        .ok()?
        .checked_mul(size_of::<Option<CacheEntry>>() as u64)?;
    let lru = u64::try_from(lru)
        .ok()?
        .checked_mul(size_of::<usize>() as u64)?;
    let hotness = u64::try_from(hotness).ok()?;
    INDEX_OVERHEAD
        .checked_add(entries)?
        .checked_add(lru)?
        .checked_add(hotness)
}

impl QueryIndexCache {
    /// Creates a cache.  Cache metadata is optional optimization state: if the
    /// fixed hotness table cannot be reserved, the cache is disabled while the
    /// source owner remains usable through its ordinary scan path.
    pub(crate) fn new(max_bytes: u64, worksheet_count: usize) -> Self {
        let requested_metadata = cache_metadata_weight(worksheet_count);
        let mut entries = Vec::new();
        let mut lru = VecDeque::new();
        let mut hotness = Vec::new();
        let enabled = max_bytes != 0
            && requested_metadata.is_some_and(|weight| weight <= max_bytes)
            && entries.try_reserve_exact(worksheet_count).is_ok()
            && lru.try_reserve_exact(worksheet_count).is_ok()
            && hotness.try_reserve_exact(worksheet_count).is_ok()
            && {
                entries.resize_with(worksheet_count, || None);
                hotness.resize(worksheet_count, 0);
                true
            };
        let metadata_weight = if enabled {
            cache_metadata_weight_from_capacities(
                entries.capacity(),
                lru.capacity(),
                hotness.capacity(),
            )
            .filter(|weight| *weight <= max_bytes)
        } else {
            None
        };
        let enabled = enabled && metadata_weight.is_some();
        if !enabled {
            entries = Vec::new();
            lru = VecDeque::new();
            hotness = Vec::new();
        }
        Self {
            state: Mutex::new(CacheState {
                max_bytes: if enabled { max_bytes } else { 0 },
                resident_weight: metadata_weight.unwrap_or(0),
                reserved_weight: 0,
                entries,
                lru,
                hotness,
            }),
        }
    }

    /// Records a selected-cell miss and returns whether this is the second
    /// observation of the worksheet.  The counter is deliberately bounded:
    /// it is a fixed worksheet table, never a coordinate-keyed map.
    pub(crate) fn observe_miss(&self, sheet_index: usize) -> bool {
        let mut state = self.lock();
        if state.max_bytes == 0 {
            return false;
        }
        let Some(hotness) = state.hotness.get_mut(sheet_index) else {
            return false;
        };
        let second = *hotness >= 1;
        *hotness = (*hotness).saturating_add(1).min(2);
        second && state.max_bytes != 0
    }

    /// Looks up and pins an index while touching its MRU position.
    pub(crate) fn lookup(
        &self,
        sheet_index: usize,
        expected_version: SourceVersion,
    ) -> Option<Arc<WorksheetCellIndex>> {
        let mut state = self.lock();
        let stale = state
            .entries
            .get(sheet_index)
            .and_then(Option::as_ref)
            .is_some_and(|entry| entry.index.expected_version != expected_version);
        if stale {
            remove_entry(&mut state, sheet_index);
            return None;
        }
        let index = state
            .entries
            .get(sheet_index)
            .and_then(Option::as_ref)
            .map(|entry| Arc::clone(&entry.index))?;
        touch(&mut state.lru, sheet_index);
        Some(index)
    }

    /// Starts a candidate with its fixed index metadata charge.
    pub(crate) fn begin_candidate<'a>(
        self: &Arc<Self>,
        sheet_index: usize,
        execution: Option<&'a ExecutionContext>,
    ) -> Option<IndexCandidate<'a>> {
        let mut candidate = IndexCandidate {
            cache: Arc::clone(self),
            sheet_index,
            execution,
            slots: Vec::new(),
            reservations: Vec::new(),
            reserved_weight: 0,
            active: true,
        };
        if !candidate.add_weight(INDEX_OVERHEAD) {
            return None;
        }
        Some(candidate)
    }

    fn reserve_weight(&self, amount: u64) -> bool {
        let mut state = self.lock();
        if amount == 0 {
            return true;
        }
        let Some(total) = state
            .resident_weight
            .checked_add(state.reserved_weight)
            .and_then(|value| value.checked_add(amount))
        else {
            return false;
        };
        if total > state.max_bytes {
            let needed = total - state.max_bytes;
            evict_unpinned(&mut state, needed);
        }
        let Some(total) = state
            .resident_weight
            .checked_add(state.reserved_weight)
            .and_then(|value| value.checked_add(amount))
        else {
            return false;
        };
        if total > state.max_bytes {
            return false;
        }
        let Some(next) = state.reserved_weight.checked_add(amount) else {
            return false;
        };
        state.reserved_weight = next;
        true
    }

    fn release_reserved(&self, amount: u64) {
        if amount == 0 {
            return;
        }
        let mut state = self.lock();
        state.reserved_weight = state.reserved_weight.saturating_sub(amount);
    }

    fn publish_candidate(
        &self,
        candidate: IndexCandidate<'_>,
        expected_version: SourceVersion,
    ) -> Option<Arc<WorksheetCellIndex>> {
        if !candidate.active || candidate.reserved_weight == 0 {
            return None;
        }
        if candidate
            .execution
            .is_some_and(|execution| execution.check().is_err())
        {
            return None;
        }
        let mut candidate = candidate;
        let sheet_index = candidate.sheet_index;
        let reserved_weight = candidate.reserved_weight;
        let mut slots = std::mem::take(&mut candidate.slots);
        let reservations = std::mem::take(&mut candidate.reservations);
        slots
            .sort_unstable_by_key(|slot| (slot.row, slot.column, slot.stream_offset, slot.ordinal));
        if candidate
            .execution
            .is_some_and(|execution| execution.check().is_err())
        {
            drop(slots);
            drop(reservations);
            return None;
        }
        let index = Arc::new(WorksheetCellIndex {
            expected_version,
            slots,
            _reservations: reservations,
        });

        let mut state = self.lock();
        if let Some(existing) = state.entries.get(sheet_index).and_then(Option::as_ref) {
            let existing = Arc::clone(&existing.index);
            drop(state);
            return Some(existing);
        }

        // Candidate reservations were admitted incrementally, so the
        // invariant already proves this checked transfer cannot exceed the
        // owner-local ceiling.  Keep the guard for arithmetic-corrupt state.
        let Some(resident) = state.resident_weight.checked_add(reserved_weight) else {
            drop(index);
            return None;
        };
        if resident > state.max_bytes {
            drop(index);
            return None;
        }
        if sheet_index >= state.entries.len() {
            drop(index);
            return None;
        }
        state.reserved_weight = state.reserved_weight.saturating_sub(reserved_weight);
        state.resident_weight = resident;
        state.entries[sheet_index] = Some(CacheEntry {
            index: Arc::clone(&index),
            weight: reserved_weight,
        });
        candidate.active = false;
        touch(&mut state.lru, sheet_index);
        drop(state);
        Some(index)
    }

    fn lock(&self) -> std::sync::MutexGuard<'_, CacheState> {
        self.state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }
}

/// A fallible, incrementally weighted worksheet-index build.
pub(crate) struct IndexCandidate<'a> {
    cache: Arc<QueryIndexCache>,
    sheet_index: usize,
    execution: Option<&'a ExecutionContext>,
    slots: Vec<CellSlot>,
    reservations: Vec<Reservation>,
    reserved_weight: u64,
    active: bool,
}

impl IndexCandidate<'_> {
    /// Records one occurrence.  Allocation or cache-budget refusal abandons
    /// only the optional candidate; the caller keeps scanning for its result.
    pub(crate) fn push(&mut self, slot: CellSlot) {
        if !self.active {
            return;
        }
        if self.slots.len() == self.slots.capacity() {
            let current = self.slots.capacity();
            let additional = if current == 0 {
                INITIAL_SLOT_CAPACITY
            } else {
                current
            };
            let Some(delta) = u64::try_from(additional)
                .ok()
                .and_then(|count| count.checked_mul(SLOT_WEIGHT))
            else {
                self.abandon();
                return;
            };
            if !self.add_weight(delta) {
                self.abandon();
                return;
            }
            let old_capacity = self.slots.capacity();
            if self.slots.try_reserve_exact(additional).is_err() {
                self.abandon();
                return;
            }
            // `try_reserve_exact` is allowed to return a capacity larger than
            // requested. Charge that conservative extra before publishing.
            let actual = self.slots.capacity().saturating_sub(old_capacity);
            if actual > additional {
                let extra = u64::try_from(actual - additional)
                    .ok()
                    .and_then(|count| count.checked_mul(SLOT_WEIGHT));
                if extra.is_none_or(|amount| !self.add_weight(amount)) {
                    self.abandon();
                    return;
                }
            }
        }
        self.slots.push(slot);
    }

    /// Publishes the completed candidate and returns a pinned index.  A
    /// concurrent winner for the same worksheet is returned instead.
    pub(crate) fn publish(
        self,
        expected_version: SourceVersion,
    ) -> Option<Arc<WorksheetCellIndex>> {
        let cache = Arc::clone(&self.cache);
        cache.publish_candidate(self, expected_version)
    }

    fn add_weight(&mut self, amount: u64) -> bool {
        if amount == 0 {
            return true;
        }
        let reservation_growth = if self.execution.is_some()
            && self.reservations.len() == self.reservations.capacity()
        {
            if self.reservations.capacity() == 0 {
                1
            } else {
                self.reservations.capacity()
            }
        } else {
            0
        };
        let storage = u64::try_from(reservation_growth)
            .ok()
            .and_then(|count| count.checked_mul(size_of::<Reservation>() as u64));
        let Some(total) = storage.and_then(|value| amount.checked_add(value)) else {
            return false;
        };
        let memory = match self.execution {
            Some(execution) => match execution.reserve(Resource::Memory, total) {
                Ok(reservation) => Some(reservation),
                Err(_) => return false,
            },
            None => None,
        };
        if !self.cache.reserve_weight(total) {
            return false;
        }
        if reservation_growth != 0
            && self
                .reservations
                .try_reserve_exact(reservation_growth)
                .is_err()
        {
            self.cache.release_reserved(total);
            return false;
        }
        if let Some(memory) = memory {
            self.reservations.push(memory);
        }
        let extra_growth = if reservation_growth == 0 {
            0
        } else {
            let actual_growth = self
                .reservations
                .capacity()
                .saturating_sub(self.reservations.len().saturating_sub(1));
            actual_growth.saturating_sub(reservation_growth)
        };
        let Some(next) = self.reserved_weight.checked_add(total) else {
            self.cache.release_reserved(total);
            return false;
        };
        self.reserved_weight = next;
        if extra_growth != 0 {
            let extra = u64::try_from(extra_growth)
                .ok()
                .and_then(|count| count.checked_mul(size_of::<Reservation>() as u64));
            let Some(extra) = extra else {
                return false;
            };
            let Some(execution) = self.execution else {
                return false;
            };
            let Ok(memory) = execution.reserve(Resource::Memory, extra) else {
                return false;
            };
            if !self.cache.reserve_weight(extra) {
                return false;
            }
            if self.reservations.len() == self.reservations.capacity() {
                self.cache.release_reserved(extra);
                return false;
            }
            self.reservations.push(memory);
            self.reserved_weight = match self.reserved_weight.checked_add(extra) {
                Some(next) => next,
                None => {
                    self.cache.release_reserved(extra);
                    return false;
                },
            };
        }
        true
    }

    fn abandon(&mut self) {
        if self.active {
            let slots = std::mem::take(&mut self.slots);
            let reservations = std::mem::take(&mut self.reservations);
            drop(slots);
            drop(reservations);
            self.cache.release_reserved(self.reserved_weight);
            self.reserved_weight = 0;
            self.active = false;
        }
    }
}

impl Drop for IndexCandidate<'_> {
    fn drop(&mut self) {
        if self.active {
            let slots = std::mem::take(&mut self.slots);
            let reservations = std::mem::take(&mut self.reservations);
            drop(slots);
            drop(reservations);
            self.cache.release_reserved(self.reserved_weight);
        }
    }
}

fn touch(lru: &mut VecDeque<usize>, key: usize) {
    if let Some(position) = lru.iter().position(|candidate| *candidate == key) {
        lru.remove(position);
    }
    lru.push_back(key);
}

fn remove_entry(state: &mut CacheState, key: usize) {
    if let Some(entry) = state.entries.get_mut(key).and_then(Option::take) {
        state.resident_weight = state.resident_weight.saturating_sub(entry.weight);
    }
    state.lru.retain(|candidate| *candidate != key);
}

fn evict_unpinned(state: &mut CacheState, needed: u64) {
    let mut removed = 0_u64;
    let mut inspected = 0_usize;
    let initial = state.lru.len();
    while removed < needed && inspected < initial {
        let Some(key) = state.lru.pop_front() else {
            break;
        };
        inspected += 1;
        let removable = state
            .entries
            .get(key)
            .and_then(Option::as_ref)
            .is_some_and(|entry| Arc::strong_count(&entry.index) == 1);
        if removable {
            if let Some(entry) = state.entries.get_mut(key).and_then(Option::take) {
                removed = removed.saturating_add(entry.weight);
                state.resident_weight = state.resident_weight.saturating_sub(entry.weight);
            }
        } else {
            state.lru.push_back(key);
        }
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "cache lifecycle assertions intentionally panic"
    )]

    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Barrier;

    fn index(version: u64, reservations: Vec<Reservation>) -> Arc<WorksheetCellIndex> {
        Arc::new(WorksheetCellIndex {
            expected_version: SourceVersion::new(version, 0),
            slots: Vec::new(),
            _reservations: reservations,
        })
    }

    fn execution(memory: u64) -> (Budget, ExecutionContext) {
        let budget = Budget::root(
            "xls-query-cache-test",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (_cancellation, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(8 * 1024 * 1024).unwrap(),
            1,
        )
        .unwrap();
        (budget.clone(), ExecutionContext::new(budget, token, limits))
    }

    fn hierarchical_execution(memory: u64) -> (Budget, Budget, ExecutionContext) {
        let parent = Budget::root(
            "xls-query-cache-parent",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let child = parent.child(
            "xls-query-cache-child",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (_cancellation, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(8 * 1024 * 1024).unwrap(),
            1,
        )
        .unwrap();
        (
            child.clone(),
            parent,
            ExecutionContext::new(child, token, limits),
        )
    }

    fn slot() -> CellSlot {
        CellSlot {
            stream_offset: 4,
            row: 0,
            column: 0,
            kind: 0x0203,
            xf: 0,
            ordinal: 0,
        }
    }

    fn cache_with_index_budget(index_budget: u64, worksheet_count: usize) -> Arc<QueryIndexCache> {
        let metadata = cache_metadata_weight(worksheet_count).unwrap();
        Arc::new(QueryIndexCache::new(
            metadata.checked_add(index_budget).unwrap(),
            worksheet_count,
        ))
    }

    #[test]
    fn disabled_cache_never_admits_a_candidate() {
        let cache = Arc::new(QueryIndexCache::new(0, 1));
        assert!(!cache.observe_miss(0));
        assert!(!cache.observe_miss(0));
        assert!(cache.begin_candidate(0, None).is_none());
    }

    #[test]
    fn second_and_later_misses_can_retry_candidate_admission() {
        let cache = Arc::new(QueryIndexCache::new(1024, 1));
        assert!(!cache.observe_miss(0));
        assert!(cache.observe_miss(0));
        assert!(cache.observe_miss(0));
    }

    #[test]
    fn pinned_entry_survives_pressure_and_unpinned_entry_is_lru_evicted() {
        let version = SourceVersion::new(1, 0);
        let cache = cache_with_index_budget(INDEX_OVERHEAD * 2, 3);
        let first = cache
            .begin_candidate(0, None)
            .unwrap()
            .publish(version)
            .unwrap();
        drop(
            cache
                .begin_candidate(1, None)
                .unwrap()
                .publish(version)
                .unwrap(),
        );
        // The oldest entry is pinned; pressure must skip it and evict entry 1.
        let third = cache
            .begin_candidate(2, None)
            .unwrap()
            .publish(version)
            .unwrap();
        assert!(cache.lookup(0, version).is_some());
        assert!(cache.lookup(1, version).is_none());
        assert!(cache.lookup(2, version).is_some());
        drop(first);
        // Entry 0 is now the oldest unpinned entry; entry 2 remains pinned.
        drop(
            cache
                .begin_candidate(1, None)
                .unwrap()
                .publish(version)
                .unwrap(),
        );
        assert!(cache.lookup(0, version).is_none());
        assert!(cache.lookup(2, version).is_some());
        drop(third);
    }

    #[test]
    fn all_pinned_entries_refuse_a_candidate() {
        let cache = cache_with_index_budget(INDEX_OVERHEAD * 2, 2);
        for sheet in 0..2 {
            assert!(!cache.observe_miss(sheet));
            assert!(cache.observe_miss(sheet));
            let candidate = cache.begin_candidate(sheet, None).unwrap();
            let _ = candidate.publish(SourceVersion::new(1, 0)).unwrap();
        }
        let first = cache.lookup(0, SourceVersion::new(1, 0)).unwrap();
        let second = cache.lookup(1, SourceVersion::new(1, 0)).unwrap();
        assert!(cache.begin_candidate(0, None).is_none());
        drop(first);
        drop(second);
    }

    #[test]
    fn source_version_is_part_of_the_cache_boundary() {
        let cache = cache_with_index_budget(INDEX_OVERHEAD, 1);
        assert!(!cache.observe_miss(0));
        assert!(cache.observe_miss(0));
        let candidate = cache.begin_candidate(0, None).unwrap();
        let _ = candidate.publish(SourceVersion::new(1, 0)).unwrap();
        assert!(cache.lookup(0, SourceVersion::new(2, 0)).is_none());
    }

    #[test]
    fn cell_slots_are_grouped_by_coordinate() {
        let index = index(1, Vec::new());
        // The empty index still exposes a total, allocation-free range.
        assert!(index.slots_for(0, 0).is_empty());
    }

    #[test]
    fn index_overhead_covers_the_retained_arc_and_cache_metadata_layout() {
        let arc_header = 2 * size_of::<usize>();
        let retained = size_of::<WorksheetCellIndex>()
            .checked_add(arc_header)
            .and_then(|size| size.checked_add(size_of::<CacheEntry>()))
            .and_then(|size| size.checked_add(size_of::<usize>()))
            .unwrap();
        assert!(u64::try_from(retained).unwrap() <= INDEX_OVERHEAD);
        assert!((size_of::<QueryIndexCache>() + arc_header) as u64 <= INDEX_OVERHEAD);
        assert!(size_of::<CellSlot>() as u64 <= SLOT_WEIGHT);
    }

    #[test]
    fn dropping_an_unpublished_candidate_releases_its_managed_reservation() {
        let cache = cache_with_index_budget(INDEX_OVERHEAD * 8, 1);
        let (budget, execution) = execution(1_024);
        assert!(!cache.observe_miss(0));
        assert!(cache.observe_miss(0));
        let candidate = cache.begin_candidate(0, Some(&execution)).unwrap();
        assert!(budget.used(Resource::Memory) >= INDEX_OVERHEAD);
        drop(candidate);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn successful_retention_charges_the_hierarchy_until_cache_drop() {
        let cache = cache_with_index_budget(INDEX_OVERHEAD * 8, 2);
        let (budget, parent, execution) = hierarchical_execution(4 * 1024);
        assert!(!cache.observe_miss(0));
        assert!(cache.observe_miss(0));
        let first = cache
            .begin_candidate(0, Some(&execution))
            .unwrap()
            .publish(SourceVersion::new(1, 0))
            .unwrap();
        drop(first);

        assert!(!cache.observe_miss(1));
        assert!(cache.observe_miss(1));
        let mut second = cache.begin_candidate(1, Some(&execution)).unwrap();
        second.push(slot());
        let second = second.publish(SourceVersion::new(1, 0)).unwrap();
        assert!(budget.used(Resource::Memory) > INDEX_OVERHEAD);
        assert_eq!(budget.used(Resource::Memory), parent.used(Resource::Memory));
        drop(second);
        drop(cache);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(parent.used(Resource::Memory), 0);
    }

    #[test]
    fn pinned_entry_blocks_admission_without_releasing_its_reservation() {
        let entry_weight = INDEX_OVERHEAD
            + INITIAL_SLOT_CAPACITY as u64 * SLOT_WEIGHT
            + 2 * size_of::<Reservation>() as u64;
        let cache = cache_with_index_budget(entry_weight * 2 - 1, 2);
        let (budget, execution) = execution(4096);
        let mut first_candidate = cache.begin_candidate(0, Some(&execution)).unwrap();
        first_candidate.push(slot());
        let first = first_candidate.publish(SourceVersion::new(1, 0)).unwrap();
        let retained = budget.used(Resource::Memory);
        let mut blocked = cache.begin_candidate(1, Some(&execution)).unwrap();
        blocked.push(slot());
        assert!(blocked.publish(SourceVersion::new(1, 0)).is_none());
        assert_eq!(budget.used(Resource::Memory), retained);
        drop(cache);
        assert_eq!(budget.used(Resource::Memory), retained);
        drop(first);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn concurrent_duplicate_publish_releases_the_loser_and_pins_the_winner() {
        let cache = cache_with_index_budget(4096, 1);
        let (budget, parent, execution) = hierarchical_execution(4096);
        let barrier = Arc::new(Barrier::new(2));
        let results = std::thread::scope(|scope| {
            let mut handles = Vec::new();
            for _ in 0..2 {
                let cache = &cache;
                let execution = &execution;
                let barrier = &barrier;
                handles.push(scope.spawn(move || {
                    let mut candidate = cache.begin_candidate(0, Some(execution)).unwrap();
                    candidate.push(slot());
                    barrier.wait();
                    candidate.publish(SourceVersion::new(1, 0)).unwrap()
                }));
            }
            handles
                .into_iter()
                .map(|handle| handle.join().unwrap())
                .collect::<Vec<_>>()
        });
        assert!(Arc::ptr_eq(&results[0], &results[1]));
        let state = cache.lock();
        assert_eq!(state.reserved_weight, 0);
        assert_eq!(
            budget.used(Resource::Memory),
            state.entries[0].as_ref().unwrap().weight
        );
        assert_eq!(budget.used(Resource::Memory), parent.used(Resource::Memory));
        drop(state);
        drop(cache);
        assert!(parent.used(Resource::Memory) > 0);
        drop(results);
        assert_eq!(parent.used(Resource::Memory), 0);
    }

    #[test]
    fn abandoning_a_grown_candidate_releases_buffers_and_both_budgets() {
        let cache = cache_with_index_budget(4096, 1);
        let (budget, parent, execution) = hierarchical_execution(4096);
        let mut candidate = cache.begin_candidate(0, Some(&execution)).unwrap();
        for _ in 0..17 {
            candidate.push(slot());
        }
        assert!(candidate.slots.capacity() >= 17);
        assert!(parent.used(Resource::Memory) > 0);
        candidate.abandon();
        assert_eq!(candidate.slots.capacity(), 0);
        assert_eq!(candidate.reservations.capacity(), 0);
        assert_eq!(cache.lock().reserved_weight, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(parent.used(Resource::Memory), 0);
    }

    #[test]
    fn concurrent_reservations_share_one_owner_ceiling() {
        let cache = cache_with_index_budget(96, 1);
        let barrier = Arc::new(Barrier::new(2));
        let first_cache = Arc::clone(&cache);
        let first_barrier = Arc::clone(&barrier);
        let first = std::thread::spawn(move || {
            first_barrier.wait();
            first_cache.reserve_weight(96)
        });
        let second_cache = Arc::clone(&cache);
        let second_barrier = Arc::clone(&barrier);
        let second = std::thread::spawn(move || {
            second_barrier.wait();
            second_cache.reserve_weight(96)
        });
        let first = first.join().unwrap();
        let second = second.join().unwrap();
        assert_ne!(
            first, second,
            "in-flight reservations exceeded the cache ceiling"
        );
        cache.release_reserved(96);
    }
}
