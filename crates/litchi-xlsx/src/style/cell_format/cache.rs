//! Bounded, evictable cache of parsed format tables.
//!
//! This is an ADR 0005 weighted clean-value cache, not an ADR 0032 memo: the
//! table's size grows with the styles payload. It follows the shape of the
//! DOCX paragraph-index cache (performance records 0680 and 0682):
//!
//! * one entry per cache, for the payload generation read last; the owner
//!   holds the cache, there is no process-wide state;
//! * the key is the payload's identity. An in-memory workbook keys on the
//!   payload allocation through a [`Weak`] reference, so the cache never
//!   keeps the payload alive, and the allocation it names cannot be reused
//!   while the key exists; a source-backed workbook keys on the captured
//!   source revision. A changed payload therefore never reuses a table;
//! * a table heavier than [`MAX_RETAINED_WEIGHT`] is returned to its caller
//!   but not retained, and is parsed again by a later read;
//! * entries are clean values: a miss replaces the entry, and
//!   [`FormatCache::trim_clean`] drops it unless an active [`super::Format`]
//!   pins it. Dropping the cache's reference never frees a table a handle
//!   still uses;
//! * a failed parse is never cached, so a refusal is always recomputed.
//!
//! Building holds the cache's lock, so concurrent first reads of one payload
//! parse it once.

use std::sync::{Arc, Mutex, MutexGuard, Weak};

use litchi_core::SourceVersion;

use super::FormatTable;
use crate::error::Result;

/// Logical ceiling, in bytes, on the clean weight one cache retains.
///
/// It bounds retention only and never refuses a read. On a 64-bit target a
/// `cellXfs` record weighs 56 bytes, so the format's maximum of 65,430
/// records weighs 3,664,168 bytes with the table header; font, fill and
/// border records add 88, 64 and 304 bytes each, plus their strings and
/// gradient stops (figures from the crate's weight test).
pub(crate) const MAX_RETAINED_WEIGHT: usize = 8 * 1024 * 1024;

/// Identity of the styles payload a table was parsed from.
#[derive(Debug, Clone)]
pub(crate) enum FormatKey {
    /// An in-memory workbook's styles payload allocation.
    Allocation(Weak<Vec<u8>>),
    /// A source-backed workbook's captured source revision. The facade is
    /// bound to one immutable source, so the revision names its payload.
    Source(SourceVersion),
}

impl FormatKey {
    pub(crate) fn allocation(payload: &Arc<Vec<u8>>) -> Self {
        Self::Allocation(Arc::downgrade(payload))
    }

    pub(crate) const fn source(version: SourceVersion) -> Self {
        Self::Source(version)
    }

    fn matches(&self, other: &Self) -> bool {
        match (self, other) {
            // The retained weak reference keeps the allocation's block
            // reserved, so an equal pointer is the same live allocation.
            (Self::Allocation(left), Self::Allocation(right)) => Weak::ptr_eq(left, right),
            (Self::Source(left), Self::Source(right)) => left == right,
            (Self::Allocation(_), Self::Source(_)) | (Self::Source(_), Self::Allocation(_)) => {
                false
            },
        }
    }
}

struct Entry {
    key: FormatKey,
    table: Arc<FormatTable>,
}

#[derive(Default)]
struct State {
    entry: Option<Entry>,
    hits: u64,
    misses: u64,
    builds: u64,
    clean_evictions: u64,
    pinned_evictions: u64,
    oversized_bypasses: u64,
}

impl State {
    fn evict(&mut self) {
        let Some(entry) = self.entry.take() else {
            return;
        };
        if Arc::strong_count(&entry.table) == 1 {
            self.clean_evictions = self.clean_evictions.saturating_add(1);
        } else {
            self.pinned_evictions = self.pinned_evictions.saturating_add(1);
        }
    }
}

/// Owner-scoped, single-entry cache of the parsed format table.
#[derive(Default)]
pub(crate) struct FormatCache {
    state: Mutex<State>,
}

impl std::fmt::Debug for FormatCache {
    // A workbook's debug output names what the cache holds, not the table.
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let state = self.lock();
        formatter
            .debug_struct("FormatCache")
            .field(
                "retained_weight",
                &state.entry.as_ref().map(|entry| entry.table.weight()),
            )
            .field("hits", &state.hits)
            .field("misses", &state.misses)
            .finish_non_exhaustive()
    }
}

impl FormatCache {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    fn lock(&self) -> MutexGuard<'_, State> {
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    /// Return the table of the payload `key` names, parsing it with `build`
    /// on a miss.
    pub(crate) fn get_or_build<F>(&self, key: FormatKey, build: F) -> Result<Arc<FormatTable>>
    where
        F: FnOnce() -> Result<FormatTable>,
    {
        let mut state = self.lock();
        if let Some(entry) = state.entry.as_ref()
            && entry.key.matches(&key)
        {
            let table = Arc::clone(&entry.table);
            state.hits = state.hits.saturating_add(1);
            return Ok(table);
        }
        state.misses = state.misses.saturating_add(1);
        // The entry belongs to another payload generation: it is no longer
        // the current value, so it is dropped before the new build.
        state.evict();
        let table = Arc::new(build()?);
        state.builds = state.builds.saturating_add(1);
        if table.weight() > MAX_RETAINED_WEIGHT {
            state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
        } else {
            state.entry = Some(Entry {
                key,
                table: Arc::clone(&table),
            });
        }
        Ok(table)
    }

    /// Drop the retained table unless an active handle pins it, returning
    /// whether a table was released. Pinned tables stay retained: dropping
    /// the cache's reference would not free them.
    pub(crate) fn trim_clean(&self) -> bool {
        let mut state = self.lock();
        let clean = state
            .entry
            .as_ref()
            .is_some_and(|entry| Arc::strong_count(&entry.table) == 1);
        if clean {
            state.evict();
        }
        clean
    }

    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> Diagnostics {
        let state = self.lock();
        let (retained_weight, pinned) = state.entry.as_ref().map_or((0, false), |entry| {
            (entry.table.weight(), Arc::strong_count(&entry.table) > 1)
        });
        Diagnostics {
            hits: state.hits,
            misses: state.misses,
            builds: state.builds,
            clean_evictions: state.clean_evictions,
            pinned_evictions: state.pinned_evictions,
            oversized_bypasses: state.oversized_bypasses,
            retained: state.entry.is_some(),
            retained_weight,
            pinned,
        }
    }
}

/// Content-free counters for focused cache tests.
#[cfg(test)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Diagnostics {
    pub(crate) hits: u64,
    pub(crate) misses: u64,
    pub(crate) builds: u64,
    pub(crate) clean_evictions: u64,
    pub(crate) pinned_evictions: u64,
    pub(crate) oversized_bypasses: u64,
    pub(crate) retained: bool,
    pub(crate) retained_weight: usize,
    pub(crate) pinned: bool,
}
