//! Archive-neutral storage for ordered iWork package entries.
//!
//! This crate deliberately stops below the package facade. It owns only the
//! ordered entry table, its checked name index, and source-checked reversible
//! entry patches; ZIP I/O, IWA framing, application protobufs, and
//! format-specific document transactions remain in the format owner. That
//! makes the substrate reusable by the eventual Pages, Numbers, and Keynote
//! package crates without introducing peer format dependencies.

#![forbid(unsafe_code)]

use std::collections::HashMap;
use std::sync::Arc;

use thiserror::Error;

/// One ordered package member.
///
/// The payload is owned so a package snapshot can share the containing store
/// through copy-on-write. Entry names are intentionally opaque here: format
/// owners validate path and format-specific naming rules at their boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    name: Arc<str>,
    data: Arc<Vec<u8>>,
}

impl Entry {
    /// Construct one package member from its owned name and payload.
    #[must_use]
    pub fn new(name: String, data: Vec<u8>) -> Self {
        Self {
            name: Arc::from(name.into_boxed_str()),
            data: Arc::new(data),
        }
    }

    /// Borrow the package member name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the package member payload.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// Borrow the package member payload for an owner-controlled mutation.
    ///
    /// The package facade must recheck its resource and format invariants
    /// before publishing or serializing a mutation.
    #[must_use]
    pub fn data_mut(&mut self) -> &mut Vec<u8> {
        Arc::make_mut(&mut self.data)
    }

    /// Consume the member and return its owned name and payload.
    #[must_use]
    pub fn into_parts(self) -> (String, Vec<u8>) {
        let data = match Arc::try_unwrap(self.data) {
            Ok(data) => data,
            Err(data) => (*data).clone(),
        };
        (self.name.to_string(), data)
    }
}

/// Errors from the archive-neutral entry table.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// An ordered entry name already exists in the table.
    #[error("duplicate package entry: {0}")]
    DuplicateEntry(String),

    /// A requested insertion position is outside the ordered table.
    #[error("package entry position {position} is outside a table of length {len}")]
    InvalidPosition { position: usize, len: usize },

    /// The table could not reserve the requested index storage.
    #[error("allocation for package entry index ({requested} items) failed")]
    Allocation { requested: usize },

    /// A patch was applied to a table other than the one it was created from.
    #[error("package patch source does not match")]
    PatchSourceMismatch,
}

/// An ordered package-entry table with a checked name-to-position index.
///
/// The table stores each name and payload once. Lookups use the compact hash
/// index while iteration and serialization retain the source order. Structural
/// mutations update the index atomically with respect to fallible reservation;
/// callers never need to rebuild a second map themselves.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EntryStore {
    state: Arc<EntryStoreState>,
}

/// An immutable, cheaply shareable snapshot of ordered package entries.
///
/// Freezing or snapshotting an [`EntryStore`] shares its exact entry table,
/// name index, names, and payloads. A later mutation of the originating store
/// detaches through copy-on-write and cannot change this view. The frozen view
/// intentionally exposes no mutable access or conversion back into a mutable
/// store.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FrozenEntryStore {
    state: Arc<EntryStoreState>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct EntryStoreState {
    entries: Vec<Entry>,
    positions: HashMap<Arc<str>, usize>,
}

impl EntryStore {
    /// Construct an empty entry table.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Construct a table from owned entries while rejecting duplicate names.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DuplicateEntry`] when two entries have the same name,
    /// or [`Error::Allocation`] when the name index cannot reserve storage.
    pub fn try_from_entries(entries: Vec<Entry>) -> Result<Self, Error> {
        let mut positions = HashMap::new();
        positions
            .try_reserve(entries.len())
            .map_err(|_error| Error::Allocation {
                requested: entries.len(),
            })?;
        for (position, entry) in entries.iter().enumerate() {
            if positions
                .insert(Arc::clone(&entry.name), position)
                .is_some()
            {
                return Err(Error::DuplicateEntry(entry.name.to_string()));
            }
        }
        Ok(Self {
            state: Arc::new(EntryStoreState { entries, positions }),
        })
    }

    /// Consume this table into an immutable view without copying its state.
    #[must_use]
    pub fn freeze(self) -> FrozenEntryStore {
        FrozenEntryStore { state: self.state }
    }

    /// Capture an immutable view without copying entries, names, or payloads.
    #[must_use]
    pub fn snapshot(&self) -> FrozenEntryStore {
        FrozenEntryStore {
            state: Arc::clone(&self.state),
        }
    }

    /// Return the number of ordered entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.state.entries.len()
    }

    /// Return whether the table has no entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.state.entries.is_empty()
    }

    /// Borrow the ordered entries for traversal or serialization.
    #[must_use]
    pub fn as_slice(&self) -> &[Entry] {
        &self.state.entries
    }

    /// Iterate over ordered entries without allocating.
    pub fn iter(&self) -> impl Iterator<Item = &Entry> {
        self.state.entries.iter()
    }

    /// Find an entry position by exact name.
    #[must_use]
    pub fn position(&self, name: &str) -> Option<usize> {
        self.state.positions.get(name).copied()
    }

    /// Borrow an entry by its exact name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Entry> {
        self.position(name)
            .and_then(|position| self.get_at(position))
    }

    /// Borrow an entry by ordered position.
    #[must_use]
    pub fn get_at(&self, position: usize) -> Option<&Entry> {
        self.state.entries.get(position)
    }

    /// Mutably borrow an entry by ordered position.
    #[must_use]
    pub fn get_at_mut(&mut self, position: usize) -> Option<&mut Entry> {
        if position >= self.state.entries.len() {
            return None;
        }
        Arc::make_mut(&mut self.state).entries.get_mut(position)
    }

    /// Replace an entry payload without changing its name or position.
    pub fn replace_data(&mut self, position: usize, data: Vec<u8>) -> Option<Vec<u8>> {
        if position >= self.state.entries.len() {
            return None;
        }
        let entry = Arc::make_mut(&mut self.state).entries.get_mut(position)?;
        Some(std::mem::replace(entry.data_mut(), data))
    }

    /// Insert an entry at an ordered position.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidPosition`] for an out-of-range insertion,
    /// [`Error::DuplicateEntry`] for an existing name, or
    /// [`Error::Allocation`] when the table cannot reserve its next slot.
    pub fn try_insert_at(&mut self, position: usize, entry: Entry) -> Result<(), Error> {
        if position > self.state.entries.len() {
            return Err(Error::InvalidPosition {
                position,
                len: self.state.entries.len(),
            });
        }
        if self.state.positions.contains_key(entry.name()) {
            return Err(Error::DuplicateEntry(entry.name.to_string()));
        }
        let state = Arc::make_mut(&mut self.state);
        state
            .entries
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { requested: 1 })?;
        state
            .positions
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { requested: 1 })?;
        state.entries.insert(position, entry);
        state.rebuild_positions();
        Ok(())
    }

    /// Remove an entry by ordered position and return its owned contents.
    pub fn remove_at(&mut self, position: usize) -> Option<Entry> {
        if position >= self.state.entries.len() {
            return None;
        }
        let state = Arc::make_mut(&mut self.state);
        let entry = state.entries.remove(position);
        state.rebuild_positions();
        Some(entry)
    }
}

impl FrozenEntryStore {
    /// Return the number of ordered entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.state.entries.len()
    }

    /// Return whether the frozen table has no entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.state.entries.is_empty()
    }

    /// Iterate over entries in their preserved order without allocating.
    pub fn iter(&self) -> impl Iterator<Item = &Entry> {
        self.state.entries.iter()
    }

    /// Find an entry position by exact name.
    #[must_use]
    pub fn position(&self, name: &str) -> Option<usize> {
        self.state.positions.get(name).copied()
    }

    /// Borrow an entry by its exact name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Entry> {
        self.position(name)
            .and_then(|position| self.get_at(position))
    }

    /// Borrow an entry by ordered position.
    #[must_use]
    pub fn get_at(&self, position: usize) -> Option<&Entry> {
        self.state.entries.get(position)
    }
}

impl EntryStoreState {
    fn rebuild_positions(&mut self) {
        self.positions.clear();
        for (position, entry) in self.entries.iter().enumerate() {
            debug_assert!(
                self.positions
                    .insert(Arc::clone(&entry.name), position)
                    .is_none(),
                "entry names must be unique"
            );
        }
    }
}

/// The kind of one ordered package-entry change.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EntryChangeKind {
    /// An entry is present only in the target table.
    Added,
    /// An entry is present only in the source table.
    Removed,
    /// An entry exists in both tables but its bytes differ.
    Replaced,
    /// An entry exists in both tables with identical bytes but a new position.
    Reordered,
}

/// Deterministic metadata for one changed package entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EntryChange {
    name: String,
    kind: EntryChangeKind,
    before_position: Option<usize>,
    after_position: Option<usize>,
    before_len: Option<usize>,
    after_len: Option<usize>,
}

impl EntryChange {
    /// Return the changed entry name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the kind of change.
    #[must_use]
    pub const fn kind(&self) -> EntryChangeKind {
        self.kind
    }

    /// Return the source position, if the entry existed before the edit.
    #[must_use]
    pub const fn before_position(&self) -> Option<usize> {
        self.before_position
    }

    /// Return the target position, if the entry exists after the edit.
    #[must_use]
    pub const fn after_position(&self) -> Option<usize> {
        self.after_position
    }

    /// Return the source payload length, if the entry existed before the edit.
    #[must_use]
    pub const fn before_len(&self) -> Option<usize> {
        self.before_len
    }

    /// Return the target payload length, if the entry exists after the edit.
    #[must_use]
    pub const fn after_len(&self) -> Option<usize> {
        self.after_len
    }

    fn inverse(&self) -> Self {
        Self {
            name: self.name.clone(),
            kind: match self.kind {
                EntryChangeKind::Added => EntryChangeKind::Removed,
                EntryChangeKind::Removed => EntryChangeKind::Added,
                EntryChangeKind::Replaced => EntryChangeKind::Replaced,
                EntryChangeKind::Reordered => EntryChangeKind::Reordered,
            },
            before_position: self.after_position,
            after_position: self.before_position,
            before_len: self.after_len,
            after_len: self.before_len,
        }
    }
}

/// A source-checked, reversible patch between two ordered entry tables.
///
/// The source and target stores are cheap copy-on-write handles. Applying a
/// patch therefore shares all entry names and payloads with the target table;
/// it never serializes or clones package bytes. ZIP, IWA framing, protobuf,
/// and format-specific validation remain above this archive-neutral primitive.
#[derive(Debug, Clone)]
pub struct Patch {
    source: EntryStore,
    target: EntryStore,
    changes: Box<[EntryChange]>,
}

impl Patch {
    /// Version of the in-memory entry-patch representation.
    pub const VERSION: u16 = 1;

    /// Build a deterministic patch between two entry tables.
    #[must_use]
    pub fn between(source: &EntryStore, target: &EntryStore) -> Self {
        let mut changes = Vec::new();

        for (before_position, before_entry) in source.iter().enumerate() {
            let name = before_entry.name();
            let Some(after_position) = target.position(name) else {
                changes.push(EntryChange {
                    name: name.to_owned(),
                    kind: EntryChangeKind::Removed,
                    before_position: Some(before_position),
                    after_position: None,
                    before_len: Some(before_entry.data().len()),
                    after_len: None,
                });
                continue;
            };
            let Some(after_entry) = target.get_at(after_position) else {
                debug_assert!(false, "entry index must resolve through the name index");
                continue;
            };
            let kind_option = if before_entry.data() != after_entry.data() {
                Some(EntryChangeKind::Replaced)
            } else if before_position != after_position {
                Some(EntryChangeKind::Reordered)
            } else {
                None
            };
            if let Some(kind) = kind_option {
                changes.push(EntryChange {
                    name: name.to_owned(),
                    kind,
                    before_position: Some(before_position),
                    after_position: Some(after_position),
                    before_len: Some(before_entry.data().len()),
                    after_len: Some(after_entry.data().len()),
                });
            }
        }

        for (after_position, after_entry) in target.iter().enumerate() {
            let name = after_entry.name();
            if source.position(name).is_none() {
                changes.push(EntryChange {
                    name: name.to_owned(),
                    kind: EntryChangeKind::Added,
                    before_position: None,
                    after_position: Some(after_position),
                    before_len: None,
                    after_len: Some(after_entry.data().len()),
                });
            }
        }

        Self {
            source: source.clone(),
            target: target.clone(),
            changes: changes.into_boxed_slice(),
        }
    }

    /// Return the in-memory patch representation version.
    #[must_use]
    pub const fn version(&self) -> u16 {
        Self::VERSION
    }

    /// Return deterministic entry-level change metadata.
    #[must_use]
    pub fn changes(&self) -> &[EntryChange] {
        &self.changes
    }

    /// Return the number of changed entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.changes.len()
    }

    /// Return whether the patch changes no entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Build the inverse patch without copying entry payloads.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self
                .changes
                .iter()
                .rev()
                .map(EntryChange::inverse)
                .collect::<Vec<_>>()
                .into_boxed_slice(),
        }
    }

    /// Apply this patch to a matching source table.
    ///
    /// The returned table shares the target's entry allocations. A table with
    /// equal ordered names and payloads is accepted even when it came from a
    /// separate package parse.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchSourceMismatch`] when the supplied source does
    /// not exactly match the patch's source table.
    pub fn apply(&self, source: &EntryStore) -> Result<EntryStore, Error> {
        if source != &self.source {
            return Err(Error::PatchSourceMismatch);
        }
        Ok(self.target.clone())
    }
}

#[cfg(test)]
mod tests {
    use super::{Entry, EntryChangeKind, EntryStore, Error, FrozenEntryStore, Patch};

    fn entry(name: &str, data: &[u8]) -> Entry {
        Entry::new(name.to_owned(), data.to_vec())
    }

    #[test]
    fn preserves_order_and_indexes_exact_names() {
        let store = EntryStore::try_from_entries(vec![entry("b", &[2]), entry("a", &[1])])
            .unwrap_or_else(|error| panic!("valid entries rejected: {error}"));

        assert_eq!(store.position("b"), Some(0));
        assert_eq!(store.position("a"), Some(1));
        assert_eq!(
            store
                .get("b")
                .unwrap_or_else(|| panic!("indexed entry is missing"))
                .data(),
            [2]
        );
        assert_eq!(
            store.iter().map(Entry::name).collect::<Vec<_>>(),
            ["b", "a"]
        );
    }

    #[test]
    fn rejects_duplicate_names_without_mutating_the_source() {
        let Err(error) = EntryStore::try_from_entries(vec![entry("a", &[]), entry("a", &[1])])
        else {
            panic!("duplicate entry unexpectedly accepted");
        };

        assert_eq!(error, Error::DuplicateEntry("a".to_owned()));
    }

    #[test]
    fn structural_mutations_rebuild_the_index() {
        let mut store = EntryStore::try_from_entries(vec![entry("a", &[1]), entry("c", &[3])])
            .unwrap_or_else(|error| panic!("valid entries rejected: {error}"));

        store
            .try_insert_at(1, entry("b", &[2]))
            .unwrap_or_else(|error| panic!("valid insertion rejected: {error}"));
        assert_eq!(store.position("b"), Some(1));
        assert_eq!(store.position("c"), Some(2));
        assert_eq!(
            store
                .remove_at(0)
                .unwrap_or_else(|| panic!("entry to remove is missing"))
                .name(),
            "a"
        );
        assert_eq!(store.position("b"), Some(0));
        assert_eq!(store.replace_data(1, vec![9]), Some(vec![3]));
        assert_eq!(
            store
                .get("c")
                .unwrap_or_else(|| panic!("replaced entry is missing"))
                .data(),
            [9]
        );
    }

    #[test]
    fn rejects_invalid_insert_positions_and_duplicates() {
        let mut store = EntryStore::try_from_entries(vec![entry("a", &[])])
            .unwrap_or_else(|error| panic!("valid entry rejected: {error}"));

        assert_eq!(
            match store.try_insert_at(2, entry("b", &[])) {
                Ok(()) => panic!("invalid position unexpectedly accepted"),
                Err(error) => error,
            },
            Error::InvalidPosition {
                position: 2,
                len: 1
            }
        );
        assert_eq!(
            match store.try_insert_at(0, entry("a", &[])) {
                Ok(()) => panic!("duplicate entry unexpectedly accepted"),
                Err(error) => error,
            },
            Error::DuplicateEntry("a".to_owned())
        );
        assert_eq!(store.iter().map(Entry::name).collect::<Vec<_>>(), ["a"]);
    }

    #[test]
    fn cloned_stores_detach_only_when_mutated() {
        let source = EntryStore::try_from_entries(vec![entry("a", &[1])])
            .unwrap_or_else(|error| panic!("valid entry rejected: {error}"));
        let mut edited = source.clone();

        assert_eq!(edited.replace_data(0, vec![2]), Some(vec![1]));

        assert_eq!(source.get("a").map(Entry::data), Some([1].as_slice()));
        assert_eq!(edited.get("a").map(Entry::data), Some([2].as_slice()));
    }

    #[test]
    fn frozen_snapshots_share_state_and_are_isolated_from_later_mutation() {
        let mut store = EntryStore::try_from_entries(vec![entry("b", &[2]), entry("a", &[1])])
            .unwrap_or_else(|error| panic!("valid entries rejected: {error}"));
        let original_payload = store
            .get("b")
            .unwrap_or_else(|| panic!("source entry is missing"))
            .data()
            .as_ptr();
        let frozen = store.snapshot();

        assert!(std::sync::Arc::ptr_eq(&store.state, &frozen.state));
        assert_eq!(frozen.len(), 2);
        assert!(!frozen.is_empty());
        assert_eq!(frozen.position("b"), Some(0));
        assert_eq!(frozen.position("a"), Some(1));
        assert_eq!(frozen.get_at(1).map(Entry::name), Some("a"));
        assert_eq!(
            frozen.iter().map(Entry::name).collect::<Vec<_>>(),
            ["b", "a"]
        );
        assert_eq!(
            frozen
                .get("b")
                .unwrap_or_else(|| panic!("frozen entry is missing"))
                .data()
                .as_ptr(),
            original_payload
        );

        assert_eq!(store.replace_data(0, vec![9]), Some(vec![2]));
        assert!(!std::sync::Arc::ptr_eq(&store.state, &frozen.state));
        assert_eq!(store.get("b").map(Entry::data), Some([9].as_slice()));
        assert_eq!(frozen.get("b").map(Entry::data), Some([2].as_slice()));
        assert_eq!(
            frozen
                .get("b")
                .unwrap_or_else(|| panic!("frozen entry is missing"))
                .data()
                .as_ptr(),
            original_payload
        );
    }

    #[test]
    fn freezing_moves_the_exact_state_and_is_send_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        let store = EntryStore::try_from_entries(vec![entry("a", &[1])])
            .unwrap_or_else(|error| panic!("valid entry rejected: {error}"));
        let state = std::sync::Arc::clone(&store.state);
        let payload = store
            .get("a")
            .unwrap_or_else(|| panic!("source entry is missing"))
            .data()
            .as_ptr();

        let frozen = store.freeze();

        assert_send_sync::<FrozenEntryStore>();
        assert!(std::sync::Arc::ptr_eq(&state, &frozen.state));
        assert_eq!(
            frozen
                .get("a")
                .unwrap_or_else(|| panic!("frozen entry is missing"))
                .data()
                .as_ptr(),
            payload
        );
    }

    #[test]
    fn patches_are_reversible_and_source_checked() {
        let source = EntryStore::try_from_entries(vec![entry("a", &[1])])
            .unwrap_or_else(|error| panic!("valid source rejected: {error}"));
        let mut target = source.clone();
        assert_eq!(target.replace_data(0, vec![2]), Some(vec![1]));
        target
            .try_insert_at(1, entry("b", &[3]))
            .unwrap_or_else(|error| panic!("valid insertion rejected: {error}"));

        let patch = Patch::between(&source, &target);
        assert_eq!(patch.version(), Patch::VERSION);
        assert_eq!(patch.len(), 2);
        assert_eq!(patch.changes()[0].kind(), EntryChangeKind::Replaced);
        assert_eq!(patch.changes()[1].kind(), EntryChangeKind::Added);

        let mut reordered = source.clone();
        reordered
            .try_insert_at(0, entry("b", &[3]))
            .unwrap_or_else(|error| panic!("valid front insertion rejected: {error}"));
        let reorder_patch = Patch::between(&source, &reordered);
        assert_eq!(
            reorder_patch.changes()[0].kind(),
            EntryChangeKind::Reordered
        );
        assert_eq!(reorder_patch.changes()[1].kind(), EntryChangeKind::Added);

        let applied = patch
            .apply(&source)
            .unwrap_or_else(|error| panic!("valid patch rejected: {error}"));
        assert_eq!(applied, target);
        let reverted = patch
            .inverse()
            .apply(&applied)
            .unwrap_or_else(|error| panic!("inverse patch rejected: {error}"));
        assert_eq!(reverted, source);

        let unrelated = EntryStore::try_from_entries(vec![entry("other", &[])])
            .unwrap_or_else(|error| panic!("valid unrelated store rejected: {error}"));
        assert_eq!(patch.apply(&unrelated), Err(Error::PatchSourceMismatch));
    }
}
