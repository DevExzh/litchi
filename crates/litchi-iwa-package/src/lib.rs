//! Archive-neutral storage for ordered iWork package entries.
//!
//! This crate deliberately stops below the package facade. It owns only the
//! ordered entry table and its checked name index; ZIP I/O, IWA framing,
//! application protobufs, and document transactions remain in the format
//! owner. That makes the substrate reusable by the eventual Pages, Numbers,
//! and Keynote package crates without introducing peer format dependencies.

#![forbid(unsafe_code)]

use std::collections::HashMap;

use thiserror::Error;

/// One ordered package member.
///
/// The payload is owned so a package snapshot can share the containing store
/// through copy-on-write. Entry names are intentionally opaque here: format
/// owners validate path and format-specific naming rules at their boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    name: String,
    data: Vec<u8>,
}

impl Entry {
    /// Construct one package member from its owned name and payload.
    #[must_use]
    pub const fn new(name: String, data: Vec<u8>) -> Self {
        Self { name, data }
    }

    /// Borrow the package member name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the package member payload.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Borrow the package member payload for an owner-controlled mutation.
    ///
    /// The package facade must recheck its resource and format invariants
    /// before publishing or serializing a mutation.
    #[must_use]
    pub fn data_mut(&mut self) -> &mut Vec<u8> {
        &mut self.data
    }

    /// Consume the member and return its owned name and payload.
    #[must_use]
    pub fn into_parts(self) -> (String, Vec<u8>) {
        (self.name, self.data)
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
}

/// An ordered package-entry table with a checked name-to-position index.
///
/// The table stores each name and payload once. Lookups use the compact hash
/// index while iteration and serialization retain the source order. Structural
/// mutations update the index atomically with respect to fallible reservation;
/// callers never need to rebuild a second map themselves.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EntryStore {
    entries: Vec<Entry>,
    positions: HashMap<String, usize>,
}

impl EntryStore {
    /// Construct an empty entry table.
    #[must_use]
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
            positions: HashMap::new(),
        }
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
            if positions.insert(entry.name.clone(), position).is_some() {
                return Err(Error::DuplicateEntry(entry.name.clone()));
            }
        }
        Ok(Self { entries, positions })
    }

    /// Return the number of ordered entries.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether the table has no entries.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Borrow the ordered entries for traversal or serialization.
    #[must_use]
    pub fn as_slice(&self) -> &[Entry] {
        &self.entries
    }

    /// Iterate over ordered entries without allocating.
    pub fn iter(&self) -> impl Iterator<Item = &Entry> {
        self.entries.iter()
    }

    /// Find an entry position by exact name.
    #[must_use]
    pub fn position(&self, name: &str) -> Option<usize> {
        self.positions.get(name).copied()
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
        self.entries.get(position)
    }

    /// Mutably borrow an entry by ordered position.
    #[must_use]
    pub fn get_at_mut(&mut self, position: usize) -> Option<&mut Entry> {
        self.entries.get_mut(position)
    }

    /// Replace an entry payload without changing its name or position.
    pub fn replace_data(&mut self, position: usize, data: Vec<u8>) -> Option<Vec<u8>> {
        let entry = self.entries.get_mut(position)?;
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
        if position > self.entries.len() {
            return Err(Error::InvalidPosition {
                position,
                len: self.entries.len(),
            });
        }
        if self.positions.contains_key(entry.name()) {
            return Err(Error::DuplicateEntry(entry.name.clone()));
        }
        self.entries
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { requested: 1 })?;
        self.positions
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { requested: 1 })?;
        self.entries.insert(position, entry);
        self.rebuild_positions();
        Ok(())
    }

    /// Remove an entry by ordered position and return its owned contents.
    pub fn remove_at(&mut self, position: usize) -> Option<Entry> {
        let entry = (position < self.entries.len()).then(|| self.entries.remove(position))?;
        self.rebuild_positions();
        Some(entry)
    }

    fn rebuild_positions(&mut self) {
        self.positions.clear();
        for (position, entry) in self.entries.iter().enumerate() {
            debug_assert!(
                self.positions
                    .insert(entry.name.clone(), position)
                    .is_none(),
                "entry names must be unique"
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{Entry, EntryStore, Error};

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
}
