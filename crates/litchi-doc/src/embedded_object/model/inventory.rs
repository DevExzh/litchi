//! Ordered DOC embedded-object inventory snapshots.

use super::{Metadata, Reference};

/// One managed DOC field together with its inert ObjectPool metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Entry {
    reference: Reference,
    metadata: Metadata,
}

impl Entry {
    /// The field and storage reference.
    #[must_use]
    pub const fn reference(&self) -> &Reference {
        &self.reference
    }

    /// The passive metadata discovered in the ObjectPool storage.
    #[must_use]
    pub const fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// The semantic storage identifier used by the managed DOC field.
    #[must_use]
    pub const fn storage_id(&self) -> u32 {
        self.reference.storage_id
    }

    pub(in crate::embedded_object) const fn from_parts(
        reference: Reference,
        metadata: Metadata,
    ) -> Self {
        Self {
            reference,
            metadata,
        }
    }
}

/// Ordered, immutable inventory of managed DOC embedded objects.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Inventory {
    entries: Vec<Entry>,
}

impl Inventory {
    /// All entries in document field order.
    #[must_use]
    pub fn as_slice(&self) -> &[Entry] {
        &self.entries
    }

    /// Number of managed embedded-object fields.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether no managed embedded-object fields were found.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Finds an entry by its semantic DOC storage identifier.
    #[must_use]
    pub fn get(&self, storage_id: u32) -> Option<&Entry> {
        self.entries
            .iter()
            .find(|entry| entry.storage_id() == storage_id)
    }

    /// Finds an entry by checked document-order position.
    #[must_use]
    pub fn at(&self, index: usize) -> Option<&Entry> {
        self.entries.get(index)
    }

    /// Iterates over entries in document field order.
    pub fn iter(&self) -> std::slice::Iter<'_, Entry> {
        self.entries.iter()
    }

    /// Consumes the inventory into its owned entries.
    #[must_use]
    pub fn into_vec(self) -> Vec<Entry> {
        self.entries
    }

    pub(in crate::embedded_object) fn from_entries(entries: Vec<Entry>) -> Self {
        Self { entries }
    }
}

impl AsRef<[Entry]> for Inventory {
    fn as_ref(&self) -> &[Entry] {
        self.as_slice()
    }
}

impl<'a> IntoIterator for &'a Inventory {
    type Item = &'a Entry;
    type IntoIter = std::slice::Iter<'a, Entry>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}
