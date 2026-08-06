//! Typed semantic values for `SttbfAtnBkmk` and `ATNBE`.

use super::validation;
use crate::package::{Error as PackageError, Result};

/// An opaque `ATNBE.lTag` identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[repr(transparent)]
pub struct TagId(u32);

impl TagId {
    /// Wrap an on-disk tag without resolving it to a comment.
    #[must_use]
    pub const fn new(raw: u32) -> Self {
        Self(raw)
    }

    /// Return the exact on-disk tag value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        self.0
    }
}

/// One inert annotation-bookmark tag (`ATNBE`, MS-DOC §2.9.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Tag {
    id: TagId,
}

impl Tag {
    /// Construct a tag with an opaque `lTag` value.
    #[must_use]
    pub const fn new(id: TagId) -> Self {
        Self { id }
    }

    /// The exact `lTag` value used to associate the annotation.
    #[must_use]
    pub const fn id(&self) -> TagId {
        self.id
    }
}

/// Ordered annotation-bookmark tags from one `SttbfAtnBkmk` table.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Tags {
    entries: Vec<Tag>,
}

impl Tags {
    /// Construct an empty present table.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Construct a table after checking count and tag uniqueness.
    pub fn try_new(entries: Vec<Tag>) -> Result<Self> {
        let value = Self { entries };
        value.validate()?;
        Ok(value)
    }

    /// Borrow tags in their on-disk order.
    #[must_use]
    pub fn entries(&self) -> &[Tag] {
        &self.entries
    }

    /// Return one tag by zero-based bookmark index.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Tag> {
        self.entries.get(index)
    }

    /// Number of annotation-bookmark entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether this present table contains no entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Append a tag atomically.
    pub fn push(&mut self, tag: Tag) -> Result<()> {
        let mut candidate = self.entries.clone();
        candidate.push(tag);
        self.entries = Self::try_new(candidate)?.entries;
        Ok(())
    }

    /// Insert a tag at an existing bookmark index atomically.
    pub fn insert(&mut self, index: usize, tag: Tag) -> Result<()> {
        if index > self.entries.len() {
            return Err(corrupted("SttbfAtnBkmk insertion index is out of bounds"));
        }
        let mut candidate = self.entries.clone();
        candidate.insert(index, tag);
        self.entries = Self::try_new(candidate)?.entries;
        Ok(())
    }

    /// Replace one tag and return the previous value.
    pub fn replace(&mut self, index: usize, tag: Tag) -> Result<Tag> {
        let previous = *self
            .entries
            .get(index)
            .ok_or_else(|| corrupted("SttbfAtnBkmk replacement index is out of bounds"))?;
        if self
            .entries
            .iter()
            .enumerate()
            .any(|(position, value)| position != index && value.id == tag.id)
        {
            return Err(corrupted("SttbfAtnBkmk lTag values must be unique"));
        }
        self.entries[index] = tag;
        Ok(previous)
    }

    /// Remove one tag and return it.
    pub fn remove(&mut self, index: usize) -> Result<Tag> {
        if index >= self.entries.len() {
            return Err(corrupted("SttbfAtnBkmk removal index is out of bounds"));
        }
        Ok(self.entries.remove(index))
    }

    /// Remove all entries while retaining a valid present table model.
    pub fn clear(&mut self) {
        self.entries.clear();
    }

    /// Validate the complete semantic table.
    pub fn validate(&self) -> Result<()> {
        validation::tags(self)
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
