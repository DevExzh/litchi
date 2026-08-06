//! Immutable snapshots and failure-atomic edits for annotation tags.

use super::model::{Tag, Tags};
use super::validation;
use crate::package::{Error as PackageError, Result as PackageResult};
use std::fmt;

/// An immutable optional `SttbfAtnBkmk` state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    tags: Option<Tags>,
}

impl Snapshot {
    /// Capture a present, validated table.
    pub fn new(tags: Tags) -> PackageResult<Self> {
        validation::tags(&tags)?;
        Ok(Self { tags: Some(tags) })
    }

    /// Capture an absent FIB range.
    #[must_use]
    pub const fn empty() -> Self {
        Self { tags: None }
    }

    pub(crate) fn from_option(tags: Option<Tags>) -> PackageResult<Self> {
        if let Some(value) = &tags {
            validation::tags(value)?;
        }
        Ok(Self { tags })
    }

    /// The present annotation tags, if any.
    #[must_use]
    pub fn tags(&self) -> Option<&Tags> {
        self.tags.as_ref()
    }

    /// Whether the FIB range is present.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.tags.is_some()
    }

    /// Start an independent transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            working: self.clone(),
        }
    }
}

impl Tags {
    /// Capture this table as an immutable snapshot.
    pub fn snapshot(&self) -> PackageResult<Snapshot> {
        Snapshot::new(self.clone())
    }

    /// Start a validated semantic transaction.
    pub fn edit(&self) -> PackageResult<Transaction> {
        Ok(self.snapshot()?.edit())
    }
}

impl From<Tags> for Snapshot {
    fn from(tags: Tags) -> Self {
        Self { tags: Some(tags) }
    }
}

/// A reversible semantic change between two snapshots.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source snapshot required by this patch.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Result snapshot produced by this patch.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the semantic state is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(self.after.clone(), self.before.clone())
    }

    /// Apply only to the exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source != &self.before {
            return Err(Error::Conflict);
        }
        Ok(self.after.clone())
    }
}

/// A committed semantic transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Post-edit immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible semantic patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

/// A staged, failure-atomic tag edit.
#[derive(Debug, Clone)]
pub struct Transaction {
    before: Snapshot,
    working: Snapshot,
}

impl Transaction {
    /// The current staged snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// Whether staged semantic values differ from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before != self.working
    }

    /// Restore the staged value to the source snapshot.
    pub fn rollback(&mut self) {
        self.working = self.before.clone();
    }

    /// Replace the complete optional table atomically.
    pub fn set(&mut self, tags: Option<Tags>) -> Result<(), Error> {
        self.working = Snapshot::from_option(tags).map_err(Error::Invalid)?;
        Ok(())
    }

    /// Replace the table while keeping its FIB range present.
    pub fn replace(&mut self, tags: Tags) -> Result<(), Error> {
        self.set(Some(tags))
    }

    /// Insert one tag into the present table.
    pub fn insert(&mut self, index: usize, tag: Tag) -> Result<(), Error> {
        let tags = self.working.tags.as_mut().ok_or(Error::Missing)?;
        tags.insert(index, tag).map_err(Error::Invalid)
    }

    /// Replace one tag in the present table.
    pub fn replace_entry(&mut self, index: usize, tag: Tag) -> Result<Tag, Error> {
        let tags = self.working.tags.as_mut().ok_or(Error::Missing)?;
        tags.replace(index, tag).map_err(Error::Invalid)
    }

    /// Remove one tag from the present table.
    pub fn remove(&mut self, index: usize) -> Result<Tag, Error> {
        let tags = self.working.tags.as_mut().ok_or(Error::Missing)?;
        tags.remove(index).map_err(Error::Invalid)
    }

    /// Remove the complete FIB range.
    pub fn clear(&mut self) {
        self.working = Snapshot::empty();
    }

    /// Commit the staged state into a reversible semantic patch.
    pub fn commit(self) -> Result<Commit, Error> {
        let patch = Patch::new(self.before, self.working.clone());
        Ok(Commit::new(self.working, patch))
    }
}

/// Failure modes for semantic edits.
#[derive(Debug)]
pub enum Error {
    /// The candidate violates an MS-DOC invariant.
    Invalid(PackageError),
    /// The transaction was created from a different snapshot.
    Conflict,
    /// An entry operation requires a present table.
    Missing,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => {
                formatter.write_str("annotation-bookmark transaction snapshot conflict")
            },
            Self::Missing => formatter.write_str("SttbfAtnBkmk table is absent"),
        }
    }
}

impl std::error::Error for Error {}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}
