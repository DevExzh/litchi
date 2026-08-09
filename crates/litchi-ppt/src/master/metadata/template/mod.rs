//! Contextual authoring of the main-master TemplateNameAtom.
//!
//! The short name API is intentional: callers already entered the template
//! facade, so the model is simply Name. Parent master snapshots also expose
//! template_name, set_template_name, and clear_template_name for composition
//! with SlideNameAtom edits.

pub(super) mod codec;
mod model;
pub(super) mod validation;

#[cfg(test)]
mod tests;

pub use crate::master_layout::{Change, ChangeSet, Context, Limits, Revision};
pub use model::{MAX_NAME_BYTES, Name};

use super as metadata;
use crate::package::Result;
use crate::records::Record;

/// An immutable main-master snapshot with a typed design-name view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    inner: metadata::Snapshot,
}

impl Snapshot {
    /// Validate and capture one main-master record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record(root: Record) -> Result<Self> {
        Self::from_inner(metadata::Snapshot::from_record(Context::Main, root)?)
    }

    /// Validate and capture one main-master record under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record_with_limits(root: Record, limits: Limits) -> Result<Self> {
        Self::from_inner(metadata::Snapshot::from_record_with_limits(
            Context::Main,
            root,
            limits,
        )?)
    }

    /// Parse one complete main-master record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_inner(metadata::Snapshot::parse(Context::Main, bytes)?)
    }

    /// Parse one complete main-master record under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_inner(metadata::Snapshot::parse_with_limits(
            Context::Main,
            bytes,
            limits,
        )?)
    }

    fn from_inner(inner: metadata::Snapshot) -> Result<Self> {
        validation::validate(Context::Main, inner.record())?;
        Ok(Self { inner })
    }

    /// The contextual master kind.
    #[must_use]
    pub const fn context(&self) -> Context {
        Context::Main
    }

    /// Borrow the validated main-master record tree.
    #[must_use]
    pub const fn record(&self) -> &Record {
        self.inner.record()
    }

    /// Borrow the exact encoded bytes represented by this snapshot.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.inner.bytes()
    }

    /// The stable content revision used for optimistic edits.
    #[must_use]
    pub fn revision(&self) -> Revision {
        self.inner.revision()
    }

    /// Read the optional main-master design name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<Option<Name>> {
        codec::read(self.context(), self.record())
    }

    /// Open an isolated semantic edit.
    #[must_use]
    pub fn edit(&self) -> Editor {
        Editor {
            inner: self.inner.edit(),
        }
    }
}

/// A transactional semantic edit of one main-master design name.
#[derive(Debug, Clone)]
pub struct Editor {
    inner: metadata::Editor,
}

impl Editor {
    /// The contextual master kind being edited.
    #[must_use]
    pub const fn context(&self) -> Context {
        Context::Main
    }

    /// Read the candidate design name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> Result<Option<Name>> {
        codec::read(self.context(), self.inner.record())
    }

    /// Whether this editor contains an uncommitted change.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.inner.is_changed()
    }

    /// The current structural change set.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        self.inner.changes()
    }

    /// Set or replace the design name atomically in the private candidate.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_name(&mut self, value: impl Into<String>) -> Result<()> {
        self.inner.set_template_name(value)
    }

    /// Remove the design name atom, returning whether one was present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn clear_name(&mut self) -> Result<bool> {
        self.inner.clear_template_name()
    }

    /// Capture the current candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<Snapshot> {
        Snapshot::from_inner(self.inner.snapshot()?)
    }

    /// Validate and publish the candidate as an immutable snapshot and patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        let commit = self.inner.commit()?;
        Ok(Commit {
            snapshot: Snapshot::from_inner(commit.snapshot().clone())?,
            changes: commit.changes().clone(),
        })
    }

    /// Discard the candidate and recover the original snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        Snapshot {
            inner: self.inner.rollback(),
        }
    }
}

/// A successful semantic commit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    changes: ChangeSet,
}

impl Commit {
    /// The immutable target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The structural patch produced by the semantic edit.
    #[must_use]
    pub const fn changes(&self) -> &ChangeSet {
        &self.changes
    }

    /// Undo this patch against its exact target snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        let inner = self.changes.undo(&current.inner.inner)?;
        Snapshot::from_inner(metadata::Snapshot::from_inner(inner)?)
    }

    /// Redo this patch against its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        let inner = self.changes.redo(&current.inner.inner)?;
        Snapshot::from_inner(metadata::Snapshot::from_inner(inner)?)
    }

    /// Split the commit into its target and reusable patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, ChangeSet) {
        (self.snapshot, self.changes)
    }
}
