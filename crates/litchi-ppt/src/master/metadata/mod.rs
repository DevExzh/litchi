//! Contextual authoring of the `[MS-PPT]` `SlideNameAtom` record.
//!
//! The same atom names a main master, title master, notes master, and handout
//! master. This owner deliberately handles only that bounded semantic field;
//! every other child remains in the parent snapshot as an opaque record.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use crate::master_layout::{Change, ChangeSet, Context, Limits, Revision};
pub use model::{MAX_NAME_BYTES, Name};

use crate::master_layout;
use crate::package::Result;
use crate::records::Record;

/// An immutable, validated master snapshot with a contextual name view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    inner: master_layout::Snapshot,
}

impl Snapshot {
    /// Validate and capture one contextual master record.
    pub fn from_record(context: Context, root: Record) -> Result<Self> {
        Self::from_inner(master_layout::Snapshot::from_record(context, root)?)
    }

    /// Validate and capture one contextual master under explicit tree limits.
    pub fn from_record_with_limits(context: Context, root: Record, limits: Limits) -> Result<Self> {
        Self::from_inner(master_layout::Snapshot::from_record_with_limits(
            context, root, limits,
        )?)
    }

    /// Parse one complete contextual master record.
    pub fn parse(context: Context, bytes: &[u8]) -> Result<Self> {
        Self::from_inner(master_layout::Snapshot::parse(context, bytes)?)
    }

    /// Parse one complete contextual master under explicit tree limits.
    pub fn parse_with_limits(context: Context, bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_inner(master_layout::Snapshot::parse_with_limits(
            context, bytes, limits,
        )?)
    }

    fn from_inner(inner: master_layout::Snapshot) -> Result<Self> {
        validation::validate(inner.context(), inner.record())?;
        Ok(Self { inner })
    }

    /// The contextual master kind.
    pub const fn context(&self) -> Context {
        self.inner.context()
    }

    /// The validated source record tree.
    pub const fn record(&self) -> &Record {
        self.inner.record()
    }

    /// The exact encoded bytes represented by this snapshot.
    pub fn bytes(&self) -> &[u8] {
        self.inner.bytes()
    }

    /// The stable content revision used for optimistic edits.
    pub fn revision(&self) -> Revision {
        self.inner.revision()
    }

    /// Read the optional `[MS-PPT]` slide name from this master.
    pub fn name(&self) -> Result<Option<Name>> {
        codec::read(self.context(), self.record())
    }

    /// Open an isolated semantic edit.
    pub fn edit(&self) -> Editor {
        Editor {
            inner: self.inner.edit(),
        }
    }

    fn into_inner(self) -> master_layout::Snapshot {
        self.inner
    }
}

/// A transactional semantic edit of one master name.
#[derive(Debug, Clone)]
pub struct Editor {
    inner: master_layout::Transaction,
}

impl Editor {
    /// The contextual master kind being edited.
    pub const fn context(&self) -> Context {
        self.inner.context()
    }

    /// Read the candidate name without committing the edit.
    pub fn name(&self) -> Result<Option<Name>> {
        codec::read(self.context(), self.inner.record())
    }

    /// Whether this editor contains an uncommitted change.
    pub fn is_changed(&self) -> bool {
        self.inner.is_changed()
    }

    /// The current semantic change set.
    pub fn changes(&self) -> &[Change] {
        self.inner.changes()
    }

    /// Set or replace the master name atomically in the private candidate.
    pub fn set_name(&mut self, value: impl Into<String>) -> Result<()> {
        let name = Name::new(value)?;
        let replacement = codec::encode(&name)?;
        let index = validation::name_index(self.context(), self.inner.record())?;
        match index {
            Some(index) => {
                self.inner
                    .replace(master_layout::Path::root().child(index), replacement)?;
            },
            None => {
                let index = validation::name_insertion_index(self.context(), self.inner.record());
                self.inner
                    .add(master_layout::Path::root(), index, replacement)?;
            },
        }
        Ok(())
    }

    /// Remove the master name atom, returning whether one was present.
    pub fn clear_name(&mut self) -> Result<bool> {
        let Some(index) = validation::name_index(self.context(), self.inner.record())? else {
            return Ok(false);
        };
        self.inner
            .remove(master_layout::Path::root().child(index))?;
        Ok(true)
    }

    /// Capture the current candidate without publishing it.
    pub fn snapshot(&self) -> Result<Snapshot> {
        Snapshot::from_inner(self.inner.snapshot()?)
    }

    /// Validate and publish the candidate as an immutable snapshot and patch.
    pub fn commit(self) -> Result<Commit> {
        let (snapshot, changes) = self.inner.commit()?.into_parts();
        Ok(Commit {
            snapshot: Snapshot::from_inner(snapshot)?,
            changes,
        })
    }

    /// Discard the candidate and recover the original snapshot.
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
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The structural patch produced by the semantic edit.
    pub const fn changes(&self) -> &ChangeSet {
        &self.changes
    }

    /// Undo this patch against its exact target snapshot.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        Snapshot::from_inner(self.changes.undo(&current.inner)?)
    }

    /// Redo this patch against its exact source snapshot.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        Snapshot::from_inner(self.changes.redo(&current.inner)?)
    }

    /// Split the commit into its target and reusable patch.
    pub fn into_parts(self) -> (Snapshot, ChangeSet) {
        (self.snapshot, self.changes)
    }
}
