//! Contextual authoring of bounded `[MS-PPT]` master metadata.
//!
//! The same atom names a main master, title master, notes master, and handout
//! master. The notes-master text-style package is a separate nested owner.
//! Every other child remains in the parent snapshot as an opaque record.
//!
//! [`Snapshot`] reads one master record; an [`Edit`] (from
//! [`crate::Package::edit_master_metadata`]) stages the name, design name and
//! notes styles of one master of a package and publishes a new package.

mod codec;
mod model;
mod validation;

/// Notes-master round-trip text-style package metadata.
pub mod notes_styles;
/// Main-master design-name metadata.
pub mod template;

mod edit;
#[cfg(test)]
mod tests;

pub use crate::master_layout::{Change, Context, Limits, Revision};
pub use edit::Edit;
pub use model::{MAX_NAME_BYTES, Name};

use crate::master_layout;
use crate::package::Result as RecordResult;
use crate::records::Record;

/// An immutable, validated master snapshot with a contextual name view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    inner: master_layout::Snapshot,
}

impl Snapshot {
    /// Validate and capture one contextual master record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record(context: Context, root: Record, limits: &Limits) -> RecordResult<Self> {
        Self::from_inner(master_layout::Snapshot::from_record(context, root, limits)?)
    }

    /// Parse one complete contextual master record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(context: Context, bytes: &[u8], limits: &Limits) -> RecordResult<Self> {
        Self::from_inner(master_layout::Snapshot::parse(context, bytes, limits)?)
    }

    fn from_inner(inner: master_layout::Snapshot) -> RecordResult<Self> {
        validation::validate(inner.context(), inner.record())?;
        Ok(Self { inner })
    }

    /// The contextual master kind.
    #[must_use]
    pub const fn context(&self) -> Context {
        self.inner.context()
    }

    /// The validated source record tree.
    #[must_use]
    pub const fn record(&self) -> &Record {
        self.inner.record()
    }

    /// The exact encoded bytes represented by this snapshot.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.inner.bytes()
    }

    /// The stable content revision used for optimistic edits.
    #[must_use]
    pub fn revision(&self) -> Revision {
        self.inner.revision()
    }

    /// Read the optional `[MS-PPT]` slide name from this master.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn name(&self) -> RecordResult<Option<Name>> {
        codec::read(self.context(), self.record())
    }

    /// Read the optional main-master design name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn template_name(&self) -> RecordResult<Option<template::Name>> {
        template::codec::read(self.context(), self.record())
    }

    /// Read the optional notes-master `txStyles` round-trip package.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn notes_styles(&self) -> RecordResult<Option<notes_styles::Styles>> {
        notes_styles::codec::read(self.context(), self.record())
    }

    /// Open an isolated semantic edit.
    #[must_use]
    pub(crate) fn edit(&self) -> RecordEdit {
        RecordEdit {
            inner: self.inner.edit(),
        }
    }
}

/// A transactional semantic edit of one master’s bounded metadata.
#[derive(Debug, Clone)]
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
pub(crate) struct RecordEdit {
    inner: master_layout::RecordEdit,
}

#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
impl RecordEdit {
    /// The contextual master kind being edited.
    #[must_use]
    pub(crate) const fn context(&self) -> Context {
        self.inner.context()
    }

    /// Read the candidate name without committing the edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn name(&self) -> RecordResult<Option<Name>> {
        codec::read(self.context(), self.inner.record())
    }

    /// Read the candidate main-master design name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn template_name(&self) -> RecordResult<Option<template::Name>> {
        template::codec::read(self.context(), self.inner.record())
    }

    /// Read the candidate notes-master `txStyles` round-trip package.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn notes_styles(&self) -> RecordResult<Option<notes_styles::Styles>> {
        notes_styles::codec::read(self.context(), self.inner.record())
    }

    /// Borrow the transaction-local master record.
    #[must_use]
    pub(crate) fn record(&self) -> &Record {
        self.inner.record()
    }

    /// The current semantic change set.
    #[must_use]
    pub(crate) fn changes(&self) -> &[Change] {
        self.inner.changes()
    }

    /// Set or replace the master name atomically in the private candidate.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn set_name(&mut self, value: impl Into<String>) -> RecordResult<()> {
        let name = Name::new(value)?;
        let replacement = codec::encode(&name)?;
        let existing = validation::name_index(self.context(), self.inner.record())?;
        if let Some(index) = existing {
            self.inner
                .replace(master_layout::Path::root().child(index), replacement)?;
        } else {
            let insertion = validation::name_insertion_index(self.context(), self.inner.record());
            self.inner
                .add(master_layout::Path::root(), insertion, replacement)?;
        }
        Ok(())
    }

    /// Set or replace the main-master design name atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn set_template_name(&mut self, value: impl Into<String>) -> RecordResult<()> {
        let name = template::Name::new(value)?;
        let replacement = template::codec::encode(&name)?;
        let existing = template::validation::template_index(self.context(), self.inner.record())?;
        if let Some(index) = existing {
            self.inner
                .replace(master_layout::Path::root().child(index), replacement)?;
        } else {
            let insertion = template::validation::template_insertion_index(
                self.context(),
                self.inner.record(),
            )?;
            self.inner
                .add(master_layout::Path::root(), insertion, replacement)?;
        }
        Ok(())
    }

    /// Set or replace the notes-master `txStyles` package atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "callers hand over the authored `Styles` package; taking it by reference would \
                  churn the established public API and its callers"
    )]
    pub(crate) fn set_notes_styles(&mut self, styles: notes_styles::Styles) -> RecordResult<()> {
        let replacement = notes_styles::codec::encode(&styles)?;
        let existing = notes_styles::validation::styles_index(self.context(), self.inner.record())?;
        if let Some(index) = existing {
            self.inner
                .replace(master_layout::Path::root().child(index), replacement)?;
        } else {
            let insertion = notes_styles::validation::styles_insertion_index(
                self.context(),
                self.inner.record(),
            )?;
            self.inner
                .add(master_layout::Path::root(), insertion, replacement)?;
        }
        Ok(())
    }

    /// Build and set a notes-master `txStyles` package from XML atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn set_notes_styles_xml(&mut self, xml: impl AsRef<[u8]>) -> RecordResult<()> {
        self.set_notes_styles(notes_styles::Styles::from_xml(xml)?)
    }

    /// Remove the master name atom, returning whether one was present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn clear_name(&mut self) -> RecordResult<bool> {
        let Some(index) = validation::name_index(self.context(), self.inner.record())? else {
            return Ok(false);
        };
        self.inner
            .remove(master_layout::Path::root().child(index))?;
        Ok(true)
    }

    /// Remove the main-master design name atom, returning whether one was present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn clear_template_name(&mut self) -> RecordResult<bool> {
        let Some(index) =
            template::validation::template_index(self.context(), self.inner.record())?
        else {
            return Ok(false);
        };
        self.inner
            .remove(master_layout::Path::root().child(index))?;
        Ok(true)
    }

    /// Remove the notes-master `txStyles` package, returning whether one was present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn clear_notes_styles(&mut self) -> RecordResult<bool> {
        let Some(index) =
            notes_styles::validation::styles_index(self.context(), self.inner.record())?
        else {
            return Ok(false);
        };
        self.inner
            .remove(master_layout::Path::root().child(index))?;
        Ok(true)
    }

    /// Validate and publish the candidate as an immutable snapshot and patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn commit(self) -> RecordResult<RecordCommit> {
        let (snapshot, changes) = self.inner.commit()?.into_parts();
        Ok(RecordCommit {
            snapshot: Snapshot::from_inner(snapshot)?,
            changes,
        })
    }
}

/// A successful semantic commit.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
pub(crate) struct RecordCommit {
    snapshot: Snapshot,
    changes: master_layout::RecordPatch,
}

#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
impl RecordCommit {
    /// The immutable target snapshot.
    #[must_use]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The structural patch produced by the semantic edit.
    #[must_use]
    pub(crate) const fn changes(&self) -> &master_layout::RecordPatch {
        &self.changes
    }

    /// Read the committed notes-master `txStyles` round-trip package.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn notes_styles(&self) -> RecordResult<Option<notes_styles::Styles>> {
        self.snapshot.notes_styles()
    }

    /// Undo this patch against its exact target snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn undo(&self, current: &Snapshot) -> RecordResult<Snapshot> {
        Snapshot::from_inner(self.changes.undo(&current.inner)?)
    }

    /// Redo this patch against its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn redo(&self, current: &Snapshot) -> RecordResult<Snapshot> {
        Snapshot::from_inner(self.changes.redo(&current.inner)?)
    }

    /// Split the commit into its target and reusable patch.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Snapshot, master_layout::RecordPatch) {
        (self.snapshot, self.changes)
    }
}

crate::scoped::family!("master-metadata");
