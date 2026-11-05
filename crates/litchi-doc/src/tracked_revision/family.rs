//! The tracked-revision edit family (record 0351 §2.3, record 0359).
//!
//! [`crate::Package::edit_tracked_revisions`] stages revision-mark edits on
//! an [`Edit`]; [`Edit::commit`] publishes a new [`crate::Package`] with an
//! exact whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_tracked_revisions`] replays the patch. The FIB,
//! piece, FKP, PLC and SPRM rewrites stay with the crate-private revision
//! engine.

use super::transaction::{Error as RecordError, Snapshot, Transaction};
use super::{Revision, RevisionKind, RevisionMetadata};

crate::scoped::family!("tracked revision");

impl From<RecordError> for Error {
    fn from(error: RecordError) -> Self {
        match error {
            RecordError::Invalid(error) => Self::Package(error),
            RecordError::Conflict => Self::PatchConflict,
        }
    }
}

/// A staged tracked-revision edit of one package. Dropping it discards it.
#[derive(Clone)]
pub struct Edit {
    base: crate::scoped::Base,
    inner: Option<Transaction>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("base", &self.base)
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn open(base: crate::scoped::Base) -> Result<Self> {
        Ok(Self { base, inner: None })
    }

    fn inner_mut(&mut self) -> Result<&mut Transaction> {
        if self.inner.is_none() {
            let snapshot = Snapshot::open_with_policy(
                self.base.to_vec()?,
                self.base.options().revision_limits_value(),
                self.base.options().protection_policy_value(),
            )?;
            self.inner = Some(snapshot.edit()?);
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "tracked revision edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<Transaction> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "tracked revision edit was not initialized".into(),
            ))
        })
    }

    /// The staged revisions in source order.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged revision tables cannot be read.
    pub fn revisions(&mut self) -> Result<Vec<Revision>> {
        Ok(self.inner_mut()?.revisions()?)
    }

    /// The staged revision-author table in its stored order.
    pub fn authors(&mut self) -> Result<&[String]> {
        Ok(self.inner_mut()?.authors())
    }

    /// Adds a revision mark to an existing main-story range.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid range or metadata; the staged state
    /// is unchanged.
    pub fn add(
        &mut self,
        start_cp: u32,
        end_cp: u32,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        Ok(self.inner_mut()?.add(start_cp, end_cp, kind, metadata)?)
    }

    /// Inserts inert plain text and marks it as an insertion or move target.
    ///
    /// `cp` must lie before the main story's final paragraph mark (MS-DOC
    /// 2.3.1).
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid position, text or metadata; the
    /// staged state is unchanged.
    pub fn add_text(
        &mut self,
        cp: u32,
        text: &str,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        Ok(self.inner_mut()?.add_text(cp, text, kind, metadata)?)
    }

    /// Replaces one revision's typed metadata without touching unrelated
    /// SPRMs. Replacing with the published metadata is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range index or invalid metadata.
    pub fn replace(&mut self, index: usize, metadata: RevisionMetadata) -> Result<Revision> {
        Ok(self.inner_mut()?.replace(index, metadata)?)
    }

    /// Removes one revision mark while retaining its text and current
    /// formatting.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range index.
    pub fn remove(&mut self, index: usize) -> Result<Revision> {
        Ok(self.inner_mut()?.remove(index)?)
    }

    /// Accepts one revision using Word redline semantics.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range index or an unsupported mark.
    pub fn accept(&mut self, index: usize) -> Result<Revision> {
        Ok(self.inner_mut()?.accept(index)?)
    }

    /// Rejects one revision using Word redline semantics.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range index or an unsupported mark.
    pub fn reject(&mut self, index: usize) -> Result<Revision> {
        Ok(self.inner_mut()?.reject(index)?)
    }

    /// Validate the staged edits and publish them as a new package.
    ///
    /// # Errors
    ///
    /// Returns an error when the protection policy refuses a changed write,
    /// the package cannot be republished, or [`Error::Signed`] for a changed
    /// signed package.
    pub fn commit(mut self) -> Result<Commit> {
        let after = self.take_inner()?.commit()?.into_snapshot().finish();
        Ok(published(self.base.publish(after)?))
    }
}

impl crate::Package {
    /// The tracked revisions of this package in source order.
    ///
    /// # Errors
    ///
    /// Returns an error when the revision tables cannot be read.
    pub fn tracked_revisions(&self) -> Result<Vec<Revision>> {
        Ok(Snapshot::open(
            self.to_bytes()?,
            self.retained_edit_options().revision_limits_value(),
        )?
        .revisions()?)
    }

    /// Start a tracked-revision edit of this package under its retained edit
    /// options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's revision tables cannot be read
    /// for editing.
    pub fn edit_tracked_revisions(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start a tracked-revision edit of this package under explicit edit
    /// options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's revision tables cannot be read
    /// for editing.
    pub fn edit_tracked_revisions_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay a tracked-revision patch onto the exact package it was made
    /// from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_tracked_revisions(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}
