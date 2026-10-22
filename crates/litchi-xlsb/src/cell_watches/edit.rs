//! Immutable cell-watch edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_cell_watches`] starts an [`Edit`] whose verbs add, remove,
//! replace or clear a worksheet's watched cells and its worksheet-wide
//! phonetic default. Opaque producer records keep their bytes and relative
//! order, and every other worksheet record is copied unchanged.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_cell_watches`] replays a patch onto an equal package. An
//! exact no-op keeps every byte, and a changed edit of a signed workbook is
//! refused with [`crate::package::PackageError::Signed`].

use super::phonetic::Info;
use super::worksheet;
use super::{Reference, Watch, Watches};
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("cell-watch", shared_error);

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Watches of one worksheet ([`super::MAX_WATCHES`]).
    Watches,
}

/// A staged cell-watch edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a cell-watch edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_cell_watches(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
        })
    }

    /// Apply a cell-watch patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::package::PackageError::PatchConflict`] when this package is not the patch's
    /// before state and [`crate::package::PackageError::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_cell_watches(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    fn snapshot(&self, sheet: usize) -> Result<worksheet::Snapshot> {
        let uri = self.staged.worksheet(sheet)?;
        worksheet::read(self.staged.blob(&uri)?, self.staged.raw_limits())
    }

    /// The staged watches of the worksheet at complete-catalog position
    /// `sheet`, in source order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing tab, a non-worksheet tab or a malformed
    /// stream.
    pub fn watches(&self, sheet: usize) -> Result<Watches> {
        Watches::new(self.snapshot(sheet)?.watches().to_vec())
    }

    /// The staged worksheet-wide phonetic default, if any.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::watches`].
    pub fn phonetic(&self, sheet: usize) -> Result<Option<Info>> {
        Ok(self.snapshot(sheet)?.phonetic())
    }

    fn stage<T>(
        &mut self,
        sheet: usize,
        operation: impl FnOnce(&mut worksheet::Edit) -> Result<T>,
    ) -> Result<T> {
        let uri = self.staged.worksheet(sheet)?;
        let mut draft = worksheet::read(self.staged.blob(&uri)?, self.staged.raw_limits())?.edit();
        let value = operation(&mut draft)?;
        let commit = draft.commit()?;
        let after = commit.patch().after().to_vec();
        self.staged.replace_worksheet(&uri, after)?;
        Ok(value)
    }

    /// Add one watched cell at the end of a worksheet's collection.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, a duplicate or
    /// invalid watch, more than [`super::MAX_WATCHES`] watches, and a changed
    /// edit of a signed workbook.
    pub fn add(&mut self, sheet: usize, watch: Watch) -> Result<()> {
        self.stage(sheet, |draft| draft.add(watch))
    }

    /// Remove the first watch of a cell reference. Returns whether one was
    /// removed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`].
    pub fn remove(&mut self, sheet: usize, reference: Reference) -> Result<bool> {
        self.stage(sheet, |draft| Ok(draft.remove(reference)))
    }

    /// Remove every typed watch of a worksheet, retaining opaque records.
    /// Returns whether any was removed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`].
    pub fn clear(&mut self, sheet: usize) -> Result<bool> {
        self.stage(sheet, |draft| Ok(draft.clear()))
    }

    /// Replace every typed watch of a worksheet, retaining every opaque slot
    /// and its relative position.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`].
    pub fn replace(&mut self, sheet: usize, watches: Watches) -> Result<()> {
        self.stage(sheet, |draft| draft.replace(watches))
    }

    /// Set or replace a worksheet's phonetic default.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`].
    pub fn set_phonetic(&mut self, sheet: usize, value: Info) -> Result<()> {
        self.stage(sheet, |draft| {
            draft.set_phonetic(value);
            Ok(())
        })
    }

    /// Remove a worksheet's phonetic default. Returns whether one was
    /// removed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`].
    pub fn clear_phonetic(&mut self, sheet: usize) -> Result<bool> {
        self.stage(sheet, |draft| Ok(draft.clear_phonetic()))
    }

    /// Validate the staged worksheets and publish them to a new package. An
    /// unchanged staging returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`crate::package::PackageError::Signed`] for a changed staging of a signed workbook
    /// and the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
