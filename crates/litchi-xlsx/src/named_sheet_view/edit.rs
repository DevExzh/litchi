//! Immutable Named Sheet View edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_named_sheet_views`] starts an [`Edit`] of one worksheet; every operation
//! runs on a private candidate graph, so an error stages nothing.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_named_sheet_views`] replays a patch onto an equal package. The
//! source package never changes. A changed edit of a signed workbook is
//! refused with [`Error::Signed`]; an exact no-op keeps every byte.

use super::model::Views;
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("Named Sheet View");

pub use crate::opc::LimitKind;

/// A staged Named Sheet View edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    worksheet: litchi_opc::PackURI,
}

impl Package {
    /// Start a Named Sheet View edit of one worksheet.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, a `NotFound` refusal for a
    /// selector that matches no sheet, and [`Error::NotWorksheet`] for a
    /// chart or dialog sheet.
    pub fn edit_named_sheet_views<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Edit> {
        let staged = Staged::new(self, "edit_named_sheet_views")?;
        let worksheet = self.worksheet_part(sheet, "edit_named_sheet_views")?;
        Ok(Edit { staged, worksheet })
    }

    /// Apply a Named Sheet View patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_named_sheet_views(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_named_sheet_views")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged Named Sheet Views of the worksheet.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the worksheet's graph.
    pub fn views(&self) -> Result<Option<Views>> {
        super::package::load_worksheet_named_sheet_views(self.staged.opc(), &self.worksheet)
    }

    /// Store or replace the worksheet's Named Sheet Views.
    ///
    /// # Errors
    ///
    /// Returns the validation errors and [`Error::Signed`] for a signed
    /// package.
    pub fn set(&mut self, value: &Views) -> Result<()> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::store_worksheet_named_sheet_views(candidate, &worksheet, value)
        })
    }

    /// Remove the worksheet's Named Sheet Views. Returns whether they were
    /// present.
    ///
    /// # Errors
    ///
    /// Returns the validation errors and [`Error::Signed`] for a change to a
    /// signed package.
    pub fn remove(&mut self) -> Result<bool> {
        let worksheet = self.worksheet.clone();
        self.staged.update(|candidate| {
            super::package::remove_worksheet_named_sheet_views(candidate, &worksheet)
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged graph to a new package. An unchanged edit returns
    /// the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed signed package and the
    /// workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
