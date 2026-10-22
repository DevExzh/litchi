//! Immutable chartsheet edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_chart_sheets`] starts an [`Edit`]; every operation
//! runs on a private candidate graph, so an error stages nothing.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_chart_sheets`] replays a patch onto an equal package. The
//! source package never changes. A changed edit of a signed workbook is
//! refused with [`Error::Signed`]; an exact no-op keeps every byte.

use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};
use litchi_opc::PackURI;

family!("chartsheet");

pub use crate::opc::LimitKind;

/// A staged chartsheet edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    workbook: PackURI,
}

impl Package {
    /// Start a chartsheet edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_chart_sheets(&self) -> Result<Edit> {
        let staged = Staged::new(self, "edit_chart_sheets")?;
        let workbook = staged.opc().main_document_part()?.partname().clone();
        Ok(Edit { staged, workbook })
    }

    /// Apply a chartsheet patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_chart_sheets(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_chart_sheets")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// Add a preflighted chartsheet package graph and its workbook sheet
    /// entry in `conformance`.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the chartsheet graph and
    /// [`Error::Signed`] for a signed package.
    pub fn add(
        &mut self,
        value: &super::package::Package,
        conformance: super::Conformance,
    ) -> Result<()> {
        let workbook = self.workbook.clone();
        self.staged.update(|candidate| {
            crate::signature_policy::ensure_unsigned(candidate)?;
            super::package::store_chartsheet(candidate, &workbook, value, conformance)
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
