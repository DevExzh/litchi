//! In-place workbook mutators for the crate's own unit tests.
//!
//! The public API is immutable: every change is a package `Edit` family that
//! commits a new [`crate::Package`]. Unit tests written against the former
//! in-place verbs use these crate-private shims, each of which runs the
//! family edit on the workbook's package and adopts the committed
//! projection, so the tests exercise the same staging and publication code.

use crate::Workbook;
use crate::package::Package;
use crate::package::error::Result;

impl Workbook {
    fn adopt(&mut self, package: &Package) -> Result<()> {
        *self = package.materialize_workbook()?;
        Ok(())
    }

    /// Run `edit` on this workbook's package and adopt the result.
    pub(crate) fn edit_in_place<T>(
        &mut self,
        edit: impl FnOnce(&Package) -> Result<(Package, T)>,
    ) -> Result<T> {
        let (package, value) = edit(&self.to_package())?;
        self.adopt(&package)?;
        Ok(value)
    }

    pub(crate) fn set_vba(
        &mut self,
        project: litchi_vba::build::Project,
    ) -> Result<crate::package::vba_project::VbaProject> {
        self.edit_in_place(|package| {
            let mut edit = package.edit_vba()?;
            let project = edit.set(project)?;
            Ok((edit.commit()?.into_package(), project))
        })
    }

    pub(crate) fn clear_vba(&mut self) -> Result<bool> {
        self.edit_in_place(|package| {
            let mut edit = package.edit_vba()?;
            let removed = edit.clear()?;
            Ok((edit.commit()?.into_package(), removed))
        })
    }
}
