//! Immutable VBA project edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_vba`] starts an [`Edit`] of the workbook's attached
//! MS-OVBA project; [`Edit::set`] attaches a cache-free authored project,
//! [`Edit::put`] a prevalidated `vbaProject.bin`, and [`Edit::clear`]
//! removes the project and its declared project-signature parts. Project
//! bytes are never inspected, verified or executed. Replacing a project
//! removes its legacy or Agile project signature, which the new bytes would
//! invalidate. [`Edit::commit`] returns a [`Commit`] with the new package,
//! the exact part-level [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_vba`] replays a patch onto an equal package. A changed
//! edit of a signed workbook is refused with [`Error::Signed`].

use crate::package::Package;
use crate::package::error::Result;
use crate::package::vba_project::{
    VbaProject, discover_vba_project, remove_vba_project, store_vba_project,
};
use crate::part_edit::{Staged, family};

family!("VBA project");

pub use crate::opc::LimitKind;

/// A staged VBA project edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Discover the attached MS-XLSB VBA project and its declared
    /// signatures without inspecting their bytes.
    ///
    /// # Errors
    ///
    /// Returns the VBA relationship-graph errors.
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let workbook = self.opc_package().main_document_part()?;
        discover_vba_project(self.opc_package(), workbook)
    }

    /// Start a VBA project edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_vba(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::unparsed(self)?,
        })
    }

    /// Apply a VBA project patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_vba(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply_unparsed(self)?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged VBA project, if any.
    ///
    /// # Errors
    ///
    /// Returns the VBA relationship-graph errors.
    pub fn project(&self) -> Result<Option<VbaProject>> {
        let workbook = self.staged.opc().main_document_part()?;
        discover_vba_project(self.staged.opc(), workbook)
    }

    /// Attach a cache-free, inert MS-OVBA project with the default limits.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_with`].
    pub fn set(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        self.set_with(project, &litchi_vba::Limits::default())
    }

    /// Attach a cache-free project with explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns the VBA authoring errors, the VBA graph errors and a changed
    /// edit of a signed workbook.
    pub fn set_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<VbaProject> {
        self.put(project.finish(limits)?)
    }

    /// Attach a prevalidated `vbaProject.bin` without executing it.
    ///
    /// # Errors
    ///
    /// Returns the VBA graph errors and a changed edit of a signed workbook.
    pub fn put(&mut self, payload: litchi_vba::Payload) -> Result<VbaProject> {
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        self.staged.update(|candidate| {
            let source = candidate.main_document_part()?.partname().clone();
            store_vba_project(candidate, &source, payload)
        })
    }

    /// Remove the VBA project and every declared project-signature part.
    /// Returns whether a project was present.
    ///
    /// # Errors
    ///
    /// Returns the VBA graph errors and a changed edit of a signed workbook.
    pub fn clear(&mut self) -> Result<bool> {
        if self.project()?.is_none() {
            return Ok(false);
        }
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        self.staged.update(|candidate| {
            let source = candidate.main_document_part()?.partname().clone();
            remove_vba_project(candidate, &source)
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged project graph to a new package. An unchanged edit
    /// returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
