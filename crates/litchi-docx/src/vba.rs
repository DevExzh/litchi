//! The VBA-project family (record 0355): the main document's MS-OVBA
//! project and its Word supplemental data. Payloads stay inert: nothing is
//! executed.
//!
//! [`Package::vba`](crate::Package::vba) discovers the attached project;
//! [`Package::edit_vba`](crate::Package::edit_vba) stages an [`Edit`] whose
//! commit publishes a new package (a macro-enabled main content type when a
//! project is attached), and [`Package::apply_vba`](crate::Package::apply_vba)
//! replays its [`Patch`].

use crate::package::family::{graph_edit, graph_family};
use crate::vba_project::{VbaProject, VbaSupplementalData};

graph_family!("VBA-project", own_limits);

/// The typed budgets of the VBA-project family: its limit errors are
/// [`litchi_vba::Error::Limit`] carried by [`crate::Error::Vba`] (record
/// 0360).
pub use litchi_vba::LimitKind;
graph_edit!("VBA-project", edit_vba, apply_vba);

impl Edit {
    /// Attach a cache-free, inert MS-OVBA project with empty Word
    /// supplemental data, replacing any attached project.
    ///
    /// # Errors
    ///
    /// Returns a project build failure, a refusal for an encrypted source and
    /// the signed-package refusal for a change to a signed package.
    pub fn set(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        self.stage(|package| package.set_vba(project))
    }

    /// Attach a cache-free project with typed Word document-event and macro
    /// metadata, built under `limits`.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_with(
        &mut self,
        project: litchi_vba::build::Project,
        supplemental_data: &VbaSupplementalData,
        limits: &litchi_vba::Limits,
    ) -> Result<VbaProject> {
        self.stage(|package| package.set_vba_with(project, supplemental_data, limits))
    }

    /// Attach a prevalidated `vbaProject.bin` payload with typed Word
    /// supplemental data.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn put(
        &mut self,
        payload: litchi_vba::Payload,
        supplemental_data: &VbaSupplementalData,
    ) -> Result<VbaProject> {
        self.stage(|package| package.put_vba(payload, supplemental_data))
    }

    /// Remove the VBA project and supplemental-data graph and restore the
    /// macro-free main content type; returns whether a project was attached.
    ///
    /// # Errors
    ///
    /// Returns a refusal for an encrypted source and the signed-package
    /// refusal for a change to a signed package.
    pub fn clear(&mut self) -> Result<bool> {
        self.stage(crate::Package::clear_vba)
    }
}
