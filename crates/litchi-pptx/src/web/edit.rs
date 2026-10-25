//! Immutable task-pane edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_task_panes`] starts an [`Edit`] over the package; [`Edit::set`] replaces the inert Office Add-in task-pane graph and
//! [`Edit::remove`] removes it.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_task_panes`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::{Conformance, Panes};
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("task-pane");

/// A staged task-pane edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a task-pane edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_task_panes(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_task_panes")?,
        })
    }

    /// Apply a task-pane patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_task_panes(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_task_panes")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Replace the complete inert task-pane graph. Returns whether anything
    /// changed; an unchanged replacement stages nothing.
    ///
    /// # Errors
    ///
    /// Returns an error for a candidate that violates bounded MS-OWEXML or OPC invariants, or a changed edit of a signed package.
    pub fn set(&mut self, panes: Panes, conformance: Conformance) -> Result<bool> {
        self.staged.update(move |opc| {
            let patch = litchi_ooxml_common::web::plan_put(
                opc,
                panes,
                conformance,
                &litchi_ooxml_common::web::Limits::standard(),
            )?;
            let commit = patch.apply(opc)?;
            let changed = commit.diagnostics().changed();
            *opc = commit.into_package();
            Ok(changed)
        })
    }

    /// Remove the inert task-pane graph. Referenced resources are removed only
    /// when OPC ownership checks prove they are not shared. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed task-pane graph, or a changed edit of a signed package.
    pub fn remove(&mut self) -> Result<bool> {
        self.staged.update(|opc| {
            let patch = litchi_ooxml_common::web::plan_remove(
                opc,
                &litchi_ooxml_common::web::Limits::standard(),
            )?;
            let commit = patch.apply(opc)?;
            let changed = commit.diagnostics().changed();
            *opc = commit.into_package();
            Ok(changed)
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Returns the part-count limit of the patch.
    pub fn commit(self) -> Result<Commit> {
        Ok(commit_of(self.staged.commit()?))
    }
}
