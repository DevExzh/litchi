//! Immutable Revision Information edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_revision_info`](crate::Package::edit_revision_info) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_revision_info`](crate::Package::apply_revision_info) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).
//!
//! The Revision Information part belongs to the presentation part, so the
//! verbs take no owner.

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::codec as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("Revision Information");

/// A staged Revision Information edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a Revision Information edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_revision_info(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_revision_info")?,
        })
    }

    /// Apply a Revision Information patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_revision_info(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_revision_info")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// The staged Revision Information part, when the presentation has one.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn revision_info(&self) -> Result<Option<Part>> {
        load(self.staged.opc())
    }

    /// Add a new Revision Information part after validating the complete graph.
    /// Existing Revision Information parts are deliberately not overwritten.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Part) -> Result<()> {
        self.staged.update(move |opc| raw::store(opc, value))
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
