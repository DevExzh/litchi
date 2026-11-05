//! Immutable table-style edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_styles`] starts an [`Edit`] over the package; [`Edit::set`] replaces the presentation table-style catalog and
//! [`Edit::remove`] removes it.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_styles`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::List;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("table-style");

/// A staged table-style edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a table-style edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_styles(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_styles")?,
        })
    }

    /// Apply a table-style patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_styles(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_styles")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Replace the complete presentation table-style catalog. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed catalog graph or an invalid catalog, or a changed edit of a signed package.
    pub fn set(&mut self, styles: List) -> Result<bool> {
        self.staged.update(move |opc| super::put(opc, styles))
    }

    /// Remove the presentation table-style catalog. Returns the removed
    /// catalog.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed catalog graph, or a changed edit of a signed package.
    pub fn remove(&mut self) -> Result<Option<List>> {
        self.staged.update(super::remove)
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
