//! Immutable embedded-font edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_fonts`] starts an [`Edit`] over the package; [`Edit::set`] replaces the presentation-owned embedded-font collection
//! and [`Edit::remove`] removes it.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_fonts`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::Fonts;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("embedded-font");

/// A staged embedded-font edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a embedded-font edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_fonts(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_fonts")?,
        })
    }

    /// Apply a embedded-font patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_fonts(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_fonts")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Replace the complete presentation-owned embedded-font collection.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed font graph or an invalid collection, or a changed edit of a signed package.
    pub fn set(&mut self, fonts: Fonts) -> Result<bool> {
        self.staged.update(move |opc| super::put(opc, fonts))
    }

    /// Remove the presentation-owned embedded-font collection and its
    /// exclusively owned parts. Returns the removed collection.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed font graph, or a changed edit of a signed package.
    pub fn remove(&mut self) -> Result<Option<Fonts>> {
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
