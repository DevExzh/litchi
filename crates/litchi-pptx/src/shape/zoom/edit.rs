//! Immutable zoom edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_zooms`] starts an [`Edit`] over the package; [`Edit::set`] replaces one slide's zoom owner and [`Edit::remove`]
//! removes it.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_zooms`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::Owner;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("zoom");

/// A staged zoom edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a zoom edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_zooms(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_zooms")?,
        })
    }

    /// Apply a zoom patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_zooms(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_zooms")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Replace the zoom owner of the selected slide. Returns the previous
    /// owner.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide,
    /// a malformed zoom graph, or a changed edit of a signed package.
    pub fn set<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        owner: Owner,
    ) -> Result<Option<Owner>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| super::store(opc, &slide, owner))
    }

    /// Remove the zoom owner of the selected slide. Returns the removed
    /// owner.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide,
    /// a malformed zoom graph, or a changed edit of a signed package.
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
    ) -> Result<Option<Owner>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| super::remove(opc, &slide))
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
