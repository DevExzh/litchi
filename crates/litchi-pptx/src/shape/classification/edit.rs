//! Immutable shape-classification edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_shape_classification`] starts an [`Edit`] over the package; [`Edit::set`] stores one shape's typed classification outcome and
//! [`Edit::remove`] removes it.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_shape_classification`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::{Outcome, Snapshot};
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("shape-classification");

/// A staged shape-classification edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a shape-classification edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_shape_classification(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_shape_classification")?,
        })
    }

    /// Apply a shape-classification patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_shape_classification(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_shape_classification",
            |package, target, patch| {
                patch.validate_changed_shape_parts(target, package.semantic_limits().shape())
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Store the classification outcome of the selected shape. Returns the
    /// previous snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide or
    /// shape, an invalid outcome, malformed markup, or a changed edit of a signed package.
    pub fn set<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        outcome: Outcome,
    ) -> Result<Option<Snapshot>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let shape = shape.into();
        let shape_limits = self.staged.shape_limits();
        self.staged
            .update(move |opc| super::put_with_limits(opc, &slide, shape, outcome, shape_limits))
    }

    /// Remove the classification of the selected shape. Returns the removed
    /// snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide or
    /// shape, malformed markup, or a changed edit of a signed package.
    pub fn remove<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<Snapshot>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let shape = shape.into();
        let shape_limits = self.staged.shape_limits();
        self.staged
            .update(move |opc| super::remove_with_limits(opc, &slide, shape, shape_limits))
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
