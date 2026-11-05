//! Immutable 3D-model edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_model3d`] starts an [`Edit`] over the package; [`Edit::set`] stores one shape's 3D model and [`Edit::remove`] removes
//! it.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_model3d`]
//! replays a patch onto an equal package. The source package never changes,
//! an exact no-op keeps every byte, and a changed edit of a signed package is
//! refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy):
//! remove the signature explicitly first with
//! [`Package::unsign`](crate::Package::unsign).

use super::Model;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("3D-model");

/// A staged 3D-model edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a 3D-model edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_model3d(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_model3d")?,
        })
    }

    /// Apply a 3D-model patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_model3d(&self, patch: &Patch) -> Result<Commit> {
        let package =
            patch
                .0
                .apply_with_replay_validation(self, "apply_model3d", |_, target, patch| {
                    let limits = super::codec::Limits::from_source_limits(target.read_limits());
                    patch.validate_changed_parts(target, |name, part| {
                        super::package::validate_replay_part(target, name, part, limits)
                    })
                })?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Store the 3D model of the selected shape. Returns the previous model.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide or
    /// shape, an invalid model, malformed markup, or a changed edit of a signed package.
    pub fn set<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        model: Model,
    ) -> Result<Option<Model>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let shape = shape.into();
        let limits = super::codec::Limits::from_source_limits(self.staged.opc().read_limits());
        self.staged
            .update(move |opc| super::package::put(opc, &slide, shape, model, limits))
    }

    /// Remove the 3D model of the selected shape. Returns the removed model.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide or
    /// shape, malformed markup, or a changed edit of a signed package.
    pub fn remove<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<Model>> {
        let slide = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let shape = shape.into();
        let limits = super::codec::Limits::from_source_limits(self.staged.opc().read_limits());
        self.staged
            .update(move |opc| super::package::remove(opc, &slide, shape, limits))
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
