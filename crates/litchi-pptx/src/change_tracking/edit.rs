//! Immutable change-tracking identifier edits of a [`Package`] (record 0351
//! §2.3).
//!
//! [`Package::edit_change_tracking`] starts an [`Edit`] whose verbs set or
//! clear a slide's creation identifier and its shapes' modification
//! identifiers. [`Edit::commit`] returns a [`Commit`] with the new package,
//! the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_change_tracking`] replays a patch onto an equal package.
//! The identifiers are document data only. An exact no-op keeps every byte,
//! and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

use super::{Draft, Id};
use crate::part_edit::{Staged, family};
use crate::shape::Key;
use crate::slide::Selector;
use crate::{Package, Result};

family!("change-tracking");

/// A staged change-tracking edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a change-tracking edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_change_tracking(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_change_tracking")?,
        })
    }

    /// Apply a change-tracking patch to a new package; this one is unchanged.
    /// An empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_change_tracking(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_change_tracking")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    fn stage(
        &mut self,
        slide: Selector<'_>,
        change: impl FnOnce(&mut Draft) -> Result<()>,
    ) -> Result<()> {
        let opc = self.staged.opc();
        let owner = crate::package::resolve_slide(opc, slide)?;
        let mut draft = super::load(opc, &owner)?.edit();
        change(&mut draft)?;
        let prepared = draft.commit()?;
        if prepared.patch().is_noop() {
            return Ok(());
        }
        self.staged
            .update(|opc| super::apply_commit(opc, prepared).map(drop))
    }

    /// The staged change-tracking identifiers of the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, or
    /// malformed extension XML.
    pub fn snapshot<'s>(&self, slide: impl Into<Selector<'s>>) -> Result<super::Snapshot> {
        let opc = self.staged.opc();
        let owner = crate::package::resolve_slide(opc, slide.into())?;
        super::load(opc, &owner)
    }

    /// Set or replace the creation identifier of the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide,
    /// malformed extension XML, or a changed edit of a signed package.
    pub fn set_creation_id<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        id: impl Into<Id>,
    ) -> Result<()> {
        let id = id.into();
        self.stage(slide.into(), |draft| {
            draft.set_creation_id(id);
            Ok(())
        })
    }

    /// Remove the creation identifier of the selected slide.
    ///
    /// # Errors
    ///
    /// As [`Self::set_creation_id`].
    pub fn clear_creation_id<'s>(&mut self, slide: impl Into<Selector<'s>>) -> Result<()> {
        self.stage(slide.into(), |draft| {
            draft.clear_creation_id();
            Ok(())
        })
    }

    /// Set the modification identifier of one shape on the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide or
    /// shape, a value another shape of the slide already uses, or a changed
    /// edit of a signed package.
    pub fn set_shape_modification_id<'s, 'k>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        shape: impl Into<Key<'k>>,
        id: impl Into<Id>,
    ) -> Result<()> {
        let (shape, id) = (shape.into(), id.into());
        self.stage(slide.into(), |draft| {
            draft.set_shape_modification_id(shape, id)
        })
    }

    /// Remove the modification identifier of one shape on the selected
    /// slide.
    ///
    /// # Errors
    ///
    /// As [`Self::set_shape_modification_id`].
    pub fn clear_shape_modification_id<'s, 'k>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        shape: impl Into<Key<'k>>,
    ) -> Result<()> {
        let shape = shape.into();
        self.stage(slide.into(), |draft| {
            draft.clear_shape_modification_id(shape)
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
