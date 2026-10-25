//! Immutable slide-ID Designer-tag edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_slide_designer_tags`] starts an [`Edit`]: [`Edit::set`]
//! creates or replaces the Designer tag list of one slide's `p:sldId` and
//! [`Edit::remove`] removes it. [`Edit::commit`] returns a [`Commit`] with the
//! new package, the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_slide_designer_tags`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

use super::Tags;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("slide Designer-tag");

/// A staged slide-ID Designer-tag edit of one [`Package`]. Dropping it
/// discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a slide-ID Designer-tag edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_slide_designer_tags(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_slide_designer_tags")?,
        })
    }

    /// Apply a slide-ID Designer-tag patch to a new package; this one is
    /// unchanged. An empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_slide_designer_tags(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_slide_designer_tags",
            |_, target, patch| {
                patch.validate_changed_parts(target, |name, part| {
                    crate::presentation_properties::metadata::validate_replay_part(
                        target,
                        crate::presentation_properties::metadata::ReplayFamily::DesignerTags,
                        name,
                        part,
                    )
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    fn slide_id(&self, slide: crate::slide::Selector<'_>) -> Result<u32> {
        Ok(crate::package::resolve_slide_identity(self.staged.opc(), slide)?.1)
    }

    /// The staged Designer tag list of the selected slide, when present.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, or a
    /// malformed or ambiguous tag owner.
    pub fn tags<'s>(&self, slide: impl Into<crate::slide::Selector<'s>>) -> Result<Option<Tags>> {
        let slide_id = self.slide_id(slide.into())?;
        super::load(self.staged.opc(), slide_id)
    }

    /// Create or replace the Designer tag list of the selected slide.
    /// Returns the previous list.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, an
    /// invalid tag list, a malformed or ambiguous tag owner, or a changed
    /// edit of a signed package.
    pub fn set<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        tags: Tags,
    ) -> Result<Option<Tags>> {
        let slide_id = self.slide_id(slide.into())?;
        self.staged.update(|opc| super::store(opc, slide_id, &tags))
    }

    /// Remove the Designer tag list of the selected slide. Returns the
    /// removed list.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
    ) -> Result<Option<Tags>> {
        let slide_id = self.slide_id(slide.into())?;
        self.staged.update(|opc| super::remove(opc, slide_id))
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
