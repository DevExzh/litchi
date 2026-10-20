//! Immutable InkAction edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_ink_actions`] starts an [`Edit`] whose verbs replace the
//! action profile of one existing InkAction target on one slide.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_ink_actions`] replays a patch onto an equal package. The
//! profiles are inert document data. An exact no-op keeps every byte, and a
//! changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

use litchi_drawingml::ink::actions::{Edit as ProfileEdit, Profile};

use super::transaction::Draft;
use super::{AnchorSelector, Limits};
use crate::part_edit::{Staged, family};
use crate::slide::Selector;
use crate::{Package, Result};

family!("InkAction");

/// A staged InkAction edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start an InkAction edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_ink_actions(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_ink_actions")?,
        })
    }

    /// Apply an InkAction patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_ink_actions(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_ink_actions")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// The staged InkAction snapshot of the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous or out-of-range slide, or a
    /// malformed or unsupported typed action closure.
    pub fn ink_actions<'s>(&self, slide: impl Into<Selector<'s>>) -> Result<super::Snapshot> {
        let opc = self.staged.opc();
        let (index, name) = crate::package::resolve_slide_position(opc, slide.into())?;
        super::load_snapshot(opc, index, opc.get_part(&name)?, &Limits::default())
    }

    fn stage(
        &mut self,
        slide: Selector<'_>,
        change: impl FnOnce(&mut Draft) -> Result<bool>,
    ) -> Result<bool> {
        let mut draft = self.ink_actions(slide)?.edit();
        if !change(&mut draft)? {
            return Ok(false);
        }
        let prepared = draft.commit()?;
        if prepared.patch().is_noop() {
            return Ok(false);
        }
        self.staged
            .update(|opc| super::apply_commit(opc, prepared).map(drop))?;
        Ok(true)
    }

    /// Run shared action-profile mutations on the selected target of the
    /// selected slide and stage the resulting profile. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or target, an invalid profile,
    /// a profile beyond the target bounds, or a changed edit of a signed
    /// package.
    pub fn edit_profile<'s, F>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        selector: AnchorSelector,
        edit: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut ProfileEdit) -> Result<()>,
    {
        self.stage(slide.into(), |draft| draft.edit_profile(selector, edit))
    }

    /// Replace the profile of the selected target on the selected slide
    /// with an already validated shared profile. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// As [`Self::edit_profile`].
    pub fn replace_profile<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        selector: AnchorSelector,
        profile: Profile,
    ) -> Result<bool> {
        self.stage(slide.into(), |draft| {
            draft.replace_profile(selector, profile)
        })
    }

    /// Parse bounded XML bytes and replace the profile of the selected target
    /// on the selected slide. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// As [`Self::edit_profile`], and for malformed profile XML.
    pub fn replace_profile_bytes<'s>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        selector: AnchorSelector,
        bytes: impl AsRef<[u8]>,
    ) -> Result<bool> {
        self.stage(slide.into(), |draft| {
            draft.replace_profile_bytes(selector, bytes)
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
