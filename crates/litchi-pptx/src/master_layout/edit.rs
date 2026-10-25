//! Immutable slide-master and slide-layout edits of a [`Package`] (record
//! 0351 §2.3).
//!
//! [`Package::edit_master_layout`] starts an [`Edit`]: [`Edit::add_slide_master`]
//! and [`Edit::add_slide_layout`] author new owners, [`Edit::remove_slide_layout`]
//! deletes an unreferenced layout and [`Edit::set_placeholder_shape`] creates
//! or updates one placeholder shape. Masters and layouts are selected by
//! their typed IDs ([`MasterId`], [`LayoutId`]); placeholder owners by
//! [`Owner`]. [`Edit::commit`] returns a [`Commit`]
//! with the new package, the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_master_layout`] replays a patch onto an equal package.
//! The source package never changes, an exact no-op keeps every byte, and a
//! changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).
//!
//! ```
//! use litchi_pptx::master_layout::SlideLayoutKind;
//!
//! # fn main() -> litchi_pptx::Result<()> {
//! let package = litchi_pptx::Package::new()?;
//! let mut edit = package.edit_master_layout()?;
//! let master = edit.add_slide_master()?;
//! let master_id = litchi_pptx::slide::MasterId(master.master_id);
//! let layout = edit.add_slide_layout(master_id, SlideLayoutKind::Blank, "Plain", &[])?;
//! let after = edit.commit()?.into_package();
//! assert!(after.opc()?.get_part(&layout.part_name).is_ok());
//! assert!(package.opc()?.get_part(&layout.part_name).is_err());
//! # Ok(())
//! # }
//! ```

use super::{AuthoredSlideLayout, AuthoredSlideMaster, PlaceholderSpec, SlideLayoutKind};
use crate::part_edit::{Staged, family};
use crate::slide::{LayoutId, MasterId, Owner};
use crate::{Package, Result};

family!("slide-master and slide-layout");

/// A staged slide-master and slide-layout edit of one [`Package`]. Dropping
/// it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a slide-master and slide-layout edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_master_layout(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_master_layout")?,
        })
    }

    /// Apply a slide-master and slide-layout patch to a new package; this one
    /// is unchanged. An empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_master_layout(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_master_layout",
            |package, target, patch| {
                patch.validate_changed_shape_parts(target, package.semantic_limits().shape())
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Author a new slide master with its theme.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed master graph, exhausted identifiers,
    /// or a changed edit of a signed package.
    pub fn add_slide_master(&mut self) -> Result<AuthoredSlideMaster> {
        self.staged.update(super::add_slide_master)
    }

    /// Author a new slide layout owned by the master `master`.
    ///
    /// The complete preflight (names, identities, the master's layout list
    /// and the generated XML limits) runs before the staged graph changes.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing master, an invalid name or placeholder,
    /// exhausted identifiers, or a changed edit of a signed package.
    pub fn add_slide_layout(
        &mut self,
        master: MasterId,
        kind: SlideLayoutKind,
        name: &str,
        placeholders: &[PlaceholderSpec],
    ) -> Result<AuthoredSlideLayout> {
        let master_part_name = &super::package::resolve_master(self.staged.opc(), master)?;
        super::preflight_add_slide_layout(
            self.staged.opc(),
            master_part_name,
            kind,
            name,
            placeholders,
        )?;
        self.staged
            .update(|opc| super::add_slide_layout(opc, master_part_name, kind, name, placeholders))
    }

    /// Delete a slide layout that no slide references, with its master's
    /// layout-list entry and relationship.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or referenced layout, a layout no
    /// master owns, or a changed edit of a signed package.
    pub fn remove_slide_layout(&mut self, layout: LayoutId) -> Result<()> {
        let layout_part_name = &super::package::resolve_layout(self.staged.opc(), layout)?;
        self.staged
            .update(|opc| super::remove_slide_layout(opc, layout_part_name))
    }

    /// Create or update one placeholder shape in the slide, layout or master
    /// `owner`. A shape that already matches stages nothing.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or foreign owner part, an invalid
    /// specification, markup this route cannot edit losslessly, or a changed
    /// edit of a signed package.
    pub fn set_placeholder_shape<'o>(
        &mut self,
        owner: impl Into<Owner<'o>>,
        spec: &PlaceholderSpec,
    ) -> Result<()> {
        let part_name = &super::package::resolve_owner(self.staged.opc(), owner.into())?;
        let shape_limits = self.staged.shape_limits();
        if super::preflight_store_placeholder_shape_with_shape_limits(
            self.staged.opc(),
            part_name,
            spec,
            shape_limits,
        )? {
            return Ok(());
        }
        self.staged.update(|opc| {
            super::store_placeholder_shape_with_shape_limits(opc, part_name, spec, shape_limits)
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
