//! Immutable slide ActiveX-control edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_slide_controls`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`Draft`] of one ActiveX control of one slide. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_slide_controls`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    clippy::missing_errors_doc,
    reason = "each verb documents the raw operation it stages"
)]

use super::Draft;
#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("slide ActiveX-control");

/// A staged slide ActiveX-control edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a slide ActiveX-control edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_slide_controls(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_slide_controls")?,
        })
    }

    /// Apply a slide ActiveX-control patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_slide_controls(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_slide_controls")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Stage one checked [`Draft`] of one ActiveX control of one slide: `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<'s, T>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        control: litchi_core::Position,
        operation: impl FnOnce(&mut Draft) -> Result<T>,
    ) -> Result<T> {
        let (slide_index, _) =
            crate::package::resolve_slide_position(self.staged.opc(), slide.into())?;
        let control_index = control.get();
        self.staged.update(move |opc| {
            let mut draft = {
                let name = crate::package::resolve_slide(opc, slide_index.into())?;
                let slide = opc.get_part(&name)?;
                load(
                    opc,
                    slide_index,
                    slide,
                    control_index,
                    &mut super::super::Limits::default(),
                )?
            }
            .edit();
            let value = operation(&mut draft)?;
            package::apply_commit(opc, draft.commit()?)?;
            Ok(value)
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
