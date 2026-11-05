//! Immutable ink annotation edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_ink`](crate::Package::edit_ink) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_ink`](crate::Package::apply_ink) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::package as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("ink annotation");

/// A staged ink annotation edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a ink annotation edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_ink(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_ink")?,
        })
    }

    /// Apply a ink annotation patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_ink(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_ink")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Store one validated `InkML` payload on a slide as a custom XML content part.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn add<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        inkml: &[u8],
    ) -> Result<StoredAnnotation> {
        let slide_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged
            .update(move |opc| raw::store_slide(opc, &slide_name, inkml))
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
