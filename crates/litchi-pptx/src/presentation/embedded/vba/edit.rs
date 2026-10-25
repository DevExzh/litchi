//! Immutable VBA project edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_vba`](crate::Package::edit_vba) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_vba`](crate::Package::apply_vba) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`Error::Signed`](crate::Error::Signed).
//!
//! The VBA project belongs to the presentation part, so the verbs take no
//! owner.

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::package as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("VBA project");

/// A staged VBA project edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a VBA project edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_vba(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_vba")?,
        })
    }

    /// Apply a VBA project patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_vba(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_vba")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Discover the one inert VBA graph attached to a presentation main part.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn project(&self) -> Result<Option<Project>> {
        let source = self.staged.opc().main_document_part()?.partname().clone();
        discover(self.staged.opc(), &source)
    }

    /// Store an opaque VBA payload through the shared transactional graph service.
    ///
    /// Storing always republishes the graph, so a signed package is refused with
    /// [`Error::Signed`](crate::Error::Signed)
    /// (ADR 0006) before it is changed: remove the signature explicitly first,
    /// with [`Package::unsign`](crate::Package::unsign).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a signed package, and an error if the output
    /// cannot be encoded or written.
    pub fn set(&mut self, payload: Vec<u8>) -> Result<Project> {
        let source = self.staged.opc().main_document_part()?.partname().clone();
        self.staged
            .update(move |opc| raw::store(opc, &source, payload))
    }

    /// Remove the complete VBA relationship graph and restore the non-macro main type.
    ///
    /// An absent graph returns `Ok(false)` and keeps a signed package's
    /// signature; removing a present graph from a signed package is refused with
    /// [`Error::Signed`](crate::Error::Signed)
    /// (ADR 0006): remove the signature explicitly first, with
    /// [`Package::unsign`](crate::Package::unsign).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a change to a signed package, and an error if
    /// the operation fails.
    pub fn remove(&mut self) -> Result<bool> {
        let source = self.staged.opc().main_document_part()?.partname().clone();
        self.staged.update(move |opc| raw::remove(opc, &source))
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
