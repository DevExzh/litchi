//! Immutable embedded OLE object edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_ole_objects`](crate::Package::edit_ole_objects) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_ole_objects`](crate::Package::apply_ole_objects) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::authoring as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("embedded OLE object");

/// A staged embedded OLE object edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a embedded OLE object edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_ole_objects(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_ole_objects")?,
        })
    }

    /// Apply a embedded OLE object patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_ole_objects(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_ole_objects")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Add an inert embedded OLE/package payload and verify it through the read
    /// inventory before publishing the package mutation.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        kind: Kind,
        program_id: Option<&str>,
        name: Option<&str>,
        frame: Frame,
        payload: &[u8],
    ) -> Result<Authored> {
        let slide_part_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        self.staged.update(move |opc| {
            raw::add(
                opc,
                slide_part_name.as_str(),
                kind,
                program_id,
                name,
                frame,
                payload,
            )
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
