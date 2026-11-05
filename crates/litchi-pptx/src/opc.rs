//! Raw clear-OPC edits of a [`Package`] (record 0351 §2.1 rule 4).
//!
//! [`Package::edit_opc`] starts an [`Edit`] over a staged copy of the
//! package's clear OPC graph; [`Edit::update`] runs one low-level graph
//! operation on it, and [`Edit::commit`] publishes a new package with the
//! reversible part-level [`Patch`] and the [`Diagnostics`].
//! [`Package::apply_opc`] replays a patch onto an equal package. The source
//! package never changes.
//!
//! Callers that need a typed, semantic operation should prefer the package's
//! scoped edit families. For a package with retained encryption provenance,
//! choosing this raw escape hatch explicitly declassifies the result: a
//! commit after at least one successful [`Edit::update`] publishes a package
//! without encryption provenance.
//!
//! Signature state is preserved (ADR 0006). An operation that changes a
//! signed graph must remove the signature explicitly with
//! [`OpcPackage::unsign`], or the caller calls [`Package::unsign`] first;
//! deleting signature parts alone does not authorize the change of a
//! signature the package was opened with. An operation that changes nothing
//! keeps the signature byte for byte.
//!
//! ```
//! use litchi_opc::{BlobPart, PackURI};
//!
//! # fn main() -> litchi_pptx::Result<()> {
//! let package = litchi_pptx::Package::new()?;
//! let name = PackURI::new("/custom/raw.bin").map_err(litchi_pptx::Error::Uri)?;
//! let mut edit = package.edit_opc()?;
//! edit.update(|opc| {
//!     opc.add_part(Box::new(BlobPart::new(
//!         name.clone(),
//!         "application/octet-stream".into(),
//!         vec![1, 2, 3],
//!     )));
//!     Ok(())
//! })?;
//! let (after, patch) = edit.commit()?.into_parts();
//! assert!(after.opc()?.get_part(&name).is_ok());
//! assert!(package.opc()?.get_part(&name).is_err());
//! let restored = after.apply_opc(&patch.inverse())?.into_package();
//! assert!(restored.same_content(&package)?);
//! # Ok(())
//! # }
//! ```

use litchi_opc::OpcPackage;

use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("raw OPC");

/// A staged raw clear-OPC edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
    updated: bool,
}

impl Package {
    /// Start a raw clear-OPC edit.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the family's entry shape.
    pub fn edit_opc(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::raw(self),
            updated: false,
        })
    }

    /// Apply a raw clear-OPC patch to a new package; this one is unchanged.
    /// An empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_opc(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_opc")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// The staged clear OPC graph.
    ///
    /// For a package with retained encryption provenance, this is plaintext.
    #[must_use]
    pub const fn opc(&self) -> &OpcPackage {
        self.staged.opc()
    }

    /// Run one low-level graph operation on the staged graph.
    ///
    /// Any error leaves the staged graph as it was before the operation.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, and [`Error::Opc`] with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// for a changed signed graph that still carries its signature.
    pub fn update<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let value = self.staged.update(|opc| {
            let signed_source = opc
                .is_signed()
                .then(|| litchi_opc::PackageWriter::to_bytes(opc))
                .transpose()?;
            let value = operation(opc)?;
            // Unless the operation removed the signature explicitly, it must
            // have changed nothing: a signed candidate that publishes other
            // bytes is refused, and one the writer refuses to publish (a
            // changed signed source) is refused with the writer's error.
            if let Some(source) = signed_source
                && crate::package::signature::requires_disposition(opc)
                && litchi_opc::PackageWriter::to_bytes(opc)? != source
            {
                return Err(Error::Opc(
                    litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
                ));
            }
            Ok(value)
        })?;
        self.updated = true;
        Ok(value)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.updated = false;
    }

    /// Publish the staged graph to a new package. After at least one
    /// successful [`Self::update`] the package has no encryption provenance.
    ///
    /// # Errors
    ///
    /// Returns the part-count limit of the patch.
    pub fn commit(self) -> Result<Commit> {
        let (mut package, patch) = self.staged.commit()?;
        if self.updated {
            package.declassify();
        }
        Ok(commit_of((package, patch)))
    }
}
