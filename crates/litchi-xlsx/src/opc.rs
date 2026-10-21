//! Raw clear-OPC edits of a [`Package`] (record 0351 §2.1 rule 4).
//!
//! [`Package::edit_opc`] starts an [`Edit`] over the package's clear OPC
//! graph; [`Edit::update`] runs one low-level graph operation on a staged
//! copy, and [`Edit::commit`] validates the staged graph as a complete
//! workbook and publishes a new package with the reversible part-level
//! [`Patch`] and the [`Diagnostics`]. [`Package::apply_opc`] replays a patch
//! onto an equal package. The source package never changes.
//!
//! Callers that need a typed, semantic operation should prefer the package's
//! scoped edit families; this is the escape hatch for graph operations no
//! family models.
//!
//! Signature state is preserved (ADR 0006): a commit that changes a signed
//! graph is refused with [`Error::Signed`] unless the operation removed the
//! signature explicitly with [`OpcPackage::unsign`] (or the caller called
//! [`Package::unsign`] first); deleting signature parts by hand does not
//! authorize the change. An operation that changes nothing keeps the
//! signature byte for byte.
//!
//! ```
//! use litchi_opc::{BlobPart, PackURI};
//!
//! # fn main() -> litchi_xlsx::Result<()> {
//! let package = litchi_xlsx::Package::new()?;
//! let name = PackURI::new("/custom/raw.bin")
//!     .map_err(|error| litchi_xlsx::Error::Invalid(error.to_string()))?;
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
//! assert!(!after.same_content(&package)?);
//! let restored = after.apply_opc(&patch.inverse())?.into_package();
//! assert!(restored.same_content(&package)?);
//! # Ok(())
//! # }
//! ```

use litchi_opc::OpcPackage;

use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("raw OPC");

/// The bounded dimension of raw OPC edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Parts one commit may add, change or remove: at most the package's
    /// source-limit part count.
    Parts,
}

/// A staged raw clear-OPC edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a raw clear-OPC edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package opened with a
    /// password, whose mutations are refused.
    pub fn edit_opc(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new(self, "edit_opc")?,
        })
    }

    /// Apply a raw clear-OPC patch to a new package; this one is unchanged.
    /// An empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// package.
    pub fn apply_opc(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_opc")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged clear OPC graph.
    #[must_use]
    pub fn opc(&self) -> &OpcPackage {
        self.staged.opc()
    }

    /// Run one low-level graph operation on the staged graph.
    ///
    /// Any error leaves the staged graph as it was before the operation.
    ///
    /// # Errors
    ///
    /// Returns the operation's error.
    pub fn update<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.staged.update(operation)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Validate the staged graph as a complete workbook and publish it to a
    /// new package. An unchanged graph returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed graph that still carries its
    /// signature, the part-count limit of the patch, and the workbook
    /// validation errors of the staged graph.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

#[cfg(test)]
mod tests {
    use litchi_opc::{BlobPart, PackURI, TargetMode};

    use crate::{Error, Package};

    fn raw_part(name: &str) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).unwrap(),
            "application/octet-stream".into(),
            vec![1, 2, 3],
        )
    }

    fn same_reopened_content(left: &Package, right: &Package) -> bool {
        let left = Package::from_bytes(&left.to_bytes().unwrap()).unwrap();
        let right = Package::from_bytes(&right.to_bytes().unwrap()).unwrap();
        left.same_content(&right).unwrap()
    }

    #[test]
    fn starting_an_edit_copies_nothing_and_a_no_op_returns_the_source() {
        let package = Package::new().unwrap();
        let edit = package.edit_opc().unwrap();
        assert!(edit.staged.staged_is_shared());
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit.package().to_bytes().unwrap(),
            package.to_bytes().unwrap()
        );
    }

    #[test]
    fn root_relationships_are_captured_replayed_and_inverted() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_opc().unwrap();
        edit.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/raw.bin")));
            opc.rels_mut().try_add_relationship(
                "urn:litchi:test".into(),
                "custom/raw.bin".into(),
                "rIdRaw".into(),
                TargetMode::Internal,
            )?;
            Ok(())
        })
        .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert!(after.opc().rels().get("rIdRaw").is_some());
        assert!(package.opc().rels().get("rIdRaw").is_none());
        let replayed = package.apply_opc(&patch).unwrap().into_package();
        assert!(replayed.same_content(&after).unwrap());
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let reopened_source = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let reopened_after = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let reopened_forward = reopened_source.apply_opc(&patch).unwrap().into_package();
        assert!(same_reopened_content(&reopened_forward, &after));
        let reopened_restored = reopened_after
            .apply_opc(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(same_reopened_content(&reopened_restored, &package));
        let restored = after.apply_opc(&patch.inverse()).unwrap().into_package();
        assert!(restored.opc().rels().get("rIdRaw").is_none());
        assert!(restored.same_content(&package).unwrap());
        assert_eq!(restored.to_bytes().unwrap(), package.to_bytes().unwrap());
        assert!(matches!(
            after.apply_opc(&patch),
            Err(Error::PatchConflict { .. })
        ));
    }

    #[test]
    fn a_failed_update_and_reset_leave_the_staging_as_it_was() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_opc().unwrap();
        let error = edit
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/a.bin")));
                Err::<(), _>(Error::Invalid("stop".into()))
            })
            .unwrap_err();
        assert!(matches!(error, Error::Invalid(_)));
        assert!(
            edit.opc()
                .get_part(&PackURI::new("/custom/a.bin").unwrap())
                .is_err()
        );
        edit.update(|opc| {
            opc.add_part(Box::new(raw_part("/custom/b.bin")));
            Ok(())
        })
        .unwrap();
        edit.reset();
        assert!(edit.commit().unwrap().patch().is_noop());
    }
}
