//! Raw OPC edits of a [`Package`] (record 0351 §2.1 rule 4).
//!
//! [`Package::edit_opc`] starts an [`Edit`] over the package's OPC graph;
//! [`Edit::update`] runs one low-level graph operation on a staged copy, and
//! [`Edit::commit`] validates the staged graph as a complete XLSB workbook
//! (the binary workbook main part, the inert VBA graph and a full reparse)
//! and publishes a new package with the reversible part-level [`Patch`] and
//! the [`Diagnostics`]. [`Package::apply_opc`] replays a patch onto an equal
//! package. The source package never changes.
//!
//! Callers that need a typed, semantic operation should prefer the package's
//! scoped edit families; this is the escape hatch for graph operations no
//! family models. A drawing-skipped projection stays drawing-skipped.
//!
//! Signature state is preserved (ADR 0006): a commit that changes a signed
//! graph is refused unless the operation removed the signature explicitly
//! with [`OpcPackage::unsign`] (or the caller called [`Package::unsign`]
//! first); deleting signature parts by hand does not authorize the change.
//! An operation that changes nothing keeps the signature byte for byte.
//!
//! ```
//! use litchi_opc::{BlobPart, PackURI};
//!
//! # fn main() -> litchi_xlsb::package::PackageResult<()> {
//! let package = litchi_xlsb::Package::new()?;
//! let name = PackURI::new("/custom/raw.bin")?;
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
//! assert!(after.opc_package().get_part(&name).is_ok());
//! let restored = after.apply_opc(&patch.inverse())?.into_package();
//! assert_eq!(restored.to_bytes()?, package.to_bytes()?);
//! # Ok(())
//! # }
//! ```

use litchi_opc::OpcPackage;

use crate::Package;
use crate::package::error::Result;
use crate::part_edit::{PartPatch, family};

family!("raw OPC");

/// The bounded dimension of raw OPC edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Parts one commit may add, change or remove: at most the package's
    /// source-limit part count.
    Parts,
}

/// A staged raw OPC edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Package,
    staged: Option<OpcPackage>,
}

impl Package {
    /// Start a raw OPC edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the family entry shape.
    pub fn edit_opc(&self) -> Result<Edit> {
        Ok(Edit {
            source: self.clone(),
            staged: None,
        })
    }

    /// Apply a raw OPC patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, [`Error::Signed`] for a changed patch of a signed
    /// package, and the workbook validation errors of the result.
    pub fn apply_opc(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// The staged OPC graph.
    #[must_use]
    pub fn opc(&self) -> &OpcPackage {
        self.staged
            .as_ref()
            .unwrap_or_else(|| self.source.opc_package())
    }

    /// Run one low-level graph operation on the staged graph.
    ///
    /// Any error leaves the staged graph as it was before the operation.
    ///
    /// # Errors
    ///
    /// Returns the operation's error.
    pub fn update<T>(&mut self, operation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut candidate = self.opc().clone();
        let value = operation(&mut candidate)?;
        self.staged = Some(candidate);
        Ok(value)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged = None;
    }

    /// Validate the staged graph as a complete workbook and publish it to a
    /// new package. An unchanged graph returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the OPC `SignedSourceRequiresExplicitPolicy` error for a
    /// changed signed graph without an explicit signature disposition, the
    /// part-count limit of the patch, and the workbook validation errors of
    /// the staged graph.
    pub fn commit(self) -> Result<Commit> {
        let Some(staged) = self.staged else {
            return Ok(Commit::new((self.source, PartPatch::default())));
        };
        let staged = std::sync::Arc::new(staged);
        let patch = PartPatch::capture_all(self.source.opc_shared(), &staged)?;
        if patch.is_noop() {
            return Ok(Commit::new((self.source, patch)));
        }
        staged.validate_signature_edit_from(self.source.opc_package())?;
        crate::Workbook::validate_edit_candidate(&staged)?;
        let package = self.source.reparse_candidate_retaining_projection(staged)?;
        Ok(Commit::new((package, patch)))
    }
}

#[cfg(test)]
mod tests {
    use litchi_opc::{BlobPart, PackURI, TargetMode};

    use crate::Package;
    use crate::package::PackageError;

    fn raw_part(name: &str) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).unwrap(),
            "application/octet-stream".into(),
            vec![1, 2, 3],
        )
    }

    #[test]
    fn starting_an_edit_copies_nothing_and_a_no_op_returns_the_source() {
        let package = Package::new().unwrap();
        let edit = package.edit_opc().unwrap();
        assert!(edit.staged.is_none());
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
        assert!(after.opc_package().rels().get("rIdRaw").is_some());
        assert!(package.opc_package().rels().get("rIdRaw").is_none());
        let replayed = package.apply_opc(&patch).unwrap().into_package();
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let restored = after.apply_opc(&patch.inverse()).unwrap().into_package();
        assert!(restored.opc_package().rels().get("rIdRaw").is_none());
        assert_eq!(restored.to_bytes().unwrap(), package.to_bytes().unwrap());
        assert!(matches!(
            after.apply_opc(&patch),
            Err(PackageError::PatchConflict { .. })
        ));
    }

    #[test]
    fn a_failed_update_and_reset_leave_the_staging_as_it_was() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_opc().unwrap();
        let error = edit
            .update(|opc| {
                opc.add_part(Box::new(raw_part("/custom/a.bin")));
                Err::<(), _>(PackageError::InvalidFormat("stop".into()))
            })
            .unwrap_err();
        assert!(matches!(error, PackageError::InvalidFormat(_)));
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
