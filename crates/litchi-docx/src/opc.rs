//! The raw OPC family (record 0355): a low-level escape for graph edits the
//! typed families do not model.
//!
//! [`Package::edit_opc`](crate::Package::edit_opc) stages an [`Edit`] retaining
//! the source graph; the structural candidate is copied on the first mutation
//! (built-in part payloads stay shared until changed); [`Edit::commit`] validates the candidate — its Word main
//! relationship and content type, its core and custom properties and its
//! signature disposition — and publishes a new package, and
//! [`Package::apply_opc`](crate::Package::apply_opc) replays the [`Patch`].
//!
//! Signature state is preserved: changed signed content requires an explicit
//! signature disposition inside the edit ([`OpcPackage::unsign`] or a signing
//! operation); deleting signature parts alone does not authorize it.
//!
//! ```
//! use litchi_docx::Package;
//! use litchi_opc::{BlobPart, PackURI};
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_opc();
//! edit.package_mut().add_part(Box::new(BlobPart::new(
//!     PackURI::new("/customXml/extra.bin").map_err(litchi_docx::Error::Uri)?,
//!     "application/octet-stream".to_owned(),
//!     vec![1, 2, 3],
//! )));
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! assert!(!package.opc_package().contains_part(&PackURI::new("/customXml/extra.bin").unwrap()));
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use litchi_opc::OpcPackage;

use crate::package::family::graph_family;
use crate::package::graph_patch::GraphPatch;

graph_family!("raw OPC");

/// A staged raw OPC change over one package snapshot. The source is never
/// changed: [`Self::commit`] publishes a new package, and dropping the edit
/// discards it.
#[derive(Clone)]
pub struct Edit {
    source: crate::Package,
    candidate: Option<OpcPackage>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn materialize(&mut self) -> &mut OpcPackage {
        if self.candidate.is_none() {
            self.candidate = Some(self.source.opc_package().clone());
        }
        self.candidate
            .as_mut()
            .expect("candidate is materialized above")
    }

    /// The staged graph.
    #[must_use]
    pub fn package(&self) -> &OpcPackage {
        match &self.candidate {
            Some(candidate) => candidate,
            None => self.source.opc_package(),
        }
    }

    /// The staged graph, for a raw change.
    pub fn package_mut(&mut self) -> &mut OpcPackage {
        self.materialize()
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.candidate = None;
    }

    /// Validate the staged graph and publish it as a new package with its
    /// exact patch. An unchanged graph publishes the source package.
    ///
    /// # Errors
    ///
    /// Returns an error for a candidate without a valid Word main part or
    /// properties, an encrypted source, or a changed signed graph without a
    /// signature disposition.
    pub fn commit(self) -> Result<Commit> {
        let Some(candidate) = self.candidate else {
            // A never-mutated edit is already an exact no-op. Keep the source
            // graph owned by the package instead of cloning the whole OPC
            // graph merely to diff it against itself at commit time.
            let source = self.source;
            let patch = GraphPatch::between(source.opc_package(), source.opc_package())?;
            return Ok(Commit::new(source, Patch(patch), Diagnostics::UNCHANGED));
        };
        let patch = GraphPatch::between(self.source.opc_package(), &candidate)?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.source,
                Patch(patch),
                Diagnostics::UNCHANGED,
            ));
        }
        let mut next = self.source.clone();
        next.edit_opc_in_place(move |opc| {
            *opc = candidate;
            Ok(())
        })?;
        Ok(Commit::new(next, Patch(patch), Diagnostics::CHANGED))
    }
}

impl crate::Package {
    /// Stage a raw OPC change over a structural copy of this package's
    /// graph.
    #[must_use]
    pub fn edit_opc(&self) -> Edit {
        Edit {
            source: self.clone(),
            candidate: None,
        }
    }

    /// Replay a raw OPC patch onto this package, which must hold the patch's
    /// before state. A changed patch refuses a signed package.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_opc(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_opc")
    }
}
