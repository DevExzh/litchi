//! Immutable calculation-chain edits of a [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_calculation_chain`] starts an [`Edit`]; every operation
//! runs on a private candidate graph, so an error stages nothing.
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_calculation_chain`] replays a patch onto an equal package. The
//! source package never changes. A changed edit of a signed workbook is
//! refused with [`Error::Signed`]; an exact no-op keeps every byte.

use super::model::{Chain, Conformance};
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("calculation-chain");

pub use crate::opc::LimitKind;

/// A staged calculation-chain edit of one [`Package`]. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a calculation-chain edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_calculation_chain(&self) -> Result<Edit> {
        let staged = Staged::new(self, "edit_calculation_chain")?;
        Ok(Edit { staged })
    }

    /// Apply a calculation-chain patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_calculation_chain(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_calculation_chain")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged calculation chain with its conformance.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the chain graph.
    pub fn chain(&self) -> Result<Option<(Chain, Conformance)>> {
        super::package::load(self.staged.opc())
    }

    /// Store a caller-authored inert calculation chain in `conformance`;
    /// formulas are neither recalculated nor inspected. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns the validation errors and [`Error::Signed`] for a change to a
    /// signed package.
    pub fn set(&mut self, chain: &Chain, conformance: Conformance) -> Result<bool> {
        self.staged
            .update(|candidate| super::package::put(candidate, chain, conformance))
    }

    /// Remove the calculation chain and its unreferenced part. Returns
    /// whether it was present.
    ///
    /// # Errors
    ///
    /// Returns the validation errors and [`Error::Signed`] for a change to a
    /// signed package.
    pub fn remove(&mut self) -> Result<bool> {
        self.staged.update(super::package::remove)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged graph to a new package. An unchanged edit returns
    /// the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed signed package and the
    /// workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
