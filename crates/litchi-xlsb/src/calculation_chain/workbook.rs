//! Workbook and package facades for the XLSB Calculation Chain owner.

use super::{Commit, Edit, EditOptions, Patch, Snapshot};
use crate::package::error::Result;
use litchi_core::edit::Diagnostics;

impl crate::Workbook {
    /// Read the optional Workbook-owned Calculation Chain as a lazy opaque
    /// source snapshot, under this projection's
    /// [`crate::SemanticLimits::calculation_chain_value`].
    ///
    /// The binary record grammar is not available in the local MS-XLSB
    /// specification, so this API exposes bounded source bytes and leaves
    /// generic BIFF12 framing to the explicit `ChainPart::record_count`
    /// diagnostic. It never evaluates, recalculates, or refreshes formulas.
    pub fn calculation_chain(&self) -> Result<Snapshot> {
        Snapshot::read_with_limits(
            &self.package,
            self.semantic_limits().calculation_chain_value(),
        )
    }
}

impl crate::Package {
    /// Read the optional Workbook-owned Calculation Chain as a lazy opaque
    /// source snapshot, under the package's
    /// [`crate::SemanticLimits::calculation_chain_value`]. `max_records` is
    /// retained only for an explicit `ChainPart::record_count` diagnostic
    /// and does not reject opaque bytes.
    pub fn calculation_chain(&self) -> Result<Snapshot> {
        Snapshot::read_with_limits(
            self.opc_package(),
            self.semantic_limits().calculation_chain_value(),
        )
    }

    /// Start a Calculation Chain edit under the package's limits.
    ///
    /// The only supported mutation is removal of the optional performance
    /// cache. Payload replacement and typed cell edits are intentionally
    /// unavailable until the binary grammar is authoritative.
    pub fn edit_calculation_chain(&self) -> Result<Edit> {
        self.edit_calculation_chain_with_options(EditOptions::default())
    }

    /// Start a Calculation Chain edit with explicit options.
    pub fn edit_calculation_chain_with_options(&self, options: EditOptions) -> Result<Edit> {
        let limits = options
            .limits_value()
            .unwrap_or_else(|| self.semantic_limits().calculation_chain_value());
        Ok(Edit::new(self.clone(), limits))
    }

    /// Apply a source-checked Calculation Chain patch to a new package; this
    /// one is unchanged.
    ///
    /// Changed signed packages are rejected until the caller explicitly
    /// removes the package signature.
    pub fn apply_calculation_chain(&self, patch: &Patch) -> Result<Commit> {
        // Keep the clone behind the source and publication policy preflight.
        // The checked seam validates the detached candidate and reads it back
        // without repeating either source capture.
        let current = patch.check_source(self.opc_package())?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        patch.check_publication_policy(self.opc_package())?;
        let mut candidate = self.opc_package().clone();
        patch.apply_checked(&mut candidate, current)?;
        Ok(Commit::new(
            self.reparse_candidate(candidate)?,
            patch.clone(),
            Diagnostics::CHANGED,
        ))
    }
}
