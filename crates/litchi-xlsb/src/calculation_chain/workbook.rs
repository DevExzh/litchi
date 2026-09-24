//! Workbook and package facades for the XLSB Calculation Chain owner.

use super::{Commit, Limits, Patch, Snapshot, Transaction};
use crate::package::error::Result;

impl crate::Workbook {
    /// Read the optional Workbook-owned Calculation Chain as a lazy opaque
    /// source snapshot.
    ///
    /// The binary record grammar is not available in the local MS-XLSB
    /// specification, so this API exposes bounded source bytes and leaves
    /// generic BIFF12 framing to the explicit `ChainPart::record_count`
    /// diagnostic. It never evaluates, recalculates, or refreshes formulas.
    pub fn calculation_chain(&self) -> Result<Snapshot> {
        self.calculation_chain_with_limits(Limits::DEFAULT)
    }

    /// Read the Calculation Chain using explicit finite byte, relationship,
    /// and graph limits. `max_records` is retained only for an explicit
    /// `ChainPart::record_count` diagnostic and does not reject opaque bytes.
    pub fn calculation_chain_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        Snapshot::read_with_limits(&self.package, limits)
    }

    /// Start a detached source-bound Calculation Chain transaction.
    ///
    /// The only supported mutation is removal of the optional performance
    /// cache. Payload replacement and typed cell edits are intentionally
    /// unavailable until the binary grammar is authoritative.
    pub fn edit_calculation_chain(&self) -> Result<Transaction> {
        Ok(self.calculation_chain()?.edit())
    }

    /// Start a Calculation Chain transaction with explicit finite limits.
    pub fn edit_calculation_chain_with_limits(&self, limits: Limits) -> Result<Transaction> {
        Ok(self.calculation_chain_with_limits(limits)?.edit())
    }

    /// Atomically publish a detached Calculation Chain commit.
    pub fn apply_calculation_chain(&mut self, commit: &Commit) -> Result<Snapshot> {
        self.apply_calculation_chain_patch(commit.patch())
    }

    /// Atomically publish a source-checked Calculation Chain patch.
    ///
    /// Signed packages are rejected for changed patches. Callers must make an
    /// explicit signature disposition before retrying; this method never
    /// silently removes signatures.
    pub fn apply_calculation_chain_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        // Prove the complete source closure before cloning the package. The
        // checked seam validates the detached candidate and reads it back,
        // without repeating the unchanged source check.
        let current = patch.check_source(&self.package)?;
        if !patch.is_empty() {
            patch.check_publication_policy(&self.package)?;
        } else {
            return Ok(current);
        }
        let mut candidate = self.package.clone();
        let resulting = patch.apply_checked(&mut candidate, current)?;
        let validated = self.reparse_candidate(candidate)?;
        *self = validated;
        Ok(resulting)
    }
}

impl crate::Package {
    /// Read the optional Workbook-owned Calculation Chain as a lazy opaque
    /// source snapshot.
    pub fn calculation_chain(&self) -> Result<Snapshot> {
        self.calculation_chain_with_limits(Limits::DEFAULT)
    }

    /// Read the Calculation Chain using explicit finite byte, relationship,
    /// and graph limits. `max_records` is retained only for an explicit
    /// `ChainPart::record_count` diagnostic and does not reject opaque bytes.
    pub fn calculation_chain_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        Snapshot::read_with_limits(self.opc_package(), limits)
    }

    /// Start a detached source-bound Calculation Chain transaction.
    pub fn edit_calculation_chain(&self) -> Result<Transaction> {
        Ok(self.calculation_chain()?.edit())
    }

    /// Start a Calculation Chain transaction with explicit finite limits.
    pub fn edit_calculation_chain_with_limits(&self, limits: Limits) -> Result<Transaction> {
        Ok(self.calculation_chain_with_limits(limits)?.edit())
    }

    /// Apply a Calculation Chain commit to a cloned package and return the
    /// validated result.
    pub fn apply_calculation_chain(&self, commit: &Commit) -> Result<Self> {
        self.apply_calculation_chain_patch(commit.patch())
    }

    /// Apply a source-checked Calculation Chain patch to a cloned package.
    ///
    /// Changed signed packages are rejected until the caller explicitly
    /// removes or otherwise handles the package signature.
    pub fn apply_calculation_chain_patch(&self, patch: &Patch) -> Result<Self> {
        // Keep the public facade's clone behind the source and publication
        // policy preflight. The checked seam validates the detached candidate
        // and reads it back without repeating either source capture.
        let current = patch.check_source(self.opc_package())?;
        if !patch.is_empty() {
            patch.check_publication_policy(self.opc_package())?;
        }
        let mut candidate = self.clone().into_opc();
        patch.apply_checked(&mut candidate, current)?;
        Self::from_opc_with_external_link_limits(candidate, self.external_link_limits())
    }
}
