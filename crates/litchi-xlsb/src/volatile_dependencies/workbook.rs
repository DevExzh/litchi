//! Package facade for the Workbook-owned Volatile Dependencies owner.

use super::{Commit, Limits, Patch, Snapshot, Transaction};
use crate::package::error::Result;

impl crate::Package {
    /// Read the optional Workbook Volatile Dependencies part.
    ///
    /// The result is a lazy, inert snapshot. It binds `BrtVolRef.ish` to the
    /// Workbook's direct `BrtBundleSh` ordinal collection and never evaluates
    /// formulas, refreshes RTD/cube values, or contacts an external service.
    ///
    /// # Example
    ///
    /// ```
    /// # use litchi_xlsb::Package;
    /// # fn main() -> litchi_xlsb::package::PackageResult<()> {
    /// let package = Package::create()?;
    /// let snapshot = package.volatile_dependencies()?;
    /// assert!(!snapshot.is_present());
    /// # Ok(())
    /// # }
    /// ```
    pub fn volatile_dependencies(&self) -> Result<Snapshot> {
        self.volatile_dependencies_with_limits(Limits::DEFAULT)
    }

    /// Read Volatile Dependencies with explicit finite semantic limits.
    pub fn volatile_dependencies_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        Snapshot::read_with_limits(self.opc_package(), limits)
    }

    /// Start a detached source-bound Volatile Dependencies transaction.
    pub fn edit_volatile_dependencies(&self) -> Result<Transaction> {
        Ok(self.volatile_dependencies()?.edit())
    }

    /// Start a detached transaction with explicit finite semantic limits.
    pub fn edit_volatile_dependencies_with_limits(&self, limits: Limits) -> Result<Transaction> {
        Ok(self.volatile_dependencies_with_limits(limits)?.edit())
    }

    /// Apply a Volatile Dependencies commit to a cloned package.
    pub fn apply_volatile_dependencies(&self, commit: &Commit) -> Result<Self> {
        self.apply_volatile_dependencies_patch(commit.patch())
    }

    /// Apply a source-checked Volatile Dependencies patch atomically.
    pub fn apply_volatile_dependencies_patch(&self, patch: &Patch) -> Result<Self> {
        let mut candidate = self.clone().into_opc();
        patch.apply(&mut candidate)?;
        Self::from_opc_with_external_link_limits(candidate, self.external_link_limits())
    }

    /// Explicitly remove package signatures before planning a changed
    /// Volatile Dependencies publication.
    pub fn without_signatures(&self) -> Result<Self> {
        let mut candidate = self.clone().into_opc();
        candidate.unsign();
        Self::from_opc_with_external_link_limits(candidate, self.external_link_limits())
    }
}
