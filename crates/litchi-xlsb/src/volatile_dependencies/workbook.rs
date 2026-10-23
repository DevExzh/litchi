//! Package facade for the Workbook-owned Volatile Dependencies owner.

use super::{Commit, Edit, EditOptions, Patch, Snapshot};
use crate::package::error::Result;
use litchi_core::edit::Diagnostics;

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
    /// let package = Package::new()?;
    /// let snapshot = package.volatile_dependencies()?;
    /// assert!(!snapshot.is_present());
    /// # Ok(())
    /// # }
    /// ```
    pub fn volatile_dependencies(&self) -> Result<Snapshot> {
        Snapshot::read_with_limits(
            self.opc_package(),
            self.semantic_limits().volatile_dependencies_value(),
        )
    }

    /// Start a Volatile Dependencies edit under the package's limits.
    /// Missing owners can be created through `replace` on the edit.
    pub fn edit_volatile_dependencies(&self) -> Result<Edit> {
        self.edit_volatile_dependencies_with_options(EditOptions::default())
    }

    /// Start a Volatile Dependencies edit with explicit options.
    pub fn edit_volatile_dependencies_with_options(&self, options: EditOptions) -> Result<Edit> {
        let limits = options
            .limits_value()
            .unwrap_or_else(|| self.semantic_limits().volatile_dependencies_value());
        Ok(Edit::new(self.clone(), limits))
    }

    /// Apply a source-checked Volatile Dependencies patch to a new package;
    /// this one is unchanged.
    pub fn apply_volatile_dependencies(&self, patch: &Patch) -> Result<Commit> {
        let mut candidate = self.opc_package().clone();
        patch.apply(&mut candidate)?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        Ok(Commit::new(
            self.reparse_candidate(candidate)?,
            patch.clone(),
            Diagnostics::CHANGED,
        ))
    }
}
