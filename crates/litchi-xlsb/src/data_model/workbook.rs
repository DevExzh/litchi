//! Workbook and package facades for the XLSB Data Model owner.

use super::{Commit, Edit, EditOptions, Patch, Snapshot};
use crate::package::error::Result;
use litchi_core::edit::Diagnostics;

impl crate::Workbook {
    /// Read the workbook Data Model records and its lazy opaque model part
    /// under this projection's [`crate::SemanticLimits::data_model_value`].
    pub fn data_model(&self) -> Result<Snapshot> {
        Snapshot::read_with_limits(&self.package, self.semantic_limits().data_model_value())
    }
}

impl crate::Package {
    /// Read the workbook Data Model records and its lazy opaque model part
    /// under the package's [`crate::SemanticLimits::data_model_value`].
    pub fn data_model(&self) -> Result<Snapshot> {
        Snapshot::read_with_limits(
            self.opc_package(),
            self.semantic_limits().data_model_value(),
        )
    }

    /// Start a Data Model edit under the package's limits.
    pub fn edit_data_model(&self) -> Result<Edit> {
        self.edit_data_model_with_options(EditOptions::default())
    }

    /// Start a Data Model edit with explicit options.
    pub fn edit_data_model_with_options(&self, options: EditOptions) -> Result<Edit> {
        let limits = options
            .limits_value()
            .unwrap_or_else(|| self.semantic_limits().data_model_value());
        Ok(Edit::new(self.clone(), limits))
    }

    /// Apply a source-checked Data Model patch to a new package; this one is
    /// unchanged.
    ///
    /// Changed signed packages are rejected until the caller explicitly
    /// removes the package signature.
    pub fn apply_data_model(&self, patch: &Patch) -> Result<Commit> {
        patch.check_source(self.opc_package())?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        let mut candidate = self.opc_package().clone();
        patch.apply(&mut candidate)?;
        Ok(Commit::new(
            self.reparse_candidate(candidate)?,
            patch.clone(),
            Diagnostics::CHANGED,
        ))
    }
}
