//! Workbook and package facades for the XLSB Data Model owner.

use super::{Commit, Limits, Patch, Snapshot};
use crate::package::error::Result;

impl crate::Workbook {
    /// Read the workbook Data Model records and its lazy opaque model part.
    pub fn data_model(&self) -> Result<Snapshot> {
        self.data_model_with_limits(Limits::DEFAULT)
    }

    /// Read the workbook Data Model using explicit finite limits.
    pub fn data_model_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        Snapshot::read_with_limits(&self.package, limits)
    }

    /// Start a detached, source-bound Data Model transaction.
    pub fn edit_data_model(&self) -> Result<super::Transaction> {
        Ok(self.data_model()?.edit())
    }

    /// Start a detached Data Model transaction with explicit limits.
    pub fn edit_data_model_with_limits(&self, limits: Limits) -> Result<super::Transaction> {
        Ok(self.data_model_with_limits(limits)?.edit())
    }

    /// Atomically publish one detached Data Model commit.
    pub fn apply_data_model(&mut self, commit: &Commit) -> Result<Snapshot> {
        self.apply_data_model_patch(commit.patch())
    }

    /// Atomically publish one reversible, source-checked Data Model patch.
    pub fn apply_data_model_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        let current = patch.check_source(&self.package)?;
        if patch.is_empty() {
            return Ok(current);
        }
        if self.package.is_signed() || self.package.requires_signature_edit_policy() {
            return Err(crate::package::error::Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        self.edit_opc(|candidate| {
            patch.materialize(candidate)?;
            let resulting = Snapshot::read_with_limits(candidate, patch.before().limits())?;
            if &resulting != patch.after() {
                return Err(crate::package::error::Error::InvalidFormat(
                    "Data Model publication changed the planned semantic or owned graph"
                        .to_string(),
                ));
            }
            Ok(resulting)
        })
    }
}

impl crate::Package {
    /// Read the workbook Data Model records and its lazy opaque model part.
    pub fn data_model(&self) -> Result<Snapshot> {
        Snapshot::read_with_limits(self.opc_package(), Limits::DEFAULT)
    }

    /// Read the workbook Data Model using explicit finite limits.
    pub fn data_model_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        Snapshot::read_with_limits(self.opc_package(), limits)
    }

    /// Apply a Data Model patch to a cloned package and return the validated result.
    pub fn apply_data_model_patch(&self, patch: &Patch) -> Result<Self> {
        patch.check_source(self.opc_package())?;
        if patch.is_empty() {
            return Ok(self.clone());
        }
        if self.opc_package().is_signed() || self.opc_package().requires_signature_edit_policy() {
            return Err(crate::package::error::Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = self.clone().into_opc();
        patch.apply(&mut candidate)?;
        Self::from_opc_with_external_link_limits(candidate, self.external_link_limits())
    }

    /// Apply a detached Data Model commit to a cloned package.
    pub fn apply_data_model(&self, commit: &Commit) -> Result<Self> {
        self.apply_data_model_patch(commit.patch())
    }
}
