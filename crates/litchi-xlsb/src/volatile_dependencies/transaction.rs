//! Detached, source-checked Volatile Dependencies transactions.

use std::sync::Arc;

use super::model::Dependencies;
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

/// A bounded draft detached from an immutable Volatile Dependencies snapshot.
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    dependencies: Option<Arc<Dependencies>>,
}

impl Transaction {
    pub(crate) fn new(before: Snapshot) -> Self {
        Self {
            dependencies: before.dependencies_arc(),
            before,
        }
    }

    /// Immutable source snapshot used by stale-source checks.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged hierarchy, or `None` when removal is staged.
    pub fn dependencies(&self) -> Option<&Dependencies> {
        self.dependencies.as_deref()
    }

    /// Replace or remove the complete Volatile Dependencies owner.
    pub fn replace(&mut self, dependencies: Option<Dependencies>) -> Result<bool> {
        if let Some(value) = dependencies.as_ref() {
            value.validate(self.before.limits())?;
        }
        if self.dependencies.as_deref() == dependencies.as_ref() {
            return Ok(false);
        }
        if dependencies.is_some() && !self.before.can_edit() {
            return Err(unsupported_unknown_records());
        }
        self.dependencies = dependencies.map(Arc::new);
        Ok(true)
    }

    /// Replace the complete typed hierarchy.
    pub fn set_dependencies(&mut self, dependencies: Dependencies) -> Result<bool> {
        self.replace(Some(dependencies))
    }

    /// Remove the optional owner part.
    pub fn remove(&mut self) -> Result<bool> {
        self.replace(None)
    }

    /// Commit the detached draft into an atomic, reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.dependencies.as_deref() != self.before.dependencies();
        if !changed {
            let before = self.before;
            return Ok(Commit::new(Patch::new(before.clone(), before), false));
        }
        if self.dependencies.is_some() && !self.before.can_edit() {
            return Err(unsupported_unknown_records());
        }
        if self.before.package().is_signed()
            || self.before.package().requires_signature_edit_policy()
        {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = self.before.package().as_ref().clone();
        let after_dependencies = self.dependencies.clone();
        super::patch::materialize(&mut candidate, &self.before, after_dependencies.as_deref())?;
        let after = Snapshot::read_with_limits(&candidate, self.before.limits())?;
        if after.dependencies() != after_dependencies.as_deref() {
            return Err(Error::InvalidFormat(
                "Volatile Dependencies transaction readback did not match staged content"
                    .to_string(),
            ));
        }
        if after.source() == self.before.source() {
            return Err(Error::InvalidFormat(
                "changed Volatile Dependencies transaction produced no source change".to_string(),
            ));
        }
        Ok(Commit::new(Patch::new(self.before, after), true))
    }
}

fn unsupported_unknown_records() -> Error {
    Error::UnsupportedFeature(
        "typed Volatile Dependencies edits are refused while unsupported source records are present; an exact no-op or owner removal remains safe".to_string(),
    )
}
