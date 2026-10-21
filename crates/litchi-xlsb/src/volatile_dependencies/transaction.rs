//! Detached, source-checked Volatile Dependencies transactions.

use std::sync::Arc;

use super::model::Dependencies;
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

crate::part_edit::limits_options!(super::SemanticLimits, "Volatile Dependencies");

/// A staged Volatile Dependencies edit of one package. Dropping it discards
/// it.
#[derive(Clone, Debug)]
pub struct Edit {
    source: crate::Package,
    before: Snapshot,
    dependencies: Option<Arc<Dependencies>>,
}

impl Edit {
    pub(crate) fn new(source: crate::Package, before: Snapshot) -> Self {
        Self {
            source,
            dependencies: before.dependencies_arc(),
            before,
        }
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.dependencies = self.before.dependencies_arc();
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
            return Ok(Commit::new(
                self.source,
                Patch::new(before.clone(), before),
                litchi_core::edit::Diagnostics::UNCHANGED,
            ));
        }
        if self.dependencies.is_some() && !self.before.can_edit() {
            return Err(unsupported_unknown_records());
        }
        if self.before.package().is_signed()
            || self.before.package().requires_signature_edit_policy()
        {
            return Err(Error::Signed);
        }
        let mut candidate = self.before.package().as_ref().clone();
        let after_dependencies = self.dependencies.clone();
        super::patch::materialize(&mut candidate, &self.before, after_dependencies.as_deref())?;
        let after = Snapshot::read_with_limits(&candidate, self.before.limits())?;
        if after.dependencies() != after_dependencies.as_deref() {
            return Err(Error::InvalidFormat(
                "Volatile Dependencies transaction readback did not match staged content",
            ));
        }
        if after.source() == self.before.source() {
            return Err(Error::InvalidFormat(
                "changed Volatile Dependencies transaction produced no source change",
            ));
        }
        self.source
            .apply_volatile_dependencies(&Patch::new(self.before, after))
    }
}

fn unsupported_unknown_records() -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "edit_volatile_dependencies",
        "the Volatile Dependencies source contains records this edit cannot preserve",
    )
}
