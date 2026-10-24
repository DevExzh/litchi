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
    limits: super::SemanticLimits,
    before: Option<Snapshot>,
    dependencies: Option<Arc<Dependencies>>,
}

impl Edit {
    pub(crate) fn new(source: crate::Package, limits: super::SemanticLimits) -> Self {
        Self {
            source,
            limits,
            before: None,
            dependencies: None,
        }
    }

    fn ensure_prepared(&mut self) -> Result<&Snapshot> {
        if self.before.is_none() {
            let before = Snapshot::read_with_limits(self.source.opc_package(), self.limits)?;
            self.dependencies = before.dependencies_arc();
            self.before = Some(before);
        }
        self.before.as_ref().ok_or(Error::InvalidFormat(
            "Volatile Dependencies edit is not prepared",
        ))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.dependencies = self.ensure_prepared()?.dependencies_arc();
        Ok(())
    }

    /// Immutable source snapshot used by stale-source checks.
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()
    }

    /// Borrow the currently staged hierarchy, or `None` when removal is staged.
    pub fn dependencies(&mut self) -> Result<Option<&Dependencies>> {
        self.ensure_prepared()?;
        Ok(self.dependencies.as_deref())
    }

    /// Replace or remove the complete Volatile Dependencies owner.
    pub fn replace(&mut self, dependencies: Option<Dependencies>) -> Result<bool> {
        let (limits, can_edit) = {
            let before = self.ensure_prepared()?;
            (before.limits(), before.can_edit())
        };
        if let Some(value) = dependencies.as_ref() {
            value.validate(limits)?;
        }
        if self.dependencies.as_deref() == dependencies.as_ref() {
            return Ok(false);
        }
        if dependencies.is_some() && !can_edit {
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
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        let before = self.before.take().ok_or(Error::InvalidFormat(
            "Volatile Dependencies edit is not prepared",
        ))?;
        let changed = self.dependencies.as_deref() != before.dependencies();
        if !changed {
            return Ok(Commit::new(
                self.source,
                Patch::new(before.clone(), before),
                litchi_core::edit::Diagnostics::UNCHANGED,
            ));
        }
        if self.dependencies.is_some() && !before.can_edit() {
            return Err(unsupported_unknown_records());
        }
        if before.package().is_signed() || before.package().requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = before.package().as_ref().clone();
        let after_dependencies = self.dependencies.clone();
        super::patch::materialize(&mut candidate, &before, after_dependencies.as_deref())?;
        let after = Snapshot::read_with_limits(&candidate, before.limits())?;
        if after.dependencies() != after_dependencies.as_deref() {
            return Err(Error::InvalidFormat(
                "Volatile Dependencies transaction readback did not match staged content",
            ));
        }
        if after.source() == before.source() {
            return Err(Error::InvalidFormat(
                "changed Volatile Dependencies transaction produced no source change",
            ));
        }
        self.source
            .apply_volatile_dependencies(&Patch::new(before, after))
    }
}

fn unsupported_unknown_records() -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "edit_volatile_dependencies",
        "the Volatile Dependencies source contains records this edit cannot preserve",
    )
}
