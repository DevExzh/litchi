//! Reversible source-checked Custom XML package patches.

use crate::Result;
use litchi_opc::OpcPackage;

use super::package;
use super::snapshot::Snapshot;

/// A complete, reversible replacement of one Custom XML package graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source graph required before this patch can be applied.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Graph produced by this patch.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact semantic and physical no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias for [`Self::is_empty`].
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.is_empty()
    }

    /// Return the exact inverse patch without reinterpreting XML.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after checking the complete captured source graph.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<bool> {
        let current = Snapshot::load_scoped(target, &self.before.scope_names())?;
        if !self.before.same_source(&current) {
            return Err(super::snapshot::source_mismatch());
        }
        if self.is_empty() {
            return Ok(false);
        }

        let mut candidate = target.clone();
        package::apply_items(&mut candidate, &self.before, self.after.items())?;
        let resulting = Snapshot::load_scoped(&candidate, &self.after.scope_names())?;
        if !self.after.same_source(&resulting) {
            return Err(super::snapshot::source_mismatch());
        }
        *target = candidate;
        Ok(true)
    }
}

/// Successful publication of one Custom XML transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Whether the package graph changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
