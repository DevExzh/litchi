//! Reversible source-checked patches for workbook Custom XML Maps.

use litchi_core::sheet::Result;
use litchi_opc::OpcPackage;

use super::invalid;
use super::snapshot::Snapshot;

/// A reversible replacement of the workbook's Custom XML Maps graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before applying this patch.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by this patch.
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact source no-op.
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after checking the workbook and owned-part source.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("custom XML maps patch source is stale"));
        }
        if self.is_empty() {
            return Ok(());
        }

        let mut candidate = target.clone();
        self.after.restore_into(&mut candidate)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after)
            || resulting.info() != self.after.info()
            || resulting.conformance() != self.after.conformance()
        {
            return Err(invalid(
                "custom XML maps patch publication changed its source",
            ));
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful publication of one Custom XML Maps transaction.
#[derive(Debug)]
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

    /// Whether the workbook Custom XML Maps owner changed.
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting immutable source snapshot.
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible source-checked patch.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the result into its snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
