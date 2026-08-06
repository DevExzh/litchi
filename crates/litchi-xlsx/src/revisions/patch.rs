//! Reversible, source-checked patches for workbook revision metadata.

use litchi_opc::OpcPackage;

use crate::error::{Error, Result};

use super::package::{replace_workbook_revisions, restore_snapshot_source};
use super::snapshot::Snapshot;

/// A reversible replacement of the complete workbook revision owner.
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

    /// Whether this patch preserves the exact source owner.
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

    /// Apply atomically after checking the complete revision source graph.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.workbook_part_name().to_owned(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }

        let mut candidate = target.clone();
        replace_workbook_revisions(
            &mut candidate,
            self.after.revisions(),
            self.after
                .conformance()
                .unwrap_or(super::model::RevisionConformance::Transitional),
        )?;
        restore_snapshot_source(&mut candidate, &self.after)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: self.after.workbook_part_name().to_owned(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful publication of one revision transaction.
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

    /// Whether the revision owner changed.
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
