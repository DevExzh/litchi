//! Reversible source-checked patches for worksheet OLE metadata.

use litchi_opc::OpcPackage;

use super::snapshot::Snapshot;
use super::validation;
use crate::error::{Error, Result};

/// A reversible replacement of the known OLE metadata spans in one worksheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before this patch can be applied.
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

    /// Apply atomically after checking the worksheet and opaque resources.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target, self.before.worksheet())?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        let mut candidate = target.clone();
        candidate
            .get_part_mut(self.after.worksheet())?
            .set_blob(self.after.source_xml().to_vec());
        validation::graph(&candidate, self.after.worksheet())?;
        let resulting = Snapshot::load(&candidate, self.after.worksheet())?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: self.after.worksheet().to_string(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful publication of one worksheet OLE transaction.
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

    /// Whether typed OLE metadata changed.
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
