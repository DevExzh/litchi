//! Exact reversible value-only worksheet patches.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::{Error, Result, invalid};

/// Exact source-bound replacement of one safe worksheet value closure.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(super) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required exact source state.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact target state.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether application preserves the selected source bytes exactly.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Exact source-bound inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically to an owning OPC package after exact source validation.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        if !self.before.matches_current_source(package) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet_part_name().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if package.is_signed() {
            return Err(Error::Signed);
        }
        let mut candidate = package.clone();
        candidate
            .get_part_mut(self.before.worksheet_part_name())?
            .set_blob_shared(self.after.source_arc());
        let result = Snapshot::load(&candidate, self.after.sheet_position())?;
        if !result.same_source(&self.after) {
            return Err(invalid("value-only patch readback differs from its target"));
        }
        *package = candidate;
        Ok(())
    }
}

/// Successful value-only transaction publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    pub(super) fn new(snapshot: Snapshot, patch: Patch, changed_cells: usize) -> Self {
        Self {
            snapshot,
            patch,
            diagnostics: Diagnostics { changed_cells },
        }
    }

    /// Whether at least one scalar payload changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.diagnostics.changed_cells != 0
    }

    /// Resulting immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free transaction diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume into snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Content-free value-only publication diagnostics.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_cells: usize,
}

impl Diagnostics {
    /// Number of scalar payloads changed by the commit.
    #[must_use]
    pub const fn changed_cells(self) -> usize {
        self.changed_cells
    }

    /// Number of worksheet parts replaced (zero or one).
    #[must_use]
    pub const fn touched_worksheets(self) -> usize {
        if self.changed_cells == 0 { 0 } else { 1 }
    }
}
