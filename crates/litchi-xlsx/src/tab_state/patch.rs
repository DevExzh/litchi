//! Exact reversible existing-tab state patches.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::{Error, Result, invalid};

/// Exact source-checked workbook and sheet-view replacement closure.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(super) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required source state.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch preserves every selected source byte.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after checking the full retained closure.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        if !self.before.matches_current_source(package) {
            return Err(Error::PatchConflict {
                part: self.before.workbook_part_name().to_string(),
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
            .get_part_mut(self.after.workbook_part_name())?
            .set_blob_shared(self.after.workbook_source_arc()?);
        for part in self.after.touched() {
            candidate
                .get_part_mut(&part.part.uri)?
                .set_blob_shared(part.part.bytes.detached_arc()?);
        }
        let resulting = Snapshot::load_owned_target(&candidate, &self.after)?;
        if !resulting.same_source(&self.after) || !resulting.same_semantics(&self.after) {
            return Err(invalid("tab-state patch readback did not match its target"));
        }
        *package = candidate;
        Ok(())
    }
}

/// Successful source-backed tab-state publication plan.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    pub(super) const fn new(snapshot: Snapshot, patch: Patch, touched_parts: u8) -> Self {
        Self {
            snapshot,
            patch,
            diagnostics: Diagnostics { touched_parts },
        }
    }

    /// Whether visibility or active selection changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.diagnostics.touched_parts != 0
    }

    /// Resulting immutable state.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free closure diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume this result into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Publication closure cardinality without exposing physical names.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    touched_parts: u8,
}

impl Diagnostics {
    /// Planned materialized/overlay Part count: one workbook or workbook plus
    /// old and new worksheet views.
    #[must_use]
    pub const fn touched_parts(self) -> usize {
        self.touched_parts as usize
    }

    /// Number of worksheet view Parts in the replacement closure.
    #[must_use]
    pub const fn touched_worksheets(self) -> usize {
        self.touched_parts.saturating_sub(1) as usize
    }
}
