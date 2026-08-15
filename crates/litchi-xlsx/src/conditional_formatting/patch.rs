//! Exact reversible worksheet conditional-formatting patches.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::{Error, Result, invalid};

/// Exact-source-checked worksheet owner replacement.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(super) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[must_use]
    /// Exact source state required for application.
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    /// Exact state produced by successful application.
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    #[must_use]
    /// Whether application preserves the selected source byte-for-byte.
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    /// Construct a patch that restores this patch's exact source state.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply after checking the workbook, worksheet, relationships, and styles closure.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.before.check_execution()?;
        self.after.check_execution()?;
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
            .set_blob_shared(self.after.source_arc()?);
        let resulting = Snapshot::load(&candidate, self.after.sheet_position())?;
        if !resulting.same_source(&self.after)
            || resulting.collections() != self.after.collections()
        {
            return Err(invalid(
                "conditional-formatting patch readback did not match its target",
            ));
        }
        *package = candidate;
        Ok(())
    }
}

/// Successful conditional-formatting publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    pub(super) const fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                touched_worksheets: if changed { 1 } else { 0 },
            },
        }
    }

    #[must_use]
    /// Whether the authored conditional-formatting collection changed.
    pub const fn changed(&self) -> bool {
        self.diagnostics.touched_worksheets != 0
    }

    #[must_use]
    /// Resulting immutable source-bound snapshot.
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    /// Exact reversible patch produced by the edit.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    /// Content-free publication diagnostics.
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    #[must_use]
    /// Consume this result into its snapshot and reversible patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    touched_worksheets: u8,
}

impl Diagnostics {
    #[must_use]
    /// Number of worksheet Parts replaced by the publication.
    pub fn touched_worksheets(self) -> usize {
        usize::from(self.touched_worksheets)
    }
}
