//! Exact reversible workbook defined-name patches.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::{Error, Result, invalid};

/// Exact source-checked workbook XML replacement.
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

    /// Whether this patch preserves the source byte-for-byte.
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

    /// Apply after checking the complete retained workbook owner.
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
            .get_part_mut(self.before.workbook_part_name())?
            .set_blob_shared(self.after.source_arc());
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after)
            || resulting.defined_names() != self.after.defined_names()
        {
            return Err(invalid(
                "defined-name patch readback did not match its target",
            ));
        }
        *package = candidate;
        Ok(())
    }
}

/// Successful defined-name transaction publication.
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
                touched_workbooks: if changed { 1 } else { 0 },
            },
        }
    }

    /// Whether the exact authored catalog changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.diagnostics.touched_workbooks != 0
    }

    /// Resulting immutable source-bound state.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
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

/// Content-free defined-name publication diagnostics.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    touched_workbooks: u8,
}

impl Diagnostics {
    /// Number of workbook Parts replaced by the commit.
    #[must_use]
    pub const fn touched_workbooks(self) -> usize {
        self.touched_workbooks as usize
    }
}
