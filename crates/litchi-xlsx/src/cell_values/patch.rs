//! Exact reversible value-only worksheet patches.

use litchi_opc::OpcPackage;

use super::{MultiSnapshot, Snapshot};
use crate::error::{Error, Result, invalid};

/// Exact source-bound replacement of one safe worksheet value closure.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

/// Exact source-bound replacement of a bounded worksheet set.
///
/// The patch owns exact bytes for selected worksheet Parts and the complete
/// workbook ownership graph. Unselected worksheet payload bytes are not part
/// of the closure: applying a patch preserves valid independent edits to
/// those payloads while still rejecting graph, relationship, content-type,
/// signature, and source-lineage conflicts.
#[derive(Clone, Debug)]
pub struct MultiPatch {
    before: MultiSnapshot,
    after: MultiSnapshot,
}

impl MultiPatch {
    pub(crate) const fn new(before: MultiSnapshot, after: MultiSnapshot) -> Self {
        Self { before, after }
    }

    /// Required exact source state for every selected worksheet.
    #[must_use]
    pub const fn before(&self) -> &MultiSnapshot {
        &self.before
    }

    /// Exact target state for every selected worksheet.
    #[must_use]
    pub const fn after(&self) -> &MultiSnapshot {
        &self.after
    }

    /// Whether every selected worksheet remains byte-identical to its source.
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

    /// Apply all worksheet replacements atomically after exact source checks.
    ///
    /// A patch whose target still carries a managed source payload returns a
    /// typed `ManagedPartDataArcEscape` error. Use
    /// [`Self::apply_materialized`] when an owning [`OpcPackage`] handoff is
    /// explicitly desired.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply_inner(package, None)
    }

    /// Apply after explicitly copying managed target payloads into the owning
    /// package, bounded by `maximum_bytes` across this patch's replacements.
    pub fn apply_materialized(&self, package: &mut OpcPackage, maximum_bytes: usize) -> Result<()> {
        self.apply_inner(package, Some(maximum_bytes))
    }

    fn apply_inner(&self, package: &mut OpcPackage, maximum_bytes: Option<usize>) -> Result<()> {
        self.before.check_execution()?;
        self.after.check_execution()?;
        let current = MultiSnapshot::load_owned(
            package,
            self.before.sheets().iter().map(Snapshot::sheet_position),
        )
        .map_err(|_| Error::PatchConflict {
            part: self
                .before
                .sheets()
                .first()
                .map_or_else(String::new, |snapshot| {
                    snapshot.worksheet_part_name().to_string()
                }),
        })?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self
                    .before
                    .sheets()
                    .first()
                    .map_or_else(String::new, |snapshot| {
                        snapshot.worksheet_part_name().to_string()
                    }),
            });
        }
        if self.is_empty() {
            self.before.check_execution()?;
            self.after.check_execution()?;
            return Ok(());
        }
        if package.is_signed() {
            return Err(Error::Signed);
        }
        self.after.check_execution()?;
        let mut candidate = package.clone();
        MultiSnapshot::apply_owned_target(
            &self.before,
            &self.after,
            &mut candidate,
            maximum_bytes,
        )?;
        let readback = MultiSnapshot::load_owned(
            &candidate,
            self.after.sheets().iter().map(Snapshot::sheet_position),
        )?;
        if !readback.same_source(&self.after) {
            return Err(invalid(
                "multi-sheet value-only patch readback differs from its target",
            ));
        }
        self.after.check_execution()?;
        *package = candidate;
        Ok(())
    }
}

impl Patch {
    pub(crate) const fn new(before: Snapshot, after: Snapshot) -> Self {
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
    ///
    /// A managed source target is refused unless the caller uses
    /// [`Self::apply_materialized`] explicitly.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply_inner(package, None)
    }

    /// Apply after explicitly copying a managed target payload into the owning
    /// package, bounded by `maximum_bytes`.
    pub fn apply_materialized(&self, package: &mut OpcPackage, maximum_bytes: usize) -> Result<()> {
        self.apply_inner(package, Some(maximum_bytes))
    }

    fn apply_inner(&self, package: &mut OpcPackage, maximum_bytes: Option<usize>) -> Result<()> {
        self.before.check_execution()?;
        self.after.check_execution()?;
        if !self.before.matches_current_source(package) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet_part_name().to_string(),
            });
        }
        if self.is_empty() {
            self.before.check_execution()?;
            self.after.check_execution()?;
            return Ok(());
        }
        if package.is_signed() {
            return Err(Error::Signed);
        }
        self.after.check_execution()?;
        let mut candidate = package.clone();
        Snapshot::apply_owned_target(&self.before, &self.after, &mut candidate, maximum_bytes)?;
        let result = Snapshot::load(&candidate, self.after.sheet_position())?;
        if !result.same_source(&self.after) {
            return Err(invalid("value-only patch readback differs from its target"));
        }
        self.after.check_execution()?;
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

/// Successful bounded multi-worksheet value-only publication.
#[derive(Debug)]
pub struct MultiCommit {
    snapshot: MultiSnapshot,
    patch: MultiPatch,
    diagnostics: MultiDiagnostics,
}

impl MultiCommit {
    pub(super) fn new(
        snapshot: MultiSnapshot,
        patch: MultiPatch,
        changed_cells: usize,
        touched_worksheets: usize,
    ) -> Self {
        Self {
            snapshot,
            patch,
            diagnostics: MultiDiagnostics {
                changed_cells,
                touched_worksheets,
            },
        }
    }

    /// Whether at least one scalar payload changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.diagnostics.changed_cells != 0
    }

    /// Resulting immutable worksheet snapshots.
    #[must_use]
    pub const fn snapshot(&self) -> &MultiSnapshot {
        &self.snapshot
    }

    /// Exact reversible multi-worksheet patch.
    #[must_use]
    pub const fn patch(&self) -> &MultiPatch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> MultiDiagnostics {
        self.diagnostics
    }

    /// Consume into snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (MultiSnapshot, MultiPatch) {
        (self.snapshot, self.patch)
    }
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

/// Content-free bounded multi-worksheet publication diagnostics.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct MultiDiagnostics {
    changed_cells: usize,
    touched_worksheets: usize,
}

impl MultiDiagnostics {
    /// Number of scalar payloads changed by the commit.
    #[must_use]
    pub const fn changed_cells(self) -> usize {
        self.changed_cells
    }

    /// Number of worksheet Parts replaced by the commit.
    #[must_use]
    pub const fn touched_worksheets(self) -> usize {
        self.touched_worksheets
    }
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
