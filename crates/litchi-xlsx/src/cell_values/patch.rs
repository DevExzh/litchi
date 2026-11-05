//! Exact reversible value-only worksheet records: the replay payloads of the
//! source-backed [`super::Patch`] and [`super::sheet::Patch`] families.

use super::{MultiSnapshot, Snapshot};
use crate::error::Result;

/// Exact source-bound replacement of one safe worksheet value closure.
#[derive(Clone, Debug)]
pub(crate) struct SheetRecord {
    before: Snapshot,
    after: Snapshot,
    changed_cells: usize,
}

/// Exact source-bound replacement of a bounded worksheet set.
///
/// The patch owns exact bytes for selected worksheet Parts and the complete
/// workbook ownership graph. Unselected worksheet payload bytes are not part
/// of the closure: applying a patch preserves valid independent edits to
/// those payloads while still rejecting graph, relationship, content-type,
/// signature, and source-lineage conflicts.
#[derive(Clone, Debug)]
pub(crate) struct Record {
    before: MultiSnapshot,
    after: MultiSnapshot,
    changed_cells: usize,
    touched_worksheets: usize,
}

impl Record {
    pub(crate) const fn new(
        before: MultiSnapshot,
        after: MultiSnapshot,
        changed_cells: usize,
        touched_worksheets: usize,
    ) -> Self {
        Self {
            before,
            after,
            changed_cells,
            touched_worksheets,
        }
    }

    /// The scalar payloads and worksheet parts the record changes.
    pub(crate) const fn counts(&self) -> (usize, usize) {
        (self.changed_cells, self.touched_worksheets)
    }

    /// Required exact source state for every selected worksheet.
    #[must_use]
    pub(crate) const fn before(&self) -> &MultiSnapshot {
        &self.before
    }

    /// Exact target state for every selected worksheet.
    #[must_use]
    pub(crate) const fn after(&self) -> &MultiSnapshot {
        &self.after
    }

    /// Whether every selected worksheet remains byte-identical to its source.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Exact source-bound inverse.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changed_cells: self.changed_cells,
            touched_worksheets: self.touched_worksheets,
        }
    }
}

impl SheetRecord {
    pub(crate) const fn new(before: Snapshot, after: Snapshot, changed_cells: usize) -> Self {
        Self {
            before,
            after,
            changed_cells,
        }
    }

    /// The scalar payloads the record changes.
    pub(crate) const fn changed_cells(&self) -> usize {
        self.changed_cells
    }

    /// The record with its target worksheet payload copied out of a managed
    /// reservation under `maximum_bytes`.
    pub(crate) fn materialized(&self, maximum_bytes: usize) -> Result<Self> {
        Ok(Self {
            before: self.before.clone(),
            after: self.after.materialized(maximum_bytes)?,
            changed_cells: self.changed_cells,
        })
    }

    /// Required exact source state.
    #[must_use]
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact target state.
    #[must_use]
    pub(crate) const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether application preserves the selected source bytes exactly.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Exact source-bound inverse.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changed_cells: self.changed_cells,
        }
    }
}

/// What a committed multi-worksheet value change changed.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_cells: usize,
    touched_worksheets: usize,
}

impl Diagnostics {
    pub(crate) const fn new(changed_cells: usize, touched_worksheets: usize) -> Self {
        Self {
            changed_cells,
            touched_worksheets,
        }
    }

    /// Whether at least one scalar payload changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed_cells != 0
    }

    /// Number of scalar payloads changed by the commit.
    #[must_use]
    pub const fn changed_cells(&self) -> usize {
        self.changed_cells
    }

    /// Number of worksheet parts replaced by the commit.
    #[must_use]
    pub const fn touched_worksheets(&self) -> usize {
        self.touched_worksheets
    }
}
