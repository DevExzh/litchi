//! Coalesced ODS cell writes for batches and closure edits (record 0277).
//!
//! Consecutive direct cell assignments (`value`, `text`, `clear`,
//! `formula`) to distinct coordinates share one leaf `Edit`: the run plans
//! each assignment against the snapshot it started from, keeps the
//! replacement cells per sheet, and stages them with one `set_cells` per
//! sheet when it is flushed. A repeated coordinate, a style or merge
//! assignment, a range, a query or any other verb closes the run, so every
//! operation still observes the result of the ones before it.

use std::collections::BTreeMap;

use litchi_ods::Cell;

use super::{CellAssignments, Commit, Target, finish, stage_cells};
use crate::error::{Error, Result};
use crate::format::ods::{Nav, Ods};
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// One planned cell replacement.
#[derive(Debug, Clone)]
pub(crate) struct Plan {
    sheet: u32,
    at: (u32, u32),
    /// The replacement, or `None` for an exact no-op.
    cell: Option<Cell>,
    bytes: usize,
}

impl Ods {
    /// Plan a direct single-cell content assignment for a coalesced run;
    /// `None` when the call takes the ordinary path.
    pub(crate) fn cell_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Plan>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if props.is_empty()
            || super::super::metadata::is_target(target)
            || props.len() > limits.max_operations()
        {
            return Ok(None);
        }
        // A batch resolves many coordinates of one snapshot: find each row
        // by binary search over the cached row runs (record 0287).
        let node = crate::path::resolve::get(&Nav::with_indexed_rows(self), path)?;
        let Target::Cell { sheet, row, column } = Target::of(&node) else {
            return Ok(None);
        };
        let assignments = match CellAssignments::admit(props, false) {
            Ok(assignments) if assignments.is_content_only() => assignments,
            // The ordinary path reports the same admission error.
            Ok(_) | Err(_) => return Ok(None),
        };
        let existing = match &node.item() {
            super::super::node::Item::Cell(cell) => cell.cell,
            _ => None,
        };
        Ok(Some(Plan {
            sheet,
            at: (row, column),
            cell: assignments.apply(existing)?,
            bytes: assignments.charge(),
        }))
    }
}

/// A run of coalesced cell writes over one base snapshot.
pub(crate) struct Run {
    base: Ods,
    cells: BTreeMap<u32, BTreeMap<(u32, u32), Cell>>,
    /// Every coordinate the run planned, including exact no-ops.
    planned: BTreeMap<(u32, u32, u32), ()>,
    first: Option<usize>,
    operations: usize,
    staged_bytes: usize,
    limits: SemanticLimits,
}

impl Run {
    /// An empty run over `base`.
    pub(crate) fn new(base: &Ods, limits: &SemanticLimits) -> Result<Self> {
        base.admit_edit()?;
        Ok(Self {
            base: base.clone(),
            cells: BTreeMap::new(),
            planned: BTreeMap::new(),
            first: None,
            operations: 0,
            staged_bytes: 0,
            limits: *limits,
        })
    }

    /// Plan against the run's base snapshot.
    pub(crate) fn cell_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Plan>> {
        self.base.cell_plan(target, props, limits)
    }

    /// Whether `plan` must start a new run: its coordinate was planned
    /// already, or it would exceed the staging budget.
    pub(crate) fn requires_barrier(&self, plan: &Plan, limits: &SemanticLimits) -> bool {
        self.planned
            .contains_key(&(plan.sheet, plan.at.0, plan.at.1))
            || self.staged_bytes.saturating_add(plan.bytes) > limits.max_staged_bytes()
    }

    /// The wire index of the run's first operation.
    pub(crate) fn first_index(&self) -> Option<usize> {
        self.first
    }

    /// Stage `plan` as operation `index`; returns whether it changes a cell.
    pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
        let staged = self.staged_bytes.saturating_add(plan.bytes);
        if staged > self.limits.max_staged_bytes() {
            return Err(Error::from(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed: staged,
                maximum: self.limits.max_staged_bytes(),
            })
            .with_format(litchi_core::FileFormat::Ods));
        }
        self.staged_bytes = staged;
        self.first.get_or_insert(index);
        self.operations += 1;
        self.planned.insert((plan.sheet, plan.at.0, plan.at.1), ());
        let changed = plan.cell.is_some();
        if let Some(cell) = plan.cell {
            self.cells
                .entry(plan.sheet)
                .or_default()
                .insert(plan.at, cell);
        }
        Ok(changed)
    }

    /// Commit the run as one leaf edit.
    pub(crate) fn commit(self) -> Result<(Ods, Commit)> {
        let mut edit = self.base.leaf_edit()?;
        if self.cells.values().any(|cells| !cells.is_empty()) {
            stage_cells(&mut edit, self.cells)?;
        }
        let commit = finish(&self.base, edit, self.operations, &self.limits)?;
        Ok((self.base, commit))
    }
}
