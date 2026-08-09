//! Non-applying bounded three-way planning for workbook transactions.

use std::fmt;
use std::sync::Arc;

use litchi_core::patch::{JoinedSubEdits, SubEdit, ThreeWayMergePlan as CorePlan};

pub use litchi_core::patch::{CompositionLimits as MergeLimits, MergeChoice};

use super::semantic::Edit;
use crate::error::{Result, invalid};
use crate::workbook::{Inner, Workbook};

#[derive(Clone)]
struct Lineage(Arc<Inner>);

impl PartialEq for Lineage {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl Eq for Lineage {}

impl fmt::Debug for Lineage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("WorkbookLineage(..)")
    }
}

/// Non-applying merge plan for two workbook edits from one exact snapshot.
///
/// The common planner retains disjoint work automatically and requires an
/// explicit [`MergeChoice`] for one conservative conflicting group.
pub struct ThreeWayPlan {
    base: Workbook,
    inner: CorePlan<Lineage, Edit>,
}

impl ThreeWayPlan {
    pub(super) fn new(left: Edit, right: Edit, limits: MergeLimits) -> Result<Self> {
        if !Arc::ptr_eq(&left.base.inner, &right.base.inner) {
            return Err(invalid(
                "three-way workbook branches have different source lineages",
            ));
        }
        let base = left.base.clone();
        let lineage = Lineage(Arc::clone(&base.inner));
        let (left_reads, left_writes) = effects(&left);
        let (right_reads, right_writes) = effects(&right);
        let left = SubEdit::new(
            lineage.clone(),
            limits,
            "left",
            left_reads,
            left_writes,
            left,
        )
        .map_err(|error| invalid(format!("invalid left workbook branch: {error}")))?;
        let right = SubEdit::new(
            lineage.clone(),
            limits,
            "right",
            right_reads,
            right_writes,
            right,
        )
        .map_err(|error| invalid(format!("invalid right workbook branch: {error}")))?;
        let mut left_branch = JoinedSubEdits::new(lineage.clone(), limits);
        left_branch
            .join(left)
            .map_err(|error| invalid(format!("left workbook branch was refused: {error:?}")))?;
        let mut right_branch = JoinedSubEdits::new(lineage, limits);
        right_branch
            .join(right)
            .map_err(|error| invalid(format!("right workbook branch was refused: {error:?}")))?;
        let inner = CorePlan::new(left_branch, right_branch)
            .map_err(|error| invalid(format!("workbook merge planning failed: {error:?}")))?;
        Ok(Self { base, inner })
    }

    /// Automatically retained disjoint branch count.
    #[must_use]
    pub fn automatic_len(&self) -> usize {
        self.inner.automatic().len()
    }

    /// Deterministically ordered common-planner conflict details.
    #[must_use]
    pub const fn conflicts(
        &self,
    ) -> &litchi_core::patch::ConflictSet<litchi_core::patch::SubEditConflict> {
        self.inner.conflicts()
    }

    /// Current explicit conflict-group resolution.
    #[must_use]
    pub const fn resolution(&self) -> Option<MergeChoice> {
        self.inner.resolution()
    }

    /// Resolve the complete conservative conflict group.
    pub fn resolve(&mut self, choice: MergeChoice) -> &mut Self {
        self.inner.resolve(choice);
        self
    }

    /// Finish the plan into one still-uncommitted workbook transaction.
    ///
    /// No workbook bytes are changed until the returned [`Edit`] is committed.
    ///
    /// # Errors
    ///
    /// Returns an error while conflicts remain unresolved or if the accepted
    /// sub-edits unexpectedly fail the XLSX owner's stricter join validation.
    pub fn finish(self) -> Result<Edit> {
        let joined = self
            .inner
            .finish()
            .map_err(|_| invalid("three-way workbook conflicts remain unresolved"))?;
        let mut edits = joined.into_sub_edits().map(SubEdit::into_payload);
        let Some(mut merged) = edits.next() else {
            return Edit::new(self.base);
        };
        for edit in edits {
            merged.join(edit)?;
        }
        Ok(merged)
    }
}

impl Edit {
    /// Plan a bounded three-way merge without committing either branch.
    ///
    /// Both edits must originate from the same immutable workbook. Exact
    /// semantic effect keys are passed through the common planner; disjoint
    /// effects are retained automatically and overlaps require resolution.
    ///
    /// # Errors
    ///
    /// Returns an error for different lineages, invalid/bound-exceeding effect
    /// sets, or a common-planner refusal.
    pub fn plan_three_way(self, other: Self, limits: MergeLimits) -> Result<ThreeWayPlan> {
        ThreeWayPlan::new(self, other, limits)
    }
}

fn effects(edit: &Edit) -> (Vec<String>, Vec<String>) {
    let mut reads = Vec::new();
    let mut writes = Vec::new();
    if edit.panes.is_some() {
        writes.push("workbook/task-panes".to_owned());
    }
    if edit.active.is_some() {
        writes.push("workbook/active-tab".to_owned());
    }
    if edit
        .order
        .as_ref()
        .is_some_and(super::validation::OrderPlan::is_effective)
    {
        writes.push("workbook/tab-order".to_owned());
    }
    for position in &edit.removed {
        writes.push(format!("sheet/{position}/owner"));
    }
    for (position, actions) in &edit.sheets {
        reads.push(format!("sheet/{position}/owner"));
        if let Some(name) = &actions.rename {
            writes.push(format!("sheet/{position}/name"));
            writes.push(format!("workbook/name/{}", name.identity_key()));
        }
        if actions.visibility.is_some() {
            writes.push(format!("sheet/{position}/visibility"));
        }
        if actions.defaults.is_some() {
            writes.push(format!("sheet/{position}/defaults"));
        }
        if actions.web.is_some() {
            writes.push(format!("sheet/{position}/web-bindings"));
        }
        if !actions.merges.is_empty() {
            writes.push(format!("sheet/{position}/merges"));
        }
        if actions.page_breaks.is_some() {
            writes.push(format!("sheet/{position}/page-breaks"));
        }
        writes.extend(
            actions
                .cells
                .keys()
                .map(|address| format!("sheet/{position}/cell/{address}")),
        );
        writes.extend(
            actions
                .rows
                .keys()
                .map(|row| format!("sheet/{position}/row/{}", row.get())),
        );
        writes.extend(
            actions
                .columns
                .keys()
                .map(|column| format!("sheet/{position}/column/{}", column.get())),
        );
    }
    for (index, added) in edit.added.iter().enumerate() {
        writes.push(format!("workbook/name/{}", added.name.identity_key()));
        writes.push(format!("added/{index}/owner"));
    }
    (reads, writes)
}
