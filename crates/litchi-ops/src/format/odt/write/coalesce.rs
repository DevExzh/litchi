//! Coalesced ODT property sets: consecutive direct-path sets of one
//! transaction share one leaf body edit, so a batch of N sets publishes one
//! package instead of N.
//!
//! Every set is planned against the run's base handle. A set keeps the
//! positional paths of every other node valid unless it replaces content
//! (`text`); a later set inside content an earlier set replaced closes the
//! run first, so each plan resolves against the document its sequential
//! predecessors produced. Sets whose effect would move later positional
//! paths (clearing a character-data run, replacing a run that holds notes
//! or bookmarks) are not planned here and run on their own.

use std::collections::BTreeMap;

use litchi_odt::body::NodeId;
use litchi_odt::transaction::Edit;

use super::{Commit, Planned, bound, finish};
use crate::Result;
use crate::format::odt::Odt;
use crate::transaction::{LimitKind, SemanticLimits};

/// The most leaf body edits one run stages (the leaf's own ceiling).
const MAX_RUN_OPERATIONS: usize = 65_536;

/// A run of planned sets sharing one leaf edit.
pub(crate) struct Run {
    base: Odt,
    edit: Edit,
    limits: SemanticLimits,
    first_index: Option<usize>,
    operations: usize,
    staged_bytes: usize,
    /// Source ranges of nodes the run edits, by start offset: `(end, node)`.
    touched: BTreeMap<usize, (usize, NodeId)>,
    /// Source ranges whose content the run replaces, by start offset.
    replaced: BTreeMap<usize, (usize, NodeId)>,
}

impl Run {
    /// A run over the immutable handle `base`.
    pub(crate) fn new(base: &Odt, limits: &SemanticLimits) -> Result<Self> {
        Ok(Self {
            base: base.clone(),
            edit: base.package()?.edit(),
            limits: *limits,
            first_index: None,
            operations: 0,
            staged_bytes: 0,
            touched: BTreeMap::new(),
            replaced: BTreeMap::new(),
        })
    }

    /// The transaction item index of the run's first set.
    pub(crate) const fn first_index(&self) -> Option<usize> {
        self.first_index
    }

    /// Whether `plan` must start a new run: it lies inside content an
    /// earlier set replaces, replaces content an earlier set edits, or
    /// would exceed the run's ceilings.
    pub(crate) fn requires_barrier(&self, plan: &Planned) -> bool {
        if self.operations.saturating_add(plan.operations()) > MAX_RUN_OPERATIONS {
            return true;
        }
        let Planned { id, range, staged } = plan;
        let inside_replaced = self
            .replaced
            .range(..=range.start)
            .next_back()
            .is_some_and(|(_, (end, node))| *node != *id && range.end <= *end);
        let replaces_touched = staged.replaces_content()
            && self
                .touched
                .range(range.start..range.end)
                .any(|(_, (_, node))| *node != *id);
        inside_replaced || replaces_touched
    }

    /// Stage `plan`, recording `index` as the run's first item.
    ///
    /// # Errors
    ///
    /// The leaf's staging refusal, or a transaction limit.
    pub(crate) fn stage(&mut self, plan: Planned, index: usize) -> Result<bool> {
        let operations = self.operations.saturating_add(plan.operations());
        bound(
            LimitKind::Operations,
            operations,
            self.limits.max_operations().max(MAX_RUN_OPERATIONS),
        )?;
        let staged_bytes = self.staged_bytes.saturating_add(plan.staged_bytes());
        bound(
            LimitKind::StagedBytes,
            staged_bytes,
            self.limits.max_staged_bytes(),
        )?;
        plan.stage(&mut self.edit)?;
        self.touched
            .insert(plan.range.start, (plan.range.end, plan.id));
        if plan.staged.replaces_content() {
            self.replaced
                .insert(plan.range.start, (plan.range.end, plan.id));
        }
        self.operations = operations;
        self.staged_bytes = staged_bytes;
        self.first_index.get_or_insert(index);
        Ok(true)
    }

    /// Commit the run's edit as one leaf transaction.
    ///
    /// # Errors
    ///
    /// The leaf's commit failure or the snapshot limit.
    pub(crate) fn commit(self) -> Result<(Odt, Commit)> {
        let commit = finish(&self.base, self.edit, self.operations, &self.limits)?;
        Ok((self.base, commit))
    }
}
