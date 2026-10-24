//! Guarded source-backed value-only worksheet edits of a
//! [`SourceBackedWorkbook`] (record 0360, the 0355a §1 shape).
//!
//! [`SourceBackedWorkbook::edit_cell_values`] stages value-only changes over a
//! bounded worksheet set ([`super::Edit`]); [`super::sheet`] is the same
//! family over one worksheet. A commit returns the new workbook whose output
//! is the topology publication the former stream publishers wrote: the
//! selected worksheet parts (and the invalidated calculation closure) are
//! replaced, every other member is raw-copied.

use std::collections::BTreeMap;
use std::io::Write;
use std::sync::Arc;

use litchi_opc::SourceBackedPackage;
use litchi_sheet::Cell as Address;

use super::patch::Record;
use super::snapshot::SourceProvenance;
use super::{Commit, MAX_SHEET_OWNERS, MultiSnapshot, Patch, Snapshot};
use crate::cell::{Cell, Content, Number, Value};
use crate::error::{Error, Result, invalid};
use crate::formula::Formula;
use crate::raw::worksheet::edit::{Action, rewrite_value_only_with_provenance};
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

pub mod sheet;

/// Maximum unique cells in one atomic value transaction.
pub const MAX_BATCH_EDITS: usize = 256;

/// One typed value-only edit for a scalar cell.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    /// Insert one numeric cell at an absent coordinate.
    Insert { address: Address, value: Number },
    /// Replace the stored scalar value without changing local style.
    Set { address: Address, value: Value },
    /// Replace the payload with a checked cacheless scalar formula.
    SetFormula { address: Address, formula: Formula },
    /// Replace one validated shared-formula master and all of its members.
    SetSharedFormula { address: Address, formula: Formula },
    /// Remove the stored scalar payload while retaining the cell record.
    Clear { address: Address },
    /// Remove the complete stored scalar cell record.
    ///
    /// Unlike [`Self::Clear`], this also removes cell-local type and style
    /// attributes. The containing row and the producer's declared worksheet
    /// dimension are retained conservatively.
    Remove { address: Address },
}

/// One selector-first scalar edit for a bounded multi-worksheet transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetChange<'a> {
    /// Worksheet selected by its semantic name or zero-based position.
    pub selector: Selector<'a>,
    /// Scalar operation applied to one selected worksheet.
    pub edit: Change,
}

impl<'a> SheetChange<'a> {
    /// Construct a selector-first numeric insertion for an absent coordinate.
    pub fn insert(selector: impl Into<Selector<'a>>, address: Address, value: Number) -> Self {
        Self {
            selector: selector.into(),
            edit: Change::insert(address, value),
        }
    }

    /// Construct a selector-first value replacement.
    pub fn set(
        selector: impl Into<Selector<'a>>,
        address: Address,
        value: impl Into<Value>,
    ) -> Self {
        Self {
            selector: selector.into(),
            edit: Change::set(address, value),
        }
    }

    /// Construct a selector-first scalar formula replacement.
    pub fn set_formula(
        selector: impl Into<Selector<'a>>,
        address: Address,
        formula: Formula,
    ) -> Self {
        Self {
            selector: selector.into(),
            edit: Change::set_formula(address, formula),
        }
    }

    /// Construct a selector-first shared-formula master replacement.
    pub fn set_shared_formula(
        selector: impl Into<Selector<'a>>,
        address: Address,
        formula: Formula,
    ) -> Self {
        Self {
            selector: selector.into(),
            edit: Change::set_shared_formula(address, formula),
        }
    }

    /// Construct a selector-first scalar clear.
    #[must_use]
    pub fn clear(selector: impl Into<Selector<'a>>, address: Address) -> Self {
        Self {
            selector: selector.into(),
            edit: Change::clear(address),
        }
    }

    /// Construct a selector-first complete cell-owner removal.
    #[must_use]
    pub fn remove(selector: impl Into<Selector<'a>>, address: Address) -> Self {
        Self {
            selector: selector.into(),
            edit: Change::remove(address),
        }
    }
}

impl Change {
    /// Construct a numeric insertion for an absent coordinate.
    pub fn insert(address: Address, value: Number) -> Self {
        Self::Insert { address, value }
    }

    /// Construct a checked value replacement.
    pub fn set(address: Address, value: impl Into<Value>) -> Self {
        Self::Set {
            address,
            value: value.into(),
        }
    }

    /// Construct a checked cacheless scalar formula replacement.
    #[must_use]
    pub const fn set_formula(address: Address, formula: Formula) -> Self {
        Self::SetFormula { address, formula }
    }

    /// Construct a checked shared-formula master replacement.
    #[must_use]
    pub const fn set_shared_formula(address: Address, formula: Formula) -> Self {
        Self::SetSharedFormula { address, formula }
    }

    /// Construct a scalar-value clear.
    #[must_use]
    pub const fn clear(address: Address) -> Self {
        Self::Clear { address }
    }

    /// Construct a complete scalar-cell owner removal.
    #[must_use]
    pub const fn remove(address: Address) -> Self {
        Self::Remove { address }
    }

    /// Target coordinate.
    #[must_use]
    pub const fn address(&self) -> Address {
        match self {
            Self::Insert { address, .. }
            | Self::Set { address, .. }
            | Self::SetFormula { address, .. }
            | Self::SetSharedFormula { address, .. }
            | Self::Clear { address }
            | Self::Remove { address } => *address,
        }
    }
}

#[derive(Clone, Debug)]
enum StagedValueEdit {
    Insert(Content),
    Set(Content),
    SetSharedFormula(Formula),
    Clear,
    Remove,
}

/// Clone-staged atomic value changes over a bounded worksheet set of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<MultiSnapshot>,
    selectors: Vec<Selector<'static>>,
    pending: Option<Vec<SheetChange<'static>>>,
    staged: BTreeMap<usize, Vec<(Address, StagedValueEdit)>>,
    staged_cells: usize,
}

impl SourceBackedWorkbook {
    /// Begin a bounded value-only edit over selected worksheet owners
    /// without loading unselected worksheet payloads.
    ///
    /// # Errors
    ///
    /// Returns an error when the set exceeds its bounds or source execution is
    /// refused. Selector and payload errors are deferred until first use.
    pub fn edit_cell_values<'a, I>(&self, selectors: I) -> Result<Edit>
    where
        I: IntoIterator<Item = Selector<'a>>,
    {
        self.check_execution()?;
        let mut selected = Vec::new();
        selected
            .try_reserve_exact(MAX_SHEET_OWNERS)
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet selector staging",
                source,
            })?;
        for selector in selectors {
            if selected.len() >= MAX_SHEET_OWNERS {
                return Err(invalid(format!(
                    "multi-sheet value edits exceed {MAX_SHEET_OWNERS} worksheet owners"
                )));
            }
            selected.push(own_selector(selector)?);
        }
        if selected.is_empty() {
            return Err(invalid("multi-sheet value edits require one worksheet"));
        }
        Ok(Edit::new(self.clone(), selected))
    }

    /// Begin a bounded selector-first transaction and stage its initial batch.
    pub fn edit_cell_values_batch<'a, I>(&self, edits: I) -> Result<Edit>
    where
        I: IntoIterator<Item = SheetChange<'a>>,
    {
        self.check_execution()?;
        let mut requests = edits.into_iter();
        let first = requests.next();
        self.check_execution()?;
        let first =
            first.ok_or_else(|| invalid("multi-sheet value edits require one cell edit"))?;
        let first = own_sheet_change(first)?;
        let mut selectors = Vec::new();
        selectors
            .try_reserve_exact(MAX_SHEET_OWNERS)
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet selector staging",
                source,
            })?;
        selectors.push(first.selector.clone());
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS)
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet value staging",
                source,
            })?;
        pending.push(first);
        for request in requests {
            self.check_execution()?;
            let request = own_sheet_change(request)?;
            if pending.len() >= MAX_BATCH_EDITS {
                return Err(Error::limit(
                    super::LimitKind::BatchEdits,
                    MAX_BATCH_EDITS + 1,
                    MAX_BATCH_EDITS,
                ));
            }
            if !selectors
                .iter()
                .any(|selector| selector == &request.selector)
            {
                if selectors.len() >= MAX_SHEET_OWNERS {
                    return Err(Error::limit(
                        super::LimitKind::SheetOwners,
                        MAX_SHEET_OWNERS + 1,
                        MAX_SHEET_OWNERS,
                    ));
                }
                selectors.push(request.selector.clone());
            }
            pending.push(request);
        }
        let mut transaction = Edit::new(self.clone(), selectors);
        transaction.pending = Some(pending);
        Ok(transaction)
    }

    /// Replay a multi-worksheet value patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when a selected worksheet does not hold
    /// the patch's before state, and the signed-source refusal for a changed
    /// patch on a signed source.
    pub fn apply_cell_values(&self, patch: &Patch) -> Result<Commit> {
        super::apply_change(self, patch)
    }
}

impl crate::workbook::source::deferred::Change for Record {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        Record::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        let before = self.before();
        let conflict = || Error::PatchConflict {
            part: before
                .sheets()
                .first()
                .map_or_else(String::new, |snapshot| {
                    snapshot.worksheet_part_name().to_string()
                }),
        };
        match before.matches_source_backed(package)? {
            // The patch was made over this exact source: nothing to reload.
            SourceProvenance::Matched => Ok(()),
            SourceProvenance::Mismatched if before.same_lineage(package) => Err(conflict()),
            // A replay compares content (record 0360).
            SourceProvenance::Mismatched | SourceProvenance::Unavailable => {
                let current = MultiSnapshot::load_source_backed(
                    package,
                    before
                        .sheets()
                        .iter()
                        .map(|snapshot| snapshot.sheet_position().into()),
                )?;
                if current.same_content(before) {
                    Ok(())
                } else {
                    Err(conflict())
                }
            },
        }
    }

    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        package.check_execution()?;
        crate::workbook::source::deferred::Change::check(self, package)?;
        let plan = self.after().topology_plan_from(self.before())?;
        package.clone().write_topology_to_stream(writer, plan)?;
        Ok(())
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        package.check_execution()?;
        let plan = self.after().topology_plan_from(self.before())?;
        package.clone().write_topology_to_stream(writer, plan)?;
        Ok(())
    }
}

impl crate::workbook::source::deferred::Report for Record {
    type Diagnostics = super::Diagnostics;

    fn diagnostics(&self) -> super::Diagnostics {
        let (changed_cells, touched_worksheets) = self.counts();
        super::Diagnostics::new(changed_cells, touched_worksheets)
    }
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, selectors: Vec<Selector<'static>>) -> Self {
        Self {
            source: LazySource::new(workbook),
            selectors,
            pending: None,
            staged: BTreeMap::new(),
            staged_cells: 0,
        }
    }

    fn resolve_before(&self) -> Result<&MultiSnapshot> {
        let selectors = self.selectors.clone();
        self.source
            .resolve(|package| MultiSnapshot::load_source_backed(package, selectors))
    }

    /// Resolve the exact source snapshots observed by this transaction.
    pub fn before(&self) -> Result<&MultiSnapshot> {
        self.resolve_before()
    }

    /// Number of unique staged cell selectors across all worksheets.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.staged_cells
            + match &self.pending {
                Some(pending) => pending.len(),
                None => 0,
            }
    }

    /// Whether no cell selectors have been staged.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.staged_cells == 0
            && match &self.pending {
                Some(pending) => pending.is_empty(),
                None => true,
            }
    }

    /// Number of selected worksheet owners.
    pub fn worksheet_count(&self) -> Result<usize> {
        Ok(self.resolve_before()?.len())
    }

    /// Stage one selector-first numeric insertion for an absent coordinate.
    pub fn insert<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        value: Number,
    ) -> Result<()> {
        self.apply_batch([SheetChange::insert(selector, address, value)])
    }

    /// Stage one selector-first scalar replacement.
    pub fn set<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        value: impl Into<Value>,
    ) -> Result<()> {
        self.apply_batch([SheetChange::set(selector, address, value)])
    }

    /// Stage one selector-first cacheless scalar formula replacement.
    pub fn set_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        formula: Formula,
    ) -> Result<()> {
        self.apply_batch([SheetChange::set_formula(selector, address, formula)])
    }

    /// Stage one selector-first shared-formula master replacement.
    pub fn set_shared_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        formula: Formula,
    ) -> Result<()> {
        self.apply_batch([SheetChange::set_shared_formula(selector, address, formula)])
    }

    /// Stage one selector-first scalar clear.
    pub fn clear<'a>(&mut self, selector: impl Into<Selector<'a>>, address: Address) -> Result<()> {
        self.apply_batch([SheetChange::clear(selector, address)])
    }

    /// Stage one selector-first complete cell-owner removal.
    pub fn remove<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
    ) -> Result<()> {
        self.apply_batch([SheetChange::remove(selector, address)])
    }

    /// Stage a bounded cross-worksheet batch atomically.
    pub fn apply_batch<'a, I>(&mut self, edits: I) -> Result<()>
    where
        I: IntoIterator<Item = SheetChange<'a>>,
    {
        self.flush_pending()?;
        self.apply_batch_inner(edits)
    }

    fn flush_pending(&mut self) -> Result<()> {
        let Some(pending) = self.pending.clone() else {
            return Ok(());
        };
        self.apply_batch_inner(pending)?;
        self.pending = None;
        Ok(())
    }

    fn apply_batch_inner<'a, I>(&mut self, edits: I) -> Result<()>
    where
        I: IntoIterator<Item = SheetChange<'a>>,
    {
        let before = self.resolve_before()?.clone();
        before.check_execution()?;
        let remaining = MAX_BATCH_EDITS.saturating_sub(self.staged_cells);
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(remaining)
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet pending value batch",
                source,
            })?;
        for request in edits {
            before.check_execution()?;
            if pending.len() >= remaining {
                return Err(Error::limit(
                    super::LimitKind::BatchEdits,
                    MAX_BATCH_EDITS + 1,
                    MAX_BATCH_EDITS,
                ));
            }
            let position = resolve_snapshot_selector(&before, &request.selector)?;
            let snapshot = before
                .sheets()
                .iter()
                .find(|snapshot| snapshot.sheet_position() == position)
                .ok_or_else(|| invalid("multi-sheet edit selector was not selected"))?;
            let address = request.edit.address();
            let duplicate = self
                .staged
                .get(&position)
                .is_some_and(|entries| entries.iter().any(|(stored, _)| *stored == address))
                || pending.iter().any(|(stored_position, stored, _)| {
                    *stored_position == position && *stored == address
                });
            if duplicate {
                return Err(invalid(format!(
                    "duplicate multi-sheet value-only selector '{position}:{}'",
                    address
                )));
            }
            let source = snapshot.editable_cell(address);
            let value = match request.edit {
                Change::Insert { value, .. } => {
                    snapshot.require_insertable_absence(address)?;
                    let value = Value::Number(value);
                    value.validate_for_write()?;
                    StagedValueEdit::Insert(Content::Value(value))
                },
                Change::Set { value, .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    snapshot.require_shared_string_free_target(address)?;
                    value.validate_for_write()?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(snapshot.edit_blocked(address));
                    }
                    snapshot.require_unmerged_target(address)?;
                    StagedValueEdit::Set(Content::Value(value))
                },
                Change::SetFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    snapshot.require_shared_string_free_target(address)?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    snapshot.require_formula_target(address)?;
                    snapshot.require_unmerged_target(address)?;
                    StagedValueEdit::Set(Content::Formula(formula))
                },
                Change::SetSharedFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    snapshot.require_shared_string_free_target(address)?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    snapshot.shared_formula_group(address, MAX_BATCH_EDITS)?;
                    snapshot.require_unmerged_target(address)?;
                    StagedValueEdit::SetSharedFormula(formula)
                },
                Change::Clear { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    snapshot.require_shared_string_free_target(address)?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(snapshot.edit_blocked(address));
                    }
                    snapshot.require_unmerged_target(address)?;
                    StagedValueEdit::Clear
                },
                Change::Remove { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    snapshot.require_shared_string_free_target(address)?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(snapshot.edit_blocked(address));
                    }
                    snapshot.require_unmerged_target(address)?;
                    StagedValueEdit::Remove
                },
            };
            before.check_execution()?;
            pending.push((position, address, value));
        }
        before.check_execution()?;
        let original_staged_cells = self.staged_cells;
        let pending_len = pending.len();
        let mut rollback = Vec::new();
        rollback
            .try_reserve_exact(pending.len())
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet staging rollback metadata",
                source,
            })?;
        for (position, _, _) in &pending {
            if !rollback
                .iter()
                .any(|(stored_position, _)| stored_position == position)
            {
                let original_len = self.staged.get(position).map_or(0, |entries| entries.len());
                rollback.push((*position, original_len));
            }
        }
        before.check_execution()?;
        for (position, address, value) in pending {
            self.staged
                .entry(position)
                .or_default()
                .push((address, value));
        }
        self.staged_cells = original_staged_cells + pending_len;
        if let Err(error) = before.check_execution() {
            for (position, original_len) in rollback {
                let remove_position = self.staged.get_mut(&position).is_some_and(|entries| {
                    entries.truncate(original_len);
                    entries.is_empty()
                });
                if remove_position {
                    self.staged.remove(&position);
                }
            }
            self.staged_cells = original_staged_cells;
            return Err(error);
        }
        Ok(())
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.pending = None;
        self.staged.clear();
        self.staged_cells = 0;
    }

    /// Validate and rewrite the staged changes and return the committed
    /// workbook, the exact patch and the diagnostics.
    pub fn commit(mut self) -> Result<Commit> {
        self.flush_pending()?;
        let before = self.resolve_before()?.clone();
        before.check_execution()?;
        let mut expected_actions = 0usize;
        for snapshot in before.sheets() {
            if let Some(staged) = self.staged.get(&snapshot.sheet_position()) {
                expected_actions = expected_actions
                    .checked_add(effective_action_count(snapshot, staged)?)
                    .ok_or_else(|| invalid("expanded value-only action count overflows usize"))?;
                if expected_actions > MAX_BATCH_EDITS {
                    return Err(Error::limit(
                        super::LimitKind::BatchEdits,
                        expected_actions,
                        MAX_BATCH_EDITS,
                    ));
                }
            }
        }
        let mut after = Vec::new();
        after
            .try_reserve_exact(before.len())
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet candidate snapshots",
                source,
            })?;
        let mut changed_cells = 0usize;
        let mut touched_worksheets = 0usize;
        let mut aggregate_bytes = 0usize;
        for snapshot in before.sheets() {
            snapshot.check_execution()?;
            let Some(staged) = self.staged.get(&snapshot.sheet_position()) else {
                aggregate_bytes = super::snapshot::checked_multi_bytes(
                    aggregate_bytes,
                    snapshot.source_xml().len(),
                    super::snapshot::MAX_MULTI_WORKSHEET_BYTES,
                )?;
                after.push(snapshot.clone());
                continue;
            };
            let mut actions = BTreeMap::new();
            for (address, value) in staged {
                append_actions(snapshot, *address, value, &mut actions)?;
            }
            if actions.is_empty() {
                aggregate_bytes = super::snapshot::checked_multi_bytes(
                    aggregate_bytes,
                    snapshot.source_xml().len(),
                    super::snapshot::MAX_MULTI_WORKSHEET_BYTES,
                )?;
                after.push(snapshot.clone());
                continue;
            }
            changed_cells += actions.len();
            touched_worksheets += 1;
            let output = rewrite_value_only_with_provenance(
                snapshot.source_xml(),
                snapshot.sheet_name(),
                actions,
                snapshot.source_facts(),
            )?;
            aggregate_bytes = super::snapshot::checked_multi_bytes(
                aggregate_bytes,
                output.bytes.len(),
                super::snapshot::MAX_MULTI_WORKSHEET_BYTES,
            )?;
            let candidate = Snapshot::from_rewritten_value_source(snapshot, output)?;
            for (address, expected) in staged {
                let matches = match expected {
                    StagedValueEdit::Set(content) => {
                        candidate.cell(*address) == Some(&content.as_cell())
                    },
                    StagedValueEdit::Insert(content) => {
                        candidate.cell(*address) == Some(&content.as_cell())
                    },
                    StagedValueEdit::SetSharedFormula(formula) => shared_formula_readback(
                        snapshot,
                        &candidate,
                        *address,
                        formula,
                        shared_formula_changed(snapshot, *address, formula)?,
                    )?,
                    StagedValueEdit::Clear => {
                        candidate.contains_cell(*address) && candidate.value(*address).is_none()
                    },
                    StagedValueEdit::Remove => !candidate.contains_cell(*address),
                };
                if !matches {
                    return Err(invalid(
                        "multi-sheet value-only publication readback differs from staged state",
                    ));
                }
            }
            after.push(candidate);
        }
        if changed_cells != 0 {
            let workbook = before.sheets()[0].invalidated_workbook_xml()?;
            for snapshot in &mut after {
                *snapshot = snapshot
                    .clone()
                    .with_invalidated_workbook(Arc::clone(&workbook))?;
            }
        }
        before.check_execution()?;
        let after = MultiSnapshot::from_sheets(after)?;
        super::commit_change(
            self.source.workbook(),
            Record::new(before, after, changed_cells, touched_worksheets),
        )
    }
}

fn own_sheet_change(request: SheetChange<'_>) -> Result<SheetChange<'static>> {
    Ok(SheetChange {
        selector: own_selector(request.selector)?,
        edit: request.edit,
    })
}

fn effective_action_count(
    snapshot: &Snapshot,
    staged: &[(Address, StagedValueEdit)],
) -> Result<usize> {
    let mut count = 0usize;
    for (address, value) in staged {
        let next = match value {
            StagedValueEdit::Set(content) => {
                usize::from(snapshot.cell(*address) != Some(&content.as_cell()))
            },
            StagedValueEdit::Insert(_) => 1,
            StagedValueEdit::SetSharedFormula(formula) => {
                let group = snapshot.shared_formula_group(*address, MAX_BATCH_EDITS)?;
                let current = snapshot.cell(group.master).ok_or_else(|| {
                    invalid("shared formula master disappeared from the source snapshot")
                })?;
                let Cell::Formula(current) = current else {
                    return Err(invalid("shared formula master is not a scalar formula"));
                };
                if current.text() == formula.text() {
                    0
                } else {
                    group.members.len()
                }
            },
            StagedValueEdit::Clear | StagedValueEdit::Remove => 1,
        };
        count = count
            .checked_add(next)
            .ok_or_else(|| invalid("expanded value-only action count overflows usize"))?;
        if count > MAX_BATCH_EDITS {
            return Err(invalid(format!(
                "value-only batch exceeds {MAX_BATCH_EDITS} expanded cell actions"
            )));
        }
    }
    Ok(count)
}

fn append_actions(
    snapshot: &Snapshot,
    address: Address,
    value: &StagedValueEdit,
    actions: &mut BTreeMap<Address, Action>,
) -> Result<()> {
    match value {
        StagedValueEdit::Set(content) => {
            if snapshot.cell(address) != Some(&content.as_cell()) {
                actions.insert(address, Action::set(content.clone()));
            }
        },
        StagedValueEdit::Insert(content) => {
            if snapshot.cell(address).is_some() {
                return Err(invalid(format!(
                    "cell selector '{address}' already has an existing cell owner"
                )));
            }
            actions.insert(address, Action::set(content.clone()));
        },
        StagedValueEdit::SetSharedFormula(formula) => {
            let group = snapshot.shared_formula_group(address, MAX_BATCH_EDITS)?;
            let current = snapshot.cell(group.master).ok_or_else(|| {
                invalid("shared formula master disappeared from the source snapshot")
            })?;
            let Cell::Formula(current) = current else {
                return Err(invalid("shared formula master is not a scalar formula"));
            };
            if current.text() == formula.text() {
                return Ok(());
            }
            for member in &group.members {
                let replacement = (*member == group.master).then(|| formula.clone());
                let action = Action::set_shared_formula(
                    group.storage.index,
                    group.storage.reference.clone(),
                    replacement,
                );
                if actions.insert(*member, action).is_some() {
                    return Err(invalid("shared formula groups overlap in one cell edit"));
                }
            }
        },
        StagedValueEdit::Clear => {
            actions.insert(address, Action::clear(false));
        },
        StagedValueEdit::Remove => {
            actions.insert(address, Action::Remove);
        },
    }
    Ok(())
}

fn shared_formula_changed(
    snapshot: &Snapshot,
    address: Address,
    formula: &Formula,
) -> Result<bool> {
    let group = snapshot.shared_formula_group(address, MAX_BATCH_EDITS)?;
    let current = snapshot
        .cell(group.master)
        .ok_or_else(|| invalid("shared formula master disappeared from the source snapshot"))?;
    let Cell::Formula(current) = current else {
        return Err(invalid("shared formula master is not a scalar formula"));
    };
    Ok(current.text() != formula.text())
}

fn shared_formula_readback(
    before: &Snapshot,
    after: &Snapshot,
    address: Address,
    formula: &Formula,
    require_cacheless: bool,
) -> Result<bool> {
    let expected = before.shared_formula_group(address, MAX_BATCH_EDITS)?;
    let actual = after.shared_formula_group(address, MAX_BATCH_EDITS)?;
    if expected != actual {
        return Ok(false);
    }
    let Some(Cell::Formula(master)) = after.cell(expected.master) else {
        return Ok(false);
    };
    if master.text() != formula.text() {
        return Ok(false);
    }
    if require_cacheless
        && expected.members.iter().any(|member| {
            !matches!(after.cell(*member), Some(Cell::Formula(formula)) if formula.cached().is_none())
        })
    {
        return Ok(false);
    }
    Ok(true)
}

fn resolve_snapshot_selector(snapshots: &MultiSnapshot, selector: &Selector<'_>) -> Result<usize> {
    match selector {
        litchi_core::Selector::Position(position) => snapshots
            .sheets()
            .iter()
            .find(|snapshot| snapshot.sheet_position() == position.get())
            .map(Snapshot::sheet_position)
            .ok_or_else(|| invalid("multi-sheet edit selector was not selected")),
        litchi_core::Selector::Name(name) => snapshots
            .sheets()
            .iter()
            .find(|snapshot| crate::sheet::key(snapshot.sheet_name()) == crate::sheet::key(name))
            .map(Snapshot::sheet_position)
            .ok_or_else(|| invalid("multi-sheet edit selector was not selected")),
        litchi_core::Selector::Id(_) => Err(Error::UnsupportedSelector),
        _ => Err(Error::UnsupportedSelector),
    }
}
