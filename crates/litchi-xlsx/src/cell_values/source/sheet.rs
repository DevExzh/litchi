//! Value-only edits of one worksheet of a [`SourceBackedWorkbook`]: the
//! single-worksheet form of the [`crate::cell_values`] family (record 0360).
//!
//! [`SourceBackedWorkbook::edit_sheet_values`] stages value-only changes of
//! one worksheet; [`Edit::commit`] returns a [`Commit`] whose output is the
//! one-worksheet topology publication; [`SourceBackedWorkbook::apply_sheet_values`]
//! replays a [`Patch`].

use std::collections::BTreeMap;
use std::io::Write;

use litchi_opc::SourceBackedPackage;
use litchi_sheet::Cell as Address;

use super::super::Snapshot;
use super::super::patch::SheetRecord;
use super::super::snapshot::SourceProvenance;
use super::{
    Change, MAX_BATCH_EDITS, StagedValueEdit, append_actions, effective_action_count,
    shared_formula_changed, shared_formula_readback,
};
use crate::cell::{Cell, Content, Number, Value};
use crate::error::{Result, invalid};
use crate::formula::Formula;
use crate::raw::worksheet::edit::rewrite_value_only_with_provenance;
use crate::workbook::source::{LazySource, own_selector};
use crate::{Selector, SourceBackedWorkbook};

crate::workbook::source::deferred::source_family!(
    "worksheet value",
    SheetRecord,
    own_limits,
    own_diagnostics
);

pub use super::super::LimitKind;

/// What a committed one-worksheet value change changed.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_cells: usize,
}

impl Diagnostics {
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

    /// Number of worksheet parts replaced (zero or one).
    #[must_use]
    pub const fn touched_worksheets(&self) -> usize {
        if self.changed_cells == 0 { 0 } else { 1 }
    }
}

impl crate::workbook::source::deferred::Report for SheetRecord {
    type Diagnostics = Diagnostics;

    fn diagnostics(&self) -> Diagnostics {
        Diagnostics {
            changed_cells: self.changed_cells(),
        }
    }
}

impl Patch {
    /// The exact worksheet state the patch requires.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact worksheet state the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        self.0.after()
    }

    /// The patch with its target worksheet payload copied out of a managed
    /// source's reservation under an explicit `maximum_bytes` bound, so a
    /// patch made over a managed source replays onto a workbook outside that
    /// budget.
    ///
    /// # Errors
    ///
    /// Returns an error when the payload exceeds the bound.
    pub fn materialized(&self, maximum_bytes: usize) -> Result<Self> {
        Ok(Self(
            std::sync::Arc::new(self.0.materialized(maximum_bytes)?),
            None,
        ))
    }
}

impl crate::workbook::source::deferred::Change for SheetRecord {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        SheetRecord::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        let before = self.before();
        let conflict = || Error::PatchConflict {
            part: before.worksheet_part_name().to_string(),
        };
        match before.matches_source_backed(package)? {
            SourceProvenance::Matched => Ok(()),
            SourceProvenance::Mismatched if before.same_lineage(package) => Err(conflict()),
            SourceProvenance::Mismatched | SourceProvenance::Unavailable => {
                let current = Snapshot::load_source_backed(package, before.sheet_position())?;
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
        write_record(package, self, writer)
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        package.check_execution()?;
        write_record(package, self, writer)
    }
}

/// Write `record`'s one-worksheet topology publication over `package`.
pub(crate) fn write_record(
    package: &SourceBackedPackage,
    record: &SheetRecord,
    writer: &mut dyn Write,
) -> Result<()> {
    let plan = record.after().topology_plan_from(record.before())?;
    package.clone().write_topology_to_stream(writer, plan)?;
    Ok(())
}

impl SourceBackedWorkbook {
    /// Begin a value-only edit of one worksheet.
    ///
    /// # Errors
    ///
    /// Value-closure and selector errors are deferred until first use.
    pub fn edit_sheet_values<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(MAX_BATCH_EDITS)
            .map_err(|source| Error::Allocation {
                resource: "value-only batch staging",
                source,
            })?;
        self.check_execution()?;
        Ok(Edit {
            source: LazySource::new(self.clone()),
            selector: own_selector(selector.into())?,
            staged,
        })
    }

    /// Replay a one-worksheet value patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the worksheet does not hold
    /// the patch's before state, and the signed-source refusal for a changed
    /// patch on a signed source.
    pub fn apply_sheet_values(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// Clone-staged atomic value changes over one exact source worksheet of a
/// [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    source: LazySource<Snapshot>,
    selector: Selector<'static>,
    staged: Vec<(Address, StagedValueEdit)>,
}

impl Edit {
    fn resolve_before(&self) -> Result<&Snapshot> {
        let selector = self.selector.clone();
        self.source
            .resolve(|package| Snapshot::load_source_backed(package, selector))
    }

    /// Resolve the exact source snapshot observed by this edit.
    pub fn before(&self) -> Result<&Snapshot> {
        self.resolve_before()
    }

    /// Number of unique staged selectors.
    #[must_use]
    pub fn len(&self) -> usize {
        self.staged.len()
    }

    /// Whether no selectors have been staged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.staged.is_empty()
    }

    /// Stage one numeric insertion for an absent coordinate.
    pub fn insert(&mut self, address: Address, value: Number) -> Result<()> {
        self.apply_batch([Change::insert(address, value)])
    }

    /// Stage one value replacement. Repeated selectors are rejected.
    pub fn set(&mut self, address: Address, value: impl Into<Value>) -> Result<()> {
        self.apply_batch([Change::set(address, value)])
    }

    /// Stage one cacheless scalar formula replacement.
    pub fn set_formula(&mut self, address: Address, formula: Formula) -> Result<()> {
        self.apply_batch([Change::set_formula(address, formula)])
    }

    /// Stage one shared-formula master replacement.
    pub fn set_shared_formula(&mut self, address: Address, formula: Formula) -> Result<()> {
        self.apply_batch([Change::set_shared_formula(address, formula)])
    }

    /// Stage one scalar clear. Repeated selectors are rejected.
    pub fn clear(&mut self, address: Address) -> Result<()> {
        self.apply_batch([Change::clear(address)])
    }

    /// Stage complete removal of one scalar `<c>` owner.
    ///
    /// The containing `<row>` is retained even when this is its last cell.
    /// Repeated selectors are rejected.
    pub fn remove(&mut self, address: Address) -> Result<()> {
        self.apply_batch([Change::remove(address)])
    }

    /// Stage a bounded batch atomically.
    pub fn apply_batch(&mut self, edits: impl IntoIterator<Item = Change>) -> Result<()> {
        let before = self.resolve_before()?.clone();
        before.check_execution()?;
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS.saturating_sub(self.staged.len()))
            .map_err(|source| Error::Allocation {
                resource: "value-only pending batch",
                source,
            })?;
        for edit in edits {
            before.check_execution()?;
            if self.staged.len() + pending.len() >= MAX_BATCH_EDITS {
                return Err(Error::limit(
                    LimitKind::BatchEdits,
                    MAX_BATCH_EDITS + 1,
                    MAX_BATCH_EDITS,
                ));
            }
            let address = edit.address();
            if self
                .staged
                .iter()
                .chain(&pending)
                .any(|(stored, _)| *stored == address)
            {
                return Err(invalid(format!(
                    "duplicate value-only selector '{address}'"
                )));
            }
            let source = before.editable_cell(address);
            let value = match edit {
                Change::Insert { value, .. } => {
                    before.require_insertable_absence(address)?;
                    let value = Value::Number(value);
                    value.validate_for_write()?;
                    StagedValueEdit::Insert(Content::Value(value))
                },
                Change::Set { value, .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    before.require_shared_string_free_target(address)?;
                    value.validate_for_write()?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(before.edit_blocked(address));
                    }
                    before.require_unmerged_target(address)?;
                    StagedValueEdit::Set(Content::Value(value))
                },
                Change::SetFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    before.require_shared_string_free_target(address)?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    before.require_formula_target(address)?;
                    before.require_unmerged_target(address)?;
                    StagedValueEdit::Set(Content::Formula(formula))
                },
                Change::SetSharedFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    before.require_shared_string_free_target(address)?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    before.shared_formula_group(address, MAX_BATCH_EDITS)?;
                    before.require_unmerged_target(address)?;
                    StagedValueEdit::SetSharedFormula(formula)
                },
                Change::Clear { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    before.require_shared_string_free_target(address)?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(before.edit_blocked(address));
                    }
                    before.require_unmerged_target(address)?;
                    StagedValueEdit::Clear
                },
                Change::Remove { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    before.require_shared_string_free_target(address)?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(before.edit_blocked(address));
                    }
                    before.require_unmerged_target(address)?;
                    StagedValueEdit::Remove
                },
            };
            before.check_execution()?;
            pending.push((address, value));
        }
        before.check_execution()?;
        let original_len = self.staged.len();
        self.staged.extend(pending);
        if let Err(error) = before.check_execution() {
            self.staged.truncate(original_len);
            return Err(error);
        }
        Ok(())
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.staged.clear();
    }

    /// Validate and rewrite the staged changes and return the committed
    /// workbook, the exact patch and the diagnostics.
    pub fn commit(self) -> Result<Commit> {
        let before = self.resolve_before()?.clone();
        before.check_execution()?;
        let expected_actions = effective_action_count(&before, &self.staged)?;
        if expected_actions > MAX_BATCH_EDITS {
            return Err(Error::limit(
                LimitKind::BatchEdits,
                expected_actions,
                MAX_BATCH_EDITS,
            ));
        }
        let mut actions = BTreeMap::new();
        for (address, value) in &self.staged {
            append_actions(&before, *address, value, &mut actions)?;
        }
        if actions.is_empty() {
            let patch = SheetRecord::new(before.clone(), before, 0);
            return commit_change(self.source.workbook(), patch);
        }
        let changed = actions.len();
        let output = rewrite_value_only_with_provenance(
            before.source_xml(),
            before.sheet_name(),
            actions,
            before.source_facts(),
        )?;
        let snapshot = Snapshot::from_rewritten_value_source(&before, output)?
            .with_invalidated_calculation()?;
        for (address, expected) in &self.staged {
            let matches = match expected {
                StagedValueEdit::Set(content) => {
                    snapshot.cell(*address) == Some(&content.as_cell())
                },
                StagedValueEdit::Insert(content) => {
                    snapshot.cell(*address) == Some(&content.as_cell())
                },
                StagedValueEdit::SetSharedFormula(formula) => shared_formula_readback(
                    &before,
                    &snapshot,
                    *address,
                    formula,
                    shared_formula_changed(&before, *address, formula)?,
                )?,
                StagedValueEdit::Clear => {
                    snapshot.contains_cell(*address) && snapshot.value(*address).is_none()
                },
                StagedValueEdit::Remove => !snapshot.contains_cell(*address),
            };
            if !matches {
                return Err(invalid(
                    "value-only publication readback differs from staged state",
                ));
            }
        }
        before.check_execution()?;
        commit_change(
            self.source.workbook(),
            SheetRecord::new(before, snapshot, changed),
        )
    }
}
