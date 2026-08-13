//! Guarded source-backed value-only worksheet transactions.

use std::collections::BTreeMap;
use std::io::Write;
use std::sync::Arc;

use litchi_core::ReadAt;
use litchi_opc::{ReadLimits, SourceBackedPackage};
use litchi_sheet::Cell as Address;

use super::{Commit, MAX_SHEET_OWNERS, MultiCommit, MultiPatch, MultiSnapshot, Patch, Snapshot};
use crate::Selector;
use crate::cell::{Content, Value};
use crate::error::{Error, Result, invalid};
use crate::raw::worksheet::edit::{Action, rewrite};

/// Maximum unique existing cells in one atomic value transaction.
pub const MAX_BATCH_EDITS: usize = 256;

/// One typed value-only edit for an existing scalar cell.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellValueEdit {
    /// Replace the stored scalar value without changing local style.
    Set { address: Address, value: Value },
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
pub struct SheetCellValueEdit<'a> {
    /// Worksheet selected by its semantic name or zero-based position.
    pub selector: Selector<'a>,
    /// Scalar operation applied to an existing cell owner.
    pub edit: CellValueEdit,
}

impl<'a> SheetCellValueEdit<'a> {
    /// Construct a selector-first value replacement.
    pub fn set(
        selector: impl Into<Selector<'a>>,
        address: Address,
        value: impl Into<Value>,
    ) -> Self {
        Self {
            selector: selector.into(),
            edit: CellValueEdit::set(address, value),
        }
    }

    /// Construct a selector-first scalar clear.
    #[must_use]
    pub fn clear(selector: impl Into<Selector<'a>>, address: Address) -> Self {
        Self {
            selector: selector.into(),
            edit: CellValueEdit::clear(address),
        }
    }

    /// Construct a selector-first complete cell-owner removal.
    #[must_use]
    pub fn remove(selector: impl Into<Selector<'a>>, address: Address) -> Self {
        Self {
            selector: selector.into(),
            edit: CellValueEdit::remove(address),
        }
    }
}

impl CellValueEdit {
    /// Construct a checked value replacement.
    pub fn set(address: Address, value: impl Into<Value>) -> Self {
        Self::Set {
            address,
            value: value.into(),
        }
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
            Self::Set { address, .. } | Self::Clear { address } | Self::Remove { address } => {
                *address
            },
        }
    }
}

#[derive(Clone, Debug)]
enum StagedValueEdit {
    Set(Value),
    Clear,
    Remove,
}

/// Owning source-backed editor for one exact XLSX artifact.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// Clone-staged atomic changes over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: Vec<(Address, StagedValueEdit)>,
}

/// Clone-staged atomic changes over a bounded worksheet set.
pub struct MultiSourceEdit {
    before: MultiSnapshot,
    staged: BTreeMap<usize, Vec<(Address, StagedValueEdit)>>,
    staged_cells: usize,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open with explicit OPC ingress limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Ok(Self {
            package: SourceBackedPackage::from_read_at_with_limits(source, limits)?,
        })
    }

    /// Capture the exact safe value-only closure for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin one atomic value-only edit.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(MAX_BATCH_EDITS)
            .map_err(|source| Error::Allocation {
                resource: "value-only batch staging",
                source,
            })?;
        Ok(SourceEdit {
            before: self.snapshot(selector)?,
            staged,
        })
    }

    /// Begin a bounded transaction over selected worksheet owners.
    pub fn edit_sheets<'a, I>(&self, selectors: I) -> Result<MultiSourceEdit>
    where
        I: IntoIterator<Item = Selector<'a>>,
    {
        MultiSourceEdit::new(MultiSnapshot::load_source_backed(&self.package, selectors)?)
    }

    /// Begin a bounded selector-first transaction and stage its initial batch.
    pub fn edit_many<'a, I>(&self, edits: I) -> Result<MultiSourceEdit>
    where
        I: IntoIterator<Item = SheetCellValueEdit<'a>>,
    {
        let mut requests = edits.into_iter();
        let first = requests
            .next()
            .ok_or_else(|| invalid("multi-sheet value edits require one cell edit"))?;
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
            if pending.len() >= MAX_BATCH_EDITS {
                return Err(invalid(format!(
                    "multi-sheet value-only batch exceeds {MAX_BATCH_EDITS} unique cells"
                )));
            }
            if !selectors
                .iter()
                .any(|selector| selector == &request.selector)
            {
                if selectors.len() >= MAX_SHEET_OWNERS {
                    return Err(invalid(format!(
                        "multi-sheet value edits exceed {MAX_SHEET_OWNERS} worksheet owners"
                    )));
                }
                selectors.push(request.selector.clone());
            }
            pending.push(request);
        }
        let mut transaction = self.edit_sheets(selectors)?;
        transaction.apply_batch(pending)?;
        Ok(transaction)
    }

    /// Content-free deferred-Part cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Publish one exact-source-checked worksheet overlay to a sequential sink.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        let current =
            Snapshot::load_source_backed(&self.package, commit.patch().before().sheet_position())?;
        if !current.same_source(commit.patch().before()) {
            return Err(Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            current
        } else {
            commit.patch().after().clone()
        };
        self.write_snapshot_overlay_to_stream(writer, &target)?;
        Ok(target)
    }

    pub(crate) fn write_snapshot_overlay_to_stream<W: Write>(
        self,
        writer: W,
        target: &Snapshot,
    ) -> Result<()> {
        self.package.write_part_overlay_to_stream(
            writer,
            target.worksheet_part_name(),
            target.source_xml().to_vec(),
        )?;
        Ok(())
    }

    /// Publish a bounded multi-worksheet commit through one atomic overlay.
    pub fn publish_multi_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &MultiCommit,
    ) -> Result<MultiSnapshot> {
        let current = MultiSnapshot::load_source_backed(
            &self.package,
            commit
                .patch()
                .before()
                .sheets()
                .iter()
                .map(|snapshot| snapshot.sheet_position().into()),
        )?;
        if !current.same_source(commit.patch().before()) {
            return Err(Error::PatchConflict {
                part: commit
                    .patch()
                    .before()
                    .sheets()
                    .first()
                    .map_or_else(String::new, |snapshot| {
                        snapshot.worksheet_part_name().to_string()
                    }),
            });
        }
        let target = if commit.patch().is_empty() {
            current
        } else {
            commit.patch().after().clone()
        };
        let replacements = target
            .sheets()
            .iter()
            .zip(commit.patch().before().sheets())
            .filter(|(after, before)| after.source_xml() != before.source_xml())
            .map(|(snapshot, _)| {
                (
                    snapshot.worksheet_part_name().clone(),
                    snapshot.source_xml().to_vec(),
                )
            })
            .collect();
        self.package
            .write_part_overlays_to_stream(writer, replacements)?;
        Ok(target)
    }
}

impl SourceEdit {
    /// Exact source snapshot captured at transaction start.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
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

    /// Stage one value replacement. Repeated selectors are rejected.
    pub fn set(&mut self, address: Address, value: impl Into<Value>) -> Result<()> {
        self.apply_batch([CellValueEdit::set(address, value)])
    }

    /// Stage one scalar clear. Repeated selectors are rejected.
    pub fn clear(&mut self, address: Address) -> Result<()> {
        self.apply_batch([CellValueEdit::clear(address)])
    }

    /// Stage complete removal of one scalar `<c>` owner.
    ///
    /// The containing `<row>` is retained even when this is its last cell.
    /// Repeated selectors are rejected.
    pub fn remove(&mut self, address: Address) -> Result<()> {
        self.apply_batch([CellValueEdit::remove(address)])
    }

    /// Stage a bounded batch atomically.
    pub fn apply_batch(&mut self, edits: impl IntoIterator<Item = CellValueEdit>) -> Result<()> {
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS.saturating_sub(self.staged.len()))
            .map_err(|source| Error::Allocation {
                resource: "value-only pending batch",
                source,
            })?;
        for edit in edits {
            if self.staged.len() + pending.len() >= MAX_BATCH_EDITS {
                return Err(invalid(format!(
                    "value-only batch exceeds {MAX_BATCH_EDITS} unique cells"
                )));
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
            let source = self.before.value(address).ok_or_else(|| {
                invalid(format!(
                    "value-only selector '{address}' is not an existing scalar value cell"
                ))
            })?;
            let value = match edit {
                CellValueEdit::Set { value, .. } => {
                    value.validate_for_write()?;
                    if matches!(value, Value::Date(_)) {
                        return Err(invalid("value-only batches currently refuse date cells"));
                    }
                    StagedValueEdit::Set(value)
                },
                CellValueEdit::Clear { .. } => StagedValueEdit::Clear,
                CellValueEdit::Remove { .. } => StagedValueEdit::Remove,
            };
            if matches!(source, Value::Date(_)) {
                return Err(invalid("value-only batches currently refuse date cells"));
            }
            pending.push((address, value));
        }
        self.staged.extend(pending);
        Ok(())
    }

    /// Validate, rewrite once, and freeze an exact reversible commit.
    pub fn commit(self) -> Result<Commit> {
        let mut actions = BTreeMap::new();
        for (address, value) in &self.staged {
            let action = match value {
                StagedValueEdit::Set(value) => {
                    if Some(value) == self.before.value(*address) {
                        continue;
                    }
                    Action::set(Content::Value(value.clone()))
                },
                StagedValueEdit::Clear => Action::clear(false),
                StagedValueEdit::Remove => Action::Remove,
            };
            actions.insert(*address, action);
        }
        if actions.is_empty() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, 0));
        }
        let changed = actions.len();
        let output = rewrite(self.before.source_xml(), self.before.sheet_name(), actions)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?;
        for (address, expected) in &self.staged {
            let matches = match expected {
                StagedValueEdit::Set(value) => snapshot.value(*address) == Some(value),
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
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, changed))
    }
}

impl MultiSourceEdit {
    fn new(before: MultiSnapshot) -> Result<Self> {
        Ok(Self {
            before,
            staged: BTreeMap::new(),
            staged_cells: 0,
        })
    }

    /// Exact source snapshots captured at transaction start.
    #[must_use]
    pub const fn before(&self) -> &MultiSnapshot {
        &self.before
    }

    /// Number of unique staged cell selectors across all worksheets.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.staged_cells
    }

    /// Whether no cell selectors have been staged.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.staged_cells == 0
    }

    /// Number of selected worksheet owners.
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.before.len()
    }

    /// Stage one selector-first scalar replacement.
    pub fn set<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        value: impl Into<Value>,
    ) -> Result<()> {
        self.apply_batch([SheetCellValueEdit::set(selector, address, value)])
    }

    /// Stage one selector-first scalar clear.
    pub fn clear<'a>(&mut self, selector: impl Into<Selector<'a>>, address: Address) -> Result<()> {
        self.apply_batch([SheetCellValueEdit::clear(selector, address)])
    }

    /// Stage one selector-first complete cell-owner removal.
    pub fn remove<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
    ) -> Result<()> {
        self.apply_batch([SheetCellValueEdit::remove(selector, address)])
    }

    /// Stage a bounded cross-worksheet batch atomically.
    pub fn apply_batch<'a, I>(&mut self, edits: I) -> Result<()>
    where
        I: IntoIterator<Item = SheetCellValueEdit<'a>>,
    {
        let remaining = MAX_BATCH_EDITS.saturating_sub(self.staged_cells);
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(remaining)
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet pending value batch",
                source,
            })?;
        for request in edits {
            if pending.len() >= remaining {
                return Err(invalid(format!(
                    "multi-sheet value-only batch exceeds {MAX_BATCH_EDITS} unique cells"
                )));
            }
            let position = resolve_snapshot_selector(&self.before, &request.selector)?;
            let snapshot = self
                .before
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
            let source = snapshot.value(address).ok_or_else(|| {
                invalid(format!(
                    "value-only selector '{}!{}' is not an existing scalar value cell",
                    snapshot.sheet_name(),
                    address
                ))
            })?;
            let value = match request.edit {
                CellValueEdit::Set { value, .. } => {
                    value.validate_for_write()?;
                    if matches!(value, Value::Date(_)) {
                        return Err(invalid("value-only batches currently refuse date cells"));
                    }
                    StagedValueEdit::Set(value)
                },
                CellValueEdit::Clear { .. } => StagedValueEdit::Clear,
                CellValueEdit::Remove { .. } => StagedValueEdit::Remove,
            };
            if matches!(source, Value::Date(_)) {
                return Err(invalid("value-only batches currently refuse date cells"));
            }
            pending.push((position, address, value));
        }
        for (position, address, value) in pending {
            self.staged
                .entry(position)
                .or_default()
                .push((address, value));
            self.staged_cells += 1;
        }
        Ok(())
    }

    /// Validate, rewrite, and freeze one atomic multi-worksheet commit.
    pub fn commit(self) -> Result<MultiCommit> {
        let mut after = Vec::new();
        after
            .try_reserve_exact(self.before.len())
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet candidate snapshots",
                source,
            })?;
        let mut changed_cells = 0usize;
        let mut touched_worksheets = 0usize;
        let mut aggregate_bytes = 0usize;
        for snapshot in self.before.sheets() {
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
                let action = match value {
                    StagedValueEdit::Set(value) => {
                        if Some(value) == snapshot.value(*address) {
                            continue;
                        }
                        Action::set(Content::Value(value.clone()))
                    },
                    StagedValueEdit::Clear => Action::clear(false),
                    StagedValueEdit::Remove => Action::Remove,
                };
                actions.insert(*address, action);
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
            let output = rewrite(snapshot.source_xml(), snapshot.sheet_name(), actions)?;
            aggregate_bytes = super::snapshot::checked_multi_bytes(
                aggregate_bytes,
                output.len(),
                super::snapshot::MAX_MULTI_WORKSHEET_BYTES,
            )?;
            let candidate = Snapshot::from_rewritten_source(snapshot, output)?;
            for (address, expected) in staged {
                let matches = match expected {
                    StagedValueEdit::Set(value) => candidate.value(*address) == Some(value),
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
        let after = MultiSnapshot::from_sheets(after)?;
        let patch = MultiPatch::new(self.before, after.clone());
        Ok(MultiCommit::new(
            after,
            patch,
            changed_cells,
            touched_worksheets,
        ))
    }
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
