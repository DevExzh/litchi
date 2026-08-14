//! Guarded source-backed value-only worksheet transactions.

use std::collections::BTreeMap;
use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ReadAt};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits};
use litchi_sheet::Cell as Address;

use super::snapshot::SourceProvenance;
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
///
/// On managed opens, the caller-provided [`Budget`](litchi_core::Budget)
/// accounts for retained and in-flight OPC [`PartData`](litchi_opc::PartData)
/// payload reservations. Parsed cell stores, graph and relationship metadata,
/// staging allocations, rewritten XML, candidate package clones, and output
/// buffers remain governed by this facade's separate bounds and are not
/// charged to that payload budget.
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
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Open with an explicit finite payload-cache policy. This compatibility
    /// path is bounded but unmanaged; use an execution-context constructor to
    /// charge retained [`PartData`](litchi_opc::PartData) handles to a caller
    /// budget.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open with explicit read and finite payload-cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                limits,
                cache_limits,
            )?,
        )
    }

    /// Open with an explicit caller-owned execution context and the default
    /// finite source cache. The context charges retained and in-flight OPC
    /// `PartData` payloads; parsed semantic state and rewritten/output buffers
    /// remain outside that payload accounting and use separate bounds.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source, limits, context,
        )?)
    }

    /// Open with explicit read and execution policies.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, limits, context)
    }

    /// Open with explicit read, cache, and caller-owned execution policies.
    /// The managed budget covers retained and in-flight OPC `PartData` payload
    /// reservations only; parsed stores, metadata, rewritten candidates, and
    /// output buffers are not charged to it.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                limits,
                cache_limits,
                context,
            )?,
        )
    }

    /// Build an editor from a validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        Ok(Self { package })
    }

    /// Capture the exact safe value-only closure for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Snapshot> {
        self.package.check_execution()?;
        Snapshot::load_source_backed(&self.package, selector)
    }

    /// Begin one atomic value-only edit.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<SourceEdit> {
        self.package.check_execution()?;
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
        self.package.check_execution()?;
        MultiSourceEdit::new(MultiSnapshot::load_source_backed(&self.package, selectors)?)
    }

    /// Begin a bounded selector-first transaction and stage its initial batch.
    pub fn edit_many<'a, I>(&self, edits: I) -> Result<MultiSourceEdit>
    where
        I: IntoIterator<Item = SheetCellValueEdit<'a>>,
    {
        self.package.check_execution()?;
        let mut requests = edits.into_iter();
        let first = requests.next();
        self.package.check_execution()?;
        let first =
            first.ok_or_else(|| invalid("multi-sheet value edits require one cell edit"))?;
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
            self.package.check_execution()?;
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
        self.package.check_execution()?;
        let before = commit.patch().before();
        let current = match before.matches_source_backed(&self.package)? {
            SourceProvenance::Matched => None,
            SourceProvenance::Mismatched => {
                return Err(Error::PatchConflict {
                    part: before.worksheet_part_name().to_string(),
                });
            },
            SourceProvenance::Unavailable => Some(Snapshot::load_source_backed(
                &self.package,
                before.sheet_position(),
            )?),
        };
        if let Some(current) = &current
            && !current.same_source(before)
        {
            return Err(Error::PatchConflict {
                part: current.worksheet_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            current.unwrap_or_else(|| before.clone())
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
        self.package.check_execution()?;
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
        self.package.check_execution()?;
        let before = commit.patch().before();
        let current = match before.matches_source_backed(&self.package)? {
            SourceProvenance::Matched => None,
            SourceProvenance::Mismatched => {
                return Err(Error::PatchConflict {
                    part: before
                        .sheets()
                        .first()
                        .map_or_else(String::new, |snapshot| {
                            snapshot.worksheet_part_name().to_string()
                        }),
                });
            },
            SourceProvenance::Unavailable => Some(MultiSnapshot::load_source_backed(
                &self.package,
                before
                    .sheets()
                    .iter()
                    .map(|snapshot| snapshot.sheet_position().into()),
            )?),
        };
        if let Some(current) = &current
            && !current.same_source(before)
        {
            return Err(Error::PatchConflict {
                part: before
                    .sheets()
                    .first()
                    .map_or_else(String::new, |snapshot| {
                        snapshot.worksheet_part_name().to_string()
                    }),
            });
        }
        let target = if commit.patch().is_empty() {
            current.unwrap_or_else(|| before.clone())
        } else {
            commit.patch().after().clone()
        };
        let replacements = target
            .sheets()
            .iter()
            .zip(before.sheets())
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
        self.before.check_execution()?;
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(MAX_BATCH_EDITS.saturating_sub(self.staged.len()))
            .map_err(|source| Error::Allocation {
                resource: "value-only pending batch",
                source,
            })?;
        for edit in edits {
            self.before.check_execution()?;
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
            self.before.check_execution()?;
            pending.push((address, value));
        }
        self.before.check_execution()?;
        let original_len = self.staged.len();
        self.staged.extend(pending);
        if let Err(error) = self.before.check_execution() {
            self.staged.truncate(original_len);
            return Err(error);
        }
        Ok(())
    }

    /// Validate, rewrite once, and freeze an exact reversible commit.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
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
            self.before.check_execution()?;
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
        self.before.check_execution()?;
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
        self.before.check_execution()?;
        let remaining = MAX_BATCH_EDITS.saturating_sub(self.staged_cells);
        let mut pending = Vec::new();
        pending
            .try_reserve_exact(remaining)
            .map_err(|source| Error::Allocation {
                resource: "multi-sheet pending value batch",
                source,
            })?;
        for request in edits {
            self.before.check_execution()?;
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
            self.before.check_execution()?;
            pending.push((position, address, value));
        }
        self.before.check_execution()?;
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
        self.before.check_execution()?;
        for (position, address, value) in pending {
            self.staged
                .entry(position)
                .or_default()
                .push((address, value));
        }
        self.staged_cells = original_staged_cells + pending_len;
        if let Err(error) = self.before.check_execution() {
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

    /// Validate, rewrite, and freeze one atomic multi-worksheet commit.
    pub fn commit(self) -> Result<MultiCommit> {
        self.before.check_execution()?;
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
        self.before.check_execution()?;
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
