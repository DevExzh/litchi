//! Guarded source-backed value-only worksheet transactions.

use std::collections::BTreeMap;
use std::io::Write;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::Arc;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{ExecutionContext, ReadAt};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits};
use litchi_sheet::Cell as Address;

use super::snapshot::SourceProvenance;
use super::{Commit, MAX_SHEET_OWNERS, MultiCommit, MultiPatch, MultiSnapshot, Patch, Snapshot};
use crate::Selector;
use crate::cell::{Cell, Content, Number, Value};
use crate::error::{Error, Result, invalid};
use crate::formula::Formula;
use crate::raw::worksheet::edit::{Action, rewrite};

/// Maximum unique cells in one atomic value transaction.
pub const MAX_BATCH_EDITS: usize = 256;

/// One typed value-only edit for a scalar cell.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellValueEdit {
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
pub struct SheetCellValueEdit<'a> {
    /// Worksheet selected by its semantic name or zero-based position.
    pub selector: Selector<'a>,
    /// Scalar operation applied to one selected worksheet.
    pub edit: CellValueEdit,
}

impl<'a> SheetCellValueEdit<'a> {
    /// Construct a selector-first numeric insertion for an absent coordinate.
    pub fn insert(selector: impl Into<Selector<'a>>, address: Address, value: Number) -> Self {
        Self {
            selector: selector.into(),
            edit: CellValueEdit::insert(address, value),
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
            edit: CellValueEdit::set(address, value),
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
            edit: CellValueEdit::set_formula(address, formula),
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
            edit: CellValueEdit::set_shared_formula(address, formula),
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
    /// Open an ordinary XLSX package from a regular filesystem path.
    ///
    /// The path is represented by an open positional [`FileSource`], so
    /// opening does not read the complete artifact into memory. Deferred
    /// worksheet payloads retain the same lazy behavior as the `ReadAt`
    /// constructors below.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path_with_limits(path, ReadLimits::default())
    }

    /// Open a filesystem-backed XLSX package with explicit OPC ingress
    /// limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a filesystem-backed XLSX package with an explicit finite
    /// deferred-payload cache policy.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_cache_limits(file_source(path)?, cache_limits)
    }

    /// Open a filesystem-backed XLSX package with explicit read and cache
    /// policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(file_source(path)?, limits, cache_limits)
    }

    /// Open a filesystem-backed XLSX package with explicit read and
    /// execution policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed XLSX package with explicit read and
    /// execution policies (the cache uses its default bounded policy).
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed XLSX package with explicit read, cache, and
    /// caller-owned execution policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            file_source(path)?,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open a filesystem-backed XLSX editor from a regular path.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    /// Open a filesystem-backed XLSX editor with explicit OPC limits.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_path_with_limits(path, limits)
    }

    /// Open a filesystem-backed XLSX editor with an explicit finite cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_cache_limits(path, cache_limits)
    }

    /// Open a filesystem-backed XLSX editor with explicit read and cache
    /// policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits(path, limits, cache_limits)
    }

    /// Open a filesystem-backed XLSX editor with explicit read and execution
    /// policies while retaining the default finite cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed XLSX editor with explicit read and execution
    /// policies while retaining the default finite cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed XLSX editor with explicit read, cache, and
    /// execution policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits_and_execution_context(
            path,
            limits,
            cache_limits,
            context,
        )
    }

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
        self.publish_patch_to_stream(writer, commit.patch())
    }

    /// Publish one exact-source-checked worksheet patch for an owning sibling
    /// facade that retains the same scalar worksheet closure.
    pub(crate) fn publish_patch_to_stream<W: Write>(
        self,
        writer: W,
        patch: &Patch,
    ) -> Result<Snapshot> {
        self.package.check_execution()?;
        let before = patch.before();
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
        let target = if patch.is_empty() {
            current.unwrap_or_else(|| before.clone())
        } else {
            patch.after().clone()
        };
        if patch.is_empty() {
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())?;
        } else {
            self.write_snapshot_overlay_to_stream(writer, before, &target)?;
        }
        Ok(target)
    }

    pub(crate) fn write_snapshot_overlay_to_stream<W: Write>(
        self,
        writer: W,
        before: &Snapshot,
        target: &Snapshot,
    ) -> Result<()> {
        self.package.check_execution()?;
        let plan = target.topology_plan_from(before)?;
        self.package.write_topology_to_stream(writer, plan)?;
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
        let plan = if commit.patch().is_empty() {
            litchi_opc::SourceTopologyPlan::new()
        } else {
            target.topology_plan_from(before)?
        };
        self.package.write_topology_to_stream(writer, plan)?;
        Ok(target)
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path).map_err(|error| {
        Error::Package(litchi_opc::OpcError::from(error))
    })?))
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

    /// Stage one numeric insertion for an absent coordinate.
    pub fn insert(&mut self, address: Address, value: Number) -> Result<()> {
        self.apply_batch([CellValueEdit::insert(address, value)])
    }

    /// Stage one value replacement. Repeated selectors are rejected.
    pub fn set(&mut self, address: Address, value: impl Into<Value>) -> Result<()> {
        self.apply_batch([CellValueEdit::set(address, value)])
    }

    /// Stage one cacheless scalar formula replacement.
    pub fn set_formula(&mut self, address: Address, formula: Formula) -> Result<()> {
        self.apply_batch([CellValueEdit::set_formula(address, formula)])
    }

    /// Stage one shared-formula master replacement.
    pub fn set_shared_formula(&mut self, address: Address, formula: Formula) -> Result<()> {
        self.apply_batch([CellValueEdit::set_shared_formula(address, formula)])
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
            let source = self.before.editable_cell(address);
            let value = match edit {
                CellValueEdit::Insert { value, .. } => {
                    self.before.require_insertable_absence(address)?;
                    let value = Value::Number(value);
                    value.validate_for_write()?;
                    StagedValueEdit::Insert(Content::Value(value))
                },
                CellValueEdit::Set { value, .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    value.validate_for_write()?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(self.before.edit_blocked(address));
                    }
                    StagedValueEdit::Set(Content::Value(value))
                },
                CellValueEdit::SetFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    self.before.require_formula_target(address)?;
                    StagedValueEdit::Set(Content::Formula(formula))
                },
                CellValueEdit::SetSharedFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    self.before.shared_formula_group(address, MAX_BATCH_EDITS)?;
                    StagedValueEdit::SetSharedFormula(formula)
                },
                CellValueEdit::Clear { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(self.before.edit_blocked(address));
                    }
                    StagedValueEdit::Clear
                },
                CellValueEdit::Remove { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{address}' has no existing cell owner"
                        ))
                    })?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(self.before.edit_blocked(address));
                    }
                    StagedValueEdit::Remove
                },
            };
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
        let expected_actions = effective_action_count(&self.before, &self.staged)?;
        if expected_actions > MAX_BATCH_EDITS {
            return Err(invalid(format!(
                "value-only batch exceeds {MAX_BATCH_EDITS} expanded cell actions"
            )));
        }
        let mut actions = BTreeMap::new();
        for (address, value) in &self.staged {
            append_actions(&self.before, *address, value, &mut actions)?;
        }
        if actions.is_empty() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            self.before.check_execution()?;
            return Ok(Commit::new(self.before, patch, 0));
        }
        let changed = actions.len();
        let output = rewrite(self.before.source_xml(), self.before.sheet_name(), actions)?;
        let snapshot = Snapshot::from_rewritten_source(&self.before, output)?
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
                    &self.before,
                    &snapshot,
                    *address,
                    formula,
                    shared_formula_changed(&self.before, *address, formula)?,
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

    /// Stage one selector-first numeric insertion for an absent coordinate.
    pub fn insert<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        value: Number,
    ) -> Result<()> {
        self.apply_batch([SheetCellValueEdit::insert(selector, address, value)])
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

    /// Stage one selector-first cacheless scalar formula replacement.
    pub fn set_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        formula: Formula,
    ) -> Result<()> {
        self.apply_batch([SheetCellValueEdit::set_formula(selector, address, formula)])
    }

    /// Stage one selector-first shared-formula master replacement.
    pub fn set_shared_formula<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        address: Address,
        formula: Formula,
    ) -> Result<()> {
        self.apply_batch([SheetCellValueEdit::set_shared_formula(
            selector, address, formula,
        )])
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
            let source = snapshot.editable_cell(address);
            let value = match request.edit {
                CellValueEdit::Insert { value, .. } => {
                    snapshot.require_insertable_absence(address)?;
                    let value = Value::Number(value);
                    value.validate_for_write()?;
                    StagedValueEdit::Insert(Content::Value(value))
                },
                CellValueEdit::Set { value, .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    value.validate_for_write()?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(snapshot.edit_blocked(address));
                    }
                    StagedValueEdit::Set(Content::Value(value))
                },
                CellValueEdit::SetFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    snapshot.require_formula_target(address)?;
                    StagedValueEdit::Set(Content::Formula(formula))
                },
                CellValueEdit::SetSharedFormula { formula, .. } => {
                    let _source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    Content::Formula(formula.clone()).validate_for_write()?;
                    snapshot.shared_formula_group(address, MAX_BATCH_EDITS)?;
                    StagedValueEdit::SetSharedFormula(formula)
                },
                CellValueEdit::Clear { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(snapshot.edit_blocked(address));
                    }
                    StagedValueEdit::Clear
                },
                CellValueEdit::Remove { .. } => {
                    let source = source.ok_or_else(|| {
                        invalid(format!(
                            "cell selector '{}!{}' has no existing cell owner",
                            snapshot.sheet_name(),
                            address
                        ))
                    })?;
                    if !matches!(source, Cell::Value(_)) {
                        return Err(snapshot.edit_blocked(address));
                    }
                    StagedValueEdit::Remove
                },
            };
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
        let mut expected_actions = 0usize;
        for snapshot in self.before.sheets() {
            if let Some(staged) = self.staged.get(&snapshot.sheet_position()) {
                expected_actions = expected_actions
                    .checked_add(effective_action_count(snapshot, staged)?)
                    .ok_or_else(|| invalid("expanded value-only action count overflows usize"))?;
                if expected_actions > MAX_BATCH_EDITS {
                    return Err(invalid(format!(
                        "multi-sheet value-only batch exceeds {MAX_BATCH_EDITS} expanded cell actions"
                    )));
                }
            }
        }
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
            let output = rewrite(snapshot.source_xml(), snapshot.sheet_name(), actions)?;
            aggregate_bytes = super::snapshot::checked_multi_bytes(
                aggregate_bytes,
                output.len(),
                super::snapshot::MAX_MULTI_WORKSHEET_BYTES,
            )?;
            let candidate = Snapshot::from_rewritten_source(snapshot, output)?;
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
            let workbook = self.before.sheets()[0].invalidated_workbook_xml()?;
            for snapshot in &mut after {
                *snapshot = snapshot
                    .clone()
                    .with_invalidated_workbook(Arc::clone(&workbook))?;
            }
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
