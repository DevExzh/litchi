//! Tile-local, byte-authoritative BNC rewrites.
//!
//! This module deliberately knows nothing about archives, table selectors, or
//! shared data-list ownership.  Its job is the small and safety-critical part
//! of a table edit: rewrite the BNC rows which are already attached to one
//! tile.  The caller owns string/rich-text/formula reference accounting and
//! sparse attachment.  An absent row in a populated existing tile is reported
//! as [`TileError::NeedSparse`].  A canonical modern tile with no row records
//! is materialised in place, retaining its top-level unknown fields exactly.

use std::mem::size_of;

#[cfg(test)]
use std::cell::Cell;

use litchi_iwa_common::formula::FormulaCachedValue;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;

use crate::cell::{
    FiniteF64,
    wire::{BncCellView, CachedScalar, ClearValue, Error as BncError, ScalarValue, StoredValue},
};

const MISSING_OFFSET: u16 = u16::MAX;
const BNC_STORAGE_VERSION: u32 = 5;
const MINIMAL_BNC_CELL: [u8; 12] = [5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

#[cfg(test)]
thread_local! {
    static PREPARED_EXECUTION_ALLOCATIONS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn reset_prepared_execution_allocations() {
    PREPARED_EXECUTION_ALLOCATIONS.set(0);
}

#[cfg(test)]
fn prepared_execution_allocations() -> usize {
    PREPARED_EXECUTION_ALLOCATIONS.get()
}

#[cfg(test)]
fn record_prepared_execution_allocation() {
    PREPARED_EXECUTION_ALLOCATIONS.set(PREPARED_EXECUTION_ALLOCATIONS.get() + 1);
}

/// One scalar which has already been interned by the transaction planner.
///
/// Text is deliberately an existing table-data-list identifier.  Keeping the
/// authored text out of this layer prevents one tile rewrite from owning a
/// second, accidental string copy.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum ScalarInput {
    String(u32),
    RichText(u32),
    Number(FiniteF64),
    Boolean(bool),
    Date(FiniteF64),
    Duration(FiniteF64),
}

/// A BNC value operation.  `Clear` retains every non-value BNC field (for
/// example comments, format/style references, and the opaque tail).
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum BncChange {
    Set(ScalarInput),
    Clear,
    /// Clear value/reference fields while retaining an explicit empty BNC
    /// slot. Formula removal uses this so the public stored-empty state is
    /// distinct from an absent sparse cell.
    FormulaClear,
    FormulaSet {
        identifier: u32,
        cache: Option<ScalarInput>,
    },
    FormulaCache(CacheScalarInput),
}

/// The cache subset admitted by the formula planner and raw BNC writer.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum CacheScalarInput {
    Number(FiniteF64),
    Boolean(bool),
}

/// One sorted tile-local change.  `row` is the tile's native row index (not a
/// semantic row after any caller-side translation).
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct TileChange {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) change: BncChange,
}

/// One sorted tile-local formula display-cache refresh.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct CacheChange {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) value: FormulaCachedValue,
}

/// Independent limits used before allocating a rewritten tile payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TileLimits {
    pub(crate) max_input_bytes: usize,
    pub(crate) max_output_bytes: usize,
    pub(crate) max_fields: usize,
    pub(crate) max_work: u64,
    pub(crate) max_rows: usize,
    pub(crate) max_cells: usize,
    max_retained_bytes: usize,
    max_retained_elements: usize,
    max_peak_scratch_bytes: usize,
    max_allocations: usize,
}

impl TileLimits {
    pub(crate) const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work: u64,
        max_rows: usize,
        max_cells: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work,
            max_rows,
            max_cells,
            max_retained_bytes: usize::MAX,
            max_retained_elements: usize::MAX,
            max_peak_scratch_bytes: usize::MAX,
            max_allocations: usize::MAX,
        }
    }

    #[must_use]
    pub(crate) const fn with_accounting(
        mut self,
        max_retained_bytes: usize,
        max_retained_elements: usize,
        max_peak_scratch_bytes: usize,
        max_allocations: usize,
    ) -> Self {
        self.max_retained_bytes = max_retained_bytes;
        self.max_retained_elements = max_retained_elements;
        self.max_peak_scratch_bytes = max_peak_scratch_bytes;
        self.max_allocations = max_allocations;
        self
    }
}

/// Borrowed tile input.  Changes must be sorted strictly by `(row, column)`;
/// validating that invariant here lets the hot path group without a map.
#[derive(Debug, Clone, Copy)]
pub(crate) struct TileRewriteRequest<'source, 'change> {
    pub(crate) source: &'source [u8],
    pub(crate) columns: u32,
    pub(crate) changes: &'change [TileChange],
    pub(crate) limits: TileLimits,
}

/// One sorted tile-local coordinate for read-only native classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TileReadPosition {
    pub(crate) row: u32,
    pub(crate) column: u32,
}

/// Value/reference shape observed before or after a local BNC rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellValue {
    Missing,
    Empty,
    Number,
    Text(u32),
    Formula { identifier: u32, error: Option<u32> },
    RichText(u32),
    Date,
    Boolean,
    Duration,
    Error(Option<u32>),
    Unsupported(u8),
}

/// Reference-bearing BNC fields which a higher transaction layer must
/// refcount or reject.  Style/format/comment fields are intentionally not
/// rewritten and remain in the original BNC bytes.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct CellReferences {
    pub(crate) string: Option<u32>,
    pub(crate) rich_text: Option<u32>,
    pub(crate) formula: Option<u32>,
    pub(crate) formula_error: Option<u32>,
    pub(crate) comment: Option<u32>,
}

/// One changed cell and both of its source classifications.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellTransition {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) before: CellValue,
    pub(crate) after: CellValue,
    pub(crate) before_references: CellReferences,
    pub(crate) after_references: CellReferences,
}

/// Exact final populated-cell count for one touched tile-local row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RowCellCount {
    pub(crate) row: u32,
    pub(crate) cell_count: u32,
}

/// Exact source state retained for one requested coordinate.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct PreclassifiedCell {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) before: CellValue,
    pub(crate) before_references: CellReferences,
    pub(crate) formula_cache: Option<FormulaCacheValue>,
    pub(crate) present: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum FormulaCacheValue {
    Number(FiniteF64),
    Boolean(bool),
    Date(FiniteF64),
    Duration(FiniteF64),
    TextKey(u32),
}

/// One formula cell discovered by a complete occupied-slot scan.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct ScannedFormulaCell {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) identifier: u32,
    pub(crate) cache: Option<FormulaCacheValue>,
    pub(crate) formula_error: Option<u32>,
}

#[derive(Debug)]
pub(crate) struct FormulaCellScan {
    pub(crate) cells: Vec<ScannedFormulaCell>,
    pub(crate) report: TileReport,
}

/// Read-only tile result used before list/rich-text identifiers are assigned.
#[derive(Debug)]
pub(crate) struct TilePreclassification {
    pub(crate) cells: Vec<PreclassifiedCell>,
    pub(crate) report: TileReport,
}

/// Exact counters which can be merged into the transaction-wide budget.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct TileReport {
    pub(crate) wire_bytes: u64,
    pub(crate) wire_fields: u64,
    pub(crate) wire_work: u64,
    pub(crate) rows_read: u64,
    pub(crate) rows_written: u64,
    pub(crate) cell_slots_scanned: u64,
    pub(crate) cell_slots_written: u64,
    pub(crate) cache_cells_read: u64,
    pub(crate) cache_cells_written: u64,
    pub(crate) output_bytes: u64,
    pub(crate) retained_elements: u64,
    pub(crate) retained_bytes: u64,
    pub(crate) current_scratch_bytes: u64,
    pub(crate) peak_scratch_bytes: u64,
    pub(crate) allocation_events: u64,
}

/// Output-free observations retained by one prepared tile rewrite.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct TilePrepareReport {
    report: TileReport,
}

impl TilePrepareReport {
    #[must_use]
    pub(crate) const fn report(self) -> TileReport {
        self.report
    }

    #[must_use]
    pub(crate) const fn output_bytes(self) -> usize {
        0
    }
}

/// Exact resources consumed after an output-free tile plan is accepted.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct TileExecutionRequirements {
    input_bytes: usize,
    fields: usize,
    work: u64,
    output_bytes: usize,
    retained_bytes: usize,
    retained_elements: usize,
    peak_scratch_bytes: usize,
    allocations: usize,
    rows_read: usize,
    rows_written: usize,
    cell_slots_scanned: usize,
    cell_slots_written: usize,
    cache_cells_read: usize,
    cache_cells_written: usize,
}

impl TileExecutionRequirements {
    #[must_use]
    pub(crate) const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub(crate) const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub(crate) const fn work(self) -> u64 {
        self.work
    }
    #[must_use]
    pub(crate) const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub(crate) const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub(crate) const fn retained_elements(self) -> usize {
        self.retained_elements
    }
    #[must_use]
    pub(crate) const fn peak_scratch_bytes(self) -> usize {
        self.peak_scratch_bytes
    }
    #[must_use]
    pub(crate) const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub(crate) const fn rows_read(self) -> usize {
        self.rows_read
    }
    #[must_use]
    pub(crate) const fn rows_written(self) -> usize {
        self.rows_written
    }
    #[must_use]
    pub(crate) const fn cell_slots_scanned(self) -> usize {
        self.cell_slots_scanned
    }
    #[must_use]
    pub(crate) const fn cell_slots_written(self) -> usize {
        self.cell_slots_written
    }
    #[must_use]
    pub(crate) const fn cache_cells_read(self) -> usize {
        self.cache_cells_read
    }
    #[must_use]
    pub(crate) const fn cache_cells_written(self) -> usize {
        self.cache_cells_written
    }
    #[must_use]
    pub(crate) const fn exact_limits(self) -> TileExecutionLimits {
        TileExecutionLimits {
            max_input_bytes: self.input_bytes(),
            max_fields: self.fields(),
            max_work: self.work(),
            max_output_bytes: self.output_bytes(),
            max_retained_bytes: self.retained_bytes(),
            max_retained_elements: self.retained_elements(),
            max_peak_scratch_bytes: self.peak_scratch_bytes(),
            max_allocations: self.allocations(),
        }
    }
}

/// Independent execution limits checked before the first mutation/candidate
/// allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TileExecutionLimits {
    pub(crate) max_input_bytes: usize,
    pub(crate) max_fields: usize,
    pub(crate) max_work: u64,
    pub(crate) max_output_bytes: usize,
    pub(crate) max_retained_bytes: usize,
    pub(crate) max_retained_elements: usize,
    pub(crate) max_peak_scratch_bytes: usize,
    pub(crate) max_allocations: usize,
}

/// Tile-local result.  A no-op retains no replacement payload at all.
#[derive(Debug)]
pub(crate) struct TileRewriteOutcome {
    pub(crate) payload: Option<Vec<u8>>,
    pub(crate) transitions: Vec<CellTransition>,
    /// Sorted exact counts for every row named by the request.
    pub(crate) final_rows: Vec<RowCellCount>,
    pub(crate) report: TileReport,
}

#[derive(Clone, Copy)]
enum PreparedSlot<'source> {
    Missing,
    Borrowed(&'source [u8]),
    Mutation {
        previous: Option<&'source [u8]>,
        change: BncChange,
        output_len: Option<usize>,
        transition: Option<CellTransition>,
    },
}

impl<'source> PreparedSlot<'source> {
    const fn borrowed(self) -> Option<&'source [u8]> {
        match self {
            Self::Missing => None,
            Self::Borrowed(bytes) => Some(bytes),
            Self::Mutation { previous, .. } => previous,
        }
    }

    const fn output_len(self) -> Option<usize> {
        match self {
            Self::Missing => None,
            Self::Borrowed(bytes) => Some(bytes.len()),
            Self::Mutation { output_len, .. } => output_len,
        }
    }

    const fn mutation(
        self,
    ) -> Option<(
        Option<&'source [u8]>,
        BncChange,
        Option<usize>,
        Option<CellTransition>,
    )> {
        match self {
            Self::Mutation {
                previous,
                change,
                output_len,
                transition,
            } => Some((previous, change, output_len, transition)),
            Self::Missing | Self::Borrowed(_) => None,
        }
    }
}

#[derive(Clone, Copy)]
enum PreparedRowSource<'source> {
    Borrowed { raw: &'source [u8] },
    Canonical { row: u32 },
}

struct PreparedRow<'source> {
    field_index: Option<usize>,
    row: u32,
    source: PreparedRowSource<'source>,
    slots: Vec<PreparedSlot<'source>>,
    slot_layout: SlotLayout,
    message_layout: RowMessageLayout,
    cell_count: u32,
    changed_slots: usize,
    transition_count: usize,
    cache_count: usize,
    output: Option<Vec<u8>>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PreparedTileMode {
    Existing,
    New,
    PopulatedAppend,
}

/// A complete logical tile plan. It borrows source BNC bytes and retains only
/// compact row/slot descriptors; no cell, row, scaffold, or tile candidate is
/// allocated until [`PreparedTileRewrite::execute`].
pub(crate) struct PreparedTileRewrite<'source> {
    source: &'source [u8],
    columns: u32,
    changes: Vec<TileChange>,
    rows: Vec<PreparedRow<'source>>,
    final_rows: Vec<RowCellCount>,
    mode: PreparedTileMode,
    final_row_count: u32,
    output_len: usize,
    prepare_report: TilePrepareReport,
    requirements: TileExecutionRequirements,
    prepared_retained_bytes: usize,
}

impl PreparedTileRewrite<'_> {
    #[must_use]
    pub(crate) const fn prepare_report(&self) -> TilePrepareReport {
        self.prepare_report
    }

    #[must_use]
    pub(crate) const fn execution_requirements(&self) -> TileExecutionRequirements {
        self.requirements
    }

    /// Exact allocation-free work needed to visit the logical transition
    /// registry once. The caller admits this independently because the visit
    /// happens after the prepared-state report has already been settled.
    pub(crate) fn transition_visit_work(&self) -> Result<usize> {
        self.rows.iter().try_fold(0usize, |total, row| {
            total
                .checked_add(row.slots.len())
                .ok_or(TileError::InvalidSource)
        })
    }

    /// Sorted exact row counts retained by the output-free plan.
    #[must_use]
    pub(crate) fn final_rows(&self) -> &[RowCellCount] {
        &self.final_rows
    }

    /// Visit the exact non-cache cell transitions without allocating a tile,
    /// row, cell, or transition artifact. The parent uses this logical view to
    /// settle shared-list releases and formula authority before entering the
    /// aggregate output lease.
    pub(crate) fn visit_transitions(&self, mut visit: impl FnMut(CellTransition)) -> Result<usize> {
        let mut visited = 0usize;
        for row in &self.rows {
            for (column, slot) in row.slots.iter().copied().enumerate() {
                let Some((_, _, _, transition)) = slot.mutation() else {
                    continue;
                };
                let Some(transition) = transition else {
                    continue;
                };
                if transition.row != row.row
                    || transition.column
                        != u32::try_from(column).map_err(|_| TileError::InvalidSource)?
                {
                    return Err(TileError::InvalidSource);
                }
                visit(transition);
                visited = visited.checked_add(1).ok_or(TileError::InvalidSource)?;
            }
        }
        let expected = self.rows.iter().try_fold(0usize, |total, row| {
            total
                .checked_add(row.transition_count)
                .ok_or(TileError::InvalidSource)
        })?;
        if visited != expected {
            return Err(TileError::InvalidSource);
        }
        Ok(visited)
    }

    pub(crate) fn execute(mut self, limits: TileExecutionLimits) -> Result<TileRewriteOutcome> {
        ensure_tile_execution_limits(self.requirements, limits)?;
        if self.prepare_report.output_bytes() != 0 {
            return Err(TileError::InvalidSource);
        }
        let prepare = self.prepare_report();
        let requirements = self.requirements;
        if self.columns == 0
            || self
                .changes
                .iter()
                .any(|change| change.column >= self.columns)
            || prepare.report().retained_bytes != u64_from_usize(self.prepared_retained_bytes)?
        {
            return Err(TileError::InvalidSource);
        }
        let mut counters = Counters::new(
            TileLimits::new(
                requirements.input_bytes.max(1),
                requirements
                    .peak_scratch_bytes
                    .max(requirements.output_bytes),
                requirements.fields.max(1),
                u64::MAX,
                usize::MAX,
                usize::MAX,
            )
            .with_accounting(
                requirements.retained_bytes,
                requirements.retained_elements,
                requirements.peak_scratch_bytes,
                requirements.allocations,
            ),
        );
        let transition_count = self.rows.iter().try_fold(0usize, |total, row| {
            total
                .checked_add(row.transition_count)
                .ok_or(TileError::InvalidSource)
        })?;
        let transition_bytes = transition_count
            .checked_mul(size_of::<CellTransition>())
            .ok_or(TileError::InvalidSource)?;
        let mut transitions =
            exact_scratch_vec::<CellTransition>(transition_count, transition_bytes, &mut counters)?;
        let final_row_bytes = self
            .final_rows
            .len()
            .checked_mul(size_of::<RowCellCount>())
            .ok_or(TileError::InvalidSource)?;
        let mut final_rows = exact_scratch_vec::<RowCellCount>(
            self.final_rows.len(),
            final_row_bytes,
            &mut counters,
        )?;
        final_rows.extend_from_slice(&self.final_rows);

        for row in &mut self.rows {
            let slots = materialize_prepared_row(row, &mut transitions, &mut counters)?;
            let observed_layout = plan_slot_layout(&slots, row.slot_layout.wide)?;
            if observed_layout.storage_len != row.slot_layout.storage_len
                || observed_layout.storage_capacity != row.slot_layout.storage_capacity
                || observed_layout.offsets_len != row.slot_layout.offsets_len
            {
                return Err(TileError::InvalidSource);
            }
            let (storage, offsets, wide) =
                encode_slots(&slots, row.slot_layout.wide, &mut counters)?;
            if wide != row.slot_layout.wide {
                return Err(TileError::InvalidSource);
            }
            let storage_capacity = storage.capacity();
            let offsets_capacity = offsets.capacity();
            let output = match row.source {
                PreparedRowSource::Borrowed { raw } => write_row_message(
                    raw,
                    row.cell_count,
                    &storage,
                    &offsets,
                    wide,
                    row.message_layout,
                    &mut counters,
                )?,
                PreparedRowSource::Canonical { row: row_index } => write_canonical_row_message(
                    row_index,
                    row.cell_count,
                    &storage,
                    &offsets,
                    wide,
                    row.message_layout,
                    &mut counters,
                )?,
            };
            counters.release_scratch(storage_capacity)?;
            counters.release_scratch(offsets_capacity)?;
            for slot in &slots {
                if let Slot::Owned(bytes) = slot {
                    counters.release_scratch(bytes.capacity())?;
                }
            }
            counters.release_scratch(
                slots
                    .capacity()
                    .checked_mul(size_of::<Slot<'_>>())
                    .ok_or(TileError::InvalidSource)?,
            )?;
            row.output = Some(output);
        }

        let payload = if self.output_len == 0 {
            None
        } else {
            Some(write_prepared_tile(&self, &mut counters)?)
        };
        for row in &self.rows {
            if let Some(output) = &row.output {
                counters.release_scratch(output.capacity())?;
            }
        }
        if let Some(payload) = &payload {
            counters.retain(payload.capacity(), 1)?;
            counters.release_scratch(payload.capacity())?;
        }
        counters.release_scratch(transition_bytes)?;
        counters.retain(transition_bytes, transitions.len())?;
        counters.release_scratch(final_row_bytes)?;
        counters.retain(final_row_bytes, final_rows.len())?;
        if transitions.len() != transition_count
            || payload.as_ref().map_or(0, Vec::len) != self.output_len
        {
            return Err(TileError::InvalidSource);
        }
        counters.report.wire_bytes = u64_from_usize(requirements.input_bytes)?;
        counters.report.wire_fields = u64_from_usize(requirements.fields)?;
        counters.report.wire_work = requirements.work;
        counters.report.rows_read = u64_from_usize(requirements.rows_read())?;
        counters.report.rows_written = u64_from_usize(requirements.rows_written())?;
        counters.report.cell_slots_scanned = u64_from_usize(requirements.cell_slots_scanned())?;
        counters.report.cell_slots_written = u64_from_usize(requirements.cell_slots_written())?;
        counters.report.cache_cells_read = u64_from_usize(requirements.cache_cells_read())?;
        counters.report.cache_cells_written = u64_from_usize(requirements.cache_cells_written())?;
        counters.report.output_bytes = u64_from_usize(requirements.output_bytes)?;
        counters.report.peak_scratch_bytes = u64_from_usize(requirements.peak_scratch_bytes)?;
        counters.report.current_scratch_bytes = 0;
        if counters.report.allocation_events != u64_from_usize(requirements.allocations)?
            || counters.report.retained_bytes != u64_from_usize(requirements.retained_bytes)?
            || counters.report.retained_elements != u64_from_usize(requirements.retained_elements)?
        {
            return Err(TileError::InvalidSource);
        }
        Ok(TileRewriteOutcome {
            payload,
            transitions,
            final_rows,
            report: counters.report,
        })
    }
}

fn ensure_tile_execution_limits(
    requirements: TileExecutionRequirements,
    limits: TileExecutionLimits,
) -> Result<()> {
    for (observed, maximum) in [
        (requirements.input_bytes, limits.max_input_bytes),
        (requirements.fields, limits.max_fields),
        (requirements.output_bytes, limits.max_output_bytes),
        (requirements.retained_bytes, limits.max_retained_bytes),
        (requirements.retained_elements, limits.max_retained_elements),
        (
            requirements.peak_scratch_bytes,
            limits.max_peak_scratch_bytes,
        ),
        (requirements.allocations, limits.max_allocations),
    ] {
        ensure_usize_limit(observed, maximum)?;
    }
    if requirements.work > limits.max_work {
        return Err(TileError::LimitExceeded {
            observed: requirements.work,
            maximum: limits.max_work,
        });
    }
    Ok(())
}

fn exact_scratch_vec<T>(amount: usize, bytes: usize, counters: &mut Counters) -> Result<Vec<T>> {
    if amount == 0 {
        return Ok(Vec::new());
    }
    counters.preflight_scratch_allocation(bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_| TileError::Allocation { amount })?;
    if output.capacity() != amount {
        return Err(TileError::Allocation { amount });
    }
    #[cfg(test)]
    record_prepared_execution_allocation();
    counters.record_scratch_allocation(bytes)?;
    Ok(output)
}

/// A typed, transaction-private tile failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TileError {
    InvalidSource,
    UnsupportedSource {
        row: u32,
        column: u32,
    },
    UnsupportedValue {
        row: u32,
        column: u32,
        kind: u8,
    },
    DuplicateOrUnsortedChange {
        row: u32,
        column: u32,
    },
    OutOfBounds {
        row: u32,
        column: u32,
    },
    /// The caller must allocate/attach a sparse tile or row first.
    NeedSparse {
        row: u32,
        column: u32,
    },
    LimitExceeded {
        observed: u64,
        maximum: u64,
    },
    Allocation {
        amount: usize,
    },
}

type Result<T> = std::result::Result<T, TileError>;

/// Strictly classify selected cells without producing a replacement payload.
///
/// Positions must be strictly sorted by `(row, column)`. The result remains
/// in that same order and contains only borrowed-source classifications copied
/// into compact semantic records. Rows and offset tables are streamed, so the
/// only retained allocation is the final cell vector.
pub(crate) fn preclassify_tile(
    source: &[u8],
    columns: u32,
    positions: &[TileReadPosition],
    limits: TileLimits,
) -> Result<TilePreclassification> {
    if positions.is_empty() {
        return Ok(TilePreclassification {
            cells: Vec::new(),
            report: TileReport::default(),
        });
    }
    validate_read_positions(source, columns, positions, limits)?;
    let mut counters = Counters::new(limits);
    counters.charge_bytes(source.len())?;
    let codec_options = storage::DecodeOptions::new(
        limits.max_input_bytes,
        limits.max_fields,
        usize_from_u64(limits.max_work)?,
        64,
        limits.max_cells,
        limits.max_output_bytes,
    );
    let (tile, codec_report) = storage::decode_tile_with_report(source, codec_options)
        .map_err(|_| TileError::InvalidSource)?;
    counters.report.wire_bytes = u64_from_usize(codec_report.source_bytes())?;
    counters.report.wire_fields = u64_from_usize(codec_report.fields())?;
    counters.charge(u64_from_usize(codec_report.work_bytes())?)?;

    if tile.storage_version() == Some(BNC_STORAGE_VERSION)
        && tile.last_saved_in_bnc() == Some(true)
        && tile_has_no_rows(source)?
    {
        return preclassify_empty_modern_tile(positions, &mut counters);
    }

    let retained_bytes = positions
        .len()
        .checked_mul(size_of::<PreclassifiedCell>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(retained_bytes)?;
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(positions.len())
        .map_err(|_| TileError::Allocation {
            amount: positions.len(),
        })?;
    let retained_capacity = cells
        .capacity()
        .checked_mul(size_of::<PreclassifiedCell>())
        .ok_or(TileError::InvalidSource)?;
    counters.record_scratch_allocation(retained_capacity)?;

    let tile_view = WireView::parse(source).map_err(|_| TileError::InvalidSource)?;
    let mut position_start = 0usize;
    let mut previous_row = None;
    let mut row_records = 0usize;
    for field in tile_view.fields() {
        counters.charge(1)?;
        if field.number() != 5 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(TileError::InvalidSource);
        }
        row_records = row_records.checked_add(1).ok_or(TileError::InvalidSource)?;
        if row_records > limits.max_rows {
            return Err(TileError::LimitExceeded {
                observed: u64_from_usize(row_records)?,
                maximum: u64_from_usize(limits.max_rows)?,
            });
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| TileError::InvalidSource)?;
        let (row, report) = storage::decode_tile_row_info_with_report(payload, codec_options)
            .map_err(|_| TileError::InvalidSource)?;
        counters.charge(u64_from_usize(report.work_bytes())?)?;
        counters.report.rows_read = counters
            .report
            .rows_read
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        let row_index = row.tile_row_index();
        if usize::try_from(row_index).map_or(true, |row| row >= limits.max_rows)
            || previous_row.is_some_and(|previous| previous >= row_index)
        {
            return Err(TileError::InvalidSource);
        }
        previous_row = Some(row_index);
        if positions
            .get(position_start)
            .is_some_and(|position| position.row < row_index)
        {
            let position = positions[position_start];
            return Err(TileError::NeedSparse {
                row: position.row,
                column: position.column,
            });
        }
        let position_end = positions[position_start..]
            .iter()
            .position(|position| position.row != row_index)
            .map_or(positions.len(), |offset| position_start + offset);
        if position_start != position_end {
            preclassify_row(
                row,
                columns,
                &positions[position_start..position_end],
                &mut cells,
                &mut counters,
            )?;
        }
        position_start = position_end;
    }
    if let Some(position) = positions.get(position_start) {
        return Err(TileError::NeedSparse {
            row: position.row,
            column: position.column,
        });
    }
    if cells.len() != positions.len() {
        return Err(TileError::InvalidSource);
    }
    counters.release_scratch(retained_capacity)?;
    counters.retain(retained_capacity, cells.len())?;
    Ok(TilePreclassification {
        cells,
        report: counters.report,
    })
}

/// Strictly scan every occupied BNC slot and retain every formula host.
pub(crate) fn scan_formula_cells(
    source: &[u8],
    columns: u32,
    limits: TileLimits,
) -> Result<FormulaCellScan> {
    let visitor_work_upper = source
        .len()
        .checked_mul(2)
        .ok_or(TileError::InvalidSource)?;
    let both_visitors_upper = visitor_work_upper
        .checked_mul(2)
        .ok_or(TileError::InvalidSource)?;
    let codec_work_total = limits
        .max_work
        .checked_sub(u64_from_usize(both_visitors_upper)?)
        .ok_or(TileError::LimitExceeded {
            observed: u64_from_usize(both_visitors_upper)?,
            maximum: limits.max_work,
        })?;
    let first_codec_work = codec_work_total / 2;
    let first_bytes_limit = limits.max_input_bytes / 2;
    let first_fields_limit = limits.max_fields / 2;
    if source.len() > first_bytes_limit {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(
                source
                    .len()
                    .checked_mul(2)
                    .ok_or(TileError::InvalidSource)?,
            )?,
            maximum: u64_from_usize(limits.max_input_bytes)?,
        });
    }
    let first_options = storage::DecodeOptions::new(
        first_bytes_limit,
        first_fields_limit,
        usize_from_u64(first_codec_work)?,
        64,
        limits.max_cells,
        limits.max_output_bytes,
    );
    let mut count = FormulaScanVisitor::new(columns, limits.max_rows, None);
    let (tile, first) = storage::decode_tile_with_visitor(source, first_options, &mut count)
        .map_err(|_| TileError::InvalidSource)?;
    count.ensure_complete()?;
    if tile.storage_version() != Some(BNC_STORAGE_VERSION) || tile.last_saved_in_bnc() != Some(true)
    {
        return Err(TileError::InvalidSource);
    }
    if count.formulas > limits.max_cells {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(count.formulas)?,
            maximum: u64_from_usize(limits.max_cells)?,
        });
    }
    let retained_bytes = count
        .formulas
        .checked_mul(size_of::<ScannedFormulaCell>())
        .ok_or(TileError::InvalidSource)?;
    if retained_bytes > limits.max_output_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(retained_bytes)?,
            maximum: u64_from_usize(limits.max_output_bytes)?,
        });
    }
    let first_bytes = first.source_bytes();
    let first_fields = first.fields();
    let first_work = first
        .work_bytes()
        .checked_add(usize::try_from(count.work).map_err(|_| TileError::InvalidSource)?)
        .ok_or(TileError::InvalidSource)?;
    let remaining_bytes = limits
        .max_input_bytes
        .checked_sub(first_bytes)
        .ok_or(TileError::InvalidSource)?;
    let remaining_fields = limits
        .max_fields
        .checked_sub(first_fields)
        .ok_or(TileError::InvalidSource)?;
    let remaining_work = codec_work_total
        .checked_sub(u64_from_usize(first.work_bytes())?)
        .ok_or(TileError::InvalidSource)?;
    if remaining_bytes < first_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(first_bytes.checked_mul(2).ok_or(TileError::InvalidSource)?)?,
            maximum: u64_from_usize(limits.max_input_bytes)?,
        });
    }
    if remaining_fields < first_fields {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(
                first_fields
                    .checked_mul(2)
                    .ok_or(TileError::InvalidSource)?,
            )?,
            maximum: u64_from_usize(limits.max_fields)?,
        });
    }
    if remaining_work < u64_from_usize(first.work_bytes())? {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(
                first
                    .work_bytes()
                    .checked_mul(2)
                    .ok_or(TileError::InvalidSource)?,
            )?,
            maximum: limits.max_work,
        });
    }
    let second_options = storage::DecodeOptions::new(
        remaining_bytes,
        remaining_fields,
        usize_from_u64(remaining_work)?,
        64,
        limits.max_cells,
        limits.max_output_bytes,
    );
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(count.formulas)
        .map_err(|_| TileError::Allocation {
            amount: count.formulas,
        })?;
    if cells.capacity() != count.formulas {
        return Err(TileError::Allocation {
            amount: count.formulas,
        });
    }
    let (second, collected_work, collected_rows, collected_slots) = {
        let mut collect = FormulaScanVisitor::new(columns, limits.max_rows, Some(&mut cells));
        let (_, second) = storage::decode_tile_with_visitor(source, second_options, &mut collect)
            .map_err(|_| TileError::InvalidSource)?;
        collect.ensure_complete()?;
        (second, collect.work, collect.rows, collect.slots)
    };
    if cells.len() != count.formulas
        || cells
            .windows(2)
            .any(|pair| (pair[0].row, pair[0].column) >= (pair[1].row, pair[1].column))
    {
        return Err(TileError::InvalidSource);
    }
    let wire_bytes = u64_from_usize(first.source_bytes())?
        .checked_add(u64_from_usize(second.source_bytes())?)
        .ok_or(TileError::InvalidSource)?;
    let wire_fields = u64_from_usize(first.fields())?
        .checked_add(u64_from_usize(second.fields())?)
        .ok_or(TileError::InvalidSource)?;
    let wire_work = u64_from_usize(first_work)?
        .checked_add(u64_from_usize(second.work_bytes())?)
        .and_then(|work| work.checked_add(collected_work))
        .ok_or(TileError::InvalidSource)?;
    if wire_work > limits.max_work {
        return Err(TileError::LimitExceeded {
            observed: wire_work,
            maximum: limits.max_work,
        });
    }
    Ok(FormulaCellScan {
        cells,
        report: TileReport {
            wire_bytes,
            wire_fields,
            wire_work,
            rows_read: count
                .rows
                .checked_add(collected_rows)
                .ok_or(TileError::InvalidSource)?,
            cell_slots_scanned: count
                .slots
                .checked_add(collected_slots)
                .ok_or(TileError::InvalidSource)?,
            cache_cells_read: u64_from_usize(count.formulas)?,
            retained_elements: u64_from_usize(count.formulas)?,
            retained_bytes: u64_from_usize(retained_bytes)?,
            allocation_events: u64::from(count.formulas != 0),
            ..TileReport::default()
        },
    })
}

struct FormulaScanVisitor<'cells> {
    columns: u32,
    max_rows: usize,
    cells: Option<&'cells mut Vec<ScannedFormulaCell>>,
    formulas: usize,
    rows: u64,
    slots: u64,
    work: u64,
    failed: bool,
    previous_row: Option<u32>,
}

impl<'cells> FormulaScanVisitor<'cells> {
    const fn new(
        columns: u32,
        max_rows: usize,
        cells: Option<&'cells mut Vec<ScannedFormulaCell>>,
    ) -> Self {
        Self {
            columns,
            max_rows,
            cells,
            formulas: 0,
            rows: 0,
            slots: 0,
            work: 0,
            failed: false,
            previous_row: None,
        }
    }

    fn ensure_complete(&self) -> Result<()> {
        if self.failed {
            Err(TileError::InvalidSource)
        } else {
            Ok(())
        }
    }
}

impl storage::StorageVisitor for FormulaScanVisitor<'_> {
    fn visit_tile_row(
        &mut self,
        row: storage::TileRowInfoSnapshot<'_>,
    ) -> core::result::Result<(), storage::DecodeError> {
        if let Err(_error) = scan_formula_row(row, self.columns, self) {
            self.failed = true;
        }
        Ok(())
    }
}

fn scan_formula_row(
    row: storage::TileRowInfoSnapshot<'_>,
    columns: u32,
    visitor: &mut FormulaScanVisitor<'_>,
) -> Result<()> {
    let row_index = row.tile_row_index();
    if usize::try_from(row_index).map_or(true, |value| value >= visitor.max_rows)
        || visitor
            .previous_row
            .is_some_and(|previous| previous >= row_index)
    {
        return Err(TileError::InvalidSource);
    }
    visitor.previous_row = Some(row_index);
    let storage_buffer = row.cell_storage_buffer().ok_or(TileError::InvalidSource)?;
    let offsets = row.cell_offsets().ok_or(TileError::InvalidSource)?;
    if row.storage_version() != Some(BNC_STORAGE_VERSION) || !offsets.len().is_multiple_of(2) {
        return Err(TileError::InvalidSource);
    }
    let column_count = usize::try_from(columns).map_err(|_| TileError::InvalidSource)?;
    visitor.rows = visitor
        .rows
        .checked_add(1)
        .ok_or(TileError::InvalidSource)?;
    let unit = if row.has_wide_offsets().unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let mut occupied = 0usize;
    let mut previous: Option<(usize, usize)> = None;
    for (column, encoded) in offsets.chunks_exact(2).enumerate() {
        visitor.slots = visitor
            .slots
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        visitor.work = visitor
            .work
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        if raw == MISSING_OFFSET {
            continue;
        }
        if column >= column_count {
            return Err(TileError::InvalidSource);
        }
        let start = usize::from(raw)
            .checked_mul(unit)
            .ok_or(TileError::InvalidSource)?;
        if start >= storage_buffer.len() || previous.is_some_and(|(_, prior)| prior >= start) {
            return Err(TileError::InvalidSource);
        }
        if let Some((prior_column, prior_start)) = previous {
            scan_formula_cell(
                row.tile_row_index(),
                prior_column,
                &storage_buffer[prior_start..start],
                visitor,
            )?;
        }
        occupied = occupied.checked_add(1).ok_or(TileError::InvalidSource)?;
        previous = Some((column, start));
    }
    if let Some((column, start)) = previous {
        scan_formula_cell(
            row.tile_row_index(),
            column,
            &storage_buffer[start..],
            visitor,
        )?;
    }
    if occupied != usize::try_from(row.cell_count()).map_err(|_| TileError::InvalidSource)? {
        return Err(TileError::InvalidSource);
    }
    Ok(())
}

fn scan_formula_cell(
    row: u32,
    column: usize,
    bytes: &[u8],
    visitor: &mut FormulaScanVisitor<'_>,
) -> Result<()> {
    visitor.work = visitor
        .work
        .checked_add(u64_from_usize(bytes.len())?)
        .ok_or(TileError::InvalidSource)?;
    let view = BncCellView::parse(bytes).map_err(|_| TileError::InvalidSource)?;
    let StoredValue::Formula(identifier) = view.stored_value() else {
        return Ok(());
    };
    let cache = match view.cached_scalar() {
        Some(CachedScalar::Number(value)) => Some(FormulaCacheValue::Number(value)),
        Some(CachedScalar::Boolean(value)) => Some(FormulaCacheValue::Boolean(value)),
        Some(CachedScalar::Date(value)) => Some(FormulaCacheValue::Date(value)),
        Some(CachedScalar::Duration(value)) => Some(FormulaCacheValue::Duration(value)),
        Some(CachedScalar::Unsupported(_)) | None => {
            view.formula_text_key().map(FormulaCacheValue::TextKey)
        },
    };
    visitor.formulas = visitor
        .formulas
        .checked_add(1)
        .ok_or(TileError::InvalidSource)?;
    if let Some(cells) = visitor.cells.as_deref_mut() {
        cells.push(ScannedFormulaCell {
            row,
            column: u32::try_from(column).map_err(|_| TileError::InvalidSource)?,
            identifier,
            cache,
            formula_error: view.formula_error_identifier(),
        });
    }
    Ok(())
}

fn preclassify_empty_modern_tile(
    positions: &[TileReadPosition],
    counters: &mut Counters,
) -> Result<TilePreclassification> {
    let retained_bytes = positions
        .len()
        .checked_mul(size_of::<PreclassifiedCell>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(retained_bytes)?;
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(positions.len())
        .map_err(|_| TileError::Allocation {
            amount: positions.len(),
        })?;
    let retained_capacity = cells
        .capacity()
        .checked_mul(size_of::<PreclassifiedCell>())
        .ok_or(TileError::InvalidSource)?;
    counters.record_scratch_allocation(retained_capacity)?;
    for position in positions {
        cells.push(PreclassifiedCell {
            row: position.row,
            column: position.column,
            before: CellValue::Missing,
            before_references: CellReferences::default(),
            formula_cache: None,
            present: false,
        });
    }
    counters.release_scratch(retained_capacity)?;
    counters.retain(retained_capacity, cells.len())?;
    Ok(TilePreclassification {
        cells,
        report: counters.report,
    })
}

fn preclassify_row(
    row: storage::TileRowInfoSnapshot<'_>,
    columns: u32,
    positions: &[TileReadPosition],
    cells: &mut Vec<PreclassifiedCell>,
    counters: &mut Counters,
) -> Result<()> {
    let storage_buffer = row.cell_storage_buffer().ok_or_else(|| {
        let position = positions[0];
        TileError::NeedSparse {
            row: position.row,
            column: position.column,
        }
    })?;
    let offsets = row.cell_offsets().ok_or_else(|| {
        let position = positions[0];
        TileError::NeedSparse {
            row: position.row,
            column: position.column,
        }
    })?;
    if row.storage_version() != Some(BNC_STORAGE_VERSION) || !offsets.len().is_multiple_of(2) {
        return Err(TileError::InvalidSource);
    }
    let requested_columns = usize::try_from(columns).map_err(|_| TileError::InvalidSource)?;
    let slot_count = offsets.len() / 2;
    if slot_count < requested_columns
        || usize::try_from(row.cell_count()).map_err(|_| TileError::InvalidSource)? > slot_count
    {
        return Err(TileError::InvalidSource);
    }

    let unit = if row.has_wide_offsets().unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let mut occupied = 0usize;
    let mut position_index = 0usize;
    let mut previous: Option<(usize, usize)> = None;
    for (column, encoded) in offsets.chunks_exact(2).enumerate() {
        counters.report.cell_slots_scanned = counters
            .report
            .cell_slots_scanned
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        counters.charge(1)?;
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        if raw == MISSING_OFFSET {
            continue;
        }
        occupied = occupied.checked_add(1).ok_or(TileError::InvalidSource)?;
        let start = usize::from(raw)
            .checked_mul(unit)
            .ok_or(TileError::InvalidSource)?;
        if start >= storage_buffer.len()
            || previous.is_some_and(|(_, previous_start)| previous_start >= start)
        {
            return Err(TileError::InvalidSource);
        }
        while positions.get(position_index).is_some_and(|position| {
            usize::try_from(position.column).is_ok_and(|value| value < column)
        }) {
            let position = positions[position_index];
            let bytes = previous
                .filter(|(previous_column, _)| {
                    u32::try_from(*previous_column).ok() == Some(position.column)
                })
                .map(|(_, previous_start)| &storage_buffer[previous_start..start]);
            push_preclassified(cells, position, bytes, counters)?;
            position_index += 1;
        }
        previous = Some((column, start));
    }
    if occupied != usize::try_from(row.cell_count()).map_err(|_| TileError::InvalidSource)? {
        return Err(TileError::InvalidSource);
    }
    while let Some(&position) = positions.get(position_index) {
        let bytes = previous
            .filter(|(previous_column, _)| {
                u32::try_from(*previous_column).ok() == Some(position.column)
            })
            .map(|(_, previous_start)| &storage_buffer[previous_start..]);
        push_preclassified(cells, position, bytes, counters)?;
        position_index += 1;
    }
    Ok(())
}

fn push_preclassified(
    cells: &mut Vec<PreclassifiedCell>,
    position: TileReadPosition,
    bytes: Option<&[u8]>,
    counters: &mut Counters,
) -> Result<()> {
    let (before, before_references, formula_cache) = classify_bnc_with_references(bytes)?;
    if let Some(bytes) = bytes {
        counters.charge(u64_from_usize(bytes.len())?)?;
    }
    cells.push(PreclassifiedCell {
        row: position.row,
        column: position.column,
        before,
        before_references,
        formula_cache,
        present: bytes.is_some(),
    });
    Ok(())
}

/// Rewrite all selected BNC cells in one existing tile.
///
/// Untouched protobuf fields, untouched row payloads, BNC metadata fields,
/// and BNC opaque tails are copied byte-for-byte.  The strict codec is run
/// before the raw source is trusted; raw fields remain authoritative when the
/// replacement is assembled.
pub(crate) fn prepare_tile<'source>(
    request: TileRewriteRequest<'source, '_>,
) -> Result<PreparedTileRewrite<'source>> {
    prepare_tile_internal(request, &[], false)
}

/// Prepare scalar and formula-cache changes as one output-free logical tile
/// rewrite.
pub(crate) fn prepare_tile_with_cache<'source>(
    request: TileRewriteRequest<'source, '_>,
    cache_changes: &[CacheChange],
) -> Result<PreparedTileRewrite<'source>> {
    prepare_tile_internal(request, cache_changes, false)
}

pub(crate) fn prepare_new_tile<'source>(
    template: &'source [u8],
    columns: u32,
    changes: &[TileChange],
    limits: TileLimits,
) -> Result<PreparedTileRewrite<'source>> {
    prepare_tile_internal(
        TileRewriteRequest {
            source: template,
            columns,
            changes,
            limits,
        },
        &[],
        true,
    )
}

fn prepare_tile_internal<'source>(
    request: TileRewriteRequest<'source, '_>,
    cache_changes: &[CacheChange],
    force_new: bool,
) -> Result<PreparedTileRewrite<'source>> {
    validate_request(request)?;
    validate_cache_changes(cache_changes, request.columns, request.limits)?;
    validate_disjoint_changes(request.changes, cache_changes)?;
    let merged_len = request
        .changes
        .len()
        .checked_add(cache_changes.len())
        .ok_or(TileError::InvalidSource)?;
    if merged_len > request.limits.max_cells {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(merged_len)?,
            maximum: u64_from_usize(request.limits.max_cells)?,
        });
    }
    let (distinct_rows, shape_scan_work) =
        preflight_prepare_shape(request, cache_changes, merged_len)?;
    if merged_len == 0 {
        return Ok(PreparedTileRewrite {
            source: request.source,
            columns: request.columns,
            changes: Vec::new(),
            rows: Vec::new(),
            final_rows: Vec::new(),
            mode: PreparedTileMode::Existing,
            final_row_count: 0,
            output_len: 0,
            prepare_report: TilePrepareReport::default(),
            requirements: TileExecutionRequirements::default(),
            prepared_retained_bytes: 0,
        });
    }

    let mut counters = Counters::new(request.limits);
    counters.charge(shape_scan_work)?;
    let merged_bytes = merged_len
        .checked_mul(size_of::<TileChange>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(merged_bytes)?;
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(merged_len)
        .map_err(|_| TileError::Allocation { amount: merged_len })?;
    if changes.capacity() != merged_len {
        return Err(TileError::Allocation { amount: merged_len });
    }
    counters.record_scratch_allocation(merged_bytes)?;
    merge_tile_changes(request.changes, cache_changes, &mut changes)?;
    if changes.len() != merged_len {
        return Err(TileError::InvalidSource);
    }
    counters.charge(u64_from_usize(merged_len)?)?;

    counters.charge_bytes(request.source.len())?;
    let codec_options = storage::DecodeOptions::new(
        request.limits.max_input_bytes,
        request.limits.max_fields,
        usize_from_u64(request.limits.max_work)?,
        64,
        request.limits.max_cells,
        request.limits.max_output_bytes,
    );
    let (snapshot, codec_report) = storage::decode_tile_with_report(request.source, codec_options)
        .map_err(|_| TileError::InvalidSource)?;
    counters.report.wire_bytes = u64_from_usize(codec_report.source_bytes())?;
    counters.report.wire_fields = u64_from_usize(codec_report.fields())?;
    counters.charge(u64_from_usize(codec_report.work_bytes())?)?;

    let source_view = WireView::parse(request.source).map_err(|_| TileError::InvalidSource)?;
    let row_capacity = usize::try_from(snapshot.num_rows())
        .map_err(|_| TileError::InvalidSource)?
        .min(request.limits.max_rows);
    let raw_row_bytes = row_capacity
        .checked_mul(size_of::<RawRow<'source>>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(raw_row_bytes)?;
    let mut raw_rows = Vec::new();
    raw_rows
        .try_reserve_exact(row_capacity)
        .map_err(|_| TileError::Allocation {
            amount: row_capacity,
        })?;
    if raw_rows.capacity() != row_capacity {
        return Err(TileError::Allocation {
            amount: row_capacity,
        });
    }
    counters.record_scratch_allocation(raw_row_bytes)?;
    let mut previous_row = None;
    for (field_index, field) in source_view.fields().enumerate() {
        counters.charge(1)?;
        if field.number() != 5 {
            continue;
        }
        if field.wire_type() != 2
            || raw_rows.len() == request.limits.max_rows
            || raw_rows.len() == raw_rows.capacity()
        {
            return Err(TileError::InvalidSource);
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| TileError::InvalidSource)?;
        let (row, report) = storage::decode_tile_row_info_with_report(payload, codec_options)
            .map_err(|_| TileError::InvalidSource)?;
        counters.charge(u64_from_usize(report.work_bytes())?)?;
        let row_index = row.tile_row_index();
        if usize::try_from(row_index).map_or(true, |value| value >= request.limits.max_rows)
            || previous_row.is_some_and(|previous| previous >= row_index)
        {
            return Err(TileError::InvalidSource);
        }
        previous_row = Some(row_index);
        counters.report.rows_read = counters
            .report
            .rows_read
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        raw_rows.push(RawRow {
            field_index,
            row,
            payload,
        });
    }

    let modern = snapshot.storage_version() == Some(BNC_STORAGE_VERSION)
        && snapshot.last_saved_in_bnc() == Some(true);
    let mode = if force_new {
        if !modern {
            return Err(TileError::InvalidSource);
        }
        PreparedTileMode::New
    } else {
        classify_prepared_mode(&changes, &raw_rows, modern, &mut counters)?
    };
    let row_plan_capacity = distinct_rows;
    let row_plan_bytes = row_plan_capacity
        .checked_mul(size_of::<PreparedRow<'source>>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(row_plan_bytes)?;
    let mut rows = Vec::new();
    rows.try_reserve_exact(row_plan_capacity)
        .map_err(|_| TileError::Allocation {
            amount: row_plan_capacity,
        })?;
    if rows.capacity() != row_plan_capacity {
        return Err(TileError::Allocation {
            amount: row_plan_capacity,
        });
    }
    counters.record_scratch_allocation(row_plan_bytes)?;

    let final_row_bytes = distinct_rows
        .checked_mul(size_of::<RowCellCount>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(final_row_bytes)?;
    let mut final_rows = Vec::new();
    final_rows
        .try_reserve_exact(distinct_rows)
        .map_err(|_| TileError::Allocation {
            amount: distinct_rows,
        })?;
    if final_rows.capacity() != distinct_rows {
        return Err(TileError::Allocation {
            amount: distinct_rows,
        });
    }
    counters.record_scratch_allocation(final_row_bytes)?;

    let mut start = 0usize;
    while start < changes.len() {
        counters.charge(1)?;
        let row_index = changes[start].row;
        let end = changes[start..]
            .partition_point(|change| change.row == row_index)
            .checked_add(start)
            .ok_or(TileError::InvalidSource)?;
        let row_changes = &changes[start..end];
        counters.charge(u64_from_usize(binary_search_work(raw_rows.len()))?)?;
        let existing = raw_rows.binary_search_by_key(&row_index, |row| row.row.tile_row_index());
        let plan = match existing {
            _ if mode == PreparedTileMode::New => Some(prepare_canonical_row(
                row_index,
                request.columns,
                row_changes,
                &mut counters,
            )?),
            Ok(index) => {
                let raw = raw_rows[index];
                Some(prepare_borrowed_row(
                    raw.row,
                    raw.payload,
                    raw.field_index,
                    request.columns,
                    row_changes,
                    &mut counters,
                )?)
            },
            Err(_) if mode == PreparedTileMode::Existing => {
                let change = row_changes[0];
                return Err(TileError::NeedSparse {
                    row: change.row,
                    column: change.column,
                });
            },
            Err(_) => Some(prepare_canonical_row(
                row_index,
                request.columns,
                row_changes,
                &mut counters,
            )?),
        };
        let cell_count = plan.as_ref().map_or(0, |row| row.cell_count);
        final_rows.push(RowCellCount {
            row: row_index,
            cell_count,
        });
        if let Some(plan) = plan.filter(|plan| plan.changed_slots != 0) {
            rows.push(plan);
        }
        start = end;
    }

    let source_rows = raw_rows.len();
    let appended_rows = rows.iter().filter(|row| row.field_index.is_none()).count();
    let final_row_count = match mode {
        PreparedTileMode::Existing => source_rows,
        PreparedTileMode::New => appended_rows,
        PreparedTileMode::PopulatedAppend => source_rows
            .checked_add(appended_rows)
            .ok_or(TileError::InvalidSource)?,
    };
    if final_row_count > request.limits.max_rows {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(final_row_count)?,
            maximum: u64_from_usize(request.limits.max_rows)?,
        });
    }
    let final_row_count = u32::try_from(final_row_count).map_err(|_| TileError::InvalidSource)?;
    let output_len =
        prepared_tile_output_len(request.source, &rows, mode, final_row_count, &mut counters)?;
    if output_len > request.limits.max_output_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(output_len)?,
            maximum: u64_from_usize(request.limits.max_output_bytes)?,
        });
    }

    counters.release_scratch(raw_row_bytes)?;
    counters.release_scratch(merged_bytes)?;
    counters.release_scratch(row_plan_bytes)?;
    counters.release_scratch(final_row_bytes)?;
    counters.retain(merged_bytes, changes.len())?;
    counters.retain(row_plan_bytes, rows.len())?;
    counters.retain(final_row_bytes, final_rows.len())?;
    let prepared_retained_bytes = usize_from_u64(counters.report.retained_bytes)?;
    let requirements = tile_execution_requirements(
        request.source,
        &rows,
        &final_rows,
        output_len,
        prepared_retained_bytes,
    )?;
    let prepare_report = TilePrepareReport {
        report: counters.report,
    };
    if prepare_report.output_bytes() != 0 || prepare_report.report.current_scratch_bytes != 0 {
        return Err(TileError::InvalidSource);
    }
    Ok(PreparedTileRewrite {
        source: request.source,
        columns: request.columns,
        changes,
        rows,
        final_rows,
        mode,
        final_row_count,
        output_len,
        prepare_report,
        requirements,
        prepared_retained_bytes,
    })
}

fn preflight_prepare_shape(
    request: TileRewriteRequest<'_, '_>,
    cache_changes: &[CacheChange],
    changes: usize,
) -> Result<(usize, u64)> {
    // The exact merged row count is obtainable with one allocation-free
    // merge scan. Admit its complete linear work before entering the scan;
    // treating every cache cell as a distinct row would otherwise create a
    // false O(cells * columns) retained bound for dense formula fanout.
    let scan_work = request
        .changes
        .len()
        .checked_add(cache_changes.len())
        .and_then(|value| value.checked_add(1))
        .ok_or(TileError::InvalidSource)?;
    if u64_from_usize(scan_work)? > request.limits.max_work {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(scan_work)?,
            maximum: request.limits.max_work,
        });
    }
    let distinct = merged_distinct_change_rows(request.changes, cache_changes)?;
    let columns = usize::try_from(request.columns).map_err(|_| TileError::InvalidSource)?;
    let slot_upper = distinct
        .checked_mul(columns)
        .ok_or(TileError::InvalidSource)?;
    let retained_bytes = changes
        .checked_mul(size_of::<TileChange>())
        .and_then(|bytes| {
            distinct
                .checked_mul(size_of::<PreparedRow<'_>>() + size_of::<RowCellCount>())
                .and_then(|value| bytes.checked_add(value))
        })
        .and_then(|bytes| {
            slot_upper
                .checked_mul(size_of::<PreparedSlot<'_>>())
                .and_then(|value| bytes.checked_add(value))
        })
        .ok_or(TileError::InvalidSource)?;
    let retained_elements = changes
        .checked_add(distinct)
        .and_then(|value| value.checked_add(distinct))
        .and_then(|value| value.checked_add(slot_upper))
        .ok_or(TileError::InvalidSource)?;
    let allocations = distinct
        .checked_mul(3)
        .and_then(|value| value.checked_add(4))
        .ok_or(TileError::InvalidSource)?;
    ensure_usize_limit(retained_bytes, request.limits.max_retained_bytes)?;
    ensure_usize_limit(retained_elements, request.limits.max_retained_elements)?;
    ensure_usize_limit(allocations, request.limits.max_allocations)?;
    ensure_usize_limit(retained_bytes, request.limits.max_peak_scratch_bytes)?;
    Ok((distinct, u64_from_usize(scan_work)?))
}

fn merged_distinct_change_rows(
    scalar_changes: &[TileChange],
    cache_changes: &[CacheChange],
) -> Result<usize> {
    let mut scalar = 0usize;
    let mut cache = 0usize;
    let mut previous_row = None;
    let mut rows = 0usize;
    while scalar < scalar_changes.len() || cache < cache_changes.len() {
        let scalar_key = scalar_changes
            .get(scalar)
            .map(|change| (change.row, change.column));
        let cache_key = cache_changes
            .get(cache)
            .map(|change| (change.row, change.column));
        let row = match (scalar_key, cache_key) {
            (Some(left), Some(right)) if left == right => {
                return Err(TileError::DuplicateOrUnsortedChange {
                    row: left.0,
                    column: left.1,
                });
            },
            (Some(left), Some(right)) if left < right => {
                scalar += 1;
                left.0
            },
            (Some(left), None) => {
                scalar += 1;
                left.0
            },
            (_, Some(right)) => {
                cache += 1;
                right.0
            },
            (None, None) => break,
        };
        if previous_row != Some(row) {
            rows = rows.checked_add(1).ok_or(TileError::InvalidSource)?;
            previous_row = Some(row);
        }
    }
    Ok(rows)
}

fn merge_tile_changes(
    scalar_changes: &[TileChange],
    cache_changes: &[CacheChange],
    merged: &mut Vec<TileChange>,
) -> Result<()> {
    let mut scalar = 0usize;
    let mut cache = 0usize;
    while scalar < scalar_changes.len() || cache < cache_changes.len() {
        let scalar_key = scalar_changes
            .get(scalar)
            .map(|change| (change.row, change.column));
        let cache_key = cache_changes
            .get(cache)
            .map(|change| (change.row, change.column));
        match (scalar_key, cache_key) {
            (Some(left), Some(right)) if left == right => {
                return Err(TileError::DuplicateOrUnsortedChange {
                    row: left.0,
                    column: left.1,
                });
            },
            (Some(left), Some(right)) if left < right => {
                merged.push(scalar_changes[scalar]);
                scalar += 1;
            },
            (Some(_), None) => {
                merged.push(scalar_changes[scalar]);
                scalar += 1;
            },
            (_, Some(_)) => {
                let change = &cache_changes[cache];
                merged.push(TileChange {
                    row: change.row,
                    column: change.column,
                    change: BncChange::FormulaCache(cache_scalar(&change.value, change)?),
                });
                cache += 1;
            },
            (None, None) => break,
        }
    }
    Ok(())
}

fn classify_prepared_mode(
    changes: &[TileChange],
    raw_rows: &[RawRow<'_>],
    modern: bool,
    counters: &mut Counters,
) -> Result<PreparedTileMode> {
    let last_source = raw_rows.last().map(|row| row.row.tile_row_index());
    let mut missing_materialized = false;
    for group in changes.chunk_by(|left, right| left.row == right.row) {
        counters.charge(u64_from_usize(binary_search_work(raw_rows.len()))?)?;
        if raw_rows
            .binary_search_by_key(&group[0].row, |row| row.row.tile_row_index())
            .is_ok()
        {
            continue;
        }
        let materialized = group.iter().any(|change| {
            matches!(
                change.change,
                BncChange::Set(_) | BncChange::FormulaSet { .. }
            )
        });
        if !materialized {
            continue;
        }
        if !modern {
            return Err(TileError::NeedSparse {
                row: group[0].row,
                column: group[0].column,
            });
        }
        if last_source.is_some_and(|last| group[0].row <= last) {
            return Err(TileError::NeedSparse {
                row: group[0].row,
                column: group[0].column,
            });
        }
        missing_materialized = true;
    }
    Ok(if missing_materialized && raw_rows.is_empty() {
        PreparedTileMode::New
    } else if missing_materialized {
        PreparedTileMode::PopulatedAppend
    } else {
        PreparedTileMode::Existing
    })
}

fn prepared_tile_output_len(
    source: &[u8],
    rows: &[PreparedRow<'_>],
    mode: PreparedTileMode,
    final_row_count: u32,
    counters: &mut Counters,
) -> Result<usize> {
    if rows.is_empty() {
        return Ok(0);
    }
    let view = WireView::parse(source).map_err(|_| TileError::InvalidSource)?;
    let mut length = 0usize;
    for (field_index, field) in view.fields().enumerate() {
        counters.charge(u64_from_usize(binary_search_work(rows.len()))?)?;
        let row = find_prepared_row_by_field(rows, field_index);
        let field_len = if let Some(row) = row {
            field
                .key()
                .len()
                .checked_add(varint_len(u64_from_usize(row.message_layout.output_len)?))
                .and_then(|value| value.checked_add(row.message_layout.output_len))
                .ok_or(TileError::InvalidSource)?
        } else {
            match (mode, field.number()) {
                (PreparedTileMode::New, 1..=3) => field.key().len() + 1,
                (PreparedTileMode::New | PreparedTileMode::PopulatedAppend, 4) => {
                    field.key().len() + varint_len(u64::from(final_row_count))
                },
                (PreparedTileMode::New, 5) => 0,
                _ => field.raw().len(),
            }
        };
        length = length
            .checked_add(field_len)
            .ok_or(TileError::InvalidSource)?;
    }
    for row in rows.iter().filter(|row| row.field_index.is_none()) {
        length = length
            .checked_add(1)
            .and_then(|value| {
                value.checked_add(varint_len(
                    u64::try_from(row.message_layout.output_len).ok()?,
                ))
            })
            .and_then(|value| value.checked_add(row.message_layout.output_len))
            .ok_or(TileError::InvalidSource)?;
    }
    Ok(length)
}

fn find_prepared_row_by_field<'row, 'source>(
    rows: &'row [PreparedRow<'source>],
    field_index: usize,
) -> Option<&'row PreparedRow<'source>> {
    let borrowed = rows.partition_point(|row| row.field_index.is_some());
    rows[..borrowed]
        .binary_search_by_key(&field_index, |row| row.field_index.unwrap_or(usize::MAX))
        .ok()
        .and_then(|index| rows.get(index))
}

fn tile_execution_requirements(
    source: &[u8],
    rows: &[PreparedRow<'_>],
    final_rows: &[RowCellCount],
    output_len: usize,
    prepared_retained_bytes: usize,
) -> Result<TileExecutionRequirements> {
    if output_len == 0 {
        let final_row_bytes = final_rows
            .len()
            .checked_mul(size_of::<RowCellCount>())
            .ok_or(TileError::InvalidSource)?;
        return Ok(TileExecutionRequirements {
            retained_bytes: final_row_bytes,
            retained_elements: final_rows.len(),
            peak_scratch_bytes: prepared_retained_bytes
                .checked_add(final_row_bytes)
                .ok_or(TileError::InvalidSource)?,
            allocations: usize::from(!final_rows.is_empty()),
            ..TileExecutionRequirements::default()
        });
    }
    let transition_count = rows.iter().try_fold(0usize, |total, row| {
        total
            .checked_add(row.transition_count)
            .ok_or(TileError::InvalidSource)
    })?;
    let transition_bytes = transition_count
        .checked_mul(size_of::<CellTransition>())
        .ok_or(TileError::InvalidSource)?;
    let final_row_bytes = final_rows
        .len()
        .checked_mul(size_of::<RowCellCount>())
        .ok_or(TileError::InvalidSource)?;
    let artifact_vector_bytes = transition_bytes
        .checked_add(final_row_bytes)
        .ok_or(TileError::InvalidSource)?;
    let mut input_bytes = source.len();
    let top_work = source
        .len()
        .checked_mul(binary_search_work(rows.len()))
        .and_then(|value| value.checked_add(output_len))
        .ok_or(TileError::InvalidSource)?;
    let mut work = u64_from_usize(top_work)?;
    let mut allocations = 1usize
        .checked_add(usize::from(transition_count != 0))
        .and_then(|value| value.checked_add(usize::from(!final_rows.is_empty())))
        .ok_or(TileError::InvalidSource)?;
    let mut prior_row_payloads = 0usize;
    let mut peak = prepared_retained_bytes
        .checked_add(artifact_vector_bytes)
        .ok_or(TileError::InvalidSource)?;
    let mut rows_read = 0usize;
    let mut rows_written = 0usize;
    let mut slots_scanned = 0usize;
    let mut slots_written = 0usize;
    let mut cache_count = 0usize;
    for row in rows {
        if let PreparedRowSource::Borrowed { raw } = row.source {
            input_bytes = input_bytes
                .checked_add(raw.len())
                .ok_or(TileError::InvalidSource)?;
            rows_read = rows_read.checked_add(1).ok_or(TileError::InvalidSource)?;
        }
        rows_written = rows_written
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        slots_scanned = slots_scanned
            .checked_add(row.slots.len())
            .ok_or(TileError::InvalidSource)?;
        slots_written = slots_written
            .checked_add(row.changed_slots)
            .ok_or(TileError::InvalidSource)?;
        cache_count = cache_count
            .checked_add(row.cache_count)
            .ok_or(TileError::InvalidSource)?;
        let mutation_bytes = row.slots.iter().try_fold(0usize, |total, slot| {
            let Some((previous, _, output_len, _)) = slot.mutation() else {
                return Ok(total);
            };
            input_bytes = input_bytes
                .checked_add(previous.map_or(0, <[u8]>::len))
                .ok_or(TileError::InvalidSource)?;
            work = work
                .checked_add(u64_from_usize(previous.map_or(0, <[u8]>::len))?)
                .ok_or(TileError::InvalidSource)?;
            output_len.map_or(Ok(total), |length| {
                allocations = allocations.checked_add(1).ok_or(TileError::InvalidSource)?;
                total.checked_add(length).ok_or(TileError::InvalidSource)
            })
        })?;
        allocations = allocations
            .checked_add(3 + usize::from(row.slot_layout.storage_capacity != 0))
            .ok_or(TileError::InvalidSource)?;
        let materialized_slot_bytes = row
            .slots
            .len()
            .checked_mul(size_of::<Slot<'_>>())
            .ok_or(TileError::InvalidSource)?;
        let row_transient = mutation_bytes
            .checked_add(materialized_slot_bytes)
            .and_then(|value| value.checked_add(row.slot_layout.storage_capacity))
            .and_then(|value| value.checked_add(row.slot_layout.offsets_len))
            .and_then(|value| value.checked_add(row.message_layout.output_len))
            .ok_or(TileError::InvalidSource)?;
        peak = peak.max(
            prepared_retained_bytes
                .checked_add(artifact_vector_bytes)
                .and_then(|value| value.checked_add(prior_row_payloads))
                .and_then(|value| value.checked_add(row_transient))
                .ok_or(TileError::InvalidSource)?,
        );
        prior_row_payloads = prior_row_payloads
            .checked_add(row.message_layout.output_len)
            .ok_or(TileError::InvalidSource)?;
        work = work
            .checked_add(u64_from_usize(
                mutation_bytes
                    .checked_add(row.slot_layout.storage_capacity)
                    .and_then(|value| value.checked_add(row.slot_layout.offsets_len))
                    .and_then(|value| value.checked_add(row.slots.len()))
                    .and_then(|value| value.checked_add(row.message_layout.output_len))
                    .ok_or(TileError::InvalidSource)?,
            )?)
            .ok_or(TileError::InvalidSource)?;
    }
    peak = peak.max(
        prepared_retained_bytes
            .checked_add(artifact_vector_bytes)
            .and_then(|value| value.checked_add(prior_row_payloads))
            .and_then(|value| value.checked_add(output_len))
            .ok_or(TileError::InvalidSource)?,
    );
    let retained_bytes = output_len
        .checked_add(artifact_vector_bytes)
        .ok_or(TileError::InvalidSource)?;
    let retained_elements = transition_count
        .checked_add(final_rows.len())
        .and_then(|value| value.checked_add(1))
        .ok_or(TileError::InvalidSource)?;
    let fields = input_bytes.checked_mul(8).ok_or(TileError::InvalidSource)?;
    Ok(TileExecutionRequirements {
        input_bytes,
        fields,
        work,
        output_bytes: output_len,
        retained_bytes,
        retained_elements,
        peak_scratch_bytes: peak,
        allocations,
        rows_read,
        rows_written,
        cell_slots_scanned: slots_scanned,
        cell_slots_written: slots_written,
        cache_cells_read: 0,
        cache_cells_written: cache_count,
    })
}

fn ensure_usize_limit(observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(observed)?,
            maximum: u64_from_usize(maximum)?,
        });
    }
    Ok(())
}

pub(crate) fn rewrite_tile(request: TileRewriteRequest<'_, '_>) -> Result<TileRewriteOutcome> {
    let limits = request.limits;
    let prepared = prepare_tile(request)?;
    execute_legacy_prepared(prepared, limits)
}

fn execute_legacy_prepared(
    prepared: PreparedTileRewrite<'_>,
    limits: TileLimits,
) -> Result<TileRewriteOutcome> {
    let prepare = prepared.prepare_report().report();
    let execution = prepared.execution_requirements();
    let total_input = usize_from_u64(prepare.wire_bytes)?
        .checked_add(execution.input_bytes())
        .ok_or(TileError::InvalidSource)?;
    let total_fields = usize_from_u64(prepare.wire_fields)?
        .checked_add(execution.fields())
        .ok_or(TileError::InvalidSource)?;
    let total_work = prepare
        .wire_work
        .checked_add(execution.work())
        .ok_or(TileError::InvalidSource)?;
    ensure_usize_limit(total_input, limits.max_input_bytes)?;
    ensure_usize_limit(total_fields, limits.max_fields)?;
    if total_work > limits.max_work {
        return Err(TileError::LimitExceeded {
            observed: total_work,
            maximum: limits.max_work,
        });
    }
    let mut outcome = prepared.execute(execution.exact_limits())?;
    outcome.report = merge_legacy_phase_reports(prepare, outcome.report)?;
    Ok(outcome)
}

fn merge_legacy_phase_reports(prepare: TileReport, execute: TileReport) -> Result<TileReport> {
    Ok(TileReport {
        wire_bytes: prepare
            .wire_bytes
            .checked_add(execute.wire_bytes)
            .ok_or(TileError::InvalidSource)?,
        wire_fields: prepare
            .wire_fields
            .checked_add(execute.wire_fields)
            .ok_or(TileError::InvalidSource)?,
        wire_work: prepare
            .wire_work
            .checked_add(execute.wire_work)
            .ok_or(TileError::InvalidSource)?,
        rows_read: prepare
            .rows_read
            .checked_add(execute.rows_read)
            .ok_or(TileError::InvalidSource)?,
        rows_written: execute.rows_written,
        cell_slots_scanned: prepare
            .cell_slots_scanned
            .checked_add(execute.cell_slots_scanned)
            .ok_or(TileError::InvalidSource)?,
        cell_slots_written: execute.cell_slots_written,
        cache_cells_read: prepare
            .cache_cells_read
            .checked_add(execute.cache_cells_read)
            .ok_or(TileError::InvalidSource)?,
        cache_cells_written: execute.cache_cells_written,
        output_bytes: execute.output_bytes,
        retained_elements: execute.retained_elements,
        retained_bytes: execute.retained_bytes,
        current_scratch_bytes: 0,
        peak_scratch_bytes: prepare.peak_scratch_bytes.max(execute.peak_scratch_bytes),
        allocation_events: prepare
            .allocation_events
            .checked_add(execute.allocation_events)
            .ok_or(TileError::InvalidSource)?,
    })
}

fn tile_has_no_rows(source: &[u8]) -> Result<bool> {
    let view = WireView::parse(source).map_err(|_| TileError::InvalidSource)?;
    for field in view.fields() {
        if field.number() == 5 {
            if field.wire_type() != 2 {
                return Err(TileError::InvalidSource);
            }
            return Ok(false);
        }
    }
    Ok(true)
}

/// Merge scalar/rich edits and formula display-cache refreshes into one tile
/// rewrite and one final payload allocation.
///
/// Both input slices must be strictly sorted and unique. Their coordinates
/// must be disjoint: a scalar replacement removes formula ownership, so a
/// cache refresh at that same coordinate would be contradictory.
#[cfg(test)]
pub(crate) fn rewrite_tile_with_cache(
    request: TileRewriteRequest<'_, '_>,
    cache_changes: &[CacheChange],
) -> Result<TileRewriteOutcome> {
    let limits = request.limits;
    let prepared = prepare_tile_with_cache(request, cache_changes)?;
    execute_legacy_prepared(prepared, limits)
}

/// Materialise all effective changes in one newly allocated type-6002 tile.
///
/// `template` is a strictly validated tile payload. Required summary fields
/// are rewritten into the modern BNC shape, every source row is removed, and
/// optional/unknown fields are retained exactly. Only rows containing a final
/// set operation are scaffolded; clears against the initially missing tile
/// remain absent. The scaffold then passes through [`rewrite_tile`] so new
/// and existing rows share one BNC mutation implementation.
pub(crate) fn rewrite_new_tile(
    template: &[u8],
    columns: u32,
    changes: &[TileChange],
    limits: TileLimits,
) -> Result<TileRewriteOutcome> {
    let prepared = prepare_new_tile(template, columns, changes, limits)?;
    execute_legacy_prepared(prepared, limits)
}

#[derive(Clone, Copy)]
struct RawRow<'source> {
    field_index: usize,
    row: storage::TileRowInfoSnapshot<'source>,
    payload: &'source [u8],
}

fn prepare_borrowed_row<'source>(
    row: storage::TileRowInfoSnapshot<'source>,
    raw: &'source [u8],
    field_index: usize,
    columns: u32,
    changes: &[TileChange],
    counters: &mut Counters,
) -> Result<PreparedRow<'source>> {
    let Some(storage_buffer) = row.cell_storage_buffer() else {
        let change = changes[0];
        return Err(TileError::NeedSparse {
            row: change.row,
            column: change.column,
        });
    };
    let Some(offsets) = row.cell_offsets() else {
        let change = changes[0];
        return Err(TileError::NeedSparse {
            row: change.row,
            column: change.column,
        });
    };
    if row.storage_version() != Some(BNC_STORAGE_VERSION) || !offsets.len().is_multiple_of(2) {
        return Err(TileError::InvalidSource);
    }
    let requested_columns = usize::try_from(columns).map_err(|_| TileError::InvalidSource)?;
    let slot_count = offsets.len() / 2;
    if slot_count < requested_columns
        || usize::try_from(row.cell_count()).map_err(|_| TileError::InvalidSource)? > slot_count
    {
        return Err(TileError::InvalidSource);
    }
    let slots = parse_slots(
        storage_buffer,
        offsets,
        row.has_wide_offsets().unwrap_or(false),
        slot_count,
        counters,
    )?;
    if slots.iter().filter(|slot| slot.bytes().is_some()).count()
        != usize::try_from(row.cell_count()).map_err(|_| TileError::InvalidSource)?
    {
        return Err(TileError::InvalidSource);
    }
    let prepared_bytes = slot_count
        .checked_mul(size_of::<PreparedSlot<'source>>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(prepared_bytes)?;
    let mut prepared = Vec::new();
    prepared
        .try_reserve_exact(slot_count)
        .map_err(|_| TileError::Allocation { amount: slot_count })?;
    if prepared.capacity() != slot_count {
        return Err(TileError::Allocation { amount: slot_count });
    }
    counters.record_scratch_allocation(prepared_bytes)?;
    for slot in &slots {
        prepared.push(match slot {
            Slot::Missing => PreparedSlot::Missing,
            Slot::Borrowed(bytes) => PreparedSlot::Borrowed(bytes),
            Slot::Owned(_) => return Err(TileError::InvalidSource),
        });
    }
    counters.release_scratch(
        slots
            .capacity()
            .checked_mul(size_of::<Slot<'_>>())
            .ok_or(TileError::InvalidSource)?,
    )?;
    drop(slots);

    let mut changed_slots = 0usize;
    let mut transition_count = 0usize;
    let mut cache_count = 0usize;
    for change in changes {
        let column = usize::try_from(change.column).map_err(|_| TileError::InvalidSource)?;
        if column >= requested_columns || column >= prepared.len() {
            return Err(TileError::OutOfBounds {
                row: change.row,
                column: change.column,
            });
        }
        let previous = prepared[column].borrowed();
        counters.charge(u64_from_usize(previous.map_or(0, <[u8]>::len))?)?;
        let cache_change = matches!(change.change, BncChange::FormulaCache(_));
        if cache_change {
            cache_count = cache_count.checked_add(1).ok_or(TileError::InvalidSource)?;
            counters.report.cache_cells_read = counters
                .report
                .cache_cells_read
                .checked_add(1)
                .ok_or(TileError::InvalidSource)?;
        }
        if cache_change && !matches!(classify_bnc_cell(previous)?, CellValue::Formula { .. }) {
            return Err(TileError::UnsupportedSource {
                row: change.row,
                column: change.column,
            });
        }
        if bnc_change_is_noop(previous, change.change)?
            || previous.is_none()
                && matches!(change.change, BncChange::Clear | BncChange::FormulaClear)
        {
            continue;
        }
        if let CellValue::Unsupported(kind) = classify_bnc_cell(previous)? {
            return Err(TileError::UnsupportedValue {
                row: change.row,
                column: change.column,
                kind,
            });
        }
        let output_len = plan_cell_mutation(previous, change.change)?;
        counters.charge(
            u64_from_usize(previous.map_or(0, <[u8]>::len))?
                .checked_mul(2)
                .ok_or(TileError::InvalidSource)?,
        )?;
        let transition = planned_transition(
            row.tile_row_index(),
            change.column,
            previous,
            change.change,
            output_len,
        )?;
        prepared[column] = PreparedSlot::Mutation {
            previous,
            change: change.change,
            output_len,
            transition,
        };
        changed_slots = changed_slots
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        if !cache_change {
            transition_count = transition_count
                .checked_add(1)
                .ok_or(TileError::InvalidSource)?;
        }
    }
    let wide = row.has_wide_offsets().unwrap_or(false) || prepared_slots_require_wide(&prepared)?;
    let slot_layout =
        plan_slot_lengths(prepared.iter().copied().map(PreparedSlot::output_len), wide)?;
    let cell_count = u32::try_from(
        prepared
            .iter()
            .filter(|slot| slot.output_len().is_some())
            .count(),
    )
    .map_err(|_| TileError::InvalidSource)?;
    let message_layout = plan_row_message(
        raw,
        cell_count,
        slot_layout.storage_len,
        slot_layout.offsets_len,
        row.has_wide_offsets(),
        wide,
    )?;
    counters.release_scratch(prepared_bytes)?;
    counters.retain(prepared_bytes, prepared.len())?;
    Ok(PreparedRow {
        field_index: Some(field_index),
        row: row.tile_row_index(),
        source: PreparedRowSource::Borrowed { raw },
        slots: prepared,
        slot_layout,
        message_layout,
        cell_count,
        changed_slots,
        transition_count,
        cache_count,
        output: None,
    })
}

fn prepare_canonical_row<'source>(
    row: u32,
    columns: u32,
    changes: &[TileChange],
    counters: &mut Counters,
) -> Result<PreparedRow<'source>> {
    let slot_count = usize::try_from(columns).map_err(|_| TileError::InvalidSource)?;
    let prepared_bytes = slot_count
        .checked_mul(size_of::<PreparedSlot<'source>>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(prepared_bytes)?;
    let mut slots = Vec::new();
    slots
        .try_reserve_exact(slot_count)
        .map_err(|_| TileError::Allocation { amount: slot_count })?;
    if slots.capacity() != slot_count {
        return Err(TileError::Allocation { amount: slot_count });
    }
    counters.record_scratch_allocation(prepared_bytes)?;
    slots.resize_with(slot_count, || PreparedSlot::Missing);
    let mut changed_slots = 0usize;
    let mut transition_count = 0usize;
    for change in changes {
        counters.charge(1)?;
        let column = usize::try_from(change.column).map_err(|_| TileError::InvalidSource)?;
        if column >= slot_count {
            return Err(TileError::OutOfBounds {
                row: change.row,
                column: change.column,
            });
        }
        if !matches!(
            change.change,
            BncChange::Set(_) | BncChange::FormulaSet { .. }
        ) {
            continue;
        }
        let output_len = plan_cell_mutation(None, change.change)?;
        let transition = planned_transition(row, change.column, None, change.change, output_len)?;
        slots[column] = PreparedSlot::Mutation {
            previous: None,
            change: change.change,
            output_len,
            transition,
        };
        changed_slots = changed_slots
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        transition_count = transition_count
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
    }
    let wide = prepared_slots_require_wide(&slots)?;
    let slot_layout = plan_slot_lengths(slots.iter().copied().map(PreparedSlot::output_len), wide)?;
    let cell_count = u32::try_from(
        slots
            .iter()
            .filter(|slot| slot.output_len().is_some())
            .count(),
    )
    .map_err(|_| TileError::InvalidSource)?;
    let message_layout = plan_canonical_row_message(
        row,
        cell_count,
        slot_layout.storage_len,
        slot_layout.offsets_len,
    )?;
    counters.release_scratch(prepared_bytes)?;
    counters.retain(prepared_bytes, slots.len())?;
    Ok(PreparedRow {
        field_index: None,
        row,
        source: PreparedRowSource::Canonical { row },
        slots,
        slot_layout,
        message_layout,
        cell_count,
        changed_slots,
        transition_count,
        cache_count: 0,
        output: None,
    })
}

fn prepared_slots_require_wide(slots: &[PreparedSlot<'_>]) -> Result<bool> {
    let mut length = 0usize;
    for slot in slots {
        let Some(slot_len) = slot.output_len() else {
            continue;
        };
        length = length
            .checked_add(slot_len)
            .ok_or(TileError::InvalidSource)?;
        if length > usize::from(MISSING_OFFSET) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn plan_canonical_row_message(
    row: u32,
    cell_count: u32,
    storage_len: usize,
    offsets_len: usize,
) -> Result<RowMessageLayout> {
    let output_len = 12usize
        .checked_add(varint_len(u64::from(row)))
        .and_then(|value| value.checked_add(varint_len(u64::from(cell_count))))
        .and_then(|value| value.checked_add(varint_len(u64::try_from(storage_len).ok()?)))
        .and_then(|value| value.checked_add(storage_len))
        .and_then(|value| value.checked_add(varint_len(u64::try_from(offsets_len).ok()?)))
        .and_then(|value| value.checked_add(offsets_len))
        .ok_or(TileError::InvalidSource)?;
    Ok(RowMessageLayout {
        output_len,
        append_wide: false,
    })
}

fn materialize_prepared_row<'source>(
    row: &PreparedRow<'source>,
    transitions: &mut Vec<CellTransition>,
    counters: &mut Counters,
) -> Result<Vec<Slot<'source>>> {
    let slot_bytes = row
        .slots
        .len()
        .checked_mul(size_of::<Slot<'source>>())
        .ok_or(TileError::InvalidSource)?;
    let mut materialized =
        exact_scratch_vec::<Slot<'source>>(row.slots.len(), slot_bytes, counters)?;
    for (column, slot) in row.slots.iter().copied().enumerate() {
        let Some((previous, change, planned_len, planned)) = slot.mutation() else {
            materialized.push(match slot {
                PreparedSlot::Missing => Slot::Missing,
                PreparedSlot::Borrowed(bytes) => Slot::Borrowed(bytes),
                PreparedSlot::Mutation { .. } => return Err(TileError::InvalidSource),
            });
            continue;
        };
        if planned.is_some_and(|transition| {
            transition.row != row.row || usize::try_from(transition.column).ok() != Some(column)
        }) {
            return Err(TileError::InvalidSource);
        }
        if let Some(length) = planned_len {
            counters.preflight_scratch_allocation(length)?;
        }
        let mutation = mutate_cell(previous, change, planned_len.unwrap_or(0).max(1))?;
        let after = match &mutation {
            CellMutation::Unchanged => previous,
            CellMutation::Delete => None,
            CellMutation::Replace(bytes) => Some(bytes.as_slice()),
        };
        if after.map(<[u8]>::len) != planned_len {
            return Err(TileError::InvalidSource);
        }
        let after_value = classify_bnc_cell(after)?;
        let after_references = bnc_references(after)?;
        let materialized_slot = match mutation {
            CellMutation::Unchanged => previous.map_or(Slot::Missing, Slot::Borrowed),
            CellMutation::Delete => Slot::Missing,
            CellMutation::Replace(bytes) => {
                if bytes.capacity() != bytes.len() {
                    return Err(TileError::Allocation {
                        amount: bytes.len(),
                    });
                }
                #[cfg(test)]
                record_prepared_execution_allocation();
                counters.record_scratch_allocation(bytes.capacity())?;
                Slot::Owned(bytes)
            },
        };
        materialized.push(materialized_slot);
        if let Some(planned) = planned {
            if planned.after != after_value || planned.after_references != after_references {
                return Err(TileError::InvalidSource);
            }
            transitions.push(planned);
        }
    }
    Ok(materialized)
}

fn planned_transition(
    row: u32,
    column: u32,
    previous: Option<&[u8]>,
    change: BncChange,
    output_len: Option<usize>,
) -> Result<Option<CellTransition>> {
    if matches!(change, BncChange::FormulaCache(_)) {
        return Ok(None);
    }
    let before = classify_bnc_cell(previous)?;
    let before_references = bnc_references(previous)?;
    let mut after_references = CellReferences {
        comment: before_references.comment,
        ..CellReferences::default()
    };
    let after = match change {
        BncChange::Clear | BncChange::FormulaClear => {
            if output_len.is_some() {
                CellValue::Empty
            } else {
                CellValue::Missing
            }
        },
        BncChange::Set(ScalarInput::String(identifier)) => {
            after_references.string = Some(identifier);
            CellValue::Text(identifier)
        },
        BncChange::Set(ScalarInput::RichText(identifier)) => {
            after_references.rich_text = Some(identifier);
            CellValue::RichText(identifier)
        },
        BncChange::Set(ScalarInput::Number(_)) => CellValue::Number,
        BncChange::Set(ScalarInput::Boolean(_)) => CellValue::Boolean,
        BncChange::Set(ScalarInput::Date(_)) => CellValue::Date,
        BncChange::Set(ScalarInput::Duration(_)) => CellValue::Duration,
        BncChange::FormulaSet { identifier, .. } => {
            after_references.formula = Some(identifier);
            CellValue::Formula {
                identifier,
                error: None,
            }
        },
        BncChange::FormulaCache(_) => return Ok(None),
    };
    Ok(Some(CellTransition {
        row,
        column,
        before,
        after,
        before_references,
        after_references,
    }))
}

fn write_canonical_row_message(
    row: u32,
    cell_count: u32,
    storage: &[u8],
    offsets: &[u8],
    wide: bool,
    layout: RowMessageLayout,
    counters: &mut Counters,
) -> Result<Vec<u8>> {
    let expected = plan_canonical_row_message(row, cell_count, storage.len(), offsets.len())?;
    if expected.output_len != layout.output_len || expected.append_wide != layout.append_wide {
        return Err(TileError::InvalidSource);
    }
    counters.charge(u64_from_usize(layout.output_len)?)?;
    let mut output = exact_scratch_vec::<u8>(layout.output_len, layout.output_len, counters)?;
    output.push(0x08);
    encode_varint(&mut output, u64::from(row));
    output.push(0x10);
    encode_varint(&mut output, u64::from(cell_count));
    output.extend_from_slice(&[0x1a, 0, 0x22, 0, 0x28, 5, 0x32]);
    encode_varint(
        &mut output,
        u64::try_from(storage.len()).map_err(|_| TileError::InvalidSource)?,
    );
    output.extend_from_slice(storage);
    output.push(0x3a);
    encode_varint(
        &mut output,
        u64::try_from(offsets.len()).map_err(|_| TileError::InvalidSource)?,
    );
    output.extend_from_slice(offsets);
    output.extend_from_slice(&[0x40, u8::from(wide)]);
    if output.len() != layout.output_len {
        return Err(TileError::InvalidSource);
    }
    Ok(output)
}

fn write_prepared_tile(plan: &PreparedTileRewrite<'_>, counters: &mut Counters) -> Result<Vec<u8>> {
    counters.charge(u64_from_usize(plan.output_len)?)?;
    let mut output = exact_scratch_vec::<u8>(plan.output_len, plan.output_len, counters)?;
    let view = WireView::parse(plan.source).map_err(|_| TileError::InvalidSource)?;
    for (field_index, field) in view.fields().enumerate() {
        if let Some(row) = find_prepared_row_by_field(&plan.rows, field_index) {
            let payload = row.output.as_deref().ok_or(TileError::InvalidSource)?;
            output.extend_from_slice(field.key());
            encode_varint(&mut output, u64_from_usize(payload.len())?);
            output.extend_from_slice(payload);
            continue;
        }
        match (plan.mode, field.number()) {
            (PreparedTileMode::New, 1..=3) => {
                output.extend_from_slice(field.key());
                output.push(0);
            },
            (PreparedTileMode::New | PreparedTileMode::PopulatedAppend, 4) => {
                output.extend_from_slice(field.key());
                encode_varint(&mut output, u64::from(plan.final_row_count));
            },
            (PreparedTileMode::New, 5) => {},
            _ => output.extend_from_slice(field.raw()),
        }
    }
    for row in plan.rows.iter().filter(|row| row.field_index.is_none()) {
        let payload = row.output.as_deref().ok_or(TileError::InvalidSource)?;
        output.push(0x2a);
        encode_varint(&mut output, u64_from_usize(payload.len())?);
        output.extend_from_slice(payload);
    }
    if output.len() != plan.output_len {
        return Err(TileError::InvalidSource);
    }
    Ok(output)
}

enum Slot<'source> {
    Missing,
    Borrowed(&'source [u8]),
    Owned(Vec<u8>),
}

impl Slot<'_> {
    fn bytes(&self) -> Option<&[u8]> {
        match self {
            Self::Missing => None,
            Self::Borrowed(bytes) => Some(bytes),
            Self::Owned(bytes) => Some(bytes),
        }
    }
}

fn parse_slots<'source>(
    storage: &'source [u8],
    offsets: &[u8],
    wide: bool,
    slot_count: usize,
    counters: &mut Counters,
) -> Result<Vec<Slot<'source>>> {
    let starts_bytes = slot_count
        .checked_mul(size_of::<(usize, usize)>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(starts_bytes)?;
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(slot_count)
        .map_err(|_| TileError::Allocation { amount: slot_count })?;
    if starts.capacity() != slot_count {
        return Err(TileError::Allocation { amount: slot_count });
    }
    counters.record_scratch_allocation(starts_bytes)?;
    let unit = if wide { 4usize } else { 1usize };
    for (column, bytes) in offsets.chunks_exact(2).enumerate() {
        counters.report.cell_slots_scanned = counters
            .report
            .cell_slots_scanned
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        counters.charge(1)?;
        let encoded = u16::from_le_bytes([bytes[0], bytes[1]]);
        if encoded != MISSING_OFFSET {
            let offset = usize::from(encoded)
                .checked_mul(unit)
                .ok_or(TileError::InvalidSource)?;
            if offset >= storage.len() {
                return Err(TileError::InvalidSource);
            }
            starts.push((column, offset));
        }
    }
    if starts.windows(2).any(|pair| pair[0].1 >= pair[1].1) {
        return Err(TileError::InvalidSource);
    }
    let slots_bytes = slot_count
        .checked_mul(size_of::<Slot<'source>>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(slots_bytes)?;
    let mut slots = Vec::new();
    slots
        .try_reserve_exact(slot_count)
        .map_err(|_| TileError::Allocation { amount: slot_count })?;
    if slots.capacity() != slot_count {
        return Err(TileError::Allocation { amount: slot_count });
    }
    counters.record_scratch_allocation(slots_bytes)?;
    let mut next_start = 0usize;
    for column in 0..slot_count {
        let bytes = starts
            .get(next_start)
            .filter(|(index, _)| *index == column)
            .map(|(_, start)| {
                let end = starts
                    .get(next_start + 1)
                    .map_or(storage.len(), |next| next.1);
                next_start += 1;
                &storage[*start..end]
            });
        slots.push(bytes.map_or(Slot::Missing, Slot::Borrowed));
    }
    counters.release_scratch(starts_bytes)?;
    Ok(slots)
}

enum CellMutation {
    Unchanged,
    Delete,
    Replace(Vec<u8>),
}

fn plan_cell_mutation(previous: Option<&[u8]>, change: BncChange) -> Result<Option<usize>> {
    let view = BncCellView::parse(previous.unwrap_or(&MINIMAL_BNC_CELL)).map_err(map_bnc_error)?;
    let plan = match change {
        BncChange::Clear => view.plan_clear_value(false),
        BncChange::FormulaClear => view.plan_clear_value(true),
        BncChange::Set(input) => view.plan_scalar_rewrite(input.as_wire()),
        BncChange::FormulaSet { identifier, cache } => {
            view.plan_formula_rewrite(identifier, cache.map(ScalarInput::as_wire))
        },
        BncChange::FormulaCache(input) => view.plan_formula_cache_rewrite(input.as_wire()),
    }
    .map_err(map_bnc_error)?;
    Ok(plan.output_len())
}

fn mutate_cell(
    previous: Option<&[u8]>,
    change: BncChange,
    max_output_bytes: usize,
) -> Result<CellMutation> {
    let Some(previous) = previous else {
        return match change {
            BncChange::Clear | BncChange::FormulaClear => Ok(CellMutation::Unchanged),
            BncChange::Set(input) => {
                let view =
                    BncCellView::parse(&MINIMAL_BNC_CELL).map_err(|_| TileError::InvalidSource)?;
                view.rewrite_scalar_with_limit(input.as_wire(), max_output_bytes)
                    .map(CellMutation::Replace)
                    .map_err(map_bnc_error)
            },
            BncChange::FormulaSet { identifier, cache } => {
                let view =
                    BncCellView::parse(&MINIMAL_BNC_CELL).map_err(|_| TileError::InvalidSource)?;
                match cache {
                    Some(cache) => view.rewrite_formula_with_limit(
                        identifier,
                        cache.as_wire(),
                        max_output_bytes,
                    ),
                    None => {
                        view.rewrite_formula_without_cache_with_limit(identifier, max_output_bytes)
                    },
                }
                .map(CellMutation::Replace)
                .map_err(map_bnc_error)
            },
            BncChange::FormulaCache(_) => Err(TileError::UnsupportedSource { row: 0, column: 0 }),
        };
    };
    let view = BncCellView::parse(previous).map_err(|_| TileError::InvalidSource)?;
    match change {
        BncChange::Clear | BncChange::FormulaClear => match view
            .clear_value_with_limit(max_output_bytes)
            .map_err(map_bnc_error)?
        {
            ClearValue::Delete if matches!(change, BncChange::Clear) => Ok(CellMutation::Delete),
            ClearValue::Delete => exact_minimal_cell(max_output_bytes).map(CellMutation::Replace),
            ClearValue::Retain(bytes) => Ok(CellMutation::Replace(bytes)),
        },
        BncChange::Set(input) => view
            .rewrite_scalar_with_limit(input.as_wire(), max_output_bytes)
            .map(CellMutation::Replace)
            .map_err(map_bnc_error),
        BncChange::FormulaSet { identifier, cache } => match cache {
            Some(cache) => {
                view.rewrite_formula_with_limit(identifier, cache.as_wire(), max_output_bytes)
            },
            None => view.rewrite_formula_without_cache_with_limit(identifier, max_output_bytes),
        }
        .map(CellMutation::Replace)
        .map_err(map_bnc_error),
        BncChange::FormulaCache(input) => view
            .rewrite_formula_cache_with_limit(input.as_wire(), max_output_bytes)
            .map(CellMutation::Replace)
            .map_err(map_bnc_error),
    }
}

fn exact_minimal_cell(max_output_bytes: usize) -> Result<Vec<u8>> {
    if MINIMAL_BNC_CELL.len() > max_output_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(MINIMAL_BNC_CELL.len())?,
            maximum: u64_from_usize(max_output_bytes)?,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(MINIMAL_BNC_CELL.len())
        .map_err(|_| TileError::Allocation {
            amount: MINIMAL_BNC_CELL.len(),
        })?;
    if output.capacity() != MINIMAL_BNC_CELL.len() {
        return Err(TileError::Allocation {
            amount: MINIMAL_BNC_CELL.len(),
        });
    }
    output.extend_from_slice(&MINIMAL_BNC_CELL);
    Ok(output)
}

impl ScalarInput {
    const fn as_wire(self) -> ScalarValue {
        match self {
            Self::String(identifier) => ScalarValue::String(identifier),
            Self::RichText(identifier) => ScalarValue::RichText(identifier),
            Self::Number(value) => ScalarValue::Number(value),
            Self::Boolean(value) => ScalarValue::Boolean(value),
            Self::Date(value) => ScalarValue::Date(value),
            Self::Duration(value) => ScalarValue::Duration(value),
        }
    }
}

impl CacheScalarInput {
    const fn as_wire(self) -> CachedScalar {
        match self {
            Self::Number(value) => CachedScalar::Number(value),
            Self::Boolean(value) => CachedScalar::Boolean(value),
        }
    }
}

fn map_bnc_error(error: BncError) -> TileError {
    match error {
        BncError::OutputLimitExceeded { observed, maximum } => {
            let Ok(observed) = u64::try_from(observed) else {
                return TileError::InvalidSource;
            };
            let Ok(maximum) = u64::try_from(maximum) else {
                return TileError::InvalidSource;
            };
            TileError::LimitExceeded { observed, maximum }
        },
        BncError::Allocation { requested } => TileError::Allocation { amount: requested },
        BncError::InvalidFormat(_) | BncError::ParseError(_) => TileError::InvalidSource,
    }
}

/// Classify one optional raw BNC slot without taking ownership of it.
pub(crate) fn classify_bnc_cell(cell: Option<&[u8]>) -> Result<CellValue> {
    let Some(cell) = cell else {
        return Ok(CellValue::Missing);
    };
    let cell = BncCellView::parse(cell).map_err(|_| TileError::InvalidSource)?;
    Ok(classify_bnc_view(&cell))
}

fn classify_bnc_view(cell: &BncCellView<'_>) -> CellValue {
    match cell.stored_value() {
        StoredValue::Empty => CellValue::Empty,
        StoredValue::Number => CellValue::Number,
        StoredValue::Text(identifier) => CellValue::Text(identifier),
        StoredValue::Formula(identifier) => CellValue::Formula {
            identifier,
            error: cell.formula_error_identifier(),
        },
        StoredValue::RichText(identifier) => CellValue::RichText(identifier),
        StoredValue::Date => CellValue::Date,
        StoredValue::Boolean => CellValue::Boolean,
        StoredValue::Duration => CellValue::Duration,
        StoredValue::Error => CellValue::Error(cell.formula_error_identifier()),
        StoredValue::Unsupported(kind) => CellValue::Unsupported(kind),
    }
}

/// Read the native references carried by one optional raw BNC slot.
pub(crate) fn bnc_references(cell: Option<&[u8]>) -> Result<CellReferences> {
    let Some(cell) = cell else {
        return Ok(CellReferences::default());
    };
    let cell = BncCellView::parse(cell).map_err(|_| TileError::InvalidSource)?;
    Ok(bnc_references_view(&cell))
}

fn bnc_references_view(cell: &BncCellView<'_>) -> CellReferences {
    let value = cell.stored_value();
    let (string, rich_text, formula) = match value {
        StoredValue::Text(identifier) => (Some(identifier), None, None),
        StoredValue::RichText(identifier) => (None, Some(identifier), None),
        StoredValue::Formula(identifier) => (None, None, Some(identifier)),
        _ => (None, None, None),
    };
    CellReferences {
        string,
        rich_text,
        formula,
        formula_error: cell.formula_error_identifier(),
        comment: cell.comment_identifier(),
    }
}

fn classify_bnc_with_references(
    cell: Option<&[u8]>,
) -> Result<(CellValue, CellReferences, Option<FormulaCacheValue>)> {
    let Some(cell) = cell else {
        return Ok((CellValue::Missing, CellReferences::default(), None));
    };
    let view = BncCellView::parse(cell).map_err(|_| TileError::InvalidSource)?;
    let formula_cache = if matches!(view.stored_value(), StoredValue::Formula(_)) {
        match view.cached_scalar() {
            Some(CachedScalar::Number(value)) => Some(FormulaCacheValue::Number(value)),
            Some(CachedScalar::Boolean(value)) => Some(FormulaCacheValue::Boolean(value)),
            Some(CachedScalar::Date(value)) => Some(FormulaCacheValue::Date(value)),
            Some(CachedScalar::Duration(value)) => Some(FormulaCacheValue::Duration(value)),
            Some(CachedScalar::Unsupported(_)) => None,
            None => view.formula_text_key().map(FormulaCacheValue::TextKey),
        }
    } else {
        None
    };
    Ok((
        classify_bnc_view(&view),
        bnc_references_view(&view),
        formula_cache,
    ))
}

/// Return whether applying `change` would leave the cell's public scalar
/// state unchanged.  This is deliberately view-only: it avoids allocating a
/// `BncCell` merely to discover an exact semantic no-op.
pub(crate) fn bnc_change_is_noop(cell: Option<&[u8]>, change: BncChange) -> Result<bool> {
    let Some(cell) = cell else {
        return Ok(matches!(change, BncChange::Clear));
    };
    let view = BncCellView::parse(cell).map_err(|_| TileError::InvalidSource)?;
    Ok(match change {
        BncChange::Clear => matches!(view.stored_value(), StoredValue::Empty),
        BncChange::FormulaClear => false,
        BncChange::Set(input) => view.scalar_equals(input.as_wire()),
        BncChange::FormulaSet { identifier, cache } => match cache {
            Some(cache) => view
                .formula_value_equals(identifier, cache.as_wire())
                .map_err(map_bnc_error)?,
            None => {
                view.stored_value() == StoredValue::Formula(identifier)
                    && view.cached_scalar().is_none()
            },
        },
        BncChange::FormulaCache(input) => view.formula_cache_equals(input.as_wire()),
    })
}

fn encode_slots(
    slots: &[Slot<'_>],
    prefer_wide: bool,
    counters: &mut Counters,
) -> Result<(Vec<u8>, Vec<u8>, bool)> {
    let wide = prefer_wide || slots_require_wide(slots)?;
    let layout = plan_slot_layout(slots, wide)?;
    let result = encode_slots_with_width(slots, layout, counters)?;
    Ok((result.0, result.1, wide))
}

#[derive(Clone, Copy)]
struct SlotLayout {
    wide: bool,
    storage_len: usize,
    storage_capacity: usize,
    offsets_len: usize,
}

fn plan_slot_layout(slots: &[Slot<'_>], wide: bool) -> Result<SlotLayout> {
    plan_slot_lengths(slots.iter().map(|slot| slot.bytes().map(<[u8]>::len)), wide)
}

fn plan_slot_lengths(
    lengths: impl ExactSizeIterator<Item = Option<usize>>,
    wide: bool,
) -> Result<SlotLayout> {
    let slot_count = lengths.len();
    let offsets_len = slot_count.checked_mul(2).ok_or(TileError::InvalidSource)?;
    let unit = if wide { 4usize } else { 1usize };
    let mut storage_len = 0usize;
    let mut content = 0usize;
    for length in lengths {
        let Some(length) = length else { continue };
        storage_len = storage_len
            .checked_add((unit - storage_len % unit) % unit)
            .and_then(|value| value.checked_add(length))
            .ok_or(TileError::InvalidSource)?;
        content = content
            .checked_add(length)
            .ok_or(TileError::InvalidSource)?;
    }
    storage_len = storage_len
        .checked_add((unit - storage_len % unit) % unit)
        .ok_or(TileError::InvalidSource)?;
    let storage_capacity = content
        .checked_add(slot_count.checked_mul(3).ok_or(TileError::InvalidSource)?)
        .ok_or(TileError::InvalidSource)?;
    if storage_len > storage_capacity {
        return Err(TileError::InvalidSource);
    }
    Ok(SlotLayout {
        wide,
        storage_len,
        storage_capacity,
        offsets_len,
    })
}

fn slots_require_wide(slots: &[Slot<'_>]) -> Result<bool> {
    let mut length = 0usize;
    for slot in slots {
        let Some(bytes) = slot.bytes() else {
            continue;
        };
        length = length
            .checked_add(bytes.len())
            .ok_or(TileError::InvalidSource)?;
        if length > usize::from(MISSING_OFFSET) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn encode_slots_with_width(
    slots: &[Slot<'_>],
    layout: SlotLayout,
    counters: &mut Counters,
) -> Result<(Vec<u8>, Vec<u8>)> {
    counters.charge(u64_from_usize(layout.storage_capacity)?)?;
    counters.charge(u64_from_usize(layout.offsets_len)?)?;
    let mut storage = Vec::new();
    counters.preflight_scratch_allocation(layout.storage_capacity)?;
    storage
        .try_reserve_exact(layout.storage_capacity)
        .map_err(|_| TileError::Allocation {
            amount: layout.storage_capacity,
        })?;
    if storage.capacity() != layout.storage_capacity {
        return Err(TileError::Allocation {
            amount: layout.storage_capacity,
        });
    }
    counters.record_scratch_allocation(storage.capacity())?;
    let mut offsets = Vec::new();
    counters.preflight_scratch_allocation(layout.offsets_len)?;
    offsets
        .try_reserve_exact(layout.offsets_len)
        .map_err(|_| TileError::Allocation {
            amount: layout.offsets_len,
        })?;
    if offsets.capacity() != layout.offsets_len {
        return Err(TileError::Allocation {
            amount: layout.offsets_len,
        });
    }
    counters.record_scratch_allocation(offsets.capacity())?;
    let unit = if layout.wide { 4usize } else { 1usize };
    for slot in slots {
        counters.charge(1)?;
        let Some(cell) = slot.bytes() else {
            offsets.extend_from_slice(&MISSING_OFFSET.to_le_bytes());
            continue;
        };
        while storage.len() % unit != 0 {
            storage.push(0);
        }
        let offset = storage.len() / unit;
        let offset =
            u16::try_from(offset).map_err(|_| TileError::NeedSparse { row: 0, column: 0 })?;
        if offset == MISSING_OFFSET {
            return Err(TileError::NeedSparse { row: 0, column: 0 });
        }
        offsets.extend_from_slice(&offset.to_le_bytes());
        storage.extend_from_slice(cell);
    }
    while storage.len() % unit != 0 {
        storage.push(0);
    }
    if storage.len() != layout.storage_len || offsets.len() != layout.offsets_len {
        return Err(TileError::InvalidSource);
    }
    Ok((storage, offsets))
}

#[derive(Clone, Copy)]
struct RowMessageLayout {
    output_len: usize,
    append_wide: bool,
}

fn plan_row_message(
    source: &[u8],
    cell_count: u32,
    storage_len: usize,
    offsets_len: usize,
    previous_wide: Option<bool>,
    wide: bool,
) -> Result<RowMessageLayout> {
    let view = WireView::parse(source).map_err(|_| TileError::InvalidSource)?;
    let mut count_present = false;
    let mut storage_present = false;
    let mut offsets_present = false;
    let mut wide_present = false;
    let mut length = 0usize;
    for field in view.fields() {
        length = length
            .checked_add(match field.number() {
                2 => {
                    count_present = true;
                    field.key().len() + varint_len(u64::from(cell_count))
                },
                6 => {
                    storage_present = true;
                    field.key().len()
                        + varint_len(
                            u64::try_from(storage_len).map_err(|_| TileError::InvalidSource)?,
                        )
                        + storage_len
                },
                7 => {
                    offsets_present = true;
                    field.key().len()
                        + varint_len(
                            u64::try_from(offsets_len).map_err(|_| TileError::InvalidSource)?,
                        )
                        + offsets_len
                },
                8 => {
                    wide_present = true;
                    field.key().len() + 1
                },
                _ => field.raw().len(),
            })
            .ok_or(TileError::InvalidSource)?;
    }
    let append_wide = previous_wide.is_none() && wide && !wide_present;
    if !count_present || !storage_present || !offsets_present || (previous_wide.is_none() && wide) {
        if append_wide {
            length = length.checked_add(2).ok_or(TileError::InvalidSource)?;
        } else {
            return Err(TileError::InvalidSource);
        }
    }
    Ok(RowMessageLayout {
        output_len: length,
        append_wide,
    })
}

fn write_row_message(
    source: &[u8],
    cell_count: u32,
    storage: &[u8],
    offsets: &[u8],
    wide: bool,
    layout: RowMessageLayout,
    counters: &mut Counters,
) -> Result<Vec<u8>> {
    let view = WireView::parse(source).map_err(|_| TileError::InvalidSource)?;
    counters.charge(u64_from_usize(layout.output_len)?)?;
    let mut output = Vec::new();
    counters.preflight_scratch_allocation(layout.output_len)?;
    output
        .try_reserve_exact(layout.output_len)
        .map_err(|_| TileError::Allocation {
            amount: layout.output_len,
        })?;
    if output.capacity() != layout.output_len {
        return Err(TileError::Allocation {
            amount: layout.output_len,
        });
    }
    counters.record_scratch_allocation(output.capacity())?;
    for field in view.fields() {
        match field.number() {
            2 => {
                output.extend_from_slice(field.key());
                encode_varint(&mut output, u64::from(cell_count));
            },
            6 => {
                output.extend_from_slice(field.key());
                encode_varint(
                    &mut output,
                    u64::try_from(storage.len()).map_err(|_| TileError::InvalidSource)?,
                );
                output.extend_from_slice(storage);
            },
            7 => {
                output.extend_from_slice(field.key());
                encode_varint(
                    &mut output,
                    u64::try_from(offsets.len()).map_err(|_| TileError::InvalidSource)?,
                );
                output.extend_from_slice(offsets);
            },
            8 => {
                output.extend_from_slice(field.key());
                encode_varint(&mut output, if wide { 1 } else { 0 });
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    if layout.append_wide {
        output.extend_from_slice(&[0x40, 1]);
    }
    if output.len() != layout.output_len {
        return Err(TileError::InvalidSource);
    }
    Ok(output)
}

fn validate_request(request: TileRewriteRequest<'_, '_>) -> Result<()> {
    if request.source.len() > request.limits.max_input_bytes || request.columns == 0 {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(request.source.len())?,
            maximum: u64_from_usize(request.limits.max_input_bytes)?,
        });
    }
    let mut previous = None;
    for change in request.changes {
        if change.column >= request.columns {
            return Err(TileError::OutOfBounds {
                row: change.row,
                column: change.column,
            });
        }
        if previous.is_some_and(|last: (u32, u32)| last >= (change.row, change.column)) {
            return Err(TileError::DuplicateOrUnsortedChange {
                row: change.row,
                column: change.column,
            });
        }
        previous = Some((change.row, change.column));
    }
    Ok(())
}

fn validate_read_positions(
    source: &[u8],
    columns: u32,
    positions: &[TileReadPosition],
    limits: TileLimits,
) -> Result<()> {
    if source.len() > limits.max_input_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(source.len())?,
            maximum: u64_from_usize(limits.max_input_bytes)?,
        });
    }
    if columns == 0 {
        let position = positions[0];
        return Err(TileError::OutOfBounds {
            row: position.row,
            column: position.column,
        });
    }
    if positions.len() > limits.max_cells {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(positions.len())?,
            maximum: u64_from_usize(limits.max_cells)?,
        });
    }
    let mut previous = None;
    for position in positions {
        if position.column >= columns
            || usize::try_from(position.row).map_or(true, |row| row >= limits.max_rows)
        {
            return Err(TileError::OutOfBounds {
                row: position.row,
                column: position.column,
            });
        }
        if previous.is_some_and(|last: (u32, u32)| last >= (position.row, position.column)) {
            return Err(TileError::DuplicateOrUnsortedChange {
                row: position.row,
                column: position.column,
            });
        }
        previous = Some((position.row, position.column));
    }
    Ok(())
}

fn validate_cache_changes(changes: &[CacheChange], columns: u32, limits: TileLimits) -> Result<()> {
    if changes.len() > limits.max_cells {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(changes.len())?,
            maximum: u64_from_usize(limits.max_cells)?,
        });
    }
    let mut previous = None;
    for change in changes {
        if change.column >= columns
            || usize::try_from(change.row).map_or(true, |row| row >= limits.max_rows)
        {
            return Err(TileError::OutOfBounds {
                row: change.row,
                column: change.column,
            });
        }
        if previous.is_some_and(|last: (u32, u32)| last >= (change.row, change.column)) {
            return Err(TileError::DuplicateOrUnsortedChange {
                row: change.row,
                column: change.column,
            });
        }
        let _ = cache_scalar(&change.value, change)?;
        previous = Some((change.row, change.column));
    }
    Ok(())
}

fn validate_disjoint_changes(scalar: &[TileChange], cache: &[CacheChange]) -> Result<()> {
    let mut scalar_index = 0usize;
    let mut cache_index = 0usize;
    while let (Some(left), Some(right)) = (
        scalar
            .get(scalar_index)
            .map(|change| (change.row, change.column)),
        cache
            .get(cache_index)
            .map(|change| (change.row, change.column)),
    ) {
        match left.cmp(&right) {
            std::cmp::Ordering::Less => scalar_index += 1,
            std::cmp::Ordering::Greater => cache_index += 1,
            std::cmp::Ordering::Equal => {
                return Err(TileError::DuplicateOrUnsortedChange {
                    row: left.0,
                    column: left.1,
                });
            },
        }
    }
    Ok(())
}

fn cache_scalar(value: &FormulaCachedValue, change: &CacheChange) -> Result<CacheScalarInput> {
    match value {
        FormulaCachedValue::Number(value) => Ok(CacheScalarInput::Number(*value)),
        FormulaCachedValue::Boolean(value) => Ok(CacheScalarInput::Boolean(*value)),
        FormulaCachedValue::Text(_) => Err(TileError::UnsupportedValue {
            row: change.row,
            column: change.column,
            kind: 3,
        }),
        FormulaCachedValue::Date(_) => Err(TileError::UnsupportedValue {
            row: change.row,
            column: change.column,
            kind: 5,
        }),
        FormulaCachedValue::Duration(_) => Err(TileError::UnsupportedValue {
            row: change.row,
            column: change.column,
            kind: 7,
        }),
    }
}

struct Counters {
    limits: TileLimits,
    report: TileReport,
    work: u64,
}
impl Counters {
    const fn new(limits: TileLimits) -> Self {
        Self {
            limits,
            report: TileReport {
                wire_bytes: 0,
                wire_fields: 0,
                wire_work: 0,
                rows_read: 0,
                rows_written: 0,
                cell_slots_scanned: 0,
                cell_slots_written: 0,
                cache_cells_read: 0,
                cache_cells_written: 0,
                output_bytes: 0,
                retained_elements: 0,
                retained_bytes: 0,
                current_scratch_bytes: 0,
                peak_scratch_bytes: 0,
                allocation_events: 0,
            },
            work: 0,
        }
    }
    fn charge_bytes(&mut self, bytes: usize) -> Result<()> {
        self.charge(u64_from_usize(bytes)?)
    }
    fn charge(&mut self, amount: u64) -> Result<()> {
        let next = self.checked_work(amount)?;
        self.work = next;
        self.report.wire_work = next;
        Ok(())
    }
    fn preflight_scratch_allocation(&mut self, bytes: usize) -> Result<()> {
        let _ = self.checked_work(u64_from_usize(bytes)?)?;
        let next = self
            .report
            .current_scratch_bytes
            .checked_add(u64_from_usize(bytes)?)
            .ok_or(TileError::InvalidSource)?;
        let maximum = self
            .limits
            .max_output_bytes
            .min(self.limits.max_peak_scratch_bytes);
        if next > u64_from_usize(maximum)? {
            return Err(TileError::LimitExceeded {
                observed: next,
                maximum: u64_from_usize(maximum)?,
            });
        }
        Ok(())
    }
    fn record_scratch_allocation(&mut self, bytes: usize) -> Result<()> {
        if bytes == 0 {
            return Ok(());
        }
        self.record_allocation_event()?;
        let current = self
            .report
            .current_scratch_bytes
            .checked_add(u64_from_usize(bytes)?)
            .ok_or(TileError::InvalidSource)?;
        self.report.current_scratch_bytes = current;
        self.report.peak_scratch_bytes = self.report.peak_scratch_bytes.max(current);
        Ok(())
    }
    fn release_scratch(&mut self, bytes: usize) -> Result<()> {
        self.report.current_scratch_bytes = self
            .report
            .current_scratch_bytes
            .checked_sub(u64_from_usize(bytes)?)
            .ok_or(TileError::InvalidSource)?;
        Ok(())
    }
    fn retain(&mut self, bytes: usize, elements: usize) -> Result<()> {
        let retained_bytes = self
            .report
            .retained_bytes
            .checked_add(u64_from_usize(bytes)?)
            .ok_or(TileError::InvalidSource)?;
        let retained_elements = self
            .report
            .retained_elements
            .checked_add(u64_from_usize(elements)?)
            .ok_or(TileError::InvalidSource)?;
        if retained_bytes > u64_from_usize(self.limits.max_retained_bytes)? {
            return Err(TileError::LimitExceeded {
                observed: retained_bytes,
                maximum: u64_from_usize(self.limits.max_retained_bytes)?,
            });
        }
        if retained_elements > u64_from_usize(self.limits.max_retained_elements)? {
            return Err(TileError::LimitExceeded {
                observed: retained_elements,
                maximum: u64_from_usize(self.limits.max_retained_elements)?,
            });
        }
        self.report.retained_bytes = retained_bytes;
        self.report.retained_elements = retained_elements;
        Ok(())
    }
    fn record_allocation_event(&mut self) -> Result<()> {
        let events = self
            .report
            .allocation_events
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
        if events > u64_from_usize(self.limits.max_allocations)? {
            return Err(TileError::LimitExceeded {
                observed: events,
                maximum: u64_from_usize(self.limits.max_allocations)?,
            });
        }
        self.report.allocation_events = events;
        Ok(())
    }
    fn checked_work(&self, amount: u64) -> Result<u64> {
        let next = self
            .work
            .checked_add(amount)
            .ok_or(TileError::LimitExceeded {
                observed: u64::MAX,
                maximum: self.limits.max_work,
            })?;
        if next > self.limits.max_work {
            return Err(TileError::LimitExceeded {
                observed: next,
                maximum: self.limits.max_work,
            });
        }
        Ok(next)
    }
}

fn encode_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn u64_from_usize(value: usize) -> Result<u64> {
    u64::try_from(value).map_err(|_| TileError::InvalidSource)
}

fn usize_from_u64(value: u64) -> Result<usize> {
    usize::try_from(value).map_err(|_| TileError::LimitExceeded {
        observed: value,
        maximum: u64::MAX,
    })
}

const fn varint_len(mut value: u64) -> usize {
    let mut bytes = 1usize;
    while value >= 0x80 {
        bytes += 1;
        value >>= 7;
    }
    bytes
}

const fn binary_search_work(length: usize) -> usize {
    let mut span = length;
    let mut work = 1usize;
    while span > 1 {
        span = span.div_ceil(2);
        work += 1;
    }
    work
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::wire::BncCell;

    #[derive(Default)]
    struct Rows {
        values: Vec<(u32, u32, Vec<u8>, Vec<u8>, bool)>,
    }

    impl storage::StorageVisitor for Rows {
        fn visit_tile_row(
            &mut self,
            row: storage::TileRowInfoSnapshot<'_>,
        ) -> std::result::Result<(), storage::DecodeError> {
            self.values.push((
                row.tile_row_index(),
                row.cell_count(),
                row.cell_storage_buffer().unwrap_or_default().to_vec(),
                row.cell_offsets().unwrap_or_default().to_vec(),
                row.has_wide_offsets().unwrap_or(false),
            ));
            Ok(())
        }
    }

    fn limits() -> TileLimits {
        TileLimits::new(1 << 20, 1 << 20, 1 << 16, 1 << 24, 256, 1 << 16)
    }

    fn field(output: &mut Vec<u8>, number: u32, value: u64) {
        encode_varint(output, u64::from(number) << 3);
        encode_varint(output, value);
    }

    fn bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        encode_varint(output, (u64::from(number) << 3) | 2);
        encode_varint(output, value.len() as u64);
        output.extend_from_slice(value);
    }

    fn empty_row(row: u32, columns: usize) -> Vec<u8> {
        let mut output = Vec::new();
        field(&mut output, 1, u64::from(row));
        field(&mut output, 2, 0);
        bytes_field(&mut output, 3, &[]);
        bytes_field(&mut output, 4, &[]);
        field(&mut output, 5, 5);
        bytes_field(&mut output, 6, &[]);
        bytes_field(&mut output, 7, &vec![0xff; columns * 2]);
        field(&mut output, 8, 0);
        output
    }

    fn populated_template() -> Vec<u8> {
        let mut output = Vec::new();
        field(&mut output, 1, 9);
        field(&mut output, 2, 9);
        field(&mut output, 3, 0);
        field(&mut output, 4, 1);
        bytes_field(&mut output, 5, &empty_row(4, 2));
        field(&mut output, 6, 5);
        field(&mut output, 7, 1);
        field(&mut output, 8, 0);
        bytes_field(&mut output, 99, b"opaque-tile-extension");
        output
    }

    fn empty_modern_tile() -> Vec<u8> {
        let mut output = Vec::new();
        field(&mut output, 1, 9);
        field(&mut output, 2, 9);
        field(&mut output, 3, 0);
        field(&mut output, 4, 0);
        field(&mut output, 6, 5);
        field(&mut output, 7, 1);
        field(&mut output, 8, 0);
        bytes_field(&mut output, 99, b"opaque-empty-tile-extension");
        output
    }

    fn one_row_tile(row: u32, cells: &[Option<Vec<u8>>]) -> Vec<u8> {
        let mut storage_buffer = Vec::new();
        let mut offsets = Vec::new();
        let mut count = 0u32;
        for cell in cells {
            if let Some(cell) = cell {
                offsets.extend_from_slice(&(storage_buffer.len() as u16).to_le_bytes());
                storage_buffer.extend_from_slice(cell);
                count += 1;
            } else {
                offsets.extend_from_slice(&MISSING_OFFSET.to_le_bytes());
            }
        }
        let mut row_payload = Vec::new();
        field(&mut row_payload, 1, u64::from(row));
        field(&mut row_payload, 2, u64::from(count));
        bytes_field(&mut row_payload, 3, &[]);
        bytes_field(&mut row_payload, 4, &[]);
        field(&mut row_payload, 5, 5);
        bytes_field(&mut row_payload, 6, &storage_buffer);
        bytes_field(&mut row_payload, 7, &offsets);
        field(&mut row_payload, 8, 0);

        let mut tile = Vec::new();
        field(&mut tile, 1, 0);
        field(&mut tile, 2, 0);
        field(&mut tile, 3, 0);
        field(&mut tile, 4, 1);
        bytes_field(&mut tile, 5, &row_payload);
        field(&mut tile, 6, 5);
        field(&mut tile, 7, 1);
        field(&mut tile, 8, 0);
        tile
    }

    fn row_cell(row: &(u32, u32, Vec<u8>, Vec<u8>, bool), column: usize) -> &[u8] {
        let raw = u16::from_le_bytes([row.3[column * 2], row.3[column * 2 + 1]]);
        assert_ne!(raw, MISSING_OFFSET);
        let unit = if row.4 { 4 } else { 1 };
        let start = usize::from(raw) * unit;
        let end = row
            .3
            .chunks_exact(2)
            .skip(column + 1)
            .find_map(|bytes| {
                let next = u16::from_le_bytes([bytes[0], bytes[1]]);
                (next != MISSING_OFFSET).then_some(usize::from(next) * unit)
            })
            .unwrap_or(row.2.len());
        &row.2[start..end]
    }

    fn decode_rows(payload: &[u8]) -> (storage::TileSnapshot, Rows) {
        let options = storage::DecodeOptions::new(
            payload.len().max(1),
            1 << 16,
            1 << 24,
            64,
            1 << 16,
            1 << 20,
        );
        let mut rows = Rows::default();
        let (tile, _) = storage::decode_tile_with_visitor(payload, options, &mut rows)
            .expect("strict new tile");
        (tile, rows)
    }

    fn unknown_99(payload: &[u8]) -> Vec<u8> {
        WireView::parse(payload)
            .expect("tile wire")
            .fields()
            .find(|field| field.number() == 99)
            .expect("unknown field")
            .raw()
            .to_vec()
    }

    #[test]
    fn new_tile_materializes_highest_local_row_and_preserves_unknown_fields() {
        let template = populated_template();
        let changes = [TileChange {
            row: 255,
            column: 3,
            change: BncChange::Set(ScalarInput::Boolean(true)),
        }];
        let outcome = rewrite_new_tile(&template, 4, &changes, limits()).expect("new tile");
        let payload = outcome.payload.as_deref().expect("replacement payload");
        let (tile, rows) = decode_rows(payload);

        assert_eq!(tile.max_column(), 0);
        assert_eq!(tile.max_row(), 0);
        assert_eq!(tile.num_cells(), 0);
        assert_eq!(tile.num_rows(), 1);
        assert_eq!(rows.values.len(), 1);
        let (row, cells, storage, offsets, wide) = &rows.values[0];
        assert_eq!((*row, *cells, *wide), (255, 1, false));
        assert_eq!(offsets.len(), 8);
        assert_eq!(&offsets[..6], &[0xff; 6]);
        assert_eq!(u16::from_le_bytes([offsets[6], offsets[7]]), 0);
        assert_eq!(
            classify_bnc_cell(Some(storage)).expect("BNC value"),
            CellValue::Boolean
        );
        assert_eq!(unknown_99(payload), unknown_99(&template));
        assert_eq!(
            outcome.final_rows,
            [RowCellCount {
                row: 255,
                cell_count: 1
            }]
        );
        assert_eq!(outcome.transitions.len(), 1);
        assert_eq!(outcome.report.rows_written, 1);
        assert_eq!(outcome.report.cell_slots_written, 1);
        assert_eq!(outcome.report.output_bytes, payload.len() as u64);
        assert_eq!(outcome.report.current_scratch_bytes, 0);
        assert_eq!(outcome.report.retained_elements, 3);
    }

    #[test]
    fn existing_empty_modern_tile_materializes_rows_in_place() {
        let source = empty_modern_tile();
        let changes = [
            TileChange {
                row: 0,
                column: 0,
                change: BncChange::Set(ScalarInput::Number(FiniteF64::new(121.0).expect("finite"))),
            },
            TileChange {
                row: 7,
                column: 2,
                change: BncChange::Set(ScalarInput::String(41)),
            },
        ];

        let positions = [
            TileReadPosition { row: 0, column: 0 },
            TileReadPosition { row: 7, column: 2 },
        ];
        let preclassified = preclassify_tile(&source, 3, &positions, limits())
            .expect("rowless tile preclassification");
        assert_eq!(preclassified.cells.len(), positions.len());
        assert!(preclassified.cells.iter().all(|cell| {
            cell.before == CellValue::Missing
                && cell.before_references == CellReferences::default()
                && !cell.present
        }));
        assert_eq!(preclassified.report.rows_read, 0);
        assert_eq!(preclassified.report.cell_slots_scanned, 0);
        assert_eq!(
            preclassified.report.retained_elements,
            positions.len() as u64
        );

        let mut cells_max_minus_one = limits();
        cells_max_minus_one.max_cells = positions.len() - 1;
        assert!(matches!(
            preclassify_tile(&source, 3, &positions, cells_max_minus_one),
            Err(TileError::LimitExceeded { observed, maximum })
                if observed == positions.len() as u64
                    && maximum == (positions.len() - 1) as u64
        ));

        let populated = populated_template();
        let missing_position = [TileReadPosition { row: 0, column: 0 }];
        assert!(matches!(
            preclassify_tile(&populated, 2, &missing_position, limits()),
            Err(TileError::NeedSparse { row: 0, column: 0 })
        ));
        assert!(matches!(
            rewrite_tile(TileRewriteRequest {
                source: &populated,
                columns: 2,
                changes: &[TileChange {
                    row: 0,
                    column: 0,
                    change: BncChange::Set(ScalarInput::Boolean(true)),
                }],
                limits: limits(),
            }),
            Err(TileError::NeedSparse { row: 0, column: 0 })
        ));

        let outcome = rewrite_tile(TileRewriteRequest {
            source: &source,
            columns: 3,
            changes: &changes,
            limits: limits(),
        })
        .expect("existing empty tile materializes");
        let payload = outcome.payload.as_deref().expect("replacement payload");
        let (tile, rows) = decode_rows(payload);

        assert_eq!(tile.num_rows(), 2);
        assert_eq!(
            rows.values
                .iter()
                .map(|(row, cells, ..)| (*row, *cells))
                .collect::<Vec<_>>(),
            [(0, 1), (7, 1)]
        );
        assert_eq!(
            classify_bnc_cell(Some(row_cell(&rows.values[0], 0))).expect("number"),
            CellValue::Number
        );
        assert_eq!(
            classify_bnc_cell(Some(row_cell(&rows.values[1], 2))).expect("text"),
            CellValue::Text(41)
        );
        assert_eq!(unknown_99(payload), unknown_99(&source));
        assert_eq!(
            outcome.final_rows,
            [
                RowCellCount {
                    row: 0,
                    cell_count: 1,
                },
                RowCellCount {
                    row: 7,
                    cell_count: 1,
                },
            ]
        );
        assert_eq!(outcome.transitions.len(), 2);
        assert_eq!(outcome.report.rows_written, 2);
    }

    #[test]
    fn new_tile_groups_multiple_rows_and_does_not_materialize_clear_only_row() {
        let template = populated_template();
        let changes = [
            TileChange {
                row: 0,
                column: 0,
                change: BncChange::Set(ScalarInput::Boolean(true)),
            },
            TileChange {
                row: 0,
                column: 4,
                change: BncChange::Set(ScalarInput::String(17)),
            },
            TileChange {
                row: 3,
                column: 1,
                change: BncChange::Clear,
            },
            TileChange {
                row: 7,
                column: 2,
                change: BncChange::Set(ScalarInput::Boolean(false)),
            },
        ];
        let outcome = rewrite_new_tile(&template, 5, &changes, limits()).expect("new tile");
        let payload = outcome.payload.as_deref().expect("replacement payload");
        let (tile, rows) = decode_rows(payload);

        assert_eq!(tile.num_rows(), 2);
        assert_eq!(
            rows.values
                .iter()
                .map(|(row, cells, ..)| (*row, *cells))
                .collect::<Vec<_>>(),
            [(0, 2), (7, 1)]
        );
        assert_eq!(
            outcome.final_rows,
            [
                RowCellCount {
                    row: 0,
                    cell_count: 2,
                },
                RowCellCount {
                    row: 3,
                    cell_count: 0,
                },
                RowCellCount {
                    row: 7,
                    cell_count: 1,
                },
            ]
        );
        assert_eq!(outcome.transitions.len(), 3);
        assert_eq!(outcome.report.rows_written, 2);
        assert_eq!(outcome.report.cell_slots_written, 3);
        assert_eq!(outcome.report.cell_slots_scanned, 10);
        assert_eq!(outcome.report.current_scratch_bytes, 0);
        assert!(outcome.report.allocation_events >= 9);
    }

    #[test]
    fn preclassification_streams_maximum_local_row_without_rewrite_allocations() {
        let template = populated_template();
        let source = rewrite_new_tile(
            &template,
            4,
            &[TileChange {
                row: 255,
                column: 3,
                change: BncChange::Set(ScalarInput::Boolean(true)),
            }],
            limits(),
        )
        .expect("new tile")
        .payload
        .expect("source payload");
        let positions = [
            TileReadPosition {
                row: 255,
                column: 0,
            },
            TileReadPosition {
                row: 255,
                column: 3,
            },
        ];
        let classified =
            preclassify_tile(&source, 4, &positions, limits()).expect("classification");

        assert_eq!(classified.cells.len(), 2);
        assert_eq!(classified.cells[0].before, CellValue::Missing);
        assert!(!classified.cells[0].present);
        assert_eq!(classified.cells[1].before, CellValue::Boolean);
        assert!(classified.cells[1].present);
        assert_eq!(
            classified.cells[1].before_references,
            CellReferences::default()
        );
        assert_eq!(classified.report.rows_read, 1);
        assert_eq!(classified.report.rows_written, 0);
        assert_eq!(classified.report.cell_slots_scanned, 4);
        assert_eq!(classified.report.cell_slots_written, 0);
        assert_eq!(classified.report.output_bytes, 0);
        assert_eq!(classified.report.retained_elements, 2);
        assert_eq!(classified.report.allocation_events, 1);
        assert_eq!(classified.report.current_scratch_bytes, 0);
        assert_eq!(
            classified.report.retained_bytes,
            (2 * size_of::<PreclassifiedCell>()) as u64
        );
        assert_eq!(
            classified.report.peak_scratch_bytes,
            classified.report.retained_bytes
        );

        let mut input_max_minus_one = limits();
        input_max_minus_one.max_input_bytes = source.len() - 1;
        assert!(matches!(
            preclassify_tile(&source, 4, &positions, input_max_minus_one),
            Err(TileError::LimitExceeded { observed, maximum })
                if observed == source.len() as u64 && maximum == (source.len() - 1) as u64
        ));
        let mut cells_max_minus_one = limits();
        cells_max_minus_one.max_cells = positions.len() - 1;
        assert!(matches!(
            preclassify_tile(&source, 4, &positions, cells_max_minus_one),
            Err(TileError::LimitExceeded { observed, maximum })
                if observed == positions.len() as u64
                    && maximum == (positions.len() - 1) as u64
        ));
    }

    #[test]
    fn preclassification_preserves_order_and_reference_kinds_across_rows() {
        let template = populated_template();
        let source = rewrite_new_tile(
            &template,
            5,
            &[
                TileChange {
                    row: 0,
                    column: 0,
                    change: BncChange::Set(ScalarInput::String(17)),
                },
                TileChange {
                    row: 7,
                    column: 2,
                    change: BncChange::Set(ScalarInput::RichText(29)),
                },
            ],
            limits(),
        )
        .expect("new tile")
        .payload
        .expect("source payload");
        let positions = [
            TileReadPosition { row: 0, column: 0 },
            TileReadPosition { row: 0, column: 4 },
            TileReadPosition { row: 7, column: 2 },
        ];
        let classified =
            preclassify_tile(&source, 5, &positions, limits()).expect("classification");

        assert_eq!(
            classified
                .cells
                .iter()
                .map(|cell| (cell.row, cell.column, cell.before, cell.present))
                .collect::<Vec<_>>(),
            [
                (0, 0, CellValue::Text(17), true),
                (0, 4, CellValue::Missing, false),
                (7, 2, CellValue::RichText(29), true),
            ]
        );
        assert_eq!(classified.cells[0].before_references.string, Some(17));
        assert_eq!(classified.cells[2].before_references.rich_text, Some(29));
        assert_eq!(classified.report.rows_read, 2);
        assert_eq!(classified.report.cell_slots_scanned, 10);
        assert_eq!(classified.report.allocation_events, 1);
    }

    #[test]
    fn grouped_scalar_and_formula_cache_changes_serialize_one_tile_once() {
        let mut a2 = BncCell::minimal();
        a2.set_number(120.0).unwrap();
        let a2 = a2.encode();
        let mut b2 = BncCell::minimal();
        b2.set_comment_identifier(Some(9));
        b2.set_number(323.0).unwrap();
        b2.set_formula_reference(17);
        let mut b2 = b2.encode();
        b2.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        let mut c2 = BncCell::minimal();
        c2.set_rich_text(29);
        let c2 = c2.encode();
        let source = one_row_tile(1, &[Some(a2), Some(b2.clone()), Some(c2)]);

        let cache_noop = [CacheChange {
            row: 1,
            column: 1,
            value: FormulaCachedValue::number(323.0).expect("finite"),
        }];
        let unchanged = rewrite_tile_with_cache(
            TileRewriteRequest {
                source: &source,
                columns: 3,
                changes: &[],
                limits: limits(),
            },
            &cache_noop,
        )
        .expect("equal cache");
        assert!(unchanged.payload.is_none());
        assert!(unchanged.transitions.is_empty());
        assert_eq!(unchanged.report.cache_cells_read, 1);
        assert_eq!(unchanged.report.cache_cells_written, 0);

        let scalar_changes = [TileChange {
            row: 1,
            column: 0,
            change: BncChange::Set(ScalarInput::Number(FiniteF64::new(121.0).expect("finite"))),
        }];
        let cache_changes = [CacheChange {
            row: 1,
            column: 1,
            value: FormulaCachedValue::number(324.0).expect("finite"),
        }];
        let outcome = rewrite_tile_with_cache(
            TileRewriteRequest {
                source: &source,
                columns: 3,
                changes: &scalar_changes,
                limits: limits(),
            },
            &cache_changes,
        )
        .expect("grouped tile rewrite");
        let payload = outcome.payload.as_deref().expect("one replacement");
        let (_, rows) = decode_rows(payload);
        let row = &rows.values[0];
        assert_eq!(
            BncCellView::parse(row_cell(row, 0))
                .unwrap()
                .cached_scalar(),
            Some(CachedScalar::Number(FiniteF64::new(121.0).unwrap()))
        );
        let cached_formula = BncCellView::parse(row_cell(row, 1)).unwrap();
        assert_eq!(cached_formula.stored_value(), StoredValue::Formula(17));
        assert_eq!(
            cached_formula.cached_scalar(),
            Some(CachedScalar::Number(FiniteF64::new(324.0).unwrap()))
        );
        assert_eq!(cached_formula.comment_identifier(), Some(9));
        assert!(row_cell(row, 1).ends_with(&[0xde, 0xad, 0xbe, 0xef]));
        assert_eq!(outcome.transitions.len(), 1);
        assert_eq!(outcome.report.rows_written, 1);
        assert_eq!(outcome.report.cell_slots_written, 2);
        assert_eq!(outcome.report.cache_cells_read, 1);
        assert_eq!(outcome.report.cache_cells_written, 1);
        assert_eq!(outcome.report.output_bytes, payload.len() as u64);
        assert_eq!(
            outcome.final_rows,
            [RowCellCount {
                row: 1,
                cell_count: 3,
            }]
        );

        let rich_only = [TileChange {
            row: 1,
            column: 2,
            change: BncChange::Set(ScalarInput::RichText(30)),
        }];
        let rich = rewrite_tile_with_cache(
            TileRewriteRequest {
                source: &source,
                columns: 3,
                changes: &rich_only,
                limits: limits(),
            },
            &[],
        )
        .expect("rich-only rewrite");
        let (_, rich_rows) = decode_rows(rich.payload.as_deref().expect("rich replacement"));
        assert_eq!(row_cell(&rich_rows.values[0], 1), b2);
    }

    #[test]
    fn formula_clear_retains_explicit_empty_slot_and_removes_formula_references() {
        let mut formula = BncCell::minimal();
        formula.set_number(50.0).expect("finite cache");
        formula.set_formula_reference(17);
        let formula = formula.encode();
        let source = one_row_tile(2, &[Some(formula.clone())]);
        let changes = [TileChange {
            row: 2,
            column: 0,
            change: BncChange::FormulaClear,
        }];

        let outcome = rewrite_tile_with_cache(
            TileRewriteRequest {
                source: &source,
                columns: 1,
                changes: &changes,
                limits: limits(),
            },
            &[],
        )
        .expect("formula clear");
        let payload = outcome
            .payload
            .as_deref()
            .expect("explicit empty replacement");
        let (_, rows) = decode_rows(payload);
        let cell = BncCellView::parse(row_cell(&rows.values[0], 0)).expect("empty BNC");
        assert_eq!(cell.stored_value(), StoredValue::Empty);
        assert_eq!(cell.formula_error_identifier(), None);
        assert_eq!(outcome.final_rows[0].cell_count, 1);
        assert_eq!(outcome.transitions.len(), 1);
        assert!(matches!(
            outcome.transitions[0].before,
            CellValue::Formula {
                identifier: 17,
                error: None
            }
        ));
        assert_eq!(outcome.transitions[0].after, CellValue::Empty);
        assert_eq!(
            outcome.transitions[0].after_references,
            CellReferences::default()
        );

        let mut max_minus_one = limits();
        max_minus_one.max_output_bytes = formula.len().saturating_add(63);
        assert!(matches!(
            rewrite_tile_with_cache(
                TileRewriteRequest {
                    source: &source,
                    columns: 1,
                    changes: &changes,
                    limits: max_minus_one,
                },
                &[],
            ),
            Err(TileError::LimitExceeded { .. })
        ));
    }

    #[test]
    fn populated_tile_appends_row_and_preserves_existing_clear_transition() {
        let mut formula = BncCell::minimal();
        formula.set_number(7.0).unwrap();
        formula.set_formula_reference(11);
        let source = one_row_tile(1, &[Some(formula.encode()), None]);
        let changes = [
            TileChange {
                row: 1,
                column: 0,
                change: BncChange::FormulaClear,
            },
            TileChange {
                row: 2,
                column: 1,
                change: BncChange::FormulaSet {
                    identifier: 12,
                    cache: Some(ScalarInput::Number(FiniteF64::new(9.0).unwrap())),
                },
            },
        ];
        let outcome = rewrite_tile(TileRewriteRequest {
            source: &source,
            columns: 2,
            changes: &changes,
            limits: limits(),
        })
        .expect("append row");
        let payload = outcome.payload.as_deref().expect("replacement");
        let (tile, rows) = decode_rows(payload);
        assert_eq!(tile.num_rows(), 2);
        assert_eq!(
            rows.values.iter().map(|row| row.0).collect::<Vec<_>>(),
            [1, 2]
        );
        assert_eq!(
            BncCellView::parse(row_cell(&rows.values[0], 0))
                .unwrap()
                .stored_value(),
            StoredValue::Empty
        );
        assert!(matches!(
            BncCellView::parse(row_cell(&rows.values[1], 1))
                .unwrap()
                .stored_value(),
            StoredValue::Formula(12)
        ));
        assert_eq!(outcome.transitions.len(), 2);

        for axis in 0..2 {
            let mut refused = limits();
            match axis {
                0 => refused.max_input_bytes = (outcome.report.wire_bytes - 1) as usize,
                1 => refused.max_fields = (outcome.report.wire_fields - 1) as usize,
                _ => unreachable!(),
            }
            assert!(matches!(
                rewrite_tile(TileRewriteRequest {
                    source: &source,
                    columns: 2,
                    changes: &changes,
                    limits: refused,
                }),
                Err(TileError::LimitExceeded { .. }) | Err(TileError::InvalidSource)
            ));
        }
    }

    #[test]
    fn prepared_populated_append_is_output_free_and_refuses_every_axis_before_allocation() {
        let mut formula = BncCell::minimal();
        formula.set_number(7.0).unwrap();
        formula.set_formula_reference(11);
        let source = one_row_tile(1, &[Some(formula.encode()), None]);
        let changes = [
            TileChange {
                row: 1,
                column: 0,
                change: BncChange::FormulaClear,
            },
            TileChange {
                row: 2,
                column: 0,
                change: BncChange::Clear,
            },
            TileChange {
                row: 2,
                column: 1,
                change: BncChange::FormulaSet {
                    identifier: 12,
                    cache: Some(ScalarInput::Number(FiniteF64::new(9.0).unwrap())),
                },
            },
        ];

        let prepare = || {
            prepare_tile(TileRewriteRequest {
                source: &source,
                columns: 2,
                changes: &changes,
                limits: limits(),
            })
            .expect("output-free populated append plan")
        };
        let prepared = prepare();
        let prepare_report = prepared.prepare_report();
        let requirements = prepared.execution_requirements();
        assert_eq!(prepare_report.output_bytes(), 0);
        assert_eq!(prepare_report.report().output_bytes, 0);
        assert_eq!(prepare_report.report().current_scratch_bytes, 0);
        assert!(prepare_report.report().retained_bytes > 0);
        assert!(requirements.input_bytes() > 0);
        assert!(requirements.fields() > 0);
        assert!(requirements.work() > 0);
        assert!(requirements.output_bytes() > 0);
        assert!(requirements.retained_bytes() > 0);
        assert!(requirements.retained_elements() > 0);
        assert!(requirements.peak_scratch_bytes() > 0);
        assert!(requirements.allocations() > 0);

        reset_prepared_execution_allocations();
        let outcome = prepared
            .execute(requirements.exact_limits())
            .expect("exact prepared execution");
        assert!(prepared_execution_allocations() > 0);
        let payload = outcome.payload.as_deref().expect("prepared payload");
        let (tile, rows) = decode_rows(payload);
        assert_eq!(tile.num_rows(), 2);
        assert_eq!(
            rows.values.iter().map(|row| row.0).collect::<Vec<_>>(),
            [1, 2]
        );
        assert_eq!(outcome.transitions.len(), 2);

        for axis in 0..8 {
            let prepared = prepare();
            let requirements = prepared.execution_requirements();
            let mut refused = requirements.exact_limits();
            match axis {
                0 => refused.max_input_bytes -= 1,
                1 => refused.max_fields -= 1,
                2 => refused.max_work -= 1,
                3 => refused.max_output_bytes -= 1,
                4 => refused.max_retained_bytes -= 1,
                5 => refused.max_retained_elements -= 1,
                6 => refused.max_peak_scratch_bytes -= 1,
                7 => refused.max_allocations -= 1,
                _ => unreachable!(),
            }
            reset_prepared_execution_allocations();
            assert!(matches!(
                prepared.execute(refused),
                Err(TileError::LimitExceeded { .. })
            ));
            assert_eq!(prepared_execution_allocations(), 0);
        }
    }

    #[test]
    fn complete_formula_scan_retains_text_cache_keys() {
        let mut formula = BncCell::minimal();
        formula.set_string(27);
        formula.set_formula_reference(11);
        let source = one_row_tile(2, &[Some(formula.encode())]);

        let scan = scan_formula_cells(&source, 1, limits()).expect("formula scan");

        assert_eq!(scan.cells.len(), 1);
        assert_eq!(scan.cells[0].row, 2);
        assert_eq!(scan.cells[0].column, 0);
        assert_eq!(scan.cells[0].identifier, 11);
        assert!(matches!(
            scan.cells[0].cache,
            Some(FormulaCacheValue::TextKey(27))
        ));
    }
}
