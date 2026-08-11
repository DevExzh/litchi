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

use litchi_iwa_common::formula::FormulaCachedValue;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;

use crate::cell::{
    FiniteF64,
    wire::{
        BncCell, BncCellView, CachedScalar, ClearValue, Error as BncError, ScalarValue, StoredValue,
    },
};

const MISSING_OFFSET: u16 = u16::MAX;
const BNC_STORAGE_VERSION: u32 = 5;

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
        }
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct PreclassifiedCell {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) before: CellValue,
    pub(crate) before_references: CellReferences,
    pub(crate) present: bool,
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

/// Tile-local result.  A no-op retains no replacement payload at all.
#[derive(Debug)]
pub(crate) struct TileRewriteOutcome {
    pub(crate) payload: Option<Vec<u8>>,
    pub(crate) transitions: Vec<CellTransition>,
    /// Sorted exact counts for every row named by the request.
    pub(crate) final_rows: Vec<RowCellCount>,
    pub(crate) report: TileReport,
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
    let (before, before_references) = classify_bnc_with_references(bytes)?;
    if let Some(bytes) = bytes {
        counters.charge(u64_from_usize(bytes.len())?)?;
    }
    cells.push(PreclassifiedCell {
        row: position.row,
        column: position.column,
        before,
        before_references,
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
pub(crate) fn rewrite_tile(request: TileRewriteRequest<'_, '_>) -> Result<TileRewriteOutcome> {
    if request.changes.is_empty() {
        return rewrite_tile_changes(request);
    }
    validate_request(request)?;
    let codec_options = storage::DecodeOptions::new(
        request.limits.max_input_bytes,
        request.limits.max_fields,
        usize_from_u64(request.limits.max_work)?,
        64,
        request.limits.max_cells,
        request.limits.max_output_bytes,
    );
    let (snapshot, _) = storage::decode_tile_with_report(request.source, codec_options)
        .map_err(|_| TileError::InvalidSource)?;
    if snapshot.storage_version() == Some(BNC_STORAGE_VERSION)
        && snapshot.last_saved_in_bnc() == Some(true)
        && tile_has_no_rows(request.source)?
    {
        return rewrite_new_tile(
            request.source,
            request.columns,
            request.changes,
            request.limits,
        );
    }
    rewrite_tile_changes(request)
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
pub(crate) fn rewrite_tile_with_cache(
    request: TileRewriteRequest<'_, '_>,
    cache_changes: &[CacheChange],
) -> Result<TileRewriteOutcome> {
    if cache_changes.is_empty() {
        return rewrite_tile(request);
    }
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
    let merged_bytes = merged_len
        .checked_mul(size_of::<TileChange>())
        .ok_or(TileError::InvalidSource)?;
    let mut counters = Counters::new(request.limits);
    counters.preflight_scratch_allocation(merged_bytes)?;
    let mut merged = Vec::new();
    merged
        .try_reserve_exact(merged_len)
        .map_err(|_| TileError::Allocation { amount: merged_len })?;
    let merged_capacity = merged
        .capacity()
        .checked_mul(size_of::<TileChange>())
        .ok_or(TileError::InvalidSource)?;
    counters.record_scratch_allocation(merged_capacity)?;

    let mut scalar = 0usize;
    let mut cache = 0usize;
    while scalar < request.changes.len() || cache < cache_changes.len() {
        let scalar_key = request
            .changes
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
                merged.push(request.changes[scalar]);
                scalar += 1;
            },
            (Some(_), None) => {
                merged.push(request.changes[scalar]);
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
    if merged.len() != merged_len {
        return Err(TileError::InvalidSource);
    }

    let live_prefix = counters.report.current_scratch_bytes;
    let outcome = rewrite_tile(TileRewriteRequest {
        source: request.source,
        columns: request.columns,
        changes: &merged,
        limits: request.limits,
    })?;
    let prefix = counters.report;
    counters.release_scratch(merged_capacity)?;
    let mut report = merge_reports(prefix, outcome.report, live_prefix)?;
    report.current_scratch_bytes = 0;
    Ok(TileRewriteOutcome { report, ..outcome })
}

fn rewrite_tile_changes(request: TileRewriteRequest<'_, '_>) -> Result<TileRewriteOutcome> {
    // The package entry point has already resolved the selected table.  An
    // empty edit must not force an otherwise irrelevant raw-tile traversal.
    if request.changes.is_empty() {
        return Ok(TileRewriteOutcome {
            payload: None,
            transitions: Vec::new(),
            final_rows: Vec::new(),
            report: TileReport::default(),
        });
    }
    validate_request(request)?;
    let mut counters = Counters::new(request.limits);
    counters.charge_bytes(request.source.len())?;

    let codec_options = storage::DecodeOptions::new(
        request.limits.max_input_bytes,
        request.limits.max_fields,
        usize_from_u64(request.limits.max_work)?,
        64,
        request.limits.max_cells,
        request.limits.max_output_bytes,
    );
    let (tile_snapshot, codec_report) =
        storage::decode_tile_with_report(request.source, codec_options)
            .map_err(|_| TileError::InvalidSource)?;
    counters.report.wire_bytes = u64_from_usize(codec_report.source_bytes())?;
    counters.report.wire_fields = u64_from_usize(codec_report.fields())?;
    counters.charge(u64_from_usize(codec_report.work_bytes())?)?;

    let tile_view = WireView::parse(request.source).map_err(|_| TileError::InvalidSource)?;
    let mut raw_rows = Vec::new();
    // Strict target admission has established that modern `num_rows` equals
    // the number of repeated row records, so it is a useful bounded capacity
    // hint but never authority for the records parsed below.
    let row_capacity = usize::try_from(tile_snapshot.num_rows())
        .map_err(|_| TileError::InvalidSource)?
        .min(request.limits.max_rows);
    raw_rows
        .try_reserve_exact(row_capacity)
        .map_err(|_| TileError::Allocation {
            amount: row_capacity,
        })?;
    for (field_index, field) in tile_view.fields().enumerate() {
        counters.charge(1)?;
        if field.number() != 5 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(TileError::InvalidSource);
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
        if raw_rows.len() == request.limits.max_rows {
            return Err(TileError::LimitExceeded {
                observed: u64_from_usize(
                    raw_rows
                        .len()
                        .checked_add(1)
                        .ok_or(TileError::InvalidSource)?,
                )?,
                maximum: u64_from_usize(request.limits.max_rows)?,
            });
        }
        raw_rows.push(RawRow {
            field_index,
            row,
            payload,
        });
    }
    if raw_rows.iter().any(|raw| {
        usize::try_from(raw.row.tile_row_index()).map_or(true, |row| row >= request.limits.max_rows)
    }) {
        return Err(TileError::InvalidSource);
    }

    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(raw_rows.len().min(request.changes.len()))
        .map_err(|_| TileError::Allocation {
            amount: raw_rows.len().min(request.changes.len()),
        })?;
    let mut transitions = Vec::new();
    let distinct_rows = distinct_change_rows(request.changes);
    let row_count_bytes = distinct_rows
        .checked_mul(size_of::<RowCellCount>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(row_count_bytes)?;
    let mut final_rows = Vec::new();
    final_rows
        .try_reserve_exact(distinct_rows)
        .map_err(|_| TileError::Allocation {
            amount: distinct_rows,
        })?;
    counters.record_allocation_event()?;

    let mut change_start = 0usize;
    let mut previous_row = None;
    for raw in &raw_rows {
        if previous_row.is_some_and(|previous| previous >= raw.row.tile_row_index()) {
            return Err(TileError::InvalidSource);
        }
        previous_row = Some(raw.row.tile_row_index());
        if change_start < request.changes.len()
            && request.changes[change_start].row < raw.row.tile_row_index()
        {
            let change = request.changes[change_start];
            return Err(TileError::NeedSparse {
                row: change.row,
                column: change.column,
            });
        }
        let change_end = request.changes[change_start..]
            .iter()
            .position(|change| change.row != raw.row.tile_row_index())
            .map_or(request.changes.len(), |offset| change_start + offset);
        if change_start != change_end {
            let row = rewrite_row(
                raw.row,
                raw.payload,
                request.columns,
                &request.changes[change_start..change_end],
                request.changes.len(),
                request.limits.max_output_bytes,
                &mut counters,
                &mut transitions,
            )?;
            final_rows.push(RowCellCount {
                row: raw.row.tile_row_index(),
                cell_count: row.cell_count,
            });
            if let Some(payload) = row.payload {
                replacements.push(RowReplacement {
                    field_index: raw.field_index,
                    payload,
                });
                counters.report.rows_written = counters
                    .report
                    .rows_written
                    .checked_add(1)
                    .ok_or(TileError::InvalidSource)?;
            }
        }
        change_start = change_end;
    }
    if let Some(change) = request.changes.get(change_start) {
        return Err(TileError::NeedSparse {
            row: change.row,
            column: change.column,
        });
    }

    if replacements.is_empty() {
        retain_outcome_vectors(&mut counters, &transitions, &final_rows)?;
        return Ok(TileRewriteOutcome {
            payload: None,
            transitions,
            final_rows,
            report: counters.report,
        });
    }

    let output_len = rewritten_tile_length(request.source, &replacements)?;
    if output_len > request.limits.max_output_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(output_len)?,
            maximum: u64_from_usize(request.limits.max_output_bytes)?,
        });
    }
    // Charge the entire publish artifact before its one allocation.
    counters.charge(u64_from_usize(output_len)?)?;
    let mut payload = Vec::new();
    counters.preflight_scratch_allocation(output_len)?;
    payload
        .try_reserve_exact(output_len)
        .map_err(|_| TileError::Allocation { amount: output_len })?;
    counters.record_scratch_allocation(payload.capacity())?;
    let mut replacement = 0usize;
    for (field_index, field) in tile_view.fields().enumerate() {
        if replacements
            .get(replacement)
            .is_some_and(|next| next.field_index == field_index)
        {
            let next = &replacements[replacement];
            payload.extend_from_slice(field.key());
            encode_varint(
                &mut payload,
                u64::try_from(next.payload.len()).map_err(|_| TileError::InvalidSource)?,
            );
            payload.extend_from_slice(&next.payload);
            replacement += 1;
        } else {
            payload.extend_from_slice(field.raw());
        }
    }
    if replacement != replacements.len() || payload.len() != output_len {
        return Err(TileError::InvalidSource);
    }
    counters.report.output_bytes = u64_from_usize(payload.len())?;
    counters.retain(payload.capacity(), 1)?;
    counters.release_scratch(payload.capacity())?;
    for row in &replacements {
        counters.release_scratch(row.payload.capacity())?;
    }
    drop(replacements);
    retain_outcome_vectors(&mut counters, &transitions, &final_rows)?;
    Ok(TileRewriteOutcome {
        payload: Some(payload),
        transitions,
        final_rows,
        report: counters.report,
    })
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
    let request = TileRewriteRequest {
        source: template,
        columns,
        changes,
        limits,
    };
    validate_request(request)?;
    if changes.len() > limits.max_cells {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(changes.len())?,
            maximum: u64_from_usize(limits.max_cells)?,
        });
    }

    let mut counters = Counters::new(limits);
    counters.charge_bytes(template.len())?;
    let codec_options = storage::DecodeOptions::new(
        limits.max_input_bytes,
        limits.max_fields,
        usize_from_u64(limits.max_work)?,
        64,
        limits.max_cells,
        limits.max_output_bytes,
    );
    let (snapshot, codec_report) = storage::decode_tile_with_report(template, codec_options)
        .map_err(|_| TileError::InvalidSource)?;
    counters.report.wire_bytes = u64_from_usize(codec_report.source_bytes())?;
    counters.report.wire_fields = u64_from_usize(codec_report.fields())?;
    counters.charge(u64_from_usize(codec_report.work_bytes())?)?;
    if snapshot.storage_version() != Some(BNC_STORAGE_VERSION)
        || snapshot.last_saved_in_bnc() != Some(true)
    {
        return Err(TileError::InvalidSource);
    }

    let distinct_rows = distinct_change_rows(changes);
    if distinct_rows > limits.max_rows {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(distinct_rows)?,
            maximum: u64_from_usize(limits.max_rows)?,
        });
    }
    let mut effective_rows = 0usize;
    let mut effective_changes = 0usize;
    let mut start = 0usize;
    while start < changes.len() {
        let row = changes[start].row;
        if usize::try_from(row).map_or(true, |row| row >= limits.max_rows) {
            return Err(TileError::OutOfBounds {
                row,
                column: changes[start].column,
            });
        }
        let end = changes[start..]
            .iter()
            .position(|change| change.row != row)
            .map_or(changes.len(), |offset| start + offset);
        let sets = changes[start..end]
            .iter()
            .filter(|change| matches!(change.change, BncChange::Set(_)))
            .count();
        if sets != 0 {
            effective_rows = effective_rows
                .checked_add(1)
                .ok_or(TileError::InvalidSource)?;
            effective_changes = effective_changes
                .checked_add(sets)
                .ok_or(TileError::InvalidSource)?;
        }
        start = end;
    }
    let effective_rows_observed = u64_from_usize(effective_rows)?;
    let effective_rows = u32::try_from(effective_rows).map_err(|_| TileError::LimitExceeded {
        observed: effective_rows_observed,
        maximum: u64::from(u32::MAX),
    })?;

    // Preserve a complete row-count result even when some requested clears
    // disappear before the ordinary tile rewrite.
    let final_row_bytes = distinct_rows
        .checked_mul(size_of::<RowCellCount>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(final_row_bytes)?;
    let mut requested_rows = Vec::new();
    requested_rows
        .try_reserve_exact(distinct_rows)
        .map_err(|_| TileError::Allocation {
            amount: distinct_rows,
        })?;
    let requested_row_bytes = requested_rows
        .capacity()
        .checked_mul(size_of::<RowCellCount>())
        .ok_or(TileError::InvalidSource)?;
    counters.record_scratch_allocation(requested_row_bytes)?;
    for change in changes {
        if requested_rows
            .last()
            .is_none_or(|previous: &RowCellCount| previous.row != change.row)
        {
            requested_rows.push(RowCellCount {
                row: change.row,
                cell_count: 0,
            });
        }
    }

    if effective_changes == 0 {
        counters.release_scratch(requested_row_bytes)?;
        counters.retain(requested_row_bytes, requested_rows.len())?;
        return Ok(TileRewriteOutcome {
            payload: None,
            transitions: Vec::new(),
            final_rows: requested_rows,
            report: counters.report,
        });
    }

    let effective_change_bytes = effective_changes
        .checked_mul(size_of::<TileChange>())
        .ok_or(TileError::InvalidSource)?;
    counters.preflight_scratch_allocation(effective_change_bytes)?;
    let mut selected_changes = Vec::new();
    selected_changes
        .try_reserve_exact(effective_changes)
        .map_err(|_| TileError::Allocation {
            amount: effective_changes,
        })?;
    let selected_change_bytes = selected_changes
        .capacity()
        .checked_mul(size_of::<TileChange>())
        .ok_or(TileError::InvalidSource)?;
    counters.record_scratch_allocation(selected_change_bytes)?;
    selected_changes.extend(
        changes
            .iter()
            .copied()
            .filter(|change| matches!(change.change, BncChange::Set(_))),
    );

    let scaffold_len =
        new_tile_scaffold_length(template, columns, effective_rows, &selected_changes)?;
    if scaffold_len > limits.max_output_bytes {
        return Err(TileError::LimitExceeded {
            observed: u64_from_usize(scaffold_len)?,
            maximum: u64_from_usize(limits.max_output_bytes)?,
        });
    }
    counters.charge(u64_from_usize(scaffold_len)?)?;
    counters.preflight_scratch_allocation(scaffold_len)?;
    let mut scaffold = Vec::new();
    scaffold
        .try_reserve_exact(scaffold_len)
        .map_err(|_| TileError::Allocation {
            amount: scaffold_len,
        })?;
    counters.record_scratch_allocation(scaffold.capacity())?;
    write_new_tile_scaffold(
        &mut scaffold,
        template,
        columns,
        effective_rows,
        &selected_changes,
        &mut counters,
    )?;
    if scaffold.len() != scaffold_len {
        return Err(TileError::InvalidSource);
    }

    let live_prefix = counters.report.current_scratch_bytes;
    let remaining_work =
        limits
            .max_work
            .checked_sub(counters.work)
            .ok_or(TileError::LimitExceeded {
                observed: counters.work,
                maximum: limits.max_work,
            })?;
    let inner_limits = TileLimits::new(
        scaffold_len,
        limits.max_output_bytes,
        limits.max_fields,
        remaining_work,
        limits.max_rows,
        limits.max_cells,
    );
    let mut outcome = rewrite_tile(TileRewriteRequest {
        source: &scaffold,
        columns,
        changes: &selected_changes,
        limits: inner_limits,
    })?;

    let inner_row_bytes = outcome
        .final_rows
        .capacity()
        .checked_mul(size_of::<RowCellCount>())
        .ok_or(TileError::InvalidSource)?;
    outcome.report.retained_bytes = outcome
        .report
        .retained_bytes
        .checked_sub(u64_from_usize(inner_row_bytes)?)
        .ok_or(TileError::InvalidSource)?;
    outcome.report.retained_elements = outcome
        .report
        .retained_elements
        .checked_sub(u64_from_usize(outcome.final_rows.len())?)
        .ok_or(TileError::InvalidSource)?;
    for final_row in &outcome.final_rows {
        let requested = requested_rows
            .binary_search_by_key(&final_row.row, |candidate| candidate.row)
            .map_err(|_| TileError::InvalidSource)?;
        requested_rows[requested].cell_count = final_row.cell_count;
    }
    drop(outcome.final_rows);

    let prefix = counters.report;
    counters.release_scratch(scaffold.capacity())?;
    counters.release_scratch(selected_change_bytes)?;
    counters.release_scratch(requested_row_bytes)?;
    let mut report = merge_reports(prefix, outcome.report, live_prefix)?;
    report.current_scratch_bytes = 0;
    report.retained_bytes = report
        .retained_bytes
        .checked_add(u64_from_usize(requested_row_bytes)?)
        .ok_or(TileError::InvalidSource)?;
    report.retained_elements = report
        .retained_elements
        .checked_add(u64_from_usize(requested_rows.len())?)
        .ok_or(TileError::InvalidSource)?;
    outcome.final_rows = requested_rows;
    outcome.report = report;
    Ok(outcome)
}

fn new_tile_scaffold_length(
    template: &[u8],
    columns: u32,
    row_count: u32,
    changes: &[TileChange],
) -> Result<usize> {
    let view = WireView::parse(template).map_err(|_| TileError::InvalidSource)?;
    let row_count = u64::from(row_count);
    let mut length = 0usize;
    for field in view.fields() {
        let field_len = match field.number() {
            1..=3 => field.key().len() + 1,
            4 => field.key().len() + varint_len(row_count),
            5 => 0,
            _ => field.raw().len(),
        };
        length = length
            .checked_add(field_len)
            .ok_or(TileError::InvalidSource)?;
    }
    let offsets_len = usize::try_from(columns)
        .map_err(|_| TileError::InvalidSource)?
        .checked_mul(2)
        .ok_or(TileError::InvalidSource)?;
    let mut previous_row = None;
    for change in changes {
        if previous_row == Some(change.row) {
            continue;
        }
        previous_row = Some(change.row);
        let row_len = canonical_empty_row_length(change.row, offsets_len)?;
        length = length
            .checked_add(1)
            .and_then(|value| value.checked_add(varint_len(u64::try_from(row_len).ok()?)))
            .and_then(|value| value.checked_add(row_len))
            .ok_or(TileError::InvalidSource)?;
    }
    Ok(length)
}

fn canonical_empty_row_length(row: u32, offsets_len: usize) -> Result<usize> {
    // Required f1-f4, modern BNC f5-f8, and the complete missing-slot table.
    14usize
        .checked_add(varint_len(u64::from(row)))
        .and_then(|value| value.checked_add(varint_len(u64::try_from(offsets_len).ok()?)))
        .and_then(|value| value.checked_add(offsets_len))
        .ok_or(TileError::InvalidSource)
}

fn write_new_tile_scaffold(
    output: &mut Vec<u8>,
    template: &[u8],
    columns: u32,
    row_count: u32,
    changes: &[TileChange],
    counters: &mut Counters,
) -> Result<()> {
    let view = WireView::parse(template).map_err(|_| TileError::InvalidSource)?;
    for field in view.fields() {
        counters.charge(1)?;
        match field.number() {
            1..=3 => {
                output.extend_from_slice(field.key());
                output.push(0);
            },
            4 => {
                output.extend_from_slice(field.key());
                encode_varint(output, u64::from(row_count));
            },
            5 => {},
            _ => output.extend_from_slice(field.raw()),
        }
    }

    let columns = usize::try_from(columns).map_err(|_| TileError::InvalidSource)?;
    let offsets_len = columns.checked_mul(2).ok_or(TileError::InvalidSource)?;
    let mut previous_row = None;
    for change in changes {
        if previous_row == Some(change.row) {
            continue;
        }
        previous_row = Some(change.row);
        let row_len = canonical_empty_row_length(change.row, offsets_len)?;
        output.push(0x2a);
        encode_varint(output, u64_from_usize(row_len)?);
        output.push(0x08);
        encode_varint(output, u64::from(change.row));
        output.extend_from_slice(&[0x10, 0, 0x1a, 0, 0x22, 0, 0x28, 5, 0x32, 0]);
        output.push(0x3a);
        encode_varint(output, u64_from_usize(offsets_len)?);
        for _ in 0..columns {
            output.extend_from_slice(&MISSING_OFFSET.to_le_bytes());
        }
        output.extend_from_slice(&[0x40, 0]);
    }
    Ok(())
}

fn retain_outcome_vectors(
    counters: &mut Counters,
    transitions: &Vec<CellTransition>,
    final_rows: &Vec<RowCellCount>,
) -> Result<()> {
    counters.retain(
        transitions
            .capacity()
            .checked_mul(size_of::<CellTransition>())
            .ok_or(TileError::InvalidSource)?,
        transitions.len(),
    )?;
    counters.retain(
        final_rows
            .capacity()
            .checked_mul(size_of::<RowCellCount>())
            .ok_or(TileError::InvalidSource)?,
        final_rows.len(),
    )
}

fn merge_reports(prefix: TileReport, suffix: TileReport, live_prefix: u64) -> Result<TileReport> {
    Ok(TileReport {
        wire_bytes: prefix
            .wire_bytes
            .checked_add(suffix.wire_bytes)
            .ok_or(TileError::InvalidSource)?,
        wire_fields: prefix
            .wire_fields
            .checked_add(suffix.wire_fields)
            .ok_or(TileError::InvalidSource)?,
        wire_work: prefix
            .wire_work
            .checked_add(suffix.wire_work)
            .ok_or(TileError::InvalidSource)?,
        rows_read: prefix
            .rows_read
            .checked_add(suffix.rows_read)
            .ok_or(TileError::InvalidSource)?,
        rows_written: prefix
            .rows_written
            .checked_add(suffix.rows_written)
            .ok_or(TileError::InvalidSource)?,
        cell_slots_scanned: prefix
            .cell_slots_scanned
            .checked_add(suffix.cell_slots_scanned)
            .ok_or(TileError::InvalidSource)?,
        cell_slots_written: prefix
            .cell_slots_written
            .checked_add(suffix.cell_slots_written)
            .ok_or(TileError::InvalidSource)?,
        cache_cells_read: prefix
            .cache_cells_read
            .checked_add(suffix.cache_cells_read)
            .ok_or(TileError::InvalidSource)?,
        cache_cells_written: prefix
            .cache_cells_written
            .checked_add(suffix.cache_cells_written)
            .ok_or(TileError::InvalidSource)?,
        output_bytes: suffix.output_bytes,
        retained_elements: prefix
            .retained_elements
            .checked_add(suffix.retained_elements)
            .ok_or(TileError::InvalidSource)?,
        retained_bytes: prefix
            .retained_bytes
            .checked_add(suffix.retained_bytes)
            .ok_or(TileError::InvalidSource)?,
        current_scratch_bytes: prefix
            .current_scratch_bytes
            .checked_add(suffix.current_scratch_bytes)
            .ok_or(TileError::InvalidSource)?,
        peak_scratch_bytes: prefix.peak_scratch_bytes.max(
            live_prefix
                .checked_add(suffix.peak_scratch_bytes)
                .ok_or(TileError::InvalidSource)?,
        ),
        allocation_events: prefix
            .allocation_events
            .checked_add(suffix.allocation_events)
            .ok_or(TileError::InvalidSource)?,
    })
}

fn distinct_change_rows(changes: &[TileChange]) -> usize {
    changes.first().map_or(0, |_| {
        1 + changes
            .windows(2)
            .filter(|pair| pair[0].row != pair[1].row)
            .count()
    })
}

#[derive(Clone, Copy)]
struct RawRow<'source> {
    field_index: usize,
    row: storage::TileRowInfoSnapshot<'source>,
    payload: &'source [u8],
}

struct RowReplacement {
    field_index: usize,
    payload: Vec<u8>,
}

struct RowRewrite {
    payload: Option<Vec<u8>>,
    cell_count: u32,
}

fn rewrite_row(
    row: storage::TileRowInfoSnapshot<'_>,
    raw: &[u8],
    columns: u32,
    changes: &[TileChange],
    transition_capacity: usize,
    max_output_bytes: usize,
    counters: &mut Counters,
    transitions: &mut Vec<CellTransition>,
) -> Result<RowRewrite> {
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
    let mut slots = parse_slots(
        storage_buffer,
        offsets,
        row.has_wide_offsets().unwrap_or(false),
        slot_count,
        counters,
    )?;
    let declared_cells = usize::try_from(row.cell_count()).map_err(|_| TileError::InvalidSource)?;
    if slots.iter().filter(|slot| slot.bytes().is_some()).count() != declared_cells {
        return Err(TileError::InvalidSource);
    }

    let mut changed = false;
    for change in changes {
        let column = usize::try_from(change.column).map_err(|_| TileError::InvalidSource)?;
        if column >= requested_columns || column >= slots.len() {
            return Err(TileError::OutOfBounds {
                row: change.row,
                column: change.column,
            });
        }
        let before_bytes = slots[column].bytes();
        let cache_change = matches!(change.change, BncChange::FormulaCache(_));
        if cache_change {
            counters.report.cache_cells_read = counters
                .report
                .cache_cells_read
                .checked_add(1)
                .ok_or(TileError::InvalidSource)?;
        }
        if cache_change && !matches!(classify_bnc_cell(before_bytes)?, CellValue::Formula { .. }) {
            return Err(TileError::UnsupportedSource {
                row: change.row,
                column: change.column,
            });
        }
        if bnc_change_is_noop(before_bytes, change.change)? {
            continue;
        }
        let before = classify_bnc_cell(before_bytes)?;
        if let CellValue::Unsupported(kind) = before {
            return Err(TileError::UnsupportedValue {
                row: change.row,
                column: change.column,
                kind,
            });
        }
        let before_references = bnc_references(before_bytes)?;
        let mutation = mutate_cell(before_bytes, change.change, max_output_bytes)?;
        let record_transition = !matches!(change.change, BncChange::FormulaCache(_));
        let after_bytes = match &mutation {
            CellMutation::Unchanged => before_bytes,
            CellMutation::Delete => None,
            CellMutation::Replace(bytes) => Some(bytes.as_slice()),
        };
        let after = classify_bnc_cell(after_bytes)?;
        let after_references = bnc_references(after_bytes)?;
        match mutation {
            CellMutation::Unchanged => {},
            CellMutation::Delete => {
                slots[column] = Slot::Missing;
                changed = true;
                counters.report.cell_slots_written = counters
                    .report
                    .cell_slots_written
                    .checked_add(1)
                    .ok_or(TileError::InvalidSource)?;
                if cache_change {
                    counters.report.cache_cells_written = counters
                        .report
                        .cache_cells_written
                        .checked_add(1)
                        .ok_or(TileError::InvalidSource)?;
                }
                if record_transition {
                    push_transition(
                        transitions,
                        transition_capacity,
                        counters,
                        CellTransition {
                            row: change.row,
                            column: change.column,
                            before,
                            after,
                            before_references,
                            after_references,
                        },
                    )?;
                }
            },
            CellMutation::Replace(bytes) if before_bytes != Some(bytes.as_slice()) => {
                slots[column] = Slot::Owned(bytes);
                changed = true;
                counters.report.cell_slots_written = counters
                    .report
                    .cell_slots_written
                    .checked_add(1)
                    .ok_or(TileError::InvalidSource)?;
                if cache_change {
                    counters.report.cache_cells_written = counters
                        .report
                        .cache_cells_written
                        .checked_add(1)
                        .ok_or(TileError::InvalidSource)?;
                }
                if record_transition {
                    push_transition(
                        transitions,
                        transition_capacity,
                        counters,
                        CellTransition {
                            row: change.row,
                            column: change.column,
                            before,
                            after,
                            before_references,
                            after_references,
                        },
                    )?;
                }
            },
            CellMutation::Replace(_) => {},
        }
    }
    if !changed {
        return Ok(RowRewrite {
            payload: None,
            cell_count: row.cell_count(),
        });
    }

    let prefer_wide = row.has_wide_offsets().unwrap_or(false);
    let (buffer, new_offsets, wide) = encode_slots(&slots, prefer_wide, counters)?;
    let buffer_capacity = buffer.capacity();
    let offsets_capacity = new_offsets.capacity();
    let cell_count = u32::try_from(slots.iter().filter(|slot| slot.bytes().is_some()).count())
        .map_err(|_| TileError::InvalidSource)?;
    let result = rewrite_row_message(
        raw,
        cell_count,
        &buffer,
        &new_offsets,
        row.has_wide_offsets(),
        wide,
        counters,
    );
    counters.release_scratch(buffer_capacity)?;
    counters.release_scratch(offsets_capacity)?;
    drop(buffer);
    drop(new_offsets);
    result.map(|payload| RowRewrite {
        payload: Some(payload),
        cell_count,
    })
}

enum Slot<'source> {
    Missing,
    Borrowed(&'source [u8]),
    Owned(Vec<u8>),
}

fn push_transition(
    transitions: &mut Vec<CellTransition>,
    total_changes: usize,
    counters: &mut Counters,
    transition: CellTransition,
) -> Result<()> {
    if transitions.is_empty() {
        let requested = total_changes
            .checked_mul(size_of::<CellTransition>())
            .ok_or(TileError::InvalidSource)?;
        counters.preflight_scratch_allocation(requested)?;
        transitions
            .try_reserve_exact(total_changes)
            .map_err(|_| TileError::Allocation {
                amount: total_changes,
            })?;
        counters.record_allocation_event()?;
    }
    transitions.push(transition);
    Ok(())
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
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(slot_count)
        .map_err(|_| TileError::Allocation { amount: slot_count })?;
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
    let mut slots = Vec::new();
    slots
        .try_reserve_exact(slot_count)
        .map_err(|_| TileError::Allocation { amount: slot_count })?;
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
    Ok(slots)
}

enum CellMutation {
    Unchanged,
    Delete,
    Replace(Vec<u8>),
}

fn mutate_cell(
    previous: Option<&[u8]>,
    change: BncChange,
    max_output_bytes: usize,
) -> Result<CellMutation> {
    let Some(previous) = previous else {
        return match change {
            BncChange::Clear => Ok(CellMutation::Unchanged),
            BncChange::Set(input) => {
                let mut cell = BncCell::minimal();
                apply_input(&mut cell, input)?;
                cell.try_encode_with_limit(max_output_bytes)
                    .map(CellMutation::Replace)
                    .map_err(map_bnc_error)
            },
            BncChange::FormulaCache(_) => Err(TileError::UnsupportedSource { row: 0, column: 0 }),
        };
    };
    let view = BncCellView::parse(previous).map_err(|_| TileError::InvalidSource)?;
    match change {
        BncChange::Clear => match view
            .clear_value_with_limit(max_output_bytes)
            .map_err(map_bnc_error)?
        {
            ClearValue::Delete => Ok(CellMutation::Delete),
            ClearValue::Retain(bytes) => Ok(CellMutation::Replace(bytes)),
        },
        BncChange::Set(input) => view
            .rewrite_scalar_with_limit(input.as_wire(), max_output_bytes)
            .map(CellMutation::Replace)
            .map_err(map_bnc_error),
        BncChange::FormulaCache(input) => view
            .rewrite_formula_cache_with_limit(input.as_wire(), max_output_bytes)
            .map(CellMutation::Replace)
            .map_err(map_bnc_error),
    }
}

fn apply_input(cell: &mut BncCell, input: ScalarInput) -> Result<()> {
    match input {
        ScalarInput::String(identifier) => cell.set_string(identifier),
        ScalarInput::RichText(identifier) => cell.set_rich_text(identifier),
        ScalarInput::Number(value) => cell
            .set_number(value.get())
            .map_err(|_| TileError::InvalidSource)?,
        ScalarInput::Boolean(value) => cell.set_boolean(value),
        ScalarInput::Date(value) => cell
            .set_date(value.get())
            .map_err(|_| TileError::InvalidSource)?,
        ScalarInput::Duration(value) => cell
            .set_duration(value.get())
            .map_err(|_| TileError::InvalidSource)?,
    }
    Ok(())
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

fn classify_bnc_with_references(cell: Option<&[u8]>) -> Result<(CellValue, CellReferences)> {
    let Some(cell) = cell else {
        return Ok((CellValue::Missing, CellReferences::default()));
    };
    let view = BncCellView::parse(cell).map_err(|_| TileError::InvalidSource)?;
    Ok((classify_bnc_view(&view), bnc_references_view(&view)))
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
        BncChange::Set(input) => view.scalar_equals(input.as_wire()),
        BncChange::FormulaCache(input) => view.formula_cache_equals(input.as_wire()),
    })
}

fn encode_slots(
    slots: &[Slot<'_>],
    prefer_wide: bool,
    counters: &mut Counters,
) -> Result<(Vec<u8>, Vec<u8>, bool)> {
    let wide = prefer_wide || slots_require_wide(slots)?;
    let result = encode_slots_with_width(slots, wide, counters)?;
    Ok((result.0, result.1, wide))
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
    wide: bool,
    counters: &mut Counters,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let offset_bytes = slots.len().checked_mul(2).ok_or(TileError::InvalidSource)?;
    let storage_content = slots
        .iter()
        .try_fold(0usize, |total, slot| {
            total.checked_add(slot.bytes().map_or(0, <[u8]>::len))
        })
        .ok_or(TileError::InvalidSource)?;
    let storage_capacity = storage_content
        .checked_add(slots.len().checked_mul(3).ok_or(TileError::InvalidSource)?)
        .ok_or(TileError::InvalidSource)?;
    counters.charge(u64_from_usize(storage_capacity)?)?;
    counters.charge(u64_from_usize(offset_bytes)?)?;
    let mut storage = Vec::new();
    counters.preflight_scratch_allocation(storage_capacity)?;
    storage
        .try_reserve_exact(storage_capacity)
        .map_err(|_| TileError::Allocation {
            amount: storage_capacity,
        })?;
    counters.record_scratch_allocation(storage.capacity())?;
    let mut offsets = Vec::new();
    counters.preflight_scratch_allocation(offset_bytes)?;
    offsets
        .try_reserve_exact(offset_bytes)
        .map_err(|_| TileError::Allocation {
            amount: offset_bytes,
        })?;
    counters.record_scratch_allocation(offsets.capacity())?;
    let unit = if wide { 4usize } else { 1usize };
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
    Ok((storage, offsets))
}

fn rewrite_row_message(
    source: &[u8],
    cell_count: u32,
    storage: &[u8],
    offsets: &[u8],
    previous_wide: Option<bool>,
    wide: bool,
    counters: &mut Counters,
) -> Result<Vec<u8>> {
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
                            u64::try_from(storage.len()).map_err(|_| TileError::InvalidSource)?,
                        )
                        + storage.len()
                },
                7 => {
                    offsets_present = true;
                    field.key().len()
                        + varint_len(
                            u64::try_from(offsets.len()).map_err(|_| TileError::InvalidSource)?,
                        )
                        + offsets.len()
                },
                8 => {
                    wide_present = true;
                    field.key().len() + 1
                },
                _ => field.raw().len(),
            })
            .ok_or(TileError::InvalidSource)?;
    }
    if !count_present || !storage_present || !offsets_present || (previous_wide.is_none() && wide) {
        if previous_wide.is_none() && wide && !wide_present {
            length = length.checked_add(2).ok_or(TileError::InvalidSource)?;
        } else {
            return Err(TileError::InvalidSource);
        }
    }
    counters.charge(u64_from_usize(length)?)?;
    let mut output = Vec::new();
    counters.preflight_scratch_allocation(length)?;
    output
        .try_reserve_exact(length)
        .map_err(|_| TileError::Allocation { amount: length })?;
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
    if previous_wide.is_none() && wide && !wide_present {
        output.extend_from_slice(&[0x40, 1]);
    }
    if output.len() != length {
        return Err(TileError::InvalidSource);
    }
    Ok(output)
}

fn rewritten_tile_length(source: &[u8], replacements: &[RowReplacement]) -> Result<usize> {
    let view = WireView::parse(source).map_err(|_| TileError::InvalidSource)?;
    let mut replacement = 0usize;
    let mut length = 0usize;
    for (index, field) in view.fields().enumerate() {
        let bytes = if replacements
            .get(replacement)
            .is_some_and(|next| next.field_index == index)
        {
            let next = &replacements[replacement];
            replacement += 1;
            field
                .key()
                .len()
                .checked_add(varint_len(
                    u64::try_from(next.payload.len()).map_err(|_| TileError::InvalidSource)?,
                ))
                .and_then(|value| value.checked_add(next.payload.len()))
        } else {
            Some(field.raw().len())
        };
        length = length
            .checked_add(bytes.ok_or(TileError::InvalidSource)?)
            .ok_or(TileError::InvalidSource)?;
    }
    (replacement == replacements.len())
        .then_some(length)
        .ok_or(TileError::InvalidSource)
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
        Ok(())
    }
    fn record_scratch_allocation(&mut self, bytes: usize) -> Result<()> {
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
        self.report.retained_bytes = self
            .report
            .retained_bytes
            .checked_add(u64_from_usize(bytes)?)
            .ok_or(TileError::InvalidSource)?;
        self.report.retained_elements = self
            .report
            .retained_elements
            .checked_add(u64_from_usize(elements)?)
            .ok_or(TileError::InvalidSource)?;
        Ok(())
    }
    fn record_allocation_event(&mut self) -> Result<()> {
        self.report.allocation_events = self
            .report
            .allocation_events
            .checked_add(1)
            .ok_or(TileError::InvalidSource)?;
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
