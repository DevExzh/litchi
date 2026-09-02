//! Strict, source-preserving codecs for the physical table-sort seam.
//!
//! This module owns only the small repeated-storage rewrites needed by a
//! format adapter.  It deliberately does not know about package objects,
//! native message identifiers, or a particular iWork format.  Every plan
//! borrows the caller's source bytes, validates the complete supported wire
//! envelope before staging, and emits the untouched fields byte-for-byte and
//! in their original order.  Buffa lazy views are used as a bounded parity
//! oracle for singular scalar envelopes; they never become the preservation
//! or mutation representation.  Structurally valid unknown fields remain
//! readable and source-preserved for no-op plans, but a plan that changes a
//! mutable row/index rejects those fields before candidate allocation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers are kept adjacent to their source-preserving plans."
)]

use core::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_table_physical_sort_generated::LitchiIwaTablePhysicalSortProjection as projection;

const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

/// Explicit finite resource limits for one decode or rewrite transaction.
///
/// The limits are intentionally operation-local.  A caller running several
/// plans concurrently must account for the sum of their budgets at the
/// orchestration layer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_records: usize,
    max_elements: usize,
    max_output_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Construct an explicit bytes/fields/work/nesting/record policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_records: usize,
        max_elements: usize,
        max_output_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_records,
            max_elements,
            max_output_bytes,
            max_scratch_bytes,
        }
    }

    /// A bounded policy suitable for ordinary native table objects.
    #[must_use]
    pub const fn bounded() -> Self {
        Self::new(
            64 * 1024 * 1024,
            1 << 20,
            256 * 1024 * 1024,
            MAX_RECURSION,
            1 << 20,
            1 << 20,
            64 * 1024 * 1024,
            64 * 1024 * 1024,
        )
    }

    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    #[must_use]
    pub const fn recursion_limit(self) -> u32 {
        self.recursion_limit
    }

    #[must_use]
    pub const fn max_records(self) -> usize {
        self.max_records
    }

    #[must_use]
    pub const fn max_elements(self) -> usize {
        self.max_elements
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            // Buffa 0.9.1 does not charge deferred allocations against this
            // setting. The handwritten budget remains authoritative.
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

impl Default for DecodeOptions {
    fn default() -> Self {
        Self::bounded()
    }
}

/// Exact successful consumption for one source or candidate pass.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    records: usize,
    elements: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn records(self) -> usize {
        self.records
    }

    #[must_use]
    pub const fn elements(self) -> usize {
        self.elements
    }
}

/// Typed finite resource refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Records { observed: usize, maximum: usize },
    Elements { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    ScratchBytes { observed: usize, maximum: usize },
    Allocation { requested: usize },
}

/// Strict physical-sort wire failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }

    pub(crate) const fn invalid() -> Self {
        Self { limit: None }
    }

    pub(crate) const fn limited(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid physical table-sort storage payload")
    }
}

impl std::error::Error for DecodeError {}

/// Exact preflight accounting for one prepared rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteRequirements {
    source: DecodeReport,
    output_bytes: usize,
    work_bytes: usize,
    fields: usize,
    records: usize,
    elements: usize,
}

impl RewriteRequirements {
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn records(self) -> usize {
        self.records
    }

    #[must_use]
    pub const fn elements(self) -> usize {
        self.elements
    }
}

/// Source/candidate accounting for a rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    source: DecodeReport,
    result: DecodeReport,
    output_bytes: usize,
    changed_records: usize,
}

impl RewriteReport {
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }

    #[must_use]
    pub const fn result(self) -> DecodeReport {
        self.result
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn changed_records(self) -> usize {
        self.changed_records
    }
}

/// Stable row move represented as source and destination local indexes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RowMove {
    source_index: u32,
    destination_index: u32,
}

impl RowMove {
    #[must_use]
    pub const fn new(source_index: u32, destination_index: u32) -> Self {
        Self {
            source_index,
            destination_index,
        }
    }

    #[must_use]
    pub const fn source_index(self) -> u32 {
        self.source_index
    }

    #[must_use]
    pub const fn destination_index(self) -> u32 {
        self.destination_index
    }
}

/// Alias documenting a header record move at the physical-storage boundary.
pub type HeaderRowMove = RowMove;

/// Compatibility spelling for tile adapters that call a move an edit.
pub type TileRowEdit = RowMove;

/// Compatibility spelling for header adapters that call a move an edit.
pub type HeaderRowEdit = RowMove;

/// A validated row permutation for UID-map rewrites.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RowUidPermutation {
    destination_by_source: Vec<u32>,
}

impl RowUidPermutation {
    /// Validate and copy a source-indexed destination permutation.
    pub fn new(destination_by_source: &[u32]) -> Result<Self, DecodeError> {
        validate_permutation(destination_by_source)?;
        let mut copied = Vec::new();
        try_reserve_exact(&mut copied, destination_by_source.len())?;
        copied.extend_from_slice(destination_by_source);
        Ok(Self {
            destination_by_source: copied,
        })
    }

    /// Build an identity permutation without exposing unchecked offsets.
    pub fn identity(row_count: u32) -> Result<Self, DecodeError> {
        let count = usize::try_from(row_count).map_err(|_| DecodeError::invalid())?;
        let mut values = Vec::new();
        try_reserve_exact(&mut values, count)?;
        for index in 0..count {
            values.push(u32::try_from(index).map_err(|_| DecodeError::invalid())?);
        }
        Ok(Self {
            destination_by_source: values,
        })
    }

    #[must_use]
    pub fn destination_by_source(&self) -> &[u32] {
        &self.destination_by_source
    }

    /// Return the destination for one source index.
    #[must_use]
    pub fn destination_for_source(&self, source_index: usize) -> Option<u32> {
        self.destination_by_source.get(source_index).copied()
    }

    fn validate_len(&self, row_count: usize) -> Result<(), DecodeError> {
        if self.destination_by_source.len() != row_count {
            return Err(DecodeError::invalid());
        }
        validate_permutation(&self.destination_by_source)
    }
}

/// Native UUID words used by the stable row/column UID map.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Uuid {
    lower: u64,
    upper: u64,
}

impl Uuid {
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

impl fmt::Debug for Uuid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Uuid(<redacted>)")
    }
}

/// Fully validated, owned row/column UID projection.
#[derive(Clone, PartialEq, Eq)]
pub struct ColumnRowUidMapSnapshot {
    sorted_column_uids: Vec<Uuid>,
    column_index_for_uid: Vec<u32>,
    column_uid_for_index: Vec<u32>,
    sorted_row_uids: Vec<Uuid>,
    row_index_for_uid: Vec<u32>,
    row_uid_for_index: Vec<u32>,
}

impl fmt::Debug for ColumnRowUidMapSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ColumnRowUidMapSnapshot")
            .field("column_count", &self.sorted_column_uids.len())
            .field("row_count", &self.sorted_row_uids.len())
            .finish()
    }
}

impl ColumnRowUidMapSnapshot {
    #[must_use]
    pub fn sorted_column_uids(&self) -> &[Uuid] {
        &self.sorted_column_uids
    }

    #[must_use]
    pub fn column_index_for_uid(&self) -> &[u32] {
        &self.column_index_for_uid
    }

    #[must_use]
    pub fn column_uid_for_index(&self) -> &[u32] {
        &self.column_uid_for_index
    }

    #[must_use]
    pub fn sorted_row_uids(&self) -> &[Uuid] {
        &self.sorted_row_uids
    }

    #[must_use]
    pub fn row_index_for_uid(&self) -> &[u32] {
        &self.row_index_for_uid
    }

    #[must_use]
    pub fn row_uid_for_index(&self) -> &[u32] {
        &self.row_uid_for_index
    }

    #[must_use]
    pub fn row_count(&self) -> usize {
        self.sorted_row_uids.len()
    }

    #[must_use]
    pub fn column_count(&self) -> usize {
        self.sorted_column_uids.len()
    }
}

/// Strictly inspect a row/column UID map and validate both inverse pairs.
pub fn decode_column_row_uid_map(
    source: &[u8],
    column_count: usize,
    row_count: usize,
    options: DecodeOptions,
) -> Result<(ColumnRowUidMapSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_uid_map_in(source, column_count, row_count, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

/// Tile scalar envelope, with repeated `row_infos` intentionally omitted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TileSnapshot {
    max_column: u32,
    max_row: u32,
    num_cells: u32,
    num_rows: u32,
    storage_version: Option<u32>,
    last_saved_in_bnc: Option<bool>,
    should_use_wide_rows: Option<bool>,
}

impl TileSnapshot {
    #[must_use]
    pub const fn max_column(self) -> u32 {
        self.max_column
    }

    #[must_use]
    pub const fn max_row(self) -> u32 {
        self.max_row
    }

    #[must_use]
    pub const fn num_cells(self) -> u32 {
        self.num_cells
    }

    #[must_use]
    pub const fn num_rows(self) -> u32 {
        self.num_rows
    }

    #[must_use]
    pub const fn storage_version(self) -> Option<u32> {
        self.storage_version
    }

    #[must_use]
    pub const fn last_saved_in_bnc(self) -> Option<bool> {
        self.last_saved_in_bnc
    }

    #[must_use]
    pub const fn should_use_wide_rows(self) -> Option<bool> {
        self.should_use_wide_rows
    }
}

/// Strict borrowed `TileRowInfo` scalar projection.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct TileRowInfoSnapshot<'source> {
    tile_row_index: u32,
    cell_count: u32,
    cell_storage_buffer_pre_bnc: &'source [u8],
    cell_offsets_pre_bnc: &'source [u8],
    storage_version: Option<u32>,
    cell_storage_buffer: Option<&'source [u8]>,
    cell_offsets: Option<&'source [u8]>,
    has_wide_offsets: Option<bool>,
}

impl fmt::Debug for TileRowInfoSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TileRowInfoSnapshot")
            .field("tile_row_index", &self.tile_row_index)
            .field("cell_count", &self.cell_count)
            .field("pre_bnc_bytes", &self.cell_storage_buffer_pre_bnc.len())
            .field("pre_bnc_offsets", &self.cell_offsets_pre_bnc.len())
            .finish()
    }
}

impl<'source> TileRowInfoSnapshot<'source> {
    #[must_use]
    pub const fn tile_row_index(self) -> u32 {
        self.tile_row_index
    }

    #[must_use]
    pub const fn cell_count(self) -> u32 {
        self.cell_count
    }

    #[must_use]
    pub const fn cell_storage_buffer_pre_bnc(self) -> &'source [u8] {
        self.cell_storage_buffer_pre_bnc
    }

    #[must_use]
    pub const fn cell_offsets_pre_bnc(self) -> &'source [u8] {
        self.cell_offsets_pre_bnc
    }

    #[must_use]
    pub const fn storage_version(self) -> Option<u32> {
        self.storage_version
    }

    #[must_use]
    pub const fn cell_storage_buffer(self) -> Option<&'source [u8]> {
        self.cell_storage_buffer
    }

    #[must_use]
    pub const fn cell_offsets(self) -> Option<&'source [u8]> {
        self.cell_offsets
    }

    #[must_use]
    pub const fn has_wide_offsets(self) -> Option<bool> {
        self.has_wide_offsets
    }
}

/// One source-ordered row record from a tile envelope.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct TileRowRecord<'source> {
    raw: &'source [u8],
    snapshot: TileRowInfoSnapshot<'source>,
}

impl fmt::Debug for TileRowRecord<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("TileRowRecord { payload: <redacted> }")
    }
}

impl<'source> TileRowRecord<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }

    #[must_use]
    pub const fn snapshot(self) -> TileRowInfoSnapshot<'source> {
        self.snapshot
    }
}

/// One source-ordered header record from a bucket envelope.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct HeaderRecord<'source> {
    raw: &'source [u8],
    snapshot: HeaderSnapshot,
}

impl fmt::Debug for HeaderRecord<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("HeaderRecord { payload: <redacted> }")
    }
}

impl<'source> HeaderRecord<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }

    #[must_use]
    pub const fn snapshot(self) -> HeaderSnapshot {
        self.snapshot
    }
}

/// Strict header scalar projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderSnapshot {
    index: u32,
    size_bits: u32,
    hiding_state: u32,
    number_of_cells: u32,
    has_cell_style: bool,
    has_text_style: bool,
}

impl HeaderSnapshot {
    #[must_use]
    pub const fn index(self) -> u32 {
        self.index
    }

    #[must_use]
    pub const fn size_bits(self) -> u32 {
        self.size_bits
    }

    #[must_use]
    pub const fn hiding_state(self) -> u32 {
        self.hiding_state
    }

    #[must_use]
    pub const fn number_of_cells(self) -> u32 {
        self.number_of_cells
    }

    /// Optional style payload presence. The payload itself remains opaque.
    #[must_use]
    pub const fn has_cell_style(self) -> bool {
        self.has_cell_style
    }

    /// Optional text-style payload presence. The payload itself remains opaque.
    #[must_use]
    pub const fn has_text_style(self) -> bool {
        self.has_text_style
    }
}

fn header_snapshot(
    index: u32,
    size_bits: u32,
    hiding_state: u32,
    number_of_cells: u32,
    cell_style: Option<&[u8]>,
    text_style: Option<&[u8]>,
) -> HeaderSnapshot {
    HeaderSnapshot {
        index,
        size_bits,
        hiding_state,
        number_of_cells,
        has_cell_style: cell_style.is_some(),
        has_text_style: text_style.is_some(),
    }
}

/// Prepare a tile's row repeated-field rewrite.
///
/// Present row records must have globally unique indexes and the highest
/// present index must be `num_rows - 1`; gaps below that index are valid and
/// remain sparse. Every move source must name a present record, while move
/// destinations form a unique in-range set together with untouched records.
pub fn plan_tile_rows_rewrite<'source>(
    source: &'source [u8],
    row_index_limit: u32,
    moves: &[RowMove],
    options: DecodeOptions,
) -> Result<PreparedTileRowsRewrite<'source>, DecodeError> {
    if row_index_limit == 0 {
        return Err(DecodeError::invalid());
    }
    let mut budget = Budget::new(source, options)?;
    let (tile, records, num_rows_span) =
        decode_tile_for_rewrite(source, row_index_limit, &mut budget, 1)?;
    if moves.len() > options.max_records {
        return Err(DecodeError::limited(DecodeLimit::Records {
            observed: moves.len(),
            maximum: options.max_records,
        }));
    }
    let mut ordered_moves = Vec::new();
    try_reserve_exact(&mut ordered_moves, moves.len())?;
    ordered_moves.extend_from_slice(moves);
    ordered_moves.sort_unstable_by_key(|movement| movement.source_index);
    let mut available_indices = Vec::new();
    try_reserve_exact(&mut available_indices, records.len())?;
    available_indices.extend(
        records
            .iter()
            .map(|record| record.snapshot.tile_row_index()),
    );
    available_indices.sort_unstable();
    validate_move_list(&ordered_moves, row_index_limit, &available_indices)?;

    let mut destinations = Vec::new();
    try_reserve_exact(&mut destinations, records.len())?;
    for record in &records {
        let destination = move_for(&ordered_moves, record.snapshot.tile_row_index())
            .unwrap_or(record.snapshot.tile_row_index());
        destinations.push(destination);
    }
    validate_values_within_limit(&destinations, row_index_limit)?;

    let mut desired_order = Vec::new();
    try_reserve_exact(&mut desired_order, records.len())?;
    desired_order.extend(0..records.len());
    if destinations
        .iter()
        .zip(records.iter())
        .any(|(destination, record)| *destination != record.snapshot.tile_row_index())
    {
        desired_order.sort_unstable_by_key(|index| (destinations[*index], *index));
        if desired_order
            .windows(2)
            .any(|pair| destinations[pair[0]] == destinations[pair[1]])
        {
            return Err(DecodeError::invalid());
        }
    }
    let new_num_rows = destinations.iter().copied().max().map_or(Ok(0), |row| {
        row.checked_add(1).ok_or_else(DecodeError::invalid)
    })?;
    let old_num_rows = tile.num_rows;
    let mut output_bytes = source.len();
    let mut changed_records = 0usize;
    let mut changed_row_bytes = 0usize;
    for (record_index, record) in records.iter().enumerate() {
        let destination = destinations[record_index];
        if destination != record.snapshot.tile_row_index() {
            changed_records = changed_records
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
            let original_index = encoded_varint_len(u64::from(record.snapshot.tile_row_index()));
            let replacement_index = encoded_varint_len(u64::from(destination));
            let payload_len = record
                .payload_end
                .checked_sub(record.payload_start)
                .ok_or_else(DecodeError::invalid)?;
            let (new_payload_len, original_length, replacement_length) =
                nested_varint_rewrite_sizes(payload_len, original_index, replacement_index)?;
            if new_payload_len > options.max_message_bytes {
                return Err(DecodeError::limited(DecodeLimit::Bytes {
                    observed: new_payload_len,
                    maximum: options.max_message_bytes,
                }));
            }
            output_bytes = replace_size_parts(
                output_bytes,
                [original_index, original_length],
                [replacement_index, replacement_length],
            )?;
            changed_row_bytes = changed_row_bytes
                .checked_add(new_payload_len.max(payload_len))
                .and_then(|value| value.checked_add(replacement_length.max(original_length)))
                .ok_or_else(DecodeError::invalid)?;
        }
    }
    if new_num_rows != old_num_rows {
        output_bytes = output_bytes
            .checked_add(encoded_varint_len(u64::from(new_num_rows)))
            .and_then(|value| value.checked_sub(encoded_varint_len(u64::from(old_num_rows))))
            .ok_or_else(DecodeError::invalid)?;
    }
    let has_unknown_fields = budget.has_unknown_fields();
    if has_unknown_fields && (changed_records != 0 || new_num_rows != old_num_rows) {
        return Err(DecodeError::invalid());
    }
    let fields = budget.fields;
    let records_count = records.len();
    let elements = records_count;
    let work_bytes = rewrite_work_upper_bound(
        source.len(),
        output_bytes,
        records_count,
        moves.len(),
        changed_row_bytes,
    )?;
    let requirements = RewriteRequirements {
        source: budget.report(),
        output_bytes,
        work_bytes,
        fields,
        records: records_count,
        elements,
    };
    validate_requirements(requirements, options)?;
    Ok(PreparedTileRowsRewrite {
        source,
        tile,
        records,
        num_rows_span,
        destinations,
        desired_order,
        has_unknown_fields,
        requirements,
    })
}

/// Execute a validated tile-row plan and revalidate the candidate.
pub fn execute_tile_rows_rewrite(
    plan: PreparedTileRowsRewrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_requirements(plan.requirements, options)?;
    let changed_records = plan
        .destinations
        .iter()
        .zip(plan.records.iter())
        .filter(|(destination, record)| **destination != record.snapshot.tile_row_index())
        .count();
    if plan.has_unknown_fields && changed_records != 0 {
        return Err(DecodeError::invalid());
    }
    let output = assemble_tile_rows(&plan, options)?;
    let mut verification_budget = Budget::new(&output, options)?;
    let (_tile, _records, _span) =
        decode_tile_for_rewrite(&output, u32::MAX, &mut verification_budget, 1)?;
    let result = verification_budget.report();
    Ok((
        output,
        RewriteReport {
            source: plan.requirements.source,
            result,
            output_bytes: plan.requirements.output_bytes,
            changed_records,
        },
    ))
}

/// One-shot tile-row rewrite convenience wrapper.
pub fn rewrite_tile_rows(
    source: &[u8],
    row_index_limit: u32,
    moves: &[RowMove],
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let plan = plan_tile_rows_rewrite(source, row_index_limit, moves, options)?;
    execute_tile_rows_rewrite(plan, options)
}

/// Prepare a header bucket's row-record rewrite.
///
/// Header records may be sparse: only present records are retained and the
/// record count is never inflated to fill gaps. Source and destination
/// indexes are globally unique and bounded by `row_index_limit`; every move
/// source must name one of the present records.
pub fn plan_header_storage_bucket_rows<'source>(
    source: &'source [u8],
    row_index_limit: u32,
    moves: &[HeaderRowMove],
    options: DecodeOptions,
) -> Result<PreparedHeaderRowsRewrite<'source>, DecodeError> {
    if row_index_limit == 0 {
        return Err(DecodeError::invalid());
    }
    let mut budget = Budget::new(source, options)?;
    let (bucket, records) =
        decode_header_bucket_for_rewrite(source, row_index_limit, &mut budget, 1)?;
    if moves.len() > options.max_records {
        return Err(DecodeError::limited(DecodeLimit::Records {
            observed: moves.len(),
            maximum: options.max_records,
        }));
    }
    let mut ordered_moves = Vec::new();
    try_reserve_exact(&mut ordered_moves, moves.len())?;
    ordered_moves.extend_from_slice(moves);
    ordered_moves.sort_unstable_by_key(|movement| movement.source_index);
    let mut available_indices = Vec::new();
    try_reserve_exact(&mut available_indices, records.len())?;
    available_indices.extend(records.iter().map(|record| record.snapshot.index()));
    available_indices.sort_unstable();
    validate_move_list(&ordered_moves, row_index_limit, &available_indices)?;
    let mut destinations = Vec::new();
    try_reserve_exact(&mut destinations, records.len())?;
    for record in &records {
        destinations.push(
            move_for(&ordered_moves, record.snapshot.index()).unwrap_or(record.snapshot.index()),
        );
    }
    validate_values_within_limit(&destinations, row_index_limit)?;
    let mut desired_order = Vec::new();
    try_reserve_exact(&mut desired_order, records.len())?;
    desired_order.extend(0..records.len());
    if destinations
        .iter()
        .zip(records.iter())
        .any(|(destination, record)| *destination != record.snapshot.index())
    {
        desired_order.sort_unstable_by_key(|index| (destinations[*index], *index));
        if desired_order
            .windows(2)
            .any(|pair| destinations[pair[0]] == destinations[pair[1]])
        {
            return Err(DecodeError::invalid());
        }
    }
    let mut output_bytes = source.len();
    let mut changed_records = 0usize;
    let mut changed_row_bytes = 0usize;
    for (record_index, record) in records.iter().enumerate() {
        let destination = destinations[record_index];
        if destination != record.snapshot.index() {
            changed_records = changed_records
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
            let original_index = encoded_varint_len(u64::from(record.snapshot.index()));
            let replacement_index = encoded_varint_len(u64::from(destination));
            let payload_len = record
                .payload_end
                .checked_sub(record.payload_start)
                .ok_or_else(DecodeError::invalid)?;
            let (new_payload_len, original_length, replacement_length) =
                nested_varint_rewrite_sizes(payload_len, original_index, replacement_index)?;
            if new_payload_len > options.max_message_bytes {
                return Err(DecodeError::limited(DecodeLimit::Bytes {
                    observed: new_payload_len,
                    maximum: options.max_message_bytes,
                }));
            }
            output_bytes = replace_size_parts(
                output_bytes,
                [original_index, original_length],
                [replacement_index, replacement_length],
            )?;
            changed_row_bytes = changed_row_bytes
                .checked_add(new_payload_len.max(payload_len))
                .and_then(|value| value.checked_add(replacement_length.max(original_length)))
                .ok_or_else(DecodeError::invalid)?;
        }
    }
    let has_unknown_fields = budget.has_unknown_fields();
    if has_unknown_fields && changed_records != 0 {
        return Err(DecodeError::invalid());
    }
    let requirements = RewriteRequirements {
        source: budget.report(),
        output_bytes,
        work_bytes: rewrite_work_upper_bound(
            source.len(),
            output_bytes,
            records.len(),
            moves.len(),
            changed_row_bytes,
        )?,
        fields: budget.fields,
        records: records.len(),
        elements: records.len(),
    };
    validate_requirements(requirements, options)?;
    Ok(PreparedHeaderRowsRewrite {
        source,
        bucket,
        records,
        destinations,
        desired_order,
        has_unknown_fields,
        requirements,
    })
}

/// Execute a validated header-row plan and revalidate the candidate.
pub fn execute_header_storage_bucket_rows(
    plan: PreparedHeaderRowsRewrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_requirements(plan.requirements, options)?;
    let changed_records = plan
        .destinations
        .iter()
        .zip(plan.records.iter())
        .filter(|(destination, record)| **destination != record.snapshot.index())
        .count();
    if plan.has_unknown_fields && changed_records != 0 {
        return Err(DecodeError::invalid());
    }
    let output = assemble_header_rows(&plan, options)?;
    let mut verification_budget = Budget::new(&output, options)?;
    let (_bucket, _records) =
        decode_header_bucket_for_rewrite(&output, u32::MAX, &mut verification_budget, 1)?;
    Ok((
        output,
        RewriteReport {
            source: plan.requirements.source,
            result: verification_budget.report(),
            output_bytes: plan.requirements.output_bytes,
            changed_records,
        },
    ))
}

/// One-shot header-row rewrite convenience wrapper.
pub fn rewrite_header_storage_bucket_rows(
    source: &[u8],
    row_index_limit: u32,
    moves: &[HeaderRowMove],
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let plan = plan_header_storage_bucket_rows(source, row_index_limit, moves, options)?;
    execute_header_storage_bucket_rows(plan, options)
}

/// Prepared source-preserving tile row rewrite.
pub struct PreparedTileRowsRewrite<'source> {
    source: &'source [u8],
    tile: TileSnapshot,
    records: Vec<StagedTileRow<'source>>,
    num_rows_span: FieldSpan,
    destinations: Vec<u32>,
    desired_order: Vec<usize>,
    has_unknown_fields: bool,
    requirements: RewriteRequirements,
}

impl fmt::Debug for PreparedTileRowsRewrite<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedTileRowsRewrite")
            .field("row_count", &self.records.len())
            .field("requirements", &self.requirements)
            .finish()
    }
}

impl PreparedTileRowsRewrite<'_> {
    #[must_use]
    pub const fn requirements(&self) -> RewriteRequirements {
        self.requirements
    }

    #[must_use]
    pub const fn tile(&self) -> TileSnapshot {
        self.tile
    }

    pub fn row_records(&self) -> impl Iterator<Item = TileRowRecord<'_>> {
        self.records.iter().map(|record| TileRowRecord {
            raw: record.payload(self.source),
            snapshot: record.snapshot,
        })
    }
}

/// Prepared source-preserving header row rewrite.
pub struct PreparedHeaderRowsRewrite<'source> {
    source: &'source [u8],
    bucket: BucketSnapshot,
    records: Vec<StagedHeader>,
    destinations: Vec<u32>,
    desired_order: Vec<usize>,
    has_unknown_fields: bool,
    requirements: RewriteRequirements,
}

impl fmt::Debug for PreparedHeaderRowsRewrite<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedHeaderRowsRewrite")
            .field("header_count", &self.records.len())
            .field("requirements", &self.requirements)
            .finish()
    }
}

impl PreparedHeaderRowsRewrite<'_> {
    #[must_use]
    pub const fn requirements(&self) -> RewriteRequirements {
        self.requirements
    }

    #[must_use]
    pub const fn bucket_hash_function(&self) -> u32 {
        self.bucket.bucket_hash_function
    }

    pub fn records(&self) -> impl Iterator<Item = HeaderRecord<'_>> {
        self.records.iter().map(|record| HeaderRecord {
            raw: record.payload(self.source),
            snapshot: record.snapshot,
        })
    }
}

#[derive(Clone, Copy)]
struct FieldSpan {
    start: usize,
    end: usize,
    value_start: usize,
    value_end: usize,
}

struct StagedTileRow<'source> {
    payload_start: usize,
    payload_end: usize,
    index_field: FieldSpan,
    snapshot: TileRowInfoSnapshot<'source>,
}

impl StagedTileRow<'_> {
    fn payload<'source>(&self, source: &'source [u8]) -> &'source [u8] {
        &source[self.payload_start..self.payload_end]
    }
}

struct StagedHeader {
    payload_start: usize,
    payload_end: usize,
    index_field: FieldSpan,
    snapshot: HeaderSnapshot,
}

impl StagedHeader {
    fn payload<'source>(&self, source: &'source [u8]) -> &'source [u8] {
        &source[self.payload_start..self.payload_end]
    }
}

#[derive(Debug, Clone, Copy)]
struct BucketSnapshot {
    bucket_hash_function: u32,
}

fn decode_tile_for_rewrite<'source>(
    source: &'source [u8],
    row_index_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(TileSnapshot, Vec<StagedTileRow<'source>>, FieldSpan), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut max_column = None;
    let mut max_row = None;
    let mut num_cells = None;
    let mut num_rows = None;
    let mut storage_version = None;
    let mut last_saved_in_bnc = None;
    let mut should_use_wide_rows = None;
    let mut num_rows_span = None;
    let mut rows = Vec::new();
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            1 => set_once(&mut max_column, field.varint().and_then(canonical_u32)?)?,
            2 => set_once(&mut max_row, field.varint().and_then(canonical_u32)?)?,
            3 => set_once(&mut num_cells, field.varint().and_then(canonical_u32)?)?,
            4 => {
                set_once(&mut num_rows, field.varint().and_then(canonical_u32)?)?;
                num_rows_span = Some(field.span);
            },
            5 => {
                let payload = field.bytes()?;
                let payload_start = field.span.value_start;
                let payload_end = field.span.value_end;
                let (snapshot, index_field) =
                    decode_tile_row_info_in(payload, row_index_limit, budget, child_depth)?;
                if rows.len() >= budget.options.max_records {
                    let observed = rows.len().checked_add(1).ok_or_else(DecodeError::invalid)?;
                    return Err(DecodeError::limited(DecodeLimit::Records {
                        observed,
                        maximum: budget.options.max_records,
                    }));
                }
                let requested = rows.len().checked_add(1).ok_or_else(DecodeError::invalid)?;
                rows.try_reserve(1)
                    .map_err(|_| DecodeError::limited(DecodeLimit::Allocation { requested }))?;
                rows.push(StagedTileRow {
                    payload_start,
                    payload_end,
                    index_field,
                    snapshot,
                });
                budget.records(1)?;
            },
            6 => set_once(
                &mut storage_version,
                field.varint().and_then(canonical_u32)?,
            )?,
            7 => set_once(
                &mut last_saved_in_bnc,
                field.varint().and_then(canonical_bool)?,
            )?,
            8 => set_once(
                &mut should_use_wide_rows,
                field.varint().and_then(canonical_bool)?,
            )?,
            _ => budget.mark_unknown_field(),
        }
    }
    let tile = TileSnapshot {
        max_column: max_column.ok_or_else(DecodeError::invalid)?,
        max_row: max_row.ok_or_else(DecodeError::invalid)?,
        num_cells: num_cells.ok_or_else(DecodeError::invalid)?,
        num_rows: num_rows.ok_or_else(DecodeError::invalid)?,
        storage_version,
        last_saved_in_bnc,
        should_use_wide_rows,
    };
    let num_rows_span = num_rows_span.ok_or_else(DecodeError::invalid)?;
    if rows
        .iter()
        .any(|row| row.snapshot.tile_row_index() >= tile.num_rows)
    {
        return Err(DecodeError::invalid());
    }
    if rows
        .iter()
        .map(|row| row.snapshot.tile_row_index())
        .max()
        .map_or(Ok(0), |row| {
            row.checked_add(1).ok_or_else(DecodeError::invalid)
        })?
        != tile.num_rows
    {
        return Err(DecodeError::invalid());
    }
    parity_tile(source, tile, budget, depth)?;
    Ok((tile, rows, num_rows_span))
}

fn decode_tile_row_info_in<'source>(
    source: &'source [u8],
    row_index_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(TileRowInfoSnapshot<'source>, FieldSpan), DecodeError> {
    budget.message(source, depth)?;
    let mut tile_row_index = None;
    let mut cell_count = None;
    let mut cell_storage_buffer_pre_bnc = None;
    let mut cell_offsets_pre_bnc = None;
    let mut storage_version = None;
    let mut cell_storage_buffer = None;
    let mut cell_offsets = None;
    let mut has_wide_offsets = None;
    let mut index_span = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            1 => {
                let index = field.varint().and_then(canonical_u32)?;
                if index >= row_index_limit {
                    return Err(DecodeError::invalid());
                }
                set_once(&mut tile_row_index, index)?;
                index_span = Some(field.span);
            },
            2 => set_once(&mut cell_count, field.varint().and_then(canonical_u32)?)?,
            3 => set_once(&mut cell_storage_buffer_pre_bnc, field.bytes()?)?,
            4 => set_once(&mut cell_offsets_pre_bnc, field.bytes()?)?,
            5 => set_once(
                &mut storage_version,
                field.varint().and_then(canonical_u32)?,
            )?,
            6 => set_once(&mut cell_storage_buffer, field.bytes()?)?,
            7 => set_once(&mut cell_offsets, field.bytes()?)?,
            8 => set_once(
                &mut has_wide_offsets,
                field.varint().and_then(canonical_bool)?,
            )?,
            _ => budget.mark_unknown_field(),
        }
    }
    let snapshot = TileRowInfoSnapshot {
        tile_row_index: tile_row_index.ok_or_else(DecodeError::invalid)?,
        cell_count: cell_count.ok_or_else(DecodeError::invalid)?,
        cell_storage_buffer_pre_bnc: cell_storage_buffer_pre_bnc
            .ok_or_else(DecodeError::invalid)?,
        cell_offsets_pre_bnc: cell_offsets_pre_bnc.ok_or_else(DecodeError::invalid)?,
        storage_version,
        cell_storage_buffer,
        cell_offsets,
        has_wide_offsets,
    };
    parity_tile_row(source, snapshot, budget, depth)?;
    Ok((snapshot, index_span.ok_or_else(DecodeError::invalid)?))
}

fn decode_header_bucket_for_rewrite(
    source: &[u8],
    row_index_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(BucketSnapshot, Vec<StagedHeader>), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut bucket_hash_function = None;
    let mut records = Vec::new();
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            1 => set_once(
                &mut bucket_hash_function,
                field.varint().and_then(canonical_u32)?,
            )?,
            2 => {
                let payload = field.bytes()?;
                let (snapshot, index_field) =
                    decode_header_in(payload, row_index_limit, budget, child_depth)?;
                if records.len() >= budget.options.max_records {
                    let observed = records
                        .len()
                        .checked_add(1)
                        .ok_or_else(DecodeError::invalid)?;
                    return Err(DecodeError::limited(DecodeLimit::Records {
                        observed,
                        maximum: budget.options.max_records,
                    }));
                }
                let requested = records
                    .len()
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                records
                    .try_reserve(1)
                    .map_err(|_| DecodeError::limited(DecodeLimit::Allocation { requested }))?;
                records.push(StagedHeader {
                    payload_start: field.span.value_start,
                    payload_end: field.span.value_end,
                    index_field,
                    snapshot,
                });
                budget.records(1)?;
            },
            _ => budget.mark_unknown_field(),
        }
    }
    let bucket = BucketSnapshot {
        bucket_hash_function: bucket_hash_function.ok_or_else(DecodeError::invalid)?,
    };
    parity_header_bucket(source, bucket, budget, depth)?;
    Ok((bucket, records))
}

fn decode_header_in(
    source: &[u8],
    row_index_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(HeaderSnapshot, FieldSpan), DecodeError> {
    budget.message(source, depth)?;
    let mut index = None;
    let mut size_bits = None;
    let mut hiding_state = None;
    let mut number_of_cells = None;
    let mut cell_style = None;
    let mut text_style = None;
    let mut index_span = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            1 => {
                let value = field.varint().and_then(canonical_u32)?;
                if value >= row_index_limit {
                    return Err(DecodeError::invalid());
                }
                set_once(&mut index, value)?;
                index_span = Some(field.span);
            },
            2 => set_once(&mut size_bits, field.fixed32()?)?,
            3 => set_once(&mut hiding_state, field.varint().and_then(canonical_u32)?)?,
            4 => set_once(
                &mut number_of_cells,
                field.varint().and_then(canonical_u32)?,
            )?,
            5 => set_once(&mut cell_style, field.bytes()?)?,
            6 => set_once(&mut text_style, field.bytes()?)?,
            _ => budget.mark_unknown_field(),
        }
    }
    let snapshot = header_snapshot(
        index.ok_or_else(DecodeError::invalid)?,
        size_bits.ok_or_else(DecodeError::invalid)?,
        hiding_state.ok_or_else(DecodeError::invalid)?,
        number_of_cells.ok_or_else(DecodeError::invalid)?,
        cell_style,
        text_style,
    );
    parity_header(source, snapshot, budget, depth)?;
    Ok((snapshot, index_span.ok_or_else(DecodeError::invalid)?))
}

fn assemble_tile_rows(
    plan: &PreparedTileRowsRewrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    try_reserve_exact(&mut output, plan.requirements.output_bytes)?;
    let mut cursor = 0usize;
    let mut row_ordinal = 0usize;
    let mut num_rows_written = false;
    let new_num_rows = plan
        .destinations
        .iter()
        .copied()
        .max()
        .map_or(Ok(0), |row| {
            row.checked_add(1).ok_or_else(DecodeError::invalid)
        })?;
    let mut remaining = plan.source;
    let mut budget = Budget::new(plan.source, options)?;
    while let Some(field) = next_field(&mut remaining, plan.source.len(), &mut budget, 1)? {
        output.extend_from_slice(&plan.source[cursor..field.span.start]);
        match field.number {
            4 => {
                if num_rows_written {
                    return Err(DecodeError::invalid());
                }
                if field.span.start != plan.num_rows_span.start
                    || field.span.end != plan.num_rows_span.end
                {
                    return Err(DecodeError::invalid());
                }
                append_replaced_varint_field(
                    &mut output,
                    &plan.source[field.span.start..field.span.end],
                    field.span,
                    new_num_rows,
                )?;
                num_rows_written = true;
            },
            5 => {
                let source_index = plan
                    .desired_order
                    .get(row_ordinal)
                    .copied()
                    .ok_or_else(DecodeError::invalid)?;
                let record = plan
                    .records
                    .get(source_index)
                    .ok_or_else(DecodeError::invalid)?;
                let destination = plan
                    .destinations
                    .get(source_index)
                    .copied()
                    .ok_or_else(DecodeError::invalid)?;
                let payload = record.payload(plan.source);
                append_replaced_length_delimited_field(
                    &mut output,
                    5,
                    payload,
                    record.index_field,
                    destination,
                    options,
                )?;
                row_ordinal = row_ordinal
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
            },
            _ => output.extend_from_slice(&plan.source[field.span.start..field.span.end]),
        }
        cursor = field.span.end;
    }
    if !num_rows_written || row_ordinal != plan.records.len() {
        return Err(DecodeError::invalid());
    }
    output.extend_from_slice(&plan.source[cursor..]);
    if output.len() != plan.requirements.output_bytes {
        return Err(DecodeError::invalid());
    }
    Ok(output)
}

fn assemble_header_rows(
    plan: &PreparedHeaderRowsRewrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    try_reserve_exact(&mut output, plan.requirements.output_bytes)?;
    let mut cursor = 0usize;
    let mut row_ordinal = 0usize;
    let mut remaining = plan.source;
    let mut budget = Budget::new(plan.source, options)?;
    while let Some(field) = next_field(&mut remaining, plan.source.len(), &mut budget, 1)? {
        output.extend_from_slice(&plan.source[cursor..field.span.start]);
        if field.number == 2 {
            let source_index = plan
                .desired_order
                .get(row_ordinal)
                .copied()
                .ok_or_else(DecodeError::invalid)?;
            let record = plan
                .records
                .get(source_index)
                .ok_or_else(DecodeError::invalid)?;
            let destination = plan
                .destinations
                .get(source_index)
                .copied()
                .ok_or_else(DecodeError::invalid)?;
            append_replaced_length_delimited_field(
                &mut output,
                2,
                record.payload(plan.source),
                record.index_field,
                destination,
                options,
            )?;
            row_ordinal = row_ordinal
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
        } else {
            output.extend_from_slice(&plan.source[field.span.start..field.span.end]);
        }
        cursor = field.span.end;
    }
    if row_ordinal != plan.records.len() {
        return Err(DecodeError::invalid());
    }
    output.extend_from_slice(&plan.source[cursor..]);
    if output.len() != plan.requirements.output_bytes {
        return Err(DecodeError::invalid());
    }
    Ok(output)
}

/// Prepare a row UID map rewrite using a source-indexed destination
/// permutation. The column arrays and sorted UUID records remain untouched.
pub fn plan_column_row_uid_map_rewrite<'source>(
    source: &'source [u8],
    column_count: usize,
    row_count: usize,
    permutation: &RowUidPermutation,
    options: DecodeOptions,
) -> Result<PreparedColumnRowUidMapRewrite<'source>, DecodeError> {
    permutation.validate_len(row_count)?;
    let mut budget = Budget::new(source, options)?;
    let (snapshot, fields) =
        decode_uid_map_for_rewrite(source, column_count, row_count, &mut budget, 1)?;
    let mut row_uid_for_index = Vec::new();
    let mut row_index_for_uid = Vec::new();
    try_reserve_exact(&mut row_uid_for_index, row_count)?;
    try_reserve_exact(&mut row_index_for_uid, row_count)?;
    row_uid_for_index.resize(row_count, 0);
    row_index_for_uid.resize(row_count, 0);
    for (source_index, &destination) in permutation.destination_by_source.iter().enumerate() {
        let destination = usize::try_from(destination).map_err(|_| DecodeError::invalid())?;
        row_uid_for_index[destination] = snapshot.row_uid_for_index[source_index];
    }
    for (destination, &uid) in row_uid_for_index.iter().enumerate() {
        let uid = usize::try_from(uid).map_err(|_| DecodeError::invalid())?;
        if uid >= row_count {
            return Err(DecodeError::invalid());
        }
        row_index_for_uid[uid] = u32::try_from(destination).map_err(|_| DecodeError::invalid())?;
    }
    let mut output_bytes = source.len();
    for field in fields
        .row_index_for_uid
        .iter()
        .chain(fields.row_uid_for_index.iter())
    {
        let old = field.original_value;
        let replacement = if field.kind == UidFieldKind::RowIndexForUid {
            let index = field.ordinal;
            row_index_for_uid[index]
        } else {
            let index = field.ordinal;
            row_uid_for_index[index]
        };
        output_bytes = output_bytes
            .checked_add(encoded_varint_len(u64::from(replacement)))
            .and_then(|value| value.checked_sub(encoded_varint_len(u64::from(old))))
            .ok_or_else(DecodeError::invalid)?;
    }
    let has_unknown_fields = budget.has_unknown_fields();
    let changed = row_index_for_uid.as_slice() != snapshot.row_index_for_uid.as_slice()
        || row_uid_for_index.as_slice() != snapshot.row_uid_for_index.as_slice();
    if has_unknown_fields && changed {
        return Err(DecodeError::invalid());
    }
    let elements = column_count
        .checked_mul(6)
        .and_then(|value| {
            row_count
                .checked_mul(6)
                .and_then(|rows| value.checked_add(rows))
        })
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = rewrite_work_upper_bound(
        source.len(),
        output_bytes,
        fields
            .row_index_for_uid
            .len()
            .checked_add(fields.row_uid_for_index.len())
            .ok_or_else(DecodeError::invalid)?,
        0,
        0,
    )?;
    let records = fields
        .row_index_for_uid
        .len()
        .checked_add(fields.row_uid_for_index.len())
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteRequirements {
        source: budget.report(),
        output_bytes,
        work_bytes,
        fields: budget.fields,
        records,
        elements,
    };
    validate_requirements(requirements, options)?;
    Ok(PreparedColumnRowUidMapRewrite {
        source,
        snapshot,
        fields,
        row_index_for_uid,
        row_uid_for_index,
        has_unknown_fields,
        requirements,
    })
}

/// Execute a validated UID-map rewrite and strictly verify both inverse
/// arrays in the candidate bytes.
pub fn execute_column_row_uid_map_rewrite(
    plan: PreparedColumnRowUidMapRewrite<'_>,
    column_count: usize,
    row_count: usize,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_requirements(plan.requirements, options)?;
    let changed_records = plan
        .fields
        .row_index_for_uid
        .iter()
        .chain(plan.fields.row_uid_for_index.iter())
        .filter(|field| {
            let replacement = if field.kind == UidFieldKind::RowIndexForUid {
                plan.row_index_for_uid[field.ordinal]
            } else {
                plan.row_uid_for_index[field.ordinal]
            };
            replacement != field.original_value
        })
        .count();
    if plan.has_unknown_fields && changed_records != 0 {
        return Err(DecodeError::invalid());
    }
    let output = assemble_uid_map(&plan, options)?;
    let (candidate, report) = decode_column_row_uid_map(&output, column_count, row_count, options)?;
    if candidate != plan.expected_snapshot() {
        return Err(DecodeError::invalid());
    }
    Ok((
        output,
        RewriteReport {
            source: plan.requirements.source,
            result: report,
            output_bytes: plan.requirements.output_bytes,
            changed_records,
        },
    ))
}

/// One-shot UID-map rewrite convenience wrapper.
pub fn rewrite_column_row_uid_map(
    source: &[u8],
    column_count: usize,
    row_count: usize,
    permutation: &RowUidPermutation,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let plan =
        plan_column_row_uid_map_rewrite(source, column_count, row_count, permutation, options)?;
    execute_column_row_uid_map_rewrite(plan, column_count, row_count, options)
}

/// Prepared source-preserving UID-map rewrite.
pub struct PreparedColumnRowUidMapRewrite<'source> {
    source: &'source [u8],
    snapshot: ColumnRowUidMapSnapshot,
    fields: UidFieldSpans,
    row_index_for_uid: Vec<u32>,
    row_uid_for_index: Vec<u32>,
    has_unknown_fields: bool,
    requirements: RewriteRequirements,
}

impl fmt::Debug for PreparedColumnRowUidMapRewrite<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedColumnRowUidMapRewrite")
            .field("row_count", &self.snapshot.row_count())
            .field("column_count", &self.snapshot.column_count())
            .field("requirements", &self.requirements)
            .finish()
    }
}

impl PreparedColumnRowUidMapRewrite<'_> {
    #[must_use]
    pub const fn requirements(&self) -> RewriteRequirements {
        self.requirements
    }

    #[must_use]
    pub const fn source_snapshot(&self) -> &ColumnRowUidMapSnapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn row_index_for_uid(&self) -> &[u32] {
        &self.row_index_for_uid
    }

    #[must_use]
    pub fn row_uid_for_index(&self) -> &[u32] {
        &self.row_uid_for_index
    }

    fn expected_snapshot(&self) -> ColumnRowUidMapSnapshot {
        let mut expected = self.snapshot.clone();
        expected
            .row_index_for_uid
            .clone_from(&self.row_index_for_uid);
        expected
            .row_uid_for_index
            .clone_from(&self.row_uid_for_index);
        expected
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum UidFieldKind {
    RowIndexForUid,
    RowUidForIndex,
}

#[derive(Clone, Copy)]
struct StagedUidField {
    ordinal: usize,
    original_value: u32,
    kind: UidFieldKind,
}

struct UidFieldSpans {
    row_index_for_uid: Vec<StagedUidField>,
    row_uid_for_index: Vec<StagedUidField>,
}

fn decode_uid_map_for_rewrite(
    source: &[u8],
    column_count: usize,
    row_count: usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<(ColumnRowUidMapSnapshot, UidFieldSpans), DecodeError> {
    let snapshot = decode_uid_map_in(source, column_count, row_count, budget, depth)?;
    let mut row_index_for_uid = Vec::new();
    let mut row_uid_for_index = Vec::new();
    try_reserve_exact(&mut row_index_for_uid, row_count)?;
    try_reserve_exact(&mut row_uid_for_index, row_count)?;
    let mut row_index_ordinal = 0usize;
    let mut row_uid_ordinal = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            5 => {
                let value = field.varint().and_then(canonical_u32)?;
                let ordinal = row_index_ordinal;
                row_index_ordinal = row_index_ordinal
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                row_index_for_uid.push(StagedUidField {
                    ordinal,
                    original_value: value,
                    kind: UidFieldKind::RowIndexForUid,
                });
            },
            6 => {
                let value = field.varint().and_then(canonical_u32)?;
                let ordinal = row_uid_ordinal;
                row_uid_ordinal = row_uid_ordinal
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                row_uid_for_index.push(StagedUidField {
                    ordinal,
                    original_value: value,
                    kind: UidFieldKind::RowUidForIndex,
                });
            },
            _ => {},
        }
    }
    if row_index_for_uid.len() != row_count || row_uid_for_index.len() != row_count {
        return Err(DecodeError::invalid());
    }
    Ok((
        snapshot,
        UidFieldSpans {
            row_index_for_uid,
            row_uid_for_index,
        },
    ))
}

fn decode_uid_map_in(
    source: &[u8],
    column_count: usize,
    row_count: usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<ColumnRowUidMapSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let elements_upper = column_count
        .checked_mul(3)
        .and_then(|value| {
            row_count
                .checked_mul(3)
                .and_then(|rows| value.checked_add(rows))
        })
        .ok_or_else(DecodeError::invalid)?;
    if elements_upper > budget.options.max_elements {
        return Err(DecodeError::limited(DecodeLimit::Elements {
            observed: elements_upper,
            maximum: budget.options.max_elements,
        }));
    }
    let mut sorted_column_uids = Vec::new();
    let mut column_index_for_uid = Vec::new();
    let mut column_uid_for_index = Vec::new();
    let mut sorted_row_uids = Vec::new();
    let mut row_index_for_uid = Vec::new();
    let mut row_uid_for_index = Vec::new();
    try_reserve_exact(&mut sorted_column_uids, column_count)?;
    try_reserve_exact(&mut column_index_for_uid, column_count)?;
    try_reserve_exact(&mut column_uid_for_index, column_count)?;
    try_reserve_exact(&mut sorted_row_uids, row_count)?;
    try_reserve_exact(&mut row_index_for_uid, row_count)?;
    try_reserve_exact(&mut row_uid_for_index, row_count)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            1 => {
                if sorted_column_uids.len() >= column_count {
                    return Err(DecodeError::invalid());
                }
                let uuid_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
                let uuid = decode_uuid(field.bytes()?, budget, uuid_depth)?;
                sorted_column_uids.push(uuid);
                budget.element(1)?;
            },
            2 => {
                if column_index_for_uid.len() >= column_count {
                    return Err(DecodeError::invalid());
                }
                column_index_for_uid.push(field.varint().and_then(canonical_u32)?);
                budget.element(1)?;
            },
            3 => {
                if column_uid_for_index.len() >= column_count {
                    return Err(DecodeError::invalid());
                }
                column_uid_for_index.push(field.varint().and_then(canonical_u32)?);
                budget.element(1)?;
            },
            4 => {
                if sorted_row_uids.len() >= row_count {
                    return Err(DecodeError::invalid());
                }
                let uuid_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
                let uuid = decode_uuid(field.bytes()?, budget, uuid_depth)?;
                sorted_row_uids.push(uuid);
                budget.element(1)?;
            },
            5 => {
                if row_index_for_uid.len() >= row_count {
                    return Err(DecodeError::invalid());
                }
                row_index_for_uid.push(field.varint().and_then(canonical_u32)?);
                budget.element(1)?;
            },
            6 => {
                if row_uid_for_index.len() >= row_count {
                    return Err(DecodeError::invalid());
                }
                row_uid_for_index.push(field.varint().and_then(canonical_u32)?);
                budget.element(1)?;
            },
            _ => budget.mark_unknown_field(),
        }
    }
    if sorted_column_uids.len() != column_count
        || column_index_for_uid.len() != column_count
        || column_uid_for_index.len() != column_count
        || sorted_row_uids.len() != row_count
        || row_index_for_uid.len() != row_count
        || row_uid_for_index.len() != row_count
    {
        return Err(DecodeError::invalid());
    }
    validate_uuid_uniqueness(&sorted_column_uids)?;
    validate_uuid_uniqueness(&sorted_row_uids)?;
    validate_inverse_pair(&column_index_for_uid, &column_uid_for_index, column_count)?;
    validate_inverse_pair(&row_index_for_uid, &row_uid_for_index, row_count)?;
    parity_uid_map(source, budget, depth)?;
    Ok(ColumnRowUidMapSnapshot {
        sorted_column_uids,
        column_index_for_uid,
        column_uid_for_index,
        sorted_row_uids,
        row_index_for_uid,
        row_uid_for_index,
    })
}

fn assemble_uid_map(
    plan: &PreparedColumnRowUidMapRewrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    try_reserve_exact(&mut output, plan.requirements.output_bytes)?;
    let mut cursor = 0usize;
    let mut row_index_ordinal = 0usize;
    let mut row_uid_ordinal = 0usize;
    let mut remaining = plan.source;
    let mut budget = Budget::new(plan.source, options)?;
    while let Some(field) = next_field(&mut remaining, plan.source.len(), &mut budget, 1)? {
        output.extend_from_slice(&plan.source[cursor..field.span.start]);
        let replacement = match field.number {
            5 => {
                let ordinal = row_index_ordinal;
                row_index_ordinal = row_index_ordinal
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                Some(
                    *plan
                        .row_index_for_uid
                        .get(ordinal)
                        .ok_or_else(DecodeError::invalid)?,
                )
            },
            6 => {
                let ordinal = row_uid_ordinal;
                row_uid_ordinal = row_uid_ordinal
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                Some(
                    *plan
                        .row_uid_for_index
                        .get(ordinal)
                        .ok_or_else(DecodeError::invalid)?,
                )
            },
            _ => None,
        };
        if let Some(value) = replacement {
            append_replaced_varint_field(
                &mut output,
                &plan.source[field.span.start..field.span.end],
                field.span,
                value,
            )?;
        } else {
            output.extend_from_slice(&plan.source[field.span.start..field.span.end]);
        }
        cursor = field.span.end;
    }
    if row_index_ordinal != plan.fields.row_index_for_uid.len()
        || row_uid_ordinal != plan.fields.row_uid_for_index.len()
    {
        return Err(DecodeError::invalid());
    }
    output.extend_from_slice(&plan.source[cursor..]);
    if output.len() != plan.requirements.output_bytes {
        return Err(DecodeError::invalid());
    }
    Ok(output)
}

fn decode_uuid(source: &[u8], budget: &mut Budget, depth: u32) -> Result<Uuid, DecodeError> {
    budget.message(source, depth)?;
    let mut lower = None;
    let mut upper = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, source.len(), budget, depth)? {
        match field.number {
            1 => set_once(&mut lower, field.varint()?)?,
            2 => set_once(&mut upper, field.varint()?)?,
            _ => budget.mark_unknown_field(),
        }
    }
    let uuid = Uuid::new(
        lower.ok_or_else(DecodeError::invalid)?,
        upper.ok_or_else(DecodeError::invalid)?,
    );
    let view: projection::UuidLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.lower != uuid.lower || view.upper != uuid.upper {
        return Err(DecodeError::invalid());
    }
    Ok(uuid)
}

fn validate_uuid_uniqueness(values: &[Uuid]) -> Result<(), DecodeError> {
    let mut sorted = Vec::new();
    try_reserve_exact(&mut sorted, values.len())?;
    sorted.extend_from_slice(values);
    sorted.sort_unstable_by_key(|value| (value.lower, value.upper));
    if sorted.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn validate_inverse_pair(
    index_for_uid: &[u32],
    uid_for_index: &[u32],
    count: usize,
) -> Result<(), DecodeError> {
    if index_for_uid.len() != count || uid_for_index.len() != count {
        return Err(DecodeError::invalid());
    }
    for (uid, &index_value) in index_for_uid.iter().enumerate() {
        let index = usize::try_from(index_value).map_err(|_| DecodeError::invalid())?;
        let uid = u32::try_from(uid).map_err(|_| DecodeError::invalid())?;
        if index >= count || uid_for_index[index] != uid {
            return Err(DecodeError::invalid());
        }
    }
    Ok(())
}

fn parity_tile(
    source: &[u8],
    snapshot: TileSnapshot,
    budget: &mut Budget,
    _depth: u32,
) -> Result<(), DecodeError> {
    let view: projection::TileArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.max_column != snapshot.max_column
        || view.max_row != snapshot.max_row
        || view.num_cells != snapshot.num_cells
        || view.num_rows != snapshot.num_rows
        || view.storage_version != snapshot.storage_version
        || view.last_saved_in_bnc != snapshot.last_saved_in_bnc
        || view.should_use_wide_rows != snapshot.should_use_wide_rows
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parity_tile_row(
    source: &[u8],
    snapshot: TileRowInfoSnapshot<'_>,
    budget: &mut Budget,
    _depth: u32,
) -> Result<(), DecodeError> {
    let view: projection::TileRowInfoArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.tile_row_index != snapshot.tile_row_index
        || view.cell_count != snapshot.cell_count
        || view.cell_storage_buffer_pre_bnc != snapshot.cell_storage_buffer_pre_bnc
        || view.cell_offsets_pre_bnc != snapshot.cell_offsets_pre_bnc
        || view.storage_version != snapshot.storage_version
        || view.cell_storage_buffer != snapshot.cell_storage_buffer
        || view.cell_offsets != snapshot.cell_offsets
        || view.has_wide_offsets != snapshot.has_wide_offsets
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parity_header_bucket(
    source: &[u8],
    snapshot: BucketSnapshot,
    budget: &mut Budget,
    _depth: u32,
) -> Result<(), DecodeError> {
    let view: projection::HeaderStorageBucketArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.bucket_hash_function != snapshot.bucket_hash_function {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parity_header(
    source: &[u8],
    snapshot: HeaderSnapshot,
    budget: &mut Budget,
    _depth: u32,
) -> Result<(), DecodeError> {
    let view: projection::HeaderArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.index != snapshot.index
        || view.size_bits != snapshot.size_bits
        || view.hiding_state != snapshot.hiding_state
        || view.number_of_cells != snapshot.number_of_cells
        || view.cell_style.is_some() != snapshot.has_cell_style
        || view.text_style.is_some() != snapshot.has_text_style
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parity_uid_map(source: &[u8], budget: &mut Budget, _depth: u32) -> Result<(), DecodeError> {
    let _view: projection::ColumnRowUidMapArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    Ok(())
}

fn validate_move_list(
    moves: &[RowMove],
    row_index_limit: u32,
    available_indices: &[u32],
) -> Result<(), DecodeError> {
    if moves.len() > available_indices.len() {
        return Err(DecodeError::invalid());
    }
    if moves.iter().any(|movement| {
        movement.source_index >= row_index_limit || movement.destination_index >= row_index_limit
    }) {
        return Err(DecodeError::invalid());
    }
    if available_indices.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(DecodeError::invalid());
    }
    if moves.iter().any(|movement| {
        available_indices
            .binary_search(&movement.source_index)
            .is_err()
    }) {
        return Err(DecodeError::invalid());
    }
    if moves
        .windows(2)
        .any(|pair| pair[0].source_index == pair[1].source_index)
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn move_for(moves: &[RowMove], source_index: u32) -> Option<u32> {
    moves
        .binary_search_by_key(&source_index, |movement| movement.source_index)
        .ok()
        .map(|position| moves[position].destination_index)
}

fn validate_values_within_limit(values: &[u32], row_index_limit: u32) -> Result<(), DecodeError> {
    if values.iter().any(|value| *value >= row_index_limit) {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn validate_permutation(values: &[u32]) -> Result<(), DecodeError> {
    let mut seen = Vec::new();
    try_reserve_exact(&mut seen, values.len())?;
    seen.resize(values.len(), false);
    for value in values.iter().copied() {
        let value = usize::try_from(value).map_err(|_| DecodeError::invalid())?;
        if value >= values.len() || seen[value] {
            return Err(DecodeError::invalid());
        }
        seen[value] = true;
    }
    Ok(())
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), DecodeError> {
    if slot.replace(value).is_some() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn validate_requirements(
    requirements: RewriteRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if requirements.output_bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    if requirements.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if requirements.fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: options.max_fields,
        }));
    }
    if requirements.records > options.max_records {
        return Err(DecodeError::limited(DecodeLimit::Records {
            observed: requirements.records,
            maximum: options.max_records,
        }));
    }
    if requirements.elements > options.max_elements {
        return Err(DecodeError::limited(DecodeLimit::Elements {
            observed: requirements.elements,
            maximum: options.max_elements,
        }));
    }
    let scratch = requirements
        .records
        .checked_mul(size_of::<FieldSpan>())
        .and_then(|bytes| {
            requirements
                .elements
                .checked_mul(size_of_u32())
                .and_then(|elements| bytes.checked_add(elements))
        })
        .ok_or_else(DecodeError::invalid)?;
    if scratch > options.max_scratch_bytes {
        return Err(DecodeError::limited(DecodeLimit::ScratchBytes {
            observed: scratch,
            maximum: options.max_scratch_bytes,
        }));
    }
    Ok(())
}

const fn size_of_u32() -> usize {
    size_of::<u32>()
}

fn rewrite_work_upper_bound(
    source_bytes: usize,
    output_bytes: usize,
    record_count: usize,
    edit_count: usize,
    changed_row_bytes: usize,
) -> Result<usize, DecodeError> {
    source_bytes
        .checked_add(output_bytes)
        .and_then(|work| work.checked_add(record_count))
        .and_then(|work| work.checked_add(edit_count))
        .and_then(|work| work.checked_add(changed_row_bytes))
        .ok_or_else(DecodeError::invalid)
}

fn nested_varint_rewrite_sizes(
    payload_len: usize,
    original_index_len: usize,
    replacement_index_len: usize,
) -> Result<(usize, usize, usize), DecodeError> {
    let new_payload_len = payload_len
        .checked_add(replacement_index_len)
        .and_then(|value| value.checked_sub(original_index_len))
        .ok_or_else(DecodeError::invalid)?;
    let original_length =
        encoded_varint_len(u64::try_from(payload_len).map_err(|_| DecodeError::invalid())?);
    let replacement_length =
        encoded_varint_len(u64::try_from(new_payload_len).map_err(|_| DecodeError::invalid())?);
    Ok((new_payload_len, original_length, replacement_length))
}

fn replace_size_parts(
    current: usize,
    original: [usize; 2],
    replacement: [usize; 2],
) -> Result<usize, DecodeError> {
    let original_total = original[0]
        .checked_add(original[1])
        .ok_or_else(DecodeError::invalid)?;
    let replacement_total = replacement[0]
        .checked_add(replacement[1])
        .ok_or_else(DecodeError::invalid)?;
    if replacement_total >= original_total {
        current
            .checked_add(replacement_total - original_total)
            .ok_or_else(DecodeError::invalid)
    } else {
        current
            .checked_sub(original_total - replacement_total)
            .ok_or_else(DecodeError::invalid)
    }
}

fn try_reserve_exact<T>(output: &mut Vec<T>, additional: usize) -> Result<(), DecodeError> {
    let requested = output
        .len()
        .checked_add(additional)
        .ok_or_else(DecodeError::invalid)?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| DecodeError::limited(DecodeLimit::Allocation { requested }))?;
    if output.capacity() != requested {
        return Err(DecodeError::limited(DecodeLimit::Allocation { requested }));
    }
    Ok(())
}

fn append_replaced_varint_field(
    output: &mut Vec<u8>,
    original: &[u8],
    span: FieldSpan,
    replacement: u32,
) -> Result<(), DecodeError> {
    let local_value_start = span
        .value_start
        .checked_sub(span.start)
        .ok_or_else(DecodeError::invalid)?;
    let local_value_end = span
        .value_end
        .checked_sub(span.start)
        .ok_or_else(DecodeError::invalid)?;
    if local_value_end > original.len() || local_value_start > local_value_end {
        return Err(DecodeError::invalid());
    }
    output.extend_from_slice(&original[..local_value_start]);
    encode_varint(output, u64::from(replacement));
    output.extend_from_slice(&original[local_value_end..]);
    Ok(())
}

fn append_replaced_length_delimited_field(
    output: &mut Vec<u8>,
    field_number: u32,
    source: &[u8],
    index_field: FieldSpan,
    replacement: u32,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let local_value_start = index_field
        .value_start
        .checked_sub(index_field.start)
        .ok_or_else(DecodeError::invalid)?;
    let local_value_end = index_field
        .value_end
        .checked_sub(index_field.start)
        .ok_or_else(DecodeError::invalid)?;
    if local_value_end > source.len() || local_value_start > local_value_end {
        return Err(DecodeError::invalid());
    }
    let original_index_len = local_value_end
        .checked_sub(local_value_start)
        .ok_or_else(DecodeError::invalid)?;
    let (new_payload_len, _, _) = nested_varint_rewrite_sizes(
        source.len(),
        original_index_len,
        encoded_varint_len(u64::from(replacement)),
    )?;
    if new_payload_len > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: new_payload_len,
            maximum: options.max_message_bytes,
        }));
    }
    encode_varint(output, (u64::from(field_number) << 3) | 2);
    encode_varint(
        output,
        u64::try_from(new_payload_len).map_err(|_| DecodeError::invalid())?,
    );
    output.extend_from_slice(&source[..local_value_start]);
    encode_varint(output, u64::from(replacement));
    output.extend_from_slice(&source[local_value_end..]);
    Ok(())
}

fn encode_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn canonical_u32(value: u64) -> Result<u32, DecodeError> {
    u32::try_from(value).map_err(|_| DecodeError::invalid())
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid()),
    }
}

struct Budget {
    options: DecodeOptions,
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    records: usize,
    elements: usize,
    has_unknown_fields: bool,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_limit =
            usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError::invalid())?;
        if options.max_message_bytes > hard_limit {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_limit,
            }));
        }
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION,
            }));
        }
        Ok(Self {
            options,
            source_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            records: 0,
            elements: 0,
            has_unknown_fields: false,
        })
    }

    fn mark_unknown_field(&mut self) {
        self.has_unknown_fields = true;
    }

    #[must_use]
    const fn has_unknown_fields(&self) -> bool {
        self.has_unknown_fields
    }

    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        if source.len() > self.options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: self.options.max_message_bytes,
            }));
        }
        self.observe_depth(depth)?;
        self.work(source.len())
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn records(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .records
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_records {
            return Err(DecodeError::limited(DecodeLimit::Records {
                observed,
                maximum: self.options.max_records,
            }));
        }
        self.records = observed;
        Ok(())
    }

    fn element(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .elements
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_elements {
            return Err(DecodeError::limited(DecodeLimit::Elements {
                observed,
                maximum: self.options.max_elements,
            }));
        }
        self.elements = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            records: self.records,
            elements: self.elements,
        }
    }
}

#[derive(Clone, Copy)]
enum Value<'source> {
    Varint(u64),
    Fixed64,
    Bytes(&'source [u8]),
    Group,
    Fixed32(u32),
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire_type: u8,
    value: Value<'source>,
    span: FieldSpan,
}

impl<'source> Field<'source> {
    fn varint(self) -> Result<u64, DecodeError> {
        match (self.wire_type, self.value) {
            (0, Value::Varint(value)) => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    fn fixed32(self) -> Result<u32, DecodeError> {
        match (self.wire_type, self.value) {
            (5, Value::Fixed32(value)) => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        match (self.wire_type, self.value) {
            (2, Value::Bytes(value)) => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }
}

enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

fn next_field<'source>(
    source: &mut &'source [u8],
    root_len: usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, DecodeError> {
    match parse_field(source, root_len, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(_)) => Err(DecodeError::invalid()),
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    root_len: usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    budget.field()?;
    let start = root_len
        .checked_sub(source.len())
        .ok_or_else(DecodeError::invalid)?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_| DecodeError::invalid())?;
    let wire_type = u8::try_from(tag & 7).map_err(|_| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let value_start;
    let value = match wire_type {
        0 => {
            value_start = root_len
                .checked_sub(source.len())
                .ok_or_else(DecodeError::invalid)?;
            let value = take_varint(source)?;
            Value::Varint(value)
        },
        1 => {
            value_start = root_len
                .checked_sub(source.len())
                .ok_or_else(DecodeError::invalid)?;
            let _ = take(source, 8)?;
            Value::Fixed64
        },
        2 => {
            let length =
                usize::try_from(take_varint(source)?).map_err(|_| DecodeError::invalid())?;
            value_start = root_len
                .checked_sub(source.len())
                .ok_or_else(DecodeError::invalid)?;
            Value::Bytes(take(source, length)?)
        },
        3 => {
            value_start = root_len
                .checked_sub(source.len())
                .ok_or_else(DecodeError::invalid)?;
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            skip_group(source, number, root_len, budget, child_depth)?;
            Value::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            value_start = root_len
                .checked_sub(source.len())
                .ok_or_else(DecodeError::invalid)?;
            Value::Fixed32(u32::from_le_bytes(
                take(source, 4)?
                    .try_into()
                    .map_err(|_| DecodeError::invalid())?,
            ))
        },
        _ => return Err(DecodeError::invalid()),
    };
    let end = root_len
        .checked_sub(source.len())
        .ok_or_else(DecodeError::invalid)?;
    Ok(Some(ParseItem::Field(Field {
        number,
        wire_type,
        value,
        span: FieldSpan {
            start,
            end,
            value_start,
            value_end: end,
        },
    })))
}

fn skip_group(
    source: &mut &'_ [u8],
    expected: u32,
    root_len: usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, root_len, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(_)) | None => return Err(DecodeError::invalid()),
        }
    }
}

fn take<'source>(source: &mut &'source [u8], amount: usize) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(DecodeError::invalid());
    }
    let (selected, remaining) = source.split_at(amount);
    *source = remaining;
    Ok(selected)
}

fn take_varint(source: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(DecodeError::invalid)?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            if encoded_varint_len(value) != consumed {
                return Err(DecodeError::invalid());
            }
            *source = &original[consumed..];
            return Ok(value);
        }
    }
    Err(DecodeError::invalid())
}

// Keep a small in-module test marker as well as the crate-level integration
// tests.  The physical-sort boundary is intentionally source-local: a future
// checker must not be able to mistake an unrelated downstream test for proof
// that this strict parser still rejects malformed mutable roots or still runs
// the private Buffa parity pass.
#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::bounded()
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn field_varint(output: &mut Vec<u8>, field: u32, value: u64) {
        push_varint(output, u64::from(field) << 3);
        push_varint(output, value);
    }

    fn field_bytes(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
        push_varint(output, (u64::from(field) << 3) | 2);
        push_varint(
            output,
            u64::try_from(payload.len()).expect("test payload length fits"),
        );
        output.extend_from_slice(payload);
    }

    fn field_fixed32(output: &mut Vec<u8>, field: u32, value: u32) {
        push_varint(output, (u64::from(field) << 3) | 5);
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn group_field(field: u32, body: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 3);
        output.extend_from_slice(body);
        push_varint(&mut output, (u64::from(field) << 3) | 4);
        output
    }

    fn tile_row(index: u32, marker: u8) -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, 1, u64::from(index));
        field_varint(&mut output, 2, 1);
        field_bytes(&mut output, 3, &[marker]);
        field_bytes(&mut output, 4, &[marker, 0x42]);
        // This extension is not part of the generated projection.  Its raw
        // bytes must remain attached to the row during a physical reorder.
        field_bytes(&mut output, 90, &[0xe0, marker]);
        output
    }

    fn tile_with_rows(rows: &[Vec<u8>]) -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, 1, 3);
        field_varint(&mut output, 2, 2);
        let row_count = u64::try_from(rows.len()).expect("test row count fits");
        field_varint(&mut output, 3, row_count);
        field_varint(&mut output, 4, row_count);
        field_bytes(&mut output, 90, b"tile-root-unknown");
        output.extend_from_slice(&group_field(91, &[0x08, 0x01]));
        for row in rows {
            field_bytes(&mut output, 5, row);
        }
        output
    }

    fn tile() -> Vec<u8> {
        tile_with_rows(&[tile_row(0, b'a'), tile_row(1, b'b')])
    }

    fn header(index: u32, marker: u8) -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, 1, u64::from(index));
        field_fixed32(&mut output, 2, u32::from(marker));
        field_varint(&mut output, 3, 0);
        field_varint(&mut output, 4, 1);
        field_bytes(&mut output, 5, &[0x10, marker]);
        field_bytes(&mut output, 6, &[0x20, marker]);
        field_bytes(&mut output, 77, &[0x70, marker]);
        output
    }

    fn header_bucket() -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, 1, 4);
        field_bytes(&mut output, 90, b"header-root-unknown");
        output.extend_from_slice(&group_field(91, &[0x08, 0x02]));
        field_bytes(&mut output, 2, &header(0, b'x'));
        field_bytes(&mut output, 2, &header(1, b'y'));
        output
    }

    fn uuid(lower: u64, upper: u64) -> Vec<u8> {
        let mut output = Vec::new();
        field_varint(&mut output, 1, lower);
        field_varint(&mut output, 2, upper);
        field_bytes(&mut output, 90, b"uuid-unknown");
        output
    }

    fn uid_map() -> Vec<u8> {
        let mut output = Vec::new();
        for (lower, upper) in [(11, 101), (12, 102)] {
            field_bytes(&mut output, 1, &uuid(lower, upper));
        }
        for value in [0, 1] {
            field_varint(&mut output, 2, value);
            field_varint(&mut output, 3, value);
        }
        for (lower, upper) in [(21, 201), (22, 202)] {
            field_bytes(&mut output, 4, &uuid(lower, upper));
        }
        for value in [0, 1] {
            field_varint(&mut output, 5, value);
            field_varint(&mut output, 6, value);
        }
        output.extend_from_slice(&group_field(90, &[0x08, 0x01]));
        field_bytes(&mut output, 91, b"uid-root-unknown");
        output
    }

    fn contains_bytes(source: &[u8], needle: &[u8]) -> bool {
        !needle.is_empty() && source.windows(needle.len()).any(|window| window == needle)
    }

    #[track_caller]
    fn rejected<T>(result: Result<T, DecodeError>) {
        assert!(
            result.is_err(),
            "malformed physical-sort source was accepted at {}",
            std::panic::Location::caller()
        );
    }

    #[test]
    fn strict_mutable_roots_reject_duplicate_wrong_wire_and_noncanonical_fields() {
        let source = tile();

        let mut duplicate_tile = source.clone();
        field_varint(&mut duplicate_tile, 1, 4);
        rejected(plan_tile_rows_rewrite(&duplicate_tile, 8, &[], options()));

        let mut wrong_wire_tile = source.clone();
        field_bytes(&mut wrong_wire_tile, 1, &[4]);
        rejected(plan_tile_rows_rewrite(&wrong_wire_tile, 8, &[], options()));

        let mut duplicate_row = tile_row(0, b'a');
        field_varint(&mut duplicate_row, 1, 0);
        rejected(plan_tile_rows_rewrite(
            &tile_with_rows(&[duplicate_row, tile_row(1, b'b')]),
            8,
            &[],
            options(),
        ));

        let mut wrong_wire_row = tile_row(0, b'a');
        field_bytes(&mut wrong_wire_row, 1, &[0]);
        rejected(plan_tile_rows_rewrite(
            &tile_with_rows(&[wrong_wire_row, tile_row(1, b'b')]),
            8,
            &[],
            options(),
        ));

        // Field 90 is unknown to both the handwritten model and its lazy
        // projection, but its key, length, and scalar framing remain strict.
        let mut overlong_unknown_key = source.clone();
        overlong_unknown_key.extend_from_slice(&[0xd2, 0x85, 0x00, 0x00]);
        rejected(plan_tile_rows_rewrite(
            &overlong_unknown_key,
            8,
            &[],
            options(),
        ));
        let mut overlong_unknown_length = source.clone();
        overlong_unknown_length.extend_from_slice(&[0xd2, 0x05, 0x80, 0x00]);
        rejected(plan_tile_rows_rewrite(
            &overlong_unknown_length,
            8,
            &[],
            options(),
        ));
        let mut overlong_unknown_value = source.clone();
        overlong_unknown_value.extend_from_slice(&[0xd0, 0x05, 0x80, 0x00]);
        rejected(plan_tile_rows_rewrite(
            &overlong_unknown_value,
            8,
            &[],
            options(),
        ));
        let mut unsupported_unknown_wire = source;
        unsupported_unknown_wire.extend_from_slice(&[0xd6, 0x05, 0x00]);
        rejected(plan_tile_rows_rewrite(
            &unsupported_unknown_wire,
            8,
            &[],
            options(),
        ));

        let header_source = header_bucket();
        let mut duplicate_header = header_source.clone();
        field_varint(&mut duplicate_header, 1, 5);
        rejected(plan_header_storage_bucket_rows(
            &duplicate_header,
            8,
            &[],
            options(),
        ));
        let mut wrong_wire_header = header_source.clone();
        field_bytes(&mut wrong_wire_header, 1, &[5]);
        rejected(plan_header_storage_bucket_rows(
            &wrong_wire_header,
            8,
            &[],
            options(),
        ));
        let mut malformed_header = Vec::new();
        field_varint(&mut malformed_header, 1, 4);
        field_bytes(&mut malformed_header, 2, &[0x08, 0x00]);
        rejected(plan_header_storage_bucket_rows(
            &malformed_header,
            8,
            &[],
            options(),
        ));

        let uid_source = uid_map();
        let mut duplicate_uid_root = uid_source.clone();
        field_varint(&mut duplicate_uid_root, 5, 0);
        rejected(decode_column_row_uid_map(
            &duplicate_uid_root,
            1,
            2,
            options(),
        ));
        let mut wrong_wire_uid_root = uid_source.clone();
        field_bytes(&mut wrong_wire_uid_root, 5, &[0]);
        rejected(decode_column_row_uid_map(
            &wrong_wire_uid_root,
            1,
            2,
            options(),
        ));
        let mut duplicate_uuid = Vec::new();
        field_varint(&mut duplicate_uuid, 1, 11);
        field_varint(&mut duplicate_uuid, 1, 12);
        field_varint(&mut duplicate_uuid, 2, 101);
        let mut duplicate_uuid_map = Vec::new();
        field_bytes(&mut duplicate_uuid_map, 1, &duplicate_uuid);
        for value in [0, 1] {
            field_varint(&mut duplicate_uuid_map, 2, value);
            field_varint(&mut duplicate_uuid_map, 3, value);
        }
        for (lower, upper) in [(21, 201), (22, 202)] {
            field_bytes(&mut duplicate_uuid_map, 4, &uuid(lower, upper));
        }
        for value in [0, 1] {
            field_varint(&mut duplicate_uuid_map, 5, value);
            field_varint(&mut duplicate_uuid_map, 6, value);
        }
        rejected(decode_column_row_uid_map(
            &duplicate_uuid_map,
            1,
            2,
            options(),
        ));
    }

    #[test]
    fn unknown_mutable_fields_are_preserved_for_no_ops_and_rejected_for_rewrites() {
        let tile_source = tile();
        let tile_root_unknown = field_bytes_vec(90, b"tile-root-unknown");
        let tile_group_unknown = group_field(91, &[0x08, 0x01]);
        let (tile_no_op, tile_report) = rewrite_tile_rows(&tile_source, 8, &[], options())
            .expect("canonical tile with unknown roots should support a no-op");
        assert_eq!(tile_no_op, tile_source);
        assert_eq!(tile_report.changed_records(), 0);
        assert!(contains_bytes(&tile_no_op, &tile_root_unknown));
        assert!(contains_bytes(&tile_no_op, &tile_group_unknown));
        let tile_identity_moves = [RowMove::new(0, 0), RowMove::new(1, 1)];
        let (tile_identity, _) =
            rewrite_tile_rows(&tile_source, 8, &tile_identity_moves, options())
                .expect("identity tile moves should remain a no-op");
        assert_eq!(tile_identity, tile_source);
        assert!(
            rewrite_tile_rows(
                &tile_source,
                8,
                &[RowMove::new(0, 1), RowMove::new(1, 0)],
                options(),
            )
            .is_err()
        );
        let tile_plan = plan_tile_rows_rewrite(&tile_source, 8, &[], options())
            .expect("unknown tile should pass strict and lazy parity for a no-op");
        let rows: Vec<_> = tile_plan.row_records().collect();
        assert_eq!(rows[0].snapshot().cell_storage_buffer_pre_bnc()[0], b'a');
        assert_eq!(rows[1].snapshot().cell_storage_buffer_pre_bnc()[0], b'b');

        let header_source = header_bucket();
        let header_root_unknown = field_bytes_vec(90, b"header-root-unknown");
        let header_group_unknown = group_field(91, &[0x08, 0x02]);
        let (header_no_op, header_report) =
            rewrite_header_storage_bucket_rows(&header_source, 8, &[], options())
                .expect("canonical headers with unknown roots should support a no-op");
        assert_eq!(header_no_op, header_source);
        assert_eq!(header_report.changed_records(), 0);
        assert!(contains_bytes(&header_no_op, &header_root_unknown));
        assert!(contains_bytes(&header_no_op, &header_group_unknown));
        let header_identity_moves = [HeaderRowMove::new(0, 0), HeaderRowMove::new(1, 1)];
        let (header_identity, _) = rewrite_header_storage_bucket_rows(
            &header_source,
            8,
            &header_identity_moves,
            options(),
        )
        .expect("identity header moves should remain a no-op");
        assert_eq!(header_identity, header_source);
        assert!(
            rewrite_header_storage_bucket_rows(
                &header_source,
                8,
                &[HeaderRowMove::new(0, 1), HeaderRowMove::new(1, 0)],
                options(),
            )
            .is_err()
        );
        plan_header_storage_bucket_rows(&header_source, 8, &[], options())
            .expect("rewritten headers should pass strict and lazy parity");

        let uid_source = uid_map();
        let uid_root_unknown = field_bytes_vec(91, b"uid-root-unknown");
        let uid_group_unknown = group_field(90, &[0x08, 0x01]);
        let identity = RowUidPermutation::identity(2).expect("identity row permutation");
        let (uid_no_op, uid_report) =
            rewrite_column_row_uid_map(&uid_source, 2, 2, &identity, options())
                .expect("canonical UID map with unknown roots should support a no-op");
        assert_eq!(uid_no_op, uid_source);
        assert_eq!(uid_report.changed_records(), 0);
        assert!(contains_bytes(&uid_no_op, &uid_root_unknown));
        assert!(contains_bytes(&uid_no_op, &uid_group_unknown));
        let permutation = RowUidPermutation::new(&[1, 0]).expect("valid row permutation");
        assert!(rewrite_column_row_uid_map(&uid_source, 2, 2, &permutation, options()).is_err());
        decode_column_row_uid_map(&uid_source, 2, 2, options())
            .expect("unknown UID map should pass strict and lazy parity for a no-op");
    }

    #[test]
    fn lazy_buffa_parity_rejects_a_handwritten_snapshot_mismatch() {
        let source = tile();
        let mut budget = Budget::new(&source, options()).expect("bounded source budget");
        let snapshot = TileSnapshot {
            max_column: 999,
            max_row: 2,
            num_cells: 2,
            num_rows: 2,
            storage_version: None,
            last_saved_in_bnc: None,
            should_use_wide_rows: None,
        };
        assert!(parity_tile(&source, snapshot, &mut budget, 1).is_err());
    }

    fn field_bytes_vec(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        field_bytes(&mut output, field, payload);
        output
    }
}
