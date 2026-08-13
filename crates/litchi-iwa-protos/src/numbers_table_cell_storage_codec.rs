//! Strict generated-free Numbers table-cell storage projections.
//!
//! Handwritten routing owns canonical wire validation, aggregate resource
//! accounting, and repeated-field streaming. Private Buffa lazy views are
//! forced only as borrowed parity oracles; generated values never escape and
//! caller-owned bytes remain authoritative.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers stay beside the generated-free snapshots they construct."
)]

use core::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_sheet_order_generated::LitchiIwaProjection as reference_projection;
use crate::buffa_numbers_table_cell_storage_generated::LitchiIwaTableCellProjection as projection;

pub(crate) const MAX_RECURSION: u32 = 64;
pub(crate) const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;

/// Finite aggregate policy for one storage-root traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    pub(crate) max_message_bytes: usize,
    pub(crate) max_fields: usize,
    pub(crate) max_work_bytes: usize,
    pub(crate) recursion_limit: u32,
    pub(crate) max_references: usize,
    pub(crate) max_text_bytes: usize,
}

impl DecodeOptions {
    /// Construct an explicit bytes/fields/work/nesting/reference/text policy.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
            max_text_bytes,
        }
    }

    pub(crate) fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact successful aggregate consumption for transaction-budget merging.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
    text_bytes: usize,
}

impl DecodeReport {
    /// Bytes in the caller-owned root payload.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Encoded fields inspected by all strict owner/reference traversals.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Exact bytes inspected by handwritten and Buffa passes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest protobuf message or unknown-group depth reached.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Selected `TSP.Reference` occurrences.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Aggregate bytes inside selected reference envelopes.
    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }

    /// Aggregate UTF-8 bytes in selected string fields.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Root payload or configured Buffa ceiling is too large.
    Bytes { observed: usize, maximum: usize },
    /// Selected reference occurrences exceed their aggregate ceiling.
    References { observed: usize, maximum: usize },
    /// Selected UTF-8 bytes exceed their aggregate ceiling.
    Text { observed: usize, maximum: usize },
    /// Strictly inspected fields exceed their aggregate ceiling.
    Fields { observed: usize, maximum: usize },
    /// Handwritten plus Buffa work exceeds its aggregate ceiling.
    Work { observed: usize, maximum: usize },
    /// Configured or traversed nesting exceeds its finite ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// A fallible transaction-staging allocation was refused.
    Allocation { requested: usize },
}

/// Strict table-cell storage decode failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    /// Return the exact finite resource observation, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }

    /// Requested element count for a refused staging allocation.
    #[must_use]
    pub const fn allocation_requested(&self) -> Option<usize> {
        match self.limit {
            Some(DecodeLimit::Allocation { requested }) => Some(requested),
            _ => None,
        }
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
        formatter.write_str("invalid Numbers table-cell storage payload")
    }
}

impl std::error::Error for DecodeError {}

/// Generated-free scalar projection of one canonical `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ReferenceSnapshot {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// One source-ordered tile-storage record.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct TileReferenceRecord<'source> {
    raw: &'source [u8],
    tile_id: u32,
    reference: ReferenceSnapshot,
}

impl<'source> TileReferenceRecord<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
    #[must_use]
    pub const fn tile_id(self) -> u32 {
        self.tile_id
    }
    #[must_use]
    pub const fn reference(self) -> ReferenceSnapshot {
        self.reference
    }
}

/// A source-ordered selected reference record.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ReferenceRecord<'source> {
    pub(crate) raw: &'source [u8],
    pub(crate) reference: ReferenceSnapshot,
}

/// One source-ordered, fully validated table header record.
///
/// The raw payload remains caller-owned and authoritative. `snapshot` is the
/// generated-free strict/Buffa parity result for those exact bytes.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct HeaderRecord<'source> {
    raw: &'source [u8],
    snapshot: HeaderSnapshot,
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

impl<'source> ReferenceRecord<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
    #[must_use]
    pub const fn reference(self) -> ReferenceSnapshot {
        self.reference
    }
}

/// Streaming hooks for collection fields. Default methods retain nothing.
///
/// Each callback observes a fully validated record, but it can run before the
/// enclosing owner finishes strict validation and Buffa parity. A later error
/// does not roll callbacks back. Callers must therefore stage side effects and
/// publish them only after the decode function returns `Ok`, or provide their
/// own reversible rollback discipline.
pub trait StorageVisitor {
    fn visit_tile_reference(
        &mut self,
        _record: TileReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_tile_row(&mut self, _row: TileRowInfoSnapshot<'_>) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_header_bucket(&mut self, _reference: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_header(&mut self, _header: HeaderSnapshot) -> Result<(), DecodeError> {
        Ok(())
    }
    /// Visit a validated header together with its exact source payload.
    ///
    /// The default forwards to `visit_header`, preserving existing visitor
    /// implementations. New consumers that need wire-exact mutation should
    /// override this method.
    fn visit_header_record(&mut self, record: HeaderRecord<'_>) -> Result<(), DecodeError> {
        self.visit_header(record.snapshot())
    }
    fn visit_list_entry(
        &mut self,
        _entry: TableDataListEntrySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_list_segment(&mut self, _reference: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl StorageVisitor for () {}

/// Borrowed model root required by a scalar-cell proof.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct TableModelSnapshot<'source> {
    table_id: &'source str,
    base_data_store: &'source [u8],
    number_of_rows: u32,
    number_of_columns: u32,
    hidden_state_formula_owner_for_columns: Option<ReferenceSnapshot>,
    hidden_state_formula_owner_for_rows: Option<ReferenceSnapshot>,
    conditional_style_formula_owner_id: Option<&'source [u8]>,
    pivot_owner: Option<ReferenceSnapshot>,
    category_owner: Option<ReferenceSnapshot>,
    spill_owner: Option<&'source [u8]>,
}

impl<'source> TableModelSnapshot<'source> {
    #[must_use]
    pub const fn table_id(self) -> &'source str {
        self.table_id
    }
    #[must_use]
    pub const fn base_data_store(self) -> &'source [u8] {
        self.base_data_store
    }
    #[must_use]
    pub const fn number_of_rows(self) -> u32 {
        self.number_of_rows
    }
    #[must_use]
    pub const fn number_of_columns(self) -> u32 {
        self.number_of_columns
    }
    #[must_use]
    pub const fn hidden_state_formula_owner_for_columns(self) -> Option<ReferenceSnapshot> {
        self.hidden_state_formula_owner_for_columns
    }
    #[must_use]
    pub const fn hidden_state_formula_owner_for_rows(self) -> Option<ReferenceSnapshot> {
        self.hidden_state_formula_owner_for_rows
    }
    #[must_use]
    pub const fn conditional_style_formula_owner_id(self) -> Option<&'source [u8]> {
        self.conditional_style_formula_owner_id
    }
    #[must_use]
    pub const fn pivot_owner(self) -> Option<ReferenceSnapshot> {
        self.pivot_owner
    }
    #[must_use]
    pub const fn category_owner(self) -> Option<ReferenceSnapshot> {
        self.category_owner
    }
    #[must_use]
    pub const fn spill_owner(self) -> Option<&'source [u8]> {
        self.spill_owner
    }
}

/// Borrowed base-data-store routes and scalar counters.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct DataStoreSnapshot<'source> {
    row_headers: &'source [u8],
    column_headers: ReferenceSnapshot,
    tiles: &'source [u8],
    string_table: ReferenceSnapshot,
    style_table: ReferenceSnapshot,
    formula_table: ReferenceSnapshot,
    next_row_strip_id: u32,
    next_column_strip_id: u32,
    row_tile_tree: &'source [u8],
    column_tile_tree: &'source [u8],
    format_table_pre_bnc: ReferenceSnapshot,
    formula_error_table: Option<ReferenceSnapshot>,
    merge_region_map: Option<ReferenceSnapshot>,
    storage_version_pre_bnc: Option<u32>,
    deprecated_custom_format_table: Option<ReferenceSnapshot>,
    multiple_choice_list_format_table: Option<ReferenceSnapshot>,
    rich_text_table: Option<ReferenceSnapshot>,
    conditional_style_table: Option<ReferenceSnapshot>,
    comment_storage_table: Option<ReferenceSnapshot>,
    import_warning_set_table: Option<ReferenceSnapshot>,
    control_cell_spec_table: Option<ReferenceSnapshot>,
    format_table: Option<ReferenceSnapshot>,
}

macro_rules! datastore_accessors {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {$(
        #[must_use]
        pub const fn $name(self) -> $ty { self.$name }
    )+};
}

impl<'source> DataStoreSnapshot<'source> {
    datastore_accessors!(
        (row_headers, &'source [u8]),
        (column_headers, ReferenceSnapshot),
        (tiles, &'source [u8]),
        (string_table, ReferenceSnapshot),
        (style_table, ReferenceSnapshot),
        (formula_table, ReferenceSnapshot),
        (next_row_strip_id, u32),
        (next_column_strip_id, u32),
        (row_tile_tree, &'source [u8]),
        (column_tile_tree, &'source [u8]),
        (format_table_pre_bnc, ReferenceSnapshot),
        (formula_error_table, Option<ReferenceSnapshot>),
        (merge_region_map, Option<ReferenceSnapshot>),
        (storage_version_pre_bnc, Option<u32>),
        (deprecated_custom_format_table, Option<ReferenceSnapshot>),
        (multiple_choice_list_format_table, Option<ReferenceSnapshot>),
        (rich_text_table, Option<ReferenceSnapshot>),
        (conditional_style_table, Option<ReferenceSnapshot>),
        (comment_storage_table, Option<ReferenceSnapshot>),
        (import_warning_set_table, Option<ReferenceSnapshot>),
        (control_cell_spec_table, Option<ReferenceSnapshot>),
        (format_table, Option<ReferenceSnapshot>)
    );
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TileStorageSnapshot {
    tile_size: Option<u32>,
    should_use_wide_rows: Option<bool>,
}
impl TileStorageSnapshot {
    #[must_use]
    pub const fn tile_size(self) -> Option<u32> {
        self.tile_size
    }
    #[must_use]
    pub const fn should_use_wide_rows(self) -> Option<bool> {
        self.should_use_wide_rows
    }
}

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderStorageSnapshot {
    bucket_hash_function: u32,
}
impl HeaderStorageSnapshot {
    #[must_use]
    pub const fn bucket_hash_function(self) -> u32 {
        self.bucket_hash_function
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderStorageBucketSnapshot {
    bucket_hash_function: u32,
}
impl HeaderStorageBucketSnapshot {
    #[must_use]
    pub const fn bucket_hash_function(self) -> u32 {
        self.bucket_hash_function
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderSnapshot {
    index: u32,
    size_bits: u32,
    hiding_state: u32,
    number_of_cells: u32,
    cell_style: Option<ReferenceSnapshot>,
    text_style: Option<ReferenceSnapshot>,
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
    #[must_use]
    pub const fn cell_style(self) -> Option<ReferenceSnapshot> {
        self.cell_style
    }
    #[must_use]
    pub const fn text_style(self) -> Option<ReferenceSnapshot> {
        self.text_style
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableDataListSnapshot {
    list_type: i32,
    next_list_id: u32,
    is_new_for_bnc: Option<bool>,
}
impl TableDataListSnapshot {
    #[must_use]
    pub const fn list_type(self) -> i32 {
        self.list_type
    }
    #[must_use]
    pub const fn next_list_id(self) -> u32 {
        self.next_list_id
    }
    #[must_use]
    pub const fn is_new_for_bnc(self) -> Option<bool> {
        self.is_new_for_bnc
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct TableDataListEntrySnapshot<'source> {
    key: u32,
    ref_count: u32,
    string_value: Option<&'source str>,
    reference: Option<ReferenceSnapshot>,
    formula: Option<&'source [u8]>,
    format: Option<&'source [u8]>,
    custom_format: Option<&'source [u8]>,
    rich_text_payload: Option<ReferenceSnapshot>,
    comment_storage: Option<ReferenceSnapshot>,
    import_warning_set: Option<&'source [u8]>,
    cell_spec: Option<&'source [u8]>,
}
impl<'source> TableDataListEntrySnapshot<'source> {
    #[must_use]
    pub const fn key(self) -> u32 {
        self.key
    }
    #[must_use]
    pub const fn ref_count(self) -> u32 {
        self.ref_count
    }
    #[must_use]
    pub const fn string_value(self) -> Option<&'source str> {
        self.string_value
    }
    #[must_use]
    pub const fn reference(self) -> Option<ReferenceSnapshot> {
        self.reference
    }
    #[must_use]
    pub const fn formula(self) -> Option<&'source [u8]> {
        self.formula
    }
    #[must_use]
    pub const fn format(self) -> Option<&'source [u8]> {
        self.format
    }
    #[must_use]
    pub const fn custom_format(self) -> Option<&'source [u8]> {
        self.custom_format
    }
    #[must_use]
    pub const fn rich_text_payload(self) -> Option<ReferenceSnapshot> {
        self.rich_text_payload
    }
    #[must_use]
    pub const fn comment_storage(self) -> Option<ReferenceSnapshot> {
        self.comment_storage
    }
    #[must_use]
    pub const fn import_warning_set(self) -> Option<&'source [u8]> {
        self.import_warning_set
    }
    #[must_use]
    pub const fn cell_spec(self) -> Option<&'source [u8]> {
        self.cell_spec
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct TableDataListSegmentSnapshot<'source> {
    list_type: i32,
    key_range: &'source [u8],
    key_range_location: u32,
    key_range_length: u32,
}
impl<'source> TableDataListSegmentSnapshot<'source> {
    #[must_use]
    pub const fn list_type(self) -> i32 {
        self.list_type
    }
    #[must_use]
    pub const fn key_range(self) -> &'source [u8] {
        self.key_range
    }
    #[must_use]
    pub const fn key_range_location(self) -> u32 {
        self.key_range_location
    }
    #[must_use]
    pub const fn key_range_length(self) -> u32 {
        self.key_range_length
    }
}

macro_rules! impl_redacted_debug {
    ($($snapshot:ident),+ $(,)?) => {$(
        impl fmt::Debug for $snapshot<'_> {
            fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str(concat!(stringify!($snapshot), " { payloads: <redacted> }"))
            }
        }
    )+};
}

impl_redacted_debug!(
    TileReferenceRecord,
    ReferenceRecord,
    HeaderRecord,
    TableModelSnapshot,
    DataStoreSnapshot,
    TileRowInfoSnapshot,
    TableDataListEntrySnapshot,
    TableDataListSegmentSnapshot,
);

/// One source-indexed size mutation for a header-storage bucket.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderSizeEdit {
    index: u32,
    replacement_size_bits: Option<u32>,
}

impl HeaderSizeEdit {
    /// Set the exact IEEE-754 bits, inserting a canonical minimal header when
    /// the source bucket has no record for `index`.
    #[must_use]
    pub const fn set(index: u32, replacement_size_bits: u32) -> Self {
        Self {
            index,
            replacement_size_bits: Some(replacement_size_bits),
        }
    }

    /// Clear a size override. Exact canonical minimal headers are removed;
    /// non-minimal records retain all facets/unknowns and are patched to +0.
    #[must_use]
    pub const fn remove(index: u32) -> Self {
        Self {
            index,
            replacement_size_bits: None,
        }
    }

    #[must_use]
    pub const fn index(self) -> u32 {
        self.index
    }

    #[must_use]
    pub const fn replacement_size_bits(self) -> Option<u32> {
        self.replacement_size_bits
    }
}

/// Exact preflight and readback accounting for a header-size rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderSizeRewriteReport {
    source: DecodeReport,
    result: DecodeReport,
    inserted: usize,
    removed: usize,
    updated: usize,
    header_count: usize,
    edit_count: usize,
    output_bytes: usize,
    rewrite_work_bytes: usize,
}

impl HeaderSizeRewriteReport {
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }

    #[must_use]
    pub const fn result(self) -> DecodeReport {
        self.result
    }

    #[must_use]
    pub const fn inserted(self) -> usize {
        self.inserted
    }

    #[must_use]
    pub const fn removed(self) -> usize {
        self.removed
    }

    #[must_use]
    pub const fn updated(self) -> usize {
        self.updated
    }

    #[must_use]
    pub const fn header_count(self) -> usize {
        self.header_count
    }

    #[must_use]
    pub const fn edit_count(self) -> usize {
        self.edit_count
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn rewrite_work_bytes(self) -> usize {
        self.rewrite_work_bytes
    }
}

/// Conservative result-decode resources known before output allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeResourceUpperBound {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
    text_bytes: usize,
}

impl DecodeResourceUpperBound {
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
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

/// All ledger requirements computed before an output buffer is reserved.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderSizeRewriteRequirements {
    source: DecodeReport,
    result_upper_bound: DecodeResourceUpperBound,
    inserted: usize,
    removed: usize,
    updated: usize,
    header_count: usize,
    edit_count: usize,
    output_bytes: usize,
    rewrite_work_bytes: usize,
}

impl HeaderSizeRewriteRequirements {
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }
    #[must_use]
    pub const fn result_upper_bound(self) -> DecodeResourceUpperBound {
        self.result_upper_bound
    }
    #[must_use]
    pub const fn inserted(self) -> usize {
        self.inserted
    }
    #[must_use]
    pub const fn removed(self) -> usize {
        self.removed
    }
    #[must_use]
    pub const fn updated(self) -> usize {
        self.updated
    }
    #[must_use]
    pub const fn header_count(self) -> usize {
        self.header_count
    }
    #[must_use]
    pub const fn edit_count(self) -> usize {
        self.edit_count
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn rewrite_work_bytes(self) -> usize {
        self.rewrite_work_bytes
    }
}

/// Opaque validated rewrite plan borrowing the caller-authoritative source.
pub struct HeaderSizeRewritePlan<'source> {
    source: &'source [u8],
    records: Vec<StagedHeaderRecord>,
    header_indices: Vec<u32>,
    edits: Vec<HeaderSizeEdit>,
    requirements: HeaderSizeRewriteRequirements,
}

impl fmt::Debug for HeaderSizeRewritePlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HeaderSizeRewritePlan")
            .field("requirements", &self.requirements)
            .field("payloads", &"<redacted>")
            .finish()
    }
}

impl HeaderSizeRewritePlan<'_> {
    #[must_use]
    pub const fn requirements(&self) -> HeaderSizeRewriteRequirements {
        self.requirements
    }
}

/// Decode one table model without retaining collection-width state.
pub fn decode_table_model(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelSnapshot<'_>, DecodeError> {
    Ok(decode_table_model_with_report(source, options)?.0)
}

/// Decode one table model and return exact aggregate resource consumption.
pub fn decode_table_model_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelSnapshot<'_>, DecodeReport), DecodeError> {
    decode_table_model_with_visitor(source, options, &mut ())
}

/// Decode a model and stream every selected repeated storage record.
pub fn decode_table_model_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TableModelSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_model_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_table_model_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<TableModelSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut table_id = None;
    let mut base_data_store = None;
    let mut number_of_rows = None;
    let mut number_of_columns = None;
    let mut hidden_columns = None;
    let mut hidden_rows = None;
    let mut conditional_owner = None;
    let mut pivot_owner = None;
    let mut category_owner = None;
    let mut spill_owner = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => {
                let raw = field.bytes()?;
                set_once(&mut table_id, strict_utf8(raw, budget)?)?;
            },
            4 => {
                let raw = field.bytes()?;
                if base_data_store.is_some() {
                    return Err(DecodeError::invalid());
                }
                let _ = decode_data_store_in(raw, budget, child_depth, visitor)?;
                base_data_store = Some(raw);
            },
            6 => set_once(&mut number_of_rows, canonical_u32(field.varint()?)?)?,
            7 => set_once(&mut number_of_columns, canonical_u32(field.varint()?)?)?,
            34 => {
                let raw = field.bytes()?;
                let reference = decode_reference(raw, budget, child_depth)?;
                set_once(&mut hidden_columns, (raw, reference))?;
            },
            35 => {
                let raw = field.bytes()?;
                let reference = decode_reference(raw, budget, child_depth)?;
                set_once(&mut hidden_rows, (raw, reference))?;
            },
            39 => {
                let raw = field.bytes()?;
                if conditional_owner.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                conditional_owner = Some(raw);
            },
            85 => {
                let raw = field.bytes()?;
                let reference = decode_reference(raw, budget, child_depth)?;
                set_once(&mut pivot_owner, (raw, reference))?;
            },
            86 => {
                let raw = field.bytes()?;
                let reference = decode_reference(raw, budget, child_depth)?;
                set_once(&mut category_owner, (raw, reference))?;
            },
            93 => {
                let raw = field.bytes()?;
                if spill_owner.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                spill_owner = Some(raw);
            },
            _ => {},
        }
    }
    let snapshot = TableModelSnapshot {
        table_id: table_id.ok_or_else(DecodeError::invalid)?,
        base_data_store: base_data_store.ok_or_else(DecodeError::invalid)?,
        number_of_rows: number_of_rows.ok_or_else(DecodeError::invalid)?,
        number_of_columns: number_of_columns.ok_or_else(DecodeError::invalid)?,
        hidden_state_formula_owner_for_columns: hidden_columns.map(|(_raw, reference)| reference),
        hidden_state_formula_owner_for_rows: hidden_rows.map(|(_raw, reference)| reference),
        conditional_style_formula_owner_id: conditional_owner,
        pivot_owner: pivot_owner.map(|(_raw, reference)| reference),
        category_owner: category_owner.map(|(_raw, reference)| reference),
        spill_owner,
    };
    budget.message(source, depth)?;
    let view: projection::TableModelArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.table_id != snapshot.table_id
        || view.base_data_store != snapshot.base_data_store
        || view.number_of_rows != snapshot.number_of_rows
        || view.number_of_columns != snapshot.number_of_columns
        || view.hidden_state_formula_owner_for_columns
            != hidden_columns.map(|(raw, _reference)| raw)
        || view.hidden_state_formula_owner_for_rows != hidden_rows.map(|(raw, _reference)| raw)
        || view.conditional_style_formula_owner_id != snapshot.conditional_style_formula_owner_id
        || view.pivot_owner != pivot_owner.map(|(raw, _reference)| raw)
        || view.category_owner != category_owner.map(|(raw, _reference)| raw)
        || view.spill_owner != snapshot.spill_owner
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

/// Decode one native data-store envelope.
pub fn decode_data_store(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DataStoreSnapshot<'_>, DecodeError> {
    Ok(decode_data_store_with_report(source, options)?.0)
}

pub fn decode_data_store_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DataStoreSnapshot<'_>, DecodeReport), DecodeError> {
    decode_data_store_with_visitor(source, options, &mut ())
}

pub fn decode_data_store_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(DataStoreSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_data_store_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_data_store_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<DataStoreSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut raw_fields: [Option<&'source [u8]>; 22] = [None; 22];
    let mut refs: [Option<ReferenceSnapshot>; 22] = [None; 22];
    let mut next_row_strip_id = None;
    let mut next_column_strip_id = None;
    let mut storage_version_pre_bnc = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        let number = usize::try_from(field.number).map_err(|_conversion| DecodeError::invalid())?;
        match field.number {
            1 => {
                let raw = field.bytes()?;
                if raw_fields[0].is_some() {
                    return Err(DecodeError::invalid());
                }
                let _ = decode_header_storage_in(raw, budget, child_depth, visitor)?;
                raw_fields[0] = Some(raw);
            },
            3 => {
                let raw = field.bytes()?;
                if raw_fields[2].is_some() {
                    return Err(DecodeError::invalid());
                }
                let _ = decode_tile_storage_in(raw, budget, child_depth, visitor)?;
                raw_fields[2] = Some(raw);
            },
            2 | 4 | 5 | 6 | 11 | 12 | 13 | 15..=22 => {
                let raw = field.bytes()?;
                if raw_fields[number - 1].is_some() {
                    return Err(DecodeError::invalid());
                }
                refs[number - 1] = Some(decode_reference(raw, budget, child_depth)?);
                raw_fields[number - 1] = Some(raw);
            },
            7 => set_once(&mut next_row_strip_id, canonical_u32(field.varint()?)?)?,
            8 => set_once(&mut next_column_strip_id, canonical_u32(field.varint()?)?)?,
            9 | 10 => {
                let raw = field.bytes()?;
                if raw_fields[number - 1].is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                raw_fields[number - 1] = Some(raw);
            },
            14 => set_once(
                &mut storage_version_pre_bnc,
                canonical_u32(field.varint()?)?,
            )?,
            _ => {},
        }
    }
    let snapshot = DataStoreSnapshot {
        row_headers: raw_fields[0].ok_or_else(DecodeError::invalid)?,
        column_headers: refs[1].ok_or_else(DecodeError::invalid)?,
        tiles: raw_fields[2].ok_or_else(DecodeError::invalid)?,
        string_table: refs[3].ok_or_else(DecodeError::invalid)?,
        style_table: refs[4].ok_or_else(DecodeError::invalid)?,
        formula_table: refs[5].ok_or_else(DecodeError::invalid)?,
        next_row_strip_id: next_row_strip_id.ok_or_else(DecodeError::invalid)?,
        next_column_strip_id: next_column_strip_id.ok_or_else(DecodeError::invalid)?,
        row_tile_tree: raw_fields[8].ok_or_else(DecodeError::invalid)?,
        column_tile_tree: raw_fields[9].ok_or_else(DecodeError::invalid)?,
        format_table_pre_bnc: refs[10].ok_or_else(DecodeError::invalid)?,
        formula_error_table: refs[11],
        merge_region_map: refs[12],
        storage_version_pre_bnc,
        deprecated_custom_format_table: refs[14],
        multiple_choice_list_format_table: refs[15],
        rich_text_table: refs[16],
        conditional_style_table: refs[17],
        comment_storage_table: refs[18],
        import_warning_set_table: refs[19],
        control_cell_spec_table: refs[20],
        format_table: refs[21],
    };
    budget.message(source, depth)?;
    let view: projection::DataStoreArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    let raw = |field: usize| raw_fields[field - 1];
    if view.row_headers != snapshot.row_headers
        || Some(view.column_headers) != raw(2)
        || view.tiles != snapshot.tiles
        || Some(view.string_table) != raw(4)
        || Some(view.style_table) != raw(5)
        || Some(view.formula_table) != raw(6)
        || view.next_row_strip_id != snapshot.next_row_strip_id
        || view.next_column_strip_id != snapshot.next_column_strip_id
        || view.row_tile_tree != snapshot.row_tile_tree
        || view.column_tile_tree != snapshot.column_tile_tree
        || Some(view.format_table_pre_bnc) != raw(11)
        || view.formula_error_table != raw(12)
        || view.merge_region_map != raw(13)
        || view.storage_version_pre_bnc != snapshot.storage_version_pre_bnc
        || view.deprecated_custom_format_table != raw(15)
        || view.multiple_choice_list_format_table != raw(16)
        || view.rich_text_table != raw(17)
        || view.conditional_style_table != raw(18)
        || view.comment_storage_table != raw(19)
        || view.import_warning_set_table != raw(20)
        || view.control_cell_spec_table != raw(21)
        || view.format_table != raw(22)
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_tile_storage(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TileStorageSnapshot, DecodeError> {
    Ok(decode_tile_storage_with_report(source, options)?.0)
}

pub fn decode_tile_storage_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TileStorageSnapshot, DecodeReport), DecodeError> {
    decode_tile_storage_with_visitor(source, options, &mut ())
}

pub fn decode_tile_storage_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TileStorageSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_tile_storage_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_tile_storage_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<TileStorageSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut tile_size = None;
    let mut should_use_wide_rows = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => {
                let raw = field.bytes()?;
                budget.message(raw, child_depth)?;
                let reference_depth = child_depth
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                let mut tile_id = None;
                let mut reference = None;
                let mut record = raw;
                while let Some(record_field) = next_field(&mut record, budget, child_depth)? {
                    match record_field.number {
                        1 => set_once(&mut tile_id, canonical_u32(record_field.varint()?)?)?,
                        2 => {
                            let payload = record_field.bytes()?;
                            set_once(
                                &mut reference,
                                decode_reference(payload, budget, reference_depth)?,
                            )?;
                        },
                        _ => {},
                    }
                }
                visitor.visit_tile_reference(TileReferenceRecord {
                    raw,
                    tile_id: tile_id.ok_or_else(DecodeError::invalid)?,
                    reference: reference.ok_or_else(DecodeError::invalid)?,
                })?;
            },
            2 => set_once(&mut tile_size, canonical_u32(field.varint()?)?)?,
            3 => set_once(&mut should_use_wide_rows, canonical_bool(field.varint()?)?)?,
            _ => {},
        }
    }
    let snapshot = TileStorageSnapshot {
        tile_size,
        should_use_wide_rows,
    };
    budget.message(source, depth)?;
    let view: projection::TileStorageArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.tile_size != snapshot.tile_size
        || view.should_use_wide_rows != snapshot.should_use_wide_rows
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_tile(source: &[u8], options: DecodeOptions) -> Result<TileSnapshot, DecodeError> {
    Ok(decode_tile_with_report(source, options)?.0)
}

pub fn decode_tile_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TileSnapshot, DecodeReport), DecodeError> {
    decode_tile_with_visitor(source, options, &mut ())
}

pub fn decode_tile_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TileSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_tile_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_tile_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<TileSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut max_column = None;
    let mut max_row = None;
    let mut num_cells = None;
    let mut num_rows = None;
    let mut storage_version = None;
    let mut last_saved_in_bnc = None;
    let mut should_use_wide_rows = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut max_column, canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut max_row, canonical_u32(field.varint()?)?)?,
            3 => set_once(&mut num_cells, canonical_u32(field.varint()?)?)?,
            4 => set_once(&mut num_rows, canonical_u32(field.varint()?)?)?,
            5 => {
                let row = decode_tile_row_info_in(field.bytes()?, budget, child_depth)?;
                visitor.visit_tile_row(row)?;
            },
            6 => set_once(&mut storage_version, canonical_u32(field.varint()?)?)?,
            7 => set_once(&mut last_saved_in_bnc, canonical_bool(field.varint()?)?)?,
            8 => set_once(&mut should_use_wide_rows, canonical_bool(field.varint()?)?)?,
            _ => {},
        }
    }
    let snapshot = TileSnapshot {
        max_column: max_column.ok_or_else(DecodeError::invalid)?,
        max_row: max_row.ok_or_else(DecodeError::invalid)?,
        num_cells: num_cells.ok_or_else(DecodeError::invalid)?,
        num_rows: num_rows.ok_or_else(DecodeError::invalid)?,
        storage_version,
        last_saved_in_bnc,
        should_use_wide_rows,
    };
    budget.message(source, depth)?;
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
    Ok(snapshot)
}

pub fn decode_tile_row_info(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TileRowInfoSnapshot<'_>, DecodeError> {
    Ok(decode_tile_row_info_with_report(source, options)?.0)
}

pub fn decode_tile_row_info_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TileRowInfoSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_tile_row_info_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_tile_row_info_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<TileRowInfoSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let mut tile_row_index = None;
    let mut cell_count = None;
    let mut pre_bnc_buffer = None;
    let mut pre_bnc_offsets = None;
    let mut storage_version = None;
    let mut buffer = None;
    let mut offsets = None;
    let mut has_wide_offsets = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut tile_row_index, canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut cell_count, canonical_u32(field.varint()?)?)?,
            3 => set_once(&mut pre_bnc_buffer, field.bytes()?)?,
            4 => set_once(&mut pre_bnc_offsets, field.bytes()?)?,
            5 => set_once(&mut storage_version, canonical_u32(field.varint()?)?)?,
            6 => set_once(&mut buffer, field.bytes()?)?,
            7 => set_once(&mut offsets, field.bytes()?)?,
            8 => set_once(&mut has_wide_offsets, canonical_bool(field.varint()?)?)?,
            _ => {},
        }
    }
    let snapshot = TileRowInfoSnapshot {
        tile_row_index: tile_row_index.ok_or_else(DecodeError::invalid)?,
        cell_count: cell_count.ok_or_else(DecodeError::invalid)?,
        cell_storage_buffer_pre_bnc: pre_bnc_buffer.ok_or_else(DecodeError::invalid)?,
        cell_offsets_pre_bnc: pre_bnc_offsets.ok_or_else(DecodeError::invalid)?,
        storage_version,
        cell_storage_buffer: buffer,
        cell_offsets: offsets,
        has_wide_offsets,
    };
    budget.message(source, depth)?;
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
    Ok(snapshot)
}

pub fn decode_header_storage(
    source: &[u8],
    options: DecodeOptions,
) -> Result<HeaderStorageSnapshot, DecodeError> {
    Ok(decode_header_storage_with_report(source, options)?.0)
}

pub fn decode_header_storage_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(HeaderStorageSnapshot, DecodeReport), DecodeError> {
    decode_header_storage_with_visitor(source, options, &mut ())
}

pub fn decode_header_storage_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(HeaderStorageSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_header_storage_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_header_storage_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<HeaderStorageSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut bucket_hash_function = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut bucket_hash_function, canonical_u32(field.varint()?)?)?,
            2 => {
                let raw = field.bytes()?;
                let reference = decode_reference(raw, budget, child_depth)?;
                visitor.visit_header_bucket(ReferenceRecord { raw, reference })?;
            },
            _ => {},
        }
    }
    let snapshot = HeaderStorageSnapshot {
        bucket_hash_function: bucket_hash_function.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::HeaderStorageArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.bucket_hash_function != snapshot.bucket_hash_function {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_header_storage_bucket(
    source: &[u8],
    options: DecodeOptions,
) -> Result<HeaderStorageBucketSnapshot, DecodeError> {
    Ok(decode_header_storage_bucket_with_report(source, options)?.0)
}

pub fn decode_header_storage_bucket_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(HeaderStorageBucketSnapshot, DecodeReport), DecodeError> {
    decode_header_storage_bucket_with_visitor(source, options, &mut ())
}

pub fn decode_header_storage_bucket_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(HeaderStorageBucketSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_header_storage_bucket_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

/// Rewrite a batch of header-size records after strict handwritten/Buffa
/// preflight, preserving every untouched source byte and existing record order.
pub fn rewrite_header_storage_bucket_sizes(
    source: &[u8],
    dimension_limit: u32,
    edits: &[HeaderSizeEdit],
    options: DecodeOptions,
) -> Result<(Vec<u8>, HeaderSizeRewriteReport), DecodeError> {
    let plan = plan_header_storage_bucket_sizes(source, dimension_limit, edits, options)?;
    execute_header_storage_bucket_size_plan(plan, options)
}

/// Validate and budget a rewrite without allocating its output buffer.
pub fn plan_header_storage_bucket_sizes<'source>(
    source: &'source [u8],
    dimension_limit: u32,
    edits: &[HeaderSizeEdit],
    options: DecodeOptions,
) -> Result<HeaderSizeRewritePlan<'source>, DecodeError> {
    if dimension_limit == 0 {
        return Err(DecodeError::invalid());
    }

    let mut records = HeaderRecordStage::new(source);
    let (_bucket, source_report) =
        decode_header_storage_bucket_with_visitor(source, options, &mut records)?;
    let header_indices = records.validate(dimension_limit)?;
    let mut ordered_edits = fallible_copy(edits)?;
    ordered_edits.sort_unstable_by_key(|edit| edit.index);
    validate_sorted_edits(&ordered_edits, dimension_limit)?;
    let mut output_len = source.len();
    let mut inserted = 0usize;
    let mut removed = 0usize;
    let mut updated = 0usize;
    let mut removed_payload_bytes = 0usize;
    let mut inserted_payload_bytes = 0usize;
    for record in &records.records {
        let Some(edit) = edit_for(&ordered_edits, record.snapshot.index()) else {
            continue;
        };
        match edit.replacement_size_bits {
            Some(bits) if bits != record.snapshot.size_bits() => {
                updated = updated.checked_add(1).ok_or_else(DecodeError::invalid)?;
            },
            None if is_canonical_minimal(source, record) => {
                output_len = output_len
                    .checked_sub(record.end - record.start)
                    .ok_or_else(DecodeError::invalid)?;
                removed = removed.checked_add(1).ok_or_else(DecodeError::invalid)?;
                removed_payload_bytes = removed_payload_bytes
                    .checked_add(record.payload_end - record.payload_start)
                    .ok_or_else(DecodeError::invalid)?;
            },
            None if record.snapshot.size_bits() != 0 => {
                updated = updated.checked_add(1).ok_or_else(DecodeError::invalid)?;
            },
            _ => {},
        }
    }
    for edit in &ordered_edits {
        if let Some(bits) = edit.replacement_size_bits
            && header_indices.binary_search(&edit.index).is_err()
        {
            let (_, payload_len) = canonical_minimal_header(edit.index, bits);
            output_len = output_len
                .checked_add(encoded_header_field_length(payload_len)?)
                .ok_or_else(DecodeError::invalid)?;
            inserted = inserted.checked_add(1).ok_or_else(DecodeError::invalid)?;
            inserted_payload_bytes = inserted_payload_bytes
                .checked_add(payload_len)
                .ok_or_else(DecodeError::invalid)?;
        }
    }
    if output_len > options.max_message_bytes
        || output_len
            > usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError::invalid())?
    {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: output_len,
            maximum: options.max_message_bytes,
        }));
    }
    let result_upper_bound = result_decode_upper_bound(
        source_report,
        source.len(),
        output_len,
        inserted,
        removed,
        inserted_payload_bytes,
        removed_payload_bytes,
    )?;
    validate_result_ceiling(result_upper_bound, options)?;
    let requirements = HeaderSizeRewriteRequirements {
        source: source_report,
        result_upper_bound,
        inserted,
        removed,
        updated,
        header_count: records.records.len(),
        edit_count: ordered_edits.len(),
        output_bytes: output_len,
        rewrite_work_bytes: rewrite_work_upper_bound(
            source.len(),
            output_len,
            records.records.len(),
            ordered_edits.len(),
        )?,
    };
    Ok(HeaderSizeRewritePlan {
        source,
        records: records.records,
        header_indices,
        edits: ordered_edits,
        requirements,
    })
}

/// Execute a validated plan after refusing insufficient result ceilings and
/// before reserving the exact output allocation.
pub fn execute_header_storage_bucket_size_plan(
    plan: HeaderSizeRewritePlan<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, HeaderSizeRewriteReport), DecodeError> {
    validate_result_ceiling(plan.requirements.result_upper_bound, options)?;
    let result = assemble_rewritten_bucket(
        plan.source,
        &plan.records,
        &plan.header_indices,
        &plan.edits,
        plan.requirements.output_bytes,
    )?;
    let (_bucket, result_report) = decode_header_storage_bucket_with_report(&result, options)?;
    let requirements = plan.requirements;
    Ok((
        result,
        HeaderSizeRewriteReport {
            source: requirements.source,
            result: result_report,
            inserted: requirements.inserted,
            removed: requirements.removed,
            updated: requirements.updated,
            header_count: requirements.header_count,
            edit_count: requirements.edit_count,
            output_bytes: requirements.output_bytes,
            rewrite_work_bytes: requirements.rewrite_work_bytes,
        },
    ))
}

struct StagedHeaderRecord {
    start: usize,
    end: usize,
    payload_start: usize,
    payload_end: usize,
    snapshot: HeaderSnapshot,
}

struct HeaderRecordStage {
    source_start: *const u8,
    source_len: usize,
    records: Vec<StagedHeaderRecord>,
}

impl HeaderRecordStage {
    fn new(source: &[u8]) -> Self {
        Self {
            source_start: source.as_ptr(),
            source_len: source.len(),
            records: Vec::new(),
        }
    }
}

impl HeaderRecordStage {
    fn validate(&self, dimension_limit: u32) -> Result<Vec<u32>, DecodeError> {
        let mut indices = Vec::new();
        indices.try_reserve_exact(self.records.len()).map_err(|_| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: self.records.len(),
            })
        })?;
        indices.extend(self.records.iter().map(|record| record.snapshot.index()));
        indices.sort_unstable();
        if indices
            .last()
            .is_some_and(|index| *index >= dimension_limit)
            || indices.windows(2).any(|pair| pair[0] == pair[1])
        {
            return Err(DecodeError::invalid());
        }
        Ok(indices)
    }
}

impl StorageVisitor for HeaderRecordStage {
    fn visit_header_record(&mut self, record: HeaderRecord<'_>) -> Result<(), DecodeError> {
        let payload_start = (record.raw().as_ptr() as usize)
            .checked_sub(self.source_start as usize)
            .ok_or_else(DecodeError::invalid)?;
        let payload_end = payload_start
            .checked_add(record.raw().len())
            .filter(|end| *end <= self.source_len)
            .ok_or_else(DecodeError::invalid)?;
        let prefix = protobuf_length_delimited_prefix_len(record.raw().len())?;
        let start = payload_start
            .checked_sub(prefix)
            .ok_or_else(DecodeError::invalid)?;
        self.records
            .try_reserve(1)
            .map_err(|_| DecodeError::limited(DecodeLimit::Allocation { requested: 1 }))?;
        self.records.push(StagedHeaderRecord {
            start,
            end: payload_end,
            payload_start,
            payload_end,
            snapshot: record.snapshot(),
        });
        Ok(())
    }
}

fn canonical_minimal_header(index: u32, size_bits: u32) -> ([u8; 16], usize) {
    let mut output = [0u8; 16];
    let mut length = 0usize;
    length += encode_varint_array(&mut output[length..], 8);
    length += encode_varint_array(&mut output[length..], u64::from(index));
    output[length] = 21;
    length += 1;
    output[length..length + 4].copy_from_slice(&size_bits.to_le_bytes());
    length += 4;
    output[length..length + 4].copy_from_slice(&[24, 0, 32, 0]);
    length += 4;
    (output, length)
}

fn header_size_offset(source: &[u8]) -> Result<usize, DecodeError> {
    let mut remaining = source;
    let mut offset = 0usize;
    let mut found = None;
    let mut budget = Budget::new(
        source,
        DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            MAX_RECURSION,
            usize::MAX,
            usize::MAX,
        ),
    )?;
    while !remaining.is_empty() {
        let before = remaining.len();
        let field = next_field(&mut remaining, &mut budget, 1)?.ok_or_else(DecodeError::invalid)?;
        let consumed = before
            .checked_sub(remaining.len())
            .ok_or_else(DecodeError::invalid)?;
        if field.number == 2 {
            let _ = field.fixed32()?;
            if found.replace(offset + consumed - 4).is_some() {
                return Err(DecodeError::invalid());
            }
        }
        offset = offset
            .checked_add(consumed)
            .ok_or_else(DecodeError::invalid)?;
    }
    found.ok_or_else(DecodeError::invalid)
}

fn encode_key(output: &mut Vec<u8>, field: u32, wire: u8) {
    encode_varint(output, (u64::from(field) << 3) | u64::from(wire));
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

fn encode_varint_array(output: &mut [u8], mut value: u64) -> usize {
    let mut index = 0usize;
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output[index] = byte;
        index += 1;
        if value == 0 {
            return index;
        }
    }
}

fn encoded_header_field_length(payload_length: usize) -> Result<usize, DecodeError> {
    let length = u64::try_from(payload_length).map_err(|_| DecodeError::invalid())?;
    1usize
        .checked_add(encoded_varint_len(length))
        .and_then(|value| value.checked_add(payload_length))
        .ok_or_else(DecodeError::invalid)
}

fn protobuf_length_delimited_prefix_len(payload_length: usize) -> Result<usize, DecodeError> {
    let payload = u64::try_from(payload_length).map_err(|_| DecodeError::invalid())?;
    1usize
        .checked_add(encoded_varint_len(payload))
        .ok_or_else(DecodeError::invalid)
}

fn append_header_field(output: &mut Vec<u8>, payload: &[u8]) -> Result<(), DecodeError> {
    encode_key(output, 2, 2);
    encode_varint(
        output,
        u64::try_from(payload.len()).map_err(|_| DecodeError::invalid())?,
    );
    output.extend_from_slice(payload);
    Ok(())
}

fn assemble_rewritten_bucket(
    source: &[u8],
    records: &[StagedHeaderRecord],
    header_indices: &[u32],
    edits: &[HeaderSizeEdit],
    output_len: usize,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output.try_reserve_exact(output_len).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: output_len,
        })
    })?;
    let mut cursor = 0usize;
    for record in records {
        output.extend_from_slice(&source[cursor..record.start]);
        match edit_for(edits, record.snapshot.index()) {
            None => output.extend_from_slice(&source[record.start..record.end]),
            Some(edit) => match edit.replacement_size_bits {
                None if is_canonical_minimal(source, record) => {},
                replacement => {
                    let bits = replacement.unwrap_or(0);
                    if bits == record.snapshot.size_bits() {
                        output.extend_from_slice(&source[record.start..record.end]);
                    } else {
                        let raw = &source[record.payload_start..record.payload_end];
                        let size_offset = header_size_offset(raw)?;
                        encode_key(&mut output, 2, 2);
                        encode_varint(
                            &mut output,
                            u64::try_from(raw.len()).map_err(|_| DecodeError::invalid())?,
                        );
                        output.extend_from_slice(&raw[..size_offset]);
                        output.extend_from_slice(&bits.to_le_bytes());
                        output.extend_from_slice(&raw[size_offset + 4..]);
                    }
                },
            },
        }
        cursor = record.end;
    }
    output.extend_from_slice(&source[cursor..]);
    for edit in edits {
        if let Some(bits) = edit.replacement_size_bits
            && header_indices.binary_search(&edit.index).is_err()
        {
            let (payload, length) = canonical_minimal_header(edit.index, bits);
            append_header_field(&mut output, &payload[..length])?;
        }
    }
    if output.len() != output_len {
        return Err(DecodeError::invalid());
    }
    Ok(output)
}

fn edit_for(edits: &[HeaderSizeEdit], index: u32) -> Option<HeaderSizeEdit> {
    edits
        .binary_search_by_key(&index, |edit| edit.index)
        .ok()
        .map(|position| edits[position])
}

fn validate_sorted_edits(
    edits: &[HeaderSizeEdit],
    dimension_limit: u32,
) -> Result<(), DecodeError> {
    if edits
        .last()
        .is_some_and(|edit| edit.index >= dimension_limit)
        || edits.windows(2).any(|pair| pair[0].index == pair[1].index)
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn fallible_copy<T: Copy>(source: &[T]) -> Result<Vec<T>, DecodeError> {
    let mut output = Vec::new();
    output.try_reserve_exact(source.len()).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: source.len(),
        })
    })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn ceil_log2(value: usize) -> usize {
    if value <= 1 {
        0
    } else {
        usize::BITS as usize - (value - 1).leading_zeros() as usize
    }
}

fn rewrite_work_upper_bound(
    source_bytes: usize,
    output_bytes: usize,
    header_count: usize,
    edit_count: usize,
) -> Result<usize, DecodeError> {
    let header_sort = header_count
        .checked_mul(ceil_log2(header_count))
        .and_then(|work| work.checked_mul(2))
        .ok_or_else(DecodeError::invalid)?;
    let edit_sort = edit_count
        .checked_mul(ceil_log2(edit_count))
        .and_then(|work| work.checked_mul(2))
        .ok_or_else(DecodeError::invalid)?;
    let header_searches = header_count
        .checked_mul(ceil_log2(edit_count).saturating_add(1))
        .ok_or_else(DecodeError::invalid)?;
    let edit_searches = edit_count
        .checked_mul(ceil_log2(header_count).saturating_add(1))
        .ok_or_else(DecodeError::invalid)?;
    source_bytes
        .checked_add(output_bytes)
        .and_then(|work| work.checked_add(header_count))
        .and_then(|work| work.checked_add(edit_count))
        .and_then(|work| work.checked_add(header_sort))
        .and_then(|work| work.checked_add(edit_sort))
        .and_then(|work| work.checked_add(header_searches))
        .and_then(|work| work.checked_add(edit_searches))
        .ok_or_else(DecodeError::invalid)
}

fn result_decode_upper_bound(
    source: DecodeReport,
    source_bytes: usize,
    output_bytes: usize,
    inserted: usize,
    removed: usize,
    inserted_payload_bytes: usize,
    removed_payload_bytes: usize,
) -> Result<DecodeResourceUpperBound, DecodeError> {
    let fields = source
        .fields
        .checked_sub(removed.checked_mul(5).ok_or_else(DecodeError::invalid)?)
        .and_then(|value| value.checked_add(inserted.checked_mul(5)?))
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = source
        .work_bytes
        .checked_sub(
            source_bytes
                .checked_mul(2)
                .ok_or_else(DecodeError::invalid)?,
        )
        .and_then(|value| value.checked_sub(removed_payload_bytes.checked_mul(2)?))
        .and_then(|value| value.checked_add(output_bytes.checked_mul(2)?))
        .and_then(|value| value.checked_add(inserted_payload_bytes.checked_mul(2)?))
        .ok_or_else(DecodeError::invalid)?;
    Ok(DecodeResourceUpperBound {
        source_bytes: output_bytes,
        fields,
        work_bytes,
        max_depth: source.max_depth.max(if inserted == 0 { 1 } else { 2 }),
        references: source.references,
        reference_bytes: source.reference_bytes,
        text_bytes: source.text_bytes,
    })
}

fn validate_result_ceiling(
    bound: DecodeResourceUpperBound,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if bound.source_bytes > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: bound.source_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    if bound.fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: bound.fields,
            maximum: options.max_fields,
        }));
    }
    if bound.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: bound.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if bound.max_depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: bound.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    if bound.references > options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: bound.references,
            maximum: options.max_references,
        }));
    }
    if bound.text_bytes > options.max_text_bytes {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: bound.text_bytes,
            maximum: options.max_text_bytes,
        }));
    }
    Ok(())
}

fn is_canonical_minimal(source: &[u8], record: &StagedHeaderRecord) -> bool {
    let (expected, length) =
        canonical_minimal_header(record.snapshot.index(), record.snapshot.size_bits());
    source[record.payload_start..record.payload_end] == expected[..length]
}

fn decode_header_storage_bucket_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<HeaderStorageBucketSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut bucket_hash_function = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut bucket_hash_function, canonical_u32(field.varint()?)?)?,
            2 => {
                let raw = field.bytes()?;
                let snapshot = decode_header_in(raw, budget, child_depth)?;
                visitor.visit_header_record(HeaderRecord { raw, snapshot })?;
            },
            _ => {},
        }
    }
    let snapshot = HeaderStorageBucketSnapshot {
        bucket_hash_function: bucket_hash_function.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::HeaderStorageBucketArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.bucket_hash_function != snapshot.bucket_hash_function {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_header(source: &[u8], options: DecodeOptions) -> Result<HeaderSnapshot, DecodeError> {
    Ok(decode_header_with_report(source, options)?.0)
}

pub fn decode_header_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(HeaderSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_header_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_header_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<HeaderSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut index = None;
    let mut size_bits = None;
    let mut hiding_state = None;
    let mut number_of_cells = None;
    let mut cell_style = None;
    let mut text_style = None;
    let mut raw_cell_style = None;
    let mut raw_text_style = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut index, canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut size_bits, field.fixed32()?)?,
            3 => set_once(&mut hiding_state, canonical_u32(field.varint()?)?)?,
            4 => set_once(&mut number_of_cells, canonical_u32(field.varint()?)?)?,
            5 => {
                let raw = field.bytes()?;
                set_once(&mut cell_style, decode_reference(raw, budget, child_depth)?)?;
                raw_cell_style = Some(raw);
            },
            6 => {
                let raw = field.bytes()?;
                set_once(&mut text_style, decode_reference(raw, budget, child_depth)?)?;
                raw_text_style = Some(raw);
            },
            _ => {},
        }
    }
    let snapshot = HeaderSnapshot {
        index: index.ok_or_else(DecodeError::invalid)?,
        size_bits: size_bits.ok_or_else(DecodeError::invalid)?,
        hiding_state: hiding_state.ok_or_else(DecodeError::invalid)?,
        number_of_cells: number_of_cells.ok_or_else(DecodeError::invalid)?,
        cell_style,
        text_style,
    };
    budget.message(source, depth)?;
    let view: projection::HeaderArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.index != snapshot.index
        || view.size_bits != snapshot.size_bits
        || view.hiding_state != snapshot.hiding_state
        || view.number_of_cells != snapshot.number_of_cells
        || view.cell_style != raw_cell_style
        || view.text_style != raw_text_style
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_table_data_list(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableDataListSnapshot, DecodeError> {
    Ok(decode_table_data_list_with_report(source, options)?.0)
}

pub fn decode_table_data_list_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableDataListSnapshot, DecodeReport), DecodeError> {
    decode_table_data_list_with_visitor(source, options, &mut ())
}

pub fn decode_table_data_list_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TableDataListSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_data_list_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_table_data_list_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<TableDataListSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut list_type = None;
    let mut next_list_id = None;
    let mut is_new_for_bnc = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut list_type, canonical_int32(field.varint()?)?)?,
            2 => set_once(&mut next_list_id, canonical_u32(field.varint()?)?)?,
            3 => visitor.visit_list_entry(decode_table_data_list_entry_in(
                field.bytes()?,
                budget,
                child_depth,
            )?)?,
            4 => {
                let raw = field.bytes()?;
                let reference = decode_reference(raw, budget, child_depth)?;
                visitor.visit_list_segment(ReferenceRecord { raw, reference })?;
            },
            5 => set_once(&mut is_new_for_bnc, canonical_bool(field.varint()?)?)?,
            _ => {},
        }
    }
    let snapshot = TableDataListSnapshot {
        list_type: list_type.ok_or_else(DecodeError::invalid)?,
        next_list_id: next_list_id.ok_or_else(DecodeError::invalid)?,
        is_new_for_bnc,
    };
    budget.message(source, depth)?;
    let view: projection::TableDataListArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.list_type != snapshot.list_type
        || view.next_list_id != snapshot.next_list_id
        || view.is_new_for_bnc != snapshot.is_new_for_bnc
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_table_data_list_entry(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableDataListEntrySnapshot<'_>, DecodeError> {
    Ok(decode_table_data_list_entry_with_report(source, options)?.0)
}

pub fn decode_table_data_list_entry_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableDataListEntrySnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_data_list_entry_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_table_data_list_entry_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<TableDataListEntrySnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut key = None;
    let mut ref_count = None;
    let mut string_value = None;
    let mut reference = None;
    let mut formula = None;
    let mut format = None;
    let mut custom_format = None;
    let mut rich_text_payload = None;
    let mut comment_storage = None;
    let mut import_warning_set = None;
    let mut cell_spec = None;
    let mut raw_reference = None;
    let mut raw_rich_text = None;
    let mut raw_comment = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut key, canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut ref_count, canonical_u32(field.varint()?)?)?,
            3 => {
                let raw = field.bytes()?;
                set_once(&mut string_value, strict_utf8(raw, budget)?)?;
            },
            4 => {
                let raw = field.bytes()?;
                set_once(&mut reference, decode_reference(raw, budget, child_depth)?)?;
                raw_reference = Some(raw);
            },
            5 => {
                let raw = field.bytes()?;
                if formula.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                formula = Some(raw);
            },
            6 => {
                let raw = field.bytes()?;
                if format.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                format = Some(raw);
            },
            8 => {
                let raw = field.bytes()?;
                if custom_format.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                custom_format = Some(raw);
            },
            9 => {
                let raw = field.bytes()?;
                set_once(
                    &mut rich_text_payload,
                    decode_reference(raw, budget, child_depth)?,
                )?;
                raw_rich_text = Some(raw);
            },
            10 => {
                let raw = field.bytes()?;
                set_once(
                    &mut comment_storage,
                    decode_reference(raw, budget, child_depth)?,
                )?;
                raw_comment = Some(raw);
            },
            11 => {
                let raw = field.bytes()?;
                if import_warning_set.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                import_warning_set = Some(raw);
            },
            12 => {
                let raw = field.bytes()?;
                if cell_spec.is_some() {
                    return Err(DecodeError::invalid());
                }
                scan_opaque_message(raw, budget, child_depth)?;
                cell_spec = Some(raw);
            },
            _ => {},
        }
    }
    let snapshot = TableDataListEntrySnapshot {
        key: key.ok_or_else(DecodeError::invalid)?,
        ref_count: ref_count.ok_or_else(DecodeError::invalid)?,
        string_value,
        reference,
        formula,
        format,
        custom_format,
        rich_text_payload,
        comment_storage,
        import_warning_set,
        cell_spec,
    };
    budget.message(source, depth)?;
    let view: projection::TableDataListEntryArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.key != snapshot.key
        || view.ref_count != snapshot.ref_count
        || view.string_value != snapshot.string_value
        || view.reference != raw_reference
        || view.formula != snapshot.formula
        || view.format != snapshot.format
        || view.custom_format != snapshot.custom_format
        || view.rich_text_payload != raw_rich_text
        || view.comment_storage != raw_comment
        || view.import_warning_set != snapshot.import_warning_set
        || view.cell_spec != snapshot.cell_spec
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_table_data_list_segment(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableDataListSegmentSnapshot<'_>, DecodeError> {
    Ok(decode_table_data_list_segment_with_report(source, options)?.0)
}

pub fn decode_table_data_list_segment_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableDataListSegmentSnapshot<'_>, DecodeReport), DecodeError> {
    decode_table_data_list_segment_with_visitor(source, options, &mut ())
}

pub fn decode_table_data_list_segment_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TableDataListSegmentSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_data_list_segment_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_table_data_list_segment_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
) -> Result<TableDataListSegmentSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut list_type = None;
    let mut key_range = None;
    let mut key_range_location = None;
    let mut key_range_length = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut list_type, canonical_int32(field.varint()?)?)?,
            2 => {
                let raw = field.bytes()?;
                if key_range.is_some() {
                    return Err(DecodeError::invalid());
                }
                let (location, length) = decode_range(raw, budget, child_depth)?;
                key_range = Some(raw);
                key_range_location = Some(location);
                key_range_length = Some(length);
            },
            3 => visitor.visit_list_entry(decode_table_data_list_entry_in(
                field.bytes()?,
                budget,
                child_depth,
            )?)?,
            _ => {},
        }
    }
    let snapshot = TableDataListSegmentSnapshot {
        list_type: list_type.ok_or_else(DecodeError::invalid)?,
        key_range: key_range.ok_or_else(DecodeError::invalid)?,
        key_range_location: key_range_location.ok_or_else(DecodeError::invalid)?,
        key_range_length: key_range_length.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::TableDataListSegmentArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.list_type != snapshot.list_type || view.key_range != snapshot.key_range {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn decode_range(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(u32, u32), DecodeError> {
    budget.message(source, depth)?;
    let mut location = None;
    let mut length = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut location, canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut length, canonical_u32(field.varint()?)?)?,
            _ => {},
        }
    }
    Ok((
        location.ok_or_else(DecodeError::invalid)?,
        length.ok_or_else(DecodeError::invalid)?,
    ))
}

pub(crate) struct Budget {
    pub(crate) options: DecodeOptions,
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
    text_bytes: usize,
}

impl Budget {
    pub(crate) fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
            .map_err(|_conversion| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_bytes,
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
            references: 0,
            reference_bytes: 0,
            text_bytes: 0,
        })
    }

    pub(crate) fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        if source.len() > self.options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: self.options.max_message_bytes,
            }));
        }
        self.observe_depth(depth)?;
        self.work(source.len())
    }

    pub(crate) fn field(&mut self) -> Result<(), DecodeError> {
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

    pub(crate) fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
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

    fn reference(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self
            .references
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_references {
            return Err(DecodeError::limited(DecodeLimit::References {
                observed,
                maximum: self.options.max_references,
            }));
        }
        self.references = observed;
        self.reference_bytes = self
            .reference_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed,
                maximum: self.options.max_text_bytes,
            }));
        }
        self.text_bytes = observed;
        Ok(())
    }

    pub(crate) fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    pub(crate) const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            reference_bytes: self.reference_bytes,
            text_bytes: self.text_bytes,
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct Field<'source> {
    pub(crate) number: u32,
    wire_type: u8,
    value: Value<'source>,
}

impl<'source> Field<'source> {
    pub(crate) fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            Value::Varint(value) if self.wire_type == 0 => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    pub(crate) fn fixed32(self) -> Result<u32, DecodeError> {
        match self.value {
            Value::Fixed32(value) if self.wire_type == 5 => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }

    pub(crate) fn bytes(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            Value::Bytes(value) if self.wire_type == 2 => Ok(value),
            _ => Err(DecodeError::invalid()),
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum Value<'source> {
    Varint(u64),
    Fixed64,
    Bytes(&'source [u8]),
    Group,
    Fixed32(u32),
}

enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

pub(crate) fn next_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, DecodeError> {
    match parse_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(_)) => Err(DecodeError::invalid()),
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    budget.field()?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_conversion| DecodeError::invalid())?;
    let wire_type = u8::try_from(tag & 7).map_err(|_conversion| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let value = match wire_type {
        0 => Value::Varint(take_varint(source)?),
        1 => {
            let _ = take(source, 8)?;
            Value::Fixed64
        },
        2 => {
            let length = usize::try_from(take_varint(source)?)
                .map_err(|_conversion| DecodeError::invalid())?;
            Value::Bytes(take(source, length)?)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            skip_group(source, number, budget, child_depth)?;
            Value::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => Value::Fixed32(u32::from_le_bytes(
            take(source, 4)?
                .try_into()
                .map_err(|_length| DecodeError::invalid())?,
        )),
        _ => return Err(DecodeError::invalid()),
    };
    Ok(Some(ParseItem::Field(Field {
        number,
        wire_type,
        value,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, budget, depth)? {
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

pub(crate) fn take_canonical_varint(source: &mut &[u8]) -> Result<u64, DecodeError> {
    take_varint(source)
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

pub(crate) fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid()),
    }
}

pub(crate) fn canonical_u32(value: u64) -> Result<u32, DecodeError> {
    u32::try_from(value).map_err(|_conversion| DecodeError::invalid())
}

pub(crate) fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(positive) = i32::try_from(value) {
        return Ok(positive);
    }
    if value < MIN_SIGN_EXTENDED_I32 {
        return Err(DecodeError::invalid());
    }
    i32::try_from(i64::from_ne_bytes(value.to_ne_bytes()))
        .map_err(|_conversion| DecodeError::invalid())
}

pub(crate) fn scan_opaque_message(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut remaining = source;
    while next_field(&mut remaining, budget, depth)?.is_some() {}
    Ok(())
}

pub(crate) fn decode_reference(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.reference(source.len())?;
    budget.message(source, depth)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut deprecated_type, canonical_int32(field.varint()?)?)?,
            3 => set_once(
                &mut deprecated_is_external,
                canonical_bool(field.varint()?)?,
            )?,
            _ => {},
        }
    }
    let snapshot = ReferenceSnapshot {
        identifier: identifier
            .filter(|identifier| *identifier != 0)
            .ok_or_else(DecodeError::invalid)?,
        deprecated_type,
        deprecated_is_external,
    };
    if snapshot.deprecated_is_external == Some(true) {
        return Err(DecodeError::invalid());
    }
    budget.message(source, depth)?;
    let view: reference_projection::NumbersSheetReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if !view.has_identifier()
        || view.identifier != snapshot.identifier
        || view.deprecated_type != snapshot.deprecated_type
        || view.deprecated_is_external != snapshot.deprecated_is_external
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn strict_utf8<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<&'source str, DecodeError> {
    let text = str::from_utf8(source).map_err(|_error| DecodeError::invalid())?;
    budget.text(source.len())?;
    Ok(text)
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::invalid());
    }
    *slot = Some(value);
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "Focused canonical wire fixtures require exact construction and failure checks."
)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).unwrap();
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
    fn key(output: &mut Vec<u8>, number: u32, wire: u8) {
        varint(output, (u64::from(number) << 3) | u64::from(wire));
    }
    fn v(output: &mut Vec<u8>, number: u32, value: u64) {
        key(output, number, 0);
        varint(output, value);
    }
    fn b(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        key(output, number, 2);
        varint(output, u64::try_from(value.len()).unwrap());
        output.extend_from_slice(value);
    }
    fn f32_bits(output: &mut Vec<u8>, number: u32, value: u32) {
        key(output, number, 5);
        output.extend_from_slice(&value.to_le_bytes());
    }
    fn reference(id: u64) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, id);
        out
    }
    fn external_reference(id: u64) -> Vec<u8> {
        let mut out = reference(id);
        v(&mut out, 3, 1);
        out
    }
    fn row(index: u32) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, u64::from(index));
        v(&mut out, 2, 0);
        b(&mut out, 3, &[]);
        b(&mut out, 4, &[]);
        out
    }
    fn tile(rows: usize) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, 0);
        v(&mut out, 2, 0);
        v(&mut out, 3, u64::try_from(rows).unwrap());
        v(&mut out, 4, u64::try_from(rows).unwrap());
        for index in 0..rows {
            b(&mut out, 5, &row(u32::try_from(index).unwrap()));
        }
        out
    }
    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            1_000_000,
            source.len().saturating_mul(20).max(1),
            64,
            20_000,
            1_000_000,
        )
    }
    fn minimal_store() -> Vec<u8> {
        let r = reference(7);
        let mut headers = Vec::new();
        v(&mut headers, 1, 3);
        let mut out = Vec::new();
        b(&mut out, 1, &headers);
        b(&mut out, 2, &r);
        b(&mut out, 3, &[]);
        for field in 4..=6 {
            b(&mut out, field, &r);
        }
        v(&mut out, 7, 1);
        v(&mut out, 8, 2);
        b(&mut out, 9, &[]);
        b(&mut out, 10, &[]);
        b(&mut out, 11, &r);
        out
    }

    fn header_record(index: u32, size_bits: u32) -> Vec<u8> {
        let mut header = Vec::new();
        v(&mut header, 1, u64::from(index));
        f32_bits(&mut header, 2, size_bits);
        v(&mut header, 3, 0);
        v(&mut header, 4, 0);
        header
    }

    fn header_bucket(records: &[Vec<u8>]) -> Vec<u8> {
        let mut bucket = Vec::new();
        v(&mut bucket, 1, 7);
        for record in records {
            b(&mut bucket, 2, record);
        }
        bucket
    }

    #[derive(Default)]
    struct RawHeaders {
        records: Vec<(Vec<u8>, HeaderSnapshot)>,
    }

    impl StorageVisitor for RawHeaders {
        fn visit_header_record(&mut self, record: HeaderRecord<'_>) -> Result<(), DecodeError> {
            self.records
                .push((record.raw().to_vec(), record.snapshot()));
            Ok(())
        }
    }

    fn rewrite_options() -> DecodeOptions {
        DecodeOptions::new(4096, 4096, 100_000, 64, 128, 1024)
    }

    #[test]
    fn model_store_and_private_lazy_views_have_full_parity() {
        let store = minimal_store();
        let mut model = Vec::new();
        b(&mut model, 1, b"T-1");
        b(&mut model, 4, &store);
        v(&mut model, 6, 10);
        v(&mut model, 7, 20);
        b(&mut model, 34, &reference(40));
        b(&mut model, 85, &reference(85));
        let (snapshot, report) = decode_table_model_with_report(&model, options(&model)).unwrap();
        assert_eq!(snapshot.table_id(), "T-1");
        assert_eq!(snapshot.number_of_rows(), 10);
        assert_eq!(snapshot.number_of_columns(), 20);
        assert_eq!(snapshot.pivot_owner().unwrap().identifier(), 85);
        assert_eq!(report.source_bytes(), model.len());
        assert_eq!(report.references(), 7);
        assert_eq!(
            report.reference_bytes(),
            reference(7).len() * 5 + reference(40).len() + reference(85).len()
        );
        assert!(report.work_bytes() > model.len() * 2);
        assert_eq!(report.max_depth(), 3);
    }

    #[derive(Default)]
    struct Counts {
        tiles: usize,
        rows: usize,
        buckets: usize,
        headers: usize,
        entries: usize,
        segments: usize,
    }
    impl StorageVisitor for Counts {
        fn visit_tile_reference(
            &mut self,
            record: TileReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.tiles += 1;
            assert_ne!(record.reference().identifier(), 0);
            Ok(())
        }
        fn visit_tile_row(&mut self, row: TileRowInfoSnapshot<'_>) -> Result<(), DecodeError> {
            self.rows += 1;
            assert_eq!(row.cell_count(), 0);
            Ok(())
        }
        fn visit_header_bucket(
            &mut self,
            reference: ReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.buckets += 1;
            assert_ne!(reference.reference().identifier(), 0);
            Ok(())
        }
        fn visit_header(&mut self, header: HeaderSnapshot) -> Result<(), DecodeError> {
            self.headers += 1;
            assert_eq!(header.number_of_cells(), 0);
            Ok(())
        }
        fn visit_list_entry(
            &mut self,
            entry: TableDataListEntrySnapshot<'_>,
        ) -> Result<(), DecodeError> {
            self.entries += 1;
            assert_eq!(entry.ref_count(), 1);
            Ok(())
        }
        fn visit_list_segment(
            &mut self,
            reference: ReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.segments += 1;
            assert_ne!(reference.reference().identifier(), 0);
            Ok(())
        }
    }

    #[test]
    fn every_repeated_storage_route_streams_without_retention() {
        let mut counts = Counts::default();
        let mut storage = Vec::new();
        let mut tile_record = Vec::new();
        v(&mut tile_record, 1, 3);
        b(&mut tile_record, 2, &reference(30));
        b(&mut storage, 1, &tile_record);
        b(&mut storage, 1, &tile_record);
        decode_tile_storage_with_visitor(&storage, options(&storage), &mut counts).unwrap();
        let tiles = tile(2);
        decode_tile_with_visitor(&tiles, options(&tiles), &mut counts).unwrap();
        let mut headers = Vec::new();
        v(&mut headers, 1, 1);
        b(&mut headers, 2, &reference(5));
        decode_header_storage_with_visitor(&headers, options(&headers), &mut counts).unwrap();
        let mut header = Vec::new();
        v(&mut header, 1, 0);
        f32_bits(&mut header, 2, 1.0f32.to_bits());
        v(&mut header, 3, 0);
        v(&mut header, 4, 0);
        let mut bucket = Vec::new();
        v(&mut bucket, 1, 1);
        b(&mut bucket, 2, &header);
        b(&mut bucket, 2, &header);
        decode_header_storage_bucket_with_visitor(&bucket, options(&bucket), &mut counts).unwrap();
        let mut entry = Vec::new();
        v(&mut entry, 1, 1);
        v(&mut entry, 2, 1);
        b(&mut entry, 3, b"x");
        let mut list = Vec::new();
        v(&mut list, 1, 1);
        v(&mut list, 2, 2);
        b(&mut list, 3, &entry);
        b(&mut list, 4, &reference(8));
        decode_table_data_list_with_visitor(&list, options(&list), &mut counts).unwrap();
        assert_eq!(
            (
                counts.tiles,
                counts.rows,
                counts.buckets,
                counts.headers,
                counts.entries,
                counts.segments
            ),
            (2, 2, 1, 2, 1, 1)
        );
    }

    #[test]
    fn raw_header_records_and_size_rewrite_preserve_unknowns_and_order() {
        let mut third = header_record(3, 40.0f32.to_bits());
        v(&mut third, 99, 990);
        let first = header_record(1, 20.0f32.to_bits());
        let mut source = header_bucket(&[third.clone(), first.clone()]);
        v(&mut source, 100, 1);

        let mut raw = RawHeaders::default();
        decode_header_storage_bucket_with_visitor(&source, rewrite_options(), &mut raw).unwrap();
        assert_eq!(raw.records[0].0, third);
        assert_eq!(raw.records[0].1.index(), 3);
        assert_eq!(raw.records[1].0, first);

        let edits = [
            HeaderSizeEdit::set(3, 98.0f32.to_bits()),
            HeaderSizeEdit::set(2, 44.0f32.to_bits()),
        ];
        let plan = plan_header_storage_bucket_sizes(&source, 4, &edits, rewrite_options()).unwrap();
        let requirements = plan.requirements();
        let (rewritten, report) =
            execute_header_storage_bucket_size_plan(plan, rewrite_options()).unwrap();
        assert_eq!(
            (report.updated(), report.inserted(), report.removed()),
            (1, 1, 0)
        );
        assert_eq!(requirements.output_bytes(), rewritten.len());
        assert_eq!(
            requirements.result_upper_bound().fields(),
            report.result().fields()
        );
        assert_eq!(
            requirements.result_upper_bound().work_bytes(),
            report.result().work_bytes()
        );
        let mut after = RawHeaders::default();
        decode_header_storage_bucket_with_visitor(&rewritten, rewrite_options(), &mut after)
            .unwrap();
        assert_eq!(
            after
                .records
                .iter()
                .map(|record| record.1.index())
                .collect::<Vec<_>>(),
            [3, 1, 2]
        );
        assert_eq!(after.records[0].1.size_bits(), 98.0f32.to_bits());
        assert!(after.records[0].0.ends_with(&third[third.len() - 3..]));
        assert_eq!(after.records[1].0, first);
        assert!(
            rewritten
                .windows(3)
                .any(|window| window == [0xa0, 0x06, 0x01])
        );
    }

    #[test]
    fn clear_removes_only_canonical_minimal_and_otherwise_patches_positive_zero() {
        let canonical = header_record(0, 25.0f32.to_bits());
        let mut unknown = header_record(1, 30.0f32.to_bits());
        b(&mut unknown, 5, &reference(9));
        v(&mut unknown, 99, 7);
        let source = header_bucket(&[canonical, unknown.clone()]);
        let edits = [HeaderSizeEdit::remove(0), HeaderSizeEdit::remove(1)];
        let (rewritten, report) =
            rewrite_header_storage_bucket_sizes(&source, 2, &edits, rewrite_options()).unwrap();
        assert_eq!(
            (report.updated(), report.inserted(), report.removed()),
            (1, 0, 1)
        );
        let mut after = RawHeaders::default();
        decode_header_storage_bucket_with_visitor(&rewritten, rewrite_options(), &mut after)
            .unwrap();
        assert_eq!(after.records.len(), 1);
        assert_eq!(after.records[0].1.index(), 1);
        assert_eq!(after.records[0].1.size_bits(), 0.0f32.to_bits());
        assert_eq!(after.records[0].1.cell_style().unwrap().identifier(), 9);
        assert!(after.records[0].0.ends_with(&unknown[unknown.len() - 2..]));
    }

    #[test]
    fn size_rewrite_rejects_duplicate_and_out_of_range_source_or_edits() {
        let record = header_record(1, 10.0f32.to_bits());
        let duplicate = header_bucket(&[record.clone(), record]);
        assert!(
            rewrite_header_storage_bucket_sizes(
                &duplicate,
                2,
                &[HeaderSizeEdit::set(1, 20.0f32.to_bits())],
                rewrite_options(),
            )
            .is_err()
        );

        let out_of_range = header_bucket(&[header_record(2, 10.0f32.to_bits())]);
        assert!(
            rewrite_header_storage_bucket_sizes(&out_of_range, 2, &[], rewrite_options(),).is_err()
        );

        let source = header_bucket(&[header_record(0, 10.0f32.to_bits())]);
        assert!(
            rewrite_header_storage_bucket_sizes(
                &source,
                2,
                &[HeaderSizeEdit::set(1, 1), HeaderSizeEdit::remove(1)],
                rewrite_options(),
            )
            .is_err()
        );
        assert!(
            rewrite_header_storage_bucket_sizes(
                &source,
                2,
                &[HeaderSizeEdit::set(2, 1)],
                rewrite_options(),
            )
            .is_err()
        );
    }

    #[test]
    fn size_rewrite_scales_linearly_and_refuses_output_max_minus_one() {
        let make_bucket = |count: u32| {
            let mut bucket = Vec::new();
            v(&mut bucket, 1, 7);
            for index in 0..count {
                b(&mut bucket, 2, &header_record(index, 10.0f32.to_bits()));
            }
            bucket
        };
        let small = make_bucket(4096);
        let large = make_bucket(8192);
        let scalable = |source: &[u8]| {
            DecodeOptions::new(
                source.len().saturating_add(64),
                usize::MAX,
                usize::MAX,
                64,
                0,
                0,
            )
        };
        let (_, small_report) =
            rewrite_header_storage_bucket_sizes(&small, 4096, &[], scalable(&small)).unwrap();
        let (_, large_report) =
            rewrite_header_storage_bucket_sizes(&large, 8192, &[], scalable(&large)).unwrap();
        assert!(
            large_report.rewrite_work_bytes()
                <= small_report.rewrite_work_bytes().saturating_mul(23) / 10 + 64
        );

        let source = header_bucket(&[]);
        let generous = DecodeOptions::new(128, 128, 4096, 64, 0, 0);
        let plan = plan_header_storage_bucket_sizes(
            &source,
            1,
            &[HeaderSizeEdit::set(0, 10.0f32.to_bits())],
            generous,
        )
        .unwrap();
        let requirements = plan.requirements();
        let exact = requirements.output_bytes();
        assert_eq!(requirements.result_upper_bound().source_bytes(), exact);
        let error = execute_header_storage_bucket_size_plan(
            plan,
            DecodeOptions::new(exact - 1, 128, 4096, 64, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Bytes { observed, maximum })
                if observed == exact && maximum == exact - 1
        ));
    }

    #[test]
    fn all_standalone_roots_crosscheck_presence_bits_and_borrowed_payloads() {
        let row = row(4);
        let row_snapshot = decode_tile_row_info(&row, options(&row)).unwrap();
        assert_eq!(row_snapshot.tile_row_index(), 4);
        let mut header = Vec::new();
        v(&mut header, 1, 2);
        f32_bits(&mut header, 2, f32::NAN.to_bits());
        v(&mut header, 3, 3);
        v(&mut header, 4, 4);
        b(&mut header, 5, &reference(9));
        assert_eq!(
            decode_header(&header, options(&header))
                .unwrap()
                .size_bits(),
            f32::NAN.to_bits()
        );
        let mut entry = Vec::new();
        v(&mut entry, 1, 1);
        v(&mut entry, 2, 2);
        b(&mut entry, 3, "雪".as_bytes());
        b(&mut entry, 4, &reference(2));
        let e = decode_table_data_list_entry(&entry, options(&entry)).unwrap();
        assert_eq!(e.string_value(), Some("雪"));
        let mut range = Vec::new();
        v(&mut range, 1, 4);
        v(&mut range, 2, 8);
        let mut segment = Vec::new();
        v(&mut segment, 1, 1);
        b(&mut segment, 2, &range);
        b(&mut segment, 3, &entry);
        let s = decode_table_data_list_segment(&segment, options(&segment)).unwrap();
        assert_eq!((s.key_range_location(), s.key_range_length()), (4, 8));
    }

    #[test]
    fn canonical_selected_failures_and_external_references_fail_closed() {
        let mut duplicate = row(1);
        v(&mut duplicate, 1, 2);
        assert!(decode_tile_row_info(&duplicate, options(&duplicate)).is_err());
        let overlong = [0x88, 0x00, 0x00];
        assert!(decode_tile_row_info(&overlong, options(&overlong)).is_err());
        let mut bad_bool = Vec::new();
        v(&mut bad_bool, 2, 1);
        v(&mut bad_bool, 3, 2);
        assert!(decode_tile_storage(&bad_bool, options(&bad_bool)).is_err());
        let mut headers = Vec::new();
        v(&mut headers, 1, 1);
        b(&mut headers, 2, &external_reference(9));
        assert!(decode_header_storage(&headers, options(&headers)).is_err());
        let invalid_utf8 = [0x0a, 0x01, 0xff, 0x20, 0x00, 0x30, 0x00, 0x38, 0x00];
        assert!(decode_table_model(&invalid_utf8, options(&invalid_utf8)).is_err());
    }

    #[test]
    fn exact_limits_are_inclusive_and_max_minus_one_is_typed() {
        let source = tile(3);
        let (_, report) = decode_tile_with_report(&source, options(&source)).unwrap();
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
            report.text_bytes(),
        );
        assert!(decode_tile(&source, exact).is_ok());
        let fields = decode_tile(
            &source,
            DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                usize::MAX,
                64,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let work = decode_tile(
            &source,
            DecodeOptions::new(
                source.len(),
                usize::MAX,
                report.work_bytes() - 1,
                64,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            work.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let bytes = decode_tile(
            &source,
            DecodeOptions::new(
                source.len() - 1,
                usize::MAX,
                usize::MAX,
                64,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes { .. })
        ));
    }

    #[test]
    fn text_and_reference_limits_are_exact() {
        let mut entry = Vec::new();
        v(&mut entry, 1, 1);
        v(&mut entry, 2, 1);
        b(&mut entry, 3, b"abcd");
        b(&mut entry, 4, &reference(2));
        let (_, report) =
            decode_table_data_list_entry_with_report(&entry, options(&entry)).unwrap();
        assert_eq!((report.text_bytes(), report.references()), (4, 1));
        let text = decode_table_data_list_entry(
            &entry,
            DecodeOptions::new(entry.len(), 100, 1000, 64, 1, 3),
        )
        .unwrap_err();
        assert!(matches!(
            text.resource_limit(),
            Some(DecodeLimit::Text { .. })
        ));
        let refs = decode_table_data_list_entry(
            &entry,
            DecodeOptions::new(entry.len(), 100, 1000, 64, 0, 4),
        )
        .unwrap_err();
        assert!(matches!(
            refs.resource_limit(),
            Some(DecodeLimit::References { .. })
        ));
    }

    #[test]
    fn wide_4096_to_8192_routes_scale_linearly_and_max_minus_one_preempts() {
        let small_source = tile(4096);
        let large_source = tile(8192);
        let (_, small) = decode_tile_with_report(&small_source, options(&small_source)).unwrap();
        let (_, large) = decode_tile_with_report(&large_source, options(&large_source)).unwrap();
        assert_eq!(large.fields() - 4, 2 * (small.fields() - 4));
        assert!(large.work_bytes() <= small.work_bytes() * 23 / 10 + 32);
        let error = decode_tile(
            &large_source,
            DecodeOptions::new(large_source.len(), large.fields() - 1, usize::MAX, 64, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));

        let mut references = Vec::new();
        v(&mut references, 1, 1);
        for id in 1..=8192u64 {
            b(&mut references, 2, &reference(id));
        }
        let error = decode_header_storage(
            &references,
            DecodeOptions::new(references.len(), usize::MAX, usize::MAX, 64, 8191, 0),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::References {
                observed: 8192,
                maximum: 8191
            })
        ));
    }
}
