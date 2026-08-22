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
    table_name: &'source str,
    base_data_store: &'source [u8],
    number_of_rows: u32,
    number_of_columns: u32,
    table_style: Option<ReferenceSnapshot>,
    body_text_style: Option<ReferenceSnapshot>,
    header_row_text_style: Option<ReferenceSnapshot>,
    header_column_text_style: Option<ReferenceSnapshot>,
    footer_row_text_style: Option<ReferenceSnapshot>,
    body_cell_style: Option<ReferenceSnapshot>,
    header_row_style: Option<ReferenceSnapshot>,
    header_column_style: Option<ReferenceSnapshot>,
    footer_row_style: Option<ReferenceSnapshot>,
    table_name_style: Option<ReferenceSnapshot>,
    table_name_shape_style: Option<ReferenceSnapshot>,
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
    pub const fn table_name(self) -> &'source str {
        self.table_name
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
    /// Optional table-style reference from field 3.
    #[must_use]
    pub const fn table_style(self) -> Option<ReferenceSnapshot> {
        self.table_style
    }
    /// Optional body-text-style reference from field 24.
    #[must_use]
    pub const fn body_text_style(self) -> Option<ReferenceSnapshot> {
        self.body_text_style
    }
    /// Optional header-row-text-style reference from field 25.
    #[must_use]
    pub const fn header_row_text_style(self) -> Option<ReferenceSnapshot> {
        self.header_row_text_style
    }
    /// Optional header-column-text-style reference from field 26.
    #[must_use]
    pub const fn header_column_text_style(self) -> Option<ReferenceSnapshot> {
        self.header_column_text_style
    }
    /// Optional footer-row-text-style reference from field 27.
    #[must_use]
    pub const fn footer_row_text_style(self) -> Option<ReferenceSnapshot> {
        self.footer_row_text_style
    }
    /// Optional body-cell-style reference from field 18.
    #[must_use]
    pub const fn body_cell_style(self) -> Option<ReferenceSnapshot> {
        self.body_cell_style
    }
    /// Optional header-row-style reference from field 19.
    #[must_use]
    pub const fn header_row_style(self) -> Option<ReferenceSnapshot> {
        self.header_row_style
    }
    /// Optional header-column-style reference from field 20.
    #[must_use]
    pub const fn header_column_style(self) -> Option<ReferenceSnapshot> {
        self.header_column_style
    }
    /// Optional footer-row-style reference from field 21.
    #[must_use]
    pub const fn footer_row_style(self) -> Option<ReferenceSnapshot> {
        self.footer_row_style
    }
    /// Optional title-paragraph-style reference from field 30.
    #[must_use]
    pub const fn table_name_style(self) -> Option<ReferenceSnapshot> {
        self.table_name_style
    }
    /// Optional title-shape-style reference from field 36.
    #[must_use]
    pub const fn table_name_shape_style(self) -> Option<ReferenceSnapshot> {
        self.table_name_shape_style
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

/// Decode a historical table-model payload while treating omitted proto2
/// required fields as their generated default values.
///
/// Numbers compatibility fixtures and older archives may omit metadata-only
/// required fields that Prost materializes as defaults. The selected storage
/// routes remain strict; this mode only relaxes the model/datastore envelope
/// so the archive-free projection can preserve that historical behavior.
pub fn decode_table_model_compatibility_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelSnapshot<'_>, DecodeReport), DecodeError> {
    decode_table_model_compatibility_with_visitor(source, options, &mut ())
}

/// Decode a native table-model envelope while projecting only its selected
/// base-data-store routes through the compatibility envelope.
///
/// The model root, dimensions, identity, and selected optional references
/// remain strict. This narrow recovery route exists for native archives whose
/// unselected row/header metadata is malformed: the table can still be
/// opened, while a later operation that needs that metadata remains the
/// authority for rejecting the source.
pub fn decode_table_model_with_compatibility_data_store_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelSnapshot<'_>, DecodeReport), DecodeError> {
    decode_table_model_with_visitor_mode(
        source,
        options,
        &mut (),
        false,
        DataStoreProjection::DenseNative,
    )
}

/// Decode a model and stream every selected repeated storage record.
pub fn decode_table_model_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TableModelSnapshot<'source>, DecodeReport), DecodeError> {
    decode_table_model_with_visitor_mode(
        source,
        options,
        visitor,
        false,
        DataStoreProjection::Strict,
    )
}

/// Compatibility variant of [`decode_table_model_with_visitor`].
pub fn decode_table_model_compatibility_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(TableModelSnapshot<'source>, DecodeReport), DecodeError> {
    decode_table_model_with_visitor_mode(
        source,
        options,
        visitor,
        true,
        DataStoreProjection::Compatibility,
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DataStoreProjection {
    Strict,
    Compatibility,
    DenseNative,
}

fn decode_style_reference(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    data_store_projection: DataStoreProjection,
) -> Result<ReferenceSnapshot, DecodeError> {
    if data_store_projection == DataStoreProjection::Compatibility {
        // Historical compatibility model envelopes may encode omitted
        // required style references as an explicit zero-identifier proto2
        // default. Dense-native model roots remain strict; only their nested
        // DataStore metadata routes are projected compatibly.
        decode_reference_compatibility(source, budget, depth)
    } else {
        decode_reference(source, budget, depth)
    }
}

fn decode_table_model_with_visitor_mode<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
    compatibility_defaults: bool,
    data_store_projection: DataStoreProjection,
) -> Result<(TableModelSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_model_in(
        source,
        &mut budget,
        1,
        visitor,
        compatibility_defaults,
        data_store_projection,
    )?;
    Ok((snapshot, budget.report()))
}

fn decode_table_model_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
    compatibility_defaults: bool,
    data_store_projection: DataStoreProjection,
) -> Result<TableModelSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut table_id = None;
    let mut table_name = None;
    let mut base_data_store = None;
    let mut number_of_rows = None;
    let mut number_of_columns = None;
    let mut table_style = None;
    let mut body_text_style = None;
    let mut header_row_text_style = None;
    let mut header_column_text_style = None;
    let mut footer_row_text_style = None;
    let mut body_cell_style = None;
    let mut header_row_style = None;
    let mut header_column_style = None;
    let mut footer_row_style = None;
    let mut table_name_style = None;
    let mut table_name_shape_style = None;
    let mut hidden_columns = None;
    let mut hidden_rows = None;
    let mut conditional_owner = None;
    let mut pivot_owner = None;
    let mut category_owner = None;
    let mut spill_owner = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if compatibility_defaults && !matches!(field.number, 4 | 6 | 7 | 8) {
            continue;
        }
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
                let store =
                    decode_data_store_in(raw, budget, child_depth, visitor, data_store_projection);
                let _ = store?;
                base_data_store = Some(raw);
            },
            3 => {
                let raw = field.bytes()?;
                set_once(
                    &mut table_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            18 => {
                let raw = field.bytes()?;
                set_once(
                    &mut body_cell_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            19 => {
                let raw = field.bytes()?;
                set_once(
                    &mut header_row_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            20 => {
                let raw = field.bytes()?;
                set_once(
                    &mut header_column_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            21 => {
                let raw = field.bytes()?;
                set_once(
                    &mut footer_row_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            24 => {
                let raw = field.bytes()?;
                set_once(
                    &mut body_text_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            25 => {
                let raw = field.bytes()?;
                set_once(
                    &mut header_row_text_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            26 => {
                let raw = field.bytes()?;
                set_once(
                    &mut header_column_text_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            27 => {
                let raw = field.bytes()?;
                set_once(
                    &mut footer_row_text_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            30 => {
                let raw = field.bytes()?;
                set_once(
                    &mut table_name_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            36 => {
                let raw = field.bytes()?;
                set_once(
                    &mut table_name_shape_style,
                    decode_style_reference(raw, budget, child_depth, data_store_projection)?,
                )?;
            },
            6 => set_once(&mut number_of_rows, canonical_u32(field.varint()?)?)?,
            7 => set_once(&mut number_of_columns, canonical_u32(field.varint()?)?)?,
            8 if compatibility_defaults => {
                // Historical generated decoding follows protobuf's singular
                // scalar rule for this compatibility-only projection: when a
                // malformed-but-readable archive repeats the display name,
                // the final value wins.  Keep the strict table-model route
                // below unchanged; the names transaction still uses its
                // independent strict projection and rejects the duplicate
                // source before publication.
                table_name = Some(strict_utf8(field.bytes()?, budget)?);
            },
            8 => set_once(&mut table_name, strict_utf8(field.bytes()?, budget)?)?,
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
        table_id: match table_id {
            Some(value) => value,
            None if compatibility_defaults => "",
            None => return Err(DecodeError::invalid()),
        },
        table_name: match table_name {
            Some(value) => value,
            None if compatibility_defaults => "",
            None => return Err(DecodeError::invalid()),
        },
        base_data_store: match base_data_store {
            Some(value) => value,
            None if compatibility_defaults => &[],
            None => return Err(DecodeError::invalid()),
        },
        number_of_rows: match number_of_rows {
            Some(value) => value,
            None if compatibility_defaults => 0,
            None => return Err(DecodeError::invalid()),
        },
        number_of_columns: match number_of_columns {
            Some(value) => value,
            None if compatibility_defaults => 0,
            None => return Err(DecodeError::invalid()),
        },
        table_style,
        body_text_style,
        header_row_text_style,
        header_column_text_style,
        footer_row_text_style,
        body_cell_style,
        header_row_style,
        header_column_style,
        footer_row_style,
        table_name_style,
        table_name_shape_style,
        hidden_state_formula_owner_for_columns: hidden_columns.map(|(_raw, reference)| reference),
        hidden_state_formula_owner_for_rows: hidden_rows.map(|(_raw, reference)| reference),
        conditional_style_formula_owner_id: conditional_owner,
        pivot_owner: pivot_owner.map(|(_raw, reference)| reference),
        category_owner: category_owner.map(|(_raw, reference)| reference),
        spill_owner,
    };
    if !compatibility_defaults {
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
            || view.table_name != snapshot.table_name
            || view.hidden_state_formula_owner_for_columns
                != hidden_columns.map(|(raw, _reference)| raw)
            || view.hidden_state_formula_owner_for_rows != hidden_rows.map(|(raw, _reference)| raw)
            || view.conditional_style_formula_owner_id
                != snapshot.conditional_style_formula_owner_id
            || view.pivot_owner != pivot_owner.map(|(raw, _reference)| raw)
            || view.category_owner != category_owner.map(|(raw, _reference)| raw)
            || view.spill_owner != snapshot.spill_owner
        {
            return Err(DecodeError::invalid());
        }
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

/// Decode a historical base-data-store envelope with generated proto2
/// defaults for omitted metadata-only required fields. Nested tile/header
/// records and non-empty references remain strictly validated.
pub fn decode_data_store_compatibility_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DataStoreSnapshot<'_>, DecodeReport), DecodeError> {
    decode_data_store_compatibility_with_visitor(source, options, &mut ())
}

/// Decode a native base-data-store envelope while deferring only the
/// metadata-only row/header routes. Selected sidecar references and tile
/// storage remain on the strict route; this mode is used only after a native
/// model's strict envelope has identified an unselected metadata failure.
pub fn decode_data_store_dense_native_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DataStoreSnapshot<'_>, DecodeReport), DecodeError> {
    decode_data_store_with_visitor_mode(source, options, &mut (), DataStoreProjection::DenseNative)
}

pub fn decode_data_store_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(DataStoreSnapshot<'source>, DecodeReport), DecodeError> {
    decode_data_store_with_visitor_mode(source, options, visitor, DataStoreProjection::Strict)
}

/// Compatibility variant of [`decode_data_store_with_visitor`].
pub fn decode_data_store_compatibility_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
) -> Result<(DataStoreSnapshot<'source>, DecodeReport), DecodeError> {
    decode_data_store_with_visitor_mode(
        source,
        options,
        visitor,
        DataStoreProjection::Compatibility,
    )
}

fn decode_data_store_with_visitor_mode<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn StorageVisitor,
    data_store_projection: DataStoreProjection,
) -> Result<(DataStoreSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_data_store_in(source, &mut budget, 1, visitor, data_store_projection)?;
    Ok((snapshot, budget.report()))
}

fn decode_data_store_in<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
    data_store_projection: DataStoreProjection,
) -> Result<DataStoreSnapshot<'source>, DecodeError> {
    let compatibility_defaults = data_store_projection == DataStoreProjection::Compatibility;
    let dense_native = data_store_projection == DataStoreProjection::DenseNative;
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut raw_fields: [Option<&'source [u8]>; 22] = [None; 22];
    let mut refs: [Option<ReferenceSnapshot>; 22] = [None; 22];
    let mut next_row_strip_id = None;
    let mut next_column_strip_id = None;
    let mut storage_version_pre_bnc = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        // Sparse compatibility routing only owns the tile and sidecar
        // references consumed by semantic extraction. Every other DataStore
        // field is opaque there, including fields required by the native
        // proto2 schema. Dense-native routing still visits those fields so
        // their presence, framing, and duplicate keys remain bounded, but it
        // defers their nested metadata validation below.
        if compatibility_defaults && !matches!(field.number, 3 | 4 | 6 | 12 | 17 | 19) {
            continue;
        }
        let number = usize::try_from(field.number).map_err(|_conversion| DecodeError::invalid())?;
        match field.number {
            1 => {
                let raw = field.bytes()?;
                if raw_fields[0].is_some() {
                    return Err(DecodeError::invalid());
                }
                if data_store_projection == DataStoreProjection::Strict {
                    let _ = decode_header_storage_in(
                        raw,
                        budget,
                        child_depth,
                        visitor,
                        compatibility_defaults,
                    )?;
                } else {
                    // HeaderStorage is metadata-only for the archive-free
                    // table projection. Keep its source width and nesting in
                    // the finite ledger, but leave ownership/axis agreement
                    // to the table-dimension transaction's source proof.
                    budget.message(raw, child_depth)?;
                }
                raw_fields[0] = Some(raw);
            },
            3 => {
                let raw = field.bytes()?;
                if raw_fields[2].is_some() {
                    return Err(DecodeError::invalid());
                }
                // The compatibility envelope keeps the tile-storage bytes
                // borrowed and opaque.  TileStorage is a separate selected
                // route whose strict decoder owns its own required fields,
                // repeated references, and private Buffa parity check.  Do
                // not enter that generated-backed route here: a sparse
                // compatibility projection must stay generated-free and
                // leave nested tile validation to the extractor after the
                // envelope has been admitted.
                if data_store_projection != DataStoreProjection::Compatibility {
                    let _ = decode_tile_storage_in(raw, budget, child_depth, visitor)?;
                } else {
                    // Even while opaque, the selected nested message still
                    // participates in the aggregate size/depth ledger.
                    budget.message(raw, child_depth)?;
                }
                raw_fields[2] = Some(raw);
            },
            2 | 4 | 5 | 6 | 11 | 12 | 13 | 15..=22 => {
                let raw = field.bytes()?;
                if raw_fields[number - 1].is_some() {
                    return Err(DecodeError::invalid());
                }
                let strict_route = matches!(field.number, 2 | 4 | 5 | 6 | 11 | 12 | 17 | 19);
                if (!compatibility_defaults && !dense_native) || strict_route {
                    refs[number - 1] = Some(if compatibility_defaults {
                        decode_reference_compatibility(raw, budget, child_depth)?
                    } else {
                        decode_reference(raw, budget, child_depth)?
                    });
                } else {
                    // Dense-native metadata references are not consumed by
                    // semantic extraction. Their length-delimited framing is
                    // still checked above and their bounded payload width is
                    // charged without admitting a generated/reference value.
                    budget.message(raw, child_depth)?;
                }
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
    let required_bytes = |slot: Option<&'source [u8]>| match slot {
        Some(value) => Ok(value),
        None if compatibility_defaults => Ok(&[][..]),
        None => Err(DecodeError::invalid()),
    };
    let required_metadata_bytes = |slot: Option<&'source [u8]>| match slot {
        Some(value) => Ok(value),
        None if compatibility_defaults || dense_native => Ok(&[][..]),
        None => Err(DecodeError::invalid()),
    };
    let required_reference = |slot: Option<ReferenceSnapshot>| match slot {
        Some(value) => Ok(value),
        // Dense-native recovery only defers unselected metadata envelopes.
        // Every required DataStore reference remains part of the selected
        // storage contract; an omitted route must not be synthesized as the
        // compatibility zero-ID default.
        None if compatibility_defaults => Ok(default_reference()),
        None => Err(DecodeError::invalid()),
    };
    let snapshot = DataStoreSnapshot {
        // Dense-native recovery leaves row/header metadata opaque.  Its
        // absence is therefore equivalent to the generated proto2 default;
        // selected tile and sidecar routes remain required below.
        row_headers: required_metadata_bytes(raw_fields[0])?,
        column_headers: required_reference(refs[1])?,
        tiles: required_bytes(raw_fields[2])?,
        string_table: required_reference(refs[3])?,
        style_table: required_reference(refs[4])?,
        formula_table: required_reference(refs[5])?,
        next_row_strip_id: match next_row_strip_id {
            Some(value) => value,
            None if compatibility_defaults || dense_native => 0,
            None => return Err(DecodeError::invalid()),
        },
        next_column_strip_id: match next_column_strip_id {
            Some(value) => value,
            None if compatibility_defaults || dense_native => 0,
            None => return Err(DecodeError::invalid()),
        },
        row_tile_tree: required_metadata_bytes(raw_fields[8])?,
        column_tile_tree: required_metadata_bytes(raw_fields[9])?,
        format_table_pre_bnc: required_reference(refs[10])?,
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
    if !compatibility_defaults && !dense_native {
        budget.message(source, depth)?;
        let view: projection::DataStoreArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid())?;
        let raw = |field: usize| raw_fields[field - 1];
        if view.row_headers != raw(1).ok_or_else(DecodeError::invalid)?
            || view.column_headers != raw(2).ok_or_else(DecodeError::invalid)?
            || view.tiles != raw(3).ok_or_else(DecodeError::invalid)?
            || view.string_table != raw(4).ok_or_else(DecodeError::invalid)?
            || view.style_table != raw(5).ok_or_else(DecodeError::invalid)?
            || view.formula_table != raw(6).ok_or_else(DecodeError::invalid)?
            || view.next_row_strip_id != snapshot.next_row_strip_id
            || view.next_column_strip_id != snapshot.next_column_strip_id
            || view.row_tile_tree != raw(9).ok_or_else(DecodeError::invalid)?
            || view.column_tile_tree != raw(10).ok_or_else(DecodeError::invalid)?
            || view.format_table_pre_bnc != raw(11).ok_or_else(DecodeError::invalid)?
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
    let snapshot = decode_header_storage_in(source, &mut budget, 1, visitor, false)?;
    Ok((snapshot, budget.report()))
}

fn decode_header_storage_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    visitor: &mut dyn StorageVisitor,
    compatibility_defaults: bool,
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
                let reference = if compatibility_defaults {
                    decode_reference_compatibility(raw, budget, child_depth)?
                } else {
                    decode_reference(raw, budget, child_depth)?
                };
                visitor.visit_header_bucket(ReferenceRecord { raw, reference })?;
            },
            _ => {},
        }
    }
    let snapshot = HeaderStorageSnapshot {
        bucket_hash_function: match bucket_hash_function {
            Some(value) => value,
            None if compatibility_defaults => 0,
            None => return Err(DecodeError::invalid()),
        },
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
    let (preflight_output_bytes, preflight_work_bytes) = preflight_rewrite_limits(
        source_report,
        source.len(),
        records.records.len(),
        &header_indices,
        edits,
        dimension_limit,
        options,
    )?;
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
    debug_assert!(output_len <= preflight_output_bytes);
    debug_assert!(
        rewrite_work_upper_bound(
            source.len(),
            output_len,
            records.records.len(),
            ordered_edits.len(),
        )
        .is_ok_and(|work| work <= preflight_work_bytes)
    );
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
        reserve_exact(&mut indices, self.records.len())?;
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
    reserve_exact(&mut output, output_len)?;
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

fn preflight_rewrite_limits(
    source: DecodeReport,
    source_bytes: usize,
    header_count: usize,
    header_indices: &[u32],
    edits: &[HeaderSizeEdit],
    dimension_limit: u32,
    options: DecodeOptions,
) -> Result<(usize, usize), DecodeError> {
    let edit_count = edits.len();
    if edit_count > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: edit_count,
            maximum: options.max_fields,
        }));
    }
    for edit in edits {
        if edit.index >= dimension_limit {
            return Err(DecodeError::invalid());
        }
    }

    // Estimate the largest output without retaining the caller's edit slice.
    // Removes and same-size updates never increase the wire width; only an
    // insertion can do so. Duplicate edits are still rejected after staging,
    // but counting each potential insertion here keeps this bound conservative
    // and makes the resource refusal happen before that staging allocation.
    let mut output_bytes = source_bytes;
    for edit in edits {
        if let Some(bits) = edit.replacement_size_bits
            && header_indices.binary_search(&edit.index).is_err()
        {
            let (_, payload_len) = canonical_minimal_header(edit.index, bits);
            output_bytes = output_bytes
                .checked_add(encoded_header_field_length(payload_len)?)
                .ok_or_else(DecodeError::invalid)?;
        }
    }
    if output_bytes > options.max_message_bytes
        || output_bytes
            > usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError::invalid())?
    {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: output_bytes,
            maximum: options.max_message_bytes,
        }));
    }

    let work_bytes =
        rewrite_work_upper_bound(source_bytes, output_bytes, header_count, edit_count)?;
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    // The source pass has already consumed its own field budget. Only a set
    // for an absent header adds wire fields (the canonical inserted record
    // has one outer field plus four nested fields). In-place updates and
    // removals preserve or reduce the field count, so charging every edit as
    // an insertion would reject a valid rewrite when the caller deliberately
    // gives an exact field budget for the source projection.
    let inserted_fields = edits.iter().try_fold(0usize, |count, edit| {
        let additional = usize::from(
            edit.replacement_size_bits.is_some()
                && header_indices.binary_search(&edit.index).is_err(),
        )
        .checked_mul(5)
        .ok_or_else(DecodeError::invalid)?;
        count
            .checked_add(additional)
            .ok_or_else(DecodeError::invalid)
    })?;
    let result_fields = source
        .fields
        .checked_add(inserted_fields)
        .ok_or_else(DecodeError::invalid)?;
    if result_fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: result_fields,
            maximum: options.max_fields,
        }));
    }
    Ok((output_bytes, work_bytes))
}

fn fallible_copy<T: Copy>(source: &[T]) -> Result<Vec<T>, DecodeError> {
    let mut output = Vec::new();
    reserve_exact(&mut output, source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}

fn reserve_exact<T>(output: &mut Vec<T>, additional: usize) -> Result<(), DecodeError> {
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

/// Borrowed list-type envelope used to route a candidate before streaming its
/// entries.  Repeated entry/segment payloads are intentionally not retained
/// or decoded here; the full list codec remains authoritative for those
/// records after the candidate has been admitted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableDataListTypeSnapshot {
    list_type: i32,
}

impl TableDataListTypeSnapshot {
    #[must_use]
    pub const fn list_type(self) -> i32 {
        self.list_type
    }
}

/// Strictly inspect only the root `TableDataList` envelope needed for
/// candidate routing.
///
/// The handwritten pass validates canonical field framing, unknown groups,
/// and the required root scalars while skipping repeated entry payloads.  A
/// private Buffa lazy view then checks the same scalar presence/value contract
/// without allocating the generated repeated representation.  The returned
/// report accounts for both passes, so callers can charge the dispatch work
/// before deciding whether entry values may be staged.
pub fn decode_table_data_list_type_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableDataListTypeSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_data_list_type_in(source, &mut budget, 1, false)?;
    Ok((snapshot, budget.report()))
}

/// Strictly inspect only the referenced `TableDataListSegment` envelope
/// needed for candidate routing.  The segment message type is an object-local
/// compatibility route; this function makes no claim about native archive
/// type numbers and validates only the bytes supplied by its caller.
pub fn decode_table_data_list_segment_type_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableDataListTypeSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = decode_table_data_list_type_in(source, &mut budget, 1, true)?;
    Ok((snapshot, budget.report()))
}

fn decode_table_data_list_type_in(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
    segment: bool,
) -> Result<TableDataListTypeSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut list_type = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            set_once(&mut list_type, canonical_int32(field.varint()?)?)?;
        }
    }
    let snapshot = TableDataListTypeSnapshot {
        list_type: list_type.ok_or_else(DecodeError::invalid)?,
    };

    // Keep the projection pass in lockstep with the full root/segment
    // decoders.  The generated view contains only scalar envelope fields, so
    // repeated entries remain opaque and caller-owned while Buffa still
    // enforces required scalar presence and wire types.
    budget.message(source, depth)?;
    if segment {
        let view: projection::TableDataListSegmentArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid())?;
        if !view.has_list_type() || !view.has_key_range() || view.list_type != snapshot.list_type {
            return Err(DecodeError::invalid());
        }
    } else {
        let view: projection::TableDataListArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid())?;
        if !view.has_list_type() || !view.has_next_list_id() || view.list_type != snapshot.list_type
        {
            return Err(DecodeError::invalid());
        }
    }
    Ok(snapshot)
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

fn default_reference() -> ReferenceSnapshot {
    ReferenceSnapshot {
        identifier: 0,
        deprecated_type: None,
        deprecated_is_external: None,
    }
}

/// Decode a reference used by a compatibility envelope. An empty nested
/// proto2 message is the generated default for an omitted required reference;
/// any non-empty payload still goes through the strict reference decoder.
fn decode_reference_compatibility(
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
    if deprecated_is_external == Some(true) {
        return Err(DecodeError::invalid());
    }
    Ok(ReferenceSnapshot {
        identifier: identifier.unwrap_or(0),
        deprecated_type,
        deprecated_is_external,
    })
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
    use prost::Message as _;

    use crate::tst;

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
    fn unknown_group(output: &mut Vec<u8>, number: u32, nested_number: u32, value: u64) {
        key(output, number, 3);
        v(output, nested_number, value);
        key(output, number, 4);
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
    fn populated_row(index: u32, storage: &[u8], offsets: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, u64::from(index));
        v(&mut out, 2, 3);
        b(&mut out, 3, storage);
        b(&mut out, 4, offsets);
        v(&mut out, 5, 9);
        b(&mut out, 6, b"current-storage");
        b(&mut out, 7, b"current-offsets");
        v(&mut out, 8, 1);
        out
    }
    fn tile_from_rows(rows: &[Vec<u8>]) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, 0);
        v(&mut out, 2, 0);
        v(&mut out, 3, u64::try_from(rows.len()).unwrap());
        v(&mut out, 4, u64::try_from(rows.len()).unwrap());
        for payload in rows {
            b(&mut out, 5, payload);
        }
        out
    }
    fn unknown_row(index: u32) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 90, 0x9000);
        v(&mut out, 1, u64::from(index));
        b(&mut out, 91, b"unknown-before-cell-count");
        v(&mut out, 2, 4);
        b(&mut out, 3, b"pre-storage");
        v(&mut out, 92, 0x9200);
        b(&mut out, 4, b"pre-offsets");
        v(&mut out, 5, 8);
        b(&mut out, 93, b"unknown-before-current-storage");
        b(&mut out, 6, b"current-storage");
        v(&mut out, 94, 0);
        b(&mut out, 7, b"current-offsets");
        v(&mut out, 8, 0);
        b(&mut out, 95, b"unknown-after-row");
        out
    }
    fn unknown_row_with_groups(index: u32) -> Vec<u8> {
        let mut out = Vec::new();
        unknown_group(&mut out, 90, 91, 0x9000);
        v(&mut out, 1, u64::from(index));
        b(&mut out, 3, b"pre-storage");
        unknown_group(&mut out, 92, 93, 0x9200);
        v(&mut out, 2, 4);
        b(&mut out, 4, b"pre-offsets");
        v(&mut out, 5, 8);
        b(&mut out, 6, b"current-storage");
        b(&mut out, 7, b"current-offsets");
        v(&mut out, 8, 0);
        unknown_group(&mut out, 94, 95, 0x9400);
        out
    }
    fn tile_with_unknowns(row: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 90, 0x9000);
        v(&mut out, 1, 12);
        b(&mut out, 91, b"unknown-between-scalars");
        v(&mut out, 2, 34);
        v(&mut out, 92, 56);
        v(&mut out, 3, 78);
        b(&mut out, 93, b"unknown-before-numrows");
        v(&mut out, 4, 90);
        b(&mut out, 94, b"unknown-before-row");
        b(&mut out, 5, row);
        v(&mut out, 95, 0x9500);
        v(&mut out, 6, 7);
        b(&mut out, 96, b"unknown-before-bools");
        v(&mut out, 7, 0);
        v(&mut out, 8, 0);
        b(&mut out, 97, b"unknown-after-tile");
        out
    }
    fn tile_with_unknown_groups(row: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        unknown_group(&mut out, 90, 91, 0x9000);
        v(&mut out, 1, 12);
        v(&mut out, 2, 34);
        unknown_group(&mut out, 92, 93, 0x9200);
        v(&mut out, 3, 78);
        v(&mut out, 4, 90);
        b(&mut out, 5, row);
        v(&mut out, 6, 7);
        v(&mut out, 7, 0);
        unknown_group(&mut out, 94, 95, 0x9400);
        v(&mut out, 8, 0);
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

    fn minimal_store_without(omitted: u32) -> Vec<u8> {
        let r = reference(7);
        let mut headers = Vec::new();
        v(&mut headers, 1, 3);
        let mut out = Vec::new();
        if omitted != 1 {
            b(&mut out, 1, &headers);
        }
        if omitted != 2 {
            b(&mut out, 2, &r);
        }
        if omitted != 3 {
            b(&mut out, 3, &[]);
        }
        for field in 4..=6 {
            if omitted != field {
                b(&mut out, field, &r);
            }
        }
        if omitted != 7 {
            v(&mut out, 7, 1);
        }
        if omitted != 8 {
            v(&mut out, 8, 2);
        }
        if omitted != 9 {
            b(&mut out, 9, &[]);
        }
        if omitted != 10 {
            b(&mut out, 10, &[]);
        }
        if omitted != 11 {
            b(&mut out, 11, &r);
        }
        out
    }

    fn strict_model(table_name: Option<&[u8]>) -> Vec<u8> {
        let store = minimal_store();
        let mut out = Vec::new();
        b(&mut out, 1, b"T-1");
        b(&mut out, 4, &store);
        v(&mut out, 6, 10);
        v(&mut out, 7, 20);
        if let Some(table_name) = table_name {
            b(&mut out, 8, table_name);
        }
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
        b(&mut model, 8, b"Table");
        b(&mut model, 34, &reference(40));
        b(&mut model, 85, &reference(85));
        let (snapshot, report) = decode_table_model_with_report(&model, options(&model)).unwrap();
        assert_eq!(snapshot.table_id(), "T-1");
        assert_eq!(snapshot.table_name(), "Table");
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

    #[test]
    fn strict_model_requires_table_name_but_keeps_unknown_wire_opaque() {
        let store = minimal_store();
        let mut valid = strict_model(Some(b"Table"));
        // Keep the generated projection's high-numbered known `spill_owner`
        // field (93) out of this unknown-field matrix: `unknown_fields`
        // appends three additional field numbers after its base.
        unknown_fields(&mut valid, 100);
        let before = valid.clone();
        let (snapshot, _) = decode_table_model_with_report(&valid, options(&valid)).unwrap();
        assert_eq!(snapshot.table_name(), "Table");
        assert_eq!(snapshot.base_data_store(), store.as_slice());
        assert_eq!(valid, before);

        let missing = strict_model(None);
        assert!(decode_table_model_with_report(&missing, options(&missing)).is_err());

        let mut wrong_wire = strict_model(None);
        v(&mut wrong_wire, 8, 1);
        assert!(decode_table_model_with_report(&wrong_wire, options(&wrong_wire)).is_err());

        let mut duplicate = strict_model(Some(b"Table"));
        b(&mut duplicate, 8, b"Second");
        assert!(decode_table_model_with_report(&duplicate, options(&duplicate)).is_err());
    }

    #[test]
    fn compatibility_model_store_projection_is_sparse_and_generated_free() {
        // The compatibility envelope intentionally omits the native proto2
        // required metadata and carries an opaque TileStorage payload.  The
        // selected sidecar references use both an empty nested message and a
        // zero identifier, matching generated proto2 defaults while retaining
        // borrowed source slices.
        let mut store = Vec::new();
        b(&mut store, 1, &[0xff]); // unselected/opaque HeaderStorage
        b(&mut store, 1, &[0x80]); // duplicate unselected key is skipped
        b(&mut store, 3, &[0xff]); // nested TileStorage is validated later
        b(&mut store, 4, &[]); // empty Reference => identifier 0
        b(&mut store, 6, &reference(0)); // explicit zero identifier
        b(&mut store, 12, &[]);
        b(&mut store, 17, &[]);
        b(&mut store, 19, &[]);
        unknown_group(&mut store, 90, 91, 0x9000);

        let mut model = Vec::new();
        b(&mut model, 4, &store);
        v(&mut model, 7, 3);
        b(&mut model, 8, b"Sparse");
        b(&mut model, 1, &[0xff]); // unselected table-id wire is opaque
        unknown_group(&mut model, 92, 93, 0x9200);

        let (snapshot, report) =
            decode_table_model_compatibility_with_report(&model, options(&model)).unwrap();
        assert_eq!(snapshot.table_id(), "");
        assert_eq!(snapshot.table_name(), "Sparse");
        assert_eq!(snapshot.number_of_rows(), 0);
        assert_eq!(snapshot.number_of_columns(), 3);
        assert_eq!(snapshot.base_data_store(), store.as_slice());
        assert!(report.fields() > 0);

        let store_options = options(&store);
        let (store_snapshot, _store_report) =
            decode_data_store_compatibility_with_report(&store, store_options).unwrap();
        assert_eq!(store_snapshot.tiles(), &[0xff]);
        assert_eq!(store_snapshot.string_table().identifier(), 0);
        assert_eq!(store_snapshot.formula_table().identifier(), 0);
        assert_eq!(
            store_snapshot
                .formula_error_table()
                .expect("explicit empty optional reference")
                .identifier(),
            0
        );
        assert_eq!(
            store_snapshot
                .rich_text_table()
                .expect("explicit empty optional reference")
                .identifier(),
            0
        );
        assert_eq!(
            store_snapshot
                .comment_storage_table()
                .expect("explicit empty optional reference")
                .identifier(),
            0
        );

        // Strict mode still preserves its required-field/Buffa parity
        // contract and therefore does not silently become the compatibility
        // route.
        assert!(decode_table_model_with_report(&model, options(&model)).is_err());
    }

    #[test]
    fn compatibility_selected_duplicates_and_wrong_wires_fail() {
        let mut store = Vec::new();
        b(&mut store, 4, &[]);
        b(&mut store, 4, &[]);
        let mut model = Vec::new();
        b(&mut model, 4, &store);
        assert!(decode_table_model_compatibility_with_report(&model, options(&model)).is_err());

        let mut wrong_store = Vec::new();
        v(&mut wrong_store, 4, 0);
        let mut wrong_model = Vec::new();
        b(&mut wrong_model, 4, &wrong_store);
        assert!(
            decode_table_model_compatibility_with_report(&wrong_model, options(&wrong_model))
                .is_err()
        );
    }

    #[test]
    fn dense_native_keeps_required_and_selected_references_strict() {
        for field in [2, 4, 5, 6, 11] {
            let source = minimal_store_without(field);
            assert!(
                decode_data_store_dense_native_with_report(&source, options(&source)).is_err(),
                "dense-native accepted omitted required datastore field {field}"
            );
        }

        for field in [4, 12, 17, 19] {
            let mut source = minimal_store();
            b(&mut source, field, &[0x80]);
            assert!(
                decode_data_store_dense_native_with_report(&source, options(&source)).is_err(),
                "dense-native accepted malformed selected datastore reference {field}"
            );
        }
    }

    #[test]
    fn dense_native_accepts_valid_store() {
        let source = minimal_store();
        let (snapshot, _report) =
            decode_data_store_dense_native_with_report(&source, options(&source)).unwrap();

        assert_eq!(snapshot.column_headers().identifier(), 7);
        assert_eq!(snapshot.string_table().identifier(), 7);
        assert_eq!(snapshot.style_table().identifier(), 7);
        assert_eq!(snapshot.formula_table().identifier(), 7);
        assert_eq!(snapshot.format_table_pre_bnc().identifier(), 7);
    }

    #[test]
    fn dense_native_accepts_malformed_unselected_reference() {
        let mut source = minimal_store();
        b(&mut source, 13, &[0x80]);

        let (snapshot, _report) =
            decode_data_store_dense_native_with_report(&source, options(&source)).unwrap();
        assert!(snapshot.merge_region_map().is_none());
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

    #[derive(Debug, PartialEq, Eq)]
    struct RowFact {
        tile_row_index: u32,
        cell_count: u32,
        storage_version: Option<u32>,
        has_wide_offsets: Option<bool>,
        cell_storage_buffer_pre_bnc: Vec<u8>,
        cell_offsets_pre_bnc: Vec<u8>,
        cell_storage_buffer: Option<Vec<u8>>,
        cell_offsets: Option<Vec<u8>>,
    }

    fn row_fact(row: TileRowInfoSnapshot<'_>) -> RowFact {
        RowFact {
            tile_row_index: row.tile_row_index(),
            cell_count: row.cell_count(),
            storage_version: row.storage_version(),
            has_wide_offsets: row.has_wide_offsets(),
            cell_storage_buffer_pre_bnc: row.cell_storage_buffer_pre_bnc().to_vec(),
            cell_offsets_pre_bnc: row.cell_offsets_pre_bnc().to_vec(),
            cell_storage_buffer: row.cell_storage_buffer().map(<[u8]>::to_vec),
            cell_offsets: row.cell_offsets().map(<[u8]>::to_vec),
        }
    }

    fn prost_row_fact(row: &tst::TileRowInfo) -> RowFact {
        RowFact {
            tile_row_index: row.tile_row_index,
            cell_count: row.cell_count,
            storage_version: row.storage_version,
            has_wide_offsets: row.has_wide_offsets,
            cell_storage_buffer_pre_bnc: row.cell_storage_buffer_pre_bnc.clone(),
            cell_offsets_pre_bnc: row.cell_offsets_pre_bnc.clone(),
            cell_storage_buffer: row.cell_storage_buffer.clone(),
            cell_offsets: row.cell_offsets.clone(),
        }
    }

    #[derive(Default)]
    struct RowCollector {
        rows: Vec<RowFact>,
        source_range: Option<(usize, usize)>,
        borrowed_payloads: Vec<(usize, usize)>,
    }

    impl RowCollector {
        fn for_source(source: &[u8]) -> Self {
            let start = source.as_ptr() as usize;
            let end = start.saturating_add(source.len());
            Self {
                rows: Vec::new(),
                source_range: Some((start, end)),
                borrowed_payloads: Vec::new(),
            }
        }

        fn assert_borrowed(&mut self, payload: &[u8]) {
            let Some((source_start, source_end)) = self.source_range else {
                return;
            };
            if payload.is_empty() {
                return;
            }
            let payload_start = payload.as_ptr() as usize;
            let payload_end = payload_start.saturating_add(payload.len());
            assert!(payload_start >= source_start);
            assert!(payload_end <= source_end);
            self.borrowed_payloads.push((payload_start, payload.len()));
        }
    }

    impl StorageVisitor for RowCollector {
        fn visit_tile_row(&mut self, row: TileRowInfoSnapshot<'_>) -> Result<(), DecodeError> {
            self.assert_borrowed(row.cell_storage_buffer_pre_bnc());
            self.assert_borrowed(row.cell_offsets_pre_bnc());
            if let Some(payload) = row.cell_storage_buffer() {
                self.assert_borrowed(payload);
            }
            if let Some(payload) = row.cell_offsets() {
                self.assert_borrowed(payload);
            }
            self.rows.push(row_fact(row));
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
    fn rewrite_preflight_rejects_exhausted_edit_budget_before_staging() {
        let source = header_bucket(&[]);
        let before = source.clone();
        let edits = [HeaderSizeEdit::set(0, 10.0f32.to_bits())];
        let error = plan_header_storage_bucket_sizes(
            &source,
            1,
            &edits,
            DecodeOptions::new(128, 1, 4096, 64, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 6,
                maximum: 1
            })
        ));
        assert_eq!(source, before);
        assert_eq!(edits[0].index(), 0);
    }

    #[test]
    fn rewrite_in_place_update_accepts_exact_source_field_budget_and_preserves_unknowns() {
        let mut record = header_record(1, 10.0f32.to_bits());
        b(&mut record, 99, b"opaque");
        let mut source = header_bucket(&[record.clone()]);
        v(&mut source, 100, 77);
        let before = source.clone();
        let (_, source_report) =
            decode_header_storage_bucket_with_report(&source, rewrite_options()).unwrap();
        let exact_fields = DecodeOptions::new(
            source.len(),
            source_report.fields(),
            usize::MAX,
            source_report.max_depth(),
            source_report.references(),
            source_report.text_bytes(),
        );

        let (rewritten, report) = rewrite_header_storage_bucket_sizes(
            &source,
            2,
            &[HeaderSizeEdit::set(1, 20.0f32.to_bits())],
            exact_fields,
        )
        .expect("an in-place update should not consume additional fields");

        assert_eq!(source, before);
        assert_eq!(report.source().fields(), source_report.fields());
        assert_eq!(report.result().fields(), source_report.fields());
        assert_eq!(
            (report.updated(), report.inserted(), report.removed()),
            (1, 0, 0)
        );
        assert!(rewritten.ends_with(&[0xa0, 0x06, 77]));

        let mut after = RawHeaders::default();
        decode_header_storage_bucket_with_visitor(&rewritten, rewrite_options(), &mut after)
            .unwrap();
        assert_eq!(after.records.len(), 1);
        assert_eq!(after.records[0].1.size_bits(), 20.0f32.to_bits());
        assert!(after.records[0].0.ends_with(b"opaque"));
        assert_eq!(after.records[0].0.len(), record.len());
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
    fn tile_visitor_matches_prost_for_unsorted_duplicate_indices_and_presence() {
        let expected = tst::Tile {
            max_column: 12,
            max_row: 34,
            num_cells: 56,
            numrows: 78,
            row_infos: vec![
                tst::TileRowInfo {
                    tile_row_index: 11,
                    cell_count: 2,
                    cell_storage_buffer_pre_bnc: vec![1, 2],
                    cell_offsets_pre_bnc: vec![3],
                    storage_version: Some(7),
                    cell_storage_buffer: Some(vec![4, 5]),
                    cell_offsets: Some(vec![6]),
                    has_wide_offsets: Some(false),
                },
                tst::TileRowInfo {
                    tile_row_index: 3,
                    cell_count: 0,
                    cell_storage_buffer_pre_bnc: Vec::new(),
                    cell_offsets_pre_bnc: Vec::new(),
                    storage_version: None,
                    cell_storage_buffer: None,
                    cell_offsets: Some(vec![8]),
                    has_wide_offsets: Some(true),
                },
                tst::TileRowInfo {
                    tile_row_index: 11,
                    cell_count: 9,
                    cell_storage_buffer_pre_bnc: vec![9],
                    cell_offsets_pre_bnc: vec![10],
                    storage_version: Some(12),
                    cell_storage_buffer: Some(vec![11]),
                    cell_offsets: None,
                    has_wide_offsets: None,
                },
            ],
            storage_version: Some(90),
            last_saved_in_bnc: Some(false),
            should_use_wide_rows: Some(false),
        };
        let source = expected.encode_to_vec();
        let before = source.clone();
        let prost = tst::Tile::decode(source.as_slice()).unwrap();
        let mut visitor = RowCollector::for_source(&source);
        let (snapshot, report) =
            decode_tile_with_visitor(&source, options(&source), &mut visitor).unwrap();

        assert_eq!(source, before);
        assert_eq!(snapshot.max_column(), prost.max_column);
        assert_eq!(snapshot.max_row(), prost.max_row);
        assert_eq!(snapshot.num_cells(), prost.num_cells);
        assert_eq!(snapshot.num_rows(), prost.numrows);
        assert_eq!(snapshot.storage_version(), prost.storage_version);
        assert_eq!(snapshot.last_saved_in_bnc(), prost.last_saved_in_bnc);
        assert_eq!(snapshot.should_use_wide_rows(), prost.should_use_wide_rows);
        assert_eq!(visitor.rows.len(), prost.row_infos.len());
        assert_eq!(
            visitor.rows,
            prost
                .row_infos
                .iter()
                .map(prost_row_fact)
                .collect::<Vec<_>>()
        );
        assert_ne!(
            usize::try_from(snapshot.num_rows()).unwrap(),
            visitor.rows.len()
        );
        assert_ne!(
            usize::try_from(snapshot.num_cells()).unwrap(),
            visitor
                .rows
                .iter()
                .map(|row| usize::try_from(row.cell_count).unwrap())
                .sum::<usize>()
        );
        assert!(report.fields() > 0);
    }

    #[test]
    fn tile_visitor_accepts_unknowns_and_keeps_source_unchanged() {
        let source = tile_with_unknowns(&unknown_row(5));
        let before = source.clone();
        let mut visitor = RowCollector::for_source(&source);
        let snapshot = decode_tile_with_visitor(&source, options(&source), &mut visitor)
            .unwrap()
            .0;

        assert_eq!(source, before);
        assert_eq!(snapshot.max_column(), 12);
        assert_eq!(snapshot.max_row(), 34);
        assert_eq!(snapshot.num_cells(), 78);
        assert_eq!(snapshot.num_rows(), 90);
        assert_eq!(snapshot.storage_version(), Some(7));
        assert_eq!(snapshot.last_saved_in_bnc(), Some(false));
        assert_eq!(snapshot.should_use_wide_rows(), Some(false));
        assert_eq!(visitor.rows.len(), 1);
        let decoded_row = &visitor.rows[0];
        assert_eq!(decoded_row.tile_row_index, 5);
        assert_eq!(decoded_row.cell_count, 4);
        assert_eq!(decoded_row.cell_storage_buffer_pre_bnc, b"pre-storage");
        assert_eq!(decoded_row.cell_offsets_pre_bnc, b"pre-offsets");
        assert_eq!(decoded_row.storage_version, Some(8));
        assert_eq!(
            decoded_row.cell_storage_buffer,
            Some(b"current-storage".to_vec())
        );
        assert_eq!(decoded_row.cell_offsets, Some(b"current-offsets".to_vec()));
        assert_eq!(decoded_row.has_wide_offsets, Some(false));
    }

    #[test]
    fn tile_visitor_accepts_matched_unknown_groups_at_root_and_row_scopes() {
        let row_source = unknown_row_with_groups(5);
        let row_before = row_source.clone();
        let row_snapshot = decode_tile_row_info(&row_source, options(&row_source)).unwrap();
        assert_eq!(row_source, row_before);
        assert_eq!(row_snapshot.tile_row_index(), 5);
        assert_eq!(row_snapshot.cell_count(), 4);
        assert_eq!(row_snapshot.cell_storage_buffer_pre_bnc(), b"pre-storage");
        assert_eq!(row_snapshot.cell_offsets_pre_bnc(), b"pre-offsets");
        assert_eq!(row_snapshot.storage_version(), Some(8));
        assert_eq!(
            row_snapshot.cell_storage_buffer(),
            Some(&b"current-storage"[..])
        );
        assert_eq!(row_snapshot.cell_offsets(), Some(&b"current-offsets"[..]));
        assert_eq!(row_snapshot.has_wide_offsets(), Some(false));

        let source = tile_with_unknown_groups(&row_source);
        let before = source.clone();
        let mut visitor = RowCollector::for_source(&source);
        let (snapshot, visitor_report) =
            decode_tile_with_visitor(&source, options(&source), &mut visitor).unwrap();
        let (scalar, scalar_report) = decode_tile_with_report(&source, options(&source)).unwrap();

        assert_eq!(source, before);
        assert_eq!(snapshot, scalar);
        assert_eq!(visitor_report, scalar_report);
        assert_eq!(
            (
                snapshot.max_column(),
                snapshot.max_row(),
                snapshot.num_cells(),
                snapshot.num_rows(),
                snapshot.storage_version(),
                snapshot.last_saved_in_bnc(),
                snapshot.should_use_wide_rows(),
            ),
            (12, 34, 78, 90, Some(7), Some(false), Some(false))
        );
        assert_eq!(visitor.rows.len(), 1);
        assert_eq!(visitor.rows[0].tile_row_index, 5);
        assert_eq!(visitor.rows[0].cell_count, 4);
        assert_eq!(visitor.rows[0].cell_storage_buffer_pre_bnc, b"pre-storage");
        assert_eq!(visitor.rows[0].cell_offsets_pre_bnc, b"pre-offsets");
        assert_eq!(visitor.rows[0].storage_version, Some(8));
        assert_eq!(
            visitor.rows[0].cell_storage_buffer,
            Some(b"current-storage".to_vec())
        );
        assert_eq!(
            visitor.rows[0].cell_offsets,
            Some(b"current-offsets".to_vec())
        );
        assert_eq!(visitor.rows[0].has_wide_offsets, Some(false));
    }

    #[test]
    fn tile_visitor_and_scalar_paths_have_report_and_row_parity() {
        let first = populated_row(7, b"pre-storage", b"pre-offsets");
        let second = unknown_row(2);
        let source = tile_from_rows(&[first, second]);
        let mut visitor = RowCollector::for_source(&source);
        let (visited, visitor_report) =
            decode_tile_with_visitor(&source, options(&source), &mut visitor).unwrap();
        let (scalar, scalar_report) = decode_tile_with_report(&source, options(&source)).unwrap();

        assert_eq!(visited, scalar);
        assert_eq!(visitor_report, scalar_report);
        assert_eq!(visitor.rows.len(), 2);
    }

    #[test]
    fn tile_visitor_exact_nesting_limit_is_inclusive() {
        let source = tile(1);
        let (_, report) = decode_tile_with_report(&source, options(&source)).unwrap();
        assert_eq!(report.max_depth(), 2);
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
            report.text_bytes(),
        );
        let mut visitor = RowCollector::for_source(&source);
        let (_, exact_report) = decode_tile_with_visitor(&source, exact, &mut visitor).unwrap();
        assert_eq!(exact_report, report);

        let one_less = DecodeOptions::new(
            source.len(),
            usize::MAX,
            usize::MAX,
            report.max_depth() - 1,
            usize::MAX,
            usize::MAX,
        );
        let error = decode_tile_with_visitor(&source, one_less, &mut ()).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1
            })
        ));
    }

    #[test]
    fn tile_visitor_rejects_malformed_required_wire_and_group_forms() {
        let valid = tile(1);

        let mut missing_tile_required = valid.clone();
        assert_eq!(&missing_tile_required[..2], [0x08, 0x00]);
        missing_tile_required.drain(..2);

        let mut duplicate_tile_required = valid.clone();
        v(&mut duplicate_tile_required, 1, 0);

        let mut wrong_tile_wire = Vec::new();
        b(&mut wrong_tile_wire, 1, &[]);
        wrong_tile_wire.extend_from_slice(&valid[2..]);

        let mut noncanonical_tile_varint = Vec::new();
        key(&mut noncanonical_tile_varint, 1, 0);
        noncanonical_tile_varint.extend_from_slice(&[0x80, 0x00]);
        noncanonical_tile_varint.extend_from_slice(&valid[2..]);

        let mut overflowing_tile_u32 = Vec::new();
        v(
            &mut overflowing_tile_u32,
            1,
            u64::from(u32::MAX).saturating_add(1),
        );
        overflowing_tile_u32.extend_from_slice(&valid[2..]);

        let mut invalid_tile_bool = valid.clone();
        v(&mut invalid_tile_bool, 7, 2);

        let mut truncated_tile = valid.clone();
        truncated_tile.pop();

        let mut missing_row_required = row(1);
        missing_row_required.drain(..2);
        let missing_row_required = tile_from_rows(&[missing_row_required]);

        let mut duplicate_row_required = row(1);
        v(&mut duplicate_row_required, 1, 2);
        let duplicate_row_required = tile_from_rows(&[duplicate_row_required]);

        let canonical_row = row(1);
        let mut wrong_row_wire = Vec::new();
        b(&mut wrong_row_wire, 1, &[]);
        wrong_row_wire.extend_from_slice(&canonical_row[2..]);
        let wrong_row_wire = tile_from_rows(&[wrong_row_wire]);

        let mut noncanonical_row_varint = Vec::new();
        key(&mut noncanonical_row_varint, 1, 0);
        noncanonical_row_varint.extend_from_slice(&[0x80, 0x00]);
        noncanonical_row_varint.extend_from_slice(&canonical_row[2..]);
        let noncanonical_row_varint = tile_from_rows(&[noncanonical_row_varint]);

        let mut overflowing_row_u32 = Vec::new();
        v(
            &mut overflowing_row_u32,
            1,
            u64::from(u32::MAX).saturating_add(1),
        );
        overflowing_row_u32.extend_from_slice(&canonical_row[2..]);
        let overflowing_row_u32 = tile_from_rows(&[overflowing_row_u32]);

        let mut invalid_row_bool = canonical_row.clone();
        v(&mut invalid_row_bool, 8, 2);
        let invalid_row_bool = tile_from_rows(&[invalid_row_bool]);

        let mut truncated_row = tile_from_rows(&[canonical_row]);
        truncated_row.pop();

        let mut unclosed_group = valid.clone();
        key(&mut unclosed_group, 90, 3);

        let mut mismatched_group = valid.clone();
        key(&mut mismatched_group, 90, 3);
        v(&mut mismatched_group, 91, 1);
        key(&mut mismatched_group, 91, 4);

        let mut stray_end_group = valid;
        key(&mut stray_end_group, 90, 4);

        let cases = [
            ("missing tile required", missing_tile_required),
            ("duplicate tile required", duplicate_tile_required),
            ("wrong tile wire", wrong_tile_wire),
            ("noncanonical tile varint", noncanonical_tile_varint),
            ("overflowing tile uint32", overflowing_tile_u32),
            ("invalid tile bool", invalid_tile_bool),
            ("truncated tile", truncated_tile),
            ("missing row required", missing_row_required),
            ("duplicate row required", duplicate_row_required),
            ("wrong row wire", wrong_row_wire),
            ("noncanonical row varint", noncanonical_row_varint),
            ("overflowing row uint32", overflowing_row_u32),
            ("invalid row bool", invalid_row_bool),
            ("truncated row", truncated_row),
            ("unclosed group", unclosed_group),
            ("mismatched group", mismatched_group),
            ("stray end group", stray_end_group),
        ];
        for (label, source) in cases {
            let mut visitor = RowCollector::default();
            assert!(
                decode_tile_with_visitor(&source, options(&source), &mut visitor).is_err(),
                "{label} unexpectedly decoded"
            );
        }
    }

    #[test]
    fn tile_visitor_callbacks_can_observe_rows_before_later_malformed_row() {
        let first = row(4);
        let mut malformed_later = row(9);
        v(&mut malformed_later, 1, 10);
        let source = tile_from_rows(&[first.clone(), malformed_later]);
        let before = source.clone();

        // Visitor callbacks are streaming by contract. A later row failure
        // cannot roll back a callback that already observed the first row.
        let mut visitor = RowCollector::for_source(&source);
        assert!(decode_tile_with_visitor(&source, options(&source), &mut visitor).is_err());
        assert_eq!(visitor.rows.len(), 1);
        assert_eq!(
            visitor.rows[0],
            row_fact(decode_tile_row_info(&first, options(&first)).unwrap())
        );

        assert_eq!(source, before);
    }

    #[test]
    fn tile_visitor_returns_source_borrowed_rows_and_exact_report() {
        let storage = [0x11, 0x22, 0x33];
        let offsets = [0x44, 0x55];
        let first = populated_row(4, &storage, &offsets);
        let second = populated_row(7, b"second-storage", b"second-offsets");
        let mut source = Vec::new();
        v(&mut source, 1, 12);
        v(&mut source, 2, 8);
        v(&mut source, 3, 3);
        v(&mut source, 4, 9);
        b(&mut source, 5, &first);
        b(&mut source, 5, &second);
        v(&mut source, 6, 10);
        v(&mut source, 7, 1);
        v(&mut source, 8, 0);

        let mut visitor = RowCollector::for_source(&source);
        let (snapshot, rows_report) =
            decode_tile_with_visitor(&source, options(&source), &mut visitor).unwrap();
        let (_, scalar_report) = decode_tile_with_report(&source, options(&source)).unwrap();
        assert_eq!(rows_report, scalar_report);
        assert_eq!(snapshot.max_column(), 12);
        assert_eq!(snapshot.max_row(), 8);
        assert_eq!(snapshot.num_cells(), 3);
        assert_eq!(snapshot.num_rows(), 9);
        assert_eq!(snapshot.storage_version(), Some(10));
        assert_eq!(snapshot.last_saved_in_bnc(), Some(true));
        assert_eq!(snapshot.should_use_wide_rows(), Some(false));
        assert_eq!(visitor.rows.len(), 2);
        let first_row = &visitor.rows[0];
        assert_eq!(first_row.tile_row_index, 4);
        assert_eq!(first_row.cell_count, 3);
        assert_eq!(first_row.storage_version, Some(9));
        assert_eq!(
            first_row.cell_storage_buffer,
            Some(b"current-storage".to_vec())
        );
        assert_eq!(first_row.cell_offsets, Some(b"current-offsets".to_vec()));
        assert_eq!(first_row.has_wide_offsets, Some(true));
        let storage_offset = source
            .windows(storage.len())
            .position(|window| window == storage)
            .unwrap();
        let offsets_offset = source
            .windows(offsets.len())
            .position(|window| window == offsets)
            .unwrap();
        assert!(
            visitor
                .borrowed_payloads
                .contains(&(source[storage_offset..].as_ptr() as usize, storage.len()))
        );
        assert!(
            visitor
                .borrowed_payloads
                .contains(&(source[offsets_offset..].as_ptr() as usize, offsets.len()))
        );
        assert_eq!(visitor.rows[1].tile_row_index, 7);
    }

    #[test]
    fn tile_visitor_preserves_failure_atomicity_and_all_decode_limits() {
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
        let mut exact_visitor = RowCollector::for_source(&source);
        let (_, exact_report) =
            decode_tile_with_visitor(&source, exact, &mut exact_visitor).unwrap();
        assert_eq!(exact_report, report);

        let fields = decode_tile_with_visitor(
            &source,
            DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                usize::MAX,
                64,
                usize::MAX,
                usize::MAX,
            ),
            &mut (),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));

        let work = decode_tile_with_visitor(
            &source,
            DecodeOptions::new(
                source.len(),
                usize::MAX,
                report.work_bytes() - 1,
                64,
                usize::MAX,
                usize::MAX,
            ),
            &mut (),
        )
        .unwrap_err();
        assert!(matches!(
            work.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));

        let bytes = decode_tile_with_visitor(
            &source,
            DecodeOptions::new(
                source.len() - 1,
                usize::MAX,
                usize::MAX,
                64,
                usize::MAX,
                usize::MAX,
            ),
            &mut (),
        )
        .unwrap_err();
        assert!(matches!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes { .. })
        ));

        let mut malformed_row = row(2);
        v(&mut malformed_row, 1, 3);
        let mut malformed = tile(2);
        b(&mut malformed, 5, &malformed_row);
        assert!(decode_tile_with_visitor(&malformed, options(&malformed), &mut ()).is_err());
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

    fn fixed64(output: &mut Vec<u8>, number: u32, value: u64) {
        key(output, number, 1);
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn fixed32(output: &mut Vec<u8>, number: u32, value: u32) {
        key(output, number, 5);
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn unknown_fields(output: &mut Vec<u8>, number: u32) {
        v(output, number, 0xfeed);
        b(output, number + 1, b"unknown-bytes");
        fixed64(output, number + 2, 0x0102_0304_0506_0708);
        fixed32(output, number + 3, 0x090a_0b0c);
        unknown_group(output, number + 4, number + 5, 0xbeef);
    }

    fn unknown_groups_to_depth(output: &mut Vec<u8>, number: u32, depth: usize) {
        key(output, number, 3);
        if depth > 1 {
            unknown_groups_to_depth(output, number + 1, depth - 1);
        } else {
            v(output, number + 1, 1);
        }
        key(output, number, 4);
    }

    fn prost_reference(
        identifier: u64,
        deprecated_type: Option<i32>,
        deprecated_is_external: Option<bool>,
    ) -> crate::tsp::Reference {
        crate::tsp::Reference {
            identifier,
            deprecated_type,
            deprecated_is_external,
        }
    }

    fn formula_oracle() -> crate::tsce::FormulaArchive {
        crate::tsce::FormulaArchive {
            ast_node_array: crate::tsce::AstNodeArrayArchive::default(),
            host_column: Some(7),
            host_row: Some(8),
            host_column_is_negative: Some(false),
            host_row_is_negative: Some(true),
            ..Default::default()
        }
    }

    fn format_oracle() -> crate::tsk::FormatStructArchive {
        crate::tsk::FormatStructArchive {
            format_type: Some(4),
            decimal_places: Some(2),
            currency_code: Some("CNY".to_owned()),
            show_thousands_separator: Some(false),
            ..Default::default()
        }
    }

    fn custom_format_oracle() -> crate::tsk::CustomFormatArchive {
        crate::tsk::CustomFormatArchive {
            name: "custom".to_owned(),
            format_type_pre_bnc: 5,
            default_format: Box::new(format_oracle()),
            format_type: Some(6),
            ..Default::default()
        }
    }

    fn import_warning_oracle() -> tst::ImportWarningSetArchive {
        tst::ImportWarningSetArchive {
            cond_format_expr: Some(true),
            cond_format_stop_if_true: Some(false),
            ..Default::default()
        }
    }

    fn cell_spec_oracle() -> tst::CellSpecArchive {
        tst::CellSpecArchive {
            interaction_type: 3,
            chooser_control_start_w_first: Some(false),
            ..Default::default()
        }
    }

    fn list_entry_oracle(key: u32, populated: bool) -> tst::table_data_list::ListEntry {
        if populated {
            tst::table_data_list::ListEntry {
                key,
                refcount: 12,
                string: Some("雪 value".to_owned()),
                reference: Some(prost_reference(101, Some(-7), Some(false))),
                formula: Some(formula_oracle()),
                format: Some(format_oracle()),
                custom_format: Some(custom_format_oracle()),
                rich_text_payload: Some(prost_reference(102, None, None)),
                comment_storage: Some(prost_reference(103, Some(4), None)),
                import_warning_set: Some(import_warning_oracle()),
                cell_spec: Some(cell_spec_oracle()),
            }
        } else {
            tst::table_data_list::ListEntry {
                key,
                refcount: 0,
                ..Default::default()
            }
        }
    }

    fn list_oracle() -> tst::TableDataList {
        tst::TableDataList {
            list_type: tst::table_data_list::ListType::Format as i32,
            next_list_id: 77,
            entries: vec![list_entry_oracle(19, true), list_entry_oracle(3, false)],
            segments: vec![
                prost_reference(201, Some(9), Some(false)),
                prost_reference(202, None, None),
            ],
            is_new_for_bnc: Some(false),
        }
    }

    fn segment_oracle() -> tst::TableDataListSegment {
        tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::Formula as i32,
            key_range: crate::tsp::Range {
                location: 0x1234,
                length: 0x5678,
            },
            entries: vec![list_entry_oracle(41, true), list_entry_oracle(2, false)],
        }
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct ReferenceFact {
        identifier: u64,
        deprecated_type: Option<i32>,
        deprecated_is_external: Option<bool>,
    }

    fn strict_reference_fact(reference: ReferenceSnapshot) -> ReferenceFact {
        ReferenceFact {
            identifier: reference.identifier(),
            deprecated_type: reference.deprecated_type(),
            deprecated_is_external: reference.deprecated_is_external(),
        }
    }

    fn prost_reference_fact(reference: &crate::tsp::Reference) -> ReferenceFact {
        ReferenceFact {
            identifier: reference.identifier,
            deprecated_type: reference.deprecated_type,
            deprecated_is_external: reference.deprecated_is_external,
        }
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct ListEntryFact {
        key: u32,
        ref_count: u32,
        string_value: Option<String>,
        reference: Option<ReferenceFact>,
        formula: Option<Vec<u8>>,
        format: Option<Vec<u8>>,
        custom_format: Option<Vec<u8>>,
        rich_text_payload: Option<ReferenceFact>,
        comment_storage: Option<ReferenceFact>,
        import_warning_set: Option<Vec<u8>>,
        cell_spec: Option<Vec<u8>>,
    }

    fn strict_entry_fact(entry: TableDataListEntrySnapshot<'_>) -> ListEntryFact {
        ListEntryFact {
            key: entry.key(),
            ref_count: entry.ref_count(),
            string_value: entry.string_value().map(str::to_owned),
            reference: entry.reference().map(strict_reference_fact),
            formula: entry.formula().map(<[u8]>::to_vec),
            format: entry.format().map(<[u8]>::to_vec),
            custom_format: entry.custom_format().map(<[u8]>::to_vec),
            rich_text_payload: entry.rich_text_payload().map(strict_reference_fact),
            comment_storage: entry.comment_storage().map(strict_reference_fact),
            import_warning_set: entry.import_warning_set().map(<[u8]>::to_vec),
            cell_spec: entry.cell_spec().map(<[u8]>::to_vec),
        }
    }

    fn prost_entry_fact(entry: &tst::table_data_list::ListEntry) -> ListEntryFact {
        ListEntryFact {
            key: entry.key,
            ref_count: entry.refcount,
            string_value: entry.string.clone(),
            reference: entry.reference.as_ref().map(prost_reference_fact),
            formula: entry.formula.as_ref().map(|value| value.encode_to_vec()),
            format: entry.format.as_ref().map(|value| value.encode_to_vec()),
            custom_format: entry
                .custom_format
                .as_ref()
                .map(|value| value.encode_to_vec()),
            rich_text_payload: entry.rich_text_payload.as_ref().map(prost_reference_fact),
            comment_storage: entry.comment_storage.as_ref().map(prost_reference_fact),
            import_warning_set: entry
                .import_warning_set
                .as_ref()
                .map(|value| value.encode_to_vec()),
            cell_spec: entry.cell_spec.as_ref().map(|value| value.encode_to_vec()),
        }
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct SegmentFact {
        raw: Vec<u8>,
        reference: ReferenceFact,
    }

    #[derive(Default)]
    struct ListCollector {
        entries: Vec<ListEntryFact>,
        segments: Vec<SegmentFact>,
        source_range: Option<(usize, usize)>,
        borrowed_payloads: Vec<(usize, usize)>,
    }

    impl ListCollector {
        fn for_source(source: &[u8]) -> Self {
            let start = source.as_ptr() as usize;
            Self {
                source_range: Some((start, start.saturating_add(source.len()))),
                ..Default::default()
            }
        }

        fn assert_borrowed(&mut self, payload: &[u8]) {
            let Some((source_start, source_end)) = self.source_range else {
                return;
            };
            if payload.is_empty() {
                return;
            }
            let payload_start = payload.as_ptr() as usize;
            let payload_end = payload_start.saturating_add(payload.len());
            assert!(payload_start >= source_start);
            assert!(payload_end <= source_end);
            self.borrowed_payloads.push((payload_start, payload.len()));
        }
    }

    impl StorageVisitor for ListCollector {
        fn visit_list_entry(
            &mut self,
            entry: TableDataListEntrySnapshot<'_>,
        ) -> Result<(), DecodeError> {
            if let Some(string) = entry.string_value() {
                self.assert_borrowed(string.as_bytes());
            }
            for payload in [
                entry.formula(),
                entry.format(),
                entry.custom_format(),
                entry.import_warning_set(),
                entry.cell_spec(),
            ]
            .into_iter()
            .flatten()
            {
                self.assert_borrowed(payload);
            }
            self.entries.push(strict_entry_fact(entry));
            Ok(())
        }

        fn visit_list_segment(
            &mut self,
            reference: ReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.assert_borrowed(reference.raw());
            self.segments.push(SegmentFact {
                raw: reference.raw().to_vec(),
                reference: strict_reference_fact(reference.reference()),
            });
            Ok(())
        }
    }

    fn list_minimal() -> Vec<u8> {
        let mut source = Vec::new();
        v(&mut source, 1, 1);
        v(&mut source, 2, 2);
        source
    }

    fn entry_minimal() -> Vec<u8> {
        let mut source = Vec::new();
        v(&mut source, 1, 1);
        v(&mut source, 2, 2);
        source
    }

    fn range_minimal() -> Vec<u8> {
        let mut source = Vec::new();
        v(&mut source, 1, 4);
        v(&mut source, 2, 8);
        source
    }

    fn segment_minimal() -> Vec<u8> {
        let mut source = Vec::new();
        v(&mut source, 1, 1);
        b(&mut source, 2, &range_minimal());
        source
    }

    fn assert_invalid_list(source: &[u8], label: &str) {
        assert!(
            decode_table_data_list(source, options(source)).is_err(),
            "{label} unexpectedly decoded"
        );
    }

    fn assert_invalid_entry(source: &[u8], label: &str) {
        assert!(
            decode_table_data_list_entry(source, options(source)).is_err(),
            "{label} unexpectedly decoded"
        );
    }

    fn assert_invalid_segment(source: &[u8], label: &str) {
        assert!(
            decode_table_data_list_segment(source, options(source)).is_err(),
            "{label} unexpectedly decoded"
        );
    }

    #[test]
    fn list_root_entry_and_segment_match_prost_presence_order_borrowing_and_reports() {
        let expected = list_oracle();
        let source = expected.encode_to_vec();
        let before = source.clone();
        let prost = tst::TableDataList::decode(source.as_slice()).unwrap();
        let expected_entries = prost
            .entries
            .iter()
            .map(prost_entry_fact)
            .collect::<Vec<_>>();
        let expected_segments = prost
            .segments
            .iter()
            .map(|reference| SegmentFact {
                raw: reference.encode_to_vec(),
                reference: prost_reference_fact(reference),
            })
            .collect::<Vec<_>>();

        let mut visitor = ListCollector::for_source(&source);
        let (visited, visitor_report) =
            decode_table_data_list_with_visitor(&source, options(&source), &mut visitor).unwrap();
        let (scalar, scalar_report) =
            decode_table_data_list_with_report(&source, options(&source)).unwrap();

        assert_eq!(source, before);
        assert_eq!(visited, scalar);
        assert_eq!(visitor_report, scalar_report);
        assert_eq!(visited.list_type(), prost.list_type);
        assert_eq!(visited.next_list_id(), prost.next_list_id);
        assert_eq!(visited.is_new_for_bnc(), prost.is_new_for_bnc);
        assert_eq!(visitor.entries, expected_entries);
        assert_eq!(visitor.segments, expected_segments);
        assert_eq!(visitor.entries[0].key, 19);
        assert_eq!(visitor.entries[1].key, 3);
        assert_eq!(visitor.segments[0].reference.identifier, 201);
        assert_eq!(visitor.segments[1].reference.identifier, 202);
        assert!(visitor.borrowed_payloads.len() >= 8);

        let absent_source = list_minimal();
        let absent_prost = tst::TableDataList::decode(absent_source.as_slice()).unwrap();
        let absent = decode_table_data_list(&absent_source, options(&absent_source)).unwrap();
        assert_eq!(absent.is_new_for_bnc(), None);
        assert_eq!(absent_prost.is_new_for_bnc, None);
    }

    #[test]
    fn list_type_routes_match_prost_and_skip_repeated_payloads() {
        let mut root = list_oracle().encode_to_vec();
        // The dispatch projection must not descend into repeated entry
        // payloads; the full list decoder remains responsible for this
        // malformed child once the candidate is selected.
        b(&mut root, 3, &[0x80]);
        unknown_fields(&mut root, 90);
        let before = root.clone();
        let (root_probe, root_report) =
            decode_table_data_list_type_with_report(&root, options(&root)).unwrap();
        assert_eq!(root, before);
        assert_eq!(
            root_probe.list_type(),
            tst::table_data_list::ListType::Format as i32
        );
        assert!(root_report.fields() > 0);
        assert!(root_report.work_bytes() >= root.len().saturating_mul(2));
        assert!(decode_table_data_list(&root, options(&root)).is_err());

        let mut segment = segment_oracle().encode_to_vec();
        b(&mut segment, 3, &[0x80]);
        unknown_fields(&mut segment, 120);
        let before = segment.clone();
        let (segment_probe, segment_report) =
            decode_table_data_list_segment_type_with_report(&segment, options(&segment)).unwrap();
        assert_eq!(segment, before);
        assert_eq!(
            segment_probe.list_type(),
            tst::table_data_list::ListType::Formula as i32
        );
        assert!(segment_report.fields() > 0);
        assert!(segment_report.work_bytes() >= segment.len().saturating_mul(2));
        assert!(decode_table_data_list_segment(&segment, options(&segment)).is_err());
    }

    #[test]
    fn list_type_routes_reject_missing_scalars_duplicate_knowns_and_noncanonical_wire() {
        let mut missing_root_scalar = Vec::new();
        v(&mut missing_root_scalar, 1, 1);
        assert!(
            decode_table_data_list_type_with_report(
                &missing_root_scalar,
                options(&missing_root_scalar)
            )
            .is_err()
        );

        let mut duplicate_root_type = list_minimal();
        v(&mut duplicate_root_type, 1, 2);
        assert!(
            decode_table_data_list_type_with_report(
                &duplicate_root_type,
                options(&duplicate_root_type)
            )
            .is_err()
        );

        let mut noncanonical_root_type = Vec::new();
        key(&mut noncanonical_root_type, 1, 0);
        noncanonical_root_type.extend_from_slice(&[0x81, 0x00]);
        v(&mut noncanonical_root_type, 2, 2);
        assert!(
            decode_table_data_list_type_with_report(
                &noncanonical_root_type,
                options(&noncanonical_root_type)
            )
            .is_err()
        );

        let mut missing_segment_range = Vec::new();
        v(&mut missing_segment_range, 1, 1);
        assert!(
            decode_table_data_list_segment_type_with_report(
                &missing_segment_range,
                options(&missing_segment_range)
            )
            .is_err()
        );

        let mut duplicate_segment_type = segment_minimal();
        v(&mut duplicate_segment_type, 1, 2);
        assert!(
            decode_table_data_list_segment_type_with_report(
                &duplicate_segment_type,
                options(&duplicate_segment_type)
            )
            .is_err()
        );
    }

    #[test]
    fn list_entry_all_fields_match_prost_wire_payloads_and_presence() {
        let expected = list_entry_oracle(29, true);
        let source = expected.encode_to_vec();
        let before = source.clone();
        let prost = tst::table_data_list::ListEntry::decode(source.as_slice()).unwrap();
        let strict = decode_table_data_list_entry(&source, options(&source)).unwrap();

        assert_eq!(source, before);
        assert_eq!(strict_entry_fact(strict), prost_entry_fact(&prost));
        assert_eq!(strict.key(), prost.key);
        assert_eq!(strict.ref_count(), prost.refcount);
        assert_eq!(strict.string_value(), prost.string.as_deref());
        assert_eq!(
            strict.reference().map(strict_reference_fact),
            prost.reference.as_ref().map(prost_reference_fact)
        );
        assert_eq!(
            strict.rich_text_payload().map(strict_reference_fact),
            prost.rich_text_payload.as_ref().map(prost_reference_fact)
        );
        assert_eq!(
            strict.comment_storage().map(strict_reference_fact),
            prost.comment_storage.as_ref().map(prost_reference_fact)
        );
        assert!(strict.formula().is_some() && prost.formula.is_some());
        assert!(strict.format().is_some() && prost.format.is_some());
        assert!(strict.custom_format().is_some() && prost.custom_format.is_some());
        assert!(strict.import_warning_set().is_some() && prost.import_warning_set.is_some());
        assert!(strict.cell_spec().is_some() && prost.cell_spec.is_some());

        let sparse = list_entry_oracle(30, false).encode_to_vec();
        let sparse_prost = tst::table_data_list::ListEntry::decode(sparse.as_slice()).unwrap();
        let sparse_strict = decode_table_data_list_entry(&sparse, options(&sparse)).unwrap();
        assert_eq!(
            strict_entry_fact(sparse_strict),
            prost_entry_fact(&sparse_prost)
        );
        assert_eq!(sparse_strict.string_value(), None);
        assert_eq!(sparse_strict.reference(), None);
        assert_eq!(sparse_strict.formula(), None);
        assert_eq!(sparse_strict.format(), None);
        assert_eq!(sparse_strict.custom_format(), None);
        assert_eq!(sparse_strict.rich_text_payload(), None);
        assert_eq!(sparse_strict.comment_storage(), None);
        assert_eq!(sparse_strict.import_warning_set(), None);
        assert_eq!(sparse_strict.cell_spec(), None);
    }

    #[test]
    fn list_segment_matches_prost_range_entry_order_borrowing_and_reports() {
        let expected = segment_oracle();
        let source = expected.encode_to_vec();
        let before = source.clone();
        let prost = tst::TableDataListSegment::decode(source.as_slice()).unwrap();
        let mut visitor = ListCollector::for_source(&source);
        let (visited, visitor_report) =
            decode_table_data_list_segment_with_visitor(&source, options(&source), &mut visitor)
                .unwrap();
        let (scalar, scalar_report) =
            decode_table_data_list_segment_with_report(&source, options(&source)).unwrap();

        assert_eq!(source, before);
        assert_eq!(visited, scalar);
        assert_eq!(visitor_report, scalar_report);
        assert_eq!(visited.list_type(), prost.list_type);
        assert_eq!(visited.key_range_location(), prost.key_range.location);
        assert_eq!(visited.key_range_length(), prost.key_range.length);
        assert_eq!(
            visitor.entries,
            prost
                .entries
                .iter()
                .map(prost_entry_fact)
                .collect::<Vec<_>>()
        );
        assert_eq!(visitor.entries[0].key, 41);
        assert_eq!(visitor.entries[1].key, 2);
        let range_bytes = prost.key_range.encode_to_vec();
        assert_eq!(visited.key_range(), range_bytes.as_slice());
        assert!(visitor.borrowed_payloads.len() >= 6);
    }

    #[test]
    fn list_routes_accept_unknown_scalar_bytes_fixed_and_matched_groups_without_mutation() {
        let expected = list_oracle();
        let mut root = expected.encode_to_vec();
        unknown_fields(&mut root, 90);
        let root_before = root.clone();
        let mut root_visitor = ListCollector::for_source(&root);
        let root_prost = tst::TableDataList::decode(root.as_slice()).unwrap();
        let (root_snapshot, root_report) =
            decode_table_data_list_with_visitor(&root, options(&root), &mut root_visitor).unwrap();
        assert_eq!(root, root_before);
        assert_eq!(root_snapshot.list_type(), root_prost.list_type);
        assert_eq!(root_snapshot.next_list_id(), root_prost.next_list_id);
        assert_eq!(root_snapshot.is_new_for_bnc(), root_prost.is_new_for_bnc);
        assert_eq!(root_visitor.entries.len(), root_prost.entries.len());
        assert_eq!(root_visitor.segments.len(), root_prost.segments.len());
        assert!(root_report.fields() > 0);

        let mut entry = list_entry_oracle(31, true).encode_to_vec();
        unknown_fields(&mut entry, 120);
        let entry_before = entry.clone();
        let entry_prost = tst::table_data_list::ListEntry::decode(entry.as_slice()).unwrap();
        let entry_strict = decode_table_data_list_entry(&entry, options(&entry)).unwrap();
        assert_eq!(entry, entry_before);
        assert_eq!(
            strict_entry_fact(entry_strict),
            prost_entry_fact(&entry_prost)
        );

        let mut segment = segment_oracle().encode_to_vec();
        unknown_fields(&mut segment, 150);
        let segment_before = segment.clone();
        let segment_prost = tst::TableDataListSegment::decode(segment.as_slice()).unwrap();
        let segment_strict = decode_table_data_list_segment(&segment, options(&segment)).unwrap();
        assert_eq!(segment, segment_before);
        assert_eq!(segment_strict.list_type(), segment_prost.list_type);
        assert_eq!(
            segment_strict.key_range_location(),
            segment_prost.key_range.location
        );
        assert_eq!(
            segment_strict.key_range_length(),
            segment_prost.key_range.length
        );
    }

    #[test]
    fn list_root_malformed_matrix_rejects_wire_varint_presence_reference_truncation_and_groups() {
        let valid = list_minimal();
        let mut cases = Vec::<(&str, Vec<u8>)>::new();

        let mut missing_list_type = Vec::new();
        v(&mut missing_list_type, 2, 2);
        cases.push(("missing required list type", missing_list_type));
        let mut missing_next_list_id = Vec::new();
        v(&mut missing_next_list_id, 1, 1);
        cases.push(("missing required next list id", missing_next_list_id));
        let mut duplicate_list_type = valid.clone();
        v(&mut duplicate_list_type, 1, 2);
        cases.push(("duplicate required list type", duplicate_list_type));
        let mut duplicate_next_list_id = valid.clone();
        v(&mut duplicate_next_list_id, 2, 3);
        cases.push(("duplicate required next list id", duplicate_next_list_id));
        let mut duplicate_optional = valid.clone();
        v(&mut duplicate_optional, 5, 0);
        v(&mut duplicate_optional, 5, 1);
        cases.push(("duplicate optional is_new", duplicate_optional));
        let mut wrong_list_type_wire = Vec::new();
        b(&mut wrong_list_type_wire, 1, &[]);
        v(&mut wrong_list_type_wire, 2, 2);
        cases.push(("wrong list type wire", wrong_list_type_wire));
        let mut wrong_next_id_wire = Vec::new();
        v(&mut wrong_next_id_wire, 1, 1);
        b(&mut wrong_next_id_wire, 2, &[]);
        cases.push(("wrong next list id wire", wrong_next_id_wire));
        let mut wrong_bool_wire = valid.clone();
        b(&mut wrong_bool_wire, 5, &[]);
        cases.push(("wrong is_new wire", wrong_bool_wire));
        let mut noncanonical = Vec::new();
        key(&mut noncanonical, 1, 0);
        noncanonical.extend_from_slice(&[0x81, 0x00]);
        v(&mut noncanonical, 2, 2);
        cases.push(("noncanonical list type varint", noncanonical));
        let mut overflowing = Vec::new();
        v(&mut overflowing, 1, 1);
        v(&mut overflowing, 2, u64::from(u32::MAX) + 1);
        cases.push(("overflowing next list id", overflowing));
        let mut invalid_bool = valid.clone();
        v(&mut invalid_bool, 5, 2);
        cases.push(("invalid is_new bool", invalid_bool));
        let mut zero_reference = valid.clone();
        b(&mut zero_reference, 4, &reference(0));
        cases.push(("zero segment reference", zero_reference));
        let mut external = valid.clone();
        b(&mut external, 4, &external_reference(301));
        cases.push(("external segment reference", external));
        let mut malformed_reference = valid.clone();
        b(&mut malformed_reference, 4, &[0x80]);
        cases.push(("truncated segment reference", malformed_reference));
        let mut truncated = valid.clone();
        truncated.pop();
        cases.push(("truncated root", truncated));
        let mut unclosed_group = valid.clone();
        key(&mut unclosed_group, 90, 3);
        v(&mut unclosed_group, 91, 1);
        cases.push(("unclosed root group", unclosed_group));
        let mut mismatched_group = valid.clone();
        key(&mut mismatched_group, 90, 3);
        key(&mut mismatched_group, 91, 4);
        cases.push(("mismatched root group", mismatched_group));
        let mut stray_end_group = valid.clone();
        key(&mut stray_end_group, 90, 4);
        cases.push(("stray root end group", stray_end_group));

        for (label, source) in cases {
            assert_invalid_list(&source, label);
        }

        let mut deep = valid;
        unknown_groups_to_depth(&mut deep, 100, 6);
        let error = decode_table_data_list(
            &deep,
            DecodeOptions::new(
                deep.len(),
                usize::MAX,
                usize::MAX,
                3,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn list_entry_malformed_matrix_rejects_required_optional_and_nested_payload_failures() {
        let mut cases = Vec::<(&str, Vec<u8>)>::new();
        cases.push(("missing required key", {
            let mut source = Vec::new();
            v(&mut source, 2, 2);
            source
        }));
        cases.push(("missing required refcount", {
            let mut source = Vec::new();
            v(&mut source, 1, 1);
            source
        }));
        let valid = entry_minimal();
        let mut duplicate_key = valid.clone();
        v(&mut duplicate_key, 1, 2);
        cases.push(("duplicate required key", duplicate_key));
        let mut duplicate_refcount = valid.clone();
        v(&mut duplicate_refcount, 2, 3);
        cases.push(("duplicate required refcount", duplicate_refcount));

        let populated = list_entry_oracle(17, true).encode_to_vec();
        for (label, number, payload) in [
            ("string", 3, b"second".to_vec()),
            ("reference", 4, reference(401)),
            ("formula", 5, formula_oracle().encode_to_vec()),
            ("format", 6, format_oracle().encode_to_vec()),
            ("custom format", 8, custom_format_oracle().encode_to_vec()),
            ("rich text payload", 9, reference(402)),
            ("comment storage", 10, reference(403)),
            (
                "import warning set",
                11,
                import_warning_oracle().encode_to_vec(),
            ),
            ("cell spec", 12, cell_spec_oracle().encode_to_vec()),
        ] {
            let mut duplicate = populated.clone();
            b(&mut duplicate, number, &payload);
            cases.push((
                match label {
                    "string" => "duplicate optional string",
                    "reference" => "duplicate optional reference",
                    "formula" => "duplicate optional formula",
                    "format" => "duplicate optional format",
                    "custom format" => "duplicate optional custom format",
                    "rich text payload" => "duplicate optional rich text payload",
                    "comment storage" => "duplicate optional comment storage",
                    "import warning set" => "duplicate optional import warning set",
                    "cell spec" => "duplicate optional cell spec",
                    _ => unreachable!(),
                },
                duplicate,
            ));
        }

        for number in [1_u32, 2] {
            let mut wrong_wire = Vec::new();
            b(&mut wrong_wire, number, &[]);
            if number == 1 {
                v(&mut wrong_wire, 2, 2);
            } else {
                v(&mut wrong_wire, 1, 1);
            }
            cases.push(("wrong required field wire", wrong_wire));
        }
        for number in [3_u32, 4, 5, 6, 8, 9, 10, 11, 12] {
            let mut wrong_wire = entry_minimal();
            v(&mut wrong_wire, number, 1);
            cases.push(("wrong optional bytes wire", wrong_wire));
        }

        let mut noncanonical_key = Vec::new();
        key(&mut noncanonical_key, 1, 0);
        noncanonical_key.extend_from_slice(&[0x81, 0x00]);
        v(&mut noncanonical_key, 2, 2);
        cases.push(("noncanonical key varint", noncanonical_key));
        let mut overflowing_refcount = Vec::new();
        v(&mut overflowing_refcount, 1, 1);
        v(&mut overflowing_refcount, 2, u64::from(u32::MAX) + 1);
        cases.push(("overflowing refcount", overflowing_refcount));
        let mut invalid_utf8 = entry_minimal();
        b(&mut invalid_utf8, 3, &[0xff]);
        cases.push(("invalid UTF-8 string", invalid_utf8));

        for number in [4_u32, 9, 10] {
            let mut zero = entry_minimal();
            b(&mut zero, number, &reference(0));
            cases.push(("zero selected reference", zero));
            let mut external = entry_minimal();
            b(&mut external, number, &external_reference(501));
            cases.push(("external selected reference", external));
            let mut malformed = entry_minimal();
            b(&mut malformed, number, &[0x80]);
            cases.push(("malformed selected reference", malformed));
        }
        for number in [5_u32, 6, 8, 11, 12] {
            let mut malformed = entry_minimal();
            b(&mut malformed, number, &[0x80]);
            cases.push(("malformed opaque payload", malformed));
            let mut wrong_wire = entry_minimal();
            b(&mut wrong_wire, number, &[0x0f]);
            cases.push(("invalid opaque payload wire", wrong_wire));
        }
        let mut truncated = populated.clone();
        truncated.pop();
        cases.push(("truncated entry", truncated));
        let mut unclosed_group = populated.clone();
        key(&mut unclosed_group, 90, 3);
        cases.push(("unclosed entry group", unclosed_group));
        let mut mismatched_group = populated.clone();
        key(&mut mismatched_group, 90, 3);
        key(&mut mismatched_group, 91, 4);
        cases.push(("mismatched entry group", mismatched_group));

        for (label, source) in cases {
            assert_invalid_entry(&source, label);
        }

        let mut deep = entry_minimal();
        unknown_groups_to_depth(&mut deep, 100, 6);
        let error = decode_table_data_list_entry(
            &deep,
            DecodeOptions::new(
                deep.len(),
                usize::MAX,
                usize::MAX,
                3,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn list_segment_and_range_malformed_matrix_rejects_required_wire_varint_truncation_and_groups()
    {
        let valid = segment_minimal();
        let mut cases = Vec::<(&str, Vec<u8>)>::new();
        let mut missing_list_type = Vec::new();
        b(&mut missing_list_type, 2, &range_minimal());
        cases.push(("missing segment list type", missing_list_type));
        let mut missing_range = Vec::new();
        v(&mut missing_range, 1, 1);
        cases.push(("missing segment range", missing_range));
        let mut duplicate_list_type = valid.clone();
        v(&mut duplicate_list_type, 1, 2);
        cases.push(("duplicate segment list type", duplicate_list_type));
        let mut duplicate_range = valid.clone();
        b(&mut duplicate_range, 2, &range_minimal());
        cases.push(("duplicate segment range", duplicate_range));
        let mut wrong_list_type_wire = Vec::new();
        b(&mut wrong_list_type_wire, 1, &[]);
        b(&mut wrong_list_type_wire, 2, &range_minimal());
        cases.push(("wrong segment list type wire", wrong_list_type_wire));
        let mut wrong_range_wire = Vec::new();
        v(&mut wrong_range_wire, 1, 1);
        v(&mut wrong_range_wire, 2, 1);
        cases.push(("wrong segment range wire", wrong_range_wire));
        let mut malformed_range = Vec::new();
        v(&mut malformed_range, 1, 1);
        let mut range = Vec::new();
        v(&mut range, 1, 4);
        b(&mut malformed_range, 2, &range);
        cases.push(("missing range length", malformed_range));
        let mut duplicate_range_location = Vec::new();
        v(&mut duplicate_range_location, 1, 1);
        let mut range = Vec::new();
        v(&mut range, 1, 4);
        v(&mut range, 1, 5);
        v(&mut range, 2, 8);
        b(&mut duplicate_range_location, 2, &range);
        cases.push(("duplicate range location", duplicate_range_location));
        let mut duplicate_range_length = Vec::new();
        v(&mut duplicate_range_length, 1, 1);
        let mut range = Vec::new();
        v(&mut range, 1, 4);
        v(&mut range, 2, 8);
        v(&mut range, 2, 9);
        b(&mut duplicate_range_length, 2, &range);
        cases.push(("duplicate range length", duplicate_range_length));
        let mut wrong_range_location_wire = Vec::new();
        v(&mut wrong_range_location_wire, 1, 1);
        let mut range = Vec::new();
        b(&mut range, 1, &[]);
        v(&mut range, 2, 8);
        b(&mut wrong_range_location_wire, 2, &range);
        cases.push(("wrong range location wire", wrong_range_location_wire));
        let mut noncanonical_range = Vec::new();
        v(&mut noncanonical_range, 1, 1);
        let mut range = Vec::new();
        key(&mut range, 1, 0);
        range.extend_from_slice(&[0x84, 0x00]);
        v(&mut range, 2, 8);
        b(&mut noncanonical_range, 2, &range);
        cases.push(("noncanonical range location", noncanonical_range));
        let mut overflowing_range = Vec::new();
        v(&mut overflowing_range, 1, 1);
        let mut range = Vec::new();
        v(&mut range, 1, u64::from(u32::MAX) + 1);
        v(&mut range, 2, 8);
        b(&mut overflowing_range, 2, &range);
        cases.push(("overflowing range location", overflowing_range));
        let mut truncated = valid.clone();
        truncated.pop();
        cases.push(("truncated segment", truncated));
        let mut truncated_range = Vec::new();
        v(&mut truncated_range, 1, 1);
        b(&mut truncated_range, 2, &[0x0a]);
        cases.push(("truncated range", truncated_range));
        let mut unclosed_group = valid.clone();
        key(&mut unclosed_group, 90, 3);
        cases.push(("unclosed segment group", unclosed_group));
        let mut mismatched_group = valid.clone();
        key(&mut mismatched_group, 90, 3);
        key(&mut mismatched_group, 91, 4);
        cases.push(("mismatched segment group", mismatched_group));

        for (label, source) in cases {
            assert_invalid_segment(&source, label);
        }

        let mut deep = valid;
        unknown_groups_to_depth(&mut deep, 100, 6);
        let error = decode_table_data_list_segment(
            &deep,
            DecodeOptions::new(
                deep.len(),
                usize::MAX,
                usize::MAX,
                3,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn list_exact_limits_are_inclusive_and_each_max_minus_one_is_typed() {
        let source = list_oracle().encode_to_vec();
        let (_, report) = decode_table_data_list_with_report(&source, options(&source)).unwrap();
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
            report.text_bytes(),
        );
        let (_, exact_report) = decode_table_data_list_with_report(&source, exact).unwrap();
        assert_eq!(exact_report, report);

        let error = decode_table_data_list(
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
            error.resource_limit(),
            Some(DecodeLimit::Bytes { .. })
        ));
        let error = decode_table_data_list(
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
            error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let error = decode_table_data_list(
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
            error.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let error = decode_table_data_list(
            &source,
            DecodeOptions::new(
                source.len(),
                usize::MAX,
                usize::MAX,
                64,
                report.references() - 1,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::References { .. })
        ));
        let error = decode_table_data_list(
            &source,
            DecodeOptions::new(
                source.len(),
                usize::MAX,
                usize::MAX,
                64,
                usize::MAX,
                report.text_bytes() - 1,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Text { .. })
        ));
        let error = decode_table_data_list(
            &source,
            DecodeOptions::new(
                source.len(),
                usize::MAX,
                usize::MAX,
                report.max_depth() - 1,
                usize::MAX,
                usize::MAX,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn list_callbacks_are_observable_before_a_later_entry_error_and_source_stays_immutable() {
        let first = list_entry_oracle(61, false).encode_to_vec();
        let mut malformed_later = entry_minimal();
        v(&mut malformed_later, 1, 62);
        let mut source = list_minimal();
        b(&mut source, 3, &first);
        b(&mut source, 3, &malformed_later);
        let before = source.clone();

        // StorageVisitor is intentionally streaming: a callback for the first
        // entry can run before strict validation discovers a later duplicate.
        // The codec cannot roll that callback back; callers must stage effects.
        let mut visitor = ListCollector::for_source(&source);
        assert!(
            decode_table_data_list_with_visitor(&source, options(&source), &mut visitor).is_err()
        );
        assert_eq!(visitor.entries.len(), 1);
        assert_eq!(visitor.entries[0].key, 61);
        assert_eq!(source, before);
    }
}
