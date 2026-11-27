//! Stateful semantic models used while decoding pivot parts.

use super::super::PivotValueFunction;
use super::super::cache::{CacheRecord, Definition, Field, Records};

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum TableContext {
    Root,
    Location,
    PivotFields,
    PivotField,
    RowFields,
    ColFields,
    PageFields,
    DataFields,
    Other,
}

pub(super) struct RawDataField {
    pub(super) field_index: u32,
    pub(super) function: PivotValueFunction,
    pub(super) display_name: Option<String>,
}

pub(super) struct PivotTableParser {
    pub(super) name: String,
    pub(super) cache_id: u32,
    pub(super) sheet_name: String,
    pub(super) location_ref: String,
    pub(super) field_names: Vec<String>,
    pub(super) row_indexes: Vec<u32>,
    pub(super) column_indexes: Vec<u32>,
    pub(super) row_field_count: usize,
    pub(super) column_field_count: usize,
    pub(super) filter_indexes: Vec<u32>,
    pub(super) data_fields: Vec<RawDataField>,
    pub(super) expected_pivot_fields: Option<u32>,
    pub(super) expected_row_fields: Option<u32>,
    pub(super) expected_col_fields: Option<u32>,
    pub(super) expected_page_fields: Option<u32>,
    pub(super) expected_data_fields: Option<u32>,
    pub(super) saw_location: bool,
    pub(super) saw_pivot_fields: bool,
    pub(super) saw_row_fields: bool,
    pub(super) saw_col_fields: bool,
    pub(super) saw_page_fields: bool,
    pub(super) saw_data_fields: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum CacheContext {
    Root,
    CacheSource,
    WorksheetSource,
    CacheFields,
    CacheField,
    Items,
    Item,
    Other,
}

pub(super) struct PivotCacheParser {
    pub(super) cache: Definition,
    pub(super) pending_field: Option<Field>,
    pub(super) expected_fields: Option<u32>,
    pub(super) expected_shared_items: Option<u32>,
    pub(super) saw_cache_source: bool,
    pub(super) saw_worksheet_source: bool,
    pub(super) saw_cache_fields: bool,
    pub(super) field_saw_shared_items: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum CacheRecordsContext {
    Root,
    Record,
    Item,
    Other,
}

pub(super) struct RecordsParser {
    pub(super) records: Records,
    pub(super) pending_record: Option<CacheRecord>,
    pub(super) expected_records: u32,
    pub(super) actual_records: usize,
    pub(super) pending_value_count: usize,
    pub(super) expected_field_count: Option<usize>,
    pub(super) shared_item_counts: Vec<usize>,
    pub(super) retain_records: bool,
}
