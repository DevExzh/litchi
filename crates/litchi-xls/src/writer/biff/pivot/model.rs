//! Typed configurations for the BIFF PivotTable record writer.

/// Configuration for an SXVIEW record.
pub(crate) struct SxViewConfig<'a> {
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
    pub first_header_row: u16,
    pub first_data_row: u16,
    pub first_data_col: u16,
    /// Pivot cache index (0-based).
    pub cache_index: u16,
    /// Axis for the data field header (0=none, 1=row, 2=col, 4=page).
    pub data_axis: u16,
    /// Position of the data field label within the axis (-1 = end).
    pub data_position: u16,
    /// Total number of fields.
    pub field_count: u16,
    pub row_field_count: u16,
    pub col_field_count: u16,
    pub page_field_count: u16,
    pub data_field_count: u16,
    pub data_row_count: u16,
    pub data_col_count: u16,
    /// Option flags: bit 0 = fRwGrand, bit 1 = fColGrand, bit 3 = fAutoFormat.
    pub flags: u16,
    /// Auto-format index (0 = none).
    pub auto_format_index: u16,
    pub name: &'a str,
    pub data_field_name: &'a str,
}

/// Configuration for an SXVD record.
pub(crate) struct SxVdConfig<'a> {
    /// Axis: 0=none, 1=row, 2=col, 4=page, 8=data.
    pub axis: u16,
    pub subtotal_count: u16,
    pub subtotal_flags: u16,
    pub item_count: u16,
    /// Optional field name override (`None` means not present).
    pub name: Option<&'a str>,
}

/// Configuration for an SXVI record.
pub(crate) struct SxViConfig<'a> {
    /// Item type: 0x0000=Data, 0x0001=Default subtotal, 0x0002=Sum, etc.
    pub item_type: u16,
    pub flags: u16,
    pub cache_index: u16,
    pub name: Option<&'a str>,
}

/// Configuration for an SXDI record.
pub(crate) struct SxDiConfig<'a> {
    pub source_field_index: u16,
    /// Aggregation function: 0=Sum,1=Count,2=Average,3=Max,4=Min,5=Product,...
    pub function: u16,
    pub display_format: u16,
    pub base_field_index: u16,
    pub base_item_index: u16,
    pub num_format_index: u16,
    pub name: &'a str,
}

/// Configuration for an SXEX record.
pub(crate) struct SxExConfig {
    pub sx_format_count: u16,
    pub sx_select_count: u16,
    pub page_rows: u16,
    pub page_cols: u16,
    /// XclPTExtInfo flags (default `0x004F0200`).
    pub flags: u32,
}

impl Default for SxExConfig {
    fn default() -> Self {
        Self {
            sx_format_count: 0,
            sx_select_count: 0,
            page_rows: 0,
            page_cols: 0,
            flags: 0x004F_0200,
        }
    }
}

/// Configuration for an SXDB pivot-cache definition record.
pub(super) struct SxDbConfig {
    pub record_count: u32,
    pub stream_id: u16,
    pub standard_field_count: u16,
    pub total_field_count: u16,
    pub flags: u16,
}

/// Configuration for an SXFDB pivot-cache field record.
pub(super) struct SxFdbConfig<'a> {
    pub item_count: u16,
    pub name: &'a str,
    pub has_items: bool,
    pub is_numeric: bool,
    pub data_flags: u16,
    pub use_16bit_indices: bool,
    pub group_child: Option<u16>,
    pub group_base: Option<u16>,
    pub group_item_count: u16,
    pub base_item_count: u16,
    pub original_item_count: u16,
    pub is_grouped: bool,
}

/// Per-field information used to generate a pivot-cache storage stream.
pub(crate) struct PivotCacheFieldInfo<'a> {
    pub name: &'a str,
    pub items: &'a [crate::PivotCacheItem],
    pub is_numeric: bool,
    pub unique_numeric_count: u16,
    pub grouping: Option<&'a crate::PivotCacheGrouping>,
    pub group_child: Option<u16>,
    pub is_source_field: bool,
}

/// One source data row in a pivot-cache storage stream.
pub(crate) struct PivotCacheSourceRow<'a> {
    pub item_indices: &'a [u16],
    pub numeric_values: &'a [f64],
}

/// Information needed to generate one pivot-cache storage stream.
pub(crate) struct PivotCacheStreamInfo<'a> {
    pub stream_id: u16,
    pub record_count: u32,
    pub fields: &'a [PivotCacheFieldInfo<'a>],
    pub source_rows: &'a [PivotCacheSourceRow<'a>],
}
