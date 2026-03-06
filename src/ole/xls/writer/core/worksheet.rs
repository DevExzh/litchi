use std::collections::{HashMap, HashSet};

use super::{XlsCellValue, XlsConditionalFormat, XlsDataValidation};
use crate::ole::xls::writer::biff::AutoFilterConditionWrite;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PivotCellXfRole {
    HeaderAccent,
    HeaderPlain,
    RowLabel,
    Value,
}

#[derive(Debug, Clone)]
pub(super) struct WritableCell {
    /// Row index (0-based)
    pub row: u32,
    /// Column index (0-based)
    pub col: u16,
    /// Cell value
    pub value: XlsCellValue,
    pub format_idx: u16,
    pub pivot_xf_role: Option<PivotCellXfRole>,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct MergedRange {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

/// Freeze panes configuration for a worksheet.
#[derive(Debug, Clone, Copy)]
pub(super) struct FreezePanes {
    /// Number of frozen rows from the top (0-based, inclusive index of last frozen row).
    pub freeze_rows: u32,
    /// Number of frozen columns from the left (0-based, inclusive index of last frozen column).
    pub freeze_cols: u16,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct AutoFilterRange {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct XlsSheetProtection {
    pub protect_objects: bool,
    pub protect_scenarios: bool,
    pub password_hash: Option<u16>,
}

/// Hyperlink target within a worksheet.
#[derive(Debug, Clone)]
pub(super) struct XlsHyperlink {
    /// First row (0-based) of the hyperlink range.
    pub first_row: u32,
    /// Last row (0-based) of the hyperlink range.
    pub last_row: u32,
    /// First column (0-based) of the hyperlink range.
    pub first_col: u16,
    /// Last column (0-based) of the hyperlink range.
    pub last_col: u16,
    /// Raw hyperlink target string.
    pub url: String,
}

/// Represents a worksheet in the writer
#[derive(Debug)]
pub(super) struct WritableWorksheet {
    /// Worksheet name
    pub name: String,
    /// Cells to write (indexed by (row, col))
    pub cells: HashMap<(u32, u16), WritableCell>,
    /// First used row
    pub first_row: u32,
    /// Last used row (exclusive)
    pub last_row: u32,
    /// First used column
    pub first_col: u16,
    /// Last used column (exclusive)
    pub last_col: u16,
    /// Per-column widths in 1/256 character units (BIFF8 COLINFO).
    pub column_widths: HashMap<u16, u16>,
    /// Hidden columns (0-based indices).
    pub hidden_columns: HashSet<u16>,
    /// Per-row heights in 1/20 point units (BIFF8 ROW).
    pub row_heights: HashMap<u32, u16>,
    /// Hidden rows (0-based indices).
    pub hidden_rows: HashSet<u32>,
    pub merged_ranges: Vec<MergedRange>,
    pub data_validations: Vec<XlsDataValidation>,
    pub conditional_formats: Vec<XlsConditionalFormat>,
    /// Optional freeze panes configuration.
    pub freeze_panes: Option<FreezePanes>,
    pub sheet_protection: Option<XlsSheetProtection>,
    pub auto_filter: Option<AutoFilterRange>,
    /// Cell or range hyperlinks stored for this worksheet.
    pub hyperlinks: Vec<XlsHyperlink>,
    /// Per-column AutoFilter conditions.
    pub auto_filter_columns: Vec<AutoFilterColumnDef>,
    /// Sort configuration.
    pub sort_config: Option<SortConfig>,
    /// Pivot tables to write.
    pub pivot_tables: Vec<WritablePivotTable>,
}

/// A column-level AutoFilter condition for the writer.
#[derive(Debug, Clone)]
pub(super) struct AutoFilterColumnDef {
    /// Column index within the filter range (0-based relative to filter start).
    pub column_index: u16,
    /// Join logic: true = OR, false = AND.
    pub join_or: bool,
    /// First condition.
    pub condition1: AutoFilterConditionWrite,
    /// Second condition.
    pub condition2: AutoFilterConditionWrite,
}

/// Sort configuration for the writer.
#[derive(Debug, Clone)]
pub(super) struct SortConfig {
    pub case_sensitive: bool,
    /// true = sort by columns (left-to-right), false = by rows (top-to-bottom)
    pub sort_by_columns: bool,
    /// Up to 3 sort keys: (column_index, descending).
    pub keys: Vec<(u16, bool)>,
}

impl WritableWorksheet {
    pub(super) fn new(name: String) -> Self {
        Self {
            name,
            cells: HashMap::new(),
            first_row: 0,
            last_row: 0,
            first_col: 0,
            last_col: 0,
            column_widths: HashMap::new(),
            hidden_columns: HashSet::new(),
            row_heights: HashMap::new(),
            hidden_rows: HashSet::new(),
            merged_ranges: Vec::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
            freeze_panes: None,
            sheet_protection: None,
            auto_filter: None,
            hyperlinks: Vec::new(),
            auto_filter_columns: Vec::new(),
            sort_config: None,
            pivot_tables: Vec::new(),
        }
    }

    pub(super) fn add_cell(&mut self, cell: WritableCell) {
        // Update dimensions
        if self.cells.is_empty() {
            self.first_row = cell.row;
            self.last_row = cell.row + 1;
            self.first_col = cell.col;
            self.last_col = cell.col + 1;
        } else {
            self.first_row = self.first_row.min(cell.row);
            self.last_row = self.last_row.max(cell.row + 1);
            self.first_col = self.first_col.min(cell.col);
            self.last_col = self.last_col.max(cell.col + 1);
        }

        self.cells.insert((cell.row, cell.col), cell);
    }

    pub(super) fn add_merged_range(&mut self, range: MergedRange) {
        self.merged_ranges.push(range);
    }

    pub(super) fn add_data_validation(&mut self, dv: XlsDataValidation) {
        self.data_validations.push(dv);
    }

    pub(super) fn add_conditional_format(&mut self, cf: XlsConditionalFormat) {
        self.conditional_formats.push(cf);
    }

    pub(super) fn set_freeze_panes(&mut self, freeze_rows: u32, freeze_cols: u16) {
        self.freeze_panes = Some(FreezePanes {
            freeze_rows,
            freeze_cols,
        });
    }

    pub(super) fn clear_freeze_panes(&mut self) {
        self.freeze_panes = None;
    }

    pub(super) fn set_column_width(&mut self, col: u16, width: u16) {
        self.column_widths.insert(col, width);
    }

    pub(super) fn hide_column(&mut self, col: u16) {
        self.hidden_columns.insert(col);
    }

    pub(super) fn add_hyperlink(&mut self, hyperlink: XlsHyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    pub(super) fn show_column(&mut self, col: u16) {
        self.hidden_columns.remove(&col);
    }

    pub(super) fn set_row_height(&mut self, row: u32, height: u16) {
        self.row_heights.insert(row, height);
    }

    pub(super) fn hide_row(&mut self, row: u32) {
        self.hidden_rows.insert(row);
    }

    pub(super) fn show_row(&mut self, row: u32) {
        self.hidden_rows.remove(&row);
    }

    pub(super) fn add_auto_filter_column(&mut self, def: AutoFilterColumnDef) {
        self.auto_filter_columns.push(def);
    }

    pub(super) fn set_sort_config(&mut self, config: SortConfig) {
        self.sort_config = Some(config);
    }

    pub(super) fn add_pivot_table(&mut self, pt: WritablePivotTable) {
        // Expand worksheet dimensions to encompass the pivot table output
        // range.  Excel validates that the DIMENSIONS record covers the
        // SXVIEW output area; a mismatch causes a "corrupt file" repair
        // dialog.
        let pt_first_row = pt.first_row as u32;
        let pt_last_row_excl = pt.last_row as u32 + 1; // DIMENSIONS uses exclusive end
        let pt_first_col = pt.first_col;
        let pt_last_col_excl = pt.last_col + 1;

        if self.cells.is_empty() && self.pivot_tables.is_empty() {
            self.first_row = pt_first_row;
            self.last_row = pt_last_row_excl;
            self.first_col = pt_first_col;
            self.last_col = pt_last_col_excl;
        } else {
            self.first_row = self.first_row.min(pt_first_row);
            self.last_row = self.last_row.max(pt_last_row_excl);
            self.first_col = self.first_col.min(pt_first_col);
            self.last_col = self.last_col.max(pt_last_col_excl);
        }

        self.pivot_tables.push(pt);
    }
}

/// A pivot table definition for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotTable {
    /// Pivot table name.
    pub name: String,
    /// Source type (0x0001 = Worksheet).
    pub source_type: u16,

    // -- Source data range (for DCONREF + SXDB cache) --
    /// Name of the source worksheet.
    pub source_sheet_name: String,
    /// Source range (0-based, inclusive).
    pub source_first_row: u16,
    pub source_last_row: u16,
    pub source_first_col: u16,
    pub source_last_col: u16,

    // -- Output range --
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
    /// First header row.
    pub first_header_row: u16,
    /// First data row.
    pub first_data_row: u16,
    /// First data column.
    pub first_data_col: u16,
    /// Data field header name (e.g. "Values").
    pub data_field_name: String,
    /// Axis for data field header.
    pub data_axis: u16,
    /// Position of data label within axis.
    pub data_position: u16,
    /// Field definitions.
    pub fields: Vec<WritablePivotField>,
    /// Data item definitions.
    pub data_items: Vec<WritablePivotDataItem>,
    /// Page field entries: (item_index, field_index, object_id).
    pub page_entries: Vec<(u16, u16, u16)>,
    /// Source data rows for the pivot cache.
    pub source_data: Vec<Vec<super::PivotCacheValue>>,
}

/// A pivot field definition for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotField {
    /// Axis: 0=none, 1=row, 2=col, 4=page, 8=data.
    pub axis: u16,
    pub subtotal_count: u16,
    pub subtotal_flags: u16,
    /// Items in this field.
    pub items: Vec<WritablePivotItem>,
    /// Optional SXVD display name override (`None` → use cache name).
    pub name: Option<String>,
    /// Source column name for the pivot cache SXFDB record.
    pub cache_name: String,
    /// Unique source data values for this field's cache items (SXSTRING records).
    pub cache_items: Vec<String>,
    /// Whether this field is numeric (data-axis).
    pub is_numeric: bool,
}

/// A pivot item for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotItem {
    /// Item type: 0x0000=Data, 0x0001=Default subtotal, 0x0002=Sum, etc.
    pub item_type: u16,
    pub flags: u16,
    pub cache_index: u16,
    pub name: Option<String>,
}

/// A pivot data item (value field) for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotDataItem {
    pub source_field_index: u16,
    /// Aggregation function: 0=Sum,1=Count,2=Average,...
    pub function: u16,
    pub display_format: u16,
    pub base_field_index: u16,
    pub base_item_index: u16,
    pub num_format_index: u16,
    pub name: String,
}
