use std::collections::{HashMap, HashSet};

use super::{XlsCellValue, XlsConditionalFormat, XlsDataValidation};

#[derive(Debug, Clone)]
pub(super) struct WritableCell {
    /// Row index (0-based)
    pub row: u32,
    /// Column index (0-based)
    pub col: u16,
    /// Cell value
    pub value: XlsCellValue,
    pub format_idx: u16,
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
    pub auto_filter: Option<AutoFilterRange>,
    /// Cell or range hyperlinks stored for this worksheet.
    pub hyperlinks: Vec<XlsHyperlink>,
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
            auto_filter: None,
            hyperlinks: Vec::new(),
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
}
