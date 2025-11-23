use std::collections::HashMap;

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
    pub merged_ranges: Vec<MergedRange>,
    pub data_validations: Vec<XlsDataValidation>,
    pub conditional_formats: Vec<XlsConditionalFormat>,
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
            merged_ranges: Vec::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
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
}
