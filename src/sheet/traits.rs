//! Traits for spreadsheet abstraction.

use super::types::{CellValue, Result};
use std::fmt::Debug;

/// Represents an individual cell in a worksheet.
pub trait Cell {
    /// Get the row number (1-based).
    fn row(&self) -> u32;

    /// Get the column number (1-based).
    fn column(&self) -> u32;

    /// Get the cell coordinate (e.g., "A1").
    fn coordinate(&self) -> String;

    /// Get the cell value.
    fn value(&self) -> &CellValue;

    /// Check if the cell is empty.
    fn is_empty(&self) -> bool {
        matches!(self.value(), CellValue::Empty)
    }

    /// Check if the cell contains a formula.
    fn is_formula(&self) -> bool {
        false // Default implementation, can be overridden
    }

    /// Check if the cell contains a date/time value.
    fn is_date(&self) -> bool {
        matches!(self.value(), CellValue::DateTime(_))
    }
}

/// Iterator over cells in a worksheet.
pub trait CellIterator<'a> {
    /// Get the next cell.
    fn next(&mut self) -> Option<Result<Box<dyn Cell + 'a>>>;
}

/// Iterator over rows in a worksheet.
pub trait RowIterator<'a> {
    /// Get the next row (as a vector of cell values).
    fn next(&mut self) -> Option<Result<Vec<CellValue>>>;
}

/// Represents a worksheet (sheet) in a workbook.
pub trait Worksheet {
    /// Get the worksheet name.
    fn name(&self) -> &str;

    /// Get the number of rows in the worksheet.
    fn row_count(&self) -> usize;

    /// Get the number of columns in the worksheet.
    fn column_count(&self) -> usize;

    /// Get the dimensions as (min_row, min_col, max_row, max_col).
    /// Returns None if the worksheet is empty.
    fn dimensions(&self) -> Option<(u32, u32, u32, u32)>;

    /// Get a cell by row and column (1-based indexing).
    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn Cell + '_>>;

    /// Get a cell by coordinate (e.g., "A1").
    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn Cell + '_>>;

    /// Get all cells as an iterator.
    fn cells(&self) -> Box<dyn CellIterator<'_> + '_>;

    /// Get all rows as an iterator.
    fn rows(&self) -> Box<dyn RowIterator<'_> + '_>;

    /// Get a specific row by index (0-based).
    fn row(&self, row_idx: usize) -> Result<Vec<CellValue>>;

    /// Get cell value by row and column (1-based indexing).
    fn cell_value(&self, row: u32, column: u32) -> Result<CellValue>;
}

/// Iterator over worksheets in a workbook.
pub trait WorksheetIterator<'a> {
    /// Get the next worksheet.
    fn next(&mut self) -> Option<Result<Box<dyn Worksheet + 'a>>>;
}

/// Trait representing a workbook (spreadsheet file).
///
/// **Note**: This is the low-level trait API. For high-level usage, use the
/// unified `Workbook` struct from `crate::sheet::Workbook`.
pub trait WorkbookTrait: Debug {
    /// Get the active worksheet.
    fn active_worksheet(&self) -> Result<Box<dyn Worksheet + '_>>;

    /// Get all worksheet names.
    fn worksheet_names(&self) -> Vec<String>;

    /// Get a worksheet by name.
    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn Worksheet + '_>>;

    /// Get a worksheet by index.
    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn Worksheet + '_>>;

    /// Get all worksheets as an iterator.
    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_>;

    /// Get the number of worksheets.
    fn worksheet_count(&self) -> usize;

    /// Get the index of the active worksheet.
    fn active_sheet_index(&self) -> usize;
}
