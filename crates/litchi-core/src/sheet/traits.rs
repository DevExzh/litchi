//! Traits for spreadsheet abstraction.

use super::types::{CellValue, Result};
use std::borrow::Cow;
use std::fmt::Debug;

/// Represents an individual cell in a worksheet.
pub trait Cell: Send + Sync {
    /// Get the column number (1-based).
    fn column(&self) -> u32;

    /// Get the cell coordinate (e.g., "A1").
    fn coordinate(&self) -> String;

    /// Check if the cell contains a date/time value.
    fn is_date(&self) -> bool {
        matches!(self.value(), CellValue::DateTime(_))
    }

    /// Check if the cell is empty.
    fn is_empty(&self) -> bool {
        matches!(self.value(), CellValue::Empty)
    }

    /// Check if the cell contains a formula.
    fn is_formula(&self) -> bool {
        false // Default implementation, can be overridden
    }

    /// Get the row number (1-based).
    fn row(&self) -> u32;

    /// Get the cell value.
    fn value(&self) -> &CellValue;
}

/// Iterator over cells in a worksheet.
pub trait CellIterator<'a> {
    /// Get the next cell.
    fn next(&mut self) -> Option<Result<Box<dyn Cell + 'a>>>;
}

/// Iterator over rows in a worksheet.
pub trait RowIterator<'a> {
    /// Get the next row (as a vector of cell values wrapped in Cow for zero-copy when possible).
    fn next(&mut self) -> Option<Result<Cow<'a, [CellValue]>>>;
}

/// Represents a worksheet (sheet) in a workbook.
pub trait Worksheet: Send + Sync {
    /// Get a cell by row and column (1-based indexing).
    ///
    /// # Errors
    ///
    /// Returns an error if the cell cannot be read.
    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn Cell + '_>>;

    /// Get a cell by coordinate (e.g., "A1").
    ///
    /// # Errors
    ///
    /// Returns an error if the coordinate is invalid or the cell cannot be read.
    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn Cell + '_>>;

    /// Get cell value by row and column (1-based indexing).
    ///
    /// Returns a Cow to allow zero-copy when possible while supporting
    /// implementations that need to compute values (e.g., shared string resolution).
    ///
    /// # Errors
    ///
    /// Returns an error if the cell value cannot be resolved.
    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>>;

    /// Get all cells as an iterator.
    fn cells(&self) -> Box<dyn CellIterator<'_> + '_>;

    /// Get the number of columns in the worksheet.
    fn column_count(&self) -> usize;

    /// Get the dimensions as (`min_row`, `min_col`, `max_row`, `max_col`).
    /// Returns None if the worksheet is empty.
    fn dimensions(&self) -> Option<(u32, u32, u32, u32)>;

    /// Get the worksheet name.
    fn name(&self) -> &str;

    /// Get a specific row by index (0-based).
    ///
    /// Returns a Cow to allow zero-copy when possible while supporting
    /// both owned and borrowed data depending on implementation.
    ///
    /// # Errors
    ///
    /// Returns an error if the row cannot be read.
    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>>;

    /// Get the number of rows in the worksheet.
    fn row_count(&self) -> usize;

    /// Get all rows as an iterator.
    fn rows(&self) -> Box<dyn RowIterator<'_> + '_>;
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
pub trait WorkbookTrait: Debug + Send + Sync {
    /// Get the index of the active worksheet.
    fn active_sheet_index(&self) -> usize;

    /// Get the active worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error if the active worksheet cannot be determined.
    fn active_worksheet(&self) -> Result<Box<dyn Worksheet + '_>>;

    /// Returns true if the workbook uses the 1904 date system (Mac).
    ///
    /// Implementations should report the correct setting based on workbook metadata.
    fn is_1904_date_system(&self) -> bool {
        false
    }

    /// Get a worksheet by index.
    ///
    /// # Errors
    ///
    /// Returns an error if the index is out of range.
    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn Worksheet + '_>>;

    /// Get a worksheet by name.
    ///
    /// # Errors
    ///
    /// Returns an error if no worksheet with that name exists.
    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn Worksheet + '_>>;

    /// Get the number of worksheets.
    fn worksheet_count(&self) -> usize;

    /// Get all worksheet names (zero-copy).
    ///
    /// Returns a slice reference to avoid cloning. Implementations
    /// should store worksheet names internally.
    fn worksheet_names(&self) -> &[String];

    /// Get all worksheets as an iterator.
    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_>;
}
