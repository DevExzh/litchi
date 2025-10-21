//! Row structures for ODS spreadsheets.

use super::Cell;
use crate::common::Result;

/// A row in an ODS spreadsheet.
///
/// Rows contain cells and maintain their position within a sheet.
#[derive(Clone)]
pub struct Row {
    /// Cells in this row
    pub cells: Vec<Cell>,
    /// Row index (0-based)
    pub index: usize,
}

impl Row {
    /// Get all cells in the row.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        Ok(self.cells.clone())
    }

    /// Get a cell by column index.
    ///
    /// Returns `Some(cell)` if a cell exists at the given column index,
    /// `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `col` - Column index (0-based)
    pub fn cell(&self, col: usize) -> Result<Option<Cell>> {
        if col < self.cells.len() {
            Ok(Some(self.cells[col].clone()))
        } else {
            Ok(None)
        }
    }

    /// Get the row index.
    ///
    /// Returns the 0-based row index within the sheet.
    pub fn index(&self) -> usize {
        self.index
    }
}
