//! Row structures for ODS spreadsheets.

use super::Cell;
use crate::common::Result;

/// A row in an ODS spreadsheet.
///
/// Rows contain cells and maintain their position within a sheet.
#[derive(Debug, Clone)]
pub struct Row {
    /// Cells in this row
    pub cells: Vec<Cell>,
    /// Row index (0-based)
    pub index: usize,
}

impl Row {
    /// Get all cells in the row.
    pub fn cells(&self) -> Result<&[Cell]> {
        Ok(&self.cells)
    }

    /// Get the number of cells in the row.
    pub fn cell_count(&self) -> Result<usize> {
        Ok(self.cells.len())
    }

    /// Get a cell by column index.
    ///
    /// Returns `Some(cell)` if a cell exists at the given column index,
    /// `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `col` - Column index (0-based)
    pub fn cell(&self, col: usize) -> Result<Option<&Cell>> {
        if col < self.cells.len() {
            Ok(Some(&self.cells[col]))
        } else {
            Ok(None)
        }
    }

    /// Get a cell by column index (alias for unified API).
    ///
    /// # Arguments
    ///
    /// * `col` - Column index (0-based)
    pub fn cell_at(&self, col: usize) -> Result<Option<&Cell>> {
        self.cell(col)
    }

    /// Get the row index.
    ///
    /// Returns the 0-based row index within the sheet.
    pub fn index(&self) -> usize {
        self.index
    }
}
