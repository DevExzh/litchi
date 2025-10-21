//! Sheet structures for ODS spreadsheets.

use super::Row;
use crate::common::Result;

/// A sheet (worksheet) in an ODS spreadsheet.
///
/// Sheets contain rows of cells and have a name for identification.
#[derive(Clone)]
pub struct Sheet {
    /// Sheet name
    pub name: String,
    /// Rows in this sheet
    pub rows: Vec<Row>,
}

impl Sheet {
    /// Get the name of the sheet.
    pub fn name(&self) -> Result<&str> {
        Ok(&self.name)
    }

    /// Get all rows in the sheet.
    pub fn rows(&self) -> Result<&[Row]> {
        Ok(&self.rows)
    }

    /// Get the number of rows in the sheet.
    ///
    /// Returns the total number of rows, including empty rows.
    pub fn row_count(&self) -> Result<usize> {
        Ok(self.rows.len())
    }

    /// Get the number of columns in the sheet.
    ///
    /// Returns the maximum number of columns across all rows.
    /// This accounts for rows with different numbers of cells.
    pub fn column_count(&self) -> Result<usize> {
        let max_cols = self
            .rows
            .iter()
            .map(|row| row.cells.len())
            .max()
            .unwrap_or(0);
        Ok(max_cols)
    }
}
