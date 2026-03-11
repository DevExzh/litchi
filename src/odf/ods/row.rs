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

#[cfg(test)]
mod tests {
    use super::super::cell::{Cell, CellValue};
    use super::*;

    #[test]
    fn test_row_new() {
        let row = Row {
            cells: vec![],
            index: 0,
        };
        assert_eq!(row.index(), 0);
        assert_eq!(row.cell_count().unwrap(), 0);
    }

    #[test]
    fn test_row_with_cells() {
        let row = Row {
            cells: vec![
                Cell {
                    value: CellValue::Text("A".to_string()),
                    text: "A".to_string(),
                    formula: None,
                    row: 0,
                    col: 0,
                },
                Cell {
                    value: CellValue::Text("B".to_string()),
                    text: "B".to_string(),
                    formula: None,
                    row: 0,
                    col: 1,
                },
            ],
            index: 5,
        };
        assert_eq!(row.index(), 5);
        assert_eq!(row.cell_count().unwrap(), 2);
    }

    #[test]
    fn test_row_cells() {
        let row = Row {
            cells: vec![Cell {
                value: CellValue::Number(1.0),
                text: "1".to_string(),
                formula: None,
                row: 0,
                col: 0,
            }],
            index: 0,
        };
        let cells = row.cells().unwrap();
        assert_eq!(cells.len(), 1);
    }

    #[test]
    fn test_row_cell() {
        let row = Row {
            cells: vec![
                Cell {
                    value: CellValue::Text("First".to_string()),
                    text: "First".to_string(),
                    formula: None,
                    row: 0,
                    col: 0,
                },
                Cell {
                    value: CellValue::Text("Second".to_string()),
                    text: "Second".to_string(),
                    formula: None,
                    row: 0,
                    col: 1,
                },
            ],
            index: 0,
        };

        assert!(row.cell(0).unwrap().is_some());
        assert!(row.cell(1).unwrap().is_some());
        assert!(row.cell(2).unwrap().is_none());
    }

    #[test]
    fn test_row_cell_at() {
        let row = Row {
            cells: vec![Cell {
                value: CellValue::Text("Test".to_string()),
                text: "Test".to_string(),
                formula: None,
                row: 0,
                col: 0,
            }],
            index: 0,
        };

        let cell = row.cell_at(0).unwrap();
        assert!(cell.is_some());
        assert_eq!(cell.unwrap().text, "Test");
    }

    #[test]
    fn test_row_index() {
        let row = Row {
            cells: vec![],
            index: 42,
        };
        assert_eq!(row.index(), 42);
    }

    #[test]
    fn test_row_empty() {
        let row = Row {
            cells: vec![],
            index: 0,
        };
        assert_eq!(row.cell_count().unwrap(), 0);
        assert!(row.cell(0).unwrap().is_none());
    }

    #[test]
    fn test_row_multiple_cells() {
        let mut cells = vec![];
        for i in 0..10 {
            cells.push(Cell {
                value: CellValue::Number(i as f64),
                text: i.to_string(),
                formula: None,
                row: 0,
                col: i,
            });
        }

        let row = Row { cells, index: 1 };
        assert_eq!(row.cell_count().unwrap(), 10);

        for i in 0..10 {
            let cell = row.cell(i).unwrap().unwrap();
            assert_eq!(cell.text, i.to_string());
        }
    }
}
