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

#[cfg(test)]
mod tests {
    use super::super::cell::{Cell, CellValue};
    use super::super::row::Row;
    use super::*;

    #[test]
    fn test_sheet_new() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            rows: vec![],
        };
        assert_eq!(sheet.name().unwrap(), "Sheet1");
        assert_eq!(sheet.row_count().unwrap(), 0);
        assert_eq!(sheet.column_count().unwrap(), 0);
    }

    #[test]
    fn test_sheet_name() {
        let sheet = Sheet {
            name: "Test Sheet".to_string(),
            rows: vec![],
        };
        assert_eq!(sheet.name().unwrap(), "Test Sheet");
    }

    #[test]
    fn test_sheet_rows() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            rows: vec![
                Row {
                    cells: vec![],
                    index: 0,
                },
                Row {
                    cells: vec![],
                    index: 1,
                },
            ],
        };
        assert_eq!(sheet.row_count().unwrap(), 2);
        let rows = sheet.rows().unwrap();
        assert_eq!(rows.len(), 2);
    }

    #[test]
    fn test_sheet_column_count() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            rows: vec![
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            row: 0,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            row: 0,
                            col: 1,
                        },
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            row: 0,
                            col: 2,
                        },
                    ],
                    index: 0,
                },
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            row: 1,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            row: 1,
                            col: 1,
                        },
                    ],
                    index: 1,
                },
            ],
        };
        // Should return max column count across all rows
        assert_eq!(sheet.column_count().unwrap(), 3);
    }

    #[test]
    fn test_sheet_column_count_empty() {
        let sheet = Sheet {
            name: "Empty".to_string(),
            rows: vec![],
        };
        assert_eq!(sheet.column_count().unwrap(), 0);
    }

    #[test]
    fn test_sheet_with_data() {
        let sheet = Sheet {
            name: "Data".to_string(),
            rows: vec![
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Text("A1".to_string()),
                            text: "A1".to_string(),
                            formula: None,
                            row: 0,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Text("B1".to_string()),
                            text: "B1".to_string(),
                            formula: None,
                            row: 0,
                            col: 1,
                        },
                    ],
                    index: 0,
                },
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Text("A2".to_string()),
                            text: "A2".to_string(),
                            formula: None,
                            row: 1,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Text("B2".to_string()),
                            text: "B2".to_string(),
                            formula: None,
                            row: 1,
                            col: 1,
                        },
                    ],
                    index: 1,
                },
            ],
        };

        assert_eq!(sheet.name().unwrap(), "Data");
        assert_eq!(sheet.row_count().unwrap(), 2);
        assert_eq!(sheet.column_count().unwrap(), 2);

        // Check we can access cells through rows
        let rows = sheet.rows().unwrap();
        assert_eq!(rows[0].cells[0].text, "A1");
        assert_eq!(rows[1].cells[1].text, "B2");
    }

    #[test]
    fn test_sheet_clone() {
        let sheet = Sheet {
            name: "Original".to_string(),
            rows: vec![],
        };
        let cloned = sheet.clone();
        assert_eq!(cloned.name().unwrap(), "Original");
    }
}
