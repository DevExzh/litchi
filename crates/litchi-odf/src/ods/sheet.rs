//! Sheet structures for ODS spreadsheets.

use super::{Column, Row, SheetProtection, TableStructure};
use litchi_core::Result;

/// A sheet (worksheet) in an ODS spreadsheet.
///
/// Sheets contain rows of cells and have a name for identification.
#[derive(Clone)]
pub struct Sheet {
    /// Sheet name
    pub name: String,
    /// Rows in this sheet
    pub rows: Vec<Row>,
    /// Logical table columns and their structural metadata.
    pub columns: Vec<Column>,
    /// Nested grouping and header structure for logical columns.
    pub column_structure: Vec<TableStructure>,
    /// Nested grouping and header structure for logical rows.
    pub row_structure: Vec<TableStructure>,
    /// Sheet protection metadata and edit permissions.
    pub protection: SheetProtection,
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

    /// Get all explicitly declared logical columns.
    pub fn columns(&self) -> &[Column] {
        &self.columns
    }

    /// Get the nested column grouping and header structure.
    pub fn column_structure(&self) -> &[TableStructure] {
        &self.column_structure
    }

    /// Get the nested row grouping and header structure.
    pub fn row_structure(&self) -> &[TableStructure] {
        &self.row_structure
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
        Ok(max_cols.max(self.columns.len()))
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
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
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
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
            name: "Test Sheet".to_string(),
            rows: vec![],
        };
        assert_eq!(sheet.name().unwrap(), "Test Sheet");
    }

    #[test]
    fn test_sheet_rows() {
        let sheet = Sheet {
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
            name: "Sheet1".to_string(),
            rows: vec![
                Row {
                    cells: vec![],
                    index: 0,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
                },
                Row {
                    cells: vec![],
                    index: 1,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
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
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
            name: "Sheet1".to_string(),
            rows: vec![
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 0,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 0,
                            col: 1,
                        },
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 0,
                            col: 2,
                        },
                    ],
                    index: 0,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
                },
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 1,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Empty,
                            text: String::new(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 1,
                            col: 1,
                        },
                    ],
                    index: 1,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
                },
            ],
        };
        // Should return max column count across all rows
        assert_eq!(sheet.column_count().unwrap(), 3);
    }

    #[test]
    fn test_sheet_column_count_empty() {
        let sheet = Sheet {
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
            name: "Empty".to_string(),
            rows: vec![],
        };
        assert_eq!(sheet.column_count().unwrap(), 0);
    }

    #[test]
    fn test_sheet_with_data() {
        let sheet = Sheet {
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
            name: "Data".to_string(),
            rows: vec![
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Text("A1".to_string()),
                            text: "A1".to_string(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 0,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Text("B1".to_string()),
                            text: "B1".to_string(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 0,
                            col: 1,
                        },
                    ],
                    index: 0,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
                },
                Row {
                    cells: vec![
                        Cell {
                            value: CellValue::Text("A2".to_string()),
                            text: "A2".to_string(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 1,
                            col: 0,
                        },
                        Cell {
                            value: CellValue::Text("B2".to_string()),
                            text: "B2".to_string(),
                            formula: None,
                            annotation: None,
                            validation_name: None,
                            style_name: None,
                            matrix_span: None,
                            merge: Default::default(),
                            protect: None,
                            protected: None,
                            row: 1,
                            col: 1,
                        },
                    ],
                    index: 1,
                    style_name: None,
                    default_cell_style_name: None,
                    visibility: Default::default(),
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
            columns: Vec::new(),
            column_structure: Vec::new(),
            row_structure: Vec::new(),
            protection: SheetProtection::default(),
            name: "Original".to_string(),
            rows: vec![],
        };
        let cloned = sheet.clone();
        assert_eq!(cloned.name().unwrap(), "Original");
    }
}
