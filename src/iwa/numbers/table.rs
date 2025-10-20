//! Numbers Table Structure
//!
//! Tables in Numbers contain cells organized in rows and columns.

use std::collections::HashMap;
use super::cell::CellValue;

/// Represents a table in a Numbers spreadsheet
#[derive(Debug, Clone)]
pub struct NumbersTable {
    /// Table name
    pub name: String,
    /// Number of rows
    pub row_count: usize,
    /// Number of columns
    pub column_count: usize,
    /// Cell data indexed by (row, column)
    pub cells: HashMap<(usize, usize), CellValue>,
    /// Column headers (if present)
    pub column_headers: Vec<String>,
    /// Row headers (if present)
    pub row_headers: Vec<String>,
}

impl NumbersTable {
    /// Create a new empty table
    pub fn new(name: String) -> Self {
        Self {
            name,
            row_count: 0,
            column_count: 0,
            cells: HashMap::new(),
            column_headers: Vec::new(),
            row_headers: Vec::new(),
        }
    }

    /// Get a cell value at the specified position
    pub fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        self.cells.get(&(row, col))
    }

    /// Set a cell value at the specified position
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) {
        self.cells.insert((row, col), value);
        self.row_count = self.row_count.max(row + 1);
        self.column_count = self.column_count.max(col + 1);
    }

    /// Get all cell values in a specific row
    pub fn get_row(&self, row: usize) -> Vec<CellValue> {
        (0..self.column_count)
            .map(|col| {
                self.get_cell(row, col)
                    .cloned()
                    .unwrap_or(CellValue::Empty)
            })
            .collect()
    }

    /// Get all cell values in a specific column
    pub fn get_column(&self, col: usize) -> Vec<CellValue> {
        (0..self.row_count)
            .map(|row| {
                self.get_cell(row, col)
                    .cloned()
                    .unwrap_or(CellValue::Empty)
            })
            .collect()
    }

    /// Convert table to CSV format
    pub fn to_csv(&self) -> String {
        let mut csv = String::new();

        // Add column headers if present
        if !self.column_headers.is_empty() {
            csv.push_str(&self.column_headers.join(","));
            csv.push('\n');
        }

        // Add data rows
        for row in 0..self.row_count {
            // Add row header if present
            if row < self.row_headers.len() && !self.row_headers[row].is_empty() {
                csv.push_str(&self.row_headers[row]);
                csv.push(',');
            }

            // Add cell values
            for col in 0..self.column_count {
                if col > 0 {
                    csv.push(',');
                }
                if let Some(cell) = self.get_cell(row, col) {
                    csv.push_str(&cell.to_string());
                }
            }
            csv.push('\n');
        }

        csv
    }

    /// Get table dimensions as (rows, columns)
    pub fn dimensions(&self) -> (usize, usize) {
        (self.row_count, self.column_count)
    }

    /// Check if table is empty
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    /// Get total number of non-empty cells
    pub fn non_empty_cell_count(&self) -> usize {
        self.cells.values().filter(|v| !v.is_empty()).count()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_creation() {
        let mut table = NumbersTable::new("Test Table".to_string());
        assert_eq!(table.name, "Test Table");
        assert_eq!(table.row_count, 0);
        assert_eq!(table.column_count, 0);
        assert!(table.is_empty());

        table.set_cell(0, 0, CellValue::Text("A1".to_string()));
        table.set_cell(0, 1, CellValue::Text("B1".to_string()));
        table.set_cell(1, 0, CellValue::Number(42.0));

        assert_eq!(table.row_count, 2);
        assert_eq!(table.column_count, 2);
        assert!(!table.is_empty());
    }

    #[test]
    fn test_table_get_row_column() {
        let mut table = NumbersTable::new("Test".to_string());
        table.set_cell(0, 0, CellValue::Number(1.0));
        table.set_cell(0, 1, CellValue::Number(2.0));
        table.set_cell(1, 0, CellValue::Number(3.0));
        table.set_cell(1, 1, CellValue::Number(4.0));

        let row0 = table.get_row(0);
        assert_eq!(row0.len(), 2);
        assert_eq!(row0[0].as_number(), Some(1.0));
        assert_eq!(row0[1].as_number(), Some(2.0));

        let col0 = table.get_column(0);
        assert_eq!(col0.len(), 2);
        assert_eq!(col0[0].as_number(), Some(1.0));
        assert_eq!(col0[1].as_number(), Some(3.0));
    }

    #[test]
    fn test_table_to_csv() {
        let mut table = NumbersTable::new("Test".to_string());
        table.column_headers = vec!["Name".to_string(), "Age".to_string()];
        table.set_cell(0, 0, CellValue::Text("Alice".to_string()));
        table.set_cell(0, 1, CellValue::Number(30.0));
        table.set_cell(1, 0, CellValue::Text("Bob".to_string()));
        table.set_cell(1, 1, CellValue::Number(25.0));

        let csv = table.to_csv();
        assert!(csv.contains("Name,Age"));
        assert!(csv.contains("Alice,30"));
        assert!(csv.contains("Bob,25"));
    }

    #[test]
    fn test_table_dimensions() {
        let mut table = NumbersTable::new("Test".to_string());
        table.set_cell(5, 10, CellValue::Number(1.0));
        
        let (rows, cols) = table.dimensions();
        assert_eq!(rows, 6); // 0-5 inclusive
        assert_eq!(cols, 11); // 0-10 inclusive
    }
}

