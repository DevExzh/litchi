//! Numbers Table Structure
//!
//! Tables in Numbers contain cells organized in rows and columns.

use super::cell::CellValue;
use std::collections::HashMap;

/// Stable UUID stored on a Numbers comment archive.
pub type NumbersCommentUuid = crate::comments::IWorkCommentUuid;

/// A comment attached to a Numbers table cell.
pub type NumbersCellComment = crate::comments::IWorkComment;

/// Represents a table in a Numbers spreadsheet
#[derive(Debug, Clone)]
pub struct NumbersTable {
    name: String,
    row_count: usize,
    column_count: usize,
    cells: HashMap<(usize, usize), CellValue>,
    comments: HashMap<(usize, usize), NumbersCellComment>,
    column_headers: Vec<String>,
    row_headers: Vec<String>,
}

impl NumbersTable {
    /// Create a new empty table
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            row_count: 0,
            column_count: 0,
            cells: HashMap::new(),
            comments: HashMap::new(),
            column_headers: Vec::new(),
            row_headers: Vec::new(),
        }
    }

    /// Create a table with dimensions declared by its native archive.
    pub(crate) fn with_dimensions(
        name: impl Into<String>,
        row_count: usize,
        column_count: usize,
    ) -> Self {
        Self {
            name: name.into(),
            row_count,
            column_count,
            cells: HashMap::new(),
            comments: HashMap::new(),
            column_headers: Vec::new(),
            row_headers: Vec::new(),
        }
    }

    /// Borrow the table name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the number of addressable rows.
    pub const fn row_count(&self) -> usize {
        self.row_count
    }

    /// Return the number of addressable columns.
    pub const fn column_count(&self) -> usize {
        self.column_count
    }

    /// Iterate over column headers in native order.
    pub fn column_headers(&self) -> impl Iterator<Item = &str> + '_ {
        self.column_headers.iter().map(String::as_str)
    }

    /// Replace the column headers while retaining their native order.
    pub fn set_column_headers<I, S>(&mut self, headers: I)
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.column_headers = headers.into_iter().map(Into::into).collect();
    }

    /// Iterate over row headers in native order.
    pub fn row_headers(&self) -> impl Iterator<Item = &str> + '_ {
        self.row_headers.iter().map(String::as_str)
    }

    /// Replace the row headers while retaining their native order.
    pub fn set_row_headers<I, S>(&mut self, headers: I)
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.row_headers = headers.into_iter().map(Into::into).collect();
    }

    /// Get a cell value at the specified position
    pub fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        self.cells.get(&(row, col))
    }

    /// Iterate over materialized cells without exposing the backing map.
    pub fn iter_cells(&self) -> impl Iterator<Item = ((usize, usize), &CellValue)> + '_ {
        self.cells
            .iter()
            .map(|(position, value)| (*position, value))
    }

    /// Return the number of materialized cells, including explicit empty cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Set a cell value at the specified position
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) {
        self.cells.insert((row, col), value);
        self.row_count = self.row_count.max(row + 1);
        self.column_count = self.column_count.max(col + 1);
    }

    /// Get the comment attached to a cell, if any.
    pub fn get_comment(&self, row: usize, col: usize) -> Option<&NumbersCellComment> {
        self.comments.get(&(row, col))
    }

    /// Iterate over cell comments without exposing the backing map.
    pub fn iter_comments(
        &self,
    ) -> impl Iterator<Item = ((usize, usize), &NumbersCellComment)> + '_ {
        self.comments
            .iter()
            .map(|(position, comment)| (*position, comment))
    }

    /// Return the number of materialized cell comments.
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Attach or replace an in-memory comment.
    pub fn set_comment(&mut self, row: usize, col: usize, comment: NumbersCellComment) {
        self.comments.insert((row, col), comment);
        self.row_count = self.row_count.max(row + 1);
        self.column_count = self.column_count.max(col + 1);
    }

    /// Remove an in-memory comment and return it.
    pub fn clear_comment(&mut self, row: usize, col: usize) -> Option<NumbersCellComment> {
        self.comments.remove(&(row, col))
    }

    /// Get all cell values in a specific row
    pub fn get_row(&self, row: usize) -> Vec<CellValue> {
        (0..self.column_count)
            .map(|col| self.get_cell(row, col).cloned().unwrap_or(CellValue::Empty))
            .collect()
    }

    /// Get all cell values in a specific column
    pub fn get_column(&self, col: usize) -> Vec<CellValue> {
        (0..self.row_count)
            .map(|row| self.get_cell(row, col).cloned().unwrap_or(CellValue::Empty))
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

    /// Move the materialized values and comments to another crate-internal
    /// table view without cloning either sparse map.
    pub(crate) fn into_parts(
        self,
    ) -> (
        HashMap<(usize, usize), CellValue>,
        HashMap<(usize, usize), NumbersCellComment>,
    ) {
        (self.cells, self.comments)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_creation() {
        let mut table = NumbersTable::new("Test Table".to_string());
        assert_eq!(table.name(), "Test Table");
        assert_eq!(table.row_count(), 0);
        assert_eq!(table.column_count(), 0);
        assert!(table.is_empty());

        table.set_cell(0, 0, CellValue::Text("A1".to_string()));
        table.set_cell(0, 1, CellValue::Text("B1".to_string()));
        table.set_cell(1, 0, CellValue::Number(42.0));

        assert_eq!(table.row_count(), 2);
        assert_eq!(table.column_count(), 2);
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
        table.set_column_headers(["Name", "Age"]);
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

    #[test]
    fn materialized_views_borrow_without_exposing_storage() {
        let mut table = NumbersTable::new("Test");
        table.set_cell(2, 3, CellValue::Number(42.0));
        table.set_column_headers(["A", "B"]);
        table.set_row_headers(["Row 1"]);

        assert_eq!(table.cell_count(), 1);
        assert_eq!(table.comment_count(), 0);
        assert_eq!(table.iter_cells().collect::<Vec<_>>().len(), 1);
        assert_eq!(table.column_headers().collect::<Vec<_>>(), ["A", "B"]);
        assert_eq!(table.row_headers().collect::<Vec<_>>(), ["Row 1"]);
    }
}
