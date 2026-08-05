//! Numbers Table Structure
//!
//! Tables in Numbers contain cells organized in rows and columns.

use super::cell::CellValue;
use litchi_iwa_common::comment::Comment;
use litchi_numbers::table::{Builder as TableBuilder, Dimensions, Error as TableError, Position};
use std::collections::HashMap;

pub(crate) fn map_table_error(error: TableError) -> crate::Error {
    match error {
        TableError::Allocation { resource, amount } => {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        other => crate::Error::InvalidFormat(other.to_string()),
    }
}

/// Represents a table in a Numbers spreadsheet
#[derive(Debug, Clone)]
pub struct NumbersTable {
    model: TableBuilder,
    comments: HashMap<(usize, usize), Comment>,
    dynamic_dimensions: bool,
}

impl NumbersTable {
    /// Create a new empty table
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            model: TableBuilder::new(name, Dimensions::new(0, 0)),
            comments: HashMap::new(),
            dynamic_dimensions: true,
        }
    }

    /// Create a table with dimensions declared by its native archive.
    pub(crate) fn with_dimensions(
        name: impl Into<String>,
        row_count: usize,
        column_count: usize,
    ) -> crate::Result<Self> {
        let dimensions =
            Dimensions::try_from_usize(row_count, column_count).map_err(map_table_error)?;
        Ok(Self {
            model: TableBuilder::new(name, dimensions),
            comments: HashMap::new(),
            dynamic_dimensions: false,
        })
    }

    /// Borrow the table name.
    pub fn name(&self) -> &str {
        self.model.name()
    }

    /// Return the number of addressable rows.
    pub const fn row_count(&self) -> usize {
        self.model.dimensions().rows() as usize
    }

    /// Return the number of addressable columns.
    pub const fn column_count(&self) -> usize {
        self.model.dimensions().columns() as usize
    }

    /// Iterate over column headers in native order.
    pub fn column_headers(&self) -> impl Iterator<Item = &str> + '_ {
        self.model.column_headers()
    }

    /// Replace the column headers while retaining their native order.
    pub fn set_column_headers<I, S>(&mut self, headers: I) -> crate::Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.model
            .set_column_headers(headers)
            .map_err(map_table_error)
    }

    /// Iterate over row headers in native order.
    pub fn row_headers(&self) -> impl Iterator<Item = &str> + '_ {
        self.model.row_headers()
    }

    /// Replace the row headers while retaining their native order.
    pub fn set_row_headers<I, S>(&mut self, headers: I) -> crate::Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.model.set_row_headers(headers).map_err(map_table_error)
    }

    /// Get a cell value at the specified position
    pub fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        let position = Position::try_from_usize(row, col).ok()?;
        self.model.get(position)
    }

    /// Iterate over materialized cells without exposing the backing map.
    pub fn iter_cells(&self) -> impl Iterator<Item = ((usize, usize), &CellValue)> + '_ {
        self.model.cells().map(|cell| {
            (
                (
                    cell.position().row() as usize,
                    cell.position().column() as usize,
                ),
                cell.value(),
            )
        })
    }

    /// Return the number of materialized cells, including explicit empty cells.
    pub fn cell_count(&self) -> usize {
        self.model.cell_count()
    }

    /// Set a cell value at the specified position.
    ///
    /// # Errors
    ///
    /// Returns an error when the coordinate is not representable, lies
    /// outside fixed archive dimensions, or a sparse-entry allocation fails.
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) -> crate::Result<()> {
        self.try_set_cell(row, col, value)
    }

    /// Fallible archive-safe cell insertion.
    pub(crate) fn try_set_cell(
        &mut self,
        row: usize,
        col: usize,
        value: CellValue,
    ) -> crate::Result<()> {
        let position = Position::try_from_usize(row, col).map_err(map_table_error)?;
        self.ensure_coordinate(row, col)?;
        self.model
            .set(position, value)
            .map_err(|failure| map_table_error(failure.into_parts().0))
    }

    fn ensure_coordinate(&mut self, row: usize, col: usize) -> crate::Result<()> {
        let dimensions = self.model.dimensions();
        if self.dynamic_dimensions {
            let rows = row.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("Numbers table row coordinate overflows".to_owned())
            })?;
            let columns = col.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("Numbers table column coordinate overflows".to_owned())
            })?;
            let requested = Dimensions::try_from_usize(
                (dimensions.rows() as usize).max(rows),
                (dimensions.columns() as usize).max(columns),
            )
            .map_err(map_table_error)?;
            self.model.resize(requested).map_err(map_table_error)?;
        } else if row >= dimensions.rows() as usize || col >= dimensions.columns() as usize {
            return Err(crate::Error::InvalidFormat(format!(
                "Numbers cell coordinate ({row}, {col}) is outside {}x{}",
                dimensions.rows(),
                dimensions.columns()
            )));
        }
        Ok(())
    }

    /// Get the comment attached to a cell, if any.
    pub fn get_comment(&self, row: usize, col: usize) -> Option<&Comment> {
        self.comments.get(&(row, col))
    }

    /// Iterate over cell comments without exposing the backing map.
    pub fn iter_comments(
        &self,
    ) -> impl Iterator<Item = ((usize, usize), &Comment)> + '_ {
        self.comments
            .iter()
            .map(|(position, comment)| (*position, comment))
    }

    /// Return the number of materialized cell comments.
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Attach or replace an in-memory comment.
    ///
    /// # Errors
    ///
    /// Returns an error when the coordinate cannot be represented, lies
    /// outside fixed archive dimensions, or comment storage cannot grow.
    pub fn set_comment(
        &mut self,
        row: usize,
        col: usize,
        comment: Comment,
    ) -> crate::Result<()> {
        self.try_set_comment(row, col, comment)
    }

    /// Fallible archive-safe comment insertion.
    pub(crate) fn try_set_comment(
        &mut self,
        row: usize,
        col: usize,
        comment: Comment,
    ) -> crate::Result<()> {
        self.ensure_coordinate(row, col)?;
        self.comments.try_reserve(1).map_err(|_| {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table comments",
                amount: 1,
            })
        })?;
        self.comments.insert((row, col), comment);
        Ok(())
    }

    /// Remove an in-memory comment and return it.
    pub fn clear_comment(&mut self, row: usize, col: usize) -> Option<Comment> {
        self.comments.remove(&(row, col))
    }

    /// Get all cell values in a specific row
    pub fn get_row(&self, row: usize) -> Vec<CellValue> {
        (0..self.column_count())
            .map(|col| self.get_cell(row, col).cloned().unwrap_or(CellValue::Empty))
            .collect()
    }

    /// Get all cell values in a specific column
    pub fn get_column(&self, col: usize) -> Vec<CellValue> {
        (0..self.row_count())
            .map(|row| self.get_cell(row, col).cloned().unwrap_or(CellValue::Empty))
            .collect()
    }

    /// Convert table to CSV format
    pub fn to_csv(&self) -> String {
        let mut csv = String::new();

        // Add column headers if present
        let column_headers: Vec<_> = self.column_headers().collect();
        if !column_headers.is_empty() {
            csv.push_str(&column_headers.join(","));
            csv.push('\n');
        }

        // Add data rows
        let row_headers: Vec<_> = self.row_headers().collect();
        for row in 0..self.row_count() {
            // Add row header if present
            if row < row_headers.len() && !row_headers[row].is_empty() {
                csv.push_str(row_headers[row]);
                csv.push(',');
            }

            // Add cell values
            for col in 0..self.column_count() {
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
        (self.row_count(), self.column_count())
    }

    /// Check if table is empty
    pub fn is_empty(&self) -> bool {
        self.model.cell_count() == 0
    }

    /// Get total number of non-empty cells
    pub fn non_empty_cell_count(&self) -> usize {
        self.model.non_empty_cell_count()
    }

    /// Consume the archive adapter and return its dependency-free semantic
    /// table.
    ///
    /// Native comments are intentionally not part of this snapshot. They
    /// remain available through the adapter while the document readers and
    /// editors migrate to the canonical leaf model.
    pub(crate) fn into_semantic(self) -> crate::Result<litchi_numbers::Table> {
        self.into_semantic_table()
    }

    /// Consume the archive adapter without allocating a comment sidecar.
    ///
    /// Structured extraction only needs the canonical sparse table. Keeping
    /// this path separate from [`Self::into_semantic_parts`] avoids sorting
    /// and boxing native comments that would otherwise be discarded.
    pub(crate) fn into_semantic_table(self) -> crate::Result<litchi_numbers::Table> {
        let Self { model, .. } = self;
        model.finish().map_err(map_table_error)
    }

    /// Consume the archive adapter while moving its canonical sparse table
    /// and format-owned comment sidecar independently.
    pub(crate) fn into_semantic_parts(
        self,
    ) -> crate::Result<(
        litchi_numbers::Table,
        Box<[((usize, usize), Comment)]>,
    )> {
        let Self {
            model, comments, ..
        } = self;
        let table = model.finish().map_err(map_table_error)?;
        let mut sorted_comments = Vec::new();
        sorted_comments.try_reserve(comments.len()).map_err(|_| {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table semantic comments",
                amount: comments.len(),
            })
        })?;
        sorted_comments.extend(comments);
        sorted_comments.sort_unstable_by_key(|(position, _comment)| *position);
        Ok((table, sorted_comments.into_boxed_slice()))
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

        assert!(
            table
                .set_cell(0, 0, CellValue::Text("A1".to_string()))
                .is_ok()
        );
        assert!(
            table
                .set_cell(0, 1, CellValue::Text("B1".to_string()))
                .is_ok()
        );
        assert!(table.set_cell(1, 0, CellValue::Number(42.0)).is_ok());

        assert_eq!(table.row_count(), 2);
        assert_eq!(table.column_count(), 2);
        assert!(!table.is_empty());
    }

    #[test]
    fn test_table_get_row_column() {
        let mut table = NumbersTable::new("Test".to_string());
        assert!(table.set_cell(0, 0, CellValue::Number(1.0)).is_ok());
        assert!(table.set_cell(0, 1, CellValue::Number(2.0)).is_ok());
        assert!(table.set_cell(1, 0, CellValue::Number(3.0)).is_ok());
        assert!(table.set_cell(1, 1, CellValue::Number(4.0)).is_ok());

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
        assert!(table.set_column_headers(["Name", "Age"]).is_ok());
        assert!(
            table
                .set_cell(0, 0, CellValue::Text("Alice".to_string()))
                .is_ok()
        );
        assert!(table.set_cell(0, 1, CellValue::Number(30.0)).is_ok());
        assert!(
            table
                .set_cell(1, 0, CellValue::Text("Bob".to_string()))
                .is_ok()
        );
        assert!(table.set_cell(1, 1, CellValue::Number(25.0)).is_ok());

        let csv = table.to_csv();
        assert!(csv.contains("Name,Age"));
        assert!(csv.contains("Alice,30"));
        assert!(csv.contains("Bob,25"));
    }

    #[test]
    fn test_table_dimensions() {
        let mut table = NumbersTable::new("Test".to_string());
        assert!(table.set_cell(5, 10, CellValue::Number(1.0)).is_ok());

        let (rows, cols) = table.dimensions();
        assert_eq!(rows, 6); // 0-5 inclusive
        assert_eq!(cols, 11); // 0-10 inclusive
    }

    #[test]
    fn materialized_views_borrow_without_exposing_storage() {
        let mut table = NumbersTable::new("Test");
        assert!(table.set_cell(2, 3, CellValue::Number(42.0)).is_ok());
        assert!(table.set_column_headers(["A", "B"]).is_ok());
        assert!(table.set_row_headers(["Row 1"]).is_ok());

        assert_eq!(table.cell_count(), 1);
        assert_eq!(table.comment_count(), 0);
        assert_eq!(table.iter_cells().collect::<Vec<_>>().len(), 1);
        assert_eq!(table.column_headers().collect::<Vec<_>>(), ["A", "B"]);
        assert_eq!(table.row_headers().collect::<Vec<_>>(), ["Row 1"]);
    }
}
