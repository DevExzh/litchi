//! Numbers Table Structure
//!
//! Tables in Numbers contain cells organized in rows and columns.

use crate::cell::Value as CellValue;
use crate::table::{Builder as TableBuilder, Dimensions, Error as TableError, Position};
use litchi_iwa_common::comment::Comment;
use std::collections::HashMap;

pub(super) fn map_table_error(error: TableError) -> super::Error {
    match error {
        TableError::Allocation { resource, amount } => {
            super::Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        other => super::Error::InvalidFormat(other.to_string()),
    }
}

/// Represents a table in a Numbers spreadsheet
#[derive(Debug, Clone)]
pub(super) struct Table {
    model: TableBuilder,
    comments: HashMap<(usize, usize), Comment>,
    dynamic_dimensions: bool,
}

impl Table {
    /// Create a new empty table
    pub(super) fn new(name: impl Into<String>) -> Self {
        Self {
            model: TableBuilder::new(name, Dimensions::new(0, 0)),
            comments: HashMap::new(),
            dynamic_dimensions: true,
        }
    }

    /// Create a table with dimensions declared by its native archive.
    pub(super) fn with_dimensions(
        name: impl Into<String>,
        row_count: usize,
        column_count: usize,
    ) -> super::Result<Self> {
        let dimensions =
            Dimensions::try_from_usize(row_count, column_count).map_err(map_table_error)?;
        Ok(Self {
            model: TableBuilder::new(name, dimensions),
            comments: HashMap::new(),
            dynamic_dimensions: false,
        })
    }

    /// Borrow the table name.
    pub(super) fn name(&self) -> &str {
        self.model.name()
    }

    /// Return the number of addressable rows.
    pub(super) const fn row_count(&self) -> usize {
        self.model.dimensions().rows() as usize
    }

    /// Return the number of addressable columns.
    pub(super) const fn column_count(&self) -> usize {
        self.model.dimensions().columns() as usize
    }

    /// Iterate over column headers in native order.
    pub(super) fn column_headers(&self) -> impl Iterator<Item = &str> + '_ {
        self.model.column_headers()
    }

    /// Replace the column headers while retaining their native order.
    pub(super) fn set_column_headers<I, S>(&mut self, headers: I) -> super::Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.model
            .set_column_headers(headers)
            .map_err(map_table_error)
    }

    /// Iterate over row headers in native order.
    pub(super) fn row_headers(&self) -> impl Iterator<Item = &str> + '_ {
        self.model.row_headers()
    }

    /// Replace the row headers while retaining their native order.
    pub(super) fn set_row_headers<I, S>(&mut self, headers: I) -> super::Result<()>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.model.set_row_headers(headers).map_err(map_table_error)
    }

    /// Get a cell value at the specified position
    pub(super) fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        let position = Position::try_from_usize(row, col).ok()?;
        self.model.get(position)
    }

    /// Iterate over materialized cells without exposing the backing map.
    pub(super) fn iter_cells(&self) -> impl Iterator<Item = ((usize, usize), &CellValue)> + '_ {
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
    pub(super) fn cell_count(&self) -> usize {
        self.model.cell_count()
    }

    /// Set a cell value at the specified position.
    ///
    /// # Errors
    ///
    /// Returns an error when the coordinate is not representable, lies
    /// outside fixed archive dimensions, or a sparse-entry allocation fails.
    pub(super) fn set_cell(
        &mut self,
        row: usize,
        col: usize,
        value: CellValue,
    ) -> super::Result<()> {
        self.try_set_cell(row, col, value)
    }

    /// Fallible archive-safe cell insertion.
    pub(super) fn try_set_cell(
        &mut self,
        row: usize,
        col: usize,
        value: CellValue,
    ) -> super::Result<()> {
        let position =
            Position::try_from_usize(row, col).map_err(|error| map_table_error(error.into()))?;
        self.ensure_coordinate(row, col)?;
        self.model
            .set(position, value)
            .map_err(|failure| map_table_error(failure.into_parts().0))
    }

    fn ensure_coordinate(&mut self, row: usize, col: usize) -> super::Result<()> {
        let dimensions = self.model.dimensions();
        if self.dynamic_dimensions {
            let rows = row.checked_add(1).ok_or_else(|| {
                super::Error::InvalidFormat("Numbers table row coordinate overflows".to_owned())
            })?;
            let columns = col.checked_add(1).ok_or_else(|| {
                super::Error::InvalidFormat("Numbers table column coordinate overflows".to_owned())
            })?;
            let requested = Dimensions::try_from_usize(
                (dimensions.rows() as usize).max(rows),
                (dimensions.columns() as usize).max(columns),
            )
            .map_err(map_table_error)?;
            self.model.resize(requested).map_err(map_table_error)?;
        } else if row >= dimensions.rows() as usize || col >= dimensions.columns() as usize {
            return Err(super::Error::InvalidFormat(format!(
                "Numbers cell coordinate ({row}, {col}) is outside {}x{}",
                dimensions.rows(),
                dimensions.columns()
            )));
        }
        Ok(())
    }

    /// Get the comment attached to a cell, if any.
    pub(super) fn get_comment(&self, row: usize, col: usize) -> Option<&Comment> {
        self.comments.get(&(row, col))
    }

    /// Iterate over cell comments without exposing the backing map.
    pub(super) fn iter_comments(&self) -> impl Iterator<Item = ((usize, usize), &Comment)> + '_ {
        self.comments
            .iter()
            .map(|(position, comment)| (*position, comment))
    }

    /// Return the number of materialized cell comments.
    pub(super) fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Attach or replace an in-memory comment.
    ///
    /// # Errors
    ///
    /// Returns an error when the coordinate cannot be represented, lies
    /// outside fixed archive dimensions, or comment storage cannot grow.
    pub(super) fn set_comment(
        &mut self,
        row: usize,
        col: usize,
        comment: Comment,
    ) -> super::Result<()> {
        self.try_set_comment(row, col, comment)
    }

    /// Fallible archive-safe comment insertion.
    pub(super) fn try_set_comment(
        &mut self,
        row: usize,
        col: usize,
        comment: Comment,
    ) -> super::Result<()> {
        self.ensure_coordinate(row, col)?;
        self.comments.try_reserve(1).map_err(|_| {
            super::Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table comments",
                amount: 1,
            })
        })?;
        self.comments.insert((row, col), comment);
        Ok(())
    }

    /// Remove an in-memory comment and return it.
    pub(super) fn clear_comment(&mut self, row: usize, col: usize) -> Option<Comment> {
        self.comments.remove(&(row, col))
    }

    /// Get all cell values in a specific row
    pub(super) fn get_row(&self, row: usize) -> Vec<CellValue> {
        (0..self.column_count())
            .map(|col| self.get_cell(row, col).cloned().unwrap_or(CellValue::Empty))
            .collect()
    }

    /// Get all cell values in a specific column
    pub(super) fn get_column(&self, col: usize) -> Vec<CellValue> {
        (0..self.row_count())
            .map(|row| self.get_cell(row, col).cloned().unwrap_or(CellValue::Empty))
            .collect()
    }

    /// Convert table to CSV format
    pub(super) fn to_csv(&self) -> String {
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
    pub(super) fn dimensions(&self) -> (usize, usize) {
        (self.row_count(), self.column_count())
    }

    /// Check if table is empty
    pub(super) fn is_empty(&self) -> bool {
        self.model.cell_count() == 0
    }

    /// Get total number of non-empty cells
    pub(super) fn non_empty_cell_count(&self) -> usize {
        self.model.non_empty_cell_count()
    }

    /// Consume the archive adapter and return its dependency-free semantic
    /// table.
    ///
    /// Native comments are intentionally not part of this snapshot. They
    /// remain available through the adapter while the document readers and
    /// editors migrate to the canonical leaf model.
    pub(super) fn into_semantic(self) -> super::Result<crate::Table> {
        self.into_semantic_table()
    }

    /// Consume the archive adapter without allocating a comment sidecar.
    ///
    /// Structured extraction only needs the canonical sparse table. Keeping
    /// this path separate from [`Self::into_semantic_parts`] avoids sorting
    /// and boxing native comments that would otherwise be discarded.
    pub(super) fn into_semantic_table(self) -> super::Result<crate::Table> {
        let Self { model, .. } = self;
        model.finish().map_err(map_table_error)
    }

    /// Consume the archive adapter while moving its canonical sparse table
    /// and format-owned comment sidecar independently.
    pub(super) fn into_semantic_parts(
        self,
    ) -> super::Result<(crate::Table, Box<[((usize, usize), Comment)]>)> {
        let Self {
            model, comments, ..
        } = self;
        let table = model.finish().map_err(map_table_error)?;
        let mut sorted_comments = Vec::new();
        sorted_comments.try_reserve(comments.len()).map_err(|_| {
            super::Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table semantic comments",
                amount: comments.len(),
            })
        })?;
        sorted_comments.extend(comments);
        sorted_comments.sort_unstable_by_key(|(position, _comment)| *position);
        Ok((table, sorted_comments.into_boxed_slice()))
    }
}
