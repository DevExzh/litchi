//! RTF table support.
//!
//! This module provides basic table parsing for RTF documents.
//! RTF tables use a complex row-based model with cell boundaries.

use crate::TextDirection;
use std::borrow::Cow;

/// A table in an RTF document.
#[derive(Debug, Clone)]
pub struct Table<'a> {
    /// Table rows
    rows: Vec<Row<'a>>,
    /// Explicit table direction from `\taprtl` or `\taprtl0`.
    direction: Option<TextDirection>,
}

impl<'a> Table<'a> {
    /// Create a new table.
    pub fn new() -> Self {
        Self {
            rows: Vec::new(),
            direction: None,
        }
    }

    /// Add a row to the table.
    pub fn add_row(&mut self, row: Row<'a>) {
        self.rows.push(row);
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get all rows.
    pub fn rows(&self) -> &[Row<'a>] {
        &self.rows
    }

    /// Return the explicit table direction.
    pub fn direction(&self) -> Option<TextDirection> {
        self.direction
    }

    /// Set or clear the explicit table direction.
    pub fn set_direction(&mut self, direction: Option<TextDirection>) {
        self.direction = direction;
    }
}

impl<'a> Default for Table<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// A table row.
#[derive(Debug, Clone)]
pub struct Row<'a> {
    /// Row cells
    cells: Vec<Cell<'a>>,
    /// Explicit row direction.
    direction: Option<TextDirection>,
}

impl<'a> Row<'a> {
    /// Create a new row.
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            direction: None,
        }
    }

    /// Add a cell to the row.
    pub fn add_cell(&mut self, cell: Cell<'a>) {
        self.cells.push(cell);
    }

    /// Get the number of cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Get all cells.
    pub fn cells(&self) -> &[Cell<'a>] {
        &self.cells
    }

    /// Return the explicit row direction.
    pub fn direction(&self) -> Option<TextDirection> {
        self.direction
    }

    /// Set or clear the explicit row direction.
    pub fn set_direction(&mut self, direction: Option<TextDirection>) {
        self.direction = direction;
    }
}

impl<'a> Default for Row<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// A table cell.
#[derive(Debug, Clone)]
pub struct Cell<'a> {
    /// Cell text content
    text: Cow<'a, str>,
}

impl<'a> Cell<'a> {
    /// Create a new cell.
    pub fn new(text: Cow<'a, str>) -> Self {
        Self { text }
    }

    /// Get the cell text.
    pub fn text(&self) -> &str {
        &self.text
    }
}
