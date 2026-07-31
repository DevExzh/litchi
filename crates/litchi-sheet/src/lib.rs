//! Format-neutral spreadsheet coordinates and selectors.
//!
//! Concrete workbook, worksheet, cell, and formula handles belong to XLS,
//! XLSB, XLSX, ODS, or other format crates. This crate contains only vocabulary
//! that can be shared without forcing one format to depend on another.

#![forbid(unsafe_code)]

use litchi_core::Selector;
use thiserror::Error;

/// Zero-based spreadsheet row coordinate.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Row(u32);

impl Row {
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    pub const fn get(self) -> u32 {
        self.0
    }
}

/// Zero-based spreadsheet column coordinate.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Column(u32);

impl Column {
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    pub const fn get(self) -> u32 {
        self.0
    }
}

/// One zero-based cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Cell {
    pub row: Row,
    pub column: Column,
}

impl Cell {
    pub const fn new(row: u32, column: u32) -> Self {
        Self {
            row: Row::new(row),
            column: Column::new(column),
        }
    }
}

/// Non-empty, zero-based, half-open rectangular range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rect {
    start: Cell,
    end: Cell,
}

impl Rect {
    pub fn new(start: Cell, end: Cell) -> Result<Self, RangeError> {
        if end.row <= start.row || end.column <= start.column {
            return Err(RangeError { start, end });
        }
        Ok(Self { start, end })
    }

    pub const fn start(self) -> Cell {
        self.start
    }

    pub const fn end(self) -> Cell {
        self.end
    }

    pub const fn rows(self) -> u32 {
        self.end.row.get() - self.start.row.get()
    }

    pub const fn columns(self) -> u32 {
        self.end.column.get() - self.start.column.get()
    }
}

/// Invalid half-open rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[error("range end {end:?} must be below and to the right of start {start:?}")]
pub struct RangeError {
    pub start: Cell,
    pub end: Cell,
}

/// Convenient sheet selector used by concrete workbook crates.
pub type SheetSelector<'a, Id> = Selector<'a, Id>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges_are_zero_based_and_half_open() {
        let range = Rect::new(Cell::new(0, 1), Cell::new(3, 5)).expect("valid rectangle");
        assert_eq!(range.rows(), 3);
        assert_eq!(range.columns(), 4);
    }

    #[test]
    fn empty_or_inverted_rectangles_are_rejected() {
        assert!(Rect::new(Cell::new(2, 2), Cell::new(2, 3)).is_err());
        assert!(Rect::new(Cell::new(2, 2), Cell::new(3, 1)).is_err());
    }
}
