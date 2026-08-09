//! Checked BIFF8 row, column, and frozen-pane locations.
//!
//! These values model the ordinary BIFF8 worksheet grid: 65,536 zero-based
//! rows and 256 zero-based columns. They keep wide caller integers from being
//! silently narrowed while worksheet properties are staged for serialization.

use crate::{Error, Result};

/// A zero-based row in the BIFF8 worksheet grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Row(u16);

impl Row {
    /// Construct a checked zero-based BIFF8 row.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCellReference`] when `index` is outside the
    /// 65,536-row BIFF8 grid.
    #[must_use = "a checked row must be used or its validation error handled"]
    pub fn new(index: u32) -> Result<Self> {
        u16::try_from(index).map(Self).map_err(|_conversion_error| {
            Error::InvalidCellReference(format!("row {index} is outside the BIFF8 grid"))
        })
    }

    /// Return the zero-based row index.
    #[must_use]
    pub const fn index(self) -> u16 {
        self.0
    }
}

/// A zero-based column in the BIFF8 worksheet grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Column(u8);

impl Column {
    /// Construct a checked zero-based BIFF8 column.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCellReference`] when `index` is outside the
    /// 256-column BIFF8 grid.
    #[must_use = "a checked column must be used or its validation error handled"]
    pub fn new(index: u16) -> Result<Self> {
        u8::try_from(index).map(Self).map_err(|_conversion_error| {
            Error::InvalidCellReference(format!("column {index} is outside the BIFF8 grid"))
        })
    }

    /// Return the zero-based column index.
    #[must_use]
    pub const fn index(self) -> u8 {
        self.0
    }
}

/// Checked frozen-pane counts for a worksheet.
///
/// `Row` and `Column` are reused here because the `Pane` record represents
/// each frozen count in the same bounded BIFF8 domains. A zero/zero value is
/// a semantic request to clear frozen panes; [`crate::writer::Writer::unfreeze_panes`]
/// is the more direct spelling for that operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FrozenPanes {
    rows: Row,
    columns: Column,
}

impl FrozenPanes {
    /// Construct frozen-pane counts from checked BIFF8 grid values.
    #[must_use]
    pub const fn new(rows: Row, columns: Column) -> Self {
        Self { rows, columns }
    }

    /// Return the number of frozen rows.
    #[must_use]
    pub const fn rows(self) -> Row {
        self.rows
    }

    /// Return the number of frozen columns.
    #[must_use]
    pub const fn columns(self) -> Column {
        self.columns
    }

    pub(crate) const fn is_empty(self) -> bool {
        self.rows.index() == 0 && self.columns.index() == 0
    }
}

#[cfg(test)]
mod tests {
    use super::{Column, FrozenPanes, Row};

    #[test]
    fn grid_locations_accept_exact_biff8_bounds() {
        assert_eq!(Row::new(u32::from(u16::MAX)).unwrap().index(), u16::MAX);
        assert_eq!(Column::new(u16::from(u8::MAX)).unwrap().index(), u8::MAX);
    }

    #[test]
    fn grid_locations_reject_overflow_without_unwinding() {
        let result = std::panic::catch_unwind(|| {
            assert!(Row::new(u32::from(u16::MAX) + 1).is_err());
            assert!(Column::new(u16::from(u8::MAX) + 1).is_err());
        });
        assert!(result.is_ok());
    }

    #[test]
    fn frozen_panes_retain_checked_counts() {
        let panes = FrozenPanes::new(Row::new(7).unwrap(), Column::new(5).unwrap());
        assert_eq!(panes.rows().index(), 7);
        assert_eq!(panes.columns().index(), 5);
    }
}
