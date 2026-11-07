//! Traits for spreadsheet abstraction.
//!
//! The traits dispatch statically: each implementation names the concrete
//! cell, worksheet and iterator types it returns as associated types (generic
//! associated types borrowing from the implementation), so a read through the
//! traits allocates nothing for dispatch. Each implementation also names its
//! error type, so a format's typed error reaches the caller unchanged; the
//! formula engine classifies it into [`super::Error`]. Code generic over the traits takes
//! `W: WorkbookTrait` (or `S: Worksheet`) rather than a trait object; a
//! collection of different formats uses an enum over the concrete types.

use super::error::Result;
use super::types::CellValue;
use std::borrow::Cow;
use std::fmt::Debug;
use std::sync::Arc;

/// Represents an individual cell in a worksheet.
pub trait Cell: Send + Sync {
    /// Get the column number (1-based).
    fn column(&self) -> u32;

    /// Get the cell coordinate (e.g., "A1").
    fn coordinate(&self) -> String;

    /// Check if the cell contains a date/time value.
    fn is_date(&self) -> bool {
        matches!(self.value(), CellValue::DateTime(_))
    }

    /// Check if the cell is empty.
    fn is_empty(&self) -> bool {
        matches!(self.value(), CellValue::Empty)
    }

    /// Check if the cell contains a formula.
    fn is_formula(&self) -> bool {
        false // Default implementation, can be overridden
    }

    /// Get the row number (1-based).
    fn row(&self) -> u32;

    /// Get the cell value.
    fn value(&self) -> &CellValue;
}

/// Represents a worksheet (sheet) in a workbook.
pub trait Worksheet: Send + Sync {
    /// The error a read reports.
    type Error: std::error::Error + Send + Sync + 'static;

    /// The cell [`Self::cell`] returns and [`Self::cells`] yields, which may
    /// borrow from the worksheet.
    type Cell<'a>: Cell
    where
        Self: 'a;

    /// The iterator [`Self::cells`] returns.
    type Cells<'a>: Iterator<Item = Result<Self::Cell<'a>, Self::Error>>
    where
        Self: 'a;

    /// The iterator [`Self::rows`] returns: each row as a slice of cell
    /// values, wrapped in `Cow` for zero-copy when possible.
    type Rows<'a>: Iterator<Item = Result<Cow<'a, [CellValue]>, Self::Error>>
    where
        Self: 'a;

    /// Get a cell by row and column (1-based indexing).
    ///
    /// # Errors
    ///
    /// Returns an error if the cell cannot be read.
    fn cell(&self, row: u32, column: u32) -> Result<Self::Cell<'_>, Self::Error>;

    /// Get a cell by coordinate (e.g., "A1").
    ///
    /// # Errors
    ///
    /// Returns an error if the coordinate is invalid or the cell cannot be read.
    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Self::Cell<'_>, Self::Error>;

    /// Get cell value by row and column (1-based indexing).
    ///
    /// Returns a Cow to allow zero-copy when possible while supporting
    /// implementations that need to compute values (e.g., shared string resolution).
    ///
    /// # Errors
    ///
    /// Returns an error if the cell value cannot be resolved.
    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>, Self::Error>;

    /// Get all cells as an iterator.
    fn cells(&self) -> Self::Cells<'_>;

    /// Get the number of columns in the worksheet.
    fn column_count(&self) -> usize;

    /// Get inclusive one-based dimensions as (`min_row`, `min_col`, `max_row`, `max_col`).
    /// Returns None if the worksheet is empty.
    fn dimensions(&self) -> Option<(u32, u32, u32, u32)>;

    /// Get the worksheet name.
    fn name(&self) -> &str;

    /// Get a specific row by index (0-based).
    ///
    /// Returns a Cow to allow zero-copy when possible while supporting
    /// both owned and borrowed data depending on implementation.
    ///
    /// # Errors
    ///
    /// Returns an error if the row cannot be read.
    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>, Self::Error>;

    /// Get the number of rows in the worksheet.
    fn row_count(&self) -> usize;

    /// Get all rows as an iterator.
    fn rows(&self) -> Self::Rows<'_>;
}

/// Trait representing a workbook (spreadsheet file).
///
/// **Note**: This is the low-level trait API. For high-level usage, use the
/// unified `Workbook` struct from `crate::sheet::Workbook`.
pub trait WorkbookTrait: Debug + Send + Sync {
    /// The error a lookup or a worksheet read reports.
    type Error: std::error::Error + Send + Sync + 'static;

    /// The worksheet the lookups return and [`Self::worksheets`] yields,
    /// which may borrow from the workbook.
    type Worksheet<'a>: Worksheet<Error = Self::Error>
    where
        Self: 'a;

    /// The iterator [`Self::worksheets`] returns.
    type Worksheets<'a>: Iterator<Item = Result<Self::Worksheet<'a>, Self::Error>>
    where
        Self: 'a;

    /// Get the index of the active worksheet.
    fn active_sheet_index(&self) -> usize;

    /// Get the active worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error if the active worksheet cannot be determined.
    fn active_worksheet(&self) -> Result<Self::Worksheet<'_>, Self::Error>;

    /// Returns true if the workbook uses the 1904 date system (Mac).
    ///
    /// Implementations should report the correct setting based on workbook metadata.
    fn is_1904_date_system(&self) -> bool {
        false
    }

    /// Get a worksheet by index.
    ///
    /// # Errors
    ///
    /// Returns an error if the index is out of range.
    fn worksheet_by_index(&self, index: usize) -> Result<Self::Worksheet<'_>, Self::Error>;

    /// Get a worksheet by name.
    ///
    /// # Errors
    ///
    /// Returns an error if no worksheet with that name exists.
    fn worksheet_by_name(&self, name: &str) -> Result<Self::Worksheet<'_>, Self::Error>;

    /// Get the number of worksheets.
    fn worksheet_count(&self) -> usize;

    /// Get all worksheet names (zero-copy).
    ///
    /// Returns a slice reference to avoid cloning. Implementations
    /// should store worksheet names internally.
    fn worksheet_names(&self) -> &[String];

    /// Get all worksheets as an iterator.
    fn worksheets(&self) -> Self::Worksheets<'_>;
}

/// A borrowed cell reads as the cell.
impl<T: Cell + ?Sized> Cell for &T {
    fn column(&self) -> u32 {
        (**self).column()
    }

    fn coordinate(&self) -> String {
        (**self).coordinate()
    }

    fn is_date(&self) -> bool {
        (**self).is_date()
    }

    fn is_empty(&self) -> bool {
        (**self).is_empty()
    }

    fn is_formula(&self) -> bool {
        (**self).is_formula()
    }

    fn row(&self) -> u32 {
        (**self).row()
    }

    fn value(&self) -> &CellValue {
        (**self).value()
    }
}

/// A borrowed-or-owned cell reads as the cell, so a worksheet can return
/// stored cells by reference and absent ones by value.
impl<T: Cell + Clone> Cell for Cow<'_, T> {
    fn column(&self) -> u32 {
        self.as_ref().column()
    }

    fn coordinate(&self) -> String {
        self.as_ref().coordinate()
    }

    fn is_date(&self) -> bool {
        self.as_ref().is_date()
    }

    fn is_empty(&self) -> bool {
        self.as_ref().is_empty()
    }

    fn is_formula(&self) -> bool {
        self.as_ref().is_formula()
    }

    fn row(&self) -> u32 {
        self.as_ref().row()
    }

    fn value(&self) -> &CellValue {
        self.as_ref().value()
    }
}

/// Forward every worksheet read through a pointer to the worksheet.
macro_rules! forward_worksheet {
    () => {
        type Error = T::Error;
        type Cell<'a>
            = T::Cell<'a>
        where
            Self: 'a;
        type Cells<'a>
            = T::Cells<'a>
        where
            Self: 'a;
        type Rows<'a>
            = T::Rows<'a>
        where
            Self: 'a;

        fn cell(&self, row: u32, column: u32) -> Result<Self::Cell<'_>, T::Error> {
            (**self).cell(row, column)
        }

        fn cell_by_coordinate(&self, coordinate: &str) -> Result<Self::Cell<'_>, T::Error> {
            (**self).cell_by_coordinate(coordinate)
        }

        fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>, T::Error> {
            (**self).cell_value(row, column)
        }

        fn cells(&self) -> Self::Cells<'_> {
            (**self).cells()
        }

        fn column_count(&self) -> usize {
            (**self).column_count()
        }

        fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
            (**self).dimensions()
        }

        fn name(&self) -> &str {
            (**self).name()
        }

        fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>, T::Error> {
            (**self).row(row_idx)
        }

        fn row_count(&self) -> usize {
            (**self).row_count()
        }

        fn rows(&self) -> Self::Rows<'_> {
            (**self).rows()
        }
    };
}

/// A borrowed worksheet reads as the worksheet.
impl<T: Worksheet + ?Sized> Worksheet for &T {
    forward_worksheet!();
}

/// A shared worksheet snapshot reads as the worksheet.
impl<T: Worksheet + ?Sized> Worksheet for Arc<T> {
    forward_worksheet!();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Clone)]
    struct Fixed(CellValue);

    impl Cell for Fixed {
        fn column(&self) -> u32 {
            2
        }

        fn coordinate(&self) -> String {
            "B1".to_owned()
        }

        fn row(&self) -> u32 {
            1
        }

        fn value(&self) -> &CellValue {
            &self.0
        }
    }

    fn read<C: Cell>(cell: &C) -> (u32, u32, String, bool, bool) {
        (
            cell.row(),
            cell.column(),
            cell.coordinate(),
            cell.is_empty(),
            cell.is_date(),
        )
    }

    #[test]
    fn pointer_cells_forward_every_read() {
        let owned = Fixed(CellValue::DateTime(1.0));
        let expected = read(&owned);
        assert_eq!(read(&&owned), expected);
        assert_eq!(read(&Cow::Borrowed(&owned)), expected);
        assert_eq!(read(&Cow::<Fixed>::Owned(owned.clone())), expected);
        assert_eq!(Cow::Borrowed(&owned).value(), &CellValue::DateTime(1.0));
    }
}
