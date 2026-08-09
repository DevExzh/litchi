//! Checked semantic coordinate and range values.

use litchi_core::{Error, Result};

/// Largest zero-based coordinate that can be represented in A1 notation with
/// the platform's `usize` storage.
///
/// A1 rows are one-based and range dimensions are inclusive.  Reserving the
/// final `usize` value keeps both conversions and inclusive width/height
/// calculations representable without a late overflow.
pub const MAX_INDEX: usize = usize::MAX - 1;

/// A checked zero-based spreadsheet cell coordinate.
///
/// Columns and rows are zero-based in the semantic model and become `A1` and
/// `1` respectively at the textual boundary.  The private fields ensure that
/// safe callers cannot construct a value which the A1 codec cannot render.
///
/// # Examples
///
/// ```
/// use litchi_odf_common::coordinates::CellCoord;
///
/// let coord = CellCoord::new(0, 0).unwrap(); // A1
/// assert_eq!(coord.to_string(), "A1");
///
/// let coord = CellCoord::new(1, 2).unwrap(); // B3
/// assert_eq!(coord.to_string(), "B3");
///
/// let coord: CellCoord = "AA10".parse().unwrap();
/// assert_eq!(coord.column(), 26);
/// assert_eq!(coord.row(), 9);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellCoord {
    column: usize,
    row: usize,
}

impl CellCoord {
    /// Create a checked zero-based cell coordinate.
    ///
    /// Values above [`MAX_INDEX`] cannot be represented by the one-based A1
    /// row and inclusive range arithmetic used by this module.
    ///
    /// # Errors
    ///
    /// Returns an error if either coordinate exceeds [`MAX_INDEX`].
    pub fn new(column: usize, row: usize) -> Result<Self> {
        if column > MAX_INDEX {
            return Err(Error::InvalidFormat(format!(
                "column index {column} exceeds the checked coordinate limit {MAX_INDEX}"
            )));
        }
        if row > MAX_INDEX {
            return Err(Error::InvalidFormat(format!(
                "row index {row} exceeds the checked coordinate limit {MAX_INDEX}"
            )));
        }
        Ok(Self { column, row })
    }

    /// Get the checked zero-based column index.
    #[inline]
    #[must_use]
    pub const fn column(&self) -> usize {
        self.column
    }

    /// Get the checked zero-based row index.
    #[inline]
    #[must_use]
    pub const fn row(&self) -> usize {
        self.row
    }
}

/// A checked inclusive rectangular cell range.
///
/// Both endpoints are inclusive.  The start must not be below or to the
/// right of the end, so width and height are always meaningful and cannot
/// silently become zero for an inverted range.
///
/// # Examples
///
/// ```
/// use litchi_odf_common::coordinates::{CellCoord, CellRange};
///
/// let range = CellRange::new(
///     CellCoord::new(0, 0).unwrap(),
///     CellCoord::new(1, 2).unwrap(),
/// )
/// .unwrap();
/// assert_eq!(range.to_string(), "A1:B3");
///
/// let range: CellRange = "A1:B3".parse().unwrap();
/// assert_eq!(range.start().column(), 0);
/// assert_eq!(range.end().row(), 2);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    start: CellCoord,
    end: CellCoord,
}

impl CellRange {
    /// Create a checked inclusive range from its endpoints.
    ///
    /// # Errors
    ///
    /// Returns an error when the start is below or to the right of the end.
    pub fn new(start: CellCoord, end: CellCoord) -> Result<Self> {
        if start.column > end.column || start.row > end.row {
            return Err(Error::InvalidFormat(format!(
                "cell range start ({}, {}) must not follow end ({}, {})",
                start.column, start.row, end.column, end.row
            )));
        }
        Ok(Self { start, end })
    }

    /// Get the inclusive start coordinate.
    #[inline]
    #[must_use]
    pub const fn start(&self) -> CellCoord {
        self.start
    }

    /// Get the inclusive end coordinate.
    #[inline]
    #[must_use]
    pub const fn end(&self) -> CellCoord {
        self.end
    }

    /// Get the number of columns in the range.
    #[inline]
    #[must_use]
    pub fn width(&self) -> usize {
        self.end.column - self.start.column + 1
    }

    /// Get the number of rows in the range.
    #[inline]
    #[must_use]
    pub fn height(&self) -> usize {
        self.end.row - self.start.row + 1
    }

    /// Whether the range contains a checked coordinate.
    #[inline]
    #[must_use]
    pub const fn contains(&self, coordinate: CellCoord) -> bool {
        coordinate.column >= self.start.column
            && coordinate.column <= self.end.column
            && coordinate.row >= self.start.row
            && coordinate.row <= self.end.row
    }
}
