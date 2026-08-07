//! Format-neutral spreadsheet coordinates and selectors.
//!
//! Concrete workbook, worksheet, value, and formula types belong to XLS, XLSB,
//! XLSX, ODS, or other format crates. This crate contains only vocabulary that
//! can be shared without forcing one format to depend on another.

#![forbid(unsafe_code)]

use std::borrow::Cow;
use std::fmt;
use std::str::FromStr;

use litchi_core::Selector;
use thiserror::Error;

/// Number of rows in the modern Excel grid.
pub const ROWS: u32 = 1_048_576;
/// Number of columns in the modern Excel grid.
pub const COLUMNS: u32 = 16_384;

/// Zero-based spreadsheet row coordinate proven to be inside the grid.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Row(u32);

impl Row {
    /// First valid row.
    pub const FIRST: Self = Self(0);
    /// Last valid row.
    pub const LAST: Self = Self(ROWS - 1);

    /// Validate a zero-based row coordinate.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError::Row`] when `value` lies outside the grid.
    #[inline]
    pub const fn new(value: u32) -> Result<Self, CoordinateError> {
        if value < ROWS {
            Ok(Self(value))
        } else {
            Err(CoordinateError::Row { value })
        }
    }

    /// Return the zero-based coordinate.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    /// Next checked row, or `None` at the grid boundary.
    #[inline]
    #[must_use]
    pub const fn next(self) -> Option<Self> {
        if self.0 < Self::LAST.0 {
            Some(Self(self.0 + 1))
        } else {
            None
        }
    }

    /// Previous checked row, or `None` at the grid boundary.
    #[inline]
    #[must_use]
    pub const fn previous(self) -> Option<Self> {
        if self.0 > Self::FIRST.0 {
            Some(Self(self.0 - 1))
        } else {
            None
        }
    }
}

impl TryFrom<u32> for Row {
    type Error = CoordinateError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Row> for u32 {
    fn from(value: Row) -> Self {
        value.get()
    }
}

impl fmt::Display for Row {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}", self.get())
    }
}

/// Convenient row input accepted by format facades.
///
/// Raw indices are zero-based and checked when resolved. A reusable [`Row`]
/// avoids repeated validation in hot paths.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RowAt {
    /// A coordinate that has already been checked.
    Checked(Row),
    /// A raw zero-based index.
    Index(u32),
}

impl RowAt {
    /// Resolve this input into a checked row coordinate.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError::Row`] when the raw index lies outside the
    /// grid.
    #[inline]
    pub const fn resolve(self) -> Result<Row, CoordinateError> {
        match self {
            Self::Checked(row) => Ok(row),
            Self::Index(index) => Row::new(index),
        }
    }
}

impl From<Row> for RowAt {
    fn from(value: Row) -> Self {
        Self::Checked(value)
    }
}

impl From<u32> for RowAt {
    fn from(value: u32) -> Self {
        Self::Index(value)
    }
}

/// Zero-based spreadsheet column coordinate proven to be inside the grid.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Column(u32);

impl Column {
    /// First valid column.
    pub const FIRST: Self = Self(0);
    /// Last valid column.
    pub const LAST: Self = Self(COLUMNS - 1);

    /// Validate a zero-based column coordinate.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError::Column`] when `value` lies outside the grid.
    #[inline]
    pub const fn new(value: u32) -> Result<Self, CoordinateError> {
        if value < COLUMNS {
            Ok(Self(value))
        } else {
            Err(CoordinateError::Column { value })
        }
    }

    /// Return the zero-based coordinate.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    /// Parse one absolute or relative A1 column label.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError::ColumnA1`] when `reference` is not one
    /// in-grid A1 column label.
    pub fn from_a1(reference: &str) -> Result<Self, CoordinateError> {
        let bytes = reference.as_bytes();
        let start = usize::from(bytes.first() == Some(&b'$'));
        if start == bytes.len() || !bytes[start..].iter().all(u8::is_ascii_alphabetic) {
            return Err(CoordinateError::ColumnA1 {
                reference: reference.into(),
            });
        }
        let mut column = 0u32;
        for byte in &bytes[start..] {
            column = column
                .checked_mul(26)
                .and_then(|value| {
                    value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
                })
                .ok_or_else(|| CoordinateError::ColumnA1 {
                    reference: reference.into(),
                })?;
        }
        let zero_based = column
            .checked_sub(1)
            .ok_or_else(|| CoordinateError::ColumnA1 {
                reference: reference.into(),
            })?;
        if zero_based < COLUMNS {
            Ok(Self(zero_based))
        } else {
            Err(CoordinateError::ColumnA1 {
                reference: reference.into(),
            })
        }
    }

    /// Render this coordinate as an A1 column label.
    #[must_use]
    pub fn a1(self) -> String {
        let mut column = self.get() + 1;
        let mut reversed = String::with_capacity(3);
        while column != 0 {
            column -= 1;
            reversed.push(char::from(b'A' + (column % 26) as u8));
            column /= 26;
        }
        reversed.chars().rev().collect()
    }

    /// Next checked column, or `None` at the grid boundary.
    #[inline]
    #[must_use]
    pub const fn next(self) -> Option<Self> {
        if self.0 < Self::LAST.0 {
            Some(Self(self.0 + 1))
        } else {
            None
        }
    }

    /// Previous checked column, or `None` at the grid boundary.
    #[inline]
    #[must_use]
    pub const fn previous(self) -> Option<Self> {
        if self.0 > Self::FIRST.0 {
            Some(Self(self.0 - 1))
        } else {
            None
        }
    }
}

impl TryFrom<u32> for Column {
    type Error = CoordinateError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Column> for u32 {
    fn from(value: Column) -> Self {
        value.get()
    }
}

impl fmt::Display for Column {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}", self.get())
    }
}

/// Convenient column input accepted by format facades.
///
/// A1 labels such as `"B"` are the primary developer-facing form. Raw indices
/// are zero-based and checked when resolved, while a reusable [`Column`] avoids
/// repeated validation in hot paths.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ColumnAt<'a> {
    /// A coordinate that has already been checked.
    Checked(Column),
    /// A raw zero-based index.
    Index(u32),
    /// A borrowed or owned A1 column label.
    A1(Cow<'a, str>),
}

impl ColumnAt<'_> {
    /// Resolve this input into a checked column coordinate.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError`] when the raw index or A1 label does not
    /// resolve to an in-grid column.
    #[inline]
    pub fn resolve(self) -> Result<Column, CoordinateError> {
        match self {
            Self::Checked(column) => Ok(column),
            Self::Index(index) => Column::new(index),
            Self::A1(reference) => Column::from_a1(&reference),
        }
    }
}

impl From<Column> for ColumnAt<'_> {
    fn from(value: Column) -> Self {
        Self::Checked(value)
    }
}

impl From<u32> for ColumnAt<'_> {
    fn from(value: u32) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for ColumnAt<'a> {
    fn from(value: &'a str) -> Self {
        Self::A1(Cow::Borrowed(value))
    }
}

impl<'a> From<&'a String> for ColumnAt<'a> {
    fn from(value: &'a String) -> Self {
        Self::A1(Cow::Borrowed(value))
    }
}

impl From<String> for ColumnAt<'_> {
    fn from(value: String) -> Self {
        Self::A1(Cow::Owned(value))
    }
}

/// One checked, zero-based cell address.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Cell {
    row: Row,
    column: Column,
}

impl Cell {
    /// Combine coordinates that have already been validated.
    #[inline]
    #[must_use]
    pub const fn new(row: Row, column: Column) -> Self {
        Self { row, column }
    }

    /// Validate raw zero-based coordinates and create an address.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError`] when either coordinate lies outside the
    /// grid.
    #[inline]
    pub fn at(row: u32, column: u32) -> Result<Self, CoordinateError> {
        let checked_row = Row::new(row)?;
        let checked_column = Column::new(column)?;
        Ok(Self::new(checked_row, checked_column))
    }

    /// Parse one absolute or relative A1 cell reference.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError::A1`] when `reference` is not one in-grid A1
    /// cell reference.
    pub fn from_a1(reference: &str) -> Result<Self, CoordinateError> {
        let bytes = reference.as_bytes();
        let mut index = usize::from(bytes.first() == Some(&b'$'));
        let column_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
            index += 1;
        }
        if index == column_start {
            return Err(CoordinateError::A1 {
                reference: reference.into(),
            });
        }
        let mut column = 0u32;
        for byte in &bytes[column_start..index] {
            column = column
                .checked_mul(26)
                .and_then(|value| {
                    value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
                })
                .ok_or_else(|| CoordinateError::A1 {
                    reference: reference.into(),
                })?;
        }
        if bytes.get(index) == Some(&b'$') {
            index += 1;
        }
        let row_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == row_start || index != bytes.len() {
            return Err(CoordinateError::A1 {
                reference: reference.into(),
            });
        }
        let row = reference[row_start..]
            .parse::<u32>()
            .ok()
            .filter(|row| (1..=ROWS).contains(row))
            .ok_or_else(|| CoordinateError::A1 {
                reference: reference.into(),
            })?;
        if !(1..=COLUMNS).contains(&column) {
            return Err(CoordinateError::A1 {
                reference: reference.into(),
            });
        }
        Self::at(row - 1, column - 1)
    }

    /// Render this address in relative A1 notation.
    #[must_use]
    pub fn a1(self) -> String {
        let mut column = self.column.get() + 1;
        let mut reversed = String::with_capacity(3);
        while column != 0 {
            column -= 1;
            reversed.push(char::from(b'A' + (column % 26) as u8));
            column /= 26;
        }
        let mut value = String::with_capacity(reversed.len() + 7);
        value.extend(reversed.chars().rev());
        value.push_str(&(self.row.get() + 1).to_string());
        value
    }

    /// Zero-based row coordinate.
    #[inline]
    #[must_use]
    pub const fn row(self) -> Row {
        self.row
    }

    /// Zero-based column coordinate.
    #[inline]
    #[must_use]
    pub const fn column(self) -> Column {
        self.column
    }
}

impl TryFrom<(u32, u32)> for Cell {
    type Error = CoordinateError;

    fn try_from((row, column): (u32, u32)) -> Result<Self, Self::Error> {
        Self::at(row, column)
    }
}

impl fmt::Display for Cell {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.a1())
    }
}

impl From<(Row, Column)> for Cell {
    fn from((row, column): (Row, Column)) -> Self {
        Self::new(row, column)
    }
}

/// Convenient cell input accepted by format facades.
///
/// Callers normally pass an A1 reference. Raw `(row, column)` indices remain a
/// convenience, and a reusable checked [`Cell`] avoids repeated validation in
/// hot loops.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum At<'a> {
    /// A checked address.
    Address(Cell),
    /// Raw zero-based indices, validated when resolved.
    Indices { row: u32, column: u32 },
    /// A borrowed or owned A1 reference.
    A1(Cow<'a, str>),
}

impl At<'_> {
    /// Resolve this input into a checked address.
    ///
    /// # Errors
    ///
    /// Returns [`CoordinateError`] when the input does not resolve to an
    /// in-grid address.
    #[inline]
    pub fn resolve(self) -> Result<Cell, CoordinateError> {
        match self {
            Self::Address(address) => Ok(address),
            Self::Indices { row, column } => Cell::at(row, column),
            Self::A1(reference) => Cell::from_a1(&reference),
        }
    }
}

impl From<Cell> for At<'_> {
    fn from(value: Cell) -> Self {
        Self::Address(value)
    }
}

impl From<(Row, Column)> for At<'_> {
    fn from(value: (Row, Column)) -> Self {
        Self::Address(value.into())
    }
}

impl From<(u32, u32)> for At<'_> {
    fn from((row, column): (u32, u32)) -> Self {
        Self::Indices { row, column }
    }
}

impl<'a> From<&'a str> for At<'a> {
    fn from(value: &'a str) -> Self {
        Self::A1(Cow::Borrowed(value))
    }
}

impl<'a> From<&'a String> for At<'a> {
    fn from(value: &'a String) -> Self {
        Self::A1(Cow::Borrowed(value))
    }
}

impl From<String> for At<'_> {
    fn from(value: String) -> Self {
        Self::A1(Cow::Owned(value))
    }
}

/// Non-empty, zero-based, half-open rectangular range.
///
/// Exclusive end coordinates may equal [`ROWS`] or [`COLUMNS`], allowing the
/// final grid cell to be selected without making an out-of-grid [`Cell`]
/// representable.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rect {
    start: Cell,
    end_row: u32,
    end_column: u32,
}

impl Rect {
    /// The complete spreadsheet grid.
    pub const ALL: Self = Self {
        start: Cell::new(Row::FIRST, Column::FIRST),
        end_row: ROWS,
        end_column: COLUMNS,
    };

    /// Create a one-cell range.
    #[inline]
    #[must_use]
    pub const fn single(cell: Cell) -> Self {
        Self {
            start: cell,
            end_row: cell.row.get() + 1,
            end_column: cell.column.get() + 1,
        }
    }

    /// Create a range from a checked start and raw exclusive end coordinates.
    ///
    /// # Errors
    ///
    /// Returns [`RangeError::EndOutsideGrid`] when an exclusive end exceeds the
    /// grid boundary, or [`RangeError::EmptyOrInverted`] when the end does not
    /// lie below and to the right of the start.
    pub const fn new(start: Cell, end_row: u32, end_column: u32) -> Result<Self, RangeError> {
        if end_row > ROWS || end_column > COLUMNS {
            return Err(RangeError::EndOutsideGrid {
                row: end_row,
                column: end_column,
            });
        }
        if end_row <= start.row.get() || end_column <= start.column.get() {
            return Err(RangeError::EmptyOrInverted {
                start,
                end_row,
                end_column,
            });
        }
        Ok(Self {
            start,
            end_row,
            end_column,
        })
    }

    /// Validate raw start and exclusive end coordinates.
    ///
    /// # Errors
    ///
    /// Returns [`RangeError::Start`] when the start lies outside the grid, or
    /// the errors of [`Rect::new`] when the end is invalid.
    pub fn at(
        start_row: u32,
        start_column: u32,
        end_row: u32,
        end_column: u32,
    ) -> Result<Self, RangeError> {
        let start = Cell::at(start_row, start_column).map_err(RangeError::Start)?;
        Self::new(start, end_row, end_column)
    }

    /// Parse one bounded A1 cell or rectangular range.
    ///
    /// # Errors
    ///
    /// Returns [`RangeError::A1`] when `reference` is not one bounded A1 cell
    /// or rectangular range, or the errors of [`Rect::new`] when the parsed
    /// range is empty or inverted.
    pub fn from_a1(reference: &str) -> Result<Self, RangeError> {
        let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
        if last.contains(':') {
            return Err(RangeError::A1 {
                reference: reference.into(),
                source: CoordinateError::A1 {
                    reference: reference.into(),
                },
            });
        }
        let first_cell = Cell::from_a1(first).map_err(|source| RangeError::A1 {
            reference: reference.into(),
            source,
        })?;
        let last_cell = Cell::from_a1(last).map_err(|source| RangeError::A1 {
            reference: reference.into(),
            source,
        })?;
        Self::new(
            first_cell,
            last_cell.row.get() + 1,
            last_cell.column.get() + 1,
        )
    }

    /// Checked inclusive start address.
    #[inline]
    #[must_use]
    pub const fn start(self) -> Cell {
        self.start
    }

    /// Raw zero-based exclusive end coordinates.
    #[inline]
    #[must_use]
    pub const fn end(self) -> (u32, u32) {
        (self.end_row, self.end_column)
    }

    /// Number of selected rows.
    #[inline]
    #[must_use]
    pub const fn rows(self) -> u32 {
        self.end_row - self.start.row.get()
    }

    /// Number of selected columns.
    #[inline]
    #[must_use]
    pub const fn columns(self) -> u32 {
        self.end_column - self.start.column.get()
    }

    /// Whether the range contains an address.
    #[inline]
    #[must_use]
    pub const fn contains(self, address: Cell) -> bool {
        address.row.get() >= self.start.row.get()
            && address.row.get() < self.end_row
            && address.column.get() >= self.start.column.get()
            && address.column.get() < self.end_column
    }

    /// Smallest rectangle containing both inputs.
    #[inline]
    #[must_use]
    pub const fn union(self, other: Self) -> Self {
        let start_row = if self.start.row.get() < other.start.row.get() {
            self.start.row
        } else {
            other.start.row
        };
        let start_column = if self.start.column.get() < other.start.column.get() {
            self.start.column
        } else {
            other.start.column
        };
        Self {
            start: Cell::new(start_row, start_column),
            end_row: if self.end_row > other.end_row {
                self.end_row
            } else {
                other.end_row
            },
            end_column: if self.end_column > other.end_column {
                self.end_column
            } else {
                other.end_column
            },
        }
    }

    /// Render this range in compact relative A1 notation.
    #[must_use]
    pub fn a1(self) -> String {
        if self.rows() == 1 && self.columns() == 1 {
            return self.start.a1();
        }
        // `Rect::new` proves both exclusive ends are inside the grid and
        // strictly beyond the start, so their predecessors are valid cells.
        let last = Cell::new(Row(self.end_row - 1), Column(self.end_column - 1));
        format!("{}:{last}", self.start.a1())
    }
}

impl fmt::Display for Rect {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.a1())
    }
}

impl FromStr for Rect {
    type Err = RangeError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::from_a1(value)
    }
}

/// Convenient rectangular-range input accepted by format facades.
///
/// A1 notation is the primary entry point. Raw zero-based half-open bounds
/// remain available, while a checked [`Rect`] avoids repeated validation.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Area<'a> {
    /// A checked rectangle.
    Range(Rect),
    /// Raw zero-based half-open bounds.
    Bounds {
        start_row: u32,
        start_column: u32,
        end_row: u32,
        end_column: u32,
    },
    /// A borrowed or owned A1 cell/range reference.
    A1(Cow<'a, str>),
}

impl Area<'_> {
    /// Resolve this input into a checked rectangle.
    ///
    /// # Errors
    ///
    /// Returns [`RangeError`] when the input does not resolve to a valid
    /// rectangle.
    #[inline]
    pub fn resolve(self) -> Result<Rect, RangeError> {
        match self {
            Self::Range(range) => Ok(range),
            Self::Bounds {
                start_row,
                start_column,
                end_row,
                end_column,
            } => Rect::at(start_row, start_column, end_row, end_column),
            Self::A1(reference) => Rect::from_a1(&reference),
        }
    }
}

impl From<Rect> for Area<'_> {
    fn from(value: Rect) -> Self {
        Self::Range(value)
    }
}

impl From<(u32, u32, u32, u32)> for Area<'_> {
    fn from((start_row, start_column, end_row, end_column): (u32, u32, u32, u32)) -> Self {
        Self::Bounds {
            start_row,
            start_column,
            end_row,
            end_column,
        }
    }
}

impl<'a> From<&'a str> for Area<'a> {
    fn from(value: &'a str) -> Self {
        Self::A1(Cow::Borrowed(value))
    }
}

impl<'a> From<&'a String> for Area<'a> {
    fn from(value: &'a String) -> Self {
        Self::A1(Cow::Borrowed(value))
    }
}

impl From<String> for Area<'_> {
    fn from(value: String) -> Self {
        Self::A1(Cow::Owned(value))
    }
}

/// Invalid spreadsheet coordinate.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum CoordinateError {
    /// The row lies outside `0..ROWS`.
    #[error("row {value} is outside the zero-based spreadsheet grid 0..{ROWS}")]
    Row { value: u32 },
    /// The column lies outside `0..COLUMNS`.
    #[error("column {value} is outside the zero-based spreadsheet grid 0..{COLUMNS}")]
    Column { value: u32 },
    /// The string is not one bounded A1 column label.
    #[error("invalid or out-of-grid A1 column label '{reference}'")]
    ColumnA1 { reference: String },
    /// The string is not one bounded A1 cell reference.
    #[error("invalid or out-of-grid A1 cell reference '{reference}'")]
    A1 { reference: String },
}

/// Invalid half-open rectangle.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum RangeError {
    /// The inclusive start is outside the grid.
    #[error("invalid range start: {0}")]
    Start(CoordinateError),
    /// The string is not one bounded A1 cell or rectangular range.
    #[error("invalid A1 range '{reference}': {source}")]
    A1 {
        reference: String,
        #[source]
        source: CoordinateError,
    },
    /// The exclusive end exceeds the grid boundary.
    #[error("range end ({row}, {column}) exceeds exclusive grid boundary ({ROWS}, {COLUMNS})")]
    EndOutsideGrid { row: u32, column: u32 },
    /// A half-open range must have positive width and height.
    #[error(
        "range end ({end_row}, {end_column}) must be below and to the right of start {start:?}"
    )]
    EmptyOrInverted {
        start: Cell,
        end_row: u32,
        end_column: u32,
    },
}

/// Convenient sheet selector used by concrete workbook crates.
pub type SheetSelector<'a, Id> = Selector<'a, Id>;

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        reason = "tests deliberately panic on unexpected parse and validation failures"
    )]

    use super::*;

    #[test]
    fn coordinates_make_out_of_grid_cells_unrepresentable() {
        assert_eq!(Cell::at(0, 0), Ok(Cell::new(Row::FIRST, Column::FIRST)));
        assert_eq!(RowAt::from(0).resolve(), Ok(Row::FIRST));
        assert_eq!(RowAt::from(Row::LAST).resolve(), Ok(Row::LAST));
        assert_eq!(ColumnAt::from(0).resolve(), Ok(Column::FIRST));
        assert_eq!(ColumnAt::from(Column::LAST).resolve(), Ok(Column::LAST));
        assert_eq!(ColumnAt::from("B").resolve(), Column::new(1));
        assert_eq!(ColumnAt::from("$xfd").resolve(), Ok(Column::LAST));
        assert!(matches!(
            RowAt::from(ROWS).resolve(),
            Err(CoordinateError::Row { .. })
        ));
        assert!(matches!(
            ColumnAt::from(COLUMNS).resolve(),
            Err(CoordinateError::Column { .. })
        ));
        assert_eq!(
            Cell::at(ROWS - 1, COLUMNS - 1).ok().map(Cell::row),
            Some(Row::LAST)
        );
        assert!(matches!(
            Cell::at(ROWS, 0),
            Err(CoordinateError::Row { .. })
        ));
        assert!(matches!(
            Cell::at(0, COLUMNS),
            Err(CoordinateError::Column { .. })
        ));
    }

    #[test]
    fn a1_is_a_semantic_checked_selector() {
        let last = Cell::from_a1("$XFD$1048576").expect("last cell");
        assert_eq!(last, Cell::new(Row::LAST, Column::LAST));
        assert_eq!(last.a1(), "XFD1048576");
        assert_eq!(Cell::from_a1("aa42").map(Cell::a1).as_deref(), Ok("AA42"));
        for invalid in ["", "A", "1", "A0", "XFE1", "A1048577", "A1:B2"] {
            assert!(Cell::from_a1(invalid).is_err(), "accepted {invalid}");
        }

        assert_eq!(Column::from_a1("aa").map(Column::a1).as_deref(), Ok("AA"));
        assert_eq!(Column::from_a1("$XFD"), Ok(Column::LAST));
        for invalid in ["", "$", "1", "A1", "XFE", "A:B"] {
            assert!(Column::from_a1(invalid).is_err(), "accepted {invalid}");
        }
    }

    #[test]
    fn ranges_are_zero_based_half_open_and_can_cover_the_full_grid() {
        let range = Rect::at(0, 1, 3, 5).expect("valid rectangle");
        assert_eq!(range.rows(), 3);
        assert_eq!(range.columns(), 4);
        assert!(range.contains(Cell::at(2, 4).expect("valid address")));
        assert!(!range.contains(Cell::at(2, 5).expect("valid address")));
        assert_eq!(Rect::ALL.end(), (ROWS, COLUMNS));
        assert!(Rect::ALL.contains(Cell::at(ROWS - 1, COLUMNS - 1).expect("last cell")));
    }

    #[test]
    fn a1_areas_are_checked_compact_and_selector_friendly() {
        let range = Rect::from_a1("$B$2:D4").expect("bounded area");
        assert_eq!(range.start(), Cell::at(1, 1).expect("start"));
        assert_eq!(range.end(), (4, 4));
        assert_eq!(range.a1(), "B2:D4");
        assert_eq!(
            Rect::from_a1("XFD1048576").map(Rect::a1).as_deref(),
            Ok("XFD1048576")
        );

        let expanded = range.union(Rect::single(Cell::from_a1("A9").expect("cell")));
        assert_eq!(expanded.to_string(), "A2:D9");
        assert_eq!(Area::from("B2:D4").resolve(), Ok(range));
        assert_eq!(Area::from((1, 1, 4, 4)).resolve(), Ok(range));

        for invalid in ["", "A", "A0", "XFE1", "A1:B2:C3", "B2:A1"] {
            assert!(Rect::from_a1(invalid).is_err(), "accepted {invalid}");
        }
    }

    #[test]
    fn empty_inverted_and_out_of_grid_rectangles_are_rejected() {
        assert!(matches!(
            Rect::at(2, 2, 2, 3),
            Err(RangeError::EmptyOrInverted { .. })
        ));
        assert!(matches!(
            Rect::at(2, 2, 3, 1),
            Err(RangeError::EmptyOrInverted { .. })
        ));
        assert!(matches!(
            Rect::at(0, 0, ROWS + 1, 1),
            Err(RangeError::EndOutsideGrid { .. })
        ));
    }
}
