//! Compact, archive-free spreadsheet coordinate selectors.
//!
//! The semantic table uses zero-based coordinates and half-open ranges. A1
//! parsing is performed directly over the caller's bytes so a selector never
//! allocates an intermediate column label or address string.

use std::fmt;
use std::fmt::Write as _;

const MAX_ONE_BASED_COMPONENT: u64 = u32::MAX as u64 + 1;

/// The reason an A1 selector could not be parsed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AddressError {
    /// The selector contains no bytes.
    Empty,
    /// A column label is required before the row number.
    MissingColumn,
    /// The column label contains a non-letter where a label is expected.
    InvalidColumn,
    /// The column label exceeds the compact `u32` coordinate domain.
    ColumnOverflow,
    /// A row number is required after the column label.
    MissingRow,
    /// The row component contains a non-decimal character.
    InvalidRow,
    /// A row number of zero is not a valid one-based A1 row.
    ZeroRow,
    /// The row component exceeds the compact `u32` coordinate domain.
    RowOverflow,
    /// A second range separator was found.
    MultipleSeparators,
    /// A range endpoint cannot be converted to a half-open bound.
    RangeOverflow,
    /// Unexpected bytes follow a complete A1 selector.
    TrailingInput,
}

impl fmt::Display for AddressError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let message = match self {
            Self::Empty => "address is empty",
            Self::MissingColumn => "address is missing its column label",
            Self::InvalidColumn => "address has an invalid column label",
            Self::ColumnOverflow => "address column exceeds the compact coordinate domain",
            Self::MissingRow => "address is missing its row number",
            Self::InvalidRow => "address has an invalid row number",
            Self::ZeroRow => "address row numbers start at one",
            Self::RowOverflow => "address row exceeds the compact coordinate domain",
            Self::MultipleSeparators => "range contains more than one separator",
            Self::RangeOverflow => "range endpoint exceeds the compact coordinate domain",
            Self::TrailingInput => "address contains trailing input",
        };
        formatter.write_str(message)
    }
}

/// Failure while constructing or parsing a compact coordinate selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// The range bounds are inverted on at least one axis.
    InvalidRange {
        /// Inclusive start coordinate.
        start: CellPosition,
        /// Exclusive end coordinate.
        end: CellPosition,
    },
    /// A platform-sized coordinate does not fit the compact representation.
    CoordinateOverflow {
        /// Supplied row coordinate.
        row: usize,
        /// Supplied column coordinate.
        column: usize,
    },
    /// An A1 selector is malformed.
    InvalidAddress {
        /// Syntax or domain failure.
        kind: AddressError,
        /// Byte offset at which parsing stopped.
        index: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidRange { start, end } => write!(
                formatter,
                "table range ({}, {})..({}, {}) is inverted",
                start.row(),
                start.column(),
                end.row(),
                end.column()
            ),
            Self::CoordinateOverflow { row, column } => {
                write!(
                    formatter,
                    "table coordinate ({row}, {column}) overflows u32"
                )
            },
            Self::InvalidAddress { kind, index } => {
                write!(formatter, "invalid A1 address at byte {index}: {kind}")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for coordinate construction and parsing.
pub type Result<T> = std::result::Result<T, Error>;

/// A compact zero-based cell coordinate.
///
/// The value is exactly eight bytes and contains no native object or package
/// identifier. Use [`Self::from_a1`] for the human-readable selector form.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellPosition {
    pub(crate) row: u32,
    pub(crate) column: u32,
}

impl CellPosition {
    /// Creates a zero-based coordinate from compact components.
    #[must_use]
    pub const fn new(row: u32, column: u32) -> Self {
        Self { row, column }
    }

    /// Converts platform-sized coordinates without truncation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::CoordinateOverflow`] if either component does not fit
    /// in the compact representation.
    pub fn try_from_usize(row: usize, column: usize) -> Result<Self> {
        let compact_row =
            u32::try_from(row).map_err(|_error| Error::CoordinateOverflow { row, column })?;
        let compact_column =
            u32::try_from(column).map_err(|_error| Error::CoordinateOverflow { row, column })?;
        Ok(Self::new(compact_row, compact_column))
    }

    /// Parses a zero-based coordinate from an A1 selector.
    ///
    /// Both relative (`B3`) and absolute (`$B$3`, `$B3`, `B$3`) markers are
    /// accepted. Markers affect neither the semantic coordinate nor its
    /// compact representation. Parsing borrows the input and performs no
    /// allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed error for malformed syntax or a component outside the
    /// compact `u32` domain.
    pub fn from_a1(address: &str) -> Result<Self> {
        let bytes = address.as_bytes();
        if bytes.is_empty() {
            return Err(invalid(AddressError::Empty, 0));
        }

        let mut cursor = 0;
        if bytes[cursor] == b'$' {
            cursor += 1;
            if cursor == bytes.len() {
                return Err(invalid(AddressError::MissingColumn, cursor));
            }
        }

        let column_start = cursor;
        let mut one_based_column = 0_u64;
        while cursor < bytes.len() && is_ascii_letter(bytes[cursor]) {
            let digit = u64::from(column_digit(bytes[cursor]));
            let next = one_based_column
                .checked_mul(26)
                .and_then(|value| value.checked_add(digit))
                .ok_or_else(|| invalid(AddressError::ColumnOverflow, cursor))?;
            if next > MAX_ONE_BASED_COMPONENT {
                return Err(invalid(AddressError::ColumnOverflow, cursor));
            }
            one_based_column = next;
            cursor += 1;
        }
        if cursor == column_start {
            return Err(invalid(AddressError::InvalidColumn, cursor));
        }

        if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
        }
        let row_start = cursor;
        let mut one_based_row = 0_u64;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            let digit = u64::from(bytes[cursor] - b'0');
            let next = one_based_row
                .checked_mul(10)
                .and_then(|value| value.checked_add(digit))
                .ok_or_else(|| invalid(AddressError::RowOverflow, cursor))?;
            if next > MAX_ONE_BASED_COMPONENT {
                return Err(invalid(AddressError::RowOverflow, cursor));
            }
            one_based_row = next;
            cursor += 1;
        }
        if cursor == row_start {
            return Err(invalid(AddressError::MissingRow, cursor));
        }
        if one_based_row == 0 {
            return Err(invalid(AddressError::ZeroRow, row_start));
        }
        if cursor != bytes.len() {
            return Err(invalid(AddressError::TrailingInput, cursor));
        }

        let row = u32::try_from(one_based_row - 1)
            .map_err(|_error| invalid(AddressError::RowOverflow, row_start))?;
        let column = u32::try_from(one_based_column - 1)
            .map_err(|_error| invalid(AddressError::ColumnOverflow, column_start))?;
        Ok(Self::new(row, column))
    }

    /// Returns the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Returns the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

impl fmt::Display for CellPosition {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut one_based_column = u64::from(self.column) + 1;
        let mut labels = [0_u8; 7];
        let mut length = 0;
        while one_based_column != 0 {
            let digit = ((one_based_column - 1) % 26) as u8;
            labels[length] = b'A' + digit;
            length += 1;
            one_based_column = (one_based_column - 1) / 26;
        }
        for digit in labels[..length].iter().rev() {
            formatter.write_char(char::from(*digit))?;
        }
        write!(formatter, "{}", u64::from(self.row) + 1)
    }
}

/// A zero-based, half-open rectangular cell range.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    pub(crate) start: CellPosition,
    pub(crate) end: CellPosition,
}

impl CellRange {
    /// Creates a half-open range without checking a table's declared extent.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidRange`] when the end precedes the start on
    /// either axis.
    pub fn new(start: CellPosition, end: CellPosition) -> Result<Self> {
        if start.row > end.row || start.column > end.column {
            return Err(Error::InvalidRange { start, end });
        }
        Ok(Self { start, end })
    }

    /// Parses an A1 cell or inclusive A1 range into a half-open range.
    ///
    /// `B3` selects one cell. `B3:D5` selects the inclusive rectangle from
    /// `B3` through `D5`; its semantic end is therefore `E6`. Parsing is
    /// allocation-free and accepts the same absolute markers as
    /// [`CellPosition::from_a1`].
    ///
    /// # Errors
    ///
    /// Returns a typed error for malformed endpoints, multiple separators, or
    /// an endpoint that cannot be advanced to an exclusive bound.
    pub fn from_a1(address: &str) -> Result<Self> {
        let Some(separator) = address.as_bytes().iter().position(|byte| *byte == b':') else {
            let position = CellPosition::from_a1(address)?;
            return Self::single_with_index(position, 0);
        };
        if let Some(extra) = address.as_bytes()[separator + 1..]
            .iter()
            .position(|byte| *byte == b':')
        {
            return Err(invalid(
                AddressError::MultipleSeparators,
                separator + 1 + extra,
            ));
        }

        let start = CellPosition::from_a1(&address[..separator])?;
        let end_inclusive = CellPosition::from_a1(&address[separator + 1..])?;
        let end = exclusive_end(end_inclusive, separator + 1)?;
        Self::new(start, end)
    }

    /// Creates a one-cell half-open range.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidAddress`] if advancing the coordinate to the
    /// exclusive end would overflow the compact domain.
    pub fn single(position: CellPosition) -> Result<Self> {
        Self::single_with_index(position, 0)
    }

    /// Returns the inclusive start and exclusive end coordinates.
    #[must_use]
    pub const fn bounds(self) -> (CellPosition, CellPosition) {
        (self.start, self.end)
    }

    /// Returns the first coordinate.
    #[must_use]
    pub const fn start(self) -> CellPosition {
        self.start
    }

    /// Returns the exclusive end coordinate.
    #[must_use]
    pub const fn end(self) -> CellPosition {
        self.end
    }

    /// Returns whether the range contains no cells.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start.row == self.end.row || self.start.column == self.end.column
    }

    /// Returns the number of cells when it fits in `usize`.
    #[must_use]
    pub const fn area(self) -> Option<usize> {
        let rows = (self.end.row - self.start.row) as usize;
        let columns = (self.end.column - self.start.column) as usize;
        rows.checked_mul(columns)
    }

    /// Returns whether a coordinate belongs to the range.
    #[must_use]
    pub const fn contains(self, position: CellPosition) -> bool {
        position.row >= self.start.row
            && position.row < self.end.row
            && position.column >= self.start.column
            && position.column < self.end.column
    }

    fn single_with_index(position: CellPosition, error_index: usize) -> Result<Self> {
        let end = exclusive_end(position, error_index)?;
        Self::new(position, end)
    }
}

fn exclusive_end(position: CellPosition, error_index: usize) -> Result<CellPosition> {
    let row = position
        .row
        .checked_add(1)
        .ok_or_else(|| invalid(AddressError::RangeOverflow, error_index))?;
    let column = position
        .column
        .checked_add(1)
        .ok_or_else(|| invalid(AddressError::RangeOverflow, error_index))?;
    Ok(CellPosition::new(row, column))
}

fn invalid(kind: AddressError, index: usize) -> Error {
    Error::InvalidAddress { kind, index }
}

fn is_ascii_letter(byte: u8) -> bool {
    byte.is_ascii_alphabetic()
}

fn column_digit(byte: u8) -> u8 {
    match byte {
        b'A'..=b'Z' => byte - b'A' + 1,
        b'a'..=b'z' => byte - b'a' + 1,
        _ => 0,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn coordinates_are_compact_and_round_trip_a1_without_allocation() {
        assert_eq!(size_of::<CellPosition>(), 8);
        assert_eq!(size_of::<CellRange>(), 16);

        let position = CellPosition::from_a1("$bc$12")
            .unwrap_or_else(|error| panic!("unexpected coordinate error: {error}"));
        assert_eq!(position, CellPosition::new(11, 54));
        assert_eq!(position.to_string(), "BC12");
    }

    #[test]
    fn ranges_accept_single_cells_and_expand_inclusive_a1_endpoints() {
        let single = CellRange::from_a1("C4")
            .unwrap_or_else(|error| panic!("unexpected range error: {error}"));
        assert_eq!(
            single.bounds(),
            (CellPosition::new(3, 2), CellPosition::new(4, 3))
        );

        let range = CellRange::from_a1("B3:D5")
            .unwrap_or_else(|error| panic!("unexpected range error: {error}"));
        assert_eq!(range.start(), CellPosition::new(2, 1));
        assert_eq!(range.end(), CellPosition::new(5, 4));
        assert_eq!(range.area(), Some(9));
        assert!(range.contains(CellPosition::new(4, 3)));
        assert!(!range.contains(CellPosition::new(5, 3)));
    }

    #[test]
    fn malformed_a1_selectors_return_precise_typed_errors() {
        assert!(matches!(
            CellPosition::from_a1("A0"),
            Err(Error::InvalidAddress {
                kind: AddressError::ZeroRow,
                ..
            })
        ));
        assert!(matches!(
            CellPosition::from_a1("A1!"),
            Err(Error::InvalidAddress {
                kind: AddressError::TrailingInput,
                index: 2
            })
        ));
        assert!(matches!(
            CellRange::from_a1("A1:B2:C3"),
            Err(Error::InvalidAddress {
                kind: AddressError::MultipleSeparators,
                index: 5
            })
        ));
        assert!(matches!(
            CellPosition::try_from_usize(usize::MAX, 0),
            Err(Error::CoordinateOverflow { .. })
        ));
    }
}
