//! Decimal Number display formats for existing Keynote table cells.
//!
//! The Keynote-local semantic value mirrors the neutral iWork Number domain.
//! Native table identifiers,
//! format-list keys, BNC records, and archive bytes remain private to the
//! Keynote package adapter.

use std::fmt;
use std::fmt::Write as _;

/// A checked semantic table-cell coordinate.
///
/// Coordinates are zero based and contain no package or native object
/// identifier. The package adapter checks the coordinate against the selected
/// table's declared extent before it inspects any cell storage.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellPosition {
    row: u32,
    column: u32,
}

impl CellPosition {
    /// Constructs a compact zero-based coordinate.
    #[must_use]
    pub const fn new(row: u32, column: u32) -> Self {
        Self { row, column }
    }

    /// Converts platform-sized coordinates without truncation.
    pub fn try_from_usize(row: usize, column: usize) -> Result<Self, CellPositionError> {
        Ok(Self::new(
            u32::try_from(row)
                .map_err(|_| CellPositionError::CoordinateOverflow { row, column })?,
            u32::try_from(column)
                .map_err(|_| CellPositionError::CoordinateOverflow { row, column })?,
        ))
    }

    /// Parses a zero-based coordinate from a one-based A1 address.
    ///
    /// Relative and absolute markers ($B$2, $B2, and B$2) are accepted.
    /// Parsing is allocation-free.
    pub fn from_a1(address: &str) -> Result<Self, CellPositionError> {
        let bytes = address.as_bytes();
        if bytes.is_empty() {
            return Err(CellPositionError::InvalidAddress);
        }
        let mut cursor = 0usize;
        if bytes[cursor] == b'$' {
            cursor += 1;
        }
        let column_start = cursor;
        let mut one_based_column = 0u64;
        while cursor < bytes.len() && bytes[cursor].is_ascii_alphabetic() {
            let byte = bytes[cursor].to_ascii_uppercase();
            let digit = u64::from(byte - b'A' + 1);
            one_based_column = one_based_column
                .checked_mul(26)
                .and_then(|value| value.checked_add(digit))
                .ok_or(CellPositionError::CoordinateOverflow {
                    row: 0,
                    column: usize::MAX,
                })?;
            cursor += 1;
        }
        if cursor == column_start || one_based_column == 0 {
            return Err(CellPositionError::InvalidAddress);
        }
        if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
        }
        let row_start = cursor;
        let mut one_based_row = 0u64;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            one_based_row = one_based_row
                .checked_mul(10)
                .and_then(|value| value.checked_add(u64::from(bytes[cursor] - b'0')))
                .ok_or(CellPositionError::CoordinateOverflow {
                    row: usize::MAX,
                    column: 0,
                })?;
            cursor += 1;
        }
        if cursor == row_start || one_based_row == 0 || cursor != bytes.len() {
            return Err(CellPositionError::InvalidAddress);
        }
        let row = u32::try_from(one_based_row - 1).map_err(|_| {
            CellPositionError::CoordinateOverflow {
                row: usize::MAX,
                column: 0,
            }
        })?;
        let column = u32::try_from(one_based_column - 1).map_err(|_| {
            CellPositionError::CoordinateOverflow {
                row: 0,
                column: usize::MAX,
            }
        })?;
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
        let mut column = u64::from(self.column) + 1;
        let mut labels = [0u8; 7];
        let mut length = 0usize;
        while column != 0 {
            labels[length] = b'A' + ((column - 1) % 26) as u8;
            length += 1;
            column = (column - 1) / 26;
        }
        for byte in labels[..length].iter().rev() {
            formatter.write_char(char::from(*byte))?;
        }
        write!(formatter, "{}", u64::from(self.row) + 1)
    }
}

/// Checked coordinate-construction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellPositionError {
    /// A platform-sized coordinate does not fit in the compact form.
    CoordinateOverflow { row: usize, column: usize },
    /// The address is not a complete A1 coordinate.
    InvalidAddress,
}

impl fmt::Display for CellPositionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::CoordinateOverflow { .. } => formatter.write_str("cell coordinate overflows u32"),
            Self::InvalidAddress => formatter.write_str("invalid A1 cell coordinate"),
        }
    }
}

impl std::error::Error for CellPositionError {}

/// Failure from a checked Number-format value constructor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberError {
    /// Fixed precision exceeds the supported native Number range.
    DecimalPlacesOutOfRange { value: u8, maximum: u8 },
}

impl fmt::Display for NumberError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::DecimalPlacesOutOfRange { value, maximum } => {
                write!(formatter, "decimal places {value} exceed maximum {maximum}")
            },
        }
    }
}

impl std::error::Error for NumberError {}

/// Largest fixed fractional precision accepted by the Keynote Number codec.
pub const MAX_DECIMAL_PLACES: u8 = 30;

/// A checked fixed fractional-digit count.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FixedDecimalPlaces(u8);

impl FixedDecimalPlaces {
    /// Zero fractional digits.
    pub const ZERO: Self = Self(0);
    /// Two fractional digits.
    pub const TWO: Self = Self(2);
    /// Maximum supported fixed precision.
    pub const MAXIMUM: Self = Self(MAX_DECIMAL_PLACES);

    /// Validates and constructs a fixed precision.
    pub const fn new(value: u8) -> Result<Self, NumberError> {
        if value > MAX_DECIMAL_PLACES {
            return Err(NumberError::DecimalPlacesOutOfRange {
                value,
                maximum: MAX_DECIMAL_PLACES,
            });
        }
        Ok(Self(value))
    }

    /// Returns the number of fractional digits.
    #[must_use]
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for FixedDecimalPlaces {
    type Error = NumberError;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// Automatic or fixed fractional digits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum DecimalPlaces {
    /// Select precision from the displayed value.
    #[default]
    Automatic,
    /// Always display the selected number of fractional digits.
    Fixed(FixedDecimalPlaces),
}

impl DecimalPlaces {
    /// Constructs checked fixed precision.
    pub const fn fixed(value: u8) -> Result<Self, NumberError> {
        match FixedDecimalPlaces::new(value) {
            Ok(value) => Ok(Self::Fixed(value)),
            Err(error) => Err(error),
        }
    }
}

/// Presentation of negative decimal values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum NegativeStyle {
    /// Use a leading minus sign.
    #[default]
    MinusSign,
    /// Use red text without a minus sign.
    Red,
    /// Enclose the value in parentheses.
    Parentheses,
    /// Use red text and parentheses.
    RedParentheses,
}

/// Whether locale-aware thousands grouping is displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ThousandsSeparator {
    /// Do not display a grouping separator.
    #[default]
    Hidden,
    /// Display the locale's grouping separator.
    Shown,
}

/// Complete semantic decimal Number display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Number {
    decimal_places: DecimalPlaces,
    negative_style: NegativeStyle,
    thousands_separator: ThousandsSeparator,
}

impl Number {
    /// Constructs a complete decimal Number display format.
    #[must_use]
    pub const fn new(
        decimal_places: DecimalPlaces,
        negative_style: NegativeStyle,
        thousands_separator: ThousandsSeparator,
    ) -> Self {
        Self {
            decimal_places,
            negative_style,
            thousands_separator,
        }
    }

    /// Returns automatic or fixed precision.
    #[must_use]
    pub const fn decimal_places(self) -> DecimalPlaces {
        self.decimal_places
    }

    /// Returns negative-value presentation.
    #[must_use]
    pub const fn negative_style(self) -> NegativeStyle {
        self.negative_style
    }

    /// Returns thousands-separator presentation.
    #[must_use]
    pub const fn thousands_separator(self) -> ThousandsSeparator {
        self.thousands_separator
    }

    /// Replaces precision.
    #[must_use]
    pub const fn with_decimal_places(mut self, value: DecimalPlaces) -> Self {
        self.decimal_places = value;
        self
    }

    /// Replaces negative-value presentation.
    #[must_use]
    pub const fn with_negative_style(mut self, value: NegativeStyle) -> Self {
        self.negative_style = value;
        self
    }

    /// Replaces thousands-separator presentation.
    #[must_use]
    pub const fn with_thousands_separator(mut self, value: ThousandsSeparator) -> Self {
        self.thousands_separator = value;
        self
    }
}

/// Exact-source transactions for one existing slide-table cell.
pub mod transaction {
    pub use crate::package::slide_table_cell_number_format::{
        SlideTableCellNumberFormatCommit as Commit,
        SlideTableCellNumberFormatDiagnostics as Diagnostics,
        SlideTableCellNumberFormatEdit as Edit, SlideTableCellNumberFormatError as Error,
        SlideTableCellNumberFormatLimitKind as LimitKind, SlideTableCellNumberFormatPatch as Patch,
        SlideTableCellNumberFormatPath as Path,
    };
}
