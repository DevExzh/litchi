//! Strongly typed settings for native positional numeral-system cells.

use super::TableCellDataFormat;
use crate::{Error, Result};

const MINIMUM_BASE: u8 = 2;
const MAXIMUM_BASE: u8 = 36;
const MAXIMUM_PLACES: u8 = 32;

/// A positional numeral-system base accepted by iWork.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellNumeralSystemBase(u8);

impl TableCellNumeralSystemBase {
    /// Binary.
    pub const BINARY: Self = Self(2);
    /// Octal.
    pub const OCTAL: Self = Self(8);
    /// Decimal.
    pub const DECIMAL: Self = Self(10);
    /// Hexadecimal.
    pub const HEXADECIMAL: Self = Self(16);
    /// Largest base supported by the native inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_BASE);

    /// Validate and construct a positional numeral-system base.
    pub fn new(value: u8) -> Result<Self> {
        if !(MINIMUM_BASE..=MAXIMUM_BASE).contains(&value) {
            return Err(Error::InvalidFormat(format!(
                "table-cell numeral-system base must be between {MINIMUM_BASE} and {MAXIMUM_BASE}, found {value}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the numeric base.
    pub const fn value(self) -> u8 {
        self.0
    }

    const fn supports_twos_complement(self) -> bool {
        matches!(self.0, 2 | 8 | 16)
    }
}

impl Default for TableCellNumeralSystemBase {
    fn default() -> Self {
        Self::DECIMAL
    }
}

impl TryFrom<u8> for TableCellNumeralSystemBase {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// A nonzero, fixed digit width accepted by iWork's Numeral System format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellNumeralSystemFixedPlaces(u8);

impl TableCellNumeralSystemFixedPlaces {
    /// One displayed digit.
    pub const ONE: Self = Self(1);
    /// Eight displayed digits.
    pub const EIGHT: Self = Self(8);
    /// Largest fixed width supported by the native inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_PLACES);

    /// Validate and construct a fixed digit width.
    pub fn new(value: u8) -> Result<Self> {
        if value == 0 || value > MAXIMUM_PLACES {
            return Err(Error::InvalidFormat(format!(
                "table-cell numeral-system places must be between 1 and {MAXIMUM_PLACES}, found {value}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the fixed digit width.
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for TableCellNumeralSystemFixedPlaces {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// Minimum-width or fixed-width numeral-system output.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellNumeralSystemPlaces {
    /// Display only the digits needed by the value.
    #[default]
    Minimum,
    /// Apply the native fixed-place setting.
    Fixed(TableCellNumeralSystemFixedPlaces),
}

impl TableCellNumeralSystemPlaces {
    /// Validate and construct a fixed-width setting.
    pub fn fixed(value: u8) -> Result<Self> {
        TableCellNumeralSystemFixedPlaces::new(value).map(Self::Fixed)
    }
}

/// How iWork represents negative values in a positional numeral system.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellNumeralSystemNegativeStyle {
    /// Prefix the converted magnitude with a minus sign.
    #[default]
    MinusSign,
    /// Encode the rounded integer using fixed-width two's complement.
    TwosComplement,
}

/// Explicit positional numeral-system display format for one native table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCellNumeralSystemFormat {
    base: TableCellNumeralSystemBase,
    places: TableCellNumeralSystemPlaces,
    negative_style: TableCellNumeralSystemNegativeStyle,
}

impl TableCellNumeralSystemFormat {
    /// Validate and construct a complete Numeral System format.
    pub fn new(
        base: TableCellNumeralSystemBase,
        places: TableCellNumeralSystemPlaces,
        negative_style: TableCellNumeralSystemNegativeStyle,
    ) -> Result<Self> {
        if negative_style == TableCellNumeralSystemNegativeStyle::TwosComplement
            && (!base.supports_twos_complement()
                || matches!(places, TableCellNumeralSystemPlaces::Minimum))
        {
            return Err(Error::InvalidFormat(
                "two's-complement numeral cells require base 2, 8, or 16 and a fixed digit width"
                    .to_owned(),
            ));
        }
        Ok(Self {
            base,
            places,
            negative_style,
        })
    }

    /// Return the positional numeral-system base.
    pub const fn base(self) -> TableCellNumeralSystemBase {
        self.base
    }

    /// Return the minimum or fixed digit width.
    pub const fn places(self) -> TableCellNumeralSystemPlaces {
        self.places
    }

    /// Return the negative-value representation.
    pub const fn negative_style(self) -> TableCellNumeralSystemNegativeStyle {
        self.negative_style
    }

    /// Validate and replace the positional numeral-system base.
    pub fn with_base(self, base: TableCellNumeralSystemBase) -> Result<Self> {
        Self::new(base, self.places, self.negative_style)
    }

    /// Validate and replace the minimum or fixed digit width.
    pub fn with_places(self, places: TableCellNumeralSystemPlaces) -> Result<Self> {
        Self::new(self.base, places, self.negative_style)
    }

    /// Validate and replace the negative-value representation.
    pub fn with_negative_style(
        self,
        negative_style: TableCellNumeralSystemNegativeStyle,
    ) -> Result<Self> {
        Self::new(self.base, self.places, negative_style)
    }
}

impl Default for TableCellNumeralSystemFormat {
    fn default() -> Self {
        Self {
            base: TableCellNumeralSystemBase::DECIMAL,
            places: TableCellNumeralSystemPlaces::Minimum,
            negative_style: TableCellNumeralSystemNegativeStyle::MinusSign,
        }
    }
}

impl From<TableCellNumeralSystemFormat> for TableCellDataFormat {
    fn from(value: TableCellNumeralSystemFormat) -> Self {
        Self::NumeralSystem(value)
    }
}
