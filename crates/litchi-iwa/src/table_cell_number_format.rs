//! Typed decimal-number formatting shared by native iWork table cells.

use crate::{Error, Result};

const MAXIMUM_DECIMAL_PLACES: u8 = 30;

/// A fixed fractional-digit count accepted by the iWork cell inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellFixedDecimalPlaces(u8);

impl TableCellFixedDecimalPlaces {
    /// Render no fractional digits.
    pub const ZERO: Self = Self(0);

    /// Largest fixed precision accepted by the native inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_DECIMAL_PLACES);

    /// Validate and construct a fixed fractional-digit count.
    pub fn new(value: u8) -> Result<Self> {
        if value > MAXIMUM_DECIMAL_PLACES {
            return Err(Error::InvalidFormat(format!(
                "table-cell decimal places must not exceed {MAXIMUM_DECIMAL_PLACES}"
            )));
        }
        Ok(Self(value))
    }

    /// Return the fractional-digit count.
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for TableCellFixedDecimalPlaces {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// Automatic or fixed fractional digits for a decimal table-cell format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellDecimalPlaces {
    /// Let iWork display the precision needed by each value.
    #[default]
    Automatic,
    /// Always display the specified number of fractional digits.
    Fixed(TableCellFixedDecimalPlaces),
}

impl TableCellDecimalPlaces {
    /// Validate and construct a fixed fractional-digit setting.
    pub fn fixed(value: u8) -> Result<Self> {
        TableCellFixedDecimalPlaces::new(value).map(Self::Fixed)
    }
}

/// Native presentation of negative decimal values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellNegativeNumberStyle {
    /// Normal text with a leading minus sign.
    #[default]
    MinusSign,
    /// Red text without a minus sign.
    Red,
    /// Normal text enclosed in parentheses.
    Parentheses,
    /// Red text enclosed in parentheses.
    RedParentheses,
}

/// Whether a decimal table cell displays locale-aware digit grouping.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellThousandsSeparator {
    /// Do not group thousands.
    #[default]
    Hidden,
    /// Display the locale's thousands separator.
    Shown,
}

/// Explicit decimal-number display format for one native table cell.
///
/// Absence of this value in editor read APIs means the cell uses iWork's
/// automatic format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellNumberFormat {
    decimal_places: TableCellDecimalPlaces,
    negative_style: TableCellNegativeNumberStyle,
    thousands_separator: TableCellThousandsSeparator,
}

impl TableCellNumberFormat {
    /// Construct a complete decimal-number format.
    pub const fn new(
        decimal_places: TableCellDecimalPlaces,
        negative_style: TableCellNegativeNumberStyle,
        thousands_separator: TableCellThousandsSeparator,
    ) -> Self {
        Self {
            decimal_places,
            negative_style,
            thousands_separator,
        }
    }

    /// Return the automatic or fixed fractional-digit setting.
    pub const fn decimal_places(self) -> TableCellDecimalPlaces {
        self.decimal_places
    }

    /// Return the negative-number presentation.
    pub const fn negative_style(self) -> TableCellNegativeNumberStyle {
        self.negative_style
    }

    /// Return whether digit grouping is displayed.
    pub const fn thousands_separator(self) -> TableCellThousandsSeparator {
        self.thousands_separator
    }

    /// Replace the fractional-digit setting.
    pub const fn with_decimal_places(mut self, decimal_places: TableCellDecimalPlaces) -> Self {
        self.decimal_places = decimal_places;
        self
    }

    /// Replace the negative-number presentation.
    pub const fn with_negative_style(
        mut self,
        negative_style: TableCellNegativeNumberStyle,
    ) -> Self {
        self.negative_style = negative_style;
        self
    }

    /// Show or hide locale-aware digit grouping.
    pub const fn with_thousands_separator(
        mut self,
        thousands_separator: TableCellThousandsSeparator,
    ) -> Self {
        self.thousands_separator = thousands_separator;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decimal_places_are_bounded_and_formats_compose() {
        assert_eq!(
            TableCellFixedDecimalPlaces::new(MAXIMUM_DECIMAL_PLACES)
                .unwrap()
                .value(),
            MAXIMUM_DECIMAL_PLACES
        );
        assert!(TableCellFixedDecimalPlaces::new(MAXIMUM_DECIMAL_PLACES + 1).is_err());

        let format = TableCellNumberFormat::default()
            .with_decimal_places(TableCellDecimalPlaces::fixed(2).unwrap())
            .with_negative_style(TableCellNegativeNumberStyle::Parentheses)
            .with_thousands_separator(TableCellThousandsSeparator::Shown);
        assert_eq!(
            format.decimal_places(),
            TableCellDecimalPlaces::Fixed(TableCellFixedDecimalPlaces::new(2).unwrap())
        );
    }
}
