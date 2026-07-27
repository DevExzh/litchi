//! Strongly typed data formats shared by native iWork table cells.

use crate::{Error, Result};
use std::fmt;
use std::str::FromStr;

const MAXIMUM_DECIMAL_PLACES: u8 = 30;

mod date_time;
mod duration;
mod numeral_system;
mod numeric_control;
mod pop_up_menu;
mod slider;
mod stepper;
pub use date_time::TableCellDateTimeFormat;
pub use duration::{
    TableCellDurationFormat, TableCellDurationStyle, TableCellDurationUnit,
    TableCellDurationUnitRange, TableCellDurationUnits,
};
pub use numeral_system::{
    TableCellNumeralSystemBase, TableCellNumeralSystemFixedPlaces, TableCellNumeralSystemFormat,
    TableCellNumeralSystemNegativeStyle, TableCellNumeralSystemPlaces,
};
pub use numeric_control::TableCellNumericControlDisplayFormat;
pub use pop_up_menu::{
    TableCellPopUpMenuFormat, TableCellPopUpMenuInitialSelection, TableCellPopUpMenuItem,
};
pub use slider::{TableCellSliderDisplayFormat, TableCellSliderFormat, TableCellSliderRange};
pub use stepper::{TableCellStepperDisplayFormat, TableCellStepperFormat, TableCellStepperRange};

/// Native interactive Checkbox format for one table cell.
///
/// iWork stores no configurable options for this format. Applying it to an
/// empty cell creates an unchecked Boolean value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellCheckboxFormat;

/// Native interactive five-star rating format for one table cell.
///
/// iWork stores a fixed zero-through-five range with whole-star increments;
/// the inspector exposes no configurable options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellStarRatingFormat;

/// Native explicit Text format for one table cell.
///
/// This preserves empty and text values verbatim. Applying Text to another
/// value type is rejected because iWork converts the locale-formatted display
/// string, which cannot be reproduced safely without the originating locale.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellTextFormat;

/// A fixed fractional-digit count accepted by the iWork cell inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellFixedDecimalPlaces(u8);

impl TableCellFixedDecimalPlaces {
    /// Render no fractional digits.
    pub const ZERO: Self = Self(0);
    /// Render two fractional digits.
    pub const TWO: Self = Self(2);
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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
struct DecimalFormat {
    decimal_places: TableCellDecimalPlaces,
    negative_style: TableCellNegativeNumberStyle,
    thousands_separator: TableCellThousandsSeparator,
}

impl DecimalFormat {
    const fn new(
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
}

macro_rules! decimal_format {
    ($name:ident, $description:literal) => {
        #[doc = $description]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
        pub struct $name(DecimalFormat);

        impl $name {
            /// Construct a complete decimal display format.
            pub const fn new(
                decimal_places: TableCellDecimalPlaces,
                negative_style: TableCellNegativeNumberStyle,
                thousands_separator: TableCellThousandsSeparator,
            ) -> Self {
                Self(DecimalFormat::new(
                    decimal_places,
                    negative_style,
                    thousands_separator,
                ))
            }

            /// Return the automatic or fixed fractional-digit setting.
            pub const fn decimal_places(self) -> TableCellDecimalPlaces {
                self.0.decimal_places
            }

            /// Return the negative-number presentation.
            pub const fn negative_style(self) -> TableCellNegativeNumberStyle {
                self.0.negative_style
            }

            /// Return whether digit grouping is displayed.
            pub const fn thousands_separator(self) -> TableCellThousandsSeparator {
                self.0.thousands_separator
            }

            /// Replace the fractional-digit setting.
            pub const fn with_decimal_places(
                mut self,
                decimal_places: TableCellDecimalPlaces,
            ) -> Self {
                self.0.decimal_places = decimal_places;
                self
            }

            /// Replace the negative-number presentation.
            pub const fn with_negative_style(
                mut self,
                negative_style: TableCellNegativeNumberStyle,
            ) -> Self {
                self.0.negative_style = negative_style;
                self
            }

            /// Show or hide locale-aware digit grouping.
            pub const fn with_thousands_separator(
                mut self,
                thousands_separator: TableCellThousandsSeparator,
            ) -> Self {
                self.0.thousands_separator = thousands_separator;
                self
            }
        }
    };
}

decimal_format!(
    TableCellNumberFormat,
    "Explicit decimal-number display format for one native table cell."
);
decimal_format!(
    TableCellPercentageFormat,
    "Explicit percentage display format for one native table cell."
);

/// Explicit scientific-notation display format for one native table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCellScientificFormat {
    decimal_places: TableCellFixedDecimalPlaces,
}

impl TableCellScientificFormat {
    /// Construct a scientific format with a fixed fractional-digit count.
    pub const fn new(decimal_places: TableCellFixedDecimalPlaces) -> Self {
        Self { decimal_places }
    }

    /// Return the fixed number of digits rendered after the decimal point.
    pub const fn decimal_places(self) -> TableCellFixedDecimalPlaces {
        self.decimal_places
    }

    /// Replace the fixed fractional-digit count.
    pub const fn with_decimal_places(
        mut self,
        decimal_places: TableCellFixedDecimalPlaces,
    ) -> Self {
        self.decimal_places = decimal_places;
        self
    }
}

impl Default for TableCellScientificFormat {
    fn default() -> Self {
        Self::new(TableCellFixedDecimalPlaces::TWO)
    }
}

/// Denominator strategy used by iWork's native Fraction cell format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellFractionAccuracy {
    /// Choose the closest fraction whose denominator has at most one digit.
    UpToOneDigit,
    /// Choose the closest fraction whose denominator has at most two digits.
    UpToTwoDigits,
    /// Choose the closest fraction whose denominator has at most three digits.
    #[default]
    UpToThreeDigits,
    /// Always use a denominator of two.
    Halves,
    /// Always use a denominator of four.
    Quarters,
    /// Always use a denominator of eight.
    Eighths,
    /// Always use a denominator of sixteen.
    Sixteenths,
    /// Always use a denominator of ten.
    Tenths,
    /// Always use a denominator of one hundred.
    Hundredths,
}

/// Explicit mixed-fraction display format for one native table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellFractionFormat {
    accuracy: TableCellFractionAccuracy,
}

impl TableCellFractionFormat {
    /// Construct a Fraction format with the requested denominator strategy.
    pub const fn new(accuracy: TableCellFractionAccuracy) -> Self {
        Self { accuracy }
    }

    /// Return the denominator strategy.
    pub const fn accuracy(self) -> TableCellFractionAccuracy {
        self.accuracy
    }

    /// Replace the denominator strategy.
    pub const fn with_accuracy(mut self, accuracy: TableCellFractionAccuracy) -> Self {
        self.accuracy = accuracy;
        self
    }
}

/// Three-letter ISO 4217-style currency code stored by native iWork.
///
/// The compact inline representation avoids allocating for every formatted
/// cell while rejecting locale labels and malformed archive values.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellCurrencyCode([u8; 3]);

impl TableCellCurrencyCode {
    /// United States dollar.
    pub const USD: Self = Self(*b"USD");
    /// Euro.
    pub const EUR: Self = Self(*b"EUR");
    /// British pound sterling.
    pub const GBP: Self = Self(*b"GBP");
    /// Japanese yen.
    pub const JPY: Self = Self(*b"JPY");

    /// Validate a three-letter uppercase ASCII currency code.
    pub fn new(value: &str) -> Result<Self> {
        let bytes = value.as_bytes();
        if bytes.len() != 3 || !bytes.iter().all(u8::is_ascii_uppercase) {
            return Err(Error::InvalidFormat(format!(
                "table-cell currency code must contain exactly three uppercase ASCII letters, found {value:?}"
            )));
        }
        Ok(Self([bytes[0], bytes[1], bytes[2]]))
    }

    /// Borrow the validated code without allocation.
    pub fn as_str(&self) -> &str {
        std::str::from_utf8(&self.0).expect("validated ASCII currency code")
    }
}

impl Default for TableCellCurrencyCode {
    fn default() -> Self {
        Self::USD
    }
}

impl fmt::Debug for TableCellCurrencyCode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("TableCellCurrencyCode")
            .field(&self.as_str())
            .finish()
    }
}

impl fmt::Display for TableCellCurrencyCode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for TableCellCurrencyCode {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for TableCellCurrencyCode {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// Whether Currency uses ordinary or accounting alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableCellCurrencyStyle {
    /// Place the currency symbol next to the number.
    #[default]
    Standard,
    /// Align currency symbols separately at the leading edge of the cell.
    Accounting,
}

/// Explicit currency display format for one native table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TableCellCurrencyFormat {
    decimal: DecimalFormat,
    currency_code: TableCellCurrencyCode,
    style: TableCellCurrencyStyle,
}

impl TableCellCurrencyFormat {
    /// Construct a complete currency display format.
    pub const fn new(
        currency_code: TableCellCurrencyCode,
        decimal_places: TableCellDecimalPlaces,
        negative_style: TableCellNegativeNumberStyle,
        thousands_separator: TableCellThousandsSeparator,
        style: TableCellCurrencyStyle,
    ) -> Self {
        Self {
            decimal: DecimalFormat::new(decimal_places, negative_style, thousands_separator),
            currency_code,
            style,
        }
    }

    /// Return the native three-letter currency code.
    pub const fn currency_code(self) -> TableCellCurrencyCode {
        self.currency_code
    }

    /// Return the automatic or fixed fractional-digit setting.
    pub const fn decimal_places(self) -> TableCellDecimalPlaces {
        self.decimal.decimal_places
    }

    /// Return the stored negative-number presentation.
    ///
    /// iWork retains this value but disables its inspector control while
    /// accounting style is active.
    pub const fn negative_style(self) -> TableCellNegativeNumberStyle {
        self.decimal.negative_style
    }

    /// Return whether digit grouping is displayed.
    pub const fn thousands_separator(self) -> TableCellThousandsSeparator {
        self.decimal.thousands_separator
    }

    /// Return the standard or accounting presentation.
    pub const fn style(self) -> TableCellCurrencyStyle {
        self.style
    }

    /// Replace the native currency code.
    pub const fn with_currency_code(mut self, currency_code: TableCellCurrencyCode) -> Self {
        self.currency_code = currency_code;
        self
    }

    /// Replace the fractional-digit setting.
    pub const fn with_decimal_places(mut self, decimal_places: TableCellDecimalPlaces) -> Self {
        self.decimal.decimal_places = decimal_places;
        self
    }

    /// Replace the stored negative-number presentation.
    pub const fn with_negative_style(
        mut self,
        negative_style: TableCellNegativeNumberStyle,
    ) -> Self {
        self.decimal.negative_style = negative_style;
        self
    }

    /// Show or hide locale-aware digit grouping.
    pub const fn with_thousands_separator(
        mut self,
        thousands_separator: TableCellThousandsSeparator,
    ) -> Self {
        self.decimal.thousands_separator = thousands_separator;
        self
    }

    /// Replace the standard or accounting presentation.
    pub const fn with_style(mut self, style: TableCellCurrencyStyle) -> Self {
        self.style = style;
        self
    }
}

/// Data format stored explicitly on one native iWork table cell.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub enum TableCellDataFormat {
    /// Let iWork infer the format from the cell value.
    #[default]
    Automatic,
    /// Display the value as a decimal number.
    Number(TableCellNumberFormat),
    /// Preserve and display the value as literal text.
    Text(TableCellTextFormat),
    /// Display the value using a native currency and optional accounting style.
    Currency(TableCellCurrencyFormat),
    /// Multiply the displayed value by one hundred and append a percent sign.
    Percentage(TableCellPercentageFormat),
    /// Display the value in scientific notation.
    Scientific(TableCellScientificFormat),
    /// Display the value as a whole number and native fraction.
    Fraction(TableCellFractionFormat),
    /// Display the rounded integer in a positional numeral system.
    NumeralSystem(TableCellNumeralSystemFormat),
    /// Display a native Date value with an ICU-style pattern.
    DateTime(TableCellDateTimeFormat),
    /// Display a native Duration value with typed style and unit settings.
    Duration(TableCellDurationFormat),
    /// Display and edit the value as a native Checkbox control.
    Checkbox(TableCellCheckboxFormat),
    /// Display and edit the numeric value as a native five-star rating.
    StarRating(TableCellStarRatingFormat),
    /// Display and edit the numeric value with a native Slider control.
    Slider(TableCellSliderFormat),
    /// Display and edit the numeric value with native increment and decrement buttons.
    Stepper(TableCellStepperFormat),
    /// Display and edit a text value with a native Pop-Up Menu control.
    PopUpMenu(TableCellPopUpMenuFormat),
}

impl From<TableCellNumberFormat> for TableCellDataFormat {
    fn from(value: TableCellNumberFormat) -> Self {
        Self::Number(value)
    }
}

impl From<TableCellTextFormat> for TableCellDataFormat {
    fn from(value: TableCellTextFormat) -> Self {
        Self::Text(value)
    }
}

impl From<TableCellPercentageFormat> for TableCellDataFormat {
    fn from(value: TableCellPercentageFormat) -> Self {
        Self::Percentage(value)
    }
}

impl From<TableCellCurrencyFormat> for TableCellDataFormat {
    fn from(value: TableCellCurrencyFormat) -> Self {
        Self::Currency(value)
    }
}

impl From<TableCellScientificFormat> for TableCellDataFormat {
    fn from(value: TableCellScientificFormat) -> Self {
        Self::Scientific(value)
    }
}

impl From<TableCellFractionFormat> for TableCellDataFormat {
    fn from(value: TableCellFractionFormat) -> Self {
        Self::Fraction(value)
    }
}

impl From<TableCellCheckboxFormat> for TableCellDataFormat {
    fn from(value: TableCellCheckboxFormat) -> Self {
        Self::Checkbox(value)
    }
}

impl From<TableCellStarRatingFormat> for TableCellDataFormat {
    fn from(value: TableCellStarRatingFormat) -> Self {
        Self::StarRating(value)
    }
}

impl From<TableCellSliderFormat> for TableCellDataFormat {
    fn from(value: TableCellSliderFormat) -> Self {
        Self::Slider(value)
    }
}

impl From<TableCellStepperFormat> for TableCellDataFormat {
    fn from(value: TableCellStepperFormat) -> Self {
        Self::Stepper(value)
    }
}

impl From<TableCellPopUpMenuFormat> for TableCellDataFormat {
    fn from(value: TableCellPopUpMenuFormat) -> Self {
        Self::PopUpMenu(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decimal_places_are_bounded_and_formats_are_distinct() {
        assert_eq!(
            TableCellFixedDecimalPlaces::new(MAXIMUM_DECIMAL_PLACES)
                .unwrap()
                .value(),
            MAXIMUM_DECIMAL_PLACES
        );
        assert!(TableCellFixedDecimalPlaces::new(MAXIMUM_DECIMAL_PLACES + 1).is_err());

        let number = TableCellNumberFormat::default()
            .with_decimal_places(TableCellDecimalPlaces::fixed(2).unwrap())
            .with_negative_style(TableCellNegativeNumberStyle::Parentheses)
            .with_thousands_separator(TableCellThousandsSeparator::Shown);
        let percentage = TableCellPercentageFormat::new(
            number.decimal_places(),
            number.negative_style(),
            number.thousands_separator(),
        );
        assert_eq!(
            TableCellDataFormat::from(number),
            TableCellDataFormat::Number(number)
        );
        assert_eq!(
            TableCellDataFormat::from(percentage),
            TableCellDataFormat::Percentage(percentage)
        );
    }

    #[test]
    fn currency_codes_are_inline_strict_and_ergonomic() {
        assert_eq!(std::mem::size_of::<TableCellCurrencyCode>(), 3);
        assert_eq!(TableCellCurrencyCode::new("EUR").unwrap().as_str(), "EUR");
        assert_eq!(
            "JPY".parse::<TableCellCurrencyCode>().unwrap(),
            TableCellCurrencyCode::JPY
        );
        assert!(TableCellCurrencyCode::new("usd").is_err());
        assert!(TableCellCurrencyCode::new("US").is_err());

        let currency = TableCellCurrencyFormat::default()
            .with_currency_code(TableCellCurrencyCode::EUR)
            .with_style(TableCellCurrencyStyle::Accounting);
        assert_eq!(currency.currency_code(), TableCellCurrencyCode::EUR);
        assert_eq!(
            TableCellDataFormat::from(currency),
            TableCellDataFormat::Currency(currency)
        );
    }

    #[test]
    fn scientific_format_has_native_default_precision() {
        let format = TableCellScientificFormat::default();
        assert_eq!(format.decimal_places(), TableCellFixedDecimalPlaces::TWO);
        assert_eq!(
            TableCellDataFormat::from(format),
            TableCellDataFormat::Scientific(format)
        );
    }

    #[test]
    fn fraction_format_has_native_default_accuracy() {
        let format = TableCellFractionFormat::default();
        assert_eq!(
            format.accuracy(),
            TableCellFractionAccuracy::UpToThreeDigits
        );
        assert_eq!(
            TableCellDataFormat::from(format),
            TableCellDataFormat::Fraction(format)
        );
    }

    #[test]
    fn numeral_system_bounds_and_signed_invariants_are_strict() {
        assert!(TableCellNumeralSystemBase::new(1).is_err());
        assert_eq!(
            TableCellNumeralSystemBase::new(36).unwrap(),
            TableCellNumeralSystemBase::MAXIMUM
        );
        assert!(TableCellNumeralSystemBase::new(37).is_err());
        assert!(TableCellNumeralSystemFixedPlaces::new(0).is_err());
        assert_eq!(
            TableCellNumeralSystemFixedPlaces::new(32).unwrap(),
            TableCellNumeralSystemFixedPlaces::MAXIMUM
        );
        assert!(TableCellNumeralSystemFixedPlaces::new(33).is_err());

        let hexadecimal = TableCellNumeralSystemFormat::new(
            TableCellNumeralSystemBase::HEXADECIMAL,
            TableCellNumeralSystemPlaces::Fixed(TableCellNumeralSystemFixedPlaces::EIGHT),
            TableCellNumeralSystemNegativeStyle::TwosComplement,
        )
        .unwrap();
        assert_eq!(
            TableCellDataFormat::from(hexadecimal),
            TableCellDataFormat::NumeralSystem(hexadecimal)
        );
        assert!(
            hexadecimal
                .with_base(TableCellNumeralSystemBase::DECIMAL)
                .is_err()
        );
        assert!(
            hexadecimal
                .with_places(TableCellNumeralSystemPlaces::Minimum)
                .is_err()
        );
        assert_eq!(
            TableCellNumeralSystemFormat::default().negative_style(),
            TableCellNumeralSystemNegativeStyle::MinusSign
        );
    }

    #[test]
    fn date_time_patterns_are_owned_strict_and_ergonomic() {
        assert!(TableCellDateTimeFormat::new("").is_err());
        assert!(TableCellDateTimeFormat::new("yyyy\0MM").is_err());
        assert_eq!(
            TableCellDateTimeFormat::iso_date_time_24_hour_with_seconds().pattern(),
            "yyyy-MM-dd H:mm:ss"
        );
        let custom = TableCellDateTimeFormat::new("EEEE, MMMM d, y").unwrap();
        assert_eq!(
            TableCellDataFormat::from(custom.clone()),
            TableCellDataFormat::DateTime(custom)
        );
    }

    #[test]
    fn duration_units_are_strict_and_ergonomic() {
        assert!(
            TableCellDurationUnitRange::new(
                TableCellDurationUnit::Seconds,
                TableCellDurationUnit::Hours,
            )
            .is_err()
        );
        let range = TableCellDurationUnitRange::hours_to_milliseconds();
        let format = TableCellDurationFormat::custom(TableCellDurationStyle::Abbreviated, range);
        assert_eq!(format.style(), TableCellDurationStyle::Abbreviated);
        assert_eq!(format.units(), TableCellDurationUnits::Custom(range));
        assert_eq!(
            TableCellDataFormat::from(format),
            TableCellDataFormat::Duration(format)
        );
    }
}
