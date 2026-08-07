//! Decimal, currency, percentage, scientific, and fraction values.

use std::fmt;
use std::str::FromStr;

macro_rules! decimal_format {
    ($name:ident, $description:literal) => {
        #[doc = $description]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
        pub struct $name(Decimal);

        impl $name {
            /// Constructs a complete decimal display format.
            #[must_use]
            pub const fn new(
                decimal_places: DecimalPlaces,
                negative_style: NegativeStyle,
                thousands_separator: ThousandsSeparator,
            ) -> Self {
                Self(Decimal::new(
                    decimal_places,
                    negative_style,
                    thousands_separator,
                ))
            }

            /// Returns the automatic or fixed precision setting.
            #[must_use]
            pub const fn decimal_places(self) -> DecimalPlaces {
                self.0.places
            }

            /// Returns the negative-value presentation.
            #[must_use]
            pub const fn negative_style(self) -> NegativeStyle {
                self.0.negative_style
            }

            /// Returns the thousands-separator setting.
            #[must_use]
            pub const fn thousands_separator(self) -> ThousandsSeparator {
                self.0.thousands_separator
            }

            /// Replaces the precision setting.
            #[must_use]
            pub const fn with_decimal_places(mut self, value: DecimalPlaces) -> Self {
                self.0.places = value;
                self
            }

            /// Replaces the negative-value presentation.
            #[must_use]
            pub const fn with_negative_style(mut self, value: NegativeStyle) -> Self {
                self.0.negative_style = value;
                self
            }

            /// Replaces the thousands-separator setting.
            #[must_use]
            pub const fn with_thousands_separator(mut self, value: ThousandsSeparator) -> Self {
                self.0.thousands_separator = value;
                self
            }
        }
    };
}

/// Largest fixed fractional precision accepted by Numbers cell formats.
pub const MAX_DECIMAL_PLACES: u8 = 30;

/// Errors returned by checked number-format constructors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// A fixed precision exceeds [`MAX_DECIMAL_PLACES`].
    DecimalPlacesOutOfRange { value: u8, maximum: u8 },
    /// A currency code does not contain exactly three bytes.
    CurrencyCodeLength { length: usize },
    /// A currency code contains a byte other than an uppercase ASCII letter.
    CurrencyCodeNotUppercase { index: usize },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::DecimalPlacesOutOfRange { value, maximum } => write!(
                formatter,
                "fixed decimal places value {value} exceeds maximum {maximum}"
            ),
            Self::CurrencyCodeLength { length } => write!(
                formatter,
                "currency code must contain exactly three ASCII letters, found {length} bytes"
            ),
            Self::CurrencyCodeNotUppercase { index } => write!(
                formatter,
                "currency code byte at index {index} is not an uppercase ASCII letter"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked number-format constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A checked fixed fractional-digit count.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FixedDecimalPlaces(u8);

impl FixedDecimalPlaces {
    /// Zero fractional digits.
    pub const ZERO: Self = Self(0);
    /// Two fractional digits.
    pub const TWO: Self = Self(2);
    /// The largest supported fixed precision.
    pub const MAXIMUM: Self = Self(MAX_DECIMAL_PLACES);

    /// Validates and constructs a fixed fractional-digit count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalPlacesOutOfRange`] for values above the
    /// supported precision.
    pub const fn new(value: u8) -> Result<Self> {
        if value > MAX_DECIMAL_PLACES {
            return Err(Error::DecimalPlacesOutOfRange {
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
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
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
    /// Validates and constructs a fixed precision setting.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalPlacesOutOfRange`] for values above the
    /// supported precision.
    pub const fn fixed(value: u8) -> Result<Self> {
        match FixedDecimalPlaces::new(value) {
            Ok(fixed_places) => Ok(Self::Fixed(fixed_places)),
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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
struct Decimal {
    places: DecimalPlaces,
    negative_style: NegativeStyle,
    thousands_separator: ThousandsSeparator,
}

impl Decimal {
    const fn new(
        decimal_places: DecimalPlaces,
        negative_style: NegativeStyle,
        thousands_separator: ThousandsSeparator,
    ) -> Self {
        Self {
            places: decimal_places,
            negative_style,
            thousands_separator,
        }
    }
}

decimal_format!(Number, "Decimal-number display format.");
decimal_format!(Percentage, "Percentage display format.");

/// A validated three-letter uppercase ASCII currency code.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CurrencyCode([u8; 3]);

impl CurrencyCode {
    /// United States dollar.
    pub const USD: Self = Self(*b"USD");
    /// Euro.
    pub const EUR: Self = Self(*b"EUR");
    /// British pound sterling.
    pub const GBP: Self = Self(*b"GBP");
    /// Japanese yen.
    pub const JPY: Self = Self(*b"JPY");

    /// Validates and constructs a currency code without allocating.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the input is not exactly three uppercase
    /// ASCII letters.
    pub fn new(value: &str) -> Result<Self> {
        let bytes = value.as_bytes();
        if bytes.len() != 3 {
            return Err(Error::CurrencyCodeLength {
                length: bytes.len(),
            });
        }
        if let Some((index, _)) = bytes
            .iter()
            .enumerate()
            .find(|(_, byte)| !byte.is_ascii_uppercase())
        {
            return Err(Error::CurrencyCodeNotUppercase { index });
        }
        Ok(Self([bytes[0], bytes[1], bytes[2]]))
    }

    /// Borrows the validated code without allocating.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match std::str::from_utf8(&self.0) {
            Ok(value) => value,
            Err(_) => unreachable!("CurrencyCode only stores uppercase ASCII bytes"),
        }
    }
}

impl Default for CurrencyCode {
    fn default() -> Self {
        Self::USD
    }
}

impl fmt::Debug for CurrencyCode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("CurrencyCode")
            .field(&self.as_str())
            .finish()
    }
}

impl fmt::Display for CurrencyCode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for CurrencyCode {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for CurrencyCode {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// Currency alignment style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum CurrencyStyle {
    /// Place the currency symbol next to the number.
    #[default]
    Standard,
    /// Align the currency symbol at the leading edge of the cell.
    Accounting,
}

/// Currency display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Currency {
    decimal: Decimal,
    code: CurrencyCode,
    style: CurrencyStyle,
}

impl Currency {
    /// Constructs a complete currency display format.
    #[must_use]
    pub const fn new(
        code: CurrencyCode,
        decimal_places: DecimalPlaces,
        negative_style: NegativeStyle,
        thousands_separator: ThousandsSeparator,
        style: CurrencyStyle,
    ) -> Self {
        Self {
            decimal: Decimal::new(decimal_places, negative_style, thousands_separator),
            code,
            style,
        }
    }

    /// Returns the currency code.
    #[must_use]
    pub const fn code(self) -> CurrencyCode {
        self.code
    }

    /// Returns the automatic or fixed precision setting.
    #[must_use]
    pub const fn decimal_places(self) -> DecimalPlaces {
        self.decimal.places
    }

    /// Returns the stored negative-value presentation.
    #[must_use]
    pub const fn negative_style(self) -> NegativeStyle {
        self.decimal.negative_style
    }

    /// Returns the thousands-separator setting.
    #[must_use]
    pub const fn thousands_separator(self) -> ThousandsSeparator {
        self.decimal.thousands_separator
    }

    /// Returns the alignment style.
    #[must_use]
    pub const fn style(self) -> CurrencyStyle {
        self.style
    }

    /// Replaces the currency code.
    #[must_use]
    pub const fn with_code(mut self, code: CurrencyCode) -> Self {
        self.code = code;
        self
    }

    /// Replaces the precision setting.
    #[must_use]
    pub const fn with_decimal_places(mut self, value: DecimalPlaces) -> Self {
        self.decimal.places = value;
        self
    }

    /// Replaces the negative-value presentation.
    #[must_use]
    pub const fn with_negative_style(mut self, value: NegativeStyle) -> Self {
        self.decimal.negative_style = value;
        self
    }

    /// Replaces the thousands-separator setting.
    #[must_use]
    pub const fn with_thousands_separator(mut self, value: ThousandsSeparator) -> Self {
        self.decimal.thousands_separator = value;
        self
    }

    /// Replaces the alignment style.
    #[must_use]
    pub const fn with_style(mut self, style: CurrencyStyle) -> Self {
        self.style = style;
        self
    }
}

impl Default for Currency {
    fn default() -> Self {
        Self::new(
            CurrencyCode::USD,
            DecimalPlaces::Automatic,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
            CurrencyStyle::Standard,
        )
    }
}

/// Scientific-notation display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Scientific {
    decimal_places: FixedDecimalPlaces,
}

impl Scientific {
    /// Constructs scientific notation with a fixed precision.
    #[must_use]
    pub const fn new(decimal_places: FixedDecimalPlaces) -> Self {
        Self { decimal_places }
    }

    /// Returns the fixed precision.
    #[must_use]
    pub const fn decimal_places(self) -> FixedDecimalPlaces {
        self.decimal_places
    }

    /// Replaces the fixed precision.
    #[must_use]
    pub const fn with_decimal_places(mut self, value: FixedDecimalPlaces) -> Self {
        self.decimal_places = value;
        self
    }
}

impl Default for Scientific {
    fn default() -> Self {
        Self::new(FixedDecimalPlaces::TWO)
    }
}

/// Denominator strategy used by a fraction format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum FractionAccuracy {
    /// Use a denominator with at most one digit.
    UpToOneDigit,
    /// Use a denominator with at most two digits.
    UpToTwoDigits,
    /// Use a denominator with at most three digits.
    #[default]
    UpToThreeDigits,
    /// Always use halves.
    Halves,
    /// Always use quarters.
    Quarters,
    /// Always use eighths.
    Eighths,
    /// Always use sixteenths.
    Sixteenths,
    /// Always use tenths.
    Tenths,
    /// Always use hundredths.
    Hundredths,
}

/// Mixed-fraction display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Fraction {
    accuracy: FractionAccuracy,
}

impl Fraction {
    /// Constructs a fraction format with the requested accuracy.
    #[must_use]
    pub const fn new(accuracy: FractionAccuracy) -> Self {
        Self { accuracy }
    }

    /// Returns the denominator strategy.
    #[must_use]
    pub const fn accuracy(self) -> FractionAccuracy {
        self.accuracy
    }

    /// Replaces the denominator strategy.
    #[must_use]
    pub const fn with_accuracy(mut self, accuracy: FractionAccuracy) -> Self {
        self.accuracy = accuracy;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn precision_and_currency_validation_is_strict() {
        assert!(matches!(
            FixedDecimalPlaces::new(MAX_DECIMAL_PLACES + 1),
            Err(Error::DecimalPlacesOutOfRange { .. })
        ));
        assert!(matches!(
            CurrencyCode::new("US"),
            Err(Error::CurrencyCodeLength { length: 2 })
        ));
        assert!(matches!(
            CurrencyCode::new("usd"),
            Err(Error::CurrencyCodeNotUppercase { index: 0 })
        ));
    }

    #[test]
    fn decimal_and_currency_values_round_trip_through_accessors() {
        let Ok(places) = DecimalPlaces::fixed(4) else {
            panic!("four decimal places should be valid");
        };
        let format = Number::default()
            .with_decimal_places(places)
            .with_negative_style(NegativeStyle::Parentheses)
            .with_thousands_separator(ThousandsSeparator::Shown);
        assert_eq!(format.decimal_places(), places);
        assert_eq!(format.negative_style(), NegativeStyle::Parentheses);
        assert_eq!(format.thousands_separator(), ThousandsSeparator::Shown);

        let currency = Currency::default()
            .with_code(CurrencyCode::EUR)
            .with_style(CurrencyStyle::Accounting);
        assert_eq!(currency.code(), CurrencyCode::EUR);
        assert_eq!(currency.code().as_str(), "EUR");
        assert_eq!(currency.style(), CurrencyStyle::Accounting);
    }
}
