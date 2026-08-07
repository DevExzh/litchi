//! Archive-free numeric-format vocabulary for table cells.

const MAXIMUM_DECIMAL_PLACES: u8 = 30;
const DECIMAL_PLACES_MASK: u8 = 0x1f;
const NEGATIVE_STYLE_SHIFT: u8 = 5;
const NEGATIVE_STYLE_MASK: u8 = 0x60;
const THOUSANDS_SEPARATOR_MASK: u8 = 0x80;

/// Validation failures for table-cell numeric formats.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The requested fixed precision is outside the native inspector domain.
    #[error("table-cell decimal places {value} must not exceed {maximum}")]
    DecimalPlacesOutOfRange { value: u8, maximum: u8 },
}

/// Result type for table-cell numeric-format construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A checked fixed number of fractional digits.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FixedDecimalPlaces(u8);

impl FixedDecimalPlaces {
    /// No fractional digits.
    pub const ZERO: Self = Self(0);

    /// Two fractional digits.
    pub const TWO: Self = Self(2);

    /// The largest fixed precision accepted by the native inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_DECIMAL_PLACES);

    /// Construct a checked fixed precision.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalPlacesOutOfRange`] when `value` exceeds 30.
    #[must_use = "use the validated decimal places or handle the validation error"]
    pub const fn new(value: u8) -> Result<Self> {
        if value > MAXIMUM_DECIMAL_PLACES {
            return Err(Error::DecimalPlacesOutOfRange {
                value,
                maximum: MAXIMUM_DECIMAL_PLACES,
            });
        }
        Ok(Self(value))
    }

    /// Return the number of fractional digits.
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

/// Automatic or fixed fractional digits for a table-cell number.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum DecimalPlaces {
    /// Let the native application choose the displayed precision.
    #[default]
    Automatic,
    /// Always display exactly this many fractional digits.
    Fixed(FixedDecimalPlaces),
}

impl DecimalPlaces {
    /// Construct a fixed fractional-digit setting.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalPlacesOutOfRange`] when `value` exceeds 30.
    #[must_use = "use the validated decimal places or handle the validation error"]
    pub const fn fixed(value: u8) -> Result<Self> {
        match FixedDecimalPlaces::new(value) {
            Ok(validated) => Ok(Self::Fixed(validated)),
            Err(error) => Err(error),
        }
    }
}

/// Presentation of negative table-cell numbers.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum NegativeStyle {
    /// Render a leading minus sign.
    #[default]
    MinusSign,
    /// Render the magnitude in red without a minus sign.
    Red,
    /// Render the magnitude in parentheses.
    Parentheses,
    /// Render the magnitude in red parentheses.
    RedParentheses,
}

/// Whether a table-cell number uses locale-aware digit grouping.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ThousandsSeparator {
    /// Do not group thousands.
    #[default]
    Hidden,
    /// Display the locale's thousands separator.
    Shown,
}

/// Complete decimal-number formatting for one table cell.
///
/// The three semantic settings are packed into one byte: five bits for the
/// automatic-or-`0..=30` decimal-place domain, two bits for negative style,
/// and one bit for thousands grouping.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct NumberFormat(u8);

impl NumberFormat {
    /// Construct a complete checked numeric format.
    #[must_use]
    pub const fn new(
        decimal_places: DecimalPlaces,
        negative_style: NegativeStyle,
        thousands_separator: ThousandsSeparator,
    ) -> Self {
        let encoded_decimal_places = match decimal_places {
            DecimalPlaces::Automatic => 0,
            DecimalPlaces::Fixed(value) => value.value() + 1,
        };
        let encoded_negative_style = (negative_style as u8) << NEGATIVE_STYLE_SHIFT;
        let encoded_thousands_separator = match thousands_separator {
            ThousandsSeparator::Hidden => 0,
            ThousandsSeparator::Shown => THOUSANDS_SEPARATOR_MASK,
        };
        Self(encoded_decimal_places | encoded_negative_style | encoded_thousands_separator)
    }

    /// Return the automatic or fixed fractional-digit setting.
    #[must_use]
    pub const fn decimal_places(self) -> DecimalPlaces {
        match self.0 & DECIMAL_PLACES_MASK {
            0 => DecimalPlaces::Automatic,
            value => DecimalPlaces::Fixed(FixedDecimalPlaces(value - 1)),
        }
    }

    /// Return the negative-number presentation.
    #[must_use]
    pub const fn negative_style(self) -> NegativeStyle {
        match (self.0 & NEGATIVE_STYLE_MASK) >> NEGATIVE_STYLE_SHIFT {
            0 => NegativeStyle::MinusSign,
            1 => NegativeStyle::Red,
            2 => NegativeStyle::Parentheses,
            _ => NegativeStyle::RedParentheses,
        }
    }

    /// Return whether locale-aware digit grouping is displayed.
    #[must_use]
    pub const fn thousands_separator(self) -> ThousandsSeparator {
        if self.0 & THOUSANDS_SEPARATOR_MASK == 0 {
            ThousandsSeparator::Hidden
        } else {
            ThousandsSeparator::Shown
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{
        DecimalPlaces, Error, FixedDecimalPlaces, NegativeStyle, NumberFormat, ThousandsSeparator,
    };

    #[test]
    fn values_are_checked_and_packed() {
        assert_eq!(size_of::<FixedDecimalPlaces>(), 1);
        assert_eq!(size_of::<DecimalPlaces>(), 2);
        assert_eq!(size_of::<NegativeStyle>(), 1);
        assert_eq!(size_of::<ThousandsSeparator>(), 1);
        assert_eq!(size_of::<NumberFormat>(), 1);
        assert_eq!(
            FixedDecimalPlaces::new(30).unwrap(),
            FixedDecimalPlaces::MAXIMUM
        );
        assert_eq!(
            FixedDecimalPlaces::new(31),
            Err(Error::DecimalPlacesOutOfRange {
                value: 31,
                maximum: 30,
            })
        );
    }

    #[test]
    fn all_semantic_fields_round_trip() {
        let format = NumberFormat::new(
            DecimalPlaces::fixed(12).unwrap(),
            NegativeStyle::RedParentheses,
            ThousandsSeparator::Shown,
        );
        assert_eq!(format.decimal_places(), DecimalPlaces::fixed(12).unwrap());
        assert_eq!(format.negative_style(), NegativeStyle::RedParentheses);
        assert_eq!(format.thousands_separator(), ThousandsSeparator::Shown);
        assert_eq!(
            NumberFormat::default(),
            NumberFormat::new(
                DecimalPlaces::Automatic,
                NegativeStyle::MinusSign,
                ThousandsSeparator::Hidden,
            )
        );
    }
}
