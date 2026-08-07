//! Positional numeral-system display values.

use std::fmt;

const MINIMUM_BASE: u8 = 2;
const MAXIMUM_BASE: u8 = 36;
const MAXIMUM_PLACES: u8 = 32;

/// Errors returned by checked numeral-system constructors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// The base is outside the supported range.
    BaseOutOfRange { value: u8, minimum: u8, maximum: u8 },
    /// Fixed places must be in the supported nonzero range.
    PlacesOutOfRange { value: u8, maximum: u8 },
    /// Two's-complement output needs a supported fixed-width base.
    InvalidTwosComplement,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::BaseOutOfRange {
                value,
                minimum,
                maximum,
            } => write!(
                formatter,
                "numeral-system base {value} is outside {minimum}..={maximum}"
            ),
            Self::PlacesOutOfRange { value, maximum } => write!(
                formatter,
                "numeral-system fixed places {value} is outside 1..={maximum}"
            ),
            Self::InvalidTwosComplement => write!(
                formatter,
                "two's-complement output requires base 2, 8, or 16 with fixed places"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked numeral-system constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A positional numeral-system base.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Base(u8);

impl Base {
    /// Binary output.
    pub const BINARY: Self = Self(2);
    /// Octal output.
    pub const OCTAL: Self = Self(8);
    /// Decimal output.
    pub const DECIMAL: Self = Self(10);
    /// Hexadecimal output.
    pub const HEXADECIMAL: Self = Self(16);
    /// The largest supported base.
    pub const MAXIMUM: Self = Self(MAXIMUM_BASE);

    /// Validates and constructs a base.
    ///
    /// # Errors
    ///
    /// Returns [`Error::BaseOutOfRange`] when the base is not between 2 and
    /// 36 inclusive.
    pub const fn new(value: u8) -> Result<Self> {
        if value < MINIMUM_BASE || value > MAXIMUM_BASE {
            return Err(Error::BaseOutOfRange {
                value,
                minimum: MINIMUM_BASE,
                maximum: MAXIMUM_BASE,
            });
        }
        Ok(Self(value))
    }

    /// Returns the numeric base.
    #[must_use]
    pub const fn value(self) -> u8 {
        self.0
    }

    const fn supports_twos_complement(self) -> bool {
        matches!(self.0, 2 | 8 | 16)
    }
}

impl Default for Base {
    fn default() -> Self {
        Self::DECIMAL
    }
}

impl TryFrom<u8> for Base {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// A checked nonzero fixed digit width.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FixedPlaces(u8);

impl FixedPlaces {
    /// One displayed digit.
    pub const ONE: Self = Self(1);
    /// Eight displayed digits.
    pub const EIGHT: Self = Self(8);
    /// The largest supported fixed width.
    pub const MAXIMUM: Self = Self(MAXIMUM_PLACES);

    /// Validates and constructs a fixed digit width.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PlacesOutOfRange`] for zero or values above the
    /// supported width.
    pub const fn new(value: u8) -> Result<Self> {
        if value == 0 || value > MAXIMUM_PLACES {
            return Err(Error::PlacesOutOfRange {
                value,
                maximum: MAXIMUM_PLACES,
            });
        }
        Ok(Self(value))
    }

    /// Returns the fixed digit width.
    #[must_use]
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for FixedPlaces {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// Minimum-width or fixed-width numeral output.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Places {
    /// Display only the digits required by the value.
    #[default]
    Minimum,
    /// Apply a fixed digit width.
    Fixed(FixedPlaces),
}

impl Places {
    /// Validates and constructs a fixed-width setting.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PlacesOutOfRange`] for an invalid width.
    pub const fn fixed(value: u8) -> Result<Self> {
        match FixedPlaces::new(value) {
            Ok(fixed_places) => Ok(Self::Fixed(fixed_places)),
            Err(error) => Err(error),
        }
    }
}

/// Representation of negative values in a positional system.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum NegativeStyle {
    /// Prefix the magnitude with a minus sign.
    #[default]
    MinusSign,
    /// Encode the rounded integer as fixed-width two's complement.
    TwosComplement,
}

/// Positional numeral-system display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct NumeralSystem {
    base: Base,
    places: Places,
    negative_style: NegativeStyle,
}

impl NumeralSystem {
    /// Validates and constructs a complete positional display format.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidTwosComplement`] when two's-complement output
    /// is requested without a supported fixed-width base.
    pub const fn new(base: Base, places: Places, negative_style: NegativeStyle) -> Result<Self> {
        if matches!(negative_style, NegativeStyle::TwosComplement)
            && (!base.supports_twos_complement() || matches!(places, Places::Minimum))
        {
            return Err(Error::InvalidTwosComplement);
        }
        Ok(Self {
            base,
            places,
            negative_style,
        })
    }

    /// Returns the positional base.
    #[must_use]
    pub const fn base(self) -> Base {
        self.base
    }

    /// Returns the minimum or fixed digit width.
    #[must_use]
    pub const fn places(self) -> Places {
        self.places
    }

    /// Returns the negative-value representation.
    #[must_use]
    pub const fn negative_style(self) -> NegativeStyle {
        self.negative_style
    }

    /// Validates and replaces the base.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidTwosComplement`] if the replacement would make
    /// the current negative style invalid.
    pub const fn with_base(self, base: Base) -> Result<Self> {
        Self::new(base, self.places, self.negative_style)
    }

    /// Validates and replaces the digit width.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidTwosComplement`] if the replacement would make
    /// the current negative style invalid.
    pub const fn with_places(self, places: Places) -> Result<Self> {
        Self::new(self.base, places, self.negative_style)
    }

    /// Validates and replaces the negative-value representation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidTwosComplement`] when the selected base and
    /// width cannot represent two's-complement values.
    pub const fn with_negative_style(self, negative_style: NegativeStyle) -> Result<Self> {
        Self::new(self.base, self.places, negative_style)
    }
}

impl Default for NumeralSystem {
    fn default() -> Self {
        Self {
            base: Base::DECIMAL,
            places: Places::Minimum,
            negative_style: NegativeStyle::MinusSign,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn base_and_width_validation_rejects_invalid_values() {
        assert!(matches!(
            Base::new(1),
            Err(Error::BaseOutOfRange { value: 1, .. })
        ));
        assert!(matches!(
            FixedPlaces::new(0),
            Err(Error::PlacesOutOfRange { value: 0, .. })
        ));
        assert!(matches!(
            NumeralSystem::new(
                Base::DECIMAL,
                Places::Fixed(FixedPlaces::EIGHT),
                NegativeStyle::TwosComplement,
            ),
            Err(Error::InvalidTwosComplement)
        ));
    }

    #[test]
    fn positional_values_round_trip_through_accessors() {
        let value = NumeralSystem::new(
            Base::HEXADECIMAL,
            Places::Fixed(FixedPlaces::EIGHT),
            NegativeStyle::TwosComplement,
        );
        let Ok(value) = value else {
            panic!("hexadecimal two's-complement format should be valid");
        };
        assert_eq!(value.base(), Base::HEXADECIMAL);
        assert_eq!(value.places(), Places::Fixed(FixedPlaces::EIGHT));
        assert_eq!(value.negative_style(), NegativeStyle::TwosComplement);
    }
}
