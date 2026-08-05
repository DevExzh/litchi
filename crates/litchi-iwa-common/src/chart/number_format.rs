//! Strict, archive-free number-format values for chart labels.
//!
//! The wire codec remains in `litchi-iwa`; these values contain only the
//! semantic settings shared by axes and series labels. `LabelAffixes` is the
//! only allocating value because native affixes are arbitrary UTF-8 strings.

const MAXIMUM_DECIMAL_PLACES: u8 = 30;
const DECIMAL_PLACES_MASK: u8 = 0x1f;
const PARENTHESES_MASK: u8 = 0x20;
const THOUSANDS_SEPARATOR_MASK: u8 = 0x40;
const MAXIMUM_LABEL_AFFIX_BYTES: usize = 4 * 1_024;

/// Validation failures for fixed chart decimal places.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The requested number exceeds the native inspector range.
    #[error("chart decimal places {value} must not exceed {maximum}")]
    DecimalPlacesOutOfRange { value: u8, maximum: u8 },
    /// The combined chart-label affixes exceed the bounded semantic budget.
    #[error("chart label affixes use {bytes} bytes, maximum is {maximum}")]
    LabelAffixesTooLong { bytes: usize, maximum: usize },
}

/// Result type for chart number-format construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A fixed number of decimal places accepted by native iWork inspectors.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FixedDecimalPlaces(u8);

impl FixedDecimalPlaces {
    /// No fractional digits.
    pub const ZERO: Self = Self(0);

    /// The largest value accepted by Pages, Numbers, and Keynote.
    pub const MAXIMUM: Self = Self(MAXIMUM_DECIMAL_PLACES);

    /// Build a fixed decimal-place count accepted by iWork.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalPlacesOutOfRange`] above the native maximum.
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

    /// Return the decimal-place count shown by iWork.
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

/// Automatic or fixed decimal places for native chart labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum DecimalPlaces {
    /// Let iWork derive the necessary number of fractional digits.
    #[default]
    Automatic,
    /// Always render exactly this many fractional digits.
    Fixed(FixedDecimalPlaces),
}

impl DecimalPlaces {
    /// Build a fixed decimal-place setting.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalPlacesOutOfRange`] above the native maximum.
    #[must_use = "use the validated decimal places or handle the validation error"]
    pub const fn fixed(value: u8) -> Result<Self> {
        match FixedDecimalPlaces::new(value) {
            Ok(validated) => Ok(Self::Fixed(validated)),
            Err(error) => Err(error),
        }
    }
}

/// Native negative-number presentation for chart labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum NegativeStyle {
    /// Render a leading minus sign, for example `-100`.
    #[default]
    MinusSign,
    /// Render the magnitude in parentheses, for example `(100)`.
    Parentheses,
}

/// Decimal number formatting applied to native chart labels.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct NumberFormat(u8);

impl NumberFormat {
    /// Native default for per-series value labels.
    pub const SERIES_VALUE_LABEL_NATIVE_DEFAULT: Self =
        Self::new(DecimalPlaces::Automatic, NegativeStyle::MinusSign, true);

    /// Native default for value-axis labels.
    pub const AXIS_NATIVE_DEFAULT: Self =
        Self::new(DecimalPlaces::Automatic, NegativeStyle::MinusSign, false);

    /// Construct a complete decimal-number format.
    #[must_use]
    pub const fn new(
        decimal_places: DecimalPlaces,
        negative_style: NegativeStyle,
        thousands_separator: bool,
    ) -> Self {
        let encoded_decimal_places = match decimal_places {
            DecimalPlaces::Automatic => 0,
            DecimalPlaces::Fixed(value) => value.value() + 1,
        };
        let encoded_negative_style = match negative_style {
            NegativeStyle::MinusSign => 0,
            NegativeStyle::Parentheses => PARENTHESES_MASK,
        };
        let encoded_thousands_separator = if thousands_separator {
            THOUSANDS_SEPARATOR_MASK
        } else {
            0
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
        if self.0 & PARENTHESES_MASK == 0 {
            NegativeStyle::MinusSign
        } else {
            NegativeStyle::Parentheses
        }
    }

    /// Whether labels include locale-aware thousands separators.
    #[must_use]
    pub const fn thousands_separator(self) -> bool {
        self.0 & THOUSANDS_SEPARATOR_MASK != 0
    }
}

/// Text placed immediately before and after native chart labels.
///
/// The two views share one allocation. This keeps the common value compact
/// without giving the archive decoder an unbounded allocation surface.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct LabelAffixes {
    text: Box<str>,
    prefix_length: usize,
}

impl LabelAffixes {
    /// The maximum combined UTF-8 length accepted for both affixes.
    pub const MAXIMUM_BYTES: usize = MAXIMUM_LABEL_AFFIX_BYTES;

    /// Construct bounded chart-label affixes in one allocation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::LabelAffixesTooLong`] when the combined UTF-8 length
    /// exceeds [`Self::MAXIMUM_BYTES`].
    #[must_use = "use the validated affixes or handle the validation error"]
    pub fn new(prefix: impl AsRef<str>, suffix: impl AsRef<str>) -> Result<Self> {
        let prefix_text = prefix.as_ref();
        let suffix_text = suffix.as_ref();
        let bytes = prefix_text.len().saturating_add(suffix_text.len());
        if bytes > Self::MAXIMUM_BYTES {
            return Err(Error::LabelAffixesTooLong {
                bytes,
                maximum: Self::MAXIMUM_BYTES,
            });
        }
        let mut text = String::with_capacity(bytes);
        text.push_str(prefix_text);
        text.push_str(suffix_text);
        Ok(Self {
            text: text.into_boxed_str(),
            prefix_length: prefix_text.len(),
        })
    }

    /// Text placed before each label.
    #[must_use]
    pub fn prefix(&self) -> &str {
        &self.text[..self.prefix_length]
    }

    /// Text placed after each label.
    #[must_use]
    pub fn suffix(&self) -> &str {
        &self.text[self.prefix_length..]
    }

    /// Whether neither affix contains text.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.text.is_empty()
    }
}

impl Default for LabelAffixes {
    fn default() -> Self {
        Self {
            text: Box::from(""),
            prefix_length: 0,
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{
        DecimalPlaces, Error, FixedDecimalPlaces, LabelAffixes, NegativeStyle, NumberFormat,
    };

    #[test]
    fn scalar_formats_are_compact_and_strict() {
        assert_eq!(size_of::<FixedDecimalPlaces>(), 1);
        assert_eq!(size_of::<DecimalPlaces>(), 2);
        assert_eq!(size_of::<NegativeStyle>(), 1);
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
        assert_eq!(DecimalPlaces::default(), DecimalPlaces::Automatic);
        assert_eq!(NegativeStyle::default(), NegativeStyle::MinusSign);
    }

    #[test]
    fn defaults_and_affixes_are_ergonomic() {
        assert_eq!(
            NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT
        );
        let format = NumberFormat::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            true,
        );
        assert_eq!(format.decimal_places(), DecimalPlaces::fixed(2).unwrap());
        assert_eq!(format.negative_style(), NegativeStyle::Parentheses);
        assert!(format.thousands_separator());
        assert_eq!(size_of::<LabelAffixes>(), 24);
        let affixes = LabelAffixes::new("$", " net").unwrap();
        assert_eq!(affixes.prefix(), "$");
        assert_eq!(affixes.suffix(), " net");
        assert!(!affixes.is_empty());
        assert!(LabelAffixes::default().is_empty());
        let oversized = "x".repeat(LabelAffixes::MAXIMUM_BYTES + 1);
        assert_eq!(
            LabelAffixes::new(oversized, ""),
            Err(Error::LabelAffixesTooLong {
                bytes: LabelAffixes::MAXIMUM_BYTES + 1,
                maximum: LabelAffixes::MAXIMUM_BYTES,
            })
        );

        let affixes = LabelAffixes::new("€", " / net").unwrap();
        assert_eq!(affixes.prefix(), "€");
        assert_eq!(affixes.suffix(), " / net");
    }
}
