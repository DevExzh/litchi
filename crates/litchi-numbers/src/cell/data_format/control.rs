//! Shared numeric-control values for Slider and Stepper cells.

use super::number::{Currency, Fraction, Number, Percentage, Scientific};
use super::numeral_system::NumeralSystem;
use std::fmt;
use std::hash::{Hash, Hasher};

const DEFAULT_MINIMUM: f64 = 1.0;
const DEFAULT_MAXIMUM: f64 = 100.0;
const DEFAULT_INCREMENT: f64 = 1.0;

/// Errors returned by checked numeric-control range construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// At least one range component is not finite.
    NonFinite,
    /// The minimum must be strictly below the maximum.
    Reversed,
    /// The increment must be positive.
    NonPositiveIncrement,
    /// The span cannot be represented by the requested increment.
    Unrepresentable,
}
impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NonFinite => formatter.write_str("numeric-control range values must be finite"),
            Self::Reversed => formatter.write_str("numeric-control minimum must be below maximum"),
            Self::NonPositiveIncrement => {
                formatter.write_str("numeric-control increment must be positive")
            },
            Self::Unrepresentable => formatter.write_str(
                "numeric-control range cannot be represented with the requested increment",
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked numeric-control constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A finite increasing numeric range with a positive step.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Range {
    minimum: f64,
    maximum: f64,
    increment: f64,
}

impl Range {
    /// The default range used for newly authored controls.
    pub const DEFAULT: Self = Self {
        minimum: DEFAULT_MINIMUM,
        maximum: DEFAULT_MAXIMUM,
        increment: DEFAULT_INCREMENT,
    };

    /// Validates and constructs a numeric-control range.
    ///
    /// # Errors
    ///
    /// Returns a typed error for non-finite, reversed, non-positive, or
    /// unrepresentable values.
    pub fn new(minimum: f64, maximum: f64, increment: f64) -> Result<Self> {
        if !minimum.is_finite() || !maximum.is_finite() || !increment.is_finite() {
            return Err(Error::NonFinite);
        }
        if minimum >= maximum {
            return Err(Error::Reversed);
        }
        if increment <= 0.0 {
            return Err(Error::NonPositiveIncrement);
        }
        let span = maximum - minimum;
        if !span.is_finite() || !(span / increment).is_finite() {
            return Err(Error::Unrepresentable);
        }
        Ok(Self {
            minimum: normalize_zero(minimum),
            maximum: normalize_zero(maximum),
            increment: normalize_zero(increment),
        })
    }

    /// Returns the inclusive minimum.
    #[must_use]
    pub const fn minimum(self) -> f64 {
        self.minimum
    }

    /// Returns the inclusive maximum.
    #[must_use]
    pub const fn maximum(self) -> f64 {
        self.maximum
    }

    /// Returns the positive step size.
    #[must_use]
    pub const fn increment(self) -> f64 {
        self.increment
    }

    /// Returns the grid value at the arithmetic midpoint of the range.
    #[must_use]
    pub fn midpoint(self) -> f64 {
        let midpoint_steps = ((self.maximum - self.minimum) / self.increment / 2.0).floor();
        self.minimum + midpoint_steps * self.increment
    }
}

impl Eq for Range {}

impl Hash for Range {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.minimum.to_bits().hash(state);
        self.maximum.to_bits().hash(state);
        self.increment.to_bits().hash(state);
    }
}

impl Default for Range {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Numeric display nested inside an interactive control.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum DisplayFormat {
    /// Display the value as a decimal number.
    Number(Number),
    /// Display the value as currency.
    Currency(Currency),
    /// Display the value as a percentage.
    Percentage(Percentage),
    /// Display the value as a mixed fraction.
    Fraction(Fraction),
    /// Display the value in scientific notation.
    Scientific(Scientific),
    /// Display the value in a positional numeral system.
    NumeralSystem(NumeralSystem),
}

impl Default for DisplayFormat {
    fn default() -> Self {
        Self::Number(Number::default())
    }
}

macro_rules! display_format_from {
    ($type:ty, $variant:ident) => {
        impl From<$type> for DisplayFormat {
            fn from(value: $type) -> Self {
                Self::$variant(value)
            }
        }
    };
}

display_format_from!(Number, Number);
display_format_from!(Currency, Currency);
display_format_from!(Percentage, Percentage);
display_format_from!(Fraction, Fraction);
display_format_from!(Scientific, Scientific);
display_format_from!(NumeralSystem, NumeralSystem);

/// Numeric slider control format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Slider {
    range: Range,
    display_format: DisplayFormat,
}

impl Slider {
    /// Constructs a slider from a validated range and display format.
    #[must_use]
    pub const fn new(range: Range, display_format: DisplayFormat) -> Self {
        Self {
            range,
            display_format,
        }
    }

    /// Returns the interactive range.
    #[must_use]
    pub const fn range(&self) -> Range {
        self.range
    }

    /// Borrows the nested display format.
    #[must_use]
    pub const fn display_format(&self) -> &DisplayFormat {
        &self.display_format
    }

    /// Replaces the interactive range.
    #[must_use]
    pub const fn with_range(mut self, range: Range) -> Self {
        self.range = range;
        self
    }

    /// Replaces the nested display format.
    #[must_use]
    pub fn with_display_format(mut self, display_format: DisplayFormat) -> Self {
        self.display_format = display_format;
        self
    }
}

impl Default for Slider {
    fn default() -> Self {
        Self::new(Range::DEFAULT, DisplayFormat::default())
    }
}

/// Numeric stepper control format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Stepper {
    range: Range,
    display_format: DisplayFormat,
}

impl Stepper {
    /// Constructs a stepper from a validated range and display format.
    #[must_use]
    pub const fn new(range: Range, display_format: DisplayFormat) -> Self {
        Self {
            range,
            display_format,
        }
    }

    /// Returns the interactive range.
    #[must_use]
    pub const fn range(&self) -> Range {
        self.range
    }

    /// Borrows the nested display format.
    #[must_use]
    pub const fn display_format(&self) -> &DisplayFormat {
        &self.display_format
    }

    /// Replaces the interactive range.
    #[must_use]
    pub const fn with_range(mut self, range: Range) -> Self {
        self.range = range;
        self
    }

    /// Replaces the nested display format.
    #[must_use]
    pub fn with_display_format(mut self, display_format: DisplayFormat) -> Self {
        self.display_format = display_format;
        self
    }
}

impl Default for Stepper {
    fn default() -> Self {
        Self::new(Range::DEFAULT, DisplayFormat::default())
    }
}

const fn normalize_zero(value: f64) -> f64 {
    if value == 0.0 { 0.0 } else { value }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges_reject_non_finite_and_unusable_steps() {
        let invalid = [
            Range::new(f64::NAN, 1.0, 1.0),
            Range::new(0.0, f64::INFINITY, 1.0),
            Range::new(1.0, 1.0, 1.0),
            Range::new(2.0, 1.0, 1.0),
            Range::new(0.0, 1.0, 0.0),
            Range::new(0.0, 1.0, -1.0),
            Range::new(-f64::MAX, f64::MAX, 1.0),
            Range::new(0.0, 1.0, f64::from_bits(1)),
        ];
        assert!(invalid.iter().all(Result::is_err));
    }

    #[test]
    fn controls_round_trip_range_and_display_values() {
        let Ok(range) = Range::new(-10.0, 30.0, 0.5) else {
            panic!("finite increasing range should construct");
        };
        let slider = Slider::new(range, Scientific::default().into());
        assert_eq!(slider.range(), range);
        assert_eq!(slider.range().midpoint(), 10.0);
        assert!(matches!(
            slider.display_format(),
            DisplayFormat::Scientific(_)
        ));
        let stepper = Stepper::new(range, Fraction::default().into());
        assert_eq!(stepper.range(), range);
        assert!(matches!(
            stepper.display_format(),
            DisplayFormat::Fraction(_)
        ));
    }
}
