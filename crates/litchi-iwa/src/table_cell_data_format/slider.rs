use super::{
    TableCellCurrencyFormat, TableCellFractionFormat, TableCellNumberFormat,
    TableCellNumeralSystemFormat, TableCellPercentageFormat, TableCellScientificFormat,
};
use crate::{Error, Result};
use std::hash::{Hash, Hasher};

const NATIVE_DEFAULT_MINIMUM: f64 = 1.0;
const NATIVE_DEFAULT_MAXIMUM: f64 = 100.0;
const NATIVE_DEFAULT_INCREMENT: f64 = 1.0;

/// Finite numeric range used by an interactive Slider table cell.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellSliderRange {
    minimum: f64,
    maximum: f64,
    increment: f64,
}

impl TableCellSliderRange {
    /// The range created by iWork when Slider is applied to an empty cell.
    pub const NATIVE_DEFAULT: Self = Self {
        minimum: NATIVE_DEFAULT_MINIMUM,
        maximum: NATIVE_DEFAULT_MAXIMUM,
        increment: NATIVE_DEFAULT_INCREMENT,
    };

    /// Validate a finite, increasing range with a positive increment.
    pub fn new(minimum: f64, maximum: f64, increment: f64) -> Result<Self> {
        if !minimum.is_finite() || !maximum.is_finite() || !increment.is_finite() {
            return Err(Error::InvalidFormat(
                "Slider minimum, maximum, and increment must be finite".to_owned(),
            ));
        }
        if minimum >= maximum {
            return Err(Error::InvalidFormat(
                "Slider minimum must be less than its maximum".to_owned(),
            ));
        }
        if increment <= 0.0 {
            return Err(Error::InvalidFormat(
                "Slider increment must be positive".to_owned(),
            ));
        }
        let span = maximum - minimum;
        if !span.is_finite() || !(span / increment).is_finite() {
            return Err(Error::InvalidFormat(
                "Slider range cannot be represented with the requested increment".to_owned(),
            ));
        }
        Ok(Self {
            minimum: normalize_zero(minimum),
            maximum: normalize_zero(maximum),
            increment: normalize_zero(increment),
        })
    }

    /// Return the inclusive minimum value.
    pub const fn minimum(self) -> f64 {
        self.minimum
    }

    /// Return the inclusive maximum value.
    pub const fn maximum(self) -> f64 {
        self.maximum
    }

    /// Return the positive step size.
    pub const fn increment(self) -> f64 {
        self.increment
    }

    pub(crate) fn native_initial_value(self) -> f64 {
        let midpoint_steps = ((self.maximum - self.minimum) / self.increment / 2.0).floor();
        self.minimum + midpoint_steps * self.increment
    }
}

impl Default for TableCellSliderRange {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

impl Eq for TableCellSliderRange {}

impl Hash for TableCellSliderRange {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.minimum.to_bits().hash(state);
        self.maximum.to_bits().hash(state);
        self.increment.to_bits().hash(state);
    }
}

/// Numeric presentation nested inside an interactive Slider format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum TableCellSliderDisplayFormat {
    /// Display the current value as a decimal number.
    Number(TableCellNumberFormat),
    /// Display the current value as currency.
    Currency(TableCellCurrencyFormat),
    /// Display the current value as a percentage.
    Percentage(TableCellPercentageFormat),
    /// Display the current value as a mixed fraction.
    Fraction(TableCellFractionFormat),
    /// Display the current value in scientific notation.
    Scientific(TableCellScientificFormat),
    /// Display the current value in a positional numeral system.
    NumeralSystem(TableCellNumeralSystemFormat),
}

impl Default for TableCellSliderDisplayFormat {
    fn default() -> Self {
        Self::Number(TableCellNumberFormat::default())
    }
}

macro_rules! slider_display_from {
    ($source:ty, $variant:ident) => {
        impl From<$source> for TableCellSliderDisplayFormat {
            fn from(value: $source) -> Self {
                Self::$variant(value)
            }
        }
    };
}

slider_display_from!(TableCellNumberFormat, Number);
slider_display_from!(TableCellCurrencyFormat, Currency);
slider_display_from!(TableCellPercentageFormat, Percentage);
slider_display_from!(TableCellFractionFormat, Fraction);
slider_display_from!(TableCellScientificFormat, Scientific);
slider_display_from!(TableCellNumeralSystemFormat, NumeralSystem);

/// Native interactive Slider format for one table cell.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct TableCellSliderFormat {
    range: TableCellSliderRange,
    display_format: TableCellSliderDisplayFormat,
}

impl TableCellSliderFormat {
    /// Construct a Slider from a validated range and numeric presentation.
    pub const fn new(
        range: TableCellSliderRange,
        display_format: TableCellSliderDisplayFormat,
    ) -> Self {
        Self {
            range,
            display_format,
        }
    }

    /// Return the interactive range.
    pub const fn range(&self) -> TableCellSliderRange {
        self.range
    }

    /// Borrow the nested numeric presentation.
    pub const fn display_format(&self) -> &TableCellSliderDisplayFormat {
        &self.display_format
    }

    /// Replace the interactive range.
    pub fn with_range(mut self, range: TableCellSliderRange) -> Self {
        self.range = range;
        self
    }

    /// Replace the nested numeric presentation.
    pub fn with_display_format(mut self, display_format: TableCellSliderDisplayFormat) -> Self {
        self.display_format = display_format;
        self
    }
}

const fn normalize_zero(value: f64) -> f64 {
    if value == 0.0 { 0.0 } else { value }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges_are_finite_ordered_and_have_positive_increments() {
        let range = TableCellSliderRange::new(-10.0, 30.0, 0.5).unwrap();
        assert_eq!(range.minimum(), -10.0);
        assert_eq!(range.maximum(), 30.0);
        assert_eq!(range.increment(), 0.5);
        assert_eq!(range.native_initial_value(), 10.0);
        assert_eq!(
            TableCellSliderRange::NATIVE_DEFAULT.native_initial_value(),
            50.0
        );

        for invalid in [
            TableCellSliderRange::new(f64::NAN, 1.0, 1.0),
            TableCellSliderRange::new(0.0, f64::INFINITY, 1.0),
            TableCellSliderRange::new(1.0, 1.0, 1.0),
            TableCellSliderRange::new(2.0, 1.0, 1.0),
            TableCellSliderRange::new(0.0, 1.0, 0.0),
            TableCellSliderRange::new(0.0, 1.0, -1.0),
            TableCellSliderRange::new(-f64::MAX, f64::MAX, 1.0),
            TableCellSliderRange::new(0.0, 1.0, f64::from_bits(1)),
        ] {
            assert!(invalid.is_err());
        }
    }

    #[test]
    fn display_formats_compose_without_losing_strong_types() {
        let range = TableCellSliderRange::new(0.0, 10.0, 0.25).unwrap();
        let format = TableCellSliderFormat::new(range, TableCellScientificFormat::default().into());
        assert_eq!(format.range(), range);
        assert!(matches!(
            format.display_format(),
            TableCellSliderDisplayFormat::Scientific(_)
        ));
    }
}
