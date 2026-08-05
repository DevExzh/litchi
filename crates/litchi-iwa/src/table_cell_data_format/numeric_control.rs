use super::{
    TableCellCurrencyFormat, TableCellFractionFormat, TableCellNumberFormat,
    TableCellNumeralSystemFormat, TableCellPercentageFormat, TableCellScientificFormat,
};
use crate::{Error, Result};
use std::hash::{Hash, Hasher};

/// Numeric presentation nested inside an interactive table-cell control.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum TableCellNumericControlDisplayFormat {
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

impl Default for TableCellNumericControlDisplayFormat {
    fn default() -> Self {
        Self::Number(TableCellNumberFormat::default())
    }
}

macro_rules! numeric_control_display_from {
    ($source:ty, $variant:ident) => {
        impl From<$source> for TableCellNumericControlDisplayFormat {
            fn from(value: $source) -> Self {
                Self::$variant(value)
            }
        }
    };
}

numeric_control_display_from!(TableCellNumberFormat, Number);
numeric_control_display_from!(TableCellCurrencyFormat, Currency);
numeric_control_display_from!(TableCellPercentageFormat, Percentage);
numeric_control_display_from!(TableCellFractionFormat, Fraction);
numeric_control_display_from!(TableCellScientificFormat, Scientific);
numeric_control_display_from!(TableCellNumeralSystemFormat, NumeralSystem);

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct NumericControlRange {
    minimum: f64,
    maximum: f64,
    increment: f64,
}

impl NumericControlRange {
    pub(crate) const fn new_unchecked(minimum: f64, maximum: f64, increment: f64) -> Self {
        Self {
            minimum,
            maximum,
            increment,
        }
    }

    pub(crate) fn new(minimum: f64, maximum: f64, increment: f64) -> Result<Self> {
        if !minimum.is_finite() || !maximum.is_finite() || !increment.is_finite() {
            return Err(Error::InvalidFormat(
                "Interactive control minimum, maximum, and increment must be finite".to_owned(),
            ));
        }
        if minimum >= maximum {
            return Err(Error::InvalidFormat(
                "Interactive control minimum must be less than its maximum".to_owned(),
            ));
        }
        if increment <= 0.0 {
            return Err(Error::InvalidFormat(
                "Interactive control increment must be positive".to_owned(),
            ));
        }
        let span = maximum - minimum;
        if !span.is_finite() || !(span / increment).is_finite() {
            return Err(Error::InvalidFormat(
                "Interactive control range cannot be represented with the requested increment"
                    .to_owned(),
            ));
        }
        Ok(Self {
            minimum: normalize_zero(minimum),
            maximum: normalize_zero(maximum),
            increment: normalize_zero(increment),
        })
    }

    pub(crate) const fn minimum(self) -> f64 {
        self.minimum
    }

    pub(crate) const fn maximum(self) -> f64 {
        self.maximum
    }

    pub(crate) const fn increment(self) -> f64 {
        self.increment
    }

    #[cfg(test)]
    pub(crate) fn midpoint_grid_value(self) -> f64 {
        let midpoint_steps = ((self.maximum - self.minimum) / self.increment / 2.0).floor();
        self.minimum + midpoint_steps * self.increment
    }
}

impl Eq for NumericControlRange {}

impl Hash for NumericControlRange {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.minimum.to_bits().hash(state);
        self.maximum.to_bits().hash(state);
        self.increment.to_bits().hash(state);
    }
}

const fn normalize_zero(value: f64) -> f64 {
    if value == 0.0 { 0.0 } else { value }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shared_ranges_are_finite_ordered_and_have_positive_increments() {
        let range = NumericControlRange::new(-10.0, 30.0, 0.5).unwrap();
        assert_eq!(range.minimum(), -10.0);
        assert_eq!(range.maximum(), 30.0);
        assert_eq!(range.increment(), 0.5);
        assert_eq!(range.midpoint_grid_value(), 10.0);

        for invalid in [
            NumericControlRange::new(f64::NAN, 1.0, 1.0),
            NumericControlRange::new(0.0, f64::INFINITY, 1.0),
            NumericControlRange::new(1.0, 1.0, 1.0),
            NumericControlRange::new(2.0, 1.0, 1.0),
            NumericControlRange::new(0.0, 1.0, 0.0),
            NumericControlRange::new(0.0, 1.0, -1.0),
            NumericControlRange::new(-f64::MAX, f64::MAX, 1.0),
            NumericControlRange::new(0.0, 1.0, f64::from_bits(1)),
        ] {
            assert!(invalid.is_err());
        }
    }
}
