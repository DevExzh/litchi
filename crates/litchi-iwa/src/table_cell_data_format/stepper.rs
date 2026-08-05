use super::numeric_control::{NumericControlRange, TableCellNumericControlDisplayFormat};
use crate::Result;

const NATIVE_DEFAULT_MINIMUM: f64 = 1.0;
const NATIVE_DEFAULT_MAXIMUM: f64 = 100.0;
const NATIVE_DEFAULT_INCREMENT: f64 = 1.0;

/// Numeric presentation nested inside an interactive Stepper format.
pub type TableCellStepperDisplayFormat = TableCellNumericControlDisplayFormat;

/// Finite numeric range used by an interactive Stepper table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCellStepperRange(NumericControlRange);

impl TableCellStepperRange {
    /// The range created by iWork when Stepper is applied to an empty cell.
    pub const NATIVE_DEFAULT: Self = Self(NumericControlRange::new_unchecked(
        NATIVE_DEFAULT_MINIMUM,
        NATIVE_DEFAULT_MAXIMUM,
        NATIVE_DEFAULT_INCREMENT,
    ));

    /// Validate a finite, increasing range with a positive increment.
    pub fn new(minimum: f64, maximum: f64, increment: f64) -> Result<Self> {
        NumericControlRange::new(minimum, maximum, increment).map(Self)
    }

    /// Return the inclusive minimum value.
    pub const fn minimum(self) -> f64 {
        self.0.minimum()
    }

    /// Return the inclusive maximum value.
    pub const fn maximum(self) -> f64 {
        self.0.maximum()
    }

    /// Return the positive step size.
    pub const fn increment(self) -> f64 {
        self.0.increment()
    }

    #[cfg(test)]
    pub(crate) const fn native_initial_value(self) -> f64 {
        self.minimum()
    }
}

impl Default for TableCellStepperRange {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

/// Native interactive Stepper format for one table cell.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct TableCellStepperFormat {
    range: TableCellStepperRange,
    display_format: TableCellStepperDisplayFormat,
}

impl TableCellStepperFormat {
    /// Construct a Stepper from a validated range and numeric presentation.
    pub const fn new(
        range: TableCellStepperRange,
        display_format: TableCellStepperDisplayFormat,
    ) -> Self {
        Self {
            range,
            display_format,
        }
    }

    /// Return the interactive range.
    pub const fn range(&self) -> TableCellStepperRange {
        self.range
    }

    /// Borrow the nested numeric presentation.
    pub const fn display_format(&self) -> &TableCellStepperDisplayFormat {
        &self.display_format
    }

    /// Replace the interactive range.
    pub fn with_range(mut self, range: TableCellStepperRange) -> Self {
        self.range = range;
        self
    }

    /// Replace the nested numeric presentation.
    pub fn with_display_format(mut self, display_format: TableCellStepperDisplayFormat) -> Self {
        self.display_format = display_format;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::table_cell_data_format::TableCellFractionFormat;

    #[test]
    fn native_default_starts_at_the_minimum() {
        assert_eq!(
            TableCellStepperRange::NATIVE_DEFAULT.native_initial_value(),
            1.0
        );
    }

    #[test]
    fn display_formats_compose_without_losing_strong_types() {
        let range = TableCellStepperRange::new(-10.0, 30.0, 0.5).unwrap();
        let format = TableCellStepperFormat::new(range, TableCellFractionFormat::default().into());
        assert_eq!(format.range(), range);
        assert!(matches!(
            format.display_format(),
            TableCellStepperDisplayFormat::Fraction(_)
        ));
    }
}
