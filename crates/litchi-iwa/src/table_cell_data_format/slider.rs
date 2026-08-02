use super::numeric_control::{NumericControlRange, TableCellNumericControlDisplayFormat};
use crate::Result;

const NATIVE_DEFAULT_MINIMUM: f64 = 1.0;
const NATIVE_DEFAULT_MAXIMUM: f64 = 100.0;
const NATIVE_DEFAULT_INCREMENT: f64 = 1.0;

/// Numeric presentation nested inside an interactive Slider format.
pub type TableCellSliderDisplayFormat = TableCellNumericControlDisplayFormat;

/// Finite numeric range used by an interactive Slider table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCellSliderRange(NumericControlRange);

impl TableCellSliderRange {
    /// The range created by iWork when Slider is applied to an empty cell.
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

    pub(crate) fn native_initial_value(self) -> f64 {
        self.0.midpoint_grid_value()
    }
}

impl Default for TableCellSliderRange {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::table_cell_data_format::TableCellScientificFormat;

    #[test]
    fn native_default_uses_the_midpoint_grid_value() {
        assert_eq!(
            TableCellSliderRange::NATIVE_DEFAULT.native_initial_value(),
            50.0
        );
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
