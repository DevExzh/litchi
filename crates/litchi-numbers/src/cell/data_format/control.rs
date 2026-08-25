//! Shared numeric-control values for Slider and Stepper cells.

use super::DataFormat;
use super::number::{Currency, Fraction, Number, Percentage, Scientific};
use super::numeral_system::NumeralSystem;
pub use super::pop_up_menu::PopUpMenu;
pub use super::{Checkbox, StarRating};
use std::fmt;
use std::hash::{Hash, Hasher};

/// Exact-source selector-first transactions for one table cell's control.
///
/// The package adapter owns the rooted table/control graph, native wire
/// preservation, and publication. This semantic namespace exposes only the
/// archive-free transaction handles; native identifiers, generated messages,
/// and package bytes never cross the boundary.
pub mod transaction {
    pub use crate::package::table_cell_control::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
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

display_format_from!(Number, Number);
display_format_from!(Currency, Currency);
display_format_from!(Percentage, Percentage);
display_format_from!(Fraction, Fraction);
display_format_from!(Scientific, Scientific);
display_format_from!(NumeralSystem, NumeralSystem);

/// One archive-free interactive control attached to a Numbers table cell.
///
/// The marker controls reuse the existing [`Checkbox`] and [`StarRating`]
/// values; numeric controls reuse their validated [`Range`] and
/// [`DisplayFormat`] values; and Pop-Up Menu reuses the bounded
/// [`PopUpMenu`] value. Keeping those values as payloads preserves the
/// established `DataFormat` representation without creating duplicate
/// semantic models for the unified control owner.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum CellControl {
    /// A Boolean checkbox control.
    Checkbox(Checkbox),
    /// A fixed five-star rating control.
    StarRating(StarRating),
    /// A bounded numeric slider control.
    Slider(Slider),
    /// A bounded numeric stepper control.
    Stepper(Stepper),
    /// A compatibility Pop-Up Menu control.
    PopUpMenu(PopUpMenu),
}

impl CellControl {
    /// Borrow the checkbox marker when this is a checkbox control.
    #[must_use]
    pub const fn as_checkbox(&self) -> Option<&Checkbox> {
        match self {
            Self::Checkbox(value) => Some(value),
            _ => None,
        }
    }

    /// Borrow the star-rating marker when this is a star-rating control.
    #[must_use]
    pub const fn as_star_rating(&self) -> Option<&StarRating> {
        match self {
            Self::StarRating(value) => Some(value),
            _ => None,
        }
    }

    /// Borrow the slider settings when this is a slider control.
    #[must_use]
    pub const fn as_slider(&self) -> Option<&Slider> {
        match self {
            Self::Slider(value) => Some(value),
            _ => None,
        }
    }

    /// Borrow the stepper settings when this is a stepper control.
    #[must_use]
    pub const fn as_stepper(&self) -> Option<&Stepper> {
        match self {
            Self::Stepper(value) => Some(value),
            _ => None,
        }
    }

    /// Borrow the Pop-Up Menu value when this is a menu control.
    #[must_use]
    pub const fn as_pop_up_menu(&self) -> Option<&PopUpMenu> {
        match self {
            Self::PopUpMenu(value) => Some(value),
            _ => None,
        }
    }

    /// Convert this control to the established complete cell-data-format sum.
    #[must_use]
    pub fn into_data_format(self) -> DataFormat {
        self.into()
    }

    /// Borrow this control as the established complete cell-data-format sum.
    #[must_use]
    pub fn to_data_format(&self) -> DataFormat {
        self.clone().into_data_format()
    }
}

impl From<Checkbox> for CellControl {
    fn from(value: Checkbox) -> Self {
        Self::Checkbox(value)
    }
}

impl From<StarRating> for CellControl {
    fn from(value: StarRating) -> Self {
        Self::StarRating(value)
    }
}

impl From<Slider> for CellControl {
    fn from(value: Slider) -> Self {
        Self::Slider(value)
    }
}

impl From<Stepper> for CellControl {
    fn from(value: Stepper) -> Self {
        Self::Stepper(value)
    }
}

impl From<PopUpMenu> for CellControl {
    fn from(value: PopUpMenu) -> Self {
        Self::PopUpMenu(value)
    }
}

impl From<CellControl> for DataFormat {
    fn from(value: CellControl) -> Self {
        match value {
            CellControl::Checkbox(value) => Self::Checkbox(value),
            CellControl::StarRating(value) => Self::StarRating(value),
            CellControl::Slider(value) => Self::Slider(value),
            CellControl::Stepper(value) => Self::Stepper(value),
            CellControl::PopUpMenu(value) => Self::PopUpMenu(value),
        }
    }
}

/// Error returned when a complete cell data format is not an interactive
/// control.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NotCellControl;

impl fmt::Display for NotCellControl {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("cell data format is not an interactive control")
    }
}

impl std::error::Error for NotCellControl {}

impl TryFrom<DataFormat> for CellControl {
    type Error = NotCellControl;

    fn try_from(value: DataFormat) -> std::result::Result<Self, Self::Error> {
        match value {
            DataFormat::Checkbox(value) => Ok(Self::Checkbox(value)),
            DataFormat::StarRating(value) => Ok(Self::StarRating(value)),
            DataFormat::Slider(value) => Ok(Self::Slider(value)),
            DataFormat::Stepper(value) => Ok(Self::Stepper(value)),
            DataFormat::PopUpMenu(value) => Ok(Self::PopUpMenu(value)),
            DataFormat::Automatic
            | DataFormat::Number(_)
            | DataFormat::Text(_)
            | DataFormat::Currency(_)
            | DataFormat::Percentage(_)
            | DataFormat::Scientific(_)
            | DataFormat::Fraction(_)
            | DataFormat::NumeralSystem(_)
            | DataFormat::DateTime(_)
            | DataFormat::Duration(_)
            | DataFormat::Custom(_) => Err(NotCellControl),
        }
    }
}

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

    #[test]
    fn unified_controls_reuse_existing_data_format_values() {
        let Ok(range) = Range::new(0.0, 10.0, 1.0) else {
            panic!("finite increasing range should construct");
        };
        let controls = [
            CellControl::from(Checkbox),
            CellControl::from(StarRating),
            CellControl::from(Slider::new(range, DisplayFormat::default())),
            CellControl::from(Stepper::new(range, DisplayFormat::default())),
            CellControl::from(PopUpMenu::default()),
        ];

        for control in controls {
            let format = control.clone().into_data_format();
            assert_eq!(CellControl::try_from(format), Ok(control));
        }
        assert_eq!(
            CellControl::try_from(DataFormat::Automatic),
            Err(NotCellControl)
        );
    }
}
