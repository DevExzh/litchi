//! Archive-free display formats for Numbers cells.
//!
//! This module is the semantic owner for cell display and control values. It
//! contains no protobuf messages, archive identifiers, package handles, or
//! transaction state. Native identifiers, codecs, unknown-field preservation,
//! and publication remain responsibilities of the concrete IWA adapter.

pub mod control;
pub mod custom;
pub mod date_time;
pub mod duration;
pub mod number;
pub mod numeral_system;
pub mod pop_up_menu;

pub use custom::Custom;
pub use date_time::DateTime;
pub use duration::Duration;
pub use number::{
    Currency, CurrencyCode, CurrencyStyle, DecimalPlaces, FixedDecimalPlaces, Fraction,
    FractionAccuracy, NegativeStyle, Number, Percentage, Scientific, ThousandsSeparator,
};
pub use numeral_system::NumeralSystem;
pub use pop_up_menu::PopUpMenu;
pub use control::{Slider, Stepper};

/// Explicit display or interaction semantics for one cell.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub enum DataFormat {
    /// Let the spreadsheet choose a display format from the cell value.
    #[default]
    Automatic,
    /// Display a value as a decimal number.
    Number(Number),
    /// Preserve and display a value as literal text.
    Text(Text),
    /// Display a value using a currency code and decimal settings.
    Currency(Currency),
    /// Multiply a displayed value by one hundred and append a percent sign.
    Percentage(Percentage),
    /// Display a value in scientific notation.
    Scientific(Scientific),
    /// Display a value as a mixed fraction.
    Fraction(Fraction),
    /// Display a rounded value in a positional numeral system.
    NumeralSystem(NumeralSystem),
    /// Display a date or time using a validated pattern.
    DateTime(DateTime),
    /// Display a duration using typed units and a presentation style.
    Duration(Duration),
    /// Edit a value with a Boolean checkbox control.
    Checkbox(Checkbox),
    /// Edit a value with a fixed five-star rating control.
    StarRating(StarRating),
    /// Edit a value with a bounded numeric slider control.
    Slider(Slider),
    /// Edit a value with a bounded numeric stepper control.
    Stepper(Stepper),
    /// Edit a text value with an ordered pop-up menu.
    PopUpMenu(PopUpMenu),
    /// Display a value with a document-registered custom format.
    Custom(Custom),
}

/// Explicit literal-text display semantics for one cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Text;

/// Edit semantics for a Boolean checkbox cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Checkbox;

/// Edit semantics for a fixed five-star rating cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct StarRating;

macro_rules! data_format_from {
    ($type:ty, $variant:ident) => {
        impl From<$type> for DataFormat {
            fn from(value: $type) -> Self {
                Self::$variant(value)
            }
        }
    };
}

data_format_from!(Number, Number);
data_format_from!(Text, Text);
data_format_from!(Currency, Currency);
data_format_from!(Percentage, Percentage);
data_format_from!(Scientific, Scientific);
data_format_from!(Fraction, Fraction);
data_format_from!(NumeralSystem, NumeralSystem);
data_format_from!(DateTime, DateTime);
data_format_from!(Duration, Duration);
data_format_from!(Checkbox, Checkbox);
data_format_from!(StarRating, StarRating);
data_format_from!(Slider, Slider);
data_format_from!(Stepper, Stepper);
data_format_from!(PopUpMenu, PopUpMenu);
data_format_from!(Custom, Custom);

#[cfg(test)]
mod tests {
    use super::{Checkbox, DataFormat, StarRating, Text};

    #[test]
    fn marker_formats_round_trip_through_the_semantic_enum() {
        assert_eq!(DataFormat::from(Text), DataFormat::Text(Text));
        assert_eq!(DataFormat::from(Checkbox), DataFormat::Checkbox(Checkbox));
        assert_eq!(
            DataFormat::from(StarRating),
            DataFormat::StarRating(StarRating)
        );
        assert_eq!(DataFormat::default(), DataFormat::Automatic);
    }
}
