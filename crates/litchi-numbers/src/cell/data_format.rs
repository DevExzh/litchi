//! Archive-free display formats for Numbers cells.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format`
//! vocabulary (record 0368), re-exported here at their established paths
//! beside the Numbers transactions.
//!
//! This module is the semantic owner for cell display and control values. It
//! contains no protobuf messages, archive identifiers, package handles, or
//! transaction state. Native identifiers, codecs, unknown-field preservation,
//! and publication remain responsibilities of the concrete IWA adapter.

pub mod control;
/// Currency display values and exact-source transactions for existing cells.
pub mod currency;
pub mod custom;
pub mod date_time;
pub mod duration;
/// Fraction display values and selector-first transactions for existing cells.
pub mod fraction;
pub mod number;
pub mod numeral_system;
/// Percentage display values and exact-source transactions for existing cells.
pub mod percentage;
pub mod pop_up_menu;
/// Scientific-notation display values and selector-first transactions for existing cells.
pub mod scientific;
/// Literal-text display values and selector-first transactions.
pub mod text;

pub use litchi_iwa_common::table::cell::data_format::{Checkbox, DataFormat, StarRating, Text};

pub use control::{CellControl, NotCellControl, Slider, Stepper};
pub use currency::{Currency, CurrencyCode, CurrencyStyle};
pub use custom::Custom;
pub use date_time::DateTime;
pub use duration::Duration;
pub use fraction::{Fraction, FractionAccuracy};
pub use number::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, Percentage, ThousandsSeparator,
};
pub use numeral_system::NumeralSystem;
pub use pop_up_menu::PopUpMenu;
pub use scientific::Scientific;
