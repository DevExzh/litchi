//! Percentage display values.
//!
//! Percentage uses the same checked decimal settings as
//! [`number`](crate::table::cell::data_format::number),
//! but remains a distinct semantic family.  In particular, a percentage
//! transaction cannot accidentally edit a Number-format cell.

pub use super::number::{DecimalPlaces, NegativeStyle, Percentage, ThousandsSeparator};
