//! Scientific-notation display values.
//!
//! Scientific notation is a distinct semantic family even though Numbers
//! stores its fixed precision beside the plain Number and Percentage formats.
//! The nominal type remains compatible with the decimal value module while
//! this namespace owns the focused Scientific transaction entry point.

pub use super::number::{FixedDecimalPlaces, Scientific};
