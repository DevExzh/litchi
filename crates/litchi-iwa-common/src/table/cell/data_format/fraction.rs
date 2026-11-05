//! Fraction display values.
//!
//! Fraction is a distinct semantic family even though Numbers stores its
//! denominator strategy beside the other numeric display formats. The
//! archive-free nominal values remain compatible with the decimal value
//! module while this namespace owns the focused Fraction transaction entry
//! point.

pub use super::number::{Fraction, FractionAccuracy};
