//! Fraction display values and selector-first transactions.
//!
//! Fraction is a distinct semantic family even though Numbers stores its
//! denominator strategy beside the other numeric display formats. The
//! archive-free nominal values remain compatible with the decimal value
//! module while this namespace owns the focused Fraction transaction entry
//! point.

pub use super::number::{Fraction, FractionAccuracy};

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Fraction format.
pub mod transaction {
    pub use crate::package::table_cell_fraction_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}
