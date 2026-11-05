//! Fraction display values and selector-first transactions.
//!
//! Fraction is a distinct semantic family even though Numbers stores its
//! denominator strategy beside the other numeric display formats. The
//! archive-free nominal values remain compatible with the decimal value
//! module while this namespace owns the focused Fraction transaction entry
//! point.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::fraction`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::fraction::{Fraction, FractionAccuracy};

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Fraction format.
pub mod transaction {
    pub use crate::package::table_cell_fraction_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
