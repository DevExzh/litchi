//! Currency display values and selector-first transactions.
//!
//! Currency is a distinct semantic family even though Numbers stores its
//! decimal settings beside the plain Number and Percentage formats.  The
//! nominal types are re-exported from the compatibility decimal module while
//! this module owns the focused Currency transaction entry point.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::currency`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::currency::{
    Currency, CurrencyCode, CurrencyStyle, DecimalPlaces, NegativeStyle, ThousandsSeparator,
};

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Currency format.
pub mod transaction {
    pub use crate::package::table_cell_currency_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
