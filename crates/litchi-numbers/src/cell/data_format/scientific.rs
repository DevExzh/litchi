//! Scientific-notation display values and selector-first transactions.
//!
//! Scientific notation is a distinct semantic family even though Numbers
//! stores its fixed precision beside the plain Number and Percentage formats.
//! The nominal type remains compatible with the decimal value module while
//! this namespace owns the focused Scientific transaction entry point.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::scientific`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::scientific::{FixedDecimalPlaces, Scientific};

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Scientific format.
pub mod transaction {
    pub use crate::package::table_cell_scientific_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
