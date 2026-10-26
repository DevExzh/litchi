//! Percentage display values and selector-first transactions.
//!
//! Percentage uses the same checked decimal settings as
//! [`number`](crate::cell::data_format::number),
//! but remains a distinct semantic family.  In particular, a percentage
//! transaction cannot accidentally edit a Number-format cell.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::percentage`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::percentage::*;

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Percentage format.
pub mod transaction {
    pub use crate::package::table_cell_percentage_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
