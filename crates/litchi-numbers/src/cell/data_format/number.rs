//! Decimal, currency, percentage, scientific, and fraction values.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::number`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::number::*;

/// Selector-first exact-source transactions for one existing table cell's
/// decimal Number format.
pub mod transaction {
    pub use crate::package::table_cell_number_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
