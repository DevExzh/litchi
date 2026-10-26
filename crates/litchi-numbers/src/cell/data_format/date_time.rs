//! Validated date-and-time pattern values.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::date_time`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::date_time::*;

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Date & Time display format.
pub mod transaction {
    pub use crate::package::table_cell_date_time_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
