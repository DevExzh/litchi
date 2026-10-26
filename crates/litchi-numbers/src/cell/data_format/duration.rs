//! Duration display values.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::duration`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::duration::*;

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Duration display format.
pub mod transaction {
    pub use crate::package::table_cell_duration_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
