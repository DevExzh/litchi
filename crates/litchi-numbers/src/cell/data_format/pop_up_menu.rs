//! Pop-Up Menu cell-control values.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::pop_up_menu`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::pop_up_menu::*;

/// Selector-first exact-source transactions for one table cell's pop-up menu.
///
/// The transaction implementation remains in the Numbers package adapter;
/// this semantic namespace exposes only archive-free values and transaction
/// handles. Native object identifiers, table-list records, and protobuf
/// payloads never cross this boundary.
pub mod transaction {
    pub use crate::package::table_cell_pop_up_menu::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
