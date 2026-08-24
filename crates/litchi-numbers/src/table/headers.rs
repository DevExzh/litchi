//! Header, footer, and repeating-row/column semantics.
//!
//! The neutral values live in [`litchi_iwa_common::table::headers`]. Numbers
//! keeps this module as the established public semantic entry point while
//! exact-source transactions remain owned by the Numbers package adapter.

/// Exact-source transactions for one rooted table's header and footer settings.
pub mod transaction {
    pub use crate::package::table_headers::{
        Commit, Diagnostics, Edit, Error, InvalidReason, LimitKind, Patch, Path,
    };
}

pub use litchi_iwa_common::table::headers::{Count, Error, Settings};
