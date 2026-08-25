//! Shared iWork table sort vocabulary exposed through the Numbers crate.
//!
//! The neutral owner is [`litchi_iwa_common::table::sort`]. Numbers keeps this
//! focused module as its direct semantic entry point while native protobuf
//! decoding and package mutation remain in the Numbers package adapter.

/// Exact-source transactions for one rooted Numbers table's persisted sort
/// configuration.
pub mod transaction {
    pub use crate::package::table_sort::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}

pub use litchi_iwa_common::table::sort::{
    ColumnIndex, Direction, Error, Order, Result, RowRange, Rule, Scope,
};
