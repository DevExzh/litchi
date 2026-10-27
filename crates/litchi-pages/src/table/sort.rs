//! Archive-free table sort semantics for Pages body tables.
//!
//! The common sort values contain no package, archive, protobuf, or native
//! identifier state. Pages owns the rooted body-table selector and the
//! exact-source transaction; those transaction details remain behind the
//! private package adapter.

/// Exact-source transactions for one rooted Pages body table's persisted sort
/// configuration.
pub mod transaction {
    pub use crate::package::body_table_sort::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}

/// Exact-source execution of a body table's persisted sort order (the
/// body rows move; record 0371).
pub mod execution {
    pub use crate::package::body_table_sort_execution::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Unsupported,
    };
}

pub use litchi_iwa_common::table::sort::{
    ColumnIndex, Direction, Error, Order, Result, RowRange, Rule, Scope,
};
