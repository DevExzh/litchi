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
        BodyTableSortCommit as Commit, BodyTableSortDiagnostics as Diagnostics,
        BodyTableSortEdit as Edit, BodyTableSortError as Error,
        BodyTableSortLimitKind as LimitKind, BodyTableSortPatch as Patch,
        BodyTableSortPath as Path,
    };
}

pub use litchi_iwa_common::table::sort::{
    ColumnIndex, Direction, Error, Order, Result, RowRange, Rule, Scope,
};
