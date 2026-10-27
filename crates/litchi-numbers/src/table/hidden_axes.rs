//! Archive-free hidden-row and hidden-column values for Numbers tables
//! (record 0379).
//!
//! The values are the shared iWork axis vocabulary; the native hidden-state
//! graph stays inside the package adapter and the shared table engine.

/// One zero-based row or column position in a table.
pub use litchi_iwa_common::table::axis::AxisIndex;
/// Validation failures while constructing a hidden-axis collection.
pub use litchi_iwa_common::table::axis::Error;
/// Canonical, duplicate-free hidden rows and columns.
pub use litchi_iwa_common::table::axis::HiddenAxes;
/// Result type for constructing hidden rows and columns.
pub use litchi_iwa_common::table::axis::Result;

/// Exact-source transactions for one table's hidden axes.
pub mod transaction {
    pub use crate::package::table_hidden_axes::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}
