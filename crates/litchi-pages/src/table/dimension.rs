//! Checked, archive-free row and column dimension values.
//!
//! The value types are format-neutral and deliberately contain no package,
//! archive, protobuf, or native identifier state. Their implementation is
//! shared by the concrete iWork owners.

pub use litchi_iwa_common::table::dimension::{Dimension, Error, Points, Size};

/// Exact-source transactions for one rooted Pages body table's row and column dimensions.
pub mod transaction {
    pub use crate::package::body_table_dimension::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}
