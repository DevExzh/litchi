//! Archive-free Keynote slide-table title settings.

/// Lossless optional visibility and outline values shared across iWork tables.
pub use litchi_iwa_common::table::title::Settings;

/// Exact-source transactions for one slide table's title settings.
pub mod transaction {
    pub use crate::package::slide_table_title::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}
