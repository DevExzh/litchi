//! Literal-text display values and selector-first transactions.
//!
//! [`Text`] is a marker value: it carries no native identifiers or archive
//! state. The transaction namespace is kept beside the semantic marker so a
//! caller can discover the selector-first package API without importing the
//! package's private native implementation.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::text`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::text::Text;

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Text format.
pub mod transaction {
    pub use crate::package::table_cell_text_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
