//! Literal-text display values and selector-first transactions.
//!
//! [`Text`] is a marker value: it carries no native identifiers or archive
//! state. The transaction namespace is kept beside the semantic marker so a
//! caller can discover the selector-first package API without importing the
//! package's private native implementation.

pub use super::Text;

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Text format.
pub mod transaction {
    pub use crate::package::table_cell_text_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}
