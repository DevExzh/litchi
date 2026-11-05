//! Dependency-free custom cell-format values.
//!
//! Semantic values retain exact user content through their getters, while
//! their [`Debug`](std::fmt::Debug) implementations expose only bounded shape metadata.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::custom`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::custom::{
    Condition, ConditionValue, Custom, DateTime, DateTimePattern, Error, MAX_NAME_BYTES,
    MAX_PATTERN_BYTES, MAX_RULES, Name, Number, NumberPattern, NumberRule, Result, Text,
};

/// Selector-first package transactions for an existing cell's document-level
/// Custom display format.
pub mod transaction {
    pub use crate::package::table_cell_custom_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
