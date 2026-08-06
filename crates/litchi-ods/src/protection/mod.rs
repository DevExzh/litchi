//! Contextual ODS protection ownership.
//!
//! Wire attributes remain in `model::protection`; this facade adds immutable
//! package-context snapshots, source-checked transactions, and atomic XML
//! publication without evaluating passwords or enforcing a security policy.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use super::model::style_protection::{ConditionalStyle, Protection, Rule};
pub use model::{Document, Key, Permissions, Sheet, Style, Styles};
pub use transaction::{Commit, Snapshot, Transaction, key, update};

/// Cell-protection vocabulary is nested to keep the document/sheet facade
/// short while retaining an intuitive semantic path.
pub mod style {
    pub use super::super::model::style_protection::{Protection, Rule};
}
