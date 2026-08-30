//! Checked row and column sizing values for Numbers tables.
//!
//! The archive-free vocabulary is shared by concrete iWork format crates.
//! Numbers retains this module path as a compatibility reexport; its exact
//! source transaction remains in [`crate::table::dimension::transaction`].

pub use litchi_iwa_common::table::dimension::{Dimension, Error, Points, Size};

/// Exact-source row and column size transactions.
pub mod transaction;
