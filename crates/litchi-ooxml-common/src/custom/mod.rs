//! Typed custom document properties shared by every OOXML host format.
//!
//! [`Props`] owns a bounded set of named [`Value`]s and can read or write the
//! package-level custom-properties part. Parsing is namespace-aware and rejects
//! ambiguous package graphs and malformed property records instead of treating
//! corruption as absence.

mod codec;
mod model;
mod package;
mod schema;

#[cfg(test)]
mod tests;

pub use model::{Props, Value};
