//! Typed custom document properties shared by every OOXML host format.
//!
//! [`Props`] owns a bounded set of named [`Value`]s and can read or write the
//! package-level custom-properties part. Parsing is namespace-aware and rejects
//! ambiguous package graphs and malformed XML instead of treating corruption
//! as absence. A package read degrades per property: a property Word would
//! ignore (unsupported type, unparseable value, bad identity, duplicate) is
//! skipped and recorded in [`Props::tolerance_report`] (LENIENCY_AUDIT 8.3).

mod codec;
mod model;
mod package;
mod reserved;
mod schema;

#[cfg(test)]
mod tests;

pub use model::{Props, Value};
pub use reserved::Host;
