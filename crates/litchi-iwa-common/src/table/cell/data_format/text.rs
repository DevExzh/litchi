//! Literal-text display values.
//!
//! [`Text`] is a marker value: it carries no native identifiers or archive
//! state. The transaction namespace is kept beside the semantic marker so a
//! caller can discover the selector-first package API without importing the
//! package's private native implementation.

pub use super::Text;
