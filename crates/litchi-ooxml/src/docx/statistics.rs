//! Facade re-exports for the canonical DOCX statistics owner.
//!
//! Document traversal remains in the OOXML package adapter, while the
//! immutable metrics value and text counters live in `litchi-docx`.

pub use litchi_docx::statistics::*;
