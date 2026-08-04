//! Compatibility re-exports for the canonical DOCX semantic enumerations.
//!
//! The value types are owned by `litchi-docx`; this path remains as a narrow
//! migration shim for callers of the historical OOXML host.

pub use litchi_docx::enums::*;
