//! Compatibility re-exports for the canonical DOCX field model.
//!
//! Field semantics and XML decoding live in `litchi_docx::field`. This
//! module remains at the historical host path so existing OOXML callers keep
//! compiling while the host only owns package/document orchestration.

pub use litchi_docx::field::*;
