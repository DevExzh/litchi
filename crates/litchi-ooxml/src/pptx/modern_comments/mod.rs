//! Compatibility adapter for the canonical PPTX modern-comments owner.
//!
//! Typed comments, authors, bounded XML codecs, and package CRUD live in
//! `litchi_pptx::modern_comments`. This adapter preserves the historical host
//! module path and error boundary.

pub use litchi_pptx::modern_comments::*;
