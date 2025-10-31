//! OpenDocument Text (.odt) implementation.
//!
//! This module provides comprehensive support for parsing and working with
//! OpenDocument Text documents (.odt files), which are the open standard
//! equivalent of Microsoft Word documents.

mod document;
mod parser;

pub use document::Document;

// Re-export ODT-specific types for external use
#[allow(unused_imports)] // Library public API
pub use parser::{ChangeType, Comment, Section, TrackChange};
