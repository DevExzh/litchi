//! Core ODF parsing functionality.
//!
//! This module provides the fundamental building blocks for parsing OpenDocument files.
//! It re-exports the main components from sibling modules for convenience.

/// ODF package handling
mod package;
/// ODF manifest parsing
mod manifest;
/// ODF XML utilities
mod xml;
/// ODF metadata parsing
mod metadata;

// Re-export main types for convenience
pub use package::Package;
pub use manifest::Manifest;
pub use xml::{Content, Styles, Meta};
