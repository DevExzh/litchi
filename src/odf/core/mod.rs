//! Core ODF parsing functionality.
//!
//! This module provides the fundamental building blocks for parsing OpenDocument files.
//! It re-exports the main components from sibling modules for convenience.

/// ODF manifest parsing
mod manifest;
/// ODF metadata parsing
mod metadata;
/// ODF package handling
mod package;
/// ODF XML utilities
mod xml;

// Re-export main types for convenience
// Manifest is internal to the package system
#[allow(unused_imports)]
pub use manifest::Manifest;
pub use package::Package;
pub use xml::{Content, Meta, Styles};
