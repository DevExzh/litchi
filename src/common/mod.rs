//! Common types, traits, and utilities shared across formats.
//!
//! This module provides unified types and traits used by both OLE2 (legacy)
//! and OOXML (modern) implementations, ensuring a consistent API for users.

// Submodule declarations
pub mod binary;
pub mod detection;
pub mod error;
pub mod metadata;
pub mod shapes;
pub mod style;

// Re-exports for convenience
pub use detection::{FileFormat, detect_file_format, detect_file_format_from_bytes};
pub use error::{Error, Result};
pub use metadata::Metadata;
pub use shapes::{PlaceholderType, ShapeType};
pub use style::{Length, RGBColor, VerticalPosition};
