//! Common types, traits, and utilities shared across formats.
//!
//! This module provides unified types and traits used by both OLE2 (legacy)
//! and OOXML (modern) implementations, ensuring a consistent API for users.

// Submodule declarations
pub mod binary;
pub mod detection;
#[cfg(any(feature = "ole", feature = "rtf"))]
pub mod encoding;
pub mod error;
pub mod metadata;
pub mod shapes;
pub mod simd;
pub mod style;
/// Common unit conversion utilities (length units used across all formats)
pub mod unit;
/// Shared byte slice for zero-copy element storage across formats
pub mod xml_slice;

// Re-exports for convenience
pub use detection::{FileFormat, detect_file_format, detect_file_format_from_bytes};
pub use error::{Error, Result};
pub use metadata::Metadata;
pub use shapes::{PlaceholderType, ShapeType};
pub use style::{Length, RGBColor, VerticalPosition};
// Unit conversions
pub use unit::{Length as MeasuredLength, LengthUnit};
// Shared slice types
pub use xml_slice::{XmlArenaBuilder, XmlSlice};
