//! Common types, traits, and utilities shared across formats.
//!
//! This module provides unified types and traits used by both OLE2 (legacy)
//! and OOXML (modern) implementations, ensuring a consistent API for users.

// Submodule declarations
pub mod binary;
pub mod bom;
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
/// XML utilities
pub mod xml;
/// Shared byte slice for zero-copy element storage across formats
pub mod xml_slice;
// ID generation utilities
pub mod id;

// Re-exports for convenience
pub use bom::{
    BomKind, UTF8_BOM, UTF16_BE_BOM, UTF16_LE_BOM, UTF32_BE_BOM, UTF32_LE_BOM, strip_bom, write_bom,
};
pub use detection::{FileFormat, detect_file_format, detect_file_format_from_bytes};
pub use error::{Error, Result};
pub use metadata::Metadata;
pub use shapes::{PlaceholderType, ShapeType};
pub use style::{Length, RGBColor, VerticalPosition};
// Unit conversions
pub use unit::{Length as MeasuredLength, LengthUnit};
// Shared slice types
pub use xml_slice::{XmlArenaBuilder, XmlSlice};
