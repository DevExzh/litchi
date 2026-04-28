// `missing_docs` is intentionally relaxed during the Phase 1 carve-out: the
// items hosted in this crate were previously private/`pub(crate)` inside the
// umbrella `litchi` crate and never had public doc coverage. A follow-up pass
// can tighten this once Phase 1 stabilises.
#![allow(missing_docs)]
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
pub mod sheet;
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

#[cfg(test)]
mod tests;

// Re-exports for convenience
pub use bom::{
    BomKind, UTF8_BOM, UTF16_BE_BOM, UTF16_LE_BOM, UTF32_BE_BOM, UTF32_LE_BOM, strip_bom, write_bom,
};
pub use detection::FileFormat;
pub use error::{Error, Result};
pub use metadata::Metadata;
pub use shapes::{PlaceholderType, ShapeType};
pub use style::{Length, RGBColor, VerticalPosition};
// Unit conversions
pub use unit::LengthUnit;
// Shared slice types — kept `pub` (not `pub(crate)` per spec) because the
// umbrella's docx code uses `XmlSlice` in public type signatures across the
// crate boundary. `#[doc(hidden)]` suppresses public docs surface.
#[doc(hidden)]
pub use xml_slice::{XmlArenaBuilder, XmlSlice};
