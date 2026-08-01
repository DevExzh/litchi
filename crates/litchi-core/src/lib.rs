// Many items in this crate were previously private inside the umbrella
// `litchi` crate and never had public doc coverage. Tighten this once a
// docs pass lands.
#![allow(missing_docs)]
//! Common types, traits, and utilities shared across formats.
//!
//! This module provides unified types and traits used by both OLE2 (legacy)
//! and OOXML (modern) implementations, ensuring a consistent API for users.

// Submodule declarations
pub mod binary;
pub mod bom;
pub mod bounded;
pub mod budget;
pub mod detection;
pub mod error;
pub mod hex;
pub mod metadata;
pub mod selector;
pub mod shapes;
pub mod sheet;
pub mod simd;
pub mod source;
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
pub use bounded::{BoundedU32, BoundsError};
pub use budget::{Budget, Limits, Profile, Reservation, Resource, ResourceLimit};
pub use detection::FileFormat;
pub use error::{Error, Result};
pub use metadata::Metadata;
pub use selector::{Position, Selector};
pub use shapes::{PlaceholderType, ShapeType};
pub use source::{OwnedSource, ReadAt, SliceSource, SourceVersion};
pub use style::{Length, RGBColor, VerticalPosition};
// Unit conversions
pub use unit::LengthUnit;
// Shared slice types — kept `pub` (not `pub(crate)` per spec) because the
// umbrella's docx code uses `XmlSlice` in public type signatures across the
// crate boundary. `#[doc(hidden)]` suppresses public docs surface.
#[doc(hidden)]
pub use xml_slice::{XmlArenaBuilder, XmlSlice};
