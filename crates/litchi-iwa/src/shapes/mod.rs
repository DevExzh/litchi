//! Shape and Drawing Element Support
//!
//! This module provides support for extracting text and metadata from shapes,
//! text boxes, and other drawing elements in iWork documents.
//!
//! Shapes in iWork documents can contain text (text boxes), images, or be
//! purely visual elements. This module helps extract meaningful content
//! from these elements.

mod geometry;
mod properties;
pub mod text_extractor;

pub use geometry::{DrawableGeometry, DrawablePoint, DrawableSize};
pub(crate) use geometry::{set_shape_geometry, shape_geometry};
pub use properties::DrawableProperties;
pub(crate) use properties::{set_shape_properties, shape_properties};
pub use text_extractor::ShapeTextExtractor;
