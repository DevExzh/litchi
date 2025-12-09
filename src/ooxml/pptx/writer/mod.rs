//! Mutable presentation writer components for PPTX.

pub(crate) mod excel_embed;
pub mod pres;
pub(crate) mod relmap;
pub mod shape;
pub mod slide;

// Re-export main types
pub use pres::{ChartParts, MutablePresentation, SmartArtParts};
pub use shape::MutableShape;
pub use slide::MutableSlide;
