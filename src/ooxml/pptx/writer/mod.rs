//! Mutable presentation writer components for PPTX.

pub mod pres;
pub(crate) mod relmap;
pub mod shape;
pub mod slide;

// Re-export main types
pub use pres::MutablePresentation;
pub use shape::MutableShape;
pub use slide::MutableSlide;
