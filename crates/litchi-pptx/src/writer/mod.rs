//! Layered mutable PresentationML authoring.

pub mod presentation;
pub(crate) mod relmap;
pub mod shape;
pub mod slide;

pub use presentation::{ChartParts, FIRST_SLIDE_ID, MutablePresentation, SmartArtParts};
pub use shape::{MutableShape, ShapeRelIds, ShapeType};
pub use slide::MutableSlide;
