//! Layered mutable PresentationML authoring.

pub(crate) mod relmap;
pub mod presentation;
pub mod shape;
pub mod slide;

pub use presentation::{ChartParts, FIRST_SLIDE_ID, MutablePresentation, SmartArtParts};
pub use shape::{MutableShape, ShapeRelIds, ShapeType};
pub use slide::MutableSlide;
