//! Layered mutable PresentationML authoring.

pub mod pres;
pub mod shape;
pub mod slide;

pub use pres::{ChartParts, FIRST_SLIDE_ID, MutablePresentation, SmartArtParts};
pub use shape::{MutableShape, ShapeRelIds, ShapeType};
pub use slide::MutableSlide;

pub mod presentation {
    pub use super::pres::*;
}
