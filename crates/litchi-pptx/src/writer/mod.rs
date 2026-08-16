//! Layered mutable `PresentationML` authoring.

pub mod presentation;
pub mod shape;
pub mod slide;
pub mod streaming;

pub use presentation::{FIRST_SLIDE_ID, MutablePresentation};
pub use shape::{MutableShape, ShapeType};
pub use slide::MutableSlide;
pub use streaming::{
    StreamingPresentationLimits, StreamingPresentationOptions, StreamingPresentationWriter,
    StreamingSlideWriter, TextBoxSpec,
};

#[cfg(test)]
mod tests;
