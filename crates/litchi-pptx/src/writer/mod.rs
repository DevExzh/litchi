//! Layered mutable `PresentationML` authoring.

pub mod presentation;
pub mod shape;
pub mod slide;

pub use presentation::{FIRST_SLIDE_ID, MutablePresentation};
pub use shape::{MutableShape, ShapeType};
pub use slide::MutableSlide;

#[cfg(test)]
mod tests;
