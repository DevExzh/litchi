//! Semantic presentation owner for the binary `PowerPoint` document stream.
//!
//! The facade keeps the established `crate::presentation` paths stable while
//! separating the typed presentation model, package loading, OfficeArt-derived
//! projections, and focused owner tests.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::{ParsedCustomShow, ParsedSlideComments, Presentation};
