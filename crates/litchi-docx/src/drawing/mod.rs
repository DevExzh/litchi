//! Typed inventory of WordprocessingML drawing anchors.
//!
//! The owner is intentionally small: [`model`] contains the owned semantic
//! inventory while [`codec`] contains the streaming `<w:drawing>` scanner.
//! DrawingML preset tokens remain the closed domain supplied by
//! [`litchi_drawingml::geom::Preset`]; unknown drawing children are inert.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Anchor, Kind, Object};

pub(crate) use codec::parse;
