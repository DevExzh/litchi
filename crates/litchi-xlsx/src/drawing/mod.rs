//! Layered SpreadsheetDrawing ownership.
//!
//! [`model`] owns the contextual object inventory and [`codec`] owns the
//! namespace-aware, bounded XML reader. Shape authoring and full DrawingML
//! text parsing remain in their existing `shapes` owner; the text facade here
//! reuses [`litchi_drawingml`] without copying its vocabulary.

mod codec;
mod model;

pub use super::chart::Anchor;
pub use codec::parse;
pub use model::{Chart, Drawing, Object, Picture, Unknown, UnknownKind, text};

#[cfg(test)]
mod tests;
