//! Mutable DOCX document authoring.
//!
//! The public facade stays intentionally small. Semantic state and editing
//! operations live in the model module, XML/settings codecs in the codec
//! module, preserved body/package integration in the package module, and
//! invariants in the tests module.

#[path = "doc/codec.rs"]
mod codec;
#[path = "doc/model.rs"]
mod model;
#[path = "doc/package.rs"]
mod package;
#[cfg(test)]
#[path = "doc/tests.rs"]
mod tests;

pub use super::super::format::ImageFormat;
pub use model::{MutableDocument, Protection};
pub(crate) use package::{BodyElement, DocumentBody};
