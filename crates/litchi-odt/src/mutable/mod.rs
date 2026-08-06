//! Mutable document structure for in-place modifications.
//!
//! The owner is deliberately layered:
//!
//! - [`model`] stores the in-memory structural projection;
//! - [`codec`] owns snapshot transitions and XML seams;
//! - [`semantic`] provides contextual content and styles views;
//! - [`package`] handles package input/output; and
//! - the test modules keep the public mutation contracts close to this owner.

mod codec;
mod elements;
mod metadata;
mod model;
mod package;
pub mod semantic;

#[cfg(test)]
mod semantic_tests;

#[cfg(test)]
mod tests;

pub use model::MutableDocument;
pub use semantic::{Content, ContentMut, Styles, StylesMut};
