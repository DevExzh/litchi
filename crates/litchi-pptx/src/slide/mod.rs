//! Contextual slide, layout, and master facades.
//!
//! The owner is deliberately split by responsibility. [`model`] contains
//! the borrowed semantic views and the checked [`Key`] selector. [`package`]
//! resolves package relationships and optional companion parts. [`codec`]
//! adapts those views to the bounded `PresentationML` and `DrawingML` readers
//! owned by the validated [`crate::parts::SlidePart`] family. This keeps
//! package graph traversal and semantic accessors separate without copying or
//! weakening the existing strict/transitional readers.

pub mod codec;
pub mod model;
pub mod package;

#[cfg(test)]
mod tests;

pub use model::{Key, Slide, SlideLayout, SlideMaster};
