//! Borrowed semantic view of one PresentationML package graph.
//!
//! The typed facade delegates XML scanning to [`codec`] and package
//! relationship traversal to [`package`]. Embedded resources and media
//! authoring remain available through their existing child modules.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub mod embedded;
pub mod media;

pub use model::Presentation;
