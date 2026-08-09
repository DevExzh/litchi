//! `ODF` office annotation vocabulary.
//!
//! The facade keeps the semantic annotation model separate from its bounded
//! `XML` event codec and package resource policy.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{Builder, write_attributes};
pub use model::{Annotation, Element, Node};
