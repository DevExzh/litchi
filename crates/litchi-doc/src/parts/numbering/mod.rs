//! Typed, layered DOC list-numbering owner.
//!
//! The facade keeps the public list model concise while the binary table
//! codecs, wire-value validation, and regression tests remain contextualized
//! below this module.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::*;
