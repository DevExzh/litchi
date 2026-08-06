//! Ergonomic ODT document facade.
//!
//! The public [`Document`] type is intentionally small at this boundary. Its
//! typed state, XML codecs, package lifecycle, semantic queries, and safety
//! limits live in focused sibling modules.

mod codec;
mod model;
mod package;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use model::Document;
