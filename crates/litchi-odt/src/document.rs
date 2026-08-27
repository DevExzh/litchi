//! Ergonomic ODT document facade.
//!
//! The public [`Document`] type is intentionally small at this boundary. Its
//! typed state, XML codecs, package lifecycle, semantic queries, and safety
//! limits live in focused sibling modules.

mod catalog;
mod codec;
mod model;
mod open_parse;
mod package;
mod semantic;
mod source;
mod validation;

#[cfg(test)]
mod tests;

pub use catalog::{SourceBackedDocumentCatalog, TextBlockCatalogEntry};
pub use model::Document;
pub use source::{ReadLimits, SourceBackedDocument};
