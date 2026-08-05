//! Contextual DOCX host facade for embedded DrawingML chart parts.
//!
//! [`litchi_drawingml::chart`] owns chart-schema semantics. This module owns
//! the DOCX placement, relationship graph, companion parts, and opaque bytes.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    Companion, Conformance, EmbeddedWorkbook, EmbeddedWorkbookContentType, Graph, Resource,
};
pub use package::{load, store};
