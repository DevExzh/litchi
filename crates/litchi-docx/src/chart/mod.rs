//! Contextual DOCX host facade for embedded `DrawingML` chart parts.
//!
//! [`litchi_drawingml::chart`] owns chart-schema semantics. This module owns
//! the DOCX placement, relationship graph, companion parts, and opaque bytes.

mod codec;
mod edit;
pub(crate) mod inline;
mod limit;
mod model;
mod package;
#[allow(
    dead_code,
    unreachable_pub,
    reason = "the OPC-level draft stages the family edit; its patch accessors serve unit tests"
)]
mod transaction;

#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use inline::InlineChart;
pub use model::{
    Companion, Conformance, EmbeddedWorkbook, EmbeddedWorkbookContentType, Graph, Resource,
};
pub use package::{load, store};
pub use transaction::{Revision, Snapshot};
