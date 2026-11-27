//! Typed table capabilities.

/// Lexical table-column edits.
pub(crate) mod columns;
/// Borrowed tables embedded in slide graphic frames.
pub mod shape;

pub use shape::{Cell, CellProperties, CellTextIndex, Properties, Row, Table};

pub mod style;
