//! Typed table capabilities.

/// Lexical table-column edits (record 0321).
pub(crate) mod columns;
/// Borrowed tables embedded in slide graphic frames.
pub mod shape;

pub use shape::{Cell, CellProperties, CellTextIndex, Properties, Row, Table};

pub mod style;
