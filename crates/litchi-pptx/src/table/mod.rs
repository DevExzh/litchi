//! Typed table capabilities.

/// Borrowed tables embedded in slide graphic frames.
pub mod shape;

pub use shape::{Cell, Properties, Row, Table};

pub mod style;
