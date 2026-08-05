//! Typed worksheet-view values for SpreadsheetML.
//!
//! The facade keeps the concise `views` owner path while separating semantic
//! values from their focused regression tests.

mod model;

#[cfg(test)]
mod tests;

pub use model::{Pane, PanePosition, PaneState, Selection, View, ViewType};
