//! Semantic SpreadsheetDrawing shape ownership.
//!
//! Typed contextual values live in model, bounded XML parsing lives in codec,
//! package relationship traversal lives in package, and focused regression
//! coverage lives in tests.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_drawing_shapes;
pub use model::*;
pub use package::{load_shapes, load_sheet_shapes};
