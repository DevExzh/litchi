//! Semantic `SpreadsheetDrawing` shape facade.
//!
//! The shared owner provides the bounded XML grammar and semantic model;
//! this host retains only workbook/package relationship traversal.

mod package;

pub use litchi_spreadsheet_drawing::shape::*;
pub use package::{load_shapes, load_sheet_shapes};
