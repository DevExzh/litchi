//! Shared DrawingML text-body vocabulary.
//!
//! SpreadsheetDrawing and PresentationML both embed the neutral
//! `a:CT_TextBody` story. Host crates retain their wrapper and relationship
//! semantics while this module owns the reusable body model.

mod model;

pub use model::{Body, Insets, Paragraph, Properties, Run};
