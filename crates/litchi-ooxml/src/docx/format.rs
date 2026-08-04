//! Compatibility re-exports for DOCX semantic formatting types.
//!
//! The canonical model is owned by [`litchi_docx::format`]. This path remains
//! available so existing `litchi_ooxml::docx` users and the package adapter
//! continue to resolve the same type names during the ownership migration.

pub use litchi_docx::format::{
    ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle,
};
