//! Canonical WordprocessingML (`.docx`) APIs.
//!
//! The concise modules own format semantics while [`litchi_opc`] remains the
//! explicit low-level package graph.

#![forbid(unsafe_code)]

mod error;

pub mod alt;
pub mod color;
pub mod enums;
pub mod font;
pub mod format;
pub mod glossary;
pub mod statistics;
pub mod web;

pub use enums::{WdHeaderFooter, WdOrientation, WdSectionStart, WdStyleType};
pub use error::{Error, Result};
pub use format::{ImageFormat, LineSpacing, ParagraphAlignment, TableBorderStyle, UnderlineStyle};
pub use statistics::{
    DocumentStatistics, count_characters, count_characters_no_spaces, count_words,
    estimate_line_count, estimate_page_count,
};
