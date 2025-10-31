//! RTF (Rich Text Format) parser module.
//!
//! This module provides high-performance parsing of RTF documents with support
//! for the RTF 1.9.1 specification. It uses arena allocation (bumpalo) for efficient
//! memory management during parsing and zero-copy patterns where possible.
//!
//! # Architecture
//!
//! The parser is organized into several components:
//! - **Lexer**: Tokenizes RTF input into control words, symbols, and text
//! - **Parser**: Builds a structured document from tokens
//! - **Document**: High-level document representation with paragraphs, runs, and tables
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::rtf::RtfDocument;
//!
//! let rtf_text = r#"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard Hello World!\par}"#;
//! let doc = RtfDocument::parse(rtf_text)?;
//! let text = doc.text();
//! # Ok::<(), litchi::common::Error>(())
//! ```

mod compressed;
mod document;
mod error;
mod field;
mod lexer;
mod parser;
mod picture;
mod table;
mod types;

// Re-exports
pub use compressed::{compress, decompress, is_compressed_rtf};
pub use document::RtfDocument;
pub use error::{RtfError, RtfResult};
pub use field::{Field, FieldType};
pub use lexer::CharacterSet;
pub use picture::{ImageType, Picture, detect_image_type};
pub use table::{Cell, Row, Table};
pub use types::{
    Alignment, Color, ColorTable, DocumentElement, Font, FontFamily, FontRef, FontTable,
    Formatting, Indentation, Paragraph, ParagraphContent, Run, Spacing, StyleBlock,
};
