//! Convert documents and presentations to Markdown.
//!
//! [`ToMarkdown`](crate::markdown::ToMarkdown) is implemented for
//! [`File`](crate::File), its document and presentation views, and the
//! element wrappers those views return. A document converts paragraph by
//! paragraph and table by table, with direct formatting, lists, headings,
//! DOCX hyperlinks, images and notes; a presentation converts to one
//! `# Slide N` section per slide. Content that Markdown cannot carry
//! faithfully is refused with a typed [`Error`](crate::Error) rather than
//! dropped, and a workbook is refused as an unsupported capability.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! use litchi::markdown::{MarkdownOptions, TableStyle, ToMarkdown};
//!
//! let path = concat!(
//!     env!("CARGO_MANIFEST_DIR"),
//!     "/../../test-data/ooxml/docx/documentProperties.docx"
//! );
//! let file = litchi::open(path)?;
//! assert_eq!(file.to_markdown()?, "Hello World\\!\n\n");
//!
//! let options = MarkdownOptions::new()
//!     .with_styles(false)
//!     .with_table_style(TableStyle::MinimalHtml);
//! assert_eq!(file.to_markdown_with_options(&options)?, "Hello World\\!\n\n");
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! The conversion lives in `litchi-ops`, next to [`File`](crate::File),
//! because the trait and `File` are both foreign to this crate.

// Re-export public API (format-agnostic helpers live in litchi-markdown)
pub use litchi_markdown::{
    FormulaStyle, MarkdownOptions, ScriptStyle, StrikethroughStyle, TableStyle, ToMarkdown, unicode,
};
