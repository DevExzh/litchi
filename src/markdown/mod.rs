/// Markdown conversion functionality for Office documents and presentations.
///
/// This module provides high-performance conversion of Word documents and PowerPoint
/// presentations to Markdown format. It supports both legacy (OLE2) and modern (OOXML)
/// formats with a unified API.
///
/// # Features
///
/// - **Format-agnostic**: Works with both .doc/.docx and .ppt/.pptx files
/// - **Style preservation**: Converts bold, italic, and other text formatting
/// - **Table conversion**: Smart handling of tables (Markdown or HTML)
/// - **High performance**: Memory-efficient with minimal allocations
/// - **Configurable**: Extensive options for customizing output
///
/// # Quick Start
///
/// ```rust,no_run
/// use litchi::{Document, markdown::ToMarkdown};
///
/// # fn main() -> Result<(), litchi::Error> {
/// // Convert a document to markdown
/// let doc = Document::open("report.docx")?;
/// let markdown = doc.to_markdown()?;
/// println!("{}", markdown);
///
/// // Or with custom options
/// use litchi::markdown::MarkdownOptions;
/// let options = MarkdownOptions::new()
///     .with_styles(true)
///     .with_metadata(false)
///     .with_html_tables(false);
/// let markdown = doc.to_markdown_with_options(&options)?;
/// # Ok(())
/// # }
/// ```
///
/// # Architecture
///
/// The module is organized around:
/// - [`ToMarkdown`] trait: Core trait for types that can be converted to Markdown
/// - [`MarkdownOptions`]: Configuration for conversion behavior
/// - [`config`]: Configuration types and enums
/// - [`writer`]: Low-level writer for efficient output generation
/// - [`document`]: Document-specific implementations
/// - [`presentation`]: Presentation-specific implementations
///
/// # Performance Considerations
///
/// This implementation is designed for high performance:
/// - Uses borrowing instead of cloning where possible
/// - Reuses buffers in parsing loops
/// - Uses pre-allocated buffers with appropriate capacity
/// - No unsafe code
///
/// # Examples
///
/// ## Basic Document Conversion
///
/// ```rust,no_run
/// use litchi::{Document, markdown::ToMarkdown};
///
/// # fn main() -> Result<(), litchi::Error> {
/// let doc = Document::open("document.docx")?;
/// let markdown = doc.to_markdown()?;
/// println!("{}", markdown);
/// # Ok(())
/// # }
/// ```
///
/// ## With Custom Options
///
/// ```rust,no_run
/// use litchi::{Document, markdown::{ToMarkdown, MarkdownOptions, TableStyle}};
///
/// # fn main() -> Result<(), litchi::Error> {
/// let doc = Document::open("document.docx")?;
///
/// let options = MarkdownOptions::new()
///     .with_styles(true)           // Include bold, italic, etc.
///     .with_metadata(true)          // Include document metadata
///     .with_table_style(TableStyle::Markdown); // Use markdown tables
///
/// let markdown = doc.to_markdown_with_options(&options)?;
/// # Ok(())
/// # }
/// ```
///
/// ## Presentation Conversion
///
/// ```rust,no_run
/// use litchi::{Presentation, markdown::ToMarkdown};
///
/// # fn main() -> Result<(), litchi::Error> {
/// let pres = Presentation::open("slides.pptx")?;
/// let markdown = pres.to_markdown()?;
/// // Slides are separated by horizontal rules (---)
/// println!("{}", markdown);
/// # Ok(())
/// # }
/// ```
// Module declarations
mod config;
mod traits;
mod writer;
pub mod unicode;

// Document and presentation markdown implementations are only available when their respective features are enabled
#[cfg(any(feature = "ole", feature = "ooxml"))]
mod document;

#[cfg(any(feature = "ole", feature = "ooxml"))]
mod presentation;

// Re-export public API
pub use config::{MarkdownOptions, TableStyle, FormulaStyle, ScriptStyle, StrikethroughStyle};
pub use traits::ToMarkdown;
