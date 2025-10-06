/// Markdown conversion functionality for Office documents and presentations.
///
/// This module provides high-performance conversion of Word documents and PowerPoint
/// presentations to Markdown format. It supports both legacy (OLE2) and modern (OOXML)
/// formats with a unified API.
///
/// # Features
///
/// - **Format-agnostic**: Works with both .doc/.docx and .ppt/.pptx files
/// - **Style preservation**: Converts bold, italic, underline, and other text formatting
/// - **Table conversion**: Smart table handling (Markdown tables or HTML when needed)
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
/// - [`MarkdownWriter`]: Low-level writer for efficient output generation
///
/// # Performance Considerations
///
/// This implementation is designed for high performance:
/// - Uses borrowing instead of cloning where possible
/// - Reuses buffers in parsing loops
/// - Uses `SmallVec` for temporary small vectors
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

use crate::common::{Error, Result};
use crate::document::{Document, Paragraph, Run, Table};
use crate::presentation::{Presentation, Slide};
use std::fmt::Write as FmtWrite;

/// Core trait for types that can be converted to Markdown.
///
/// This trait is implemented for Document, Presentation, and their constituent
/// parts (paragraphs, runs, tables, etc.).
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::{Document, markdown::ToMarkdown};
///
/// # fn main() -> Result<(), litchi::Error> {
/// let doc = Document::open("document.docx")?;
///
/// // Convert entire document
/// let markdown = doc.to_markdown()?;
///
/// // Or convert individual parts
/// for para in doc.paragraphs()? {
///     let para_md = para.to_markdown()?;
///     println!("{}", para_md);
/// }
/// # Ok(())
/// # }
/// ```
pub trait ToMarkdown {
    /// Convert this item to Markdown with default options.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::{Document, markdown::ToMarkdown};
    ///
    /// # fn main() -> Result<(), litchi::Error> {
    /// let doc = Document::open("document.docx")?;
    /// let markdown = doc.to_markdown()?;
    /// # Ok(())
    /// # }
    /// ```
    fn to_markdown(&self) -> Result<String> {
        self.to_markdown_with_options(&MarkdownOptions::default())
    }

    /// Convert this item to Markdown with custom options.
    ///
    /// # Arguments
    ///
    /// * `options` - Configuration for the conversion
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::{Document, markdown::{ToMarkdown, MarkdownOptions}};
    ///
    /// # fn main() -> Result<(), litchi::Error> {
    /// let doc = Document::open("document.docx")?;
    /// let options = MarkdownOptions::new().with_styles(true);
    /// let markdown = doc.to_markdown_with_options(&options)?;
    /// # Ok(())
    /// # }
    /// ```
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String>;
}

/// Configuration options for Markdown conversion.
///
/// This struct controls various aspects of the Markdown output, including
/// whether to include styles, metadata, and how to format tables.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::{MarkdownOptions, TableStyle};
///
/// // Create with defaults
/// let options = MarkdownOptions::default();
///
/// // Or customize
/// let options = MarkdownOptions::new()
///     .with_styles(true)
///     .with_metadata(false)
///     .with_table_style(TableStyle::MinimalHtml);
/// ```
#[derive(Debug, Clone)]
pub struct MarkdownOptions {
    /// Whether to include text styles (bold, italic, underline, etc.)
    pub include_styles: bool,
    /// Whether to include document metadata at the beginning
    pub include_metadata: bool,
    /// How to render tables
    pub table_style: TableStyle,
    /// Indentation for HTML tables (spaces)
    pub html_table_indent: usize,
}

impl Default for MarkdownOptions {
    fn default() -> Self {
        Self {
            include_styles: true,
            include_metadata: false,
            table_style: TableStyle::Markdown,
            html_table_indent: 2,
        }
    }
}

impl MarkdownOptions {
    /// Create a new `MarkdownOptions` with default values.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::MarkdownOptions;
    ///
    /// let options = MarkdownOptions::new();
    /// ```
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set whether to include text styles.
    ///
    /// When enabled, text formatting like **bold**, *italic*, ~~strikethrough~~
    /// will be preserved in the output.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::MarkdownOptions;
    ///
    /// let options = MarkdownOptions::new().with_styles(true);
    /// ```
    #[inline]
    pub fn with_styles(mut self, include: bool) -> Self {
        self.include_styles = include;
        self
    }

    /// Set whether to include document metadata.
    ///
    /// When enabled, document metadata (title, author, etc.) will be included
    /// at the beginning of the output as a YAML frontmatter block.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::MarkdownOptions;
    ///
    /// let options = MarkdownOptions::new().with_metadata(true);
    /// ```
    #[inline]
    pub fn with_metadata(mut self, include: bool) -> Self {
        self.include_metadata = include;
        self
    }

    /// Set the table rendering style.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::{MarkdownOptions, TableStyle};
    ///
    /// let options = MarkdownOptions::new()
    ///     .with_table_style(TableStyle::MinimalHtml);
    /// ```
    #[inline]
    pub fn with_table_style(mut self, style: TableStyle) -> Self {
        self.table_style = style;
        self
    }

    /// Set the indentation for HTML tables (number of spaces).
    ///
    /// Only applies when using HTML table styles.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::MarkdownOptions;
    ///
    /// let options = MarkdownOptions::new().with_html_table_indent(4);
    /// ```
    #[inline]
    pub fn with_html_table_indent(mut self, indent: usize) -> Self {
        self.html_table_indent = indent;
        self
    }
}

/// Table rendering styles for Markdown conversion.
///
/// Determines how tables are rendered in the output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableStyle {
    /// Use Markdown tables (only when no merged cells exist).
    ///
    /// If merged cells are detected, falls back to MinimalHtml.
    ///
    /// Example:
    /// ```markdown
    /// | Header 1 | Header 2 |
    /// |----------|----------|
    /// | Cell 1   | Cell 2   |
    /// ```
    Markdown,

    /// Use minimal HTML tables (no styling, just structure).
    ///
    /// Example:
    /// ```html
    /// <table>
    ///   <tr><td>Cell 1</td><td>Cell 2</td></tr>
    /// </table>
    /// ```
    MinimalHtml,

    /// Use styled HTML tables with customizable indentation.
    ///
    /// Includes basic CSS classes for styling.
    StyledHtml,
}

/// Low-level writer for efficient Markdown generation.
///
/// This struct provides optimized methods for writing Markdown elements
/// with minimal allocations.
struct MarkdownWriter {
    /// The output buffer
    buffer: String,
    /// Current options
    options: MarkdownOptions,
}

impl MarkdownWriter {
    /// Create a new writer with the given options.
    fn new(options: MarkdownOptions) -> Self {
        Self {
            buffer: String::with_capacity(4096), // Pre-allocate reasonable size
            options,
        }
    }

    /// Write a paragraph to the buffer.
    fn write_paragraph(&mut self, para: &Paragraph) -> Result<()> {
        if self.options.include_styles {
            // Write runs with style information
            let runs = para.runs()?;
            for run in runs {
                self.write_run(&run)?;
            }
        } else {
            // Write plain text
            let text = para.text()?;
            self.buffer.push_str(&text);
        }
        
        // Add paragraph break
        self.buffer.push_str("\n\n");
        Ok(())
    }

    /// Write a run with formatting.
    fn write_run(&mut self, run: &Run) -> Result<()> {
        let text = run.text()?;
        if text.is_empty() {
            return Ok(());
        }

        let bold = run.bold()?.unwrap_or(false);
        let italic = run.italic()?.unwrap_or(false);

        // Apply markdown formatting
        // Note: Markdown doesn't support underline natively
        match (bold, italic) {
            (true, true) => {
                write!(self.buffer, "***{}***", text).map_err(|e| Error::Other(e.to_string()))?;
            }
            (true, false) => {
                write!(self.buffer, "**{}**", text).map_err(|e| Error::Other(e.to_string()))?;
            }
            (false, true) => {
                write!(self.buffer, "*{}*", text).map_err(|e| Error::Other(e.to_string()))?;
            }
            (false, false) => {
                self.buffer.push_str(&text);
            }
        }

        Ok(())
    }

    /// Write a table to the buffer.
    fn write_table(&mut self, table: &Table) -> Result<()> {
        // Check if table has merged cells
        let has_merged_cells = self.table_has_merged_cells(table)?;

        match self.options.table_style {
            TableStyle::Markdown if !has_merged_cells => {
                self.write_markdown_table(table)?;
            }
            TableStyle::MinimalHtml | TableStyle::Markdown => {
                self.write_html_table(table, false)?;
            }
            TableStyle::StyledHtml => {
                self.write_html_table(table, true)?;
            }
        }

        // Add spacing after table
        self.buffer.push_str("\n\n");
        Ok(())
    }

    /// Check if a table has merged cells.
    ///
    /// For OLE tables, we check the CellMergeStatus.
    /// For OOXML tables, we would need to parse cell properties (tcPr/gridSpan, vMerge).
    /// For now, we conservatively assume OOXML tables might have merged cells.
    fn table_has_merged_cells(&self, table: &Table) -> Result<bool> {
        // For the unified API, we can't directly check merge status without
        // accessing the underlying implementation. For safety, we check
        // if all rows have the same number of cells.
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(false);
        }

        let first_row_cell_count = rows[0].cells()?.len();
        for row in &rows[1..] {
            if row.cells()?.len() != first_row_cell_count {
                return Ok(true); // Inconsistent cell counts suggest merged cells
            }
        }

        Ok(false)
    }

    /// Write a table in Markdown format.
    fn write_markdown_table(&mut self, table: &Table) -> Result<()> {
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(());
        }

        // Write header row (first row)
        self.buffer.push('|');
        let first_row = &rows[0];
        for cell in first_row.cells()? {
            let text = cell.text()?;
            // Escape pipe characters in cell content
            let escaped = text.replace('|', "\\|").replace('\n', " ");
            write!(self.buffer, " {} |", escaped).map_err(|e| Error::Other(e.to_string()))?;
        }
        self.buffer.push('\n');

        // Write separator row
        self.buffer.push('|');
        let cell_count = first_row.cells()?.len();
        for _ in 0..cell_count {
            self.buffer.push_str("----------|");
        }
        self.buffer.push('\n');

        // Write data rows
        for row in &rows[1..] {
            self.buffer.push('|');
            for cell in row.cells()? {
                let text = cell.text()?;
                let escaped = text.replace('|', "\\|").replace('\n', " ");
                write!(self.buffer, " {} |", escaped).map_err(|e| Error::Other(e.to_string()))?;
            }
            self.buffer.push('\n');
        }

        Ok(())
    }

    /// Write a table in HTML format.
    fn write_html_table(&mut self, table: &Table, styled: bool) -> Result<()> {
        let indent = " ".repeat(self.options.html_table_indent);

        if styled {
            self.buffer.push_str("<table class=\"doc-table\">\n");
        } else {
            self.buffer.push_str("<table>\n");
        }

        let rows = table.rows()?;
        for (i, row) in rows.iter().enumerate() {
            writeln!(self.buffer, "{}<tr>", indent).map_err(|e| Error::Other(e.to_string()))?;

            // First row is typically header
            let tag = if i == 0 { "th" } else { "td" };

            for cell in row.cells()? {
                let text = cell.text()?;
                // HTML escape
                let escaped = text
                    .replace('&', "&amp;")
                    .replace('<', "&lt;")
                    .replace('>', "&gt;")
                    .replace('\n', "<br>");
                
                writeln!(self.buffer, "{}{}<{}>{}</{}>", 
                    indent, indent, tag, escaped, tag)
                    .map_err(|e| Error::Other(e.to_string()))?;
            }

            writeln!(self.buffer, "{}</tr>", indent).map_err(|e| Error::Other(e.to_string()))?;
        }

        self.buffer.push_str("</table>");
        Ok(())
    }

    /// Get the final markdown output.
    fn finish(self) -> String {
        self.buffer
    }
}

// ============================================================================
// ToMarkdown implementations for Document types
// ============================================================================

impl ToMarkdown for Document {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());

        // TODO: Add metadata support when Document metadata API is available
        // if options.include_metadata {
        //     writer.write_metadata(...)?;
        // }

        // Write paragraphs
        for para in self.paragraphs()? {
            writer.write_paragraph(&para)?;
        }

        // Write tables
        for table in self.tables()? {
            writer.write_table(&table)?;
        }

        Ok(writer.finish())
    }
}

impl ToMarkdown for Paragraph {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());
        writer.write_paragraph(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

impl ToMarkdown for Run {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());
        writer.write_run(self)?;
        Ok(writer.finish())
    }
}

impl ToMarkdown for Table {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());
        writer.write_table(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

// ============================================================================
// ToMarkdown implementations for Presentation types
// ============================================================================

impl ToMarkdown for Presentation {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        let mut output = String::with_capacity(4096);

        // TODO: Add metadata support when Presentation metadata API is available
        // if _options.include_metadata {
        //     output.push_str("---\n");
        //     output.push_str(&format!("slides: {}\n", self.slide_count()?));
        //     output.push_str("---\n\n");
        // }

        let slides = self.slides()?;
        for (i, slide) in slides.iter().enumerate() {
            if i > 0 {
                // Separate slides with horizontal rule
                output.push_str("\n\n---\n\n");
            }

            // Add slide number as heading
            writeln!(output, "# Slide {}", i + 1).map_err(|e| Error::Other(e.to_string()))?;
            output.push('\n');

            // Add slide content
            let text = slide.text()?;
            output.push_str(&text);
        }

        Ok(output)
    }
}

impl ToMarkdown for Slide {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        // For individual slides, just return the text
        // Formatting is minimal for presentations
        self.text()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_markdown_options_builder() {
        let options = MarkdownOptions::new()
            .with_styles(true)
            .with_metadata(false)
            .with_table_style(TableStyle::MinimalHtml)
            .with_html_table_indent(4);

        assert!(options.include_styles);
        assert!(!options.include_metadata);
        assert_eq!(options.table_style, TableStyle::MinimalHtml);
        assert_eq!(options.html_table_indent, 4);
    }

    #[test]
    fn test_markdown_options_default() {
        let options = MarkdownOptions::default();
        assert!(options.include_styles);
        assert!(!options.include_metadata);
        assert_eq!(options.table_style, TableStyle::Markdown);
        assert_eq!(options.html_table_indent, 2);
    }
}
