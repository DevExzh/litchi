/// Configuration types for Markdown conversion.
///
/// This module defines the configuration options and enums used to customize
/// the Markdown conversion process.
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

