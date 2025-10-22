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
///
/// # Note
/// This struct is `Copy` since it only contains simple values,
/// making it cheap to pass around without cloning overhead.
#[derive(Debug, Copy, Clone)]
pub struct MarkdownOptions {
    /// Whether to include text styles (bold, italic, underline, etc.)
    pub include_styles: bool,
    /// Whether to include document metadata at the beginning
    pub include_metadata: bool,
    /// How to render tables
    pub table_style: TableStyle,
    /// Indentation for HTML tables (spaces)
    pub html_table_indent: usize,
    /// How to render mathematical formulas
    pub formula_style: FormulaStyle,
    /// Number of spaces for list indentation (default: 2)
    pub list_indent: usize,
    /// How to render superscript and subscript
    pub script_style: ScriptStyle,
    /// How to render strikethrough text
    pub strikethrough_style: StrikethroughStyle,
}

impl Default for MarkdownOptions {
    fn default() -> Self {
        Self {
            include_styles: true,
            include_metadata: false,
            table_style: TableStyle::Markdown,
            html_table_indent: 2,
            formula_style: FormulaStyle::LaTeX,
            list_indent: 2,
            script_style: ScriptStyle::Html,
            strikethrough_style: StrikethroughStyle::Markdown,
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

    /// Set the formula rendering style.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::{MarkdownOptions, FormulaStyle};
    ///
    /// let options = MarkdownOptions::new()
    ///     .with_formula_style(FormulaStyle::Dollar);
    /// ```
    #[inline]
    pub fn with_formula_style(mut self, style: FormulaStyle) -> Self {
        self.formula_style = style;
        self
    }

    /// Set the list indentation (number of spaces).
    ///
    /// Used for nested lists to indicate hierarchy.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::MarkdownOptions;
    ///
    /// let options = MarkdownOptions::new().with_list_indent(4);
    /// ```
    #[inline]
    pub fn with_list_indent(mut self, indent: usize) -> Self {
        self.list_indent = indent;
        self
    }

    /// Set the superscript and subscript rendering style.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::{MarkdownOptions, ScriptStyle};
    ///
    /// let options = MarkdownOptions::new()
    ///     .with_script_style(ScriptStyle::Unicode);
    /// ```
    #[inline]
    pub fn with_script_style(mut self, style: ScriptStyle) -> Self {
        self.script_style = style;
        self
    }

    /// Set the strikethrough rendering style.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::markdown::{MarkdownOptions, StrikethroughStyle};
    ///
    /// let options = MarkdownOptions::new()
    ///     .with_strikethrough_style(StrikethroughStyle::Html);
    /// ```
    #[inline]
    pub fn with_strikethrough_style(mut self, style: StrikethroughStyle) -> Self {
        self.strikethrough_style = style;
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

/// Formula rendering styles for Markdown conversion.
///
/// Determines how mathematical formulas are rendered in the output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaStyle {
    /// Use LaTeX delimiters: inline \(\) and display \[\].
    ///
    /// Examples:
    /// - Inline: \(\sin\pi\)
    /// - Display: \[\sin\pi\]
    LaTeX,

    /// Use dollar signs: inline $ and display $$ (GitHub flavored).
    ///
    /// Examples:
    /// - Inline: $\sin\pi$
    /// - Display: $$\sin\pi$$
    Dollar,
}

/// Superscript and subscript rendering styles.
///
/// Determines how superscript and subscript text is rendered.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScriptStyle {
    /// Use HTML tags: <sup> and <sub>.
    ///
    /// Examples:
    /// - Superscript: x<sup>2</sup>
    /// - Subscript: H<sub>2</sub>O
    Html,

    /// Use Unicode superscript/subscript characters where possible.
    ///
    /// Falls back to HTML tags for unsupported characters.
    ///
    /// Examples:
    /// - Superscript: x²
    /// - Subscript: H₂O
    Unicode,
}

/// Strikethrough rendering styles.
///
/// Determines how strikethrough text is rendered.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StrikethroughStyle {
    /// Use Markdown strikethrough: ~~text~~
    ///
    /// Example: ~~deleted text~~
    Markdown,

    /// Use HTML tags: <del>text</del>
    ///
    /// Example: <del>deleted text</del>
    Html,
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
            .with_html_table_indent(4)
            .with_formula_style(FormulaStyle::Dollar)
            .with_list_indent(4)
            .with_script_style(ScriptStyle::Unicode)
            .with_strikethrough_style(StrikethroughStyle::Html);

        assert!(options.include_styles);
        assert!(!options.include_metadata);
        assert_eq!(options.table_style, TableStyle::MinimalHtml);
        assert_eq!(options.html_table_indent, 4);
        assert_eq!(options.formula_style, FormulaStyle::Dollar);
        assert_eq!(options.list_indent, 4);
        assert_eq!(options.script_style, ScriptStyle::Unicode);
        assert_eq!(options.strikethrough_style, StrikethroughStyle::Html);
    }

    #[test]
    fn test_markdown_options_default() {
        let options = MarkdownOptions::default();
        assert!(options.include_styles);
        assert!(!options.include_metadata);
        assert_eq!(options.table_style, TableStyle::Markdown);
        assert_eq!(options.html_table_indent, 2);
        assert_eq!(options.formula_style, FormulaStyle::LaTeX);
        assert_eq!(options.list_indent, 2);
        assert_eq!(options.script_style, ScriptStyle::Html);
        assert_eq!(options.strikethrough_style, StrikethroughStyle::Markdown);
    }
}
