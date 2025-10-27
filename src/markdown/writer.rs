use super::config::{MarkdownOptions, TableStyle};
/// Low-level writer for Markdown generation.
///
/// This module provides the `MarkdownWriter` struct which handles the actual
/// conversion of document elements to Markdown format.
///
/// **Note**: Some functionality requires the `ole` or `ooxml` feature to be enabled.
use crate::common::{Error, Metadata, Result};
#[cfg(any(feature = "ole", feature = "ooxml"))]
use crate::document::{Paragraph, Run, Table};
use std::fmt::Write as FmtWrite;

#[cfg(any(feature = "ole", feature = "ooxml"))]
use memchr::memchr;

/// Information about a detected list item.
#[derive(Debug, Clone)]
struct ListItemInfo {
    /// The type of list
    list_type: ListType,
    /// The nesting level (0 = top level)
    level: usize,
    /// The marker text (e.g., "1.", "-", "*")
    marker: String,
    /// The content after the marker
    content: String,
}

/// Types of lists supported.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ListType {
    /// Ordered list (numbered)
    Ordered,
    /// Unordered list (bulleted)
    Unordered,
}

/// Low-level writer for efficient Markdown generation.
///
/// This struct provides optimized methods for writing Markdown elements
/// with minimal allocations.
pub(crate) struct MarkdownWriter {
    /// The output buffer
    buffer: String,
    /// Current options
    options: MarkdownOptions,
}

impl MarkdownWriter {
    /// Create a new writer with the given options.
    pub fn new(options: MarkdownOptions) -> Self {
        Self {
            buffer: String::with_capacity(4096), // Pre-allocate reasonable size
            options,
        }
    }

    /// Write a paragraph to the buffer.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    ///
    /// **Performance**: Optimized to avoid redundant XML parsing by extracting runs
    /// once and deriving text from them when needed.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    pub fn write_paragraph(&mut self, para: &Paragraph) -> Result<()> {
        // First check for paragraph-level formulas (display math)
        #[cfg(feature = "ooxml")]
        {
            use crate::document::Paragraph;
            if let Paragraph::Docx(docx_para) = para {
                let display_formulas = docx_para.paragraph_level_formulas()?;
                if !display_formulas.is_empty() {
                    // This paragraph contains display formulas
                    // Process runs and formulas together in order
                    self.write_paragraph_with_display_formulas(para, display_formulas)?;
                    self.buffer.push_str("\n\n");
                    return Ok(());
                }
            }
        }

        // PERFORMANCE OPTIMIZATION:
        // For styled output (which needs runs anyway), get runs first and derive text from them.
        // This avoids parsing the paragraph XML twice (once for text(), once for runs()).
        // For plain text output, we still call text() as it's more efficient than getting runs.
        if self.options.include_styles {
            // Get runs once - this parses the paragraph XML
            let runs = para.runs()?;

            // Derive text from runs for list detection (cheaper than parsing XML again)
            let text = self.extract_text_from_runs(&runs)?;

            // Check if this is a list item
            if let Some(list_info) = self.detect_list_item(&text) {
                self.write_list_item_from_runs(&runs, &list_info)?;
            } else {
                // Write runs with style information
                for run in runs {
                    self.write_run(&run)?;
                }
            }
        } else {
            // Plain text mode - just get text directly (single XML parse)
            let text = para.text()?;

            // Check if this is a list item
            if let Some(list_info) = self.detect_list_item(&text) {
                // For plain text lists, we can just write the content directly
                let indent = " ".repeat(list_info.level * self.options.list_indent);
                let marker = match list_info.list_type {
                    ListType::Ordered => {
                        // Normalize to markdown style "1."
                        if list_info.marker.contains('.') {
                            list_info.marker.clone()
                        } else if list_info.marker.starts_with('(')
                            && list_info.marker.ends_with(')')
                        {
                            let inner = &list_info.marker[1..list_info.marker.len() - 1];
                            format!("{}.", inner)
                        } else {
                            list_info.marker.replace(')', ".")
                        }
                    },
                    ListType::Unordered => "-".to_string(),
                };
                write!(self.buffer, "{}{} {}", indent, marker, list_info.content)
                    .map_err(|e| Error::Other(e.to_string()))?;
            } else {
                // Write plain text
                self.buffer.push_str(&text);
            }
        }

        // Add paragraph break
        self.buffer.push_str("\n\n");
        Ok(())
    }

    /// Write a paragraph that contains display-level formulas.
    ///
    /// This handles paragraphs where formulas are direct children of the paragraph (not within runs).
    #[cfg(all(feature = "ooxml", feature = "formula"))]
    fn write_paragraph_with_display_formulas(
        &mut self,
        para: &Paragraph,
        display_formulas: Vec<String>,
    ) -> Result<()> {
        use crate::formula::omml_to_latex;

        // For display formulas, we'll write each formula on its own line
        // and interleave with any text content from runs
        let runs = para.runs()?;

        // Write all runs first (if any)
        for run in runs {
            let text = run.text()?;
            if !text.trim().is_empty() {
                self.buffer.push_str(&text);
            }
        }

        // Add line break if there was text before formulas
        if !self.buffer.ends_with("\n\n") && !self.buffer.ends_with('\n') {
            self.buffer.push('\n');
        }

        // Write display formulas
        for omml_xml in display_formulas {
            let latex = match omml_to_latex(&omml_xml) {
                Ok(l) => l,
                Err(_) => "[Formula conversion error]".to_string(),
            };

            // Display formulas use display style (false = display mode)
            let formula_md = self.format_formula(&latex, false);
            self.buffer.push_str(&formula_md);
            self.buffer.push('\n');
        }

        Ok(())
    }

    /// Fallback for when formula feature is not enabled.
    #[cfg(all(feature = "ooxml", not(feature = "formula")))]
    fn write_paragraph_with_display_formulas(
        &mut self,
        para: &Paragraph,
        display_formulas: Vec<String>,
    ) -> Result<()> {
        // Write runs normally
        let runs = para.runs()?;
        for run in runs {
            let text = run.text()?;
            if !text.trim().is_empty() {
                self.buffer.push_str(&text);
            }
        }

        // Add placeholder for formulas
        for _ in display_formulas {
            self.buffer
                .push_str("\n[Formula - enable 'formula' feature]\n");
        }

        Ok(())
    }

    /// Write a run with formatting.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    ///
    /// **Performance**: For OOXML runs, this uses a single XML parse to extract both
    /// text and properties simultaneously, providing 2x speedup over separate calls.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    pub fn write_run(&mut self, run: &Run) -> Result<()> {
        // First check if this run contains a formula
        if let Some(formula_markdown) = self.extract_formula_from_run(run)? {
            self.buffer.push_str(&formula_markdown);
            return Ok(());
        }

        // OPTIMIZATION: Get text AND properties in a single XML parse
        // This is 2x faster than calling text() then get_properties()
        #[cfg(feature = "ooxml")]
        let (text, bold, italic, strikethrough, vertical_pos) =
            if let crate::document::Run::Docx(docx_run) = run {
                let (text, props) = docx_run.get_text_and_properties()?;
                if text.is_empty() {
                    return Ok(());
                }
                (
                    text,
                    props.bold.unwrap_or(false),
                    props.italic.unwrap_or(false),
                    props.strikethrough.unwrap_or(false),
                    props.vertical_position,
                )
            } else {
                // Fallback for non-OOXML runs (e.g., OLE format)
                let text = run.text()?;
                if text.is_empty() {
                    return Ok(());
                }
                (
                    text.to_string(),
                    run.bold()?.unwrap_or(false),
                    run.italic()?.unwrap_or(false),
                    run.strikethrough()?.unwrap_or(false),
                    run.vertical_position()?,
                )
            };

        #[cfg(all(feature = "ole", not(feature = "ooxml")))]
        let (text, bold, italic, strikethrough, vertical_pos) = {
            let text = run.text()?;
            if text.is_empty() {
                return Ok(());
            }
            (
                text.to_string(),
                run.bold()?.unwrap_or(false),
                run.italic()?.unwrap_or(false),
                run.strikethrough()?.unwrap_or(false),
                run.vertical_position()?,
            )
        };

        // Handle vertical position (superscript/subscript)
        // Note: vertical_position() is available when ole or ooxml features are enabled
        #[cfg(any(feature = "ole", feature = "ooxml"))]
        {
            use crate::common::VerticalPosition;

            // Pre-calculate buffer size needed to minimize reallocations
            let mut needed_capacity = text.len();
            if vertical_pos.is_some() {
                needed_capacity += 11; // <sup></sup> or <sub></sub>
            }
            if strikethrough {
                needed_capacity += 9; // ~~ or <del></del>
            }
            if bold && italic {
                needed_capacity += 6; // ***
            } else if bold || italic {
                needed_capacity += 4; // ** or *
            }

            // Reserve capacity to avoid reallocations
            self.buffer.reserve(needed_capacity);

            // For superscript/subscript, we apply them directly and skip other formatting
            if let Some(pos) = vertical_pos {
                match self.options.script_style {
                    super::config::ScriptStyle::Html => match pos {
                        VerticalPosition::Superscript => {
                            self.buffer.push_str("<sup>");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("</sup>");
                        },
                        VerticalPosition::Subscript => {
                            self.buffer.push_str("<sub>");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("</sub>");
                        },
                        VerticalPosition::Normal => {
                            self.buffer.push_str(&text);
                        },
                    },
                    super::config::ScriptStyle::Unicode => {
                        // Convert to Unicode superscript/subscript characters
                        // Fall back to HTML tags for characters without Unicode equivalents
                        match pos {
                            VerticalPosition::Superscript => {
                                if super::unicode::can_convert_to_superscript(&text) {
                                    // All characters can be converted to superscript
                                    let converted = super::unicode::convert_to_superscript(&text);
                                    self.buffer.push_str(&converted);
                                } else {
                                    // Fall back to HTML for partial support
                                    self.buffer.push_str("<sup>");
                                    self.buffer.push_str(&text);
                                    self.buffer.push_str("</sup>");
                                }
                            },
                            VerticalPosition::Subscript => {
                                if super::unicode::can_convert_to_subscript(&text) {
                                    // All characters can be converted to subscript
                                    let converted = super::unicode::convert_to_subscript(&text);
                                    self.buffer.push_str(&converted);
                                } else {
                                    // Fall back to HTML for partial support
                                    self.buffer.push_str("<sub>");
                                    self.buffer.push_str(&text);
                                    self.buffer.push_str("</sub>");
                                }
                            },
                            VerticalPosition::Normal => {
                                self.buffer.push_str(&text);
                            },
                        }
                    },
                }
                return Ok(());
            }
        }

        // Pre-calculate buffer size for non-vertical-position formatting
        #[cfg(not(any(feature = "ole", feature = "ooxml")))]
        {
            let mut needed_capacity = text.len();
            if strikethrough {
                needed_capacity += 9; // ~~ or <del></del>
            }
            if bold && italic {
                needed_capacity += 6; // ***
            } else if bold || italic {
                needed_capacity += 4; // ** or *
            }
            self.buffer.reserve(needed_capacity);
        }

        // Apply strikethrough and bold/italic formatting
        if strikethrough {
            match self.options.strikethrough_style {
                super::config::StrikethroughStyle::Markdown => match (bold, italic) {
                    (true, true) => {
                        self.buffer.push_str("~~***");
                        self.buffer.push_str(&text);
                        self.buffer.push_str("***~~");
                    },
                    (true, false) => {
                        self.buffer.push_str("~~**");
                        self.buffer.push_str(&text);
                        self.buffer.push_str("**~~");
                    },
                    (false, true) => {
                        self.buffer.push_str("~~*");
                        self.buffer.push_str(&text);
                        self.buffer.push_str("*~~");
                    },
                    (false, false) => {
                        self.buffer.push_str("~~");
                        self.buffer.push_str(&text);
                        self.buffer.push_str("~~");
                    },
                },
                super::config::StrikethroughStyle::Html => {
                    self.buffer.push_str("<del>");
                    match (bold, italic) {
                        (true, true) => {
                            self.buffer.push_str("***");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("***");
                        },
                        (true, false) => {
                            self.buffer.push_str("**");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("**");
                        },
                        (false, true) => {
                            self.buffer.push('*');
                            self.buffer.push_str(&text);
                            self.buffer.push('*');
                        },
                        (false, false) => {
                            self.buffer.push_str(&text);
                        },
                    }
                    self.buffer.push_str("</del>");
                },
            }
        } else {
            // Apply bold/italic only
            match (bold, italic) {
                (true, true) => {
                    self.buffer.push_str("***");
                    self.buffer.push_str(&text);
                    self.buffer.push_str("***");
                },
                (true, false) => {
                    self.buffer.push_str("**");
                    self.buffer.push_str(&text);
                    self.buffer.push_str("**");
                },
                (false, true) => {
                    self.buffer.push('*');
                    self.buffer.push_str(&text);
                    self.buffer.push('*');
                },
                (false, false) => {
                    self.buffer.push_str(&text);
                },
            }
        }

        Ok(())
    }

    /// Write a table to the buffer.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    pub fn write_table(&mut self, table: &Table) -> Result<()> {
        // Check if table has merged cells
        let has_merged_cells = self.table_has_merged_cells(table)?;

        match self.options.table_style {
            TableStyle::Markdown if !has_merged_cells => {
                self.write_markdown_table(table)?;
            },
            TableStyle::MinimalHtml | TableStyle::Markdown => {
                self.write_html_table(table, false)?;
            },
            TableStyle::StyledHtml => {
                self.write_html_table(table, true)?;
            },
        }

        // Add spacing after table
        self.buffer.push_str("\n\n");
        Ok(())
    }

    /// Check if a table has merged cells.
    ///
    /// Uses multiple heuristics to detect merged cells:
    /// - Inconsistent cell counts across rows
    /// - Empty cells in positions where content is expected
    /// - Cell spans larger than 1 (when available)
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn table_has_merged_cells(&self, table: &Table) -> Result<bool> {
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(false);
        }

        // Check for inconsistent cell counts - avoid allocating cell vectors
        let mut max_cells = 0;
        let mut min_cells = usize::MAX;

        for row in &rows {
            let cell_count = row.cell_count()?;
            max_cells = max_cells.max(cell_count);
            min_cells = min_cells.min(cell_count);
        }

        // If cell counts vary significantly, likely merged cells
        if max_cells > min_cells {
            return Ok(true);
        }

        // Check for empty cells in patterns that suggest merging
        // This is a heuristic: if we have empty cells surrounded by content,
        // it might indicate horizontal merging
        for row in &rows {
            let cells = row.cells()?;
            if cells.len() < 2 {
                continue;
            }

            let mut empty_streak = 0;
            for cell in &cells {
                let cell_text = cell.text()?;
                let text = cell_text.trim();
                if text.is_empty() {
                    empty_streak += 1;
                    // Multiple consecutive empty cells suggest merging
                    if empty_streak >= 2 {
                        return Ok(true);
                    }
                } else {
                    empty_streak = 0;
                }
            }
        }

        // For more advanced detection, we could check:
        // - Cell spans (gridSpan, rowspan attributes)
        // - Vertical merging (vMerge attributes)
        // But these require deeper parsing of the underlying formats

        Ok(false)
    }

    /// Write a table in Markdown format.
    ///
    /// **Performance**: Uses efficient single-pass escaping and minimizes allocations.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn write_markdown_table(&mut self, table: &Table) -> Result<()> {
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(());
        }

        // Pre-allocate buffer capacity
        let total_cells: usize = rows.iter().map(|r| r.cell_count().unwrap_or(0)).sum();
        self.buffer.reserve(total_cells * 50); // Estimate: ~50 bytes per cell

        // Write header row (first row)
        let first_row = &rows[0];
        let first_row_cells = first_row.cells()?;
        let cell_count = first_row_cells.len();

        self.buffer.push('|');
        for cell in &first_row_cells {
            let text = cell.text()?;
            self.buffer.push(' ');
            // Escape pipe and newline in a single pass
            self.write_markdown_escaped(&text);
            self.buffer.push_str(" |");
        }
        self.buffer.push('\n');

        // Write separator row
        self.buffer.push('|');
        for _ in 0..cell_count {
            self.buffer.push_str("----------|");
        }
        self.buffer.push('\n');

        // Write data rows
        for row in &rows[1..] {
            self.buffer.push('|');
            let cells = row.cells()?;
            for cell in &cells {
                let text = cell.text()?;
                self.buffer.push(' ');
                self.write_markdown_escaped(&text);
                self.buffer.push_str(" |");
            }
            self.buffer.push('\n');
        }

        Ok(())
    }

    /// Write markdown-escaped text (escape | and convert \n to space) directly to buffer.
    ///
    /// **Performance**: Single-pass escaping without intermediate allocations.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn write_markdown_escaped(&mut self, text: &str) {
        let bytes = text.as_bytes();
        let mut pos = 0;

        while pos < bytes.len() {
            // Use memchr to quickly find the next character that needs escaping
            let next_special = if let Some(pipe_pos) = memchr(b'|', &bytes[pos..]) {
                if let Some(newline_pos) = memchr(b'\n', &bytes[pos..]) {
                    pos + pipe_pos.min(newline_pos)
                } else {
                    pos + pipe_pos
                }
            } else if let Some(newline_pos) = memchr(b'\n', &bytes[pos..]) {
                pos + newline_pos
            } else {
                // No more special characters, write rest and return
                if pos < bytes.len() {
                    self.buffer.push_str(&text[pos..]);
                }
                return;
            };

            // Write everything up to the special character
            if next_special > pos {
                self.buffer.push_str(&text[pos..next_special]);
            }

            // Write the escape sequence
            match bytes[next_special] {
                b'|' => self.buffer.push_str("\\|"),
                b'\n' => self.buffer.push(' '),
                _ => unreachable!(),
            }

            pos = next_special + 1;
        }
    }

    /// Write a table in HTML format.
    ///
    /// **Performance**: Uses efficient single-pass HTML escaping and minimizes allocations.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn write_html_table(&mut self, table: &Table, styled: bool) -> Result<()> {
        let indent = " ".repeat(self.options.html_table_indent);
        let double_indent = format!("{}{}", indent, indent);

        if styled {
            self.buffer.push_str("<table class=\"doc-table\">\n");
        } else {
            self.buffer.push_str("<table>\n");
        }

        let rows = table.rows()?;

        // Pre-allocate buffer capacity to reduce reallocations
        // Estimate: ~100 bytes per cell on average
        let total_cells: usize = rows.iter().map(|r| r.cell_count().unwrap_or(0)).sum();
        self.buffer.reserve(total_cells * 100);

        for (i, row) in rows.iter().enumerate() {
            // First row is typically header
            let tag = if i == 0 { "th" } else { "td" };

            self.buffer.push_str(&indent);
            self.buffer.push_str("<tr>\n");

            let cells = row.cells()?;
            for cell in &cells {
                let text = cell.text()?;

                // Write opening tag
                self.buffer.push_str(&double_indent);
                self.buffer.push('<');
                self.buffer.push_str(tag);
                self.buffer.push('>');

                // HTML escape and write text in a single pass (no intermediate allocations)
                self.write_html_escaped(&text);

                // Write closing tag
                self.buffer.push_str("</");
                self.buffer.push_str(tag);
                self.buffer.push_str(">\n");
            }

            self.buffer.push_str(&indent);
            self.buffer.push_str("</tr>\n");
        }

        self.buffer.push_str("</table>");
        Ok(())
    }

    /// Write HTML-escaped text directly to the buffer without intermediate allocations.
    ///
    /// **Performance**: Single-pass escaping that writes directly to the buffer,
    /// avoiding the 4 intermediate string allocations from chained `replace()` calls.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn write_html_escaped(&mut self, text: &str) {
        let bytes = text.as_bytes();
        let mut pos = 0;

        while pos < bytes.len() {
            // Find the next character that needs escaping
            let next_special = [b'&', b'<', b'>', b'\n']
                .iter()
                .filter_map(|&ch| memchr(ch, &bytes[pos..]).map(|p| pos + p))
                .min();

            if let Some(special_pos) = next_special {
                // Write everything up to the special character
                if special_pos > pos {
                    self.buffer.push_str(&text[pos..special_pos]);
                }

                // Write the escape sequence
                match bytes[special_pos] {
                    b'&' => self.buffer.push_str("&amp;"),
                    b'<' => self.buffer.push_str("&lt;"),
                    b'>' => self.buffer.push_str("&gt;"),
                    b'\n' => self.buffer.push_str("<br>"),
                    _ => unreachable!(),
                }

                pos = special_pos + 1;
            } else {
                // No more special characters, write rest and return
                if pos < bytes.len() {
                    self.buffer.push_str(&text[pos..]);
                }
                return;
            }
        }
    }

    /// Get the final markdown output.
    pub fn finish(self) -> String {
        self.buffer
    }

    /// Append text to the buffer.
    pub fn push_str(&mut self, text: &str) {
        self.buffer.push_str(text);
    }

    /// Append a single character to the buffer.
    pub fn push(&mut self, ch: char) {
        self.buffer.push(ch);
    }

    /// Write a formatted string to the buffer.
    pub fn write_fmt(&mut self, args: std::fmt::Arguments) -> Result<()> {
        use std::fmt::Write as FmtWrite;
        self.buffer
            .write_fmt(args)
            .map_err(|e| Error::Other(e.to_string()))
    }

    /// Reserve additional capacity in the buffer.
    pub fn reserve(&mut self, additional: usize) {
        self.buffer.reserve(additional);
    }

    /// Write document metadata as YAML front matter.
    ///
    /// If metadata is available and include_metadata is enabled,
    /// this writes the metadata as YAML front matter at the beginning of the document.
    pub fn write_metadata(&mut self, metadata: &Metadata) -> Result<()> {
        if !self.options.include_metadata {
            return Ok(());
        }

        let yaml_front_matter = metadata
            .to_yaml_front_matter()
            .map_err(|e| Error::Other(format!("Failed to generate YAML front matter: {}", e)))?;

        if !yaml_front_matter.is_empty() {
            self.buffer.push_str(&yaml_front_matter);
        }

        Ok(())
    }

    /// Detect if a paragraph is a list item and extract list information.
    fn detect_list_item(&self, text: &str) -> Option<ListItemInfo> {
        let text = text.trim_start();

        // Check for ordered lists: 1. 2. 3. or 1) 2) 3) or (1) (2) (3)
        if let Some(captures) = self.extract_ordered_list_marker(text) {
            let marker = captures.0;
            let content = captures.1;
            let level = self.calculate_indent_level(text);
            return Some(ListItemInfo {
                list_type: ListType::Ordered,
                level,
                marker: marker.to_string(),
                content: content.to_string(),
            });
        }

        // Check for unordered lists: - * •
        if let Some(captures) = self.extract_unordered_list_marker(text) {
            let marker = captures.0;
            let content = captures.1;
            let level = self.calculate_indent_level(text);
            return Some(ListItemInfo {
                list_type: ListType::Unordered,
                level,
                marker: marker.to_string(),
                content: content.to_string(),
            });
        }

        None
    }

    /// Extract ordered list marker and content.
    fn extract_ordered_list_marker<'a>(&self, text: &'a str) -> Option<(&'a str, &'a str)> {
        // Match patterns like: "1. ", "2) ", "(1) ", etc.
        if let Some(pos) = text.find('.')
            && pos > 0
            && text[..pos].chars().all(|c| c.is_ascii_digit())
        {
            let marker_end = pos + 1;
            if text.len() > marker_end && text.as_bytes()[marker_end] == b' ' {
                return Some((&text[..marker_end], &text[marker_end + 1..]));
            }
        }

        if let Some(pos) = text.find(')')
            && pos > 0
            && text[..pos].chars().all(|c| c.is_ascii_digit())
        {
            let marker_end = pos + 1;
            if text.len() > marker_end && text.as_bytes()[marker_end] == b' ' {
                return Some((&text[..marker_end], &text[marker_end + 1..]));
            }
        }

        // Check for parenthesized numbers: (1) (2) (3)
        if text.starts_with('(')
            && let Some(end_pos) = text.find(") ")
        {
            let inner = &text[1..end_pos];
            if inner.chars().all(|c| c.is_ascii_digit()) {
                return Some((&text[..end_pos + 1], &text[end_pos + 2..]));
            }
        }

        None
    }

    /// Extract unordered list marker and content.
    fn extract_unordered_list_marker<'a>(&self, text: &'a str) -> Option<(&'a str, &'a str)> {
        let markers = ["-", "*", "•"];

        for &marker in &markers {
            if let Some(remaining) = text.strip_prefix(marker)
                && (remaining.starts_with(' ') || remaining.starts_with('\t'))
            {
                return Some((marker, remaining.trim_start()));
            }
        }

        None
    }

    /// Calculate the indentation level based on leading spaces/tabs.
    fn calculate_indent_level(&self, text: &str) -> usize {
        let leading = text.len() - text.trim_start().len();
        // Each indent level corresponds to list_indent spaces
        leading / self.options.list_indent
    }

    /// Extract formula content from a run and convert to markdown.
    ///
    /// Returns the markdown representation of the formula if one is found, None otherwise.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn extract_formula_from_run(&self, run: &Run) -> Result<Option<String>> {
        // Try OOXML OMML formulas first
        #[cfg(feature = "ooxml")]
        if let crate::document::Run::Docx(docx_run) = run
            && let Some(omml_xml) = docx_run.omml_formula()?
        {
            // Parse OMML and convert to LaTeX
            #[cfg(feature = "formula")]
            {
                let latex = self.convert_omml_to_latex(&omml_xml);
                return Ok(Some(self.format_formula(&latex, true))); // true = inline
            }

            #[cfg(not(feature = "formula"))]
            {
                // omml_xml is captured but not used when formula feature is disabled
                let _ = omml_xml;
                return Ok(Some(
                    self.format_formula("[Formula - enable 'formula' feature]", true),
                ));
            }
        }

        // Try OLE MTEF formulas
        #[cfg(feature = "ole")]
        {
            // When only ole feature is enabled, Run can only be Doc variant
            let ole_run = match run {
                crate::document::Run::Doc(r) => r,
                #[cfg(feature = "ooxml")]
                _ => return Ok(None),
            };

            if ole_run.has_mtef_formula() {
                // Get the MTEF formula AST
                if let Some(mtef_ast) = ole_run.mtef_formula_ast() {
                    // Convert MTEF AST to LaTeX
                    let latex = self.convert_mtef_to_latex(mtef_ast);
                    return Ok(Some(self.format_formula(&latex, true))); // true = inline
                } else {
                    // Fallback placeholder if AST is not available
                    return Ok(Some(self.format_formula("[Formula]", true)));
                }
            }
        }

        Ok(None)
    }

    /// Convert MTEF AST nodes to LaTeX string
    #[cfg(feature = "formula")]
    fn convert_mtef_to_latex(&self, nodes: &[crate::formula::MathNode]) -> String {
        use crate::formula::latex::LatexConverter;

        let mut converter = LatexConverter::new();
        match converter.convert_nodes(nodes) {
            Ok(latex) => latex.to_string(),
            Err(_) => "[Formula conversion error]".to_string(),
        }
    }

    /// Convert MTEF AST nodes to LaTeX string (fallback when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    fn convert_mtef_to_latex(&self, _nodes: &[()]) -> String {
        "[Formula support disabled - enable 'formula' feature]".to_string()
    }

    /// Convert OMML XML to LaTeX string
    #[cfg(all(feature = "ooxml", feature = "formula"))]
    #[allow(dead_code)] // Used conditionally based on feature flags
    fn convert_omml_to_latex(&self, omml_xml: &str) -> String {
        use crate::formula::omml_to_latex;

        // Use the high-level conversion function
        match omml_to_latex(omml_xml) {
            Ok(latex) => latex,
            Err(_) => "[Formula conversion error]".to_string(),
        }
    }

    /// Convert OMML XML to LaTeX string (fallback when formula feature is disabled)
    #[cfg(all(feature = "ooxml", not(feature = "formula")))]
    #[allow(dead_code)] // Used conditionally based on feature flags
    fn convert_omml_to_latex(&self, _omml_xml: &str) -> String {
        "[Formula support disabled - enable 'formula' feature]".to_string()
    }

    /// Format a formula with the appropriate delimiters.
    ///
    /// # Arguments
    /// * `formula` - The formula content (LaTeX)
    /// * `inline` - Whether this is an inline formula (true) or display formula (false)
    fn format_formula(&self, formula: &str, inline: bool) -> String {
        if inline {
            match self.options.formula_style {
                super::config::FormulaStyle::LaTeX => format!("\\({}\\)", formula),
                super::config::FormulaStyle::Dollar => format!("${}$", formula),
            }
        } else {
            match self.options.formula_style {
                super::config::FormulaStyle::LaTeX => format!("\\[{}\\]", formula),
                super::config::FormulaStyle::Dollar => format!("$${}$$", formula),
            }
        }
    }

    /// Format a formula placeholder with the appropriate delimiters.
    #[allow(dead_code)]
    fn format_formula_placeholder(&self, placeholder: &str) -> String {
        self.format_formula(placeholder, true)
    }

    /// Write a list item with proper formatting.
    #[allow(dead_code)] // Used in fallback paths
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn write_list_item(&mut self, _para: &Paragraph, list_info: &ListItemInfo) -> Result<()> {
        // Add indentation for nested lists
        let indent = " ".repeat(list_info.level * self.options.list_indent);

        // Generate the appropriate marker
        let marker = match list_info.list_type {
            ListType::Ordered => {
                // For ordered lists, we need to determine the number
                // For now, use a simple approach - in a real implementation
                // we'd track list state across paragraphs
                if list_info.marker.contains('.') {
                    // Keep "1." as is
                    list_info.marker.clone()
                } else {
                    // Convert "1)" or "(1)" to "1." for markdown
                    if list_info.marker.starts_with('(') && list_info.marker.ends_with(')') {
                        // Extract number from (1) -> 1.
                        let inner = &list_info.marker[1..list_info.marker.len() - 1];
                        format!("{}.", inner)
                    } else {
                        // Convert "1)" to "1."
                        list_info.marker.replace(')', ".")
                    }
                }
            },
            ListType::Unordered => "-".to_string(),
        };

        // Write the list item
        write!(self.buffer, "{}{} ", indent, marker).map_err(|e| Error::Other(e.to_string()))?;

        // Write the content with styles if enabled
        if self.options.include_styles && !list_info.content.trim().is_empty() {
            // For styled content, we need to skip the marker part and write the remaining runs
            // This is a simplified approach - in practice, we'd need more sophisticated
            // parsing to handle cases where the marker spans multiple runs
            self.buffer.push_str(&list_info.content);
        } else {
            // Write the content directly
            self.buffer.push_str(&list_info.content);
        }

        Ok(())
    }

    /// Extract text from runs without re-parsing paragraph XML.
    ///
    /// **Performance**: This is much faster than calling `para.text()` when we already
    /// have the runs, as it avoids re-parsing the paragraph XML.
    ///
    /// For OOXML runs, this method is optimized to extract only text efficiently.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn extract_text_from_runs(&self, runs: &[Run]) -> Result<String> {
        // Pre-allocate capacity based on number of runs
        let mut text = String::with_capacity(runs.len() * 32);

        for run in runs {
            // For OOXML, just extract text without parsing properties
            // since we only need text for list detection
            let run_text = run.text()?;
            text.push_str(&run_text);
        }

        Ok(text)
    }

    /// Write a list item from runs with proper formatting.
    ///
    /// **Performance**: Takes pre-parsed runs to avoid re-parsing XML.
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    fn write_list_item_from_runs(&mut self, runs: &[Run], list_info: &ListItemInfo) -> Result<()> {
        // Add indentation for nested lists
        let indent = " ".repeat(list_info.level * self.options.list_indent);

        // Generate the appropriate marker
        let marker = match list_info.list_type {
            ListType::Ordered => {
                // Normalize to markdown style "1."
                if list_info.marker.contains('.') {
                    list_info.marker.clone()
                } else if list_info.marker.starts_with('(') && list_info.marker.ends_with(')') {
                    let inner = &list_info.marker[1..list_info.marker.len() - 1];
                    format!("{}.", inner)
                } else {
                    list_info.marker.replace(')', ".")
                }
            },
            ListType::Unordered => "-".to_string(),
        };

        // Write the list item marker
        write!(self.buffer, "{}{} ", indent, marker).map_err(|e| Error::Other(e.to_string()))?;

        // Write runs, skipping the list marker portion
        // This is a simplified approach - we write all runs with their formatting
        // A more sophisticated implementation would skip the marker text in the first run
        let mut accumulated_len = 0;
        let marker_end_pos = list_info.marker.len() + 1; // marker + space

        for run in runs {
            // OPTIMIZATION: Get text first to check if we need to skip/process this run
            // Only parse properties if we actually need to write the run
            let run_text = run.text()?;
            let run_len = run_text.len();

            // Skip runs that are part of the marker
            if accumulated_len + run_len <= marker_end_pos {
                accumulated_len += run_len;
                continue;
            }

            // Partial skip if run contains marker end
            if accumulated_len < marker_end_pos && accumulated_len + run_len > marker_end_pos {
                let skip_chars = marker_end_pos - accumulated_len;
                // Write the portion after the marker
                let text_after_marker = &run_text[skip_chars..];

                // Create a temporary run-like structure with the remaining text
                // For now, just write the text - ideally we'd preserve formatting
                self.buffer.push_str(text_after_marker);
                accumulated_len += run_len;
            } else {
                // Write the entire run with formatting
                self.write_run(run)?;
                accumulated_len += run_len;
            }
        }

        Ok(())
    }
}
