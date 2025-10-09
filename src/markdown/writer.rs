/// Low-level writer for Markdown generation.
///
/// This module provides the `MarkdownWriter` struct which handles the actual
/// conversion of document elements to Markdown format.
use crate::common::{Error, Result};
use crate::document::{Paragraph, Run, Table};
use super::config::{MarkdownOptions, TableStyle};
use std::fmt::Write as FmtWrite;

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
    pub fn write_paragraph(&mut self, para: &Paragraph) -> Result<()> {
        let text = para.text()?;

        // Check if this is a list item
        if let Some(list_info) = self.detect_list_item(&text) {
            self.write_list_item(para, &list_info)?;
        } else {
            // Regular paragraph
            if self.options.include_styles {
                // Write runs with style information
                let runs = para.runs()?;
                for run in runs {
                    self.write_run(&run)?;
                }
            } else {
                // Write plain text
                self.buffer.push_str(&text);
            }
        }

        // Add paragraph break
        self.buffer.push_str("\n\n");
        Ok(())
    }

    /// Write a run with formatting.
    pub fn write_run(&mut self, run: &Run) -> Result<()> {
        // First check if this run contains a formula
        if let Some(formula_markdown) = self.extract_formula_from_run(run)? {
            self.buffer.push_str(&formula_markdown);
            return Ok(());
        }

        let text = run.text()?;
        if text.is_empty() {
            return Ok(());
        }

        let bold = run.bold()?.unwrap_or(false);
        let italic = run.italic()?.unwrap_or(false);
        let strikethrough = run.strikethrough()?.unwrap_or(false);
        let vertical_pos = run.vertical_position()?;

        // Pre-calculate buffer size needed to minimize reallocations
        let mut needed_capacity = text.len();
        if let Some(_) = vertical_pos {
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
                super::config::ScriptStyle::Html => {
                    match pos {
                        crate::ole::doc::parts::chp::VerticalPosition::Superscript => {
                            self.buffer.push_str("<sup>");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("</sup>");
                        }
                        crate::ole::doc::parts::chp::VerticalPosition::Subscript => {
                            self.buffer.push_str("<sub>");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("</sub>");
                        }
                        _ => {
                            self.buffer.push_str(&text);
                        }
                    }
                }
                super::config::ScriptStyle::Unicode => {
                    // For Unicode, we'd need to convert each character to superscript/subscript
                    // This is complex, so for now fall back to HTML for unsupported characters
                    match pos {
                        crate::ole::doc::parts::chp::VerticalPosition::Superscript => {
                            self.buffer.push_str("<sup>");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("</sup>");
                        }
                        crate::ole::doc::parts::chp::VerticalPosition::Subscript => {
                            self.buffer.push_str("<sub>");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("</sub>");
                        }
                        _ => {
                            self.buffer.push_str(&text);
                        }
                    }
                }
            }
            return Ok(());
        }

        // Apply strikethrough and bold/italic formatting
        if strikethrough {
            match self.options.strikethrough_style {
                super::config::StrikethroughStyle::Markdown => {
                    match (bold, italic) {
                        (true, true) => {
                            self.buffer.push_str("~~***");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("***~~");
                        }
                        (true, false) => {
                            self.buffer.push_str("~~**");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("**~~");
                        }
                        (false, true) => {
                            self.buffer.push_str("~~*");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("*~~");
                        }
                        (false, false) => {
                            self.buffer.push_str("~~");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("~~");
                        }
                    }
                }
                super::config::StrikethroughStyle::Html => {
                    self.buffer.push_str("<del>");
                    match (bold, italic) {
                        (true, true) => {
                            self.buffer.push_str("***");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("***");
                        }
                        (true, false) => {
                            self.buffer.push_str("**");
                            self.buffer.push_str(&text);
                            self.buffer.push_str("**");
                        }
                        (false, true) => {
                            self.buffer.push('*');
                            self.buffer.push_str(&text);
                            self.buffer.push('*');
                        }
                        (false, false) => {
                            self.buffer.push_str(&text);
                        }
                    }
                    self.buffer.push_str("</del>");
                }
            }
        } else {
            // Apply bold/italic only
            match (bold, italic) {
                (true, true) => {
                    self.buffer.push_str("***");
                    self.buffer.push_str(&text);
                    self.buffer.push_str("***");
                }
                (true, false) => {
                    self.buffer.push_str("**");
                    self.buffer.push_str(&text);
                    self.buffer.push_str("**");
                }
                (false, true) => {
                    self.buffer.push('*');
                    self.buffer.push_str(&text);
                    self.buffer.push('*');
                }
                (false, false) => {
                    self.buffer.push_str(&text);
                }
            }
        }

        Ok(())
    }

    /// Write a table to the buffer.
    pub fn write_table(&mut self, table: &Table) -> Result<()> {
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
    /// Uses multiple heuristics to detect merged cells:
    /// - Inconsistent cell counts across rows
    /// - Empty cells in positions where content is expected
    /// - Cell spans larger than 1 (when available)
    fn table_has_merged_cells(&self, table: &Table) -> Result<bool> {
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(false);
        }

        // Check for inconsistent cell counts
        let cell_counts: Vec<usize> = rows.iter()
            .map(|row| row.cells().map(|cells| cells.len()).unwrap_or(0))
            .collect();

        let max_cells = cell_counts.iter().max().unwrap_or(&0);
        let min_cells = cell_counts.iter().min().unwrap_or(&0);

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
    pub fn finish(self) -> String {
        self.buffer
    }

    /// Append text to the buffer.
    pub fn push_str(&mut self, text: &str) {
        self.buffer.push_str(&text);
    }

    /// Append a single character to the buffer.
    pub fn push(&mut self, ch: char) {
        self.buffer.push(ch);
    }

    /// Write a formatted string to the buffer.
    pub fn write_fmt(&mut self, args: std::fmt::Arguments) -> Result<()> {
        use std::fmt::Write as FmtWrite;
        self.buffer.write_fmt(args).map_err(|e| Error::Other(e.to_string()))
    }

    /// Reserve additional capacity in the buffer.
    pub fn reserve(&mut self, additional: usize) {
        self.buffer.reserve(additional);
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
        if let Some(pos) = text.find('.') {
            if pos > 0 && text[..pos].chars().all(|c| c.is_ascii_digit()) {
                let marker_end = pos + 1;
                if text.len() > marker_end && text.as_bytes()[marker_end] == b' ' {
                    return Some((&text[..marker_end], &text[marker_end + 1..]));
                }
            }
        }

        if let Some(pos) = text.find(')') {
            if pos > 0 && text[..pos].chars().all(|c| c.is_ascii_digit()) {
                let marker_end = pos + 1;
                if text.len() > marker_end && text.as_bytes()[marker_end] == b' ' {
                    return Some((&text[..marker_end], &text[marker_end + 1..]));
                }
            }
        }

        // Check for parenthesized numbers: (1) (2) (3)
        if text.starts_with('(') {
            if let Some(end_pos) = text.find(") ") {
                let inner = &text[1..end_pos];
                if inner.chars().all(|c| c.is_ascii_digit()) {
                    return Some((&text[..end_pos + 1], &text[end_pos + 2..]));
                }
            }
        }

        None
    }

    /// Extract unordered list marker and content.
    fn extract_unordered_list_marker<'a>(&self, text: &'a str) -> Option<(&'a str, &'a str)> {
        let markers = ["-", "*", "•"];

        for &marker in &markers {
            if let Some(remaining) = text.strip_prefix(marker) {
                if remaining.starts_with(' ') || remaining.starts_with('\t') {
                    return Some((marker, remaining.trim_start()));
                }
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
    fn extract_formula_from_run(&self, run: &Run) -> Result<Option<String>> {
        // Try OOXML OMML formulas first
        if let crate::document::Run::Docx(docx_run) = run {
            if let Some(_omml_xml) = docx_run.omml_formula()? {
                // For now, return a placeholder. In a full implementation,
                // this would parse the OMML XML and convert to LaTeX/markdown
                return Ok(Some(self.format_formula_placeholder("OMML formula detected")));
            }
        }

        // Try OLE MTEF formulas
        if let crate::document::Run::Doc(ole_run) = run {
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
    fn convert_mtef_to_latex(&self, nodes: &[crate::formula::MathNode]) -> String {
        use crate::formula::latex::LatexConverter;
        
        let mut converter = LatexConverter::new();
        match converter.convert_nodes(nodes) {
            Ok(latex) => latex.to_string(),
            Err(_) => "[Formula conversion error]".to_string(),
        }
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
                        let inner = &list_info.marker[1..list_info.marker.len()-1];
                        format!("{}.", inner)
                    } else {
                        // Convert "1)" to "1."
                        list_info.marker.replace(')', ".")
                    }
                }
            }
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
}

