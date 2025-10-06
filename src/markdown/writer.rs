/// Low-level writer for Markdown generation.
///
/// This module provides the `MarkdownWriter` struct which handles the actual
/// conversion of document elements to Markdown format.

use crate::common::{Error, Result};
use crate::document::{Paragraph, Run, Table};
use super::config::{MarkdownOptions, TableStyle};
use std::fmt::Write as FmtWrite;

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
    pub fn write_run(&mut self, run: &Run) -> Result<()> {
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
    pub fn finish(self) -> String {
        self.buffer
    }
}

