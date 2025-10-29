use super::config::{MarkdownOptions, TableStyle};
/// Low-level writer for Markdown generation.
///
/// This module provides the `MarkdownWriter` struct which handles the actual
/// conversion of document elements to Markdown format.
///
/// **Note**: Some functionality requires the `ole` or `ooxml` feature to be enabled.
use crate::common::{Error, Metadata, Result};
#[cfg(any(feature = "ole", feature = "ooxml"))]
use crate::document::{Cell, Paragraph, Run, Table};
use std::fmt::Write as FmtWrite;

#[cfg(any(feature = "ole", feature = "ooxml"))]
use memchr::memchr;

#[cfg(any(feature = "ole", feature = "ooxml"))]
use rayon::prelude::*;

/// Minimum number of table rows to justify parallel processing overhead.
/// Tables are typically smaller than documents, so we use a lower threshold.
#[cfg(any(feature = "ole", feature = "ooxml"))]
const TABLE_PARALLEL_THRESHOLD: usize = 20;

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

/// Information about cell span (colspan and rowspan) for HTML rendering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellSpan {
    /// Number of columns this cell spans (horizontal merge)
    colspan: usize,
    /// Number of rows this cell spans (vertical merge)
    rowspan: usize,
    /// Whether this cell should be skipped in rendering (it's covered by a merge)
    skip: bool,
}

impl CellSpan {
    /// Create a new cell span with default values (no merge).
    fn new() -> Self {
        Self {
            colspan: 1,
            rowspan: 1,
            skip: false,
        }
    }

    /// Create a cell span that should be skipped.
    fn skipped() -> Self {
        Self {
            colspan: 1,
            rowspan: 1,
            skip: true,
        }
    }
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
    /// Current formatting state to avoid duplicate markers
    current_bold: bool,
    current_italic: bool,
    current_strikethrough: bool,
}

/// Analyze a table to compute cell spans (colspan/rowspan) for proper HTML rendering.
///
/// This function processes a table and computes the actual colspan and rowspan for each cell,
/// taking into account:
/// - `gridSpan` (horizontal merge/colspan)
/// - `vMerge` (vertical merge/rowspan)
///
/// Returns a 2D vector where `result[row][col]` contains the span information for that cell.
///
/// **Performance**: Uses efficient single-pass analysis with minimal allocations.
#[cfg(any(
    feature = "ole",
    feature = "ooxml",
    feature = "odf",
    feature = "rtf",
    feature = "iwa"
))]
fn analyze_table_spans(table: &Table) -> Result<Vec<Vec<CellSpan>>> {
    let rows = table.rows()?;
    if rows.is_empty() {
        return Ok(Vec::new());
    }

    // First pass: determine the maximum grid width (considering gridSpan)
    let mut max_grid_cols = 0;
    for row in &rows {
        let cells = row.cells()?;
        let mut row_grid_cols = 0;
        for cell in &cells {
            let grid_span = cell.grid_span().unwrap_or(1);
            row_grid_cols += grid_span;
        }
        max_grid_cols = max_grid_cols.max(row_grid_cols);
    }

    // Initialize span info for all cells
    let mut spans: Vec<Vec<CellSpan>> = Vec::with_capacity(rows.len());
    for _ in 0..rows.len() {
        spans.push(vec![CellSpan::new(); max_grid_cols]);
    }

    // Second pass: analyze gridSpan and vMerge for each cell
    for (row_idx, row) in rows.iter().enumerate() {
        let cells = row.cells()?;
        let mut grid_col = 0; // Current grid column position

        for cell in &cells {
            // Skip grid columns that are covered by previous cells' colspan
            while grid_col < max_grid_cols && spans[row_idx][grid_col].skip {
                grid_col += 1;
            }

            if grid_col >= max_grid_cols {
                break;
            }

            // Get horizontal span (gridSpan)
            let colspan = cell.grid_span().unwrap_or(1);
            spans[row_idx][grid_col].colspan = colspan;

            // Mark columns covered by this cell's colspan as skipped
            for offset in 1..colspan {
                if grid_col + offset < max_grid_cols {
                    spans[row_idx][grid_col + offset] = CellSpan::skipped();
                }
            }

            // Get vertical merge state (vMerge)
            #[cfg(feature = "ooxml")]
            {
                use crate::ooxml::docx::VMergeState;

                if let Ok(Some(v_merge_state)) = cell.v_merge() {
                    match v_merge_state {
                        VMergeState::Restart => {
                            // This cell starts a vertical merge
                            // Count how many rows below continue this merge
                            let mut rowspan = 1;
                            for next_row_idx in (row_idx + 1)..rows.len() {
                                if let Ok(Some(next_row)) = rows[next_row_idx].cell_at(grid_col) {
                                    if let Ok(Some(VMergeState::Continue)) = next_row.v_merge() {
                                        rowspan += 1;
                                        // Mark this cell as skipped
                                        spans[next_row_idx][grid_col] = CellSpan::skipped();
                                        // Also mark colspan cells as skipped
                                        for offset in 1..colspan {
                                            if grid_col + offset < max_grid_cols {
                                                spans[next_row_idx][grid_col + offset] =
                                                    CellSpan::skipped();
                                            }
                                        }
                                    } else {
                                        break;
                                    }
                                } else {
                                    break;
                                }
                            }
                            spans[row_idx][grid_col].rowspan = rowspan;
                        },
                        VMergeState::Continue => {
                            // This cell continues a merge from above, should be skipped
                            // (already marked in the Restart case above)
                        },
                    }
                }
            }

            grid_col += colspan;
        }
    }

    Ok(spans)
}

impl MarkdownWriter {
    /// Create a new writer with the given options.
    pub fn new(options: MarkdownOptions) -> Self {
        Self {
            buffer: String::with_capacity(4096), // Pre-allocate reasonable size
            options,
            current_bold: false,
            current_italic: false,
            current_strikethrough: false,
        }
    }

    /// Write a paragraph to the buffer.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    ///
    /// **Performance**: Optimized to avoid redundant XML parsing by extracting runs
    /// once and deriving text from them when needed.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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

            // FALLBACK: If no runs found (e.g., ODF paragraphs with direct text), use paragraph text
            if runs.is_empty() {
                let text = para.text()?;
                if !text.is_empty() {
                    // Check if this is a list item
                    if let Some(list_info) = self.detect_list_item(&text) {
                        // For plain text lists, write the content directly
                        let indent = " ".repeat(list_info.level * self.options.list_indent);
                        let marker = match list_info.list_type {
                            ListType::Ordered => {
                                if list_info.marker.contains('.') {
                                    list_info.marker.clone()
                                } else if list_info.marker.starts_with('(') {
                                    format!(
                                        "{}.",
                                        list_info
                                            .marker
                                            .trim_start_matches('(')
                                            .trim_end_matches(')')
                                    )
                                } else {
                                    "1.".to_string()
                                }
                            },
                            ListType::Unordered => "-".to_string(),
                        };
                        self.buffer.push_str(&indent);
                        self.buffer.push_str(&marker);
                        self.buffer.push(' ');
                        self.buffer.push_str(
                            text.trim_start()
                                .trim_start_matches(&list_info.marker)
                                .trim_start(),
                        );
                    } else {
                        // Regular paragraph - just write the text
                        self.buffer.push_str(&text);
                    }
                }
            } else {
                // Has runs - process them normally
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

        // Close any open formatting at paragraph boundary
        self.close_formatting();

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

    /// Close any currently open formatting.
    /// This should be called at paragraph boundaries to ensure clean output.
    fn close_formatting(&mut self) {
        // Close in reverse order of opening (strikethrough -> italic -> bold)
        if self.current_strikethrough {
            self.buffer.push_str("~~");
            self.current_strikethrough = false;
        }
        if self.current_italic {
            self.buffer.push('*');
            self.current_italic = false;
        }
        if self.current_bold {
            self.buffer.push_str("**");
            self.current_bold = false;
        }
    }

    /// Apply formatting changes by closing/opening markers as needed.
    /// Returns the text with appropriate formatting markers applied.
    fn apply_formatting(&mut self, bold: bool, italic: bool, strikethrough: bool) {
        // Determine what needs to change
        let bold_changed = bold != self.current_bold;
        let italic_changed = italic != self.current_italic;
        let strike_changed = strikethrough != self.current_strikethrough;

        // If nothing changed, we're done
        if !bold_changed && !italic_changed && !strike_changed {
            return;
        }

        // Close formatting that's being removed (in reverse order)
        if strike_changed && self.current_strikethrough {
            self.buffer.push_str("~~");
            self.current_strikethrough = false;
        }
        if italic_changed && self.current_italic {
            self.buffer.push('*');
            self.current_italic = false;
        }
        if bold_changed && self.current_bold {
            self.buffer.push_str("**");
            self.current_bold = false;
        }

        // Open new formatting (in forward order)
        if bold_changed && bold {
            self.buffer.push_str("**");
            self.current_bold = true;
        }
        if italic_changed && italic {
            self.buffer.push('*');
            self.current_italic = true;
        }
        if strike_changed && strikethrough {
            self.buffer.push_str("~~");
            self.current_strikethrough = true;
        }
    }

    /// Write a run with formatting.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    ///
    /// **Performance**: For OOXML runs, this uses a single XML parse to extract both
    /// text and properties simultaneously, providing 2x speedup over separate calls.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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

        // Apply formatting changes (only add/remove markers when formatting changes)
        // Note: For HTML strikethrough style, we need special handling since HTML
        // tags can't be left open across runs
        if self.options.strikethrough_style == super::config::StrikethroughStyle::Html
            && strikethrough
        {
            // HTML strikethrough: must be self-contained per run
            // Close any open markdown formatting first
            self.close_formatting();

            // Apply HTML strikethrough with inline markdown formatting
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
        } else {
            // Markdown-style formatting: can span across runs
            self.apply_formatting(bold, italic, strikethrough);
            self.buffer.push_str(&text);
        }

        Ok(())
    }

    /// Write a table to the buffer.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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
    /// Uses proper span analysis to detect merged cells by checking for:
    /// - Horizontal merges (gridSpan/colspan > 1)
    /// - Vertical merges (vMerge/rowspan > 1)
    ///
    /// **Performance**: Efficient analysis that reuses existing span computation.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn table_has_merged_cells(&self, table: &Table) -> Result<bool> {
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(false);
        }

        // Quick check: Look for cells with gridSpan > 1 or vMerge attributes
        for row in &rows {
            let cells = row.cells()?;
            for cell in &cells {
                // Check horizontal merge (gridSpan)
                if cell.grid_span().unwrap_or(1) > 1 {
                    return Ok(true);
                }

                // Check vertical merge (vMerge) - only available for OOXML
                #[cfg(feature = "ooxml")]
                {
                    if cell.v_merge().ok().flatten().is_some() {
                        return Ok(true);
                    }
                }
            }
        }

        Ok(false)
    }

    /// Write a table in Markdown format.
    ///
    /// **Performance**: Uses efficient single-pass escaping and minimizes allocations.
    /// For large tables (20+ rows), uses parallel processing to render rows concurrently.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn write_markdown_table(&mut self, table: &Table) -> Result<()> {
        let rows = table.rows()?;
        if rows.is_empty() {
            return Ok(());
        }

        // Pre-allocate buffer capacity
        let total_cells: usize = rows.iter().map(|r| r.cell_count().unwrap_or(0)).sum();
        self.buffer.reserve(total_cells * 50); // Estimate: ~50 bytes per cell

        // Write header row (first row) - always sequential
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

        // Write data rows - parallel if large enough
        if self.options.use_parallel && rows.len() > TABLE_PARALLEL_THRESHOLD {
            // PARALLEL PATH: Process rows in parallel for large tables
            // First, extract all cell texts sequentially (to avoid borrowing issues with Row enum)
            let mut row_texts: Vec<Vec<String>> = Vec::with_capacity(rows.len() - 1);
            for row in &rows[1..] {
                let cells = row.cells()?;
                let cell_texts: Result<Vec<String>> = cells.iter().map(|c| c.text()).collect();
                row_texts.push(cell_texts?);
            }

            // Now process the texts in parallel
            let row_strings: Vec<String> = row_texts
                .par_iter()
                .map(|cell_texts| {
                    let mut row_buffer = String::with_capacity(cell_texts.len() * 50);
                    row_buffer.push('|');
                    for text in cell_texts {
                        row_buffer.push(' ');
                        Self::escape_markdown_to_buffer(&mut row_buffer, text);
                        row_buffer.push_str(" |");
                    }
                    row_buffer.push('\n');
                    row_buffer
                })
                .collect();

            // Concatenate all row strings
            for row_str in &row_strings {
                self.buffer.push_str(row_str);
            }
        } else {
            // SEQUENTIAL PATH: Process rows sequentially for small tables
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
        }

        Ok(())
    }

    /// Write markdown-escaped text (escape | and convert \n to space) directly to buffer.
    ///
    /// **Performance**: Single-pass escaping without intermediate allocations.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn write_markdown_escaped(&mut self, text: &str) {
        Self::escape_markdown_to_buffer(&mut self.buffer, text);
    }

    /// Helper function to escape markdown to a string buffer.
    ///
    /// This is extracted as a separate function so it can be used in parallel contexts.
    ///
    /// **Performance**: Single-pass escaping without intermediate allocations.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn escape_markdown_to_buffer(buffer: &mut String, text: &str) {
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
                    buffer.push_str(&text[pos..]);
                }
                return;
            };

            // Write everything up to the special character
            if next_special > pos {
                buffer.push_str(&text[pos..next_special]);
            }

            // Write the escape sequence
            match bytes[next_special] {
                b'|' => buffer.push_str("\\|"),
                b'\n' => buffer.push(' '),
                _ => unreachable!(),
            }

            pos = next_special + 1;
        }
    }

    /// Write a table in HTML format with proper colspan and rowspan attributes.
    ///
    /// **Performance**: Uses efficient single-pass HTML escaping and minimizes allocations.
    ///
    /// **Merged Cells**: Properly handles merged cells by:
    /// - Adding `colspan` attributes for horizontal merges (gridSpan)
    /// - Adding `rowspan` attributes for vertical merges (vMerge)
    /// - Skipping cells that are covered by a merge
    ///
    /// **Styling**:
    /// - Styled tables (`styled = true`): Include indentation, line feeds, and CSS class
    /// - Minimal tables (`styled = false`): No indentation, no line feeds for compact output
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn write_html_table(&mut self, table: &Table, styled: bool) -> Result<()> {
        let rows = table.rows()?;

        // Pre-allocate buffer capacity to reduce reallocations
        // Estimate: ~100 bytes per cell on average
        let total_cells: usize = rows.iter().map(|r| r.cell_count().unwrap_or(0)).sum();
        self.buffer.reserve(total_cells * 100);

        // Analyze table to get span information (colspan/rowspan)
        let spans = analyze_table_spans(table)?;

        // Helper closure to write cell opening tag with span attributes
        let write_cell_open = |buffer: &mut String, tag: &str, span: &CellSpan| -> Result<()> {
            buffer.push('<');
            buffer.push_str(tag);

            if span.colspan > 1 {
                write!(buffer, " colspan=\"{}\"", span.colspan)
                    .map_err(|e| Error::Other(e.to_string()))?;
            }

            if span.rowspan > 1 {
                write!(buffer, " rowspan=\"{}\"", span.rowspan)
                    .map_err(|e| Error::Other(e.to_string()))?;
            }

            buffer.push('>');
            Ok(())
        };

        // Helper closure to write cell closing tag
        let write_cell_close = |buffer: &mut String, tag: &str| {
            buffer.push_str("</");
            buffer.push_str(tag);
            buffer.push('>');
        };

        // Helper closure to process row cells with optional indentation
        let process_row_cells = |writer: &mut Self,
                                 cells: &[Cell],
                                 row_idx: usize,
                                 tag: &str,
                                 spans: &[Vec<CellSpan>],
                                 cell_indent: Option<&str>|
         -> Result<()> {
            let mut grid_col = 0;

            for cell in cells {
                // Skip grid columns covered by merges
                while grid_col < spans.get(row_idx).map(|r| r.len()).unwrap_or(0)
                    && spans[row_idx][grid_col].skip
                {
                    grid_col += 1;
                }

                // Get span information for this cell
                let span_info = spans
                    .get(row_idx)
                    .and_then(|r| r.get(grid_col))
                    .copied()
                    .unwrap_or_else(CellSpan::new);

                // Skip this cell if it's covered by a merge
                if span_info.skip {
                    grid_col += 1;
                    continue;
                }

                let text = cell.text()?;

                // Write cell indent if provided
                if let Some(indent) = cell_indent {
                    writer.buffer.push_str(indent);
                }

                // Write opening tag with colspan/rowspan attributes
                write_cell_open(&mut writer.buffer, tag, &span_info)?;

                // HTML escape and write text
                writer.write_html_escaped(&text);

                // Write closing tag
                write_cell_close(&mut writer.buffer, tag);

                // Add line feed if indented
                if cell_indent.is_some() {
                    writer.buffer.push('\n');
                }

                // Move to next grid column
                grid_col += span_info.colspan;
            }

            Ok(())
        };

        if styled {
            // STYLED TABLE: With indentation, line feeds, and CSS class
            let indent = " ".repeat(self.options.html_table_indent);
            let double_indent = format!("{}{}", indent, indent);

            self.buffer.push_str("<table>\n");

            for (row_idx, row) in rows.iter().enumerate() {
                let tag = if row_idx == 0 { "th" } else { "td" };

                self.buffer.push_str(&indent);
                self.buffer.push_str("<tr>\n");

                let cells = row.cells()?;
                process_row_cells(self, &cells, row_idx, tag, &spans, Some(&double_indent))?;

                self.buffer.push_str(&indent);
                self.buffer.push_str("</tr>\n");
            }

            self.buffer.push_str("</table>");
        } else {
            // MINIMAL TABLE: No indentation, no line feeds for compact output
            self.buffer.push_str("<table>");

            for (row_idx, row) in rows.iter().enumerate() {
                let tag = if row_idx == 0 { "th" } else { "td" };

                self.buffer.push_str("<tr>");

                let cells = row.cells()?;
                process_row_cells(self, &cells, row_idx, tag, &spans, None)?;

                self.buffer.push_str("</tr>");
            }

            self.buffer.push_str("</table>");
        }

        Ok(())
    }

    /// Write HTML-escaped text directly to the buffer without intermediate allocations.
    ///
    /// **Performance**: Single-pass escaping that writes directly to the buffer,
    /// avoiding the 4 intermediate string allocations from chained `replace()` calls.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn write_html_escaped(&mut self, text: &str) {
        Self::escape_html_to_buffer(&mut self.buffer, text);
    }

    /// Helper function to escape HTML to a string buffer.
    ///
    /// This is extracted as a separate function so it can be used in parallel contexts.
    ///
    /// **Performance**: Single-pass escaping that writes directly to the buffer,
    /// avoiding the 4 intermediate string allocations from chained `replace()` calls.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
    fn escape_html_to_buffer(buffer: &mut String, text: &str) {
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
                    buffer.push_str(&text[pos..special_pos]);
                }

                // Write the escape sequence
                match bytes[special_pos] {
                    b'&' => buffer.push_str("&amp;"),
                    b'<' => buffer.push_str("&lt;"),
                    b'>' => buffer.push_str("&gt;"),
                    b'\n' => buffer.push_str("<br>"),
                    _ => unreachable!(),
                }

                pos = special_pos + 1;
            } else {
                // No more special characters, write rest and return
                if pos < bytes.len() {
                    buffer.push_str(&text[pos..]);
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
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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
    #[cfg(any(
        feature = "ole",
        feature = "ooxml",
        feature = "odf",
        feature = "rtf",
        feature = "iwa"
    ))]
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
