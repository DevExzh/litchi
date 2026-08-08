#[cfg(feature = "yaml")]
use crate::MetadataYaml;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use crate::document::{Cell, Table};
use crate::document::{Paragraph, Run};
/// Low-level writer for Markdown generation.
///
/// This module provides the `MarkdownWriter` struct which handles the actual
/// conversion of document elements to Markdown format.
///
/// **Note**: Some functionality requires the `doc` or `ooxml` feature to be enabled.
use litchi_core::{Error, Metadata, Result};
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use litchi_markdown::TableStyle;
use litchi_markdown::{MarkdownOptions, escape};
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use memchr::memchr;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use rayon::iter::{IndexedParallelIterator, IntoParallelRefIterator, ParallelIterator};

/// Minimum number of table rows to justify parallel processing overhead.
/// Tables are typically smaller than documents, so we use a lower threshold.
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
const TABLE_PARALLEL_THRESHOLD: usize = 20;

/// Information about a detected list item.
#[derive(Debug, Clone)]
pub(crate) struct ListItemInfo {
    /// The type of list
    list_type: ListType,
    /// The nesting level (0 = top level)
    level: usize,
    /// The marker text (e.g., "1.", "-", "*")
    marker: String,
    /// The content after the marker
    content: String,
    /// Byte offset at which content begins in the original paragraph text.
    content_start: usize,
    /// Whether owner metadata or literal text identified the list.
    origin: ListOrigin,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ListOrigin {
    Semantic,
    Text,
}

/// Types of lists supported.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ListType {
    /// Ordered list (numbered)
    Ordered,
    /// Unordered list (bulleted)
    Unordered,
}

impl ListItemInfo {
    pub(crate) fn bullet(level: usize) -> Self {
        Self {
            list_type: ListType::Unordered,
            level,
            marker: "-".to_owned(),
            content: String::new(),
            content_start: 0,
            origin: ListOrigin::Semantic,
        }
    }

    #[cfg(feature = "docx")]
    pub(crate) fn ordered(level: usize, marker: String) -> Self {
        Self {
            list_type: ListType::Ordered,
            level,
            marker,
            content: String::new(),
            content_start: 0,
            origin: ListOrigin::Semantic,
        }
    }
}

/// Information about cell span (colspan and rowspan) for HTML rendering.
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellSpan {
    /// Number of columns this cell spans (horizontal merge)
    colspan: usize,
    /// Number of rows this cell spans (vertical merge)
    rowspan: usize,
    /// Whether this cell should be skipped in rendering (it's covered by a merge)
    skip: bool,
}

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
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

/// Pre-extracted cell information for efficient table processing.
///
/// This struct caches cell span data to avoid repeated parsing during span analysis.
/// Text content is extracted separately for better performance.
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
#[derive(Debug, Clone)]
struct CellData {
    /// Horizontal span (gridSpan/colspan)
    grid_span: usize,
    /// Vertical merge state (OOXML only)
    #[cfg(feature = "docx")]
    v_merge: Option<crate::docx::VMergeState>,
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
/// **Performance**: Optimized to extract all cell data in a single pass, avoiding repeated
/// parsing. For large tables, uses parallel processing to extract cell data concurrently.
#[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
fn analyze_table_spans(table: &Table, use_parallel: bool) -> Result<Vec<Vec<CellSpan>>> {
    let rows = table.rows()?;
    if rows.is_empty() {
        return Ok(Vec::new());
    }

    // OPTIMIZATION: Pre-extract all cell data in a single pass to avoid repeated parsing.
    // This is the key optimization - we parse each cell exactly once.
    // For large tables, use parallel processing.
    let cell_data: Result<Vec<Vec<CellData>>> =
        if use_parallel && rows.len() > TABLE_PARALLEL_THRESHOLD {
            // PARALLEL PATH: Extract cell data in parallel for large tables
            // First collect all cells to avoid borrowing issues
            let all_cells: Result<Vec<Vec<Cell>>> = rows.iter().map(|row| row.cells()).collect();
            let all_cells = all_cells?;

            all_cells
                .par_iter()
                .map(|cells| {
                    cells
                        .iter()
                        .map(|cell| {
                            Ok(CellData {
                                grid_span: cell.grid_span().unwrap_or(1),
                                #[cfg(feature = "docx")]
                                v_merge: cell.v_merge()?,
                            })
                        })
                        .collect()
                })
                .collect()
        } else {
            // SEQUENTIAL PATH: Extract cell data sequentially for small tables
            rows.iter()
                .map(|row| {
                    let cells = row.cells()?;
                    cells
                        .iter()
                        .map(|cell| {
                            Ok(CellData {
                                grid_span: cell.grid_span().unwrap_or(1),
                                #[cfg(feature = "docx")]
                                v_merge: cell.v_merge()?,
                            })
                        })
                        .collect()
                })
                .collect()
        };
    let cell_data = cell_data?;

    // First pass: determine the maximum grid width (considering gridSpan)
    let mut max_grid_cols = 0;
    for row_cells in &cell_data {
        let row_grid_cols: usize = row_cells.iter().map(|c| c.grid_span).sum();
        max_grid_cols = max_grid_cols.max(row_grid_cols);
    }

    // Initialize span info for all cells
    let mut spans: Vec<Vec<CellSpan>> = vec![vec![CellSpan::new(); max_grid_cols]; rows.len()];

    // Second pass: analyze gridSpan and vMerge for each cell
    for (row_idx, row_cells) in cell_data.iter().enumerate() {
        let mut grid_col = 0; // Current grid column position

        for cell in row_cells {
            // Skip grid columns that are covered by previous cells' colspan
            while grid_col < max_grid_cols && spans[row_idx][grid_col].skip {
                grid_col += 1;
            }

            if grid_col >= max_grid_cols {
                break;
            }

            // Get horizontal span (gridSpan)
            let colspan = cell.grid_span;
            spans[row_idx][grid_col].colspan = colspan;

            // Mark columns covered by this cell's colspan as skipped
            for offset in 1..colspan {
                if grid_col + offset < max_grid_cols {
                    spans[row_idx][grid_col + offset] = CellSpan::skipped();
                }
            }

            // Get vertical merge state (vMerge)
            #[cfg(feature = "docx")]
            {
                use crate::docx::VMergeState;

                if let Some(v_merge_state) = &cell.v_merge {
                    match v_merge_state {
                        VMergeState::Restart => {
                            // This cell starts a vertical merge
                            // Count how many rows below continue this merge
                            let mut rowspan = 1;
                            for next_row_idx in (row_idx + 1)..cell_data.len() {
                                let Some(next_cell) = cell_data[next_row_idx].get(grid_col) else {
                                    break;
                                };
                                if !matches!(next_cell.v_merge, Some(VMergeState::Continue)) {
                                    break;
                                }

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

/// Extract all cell data from a table in a single optimized pass.
///
/// **Performance**: For large tables, uses parallel processing to extract cell data concurrently.
/// This avoids repeated XML parsing during table rendering.
#[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
fn extract_table_cell_data(table: &Table, use_parallel: bool) -> Result<Vec<Vec<String>>> {
    let rows = table.rows()?;
    if rows.is_empty() {
        return Ok(Vec::new());
    }

    // OPTIMIZATION: Extract all cell texts in a single pass
    // For large tables, use parallel processing
    if use_parallel && rows.len() > TABLE_PARALLEL_THRESHOLD {
        // First collect all cells to avoid borrowing issues with enum variants
        let all_cells: Result<Vec<Vec<Cell>>> = rows.iter().map(|row| row.cells()).collect();
        let all_cells = all_cells?;

        // Now extract texts in parallel
        all_cells
            .par_iter()
            .map(|cells| cells.iter().map(|cell| cell.text()).collect())
            .collect()
    } else {
        // Sequential extraction for small tables
        rows.iter()
            .map(|row| {
                let cells = row.cells()?;
                cells.iter().map(|cell| cell.text()).collect()
            })
            .collect()
    }
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

    #[cfg(feature = "docx")]
    fn heading_level_from_style_id(style_id: &str) -> Option<u8> {
        let mut characters = style_id
            .chars()
            .filter(|character| !matches!(character, ' ' | '-' | '_'));
        for expected in "heading".chars() {
            if !characters
                .next()
                .is_some_and(|actual| actual.eq_ignore_ascii_case(&expected))
            {
                return None;
            }
        }
        let level = characters
            .next()?
            .to_digit(10)
            .and_then(|value| u8::try_from(value).ok())?;
        (characters.next().is_none() && (1..=6).contains(&level)).then_some(level)
    }

    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    fn heading_level(&self, para: &Paragraph) -> Result<Option<u8>> {
        let level = match para {
            #[cfg(feature = "doc")]
            Paragraph::Doc(paragraph) => paragraph
                .properties()
                .outline_level
                .and_then(|level| level.checked_add(1)),
            #[cfg(feature = "docx")]
            Paragraph::Docx(paragraph) => paragraph
                .style_id()
                .map_err(crate::map_ooxml_error)?
                .as_deref()
                .and_then(Self::heading_level_from_style_id),
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(paragraph) => paragraph
                .properties
                .outline_level
                .and_then(|level| level.checked_add(1)),
            #[cfg(feature = "odt")]
            Paragraph::Odt(_) => None,
            #[cfg(feature = "pages")]
            Paragraph::Pages(_) => None,
        };
        Ok(level.filter(|level| (1..=6).contains(level)))
    }

    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    fn reject_unresolved_list(&self, para: &Paragraph) -> Result<()> {
        let unresolved = match para {
            #[cfg(feature = "doc")]
            Paragraph::Doc(paragraph) => {
                let properties = paragraph.properties();
                properties.list_format_override.is_some() && properties.list_level != Some(12)
            },
            #[cfg(feature = "docx")]
            Paragraph::Docx(paragraph) => paragraph
                .numbering()
                .map_err(crate::map_ooxml_error)?
                .is_some_and(|numbering| numbering.num_id != 0),
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(_) => false,
            #[cfg(feature = "odt")]
            Paragraph::Odt(_) => false,
            #[cfg(feature = "pages")]
            Paragraph::Pages(_) => false,
        };
        if unresolved {
            return Err(Error::Unsupported(
                "Markdown export requires resolved numbering definitions to distinguish bullets, ordered lists, and unsupported formats"
                    .to_owned(),
            ));
        }
        Ok(())
    }

    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    fn permits_text_list_fallback(para: &Paragraph) -> bool {
        match para {
            #[cfg(feature = "doc")]
            Paragraph::Doc(_) => false,
            #[cfg(feature = "docx")]
            Paragraph::Docx(_) => false,
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(_) => true,
            #[cfg(feature = "odt")]
            Paragraph::Odt(_) => true,
            #[cfg(feature = "pages")]
            Paragraph::Pages(_) => true,
        }
    }

    fn write_literal(&mut self, text: &str) {
        self.buffer.push_str(&escape::text(text));
    }

    fn write_html_literal(&mut self, text: &str) {
        let mut start = 0usize;
        for (index, character) in text.char_indices() {
            match character {
                '&' | '<' | '>' => {
                    self.write_literal(&text[start..index]);
                    self.buffer.push_str(match character {
                        '&' => "&amp;",
                        '<' => "&lt;",
                        '>' => "&gt;",
                        _ => "",
                    });
                    start = index + character.len_utf8();
                },
                _ => {},
            }
        }
        self.write_literal(&text[start..]);
    }

    fn write_list_prefix(&mut self, list_info: &ListItemInfo) -> Result<()> {
        let spaces = list_info
            .level
            .checked_mul(self.options.list_indent)
            .ok_or_else(|| {
                Error::Unsupported("Markdown list indentation exceeds limits".to_owned())
            })?;
        let marker = match list_info.list_type {
            ListType::Unordered => "-".to_owned(),
            ListType::Ordered if list_info.marker.ends_with('.') => list_info.marker.clone(),
            ListType::Ordered => format!(
                "{}.",
                list_info
                    .marker
                    .trim_start_matches('(')
                    .trim_end_matches(')')
            ),
        };
        self.buffer
            .try_reserve(spaces.saturating_add(marker.len()).saturating_add(1))
            .map_err(|source| Error::Allocation {
                resource: "Markdown list output",
                source,
            })?;
        self.buffer.extend(std::iter::repeat_n(' ', spaces));
        self.buffer.push_str(&marker);
        self.buffer.push(' ');
        Ok(())
    }

    /// Write a paragraph to the buffer.
    ///
    /// **Note**: This method requires the `doc` or `ooxml` feature to be enabled.
    ///
    /// **Performance**: Optimized to avoid redundant XML parsing by extracting runs
    /// once and deriving text from them when needed.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    #[allow(irrefutable_let_patterns)]
    pub fn write_paragraph(&mut self, para: &Paragraph) -> Result<()> {
        self.write_paragraph_with_list(para, None)
    }

    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    #[allow(irrefutable_let_patterns)]
    pub(crate) fn write_paragraph_with_list(
        &mut self,
        para: &Paragraph,
        resolved_list: Option<&ListItemInfo>,
    ) -> Result<()> {
        if resolved_list.is_none() {
            self.reject_unresolved_list(para)?;
        }
        if let Some(level) = self.heading_level(para)? {
            self.buffer
                .extend(std::iter::repeat_n('#', usize::from(level)));
            self.buffer.push(' ');
        }

        // First check for paragraph-level formulas (display math)
        #[cfg(feature = "docx")]
        {
            use crate::document::Paragraph;
            if let Paragraph::Docx(docx_para) = para {
                let display_formulas = docx_para
                    .paragraph_level_formulas()
                    .map_err(crate::map_ooxml_error)?;
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
                    let list_info = resolved_list.cloned().or_else(|| {
                        Self::permits_text_list_fallback(para)
                            .then(|| self.detect_list_item(&text))
                            .flatten()
                    });
                    if let Some(list_info) = list_info {
                        self.write_list_prefix(&list_info)?;
                        if list_info.origin == ListOrigin::Semantic {
                            self.write_literal(&text);
                        } else {
                            self.write_literal(&list_info.content);
                        }
                    } else {
                        self.write_literal(&text);
                    }
                }
            } else {
                // Has runs - process them normally
                // Derive text from runs for list detection (cheaper than parsing XML again)
                let text = self.extract_text_from_runs(&runs)?;

                let list_info = resolved_list.cloned().or_else(|| {
                    Self::permits_text_list_fallback(para)
                        .then(|| self.detect_list_item(&text))
                        .flatten()
                });
                if let Some(list_info) = list_info {
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

            let list_info = resolved_list.cloned().or_else(|| {
                Self::permits_text_list_fallback(para)
                    .then(|| self.detect_list_item(&text))
                    .flatten()
            });
            if let Some(list_info) = list_info {
                self.write_list_prefix(&list_info)?;
                if list_info.origin == ListOrigin::Semantic {
                    self.write_literal(&text);
                } else {
                    self.write_literal(&list_info.content);
                }
            } else {
                self.write_literal(&text);
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
    #[cfg(all(feature = "docx", feature = "formula"))]
    fn write_paragraph_with_display_formulas(
        &mut self,
        para: &Paragraph,
        display_formulas: Vec<String>,
    ) -> Result<()> {
        use litchi_formula::omml_to_latex;

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
            let latex = omml_to_latex(&omml_xml)
                .map_err(|error| Error::ParseError(format!("OMML conversion failed: {error}")))?;

            // Display formulas use display style (false = display mode)
            let formula_md = self.format_formula(&latex, false);
            self.buffer.push_str(&formula_md);
            self.buffer.push('\n');
        }

        Ok(())
    }

    /// Fallback for when formula feature is not enabled.
    #[cfg(all(feature = "docx", not(feature = "formula")))]
    fn write_paragraph_with_display_formulas(
        &mut self,
        para: &Paragraph,
        display_formulas: Vec<String>,
    ) -> Result<()> {
        if !display_formulas.is_empty() {
            return Err(Error::FeatureDisabled("formula".to_owned()));
        }
        for run in para.runs()? {
            self.write_run(&run)?;
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
        if bold == self.current_bold
            && italic == self.current_italic
            && strikethrough == self.current_strikethrough
        {
            return;
        }

        self.close_formatting();
        if bold {
            self.buffer.push_str("**");
            self.current_bold = true;
        }
        if italic {
            self.buffer.push('*');
            self.current_italic = true;
        }
        if strikethrough {
            self.buffer.push_str("~~");
            self.current_strikethrough = true;
        }
    }

    /// Write a run with formatting.
    ///
    /// **Note**: This method requires the `doc` or `ooxml` feature to be enabled.
    ///
    /// **Performance**: For OOXML runs, this uses a single XML parse to extract both
    /// text and properties simultaneously, providing 2x speedup over separate calls.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    #[allow(irrefutable_let_patterns)]
    pub fn write_run(&mut self, run: &Run) -> Result<()> {
        // First check if this run contains a formula
        if let Some(formula_markdown) = self.extract_formula_from_run(run)? {
            self.close_formatting();
            self.buffer.push_str(&formula_markdown);
            return Ok(());
        }

        // OPTIMIZATION: Get text AND properties in a single XML parse
        // This is 2x faster than calling text() then get_properties()
        #[cfg(feature = "docx")]
        let (text, bold, italic, strikethrough, vertical_pos) =
            if let crate::document::Run::Docx(docx_run) = run {
                let (text, props) = docx_run
                    .get_text_and_properties()
                    .map_err(crate::map_ooxml_error)?;
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

        #[cfg(all(feature = "doc", not(feature = "docx")))]
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

        // For rtf, odf features (without doc/ooxml)
        #[cfg(all(
            not(any(feature = "doc", feature = "docx")),
            any(feature = "rtf", feature = "odt", feature = "pages")
        ))]
        let (text, bold, italic, strikethrough) = {
            let text = run.text()?;
            if text.is_empty() {
                return Ok(());
            }
            (
                text,
                run.bold()?.unwrap_or(false),
                run.italic()?.unwrap_or(false),
                run.strikethrough()?.unwrap_or(false),
            )
        };

        // Handle vertical position (superscript/subscript)
        // Note: vertical_position() is available when doc or ooxml features are enabled
        #[cfg(any(feature = "doc", feature = "docx"))]
        {
            use litchi_core::VerticalPosition;

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
                self.close_formatting();
                match self.options.script_style {
                    litchi_markdown::ScriptStyle::Html => match pos {
                        VerticalPosition::Superscript => {
                            self.buffer.push_str("<sup>");
                            self.write_html_literal(&text);
                            self.buffer.push_str("</sup>");
                        },
                        VerticalPosition::Subscript => {
                            self.buffer.push_str("<sub>");
                            self.write_html_literal(&text);
                            self.buffer.push_str("</sub>");
                        },
                        VerticalPosition::Normal => {
                            self.write_literal(&text);
                        },
                    },
                    litchi_markdown::ScriptStyle::Unicode => {
                        // Convert to Unicode superscript/subscript characters
                        // Fall back to HTML tags for characters without Unicode equivalents
                        match pos {
                            VerticalPosition::Superscript => {
                                if litchi_markdown::unicode::can_convert_to_superscript(&text) {
                                    // All characters can be converted to superscript
                                    let converted =
                                        litchi_markdown::unicode::convert_to_superscript(&text);
                                    self.buffer.push_str(&converted);
                                } else {
                                    // Fall back to HTML for partial support
                                    self.buffer.push_str("<sup>");
                                    self.write_html_literal(&text);
                                    self.buffer.push_str("</sup>");
                                }
                            },
                            VerticalPosition::Subscript => {
                                if litchi_markdown::unicode::can_convert_to_subscript(&text) {
                                    // All characters can be converted to subscript
                                    let converted =
                                        litchi_markdown::unicode::convert_to_subscript(&text);
                                    self.buffer.push_str(&converted);
                                } else {
                                    // Fall back to HTML for partial support
                                    self.buffer.push_str("<sub>");
                                    self.write_html_literal(&text);
                                    self.buffer.push_str("</sub>");
                                }
                            },
                            VerticalPosition::Normal => {
                                self.write_literal(&text);
                            },
                        }
                    },
                }
                return Ok(());
            }
        }

        // Pre-calculate buffer size for non-vertical-position formatting
        #[cfg(not(any(feature = "doc", feature = "docx")))]
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
        if self.options.strikethrough_style == litchi_markdown::StrikethroughStyle::Html
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
                    self.write_literal(&text);
                    self.buffer.push_str("***");
                },
                (true, false) => {
                    self.buffer.push_str("**");
                    self.write_literal(&text);
                    self.buffer.push_str("**");
                },
                (false, true) => {
                    self.buffer.push('*');
                    self.write_literal(&text);
                    self.buffer.push('*');
                },
                (false, false) => {
                    self.write_html_literal(&text);
                },
            }
            self.buffer.push_str("</del>");
        } else {
            // Markdown-style formatting: can span across runs
            self.apply_formatting(bold, italic, strikethrough);
            self.write_literal(&text);
        }

        Ok(())
    }

    /// Write a table to the buffer.
    ///
    /// **Note**: This method requires the `doc` or `ooxml` feature to be enabled.
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
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
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
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
                #[cfg(feature = "docx")]
                {
                    if cell.v_merge()?.is_some() {
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
    /// Pre-extracts all cell data in a single optimized pass to avoid repeated parsing.
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
    fn write_markdown_table(&mut self, table: &Table) -> Result<()> {
        // OPTIMIZATION: Extract all cell data in a single pass (with parallelization for large tables)
        let cell_data = extract_table_cell_data(table, self.options.use_parallel)?;
        if cell_data.is_empty() {
            return Ok(());
        }

        // Pre-allocate buffer capacity
        let total_cells: usize = cell_data.iter().map(|row| row.len()).sum();
        self.buffer.reserve(total_cells * 50); // Estimate: ~50 bytes per cell

        // Write header row (first row) - always sequential
        let cell_count = cell_data[0].len();

        self.buffer.push('|');
        for text in &cell_data[0] {
            self.buffer.push(' ');
            // Escape pipe and newline in a single pass
            self.write_markdown_escaped(text);
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
        if self.options.use_parallel && cell_data.len() > TABLE_PARALLEL_THRESHOLD {
            // PARALLEL PATH: Process rows in parallel for large tables
            // Cell data is already extracted, now just format in parallel
            let row_strings: Vec<String> = cell_data[1..]
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

            // Concatenate all row strings efficiently
            let total_len: usize = row_strings.iter().map(|s| s.len()).sum();
            self.buffer.reserve(total_len);
            for row_str in &row_strings {
                self.buffer.push_str(row_str);
            }
        } else {
            // SEQUENTIAL PATH: Process rows sequentially for small tables
            for row_texts in &cell_data[1..] {
                self.buffer.push('|');
                for text in row_texts {
                    self.buffer.push(' ');
                    self.write_markdown_escaped(text);
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
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
    fn write_markdown_escaped(&mut self, text: &str) {
        Self::escape_markdown_to_buffer(&mut self.buffer, text);
    }

    /// Helper function to escape markdown to a string buffer.
    ///
    /// This is extracted as a separate function so it can be used in parallel contexts.
    ///
    /// **Performance**: Single-pass escaping without intermediate allocations.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
    fn escape_markdown_to_buffer(buffer: &mut String, text: &str) {
        let escaped = escape::text(text);
        let mut characters = escaped.chars().peekable();
        while let Some(character) = characters.next() {
            match character {
                '\r' => {
                    if characters.peek() == Some(&'\n') {
                        characters.next();
                    }
                    buffer.push(' ');
                },
                '\n' => buffer.push(' '),
                other => buffer.push(other),
            }
        }
    }

    /// Write a table in HTML format with proper colspan and rowspan attributes.
    ///
    /// **Performance**: Uses efficient single-pass HTML escaping and minimizes allocations.
    /// For large tables, uses parallel processing to render rows concurrently.
    /// Pre-extracts all cell data in a single optimized pass to avoid repeated parsing.
    ///
    /// **Merged Cells**: Properly handles merged cells by:
    /// - Adding `colspan` attributes for horizontal merges (gridSpan)
    /// - Adding `rowspan` attributes for vertical merges (vMerge)
    /// - Skipping cells that are covered by a merge
    ///
    /// **Styling**:
    /// - Styled tables (`styled = true`): Include indentation, line feeds, and CSS class
    /// - Minimal tables (`styled = false`): No indentation, no line feeds for compact output
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
    fn write_html_table(&mut self, table: &Table, styled: bool) -> Result<()> {
        // OPTIMIZATION: Extract all cell data in a single pass (with parallelization for large tables)
        let cell_data = extract_table_cell_data(table, self.options.use_parallel)?;
        if cell_data.is_empty() {
            return Ok(());
        }

        // Pre-allocate buffer capacity to reduce reallocations
        // Estimate: ~100 bytes per cell on average
        let total_cells: usize = cell_data.iter().map(|row| row.len()).sum();
        self.buffer.reserve(total_cells * 100);

        // Analyze table to get span information (colspan/rowspan)
        // Use the same parallel setting as for cell extraction
        let spans = analyze_table_spans(table, self.options.use_parallel)?;

        // Helper to format a single cell
        let format_cell =
            |text: &str, tag: &str, span: &CellSpan, cell_indent: Option<&str>| -> String {
                let mut cell_buffer = String::with_capacity(text.len() + 50);

                // Write cell indent if provided
                if let Some(indent) = cell_indent {
                    cell_buffer.push_str(indent);
                }

                // Write opening tag with colspan/rowspan attributes
                cell_buffer.push('<');
                cell_buffer.push_str(tag);

                if span.colspan > 1 {
                    use std::fmt::Write;
                    let _ = write!(cell_buffer, " colspan=\"{}\"", span.colspan);
                }

                if span.rowspan > 1 {
                    use std::fmt::Write;
                    let _ = write!(cell_buffer, " rowspan=\"{}\"", span.rowspan);
                }

                cell_buffer.push('>');

                // HTML escape and write text
                Self::escape_html_to_buffer(&mut cell_buffer, text);

                // Write closing tag
                cell_buffer.push_str("</");
                cell_buffer.push_str(tag);
                cell_buffer.push('>');

                // Add line feed if indented
                if cell_indent.is_some() {
                    cell_buffer.push('\n');
                }

                cell_buffer
            };

        // Helper to format an entire row
        let format_row = |row_texts: &[String],
                          row_idx: usize,
                          tag: &str,
                          spans: &[Vec<CellSpan>],
                          cell_indent: Option<&str>|
         -> String {
            let mut row_buffer = String::with_capacity(row_texts.len() * 100);
            let mut grid_col = 0;
            let mut text_idx = 0;

            while text_idx < row_texts.len() {
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

                // Format this cell
                let cell_html = format_cell(&row_texts[text_idx], tag, &span_info, cell_indent);
                row_buffer.push_str(&cell_html);

                // Move to next grid column and text index
                grid_col += span_info.colspan;
                text_idx += 1;
            }

            row_buffer
        };

        if styled {
            // STYLED TABLE: With indentation, line feeds, and CSS class
            let indent = " ".repeat(self.options.html_table_indent);
            let double_indent = format!("{}{}", indent, indent);

            self.buffer.push_str("<table>\n");

            // Use parallel processing for large tables
            if self.options.use_parallel && cell_data.len() > TABLE_PARALLEL_THRESHOLD {
                // PARALLEL PATH: Format rows in parallel
                let row_htmls: Vec<String> = cell_data
                    .par_iter()
                    .enumerate()
                    .map(|(row_idx, row_texts)| {
                        let tag = if row_idx == 0 { "th" } else { "td" };
                        let mut row_html = String::with_capacity(row_texts.len() * 100 + 100);
                        row_html.push_str(&indent);
                        row_html.push_str("<tr>\n");
                        row_html.push_str(&format_row(
                            row_texts,
                            row_idx,
                            tag,
                            &spans,
                            Some(&double_indent),
                        ));
                        row_html.push_str(&indent);
                        row_html.push_str("</tr>\n");
                        row_html
                    })
                    .collect();

                // Concatenate all row HTMLs efficiently
                let total_len: usize = row_htmls.iter().map(|s| s.len()).sum();
                self.buffer.reserve(total_len);
                for row_html in &row_htmls {
                    self.buffer.push_str(row_html);
                }
            } else {
                // SEQUENTIAL PATH: Format rows sequentially
                for (row_idx, row_texts) in cell_data.iter().enumerate() {
                    let tag = if row_idx == 0 { "th" } else { "td" };

                    self.buffer.push_str(&indent);
                    self.buffer.push_str("<tr>\n");
                    self.buffer.push_str(&format_row(
                        row_texts,
                        row_idx,
                        tag,
                        &spans,
                        Some(&double_indent),
                    ));
                    self.buffer.push_str(&indent);
                    self.buffer.push_str("</tr>\n");
                }
            }

            self.buffer.push_str("</table>");
        } else {
            // MINIMAL TABLE: No indentation, no line feeds for compact output
            self.buffer.push_str("<table>");

            // Use parallel processing for large tables
            if self.options.use_parallel && cell_data.len() > TABLE_PARALLEL_THRESHOLD {
                // PARALLEL PATH: Format rows in parallel
                let row_htmls: Vec<String> = cell_data
                    .par_iter()
                    .enumerate()
                    .map(|(row_idx, row_texts)| {
                        let tag = if row_idx == 0 { "th" } else { "td" };
                        let mut row_html = String::with_capacity(row_texts.len() * 100 + 20);
                        row_html.push_str("<tr>");
                        row_html.push_str(&format_row(row_texts, row_idx, tag, &spans, None));
                        row_html.push_str("</tr>");
                        row_html
                    })
                    .collect();

                // Concatenate all row HTMLs efficiently
                let total_len: usize = row_htmls.iter().map(|s| s.len()).sum();
                self.buffer.reserve(total_len);
                for row_html in &row_htmls {
                    self.buffer.push_str(row_html);
                }
            } else {
                // SEQUENTIAL PATH: Format rows sequentially
                for (row_idx, row_texts) in cell_data.iter().enumerate() {
                    let tag = if row_idx == 0 { "th" } else { "td" };

                    self.buffer.push_str("<tr>");
                    self.buffer
                        .push_str(&format_row(row_texts, row_idx, tag, &spans, None));
                    self.buffer.push_str("</tr>");
                }
            }

            self.buffer.push_str("</table>");
        }

        Ok(())
    }

    /// Helper function to escape HTML to a string buffer.
    ///
    /// This is extracted as a separate function so it can be used in parallel contexts.
    ///
    /// **Performance**: Single-pass escaping that writes directly to the buffer,
    /// avoiding the 4 intermediate string allocations from chained `replace()` calls.
    /// Uses SIMD-accelerated memchr for fast searching.
    #[cfg(any(feature = "doc", feature = "docx", feature = "odt", feature = "rtf"))]
    fn escape_html_to_buffer(buffer: &mut String, text: &str) {
        let bytes = text.as_bytes();
        let mut pos = 0;

        while pos < bytes.len() {
            // Find the next character that needs escaping
            let next_special = b"&<>\n"
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
                    _ => return,
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

    /// Append a single character to the buffer.
    #[allow(dead_code)]
    pub fn push(&mut self, ch: char) {
        self.buffer.push(ch);
    }

    /// Write a formatted string to the buffer.
    #[allow(dead_code)]
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
    #[cfg(feature = "yaml")]
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

    /// Refuse metadata output when its serializer is not compiled in.
    #[cfg(not(feature = "yaml"))]
    pub fn write_metadata(&mut self, _metadata: &Metadata) -> Result<()> {
        Err(Error::FeatureDisabled("yaml".to_owned()))
    }

    /// Detect if a paragraph is a list item and extract list information.
    fn detect_list_item(&self, text: &str) -> Option<ListItemInfo> {
        let trimmed = text.trim_start();

        // Check for ordered lists: 1. 2. 3. or 1) 2) 3) or (1) (2) (3)
        if let Some(captures) = self.extract_ordered_list_marker(trimmed) {
            let marker = captures.0;
            let content = captures.1;
            let level = self.calculate_indent_level(text);
            return Some(ListItemInfo {
                list_type: ListType::Ordered,
                level,
                marker: marker.to_string(),
                content: content.to_string(),
                content_start: text.len().saturating_sub(content.len()),
                origin: ListOrigin::Text,
            });
        }

        // Check for unordered lists: - * •
        if let Some(captures) = self.extract_unordered_list_marker(trimmed) {
            let marker = captures.0;
            let content = captures.1;
            let level = self.calculate_indent_level(text);
            return Some(ListItemInfo {
                list_type: ListType::Unordered,
                level,
                marker: marker.to_string(),
                content: content.to_string(),
                content_start: text.len().saturating_sub(content.len()),
                origin: ListOrigin::Text,
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
        if self.options.list_indent == 0 {
            0
        } else {
            leading / self.options.list_indent
        }
    }

    /// Extract formula content from a run and convert to markdown.
    ///
    /// Returns the markdown representation of the formula if one is found, None otherwise.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    #[allow(irrefutable_let_patterns)]
    fn extract_formula_from_run(&self, _run: &Run) -> Result<Option<String>> {
        // Try OOXML OMML formulas first
        #[cfg(feature = "docx")]
        if let crate::document::Run::Docx(docx_run) = _run
            && let Some(omml_xml) = docx_run.omml_formula().map_err(crate::map_ooxml_error)?
        {
            // Parse OMML and convert to LaTeX
            #[cfg(feature = "formula")]
            {
                let latex = litchi_formula::omml_to_latex(&omml_xml).map_err(|error| {
                    Error::ParseError(format!("OMML conversion failed: {error}"))
                })?;
                return Ok(Some(self.format_formula(&latex, true))); // true = inline
            }

            #[cfg(not(feature = "formula"))]
            {
                let _ = omml_xml;
                return Err(Error::FeatureDisabled("formula".to_owned()));
            }
        }

        // Try OLE MTEF formulas
        #[cfg(feature = "doc")]
        {
            // `Run` carries one variant per enabled format feature, so this
            // pattern is refutable in most builds and irrefutable in an
            // `doc`-only build; `if let` covers both without a wildcard arm that
            // would be unreachable in the latter.
            #[allow(irrefutable_let_patterns)]
            if let crate::document::Run::Doc(ole_run) = _run
                && ole_run.has_mtef_formula()
            {
                let latex = ole_run.mtef_formula_latex().ok_or_else(|| {
                    Error::Unsupported(
                        "MTEF formula has no semantic LaTeX representation".to_owned(),
                    )
                })?;
                return Ok(Some(self.format_formula(latex, true)));
            }
        }

        Ok(None)
    }

    /// Format a formula with the appropriate delimiters.
    ///
    /// # Arguments
    /// * `formula` - The formula content (LaTeX)
    /// * `inline` - Whether this is an inline formula (true) or display formula (false)
    fn format_formula(&self, formula: &str, inline: bool) -> String {
        if inline {
            match self.options.formula_style {
                litchi_markdown::FormulaStyle::LaTeX => format!("\\({}\\)", formula),
                litchi_markdown::FormulaStyle::Dollar => format!("${}$", formula),
            }
        } else {
            match self.options.formula_style {
                litchi_markdown::FormulaStyle::LaTeX => format!("\\[{}\\]", formula),
                litchi_markdown::FormulaStyle::Dollar => format!("$${}$$", formula),
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
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    fn write_list_item(&mut self, _para: &Paragraph, list_info: &ListItemInfo) -> Result<()> {
        self.write_list_prefix(list_info)?;

        // Write the content with styles if enabled
        if self.options.include_styles && !list_info.content.trim().is_empty() {
            // For styled content, we need to skip the marker part and write the remaining runs
            // This is a simplified approach - in practice, we'd need more sophisticated
            // parsing to handle cases where the marker spans multiple runs
            self.write_literal(&list_info.content);
        } else {
            // Write the content directly
            self.write_literal(&list_info.content);
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
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
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
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "rtf",
        feature = "pages"
    ))]
    fn write_list_item_from_runs(&mut self, runs: &[Run], list_info: &ListItemInfo) -> Result<()> {
        self.write_list_prefix(list_info)?;

        // Write runs, skipping the list marker portion
        // This is a simplified approach - we write all runs with their formatting
        // A more sophisticated implementation would skip the marker text in the first run
        let mut accumulated_len = 0;
        let marker_end_pos = list_info.content_start;

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
                let text_after_marker = run_text.get(skip_chars..).ok_or_else(|| {
                    Error::ParseError("list marker splits a UTF-8 code point".to_owned())
                })?;

                // Create a temporary run-like structure with the remaining text
                // For now, just write the text - ideally we'd preserve formatting
                self.write_literal(text_after_marker);
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

#[cfg(all(test, feature = "docx"))]
mod tests {
    use super::{ListItemInfo, ListOrigin, ListType, MarkdownWriter};
    use crate::document::Paragraph;
    use litchi_markdown::{MarkdownOptions, ToMarkdown};

    fn paragraph(properties: &str, text: &str) -> Paragraph {
        Paragraph::Docx(crate::docx::Paragraph::new(
            format!(
                r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
            )
            .into_bytes(),
        ))
    }

    #[test]
    fn golden_heading_uses_style_and_escapes_literal_markdown() -> litchi_core::Result<()> {
        let paragraph = paragraph(r#"<w:pStyle w:val="Heading3"/>"#, "A *literal* [title]");
        assert_eq!(paragraph.to_markdown()?, "### A \\*literal\\* \\[title\\]");
        Ok(())
    }

    #[test]
    fn literal_number_prefix_is_not_guessed_as_docx_numbering() -> litchi_core::Result<()> {
        let literal = paragraph("", "1. not a semantic list");
        assert_eq!(literal.to_markdown()?, "1\\. not a semantic list");
        Ok(())
    }

    fn assert_unresolved_numbering_refused(num_id: u32) {
        let properties =
            format!(r#"<w:numPr><w:ilvl w:val="2"/><w:numId w:val="{num_id}"/></w:numPr>"#);
        let paragraph = paragraph(&properties, "item");
        assert!(matches!(
            paragraph.to_markdown(),
            Err(litchi_core::Error::Unsupported(_))
        ));
    }

    #[test]
    fn bullet_numbering_without_package_definitions_is_refused() {
        assert_unresolved_numbering_refused(7);
    }

    #[test]
    fn ordered_numbering_without_package_definitions_is_refused() {
        assert_unresolved_numbering_refused(8);
    }

    #[test]
    fn unknown_numbering_without_package_definitions_is_refused() {
        assert_unresolved_numbering_refused(4_294_967_295);
    }

    #[test]
    fn zero_list_indent_is_safe_for_textual_fallback() -> litchi_core::Result<()> {
        let mut writer = MarkdownWriter::new(MarkdownOptions::new().with_list_indent(0));
        let item = ListItemInfo {
            list_type: ListType::Unordered,
            level: 4,
            marker: "•".to_owned(),
            content: "item".to_owned(),
            content_start: 0,
            origin: ListOrigin::Text,
        };
        writer.write_list_prefix(&item)?;
        writer.write_literal(&item.content);
        assert_eq!(writer.finish(), "- item");
        Ok(())
    }

    #[test]
    fn textual_fallback_tracks_multibyte_prefix_boundaries() {
        let writer = MarkdownWriter::new(MarkdownOptions::new());
        let Some(info) = writer.detect_list_item("\u{2003}• text") else {
            panic!("textual bullet should be recognized");
        };
        assert_eq!(info.content, "text");
        assert_eq!("\u{2003}• text".get(info.content_start..), Some("text"));
        assert!("\u{2003}• text".is_char_boundary(info.content_start));
    }

    #[cfg(feature = "formula")]
    #[test]
    fn formula_conversion_failure_is_propagated() {
        let paragraph = Paragraph::Docx(crate::docx::Paragraph::new(
            br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath/></w:p>"#.to_vec(),
        ));
        assert!(paragraph.to_markdown().is_err());
    }

    #[test]
    fn html_script_text_is_entity_and_markdown_escaped() {
        let mut writer = MarkdownWriter::new(MarkdownOptions::new());
        writer.write_html_literal("<&> *");
        assert_eq!(writer.finish(), "&lt;&amp;&gt; \\*");
    }

    #[test]
    fn markdown_table_text_escapes_ampersands_and_normalizes_crlf() {
        let mut output = String::new();
        MarkdownWriter::escape_markdown_to_buffer(&mut output, "&copy;\r\nnext");
        assert_eq!(output, "\\&copy; next");
    }
}
