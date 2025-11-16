//! DOC file writer implementation
//!
//! This module provides functionality to create and modify Microsoft Word documents
//! in the legacy binary format (.doc files) using OLE2 structured storage.
//!
//! # Architecture
//!
//! The writer generates the complex DOC file structure including:
//! - FIB (File Information Block) - contains file metadata and stream offsets
//! - Text stream - contains the actual document text
//! - Table stream (0Table/1Table) - contains formatting and structure
//! - Data stream - contains embedded objects
//!
//! # DOC File Format Overview
//!
//! DOC files use a "piece table" architecture where:
//! 1. Text is stored in one or more pieces (continuous runs)
//! 2. Character formatting (CHP) is stored separately
//! 3. Paragraph formatting (PAP) is stored separately
//! 4. All formatting uses SPRMs (Single Property Modifiers)
//!
//! # Critical Implementation Details
//!
//! ## Stream Creation Order
//!
//! Microsoft Word requires `WordDocument` to be allocated at **sector 0** of the
//! OLE file. This is achieved by creating the `WordDocument` stream BEFORE any
//! other streams. The stream creation order in `save()` method is:
//!
//! 1. `WordDocument` → sector 0 (REQUIRED by Microsoft Word)
//! 2. `1Table` → next available sector
//!
//! ## Directory Entry Ordering
//!
//! Directory entries are sorted using Apache POI's PropertyComparator rules:
//! - Sort by name length first (shorter names before longer names)
//! - Then alphabetically (case-insensitive) for same-length names
//!
//! For DOC files, this results in the tree structure:
//! ```text
//! Root Entry
//!     └─ WordDocument (midpoint of sorted list)
//!          └─ 1Table (left child, shorter name)
//! ```
//!
//! **Note**: Stream ALLOCATION order (sector assignment) is DIFFERENT from
//! directory ENTRY order (tree structure). See `OleWriter` documentation for details.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ole::doc::DocWriter;
//!
//! let mut writer = DocWriter::new();
//!
//! // Add paragraphs
//! writer.add_paragraph("Hello, World!")?;
//! writer.add_paragraph("This is a second paragraph.")?;
//!
//! // Save the document
//! writer.save("output.doc")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Compatibility
//!
//! Generated DOC files are compatible with:
//! - Microsoft Word 97-2003
//! - Microsoft Word 2007+ (compatibility mode)
//! - LibreOffice Writer
//! - Apache POI (HWPF)
//! - Other OLE2-based Word readers

use super::fib::FibBuilder;
use super::font_table::FontTableBuilder;
use super::piece_table::{Piece, PieceTableBuilder};
use crate::ole::sprm_operations::*;
use crate::ole::writer::OleWriter;
use std::collections::HashMap;

/// Error type for DOC writing
#[derive(Debug)]
pub enum DocWriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(crate::ole::OleError),
}

impl From<std::io::Error> for DocWriteError {
    fn from(err: std::io::Error) -> Self {
        DocWriteError::Io(err)
    }
}

impl From<crate::ole::OleError> for DocWriteError {
    fn from(err: crate::ole::OleError) -> Self {
        DocWriteError::Ole(err)
    }
}

impl std::fmt::Display for DocWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DocWriteError::Io(e) => write!(f, "I/O error: {}", e),
            DocWriteError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            DocWriteError::Ole(e) => write!(f, "OLE error: {}", e),
        }
    }
}

impl std::error::Error for DocWriteError {}

/// Character formatting properties
#[derive(Debug, Clone, Default)]
pub struct CharacterFormatting {
    /// Bold
    pub bold: Option<bool>,
    /// Italic
    pub italic: Option<bool>,
    /// Underline
    pub underline: Option<bool>,
    /// Strikethrough
    pub strike: Option<bool>,
    /// Double strikethrough
    pub double_strike: Option<bool>,
    /// Superscript
    pub superscript: Option<bool>,
    /// Subscript
    pub subscript: Option<bool>,
    /// Small caps
    pub small_caps: Option<bool>,
    /// All caps
    pub all_caps: Option<bool>,
    /// Hidden text
    pub hidden: Option<bool>,
    /// Special character flag (fSpec). Required for field begin/separator/end and other control chars.
    pub special: Option<bool>,
    /// Field vanish flag. Used to hide field instruction text per Word conventions.
    pub field_vanish: Option<bool>,
    /// Font size (in half-points, e.g., 24 = 12pt)
    pub font_size: Option<u16>,
    /// Font name
    pub font_name: Option<String>,
    /// Text color as (R,G,B)
    pub color: Option<(u8, u8, u8)>,
    // Future enhancement: Additional properties (color, strikethrough, subscript, superscript, etc.)
}

/// Line spacing descriptor for paragraphs, equivalent to POI's LineSpacingDescriptor (LSPD).
#[derive(Debug, Clone, Copy, Default)]
pub struct LineSpacing {
    /// Line height. If `is_multiple` is false, value is in twips. If true, value is in 240ths of a line.
    pub dya_line: i16,
    /// Whether `dya_line` is a multiple of single line (value is 240ths of a line) instead of twips.
    pub is_multiple: bool,
}

/// Paragraph formatting properties
#[derive(Debug, Clone, Default)]
pub struct ParagraphFormatting {
    /// Alignment (0=left, 1=center, 2=right, 3=justify)
    pub alignment: Option<u8>,
    /// Left indent (in twips, 1440 twips = 1 inch)
    pub left_indent: Option<i32>,
    /// Right indent (in twips)
    pub right_indent: Option<i32>,
    /// First line indent (in twips)
    pub first_line_indent: Option<i32>,
    /// Space before paragraph (in twips)
    pub space_before: Option<u16>,
    /// Space after paragraph (in twips)
    pub space_after: Option<u16>,
    /// Use auto spacing for space before
    pub space_before_auto: Option<bool>,
    /// Use auto spacing for space after
    pub space_after_auto: Option<bool>,
    /// Widow/orphan control
    pub widow_control: Option<bool>,
    /// Keep the paragraph on one page
    pub keep: Option<bool>,
    /// Keep the paragraph with the next paragraph
    pub keep_with_next: Option<bool>,
    /// Insert a page break before this paragraph
    pub page_break_before: Option<bool>,
    /// Bi-directional paragraph
    pub bidi: Option<bool>,
    /// Outline level (0..9)
    pub outline_level: Option<u8>,
    /// Contextual spacing (ignore spacing between same style)
    pub contextual_spacing: Option<bool>,
    /// Mirror indents (for facing pages)
    pub mirror_indents: Option<bool>,
    /// Line spacing descriptor
    pub line_spacing: Option<LineSpacing>,
    // TODO: tabs, borders, shading, numbering
}

/// Represents a text run with formatting
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
struct TextRun {
    /// Text content
    text: String,
    /// Character formatting
    formatting: CharacterFormatting,
}

/// Represents a paragraph
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
struct WritableParagraph {
    /// Text runs in this paragraph
    runs: Vec<TextRun>,
    /// Paragraph formatting
    formatting: ParagraphFormatting,
}

/// Represents a table cell
#[derive(Debug, Clone)]
struct TableCell {
    /// Paragraphs in the cell
    paragraphs: Vec<WritableParagraph>,
}

/// Represents a table row
#[derive(Debug, Clone)]
struct TableRow {
    /// Cells in the row
    cells: Vec<TableCell>,
}

/// Represents a table
#[derive(Debug, Clone)]
struct WritableTable {
    /// Rows in the table
    rows: Vec<TableRow>,
}

/// DOC file writer
///
/// Provides methods to create and modify DOC files.
pub struct DocWriter {
    /// Paragraphs in the document
    paragraphs: Vec<WritableParagraph>,
    /// Tables in the document
    tables: Vec<WritableTable>,
    /// Document properties
    properties: HashMap<String, String>,
    /// Header/Footer texts (None = not set)
    /// Indices map to plcfHdd entries (following Apache POI HeaderStories indexing):
    /// 0..5: footnote/endnote separators (unused here)
    /// 6: even header, 7: odd header, 10: first header
    /// 8: even footer, 9: odd footer, 11: first footer
    header_even: Option<String>,
    header_odd: Option<String>,
    header_first: Option<String>,
    footer_even: Option<String>,
    footer_odd: Option<String>,
    footer_first: Option<String>,
}

impl DocWriter {
    /// Create a new DOC writer
    pub fn new() -> Self {
        Self {
            paragraphs: Vec::new(),
            tables: Vec::new(),
            properties: HashMap::new(),
            header_even: None,
            header_odd: None,
            header_first: None,
            footer_even: None,
            footer_odd: None,
            footer_first: None,
        }
    }

    /// Add a paragraph with plain text
    ///
    /// # Arguments
    ///
    /// * `text` - Paragraph text
    ///
    /// # Returns
    ///
    /// * `Result<(), DocWriteError>` - Success or error
    pub fn add_paragraph(&mut self, text: &str) -> Result<(), DocWriteError> {
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: text.to_string(),
                formatting: CharacterFormatting::default(),
            }],
            formatting: ParagraphFormatting::default(),
        });
        Ok(())
    }

    /// Add a paragraph with formatting
    ///
    /// # Arguments
    ///
    /// * `text` - Paragraph text
    /// * `char_fmt` - Character formatting
    /// * `para_fmt` - Paragraph formatting
    pub fn add_paragraph_with_format(
        &mut self,
        text: &str,
        char_fmt: CharacterFormatting,
        para_fmt: ParagraphFormatting,
    ) -> Result<(), DocWriteError> {
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: text.to_string(),
                formatting: char_fmt,
            }],
            formatting: para_fmt,
        });
        Ok(())
    }

    /// Add a paragraph composed of multiple runs (rich text)
    ///
    /// Each tuple is (text, character formatting) and the whole paragraph shares the
    /// given paragraph formatting.
    pub fn add_paragraph_runs(
        &mut self,
        runs: Vec<(String, CharacterFormatting)>,
        para_fmt: ParagraphFormatting,
    ) -> Result<(), DocWriteError> {
        if runs.is_empty() {
            return self.add_paragraph_with_format("", CharacterFormatting::default(), para_fmt);
        }
        let mut wruns = Vec::with_capacity(runs.len());
        for (text, formatting) in runs {
            wruns.push(TextRun { text, formatting });
        }
        self.paragraphs.push(WritableParagraph {
            runs: wruns,
            formatting: para_fmt,
        });
        Ok(())
    }

    /// Add a hyperlink paragraph using Word field codes (HYPERLINK)
    ///
    /// This creates a field sequence:
    /// - 0x0013 (field begin, fSpec=1)
    /// - Instruction text: `HYPERLINK "url"` (field-vanished)
    /// - 0x0014 (field separator, fSpec=1)
    /// - Display text
    /// - 0x0015 (field end, fSpec=1)
    ///
    /// # Arguments
    /// - `display_text` - Visible link text shown in the document
    /// - `url` - Target URL for the hyperlink (quotes will be escaped)
    /// - `para_fmt` - Paragraph formatting to apply to this paragraph
    pub fn add_hyperlink(
        &mut self,
        display_text: &str,
        url: &str,
        mut para_fmt: ParagraphFormatting,
    ) -> Result<(), DocWriteError> {
        // Stage: Implementing hyperlinks using field codes
        // TODO(stage:headers_footers): Implement header/footer subdocuments via ccpHdd and PLCFs
        // TODO(stage:notes): Implement footnotes/endnotes PLCFs (plcffndRef, plcfendRef, etc.)

        // Escape quotes inside URL by doubling them per Word field syntax
        let escaped = url.replace('"', "\"\"");
        let instr = format!("HYPERLINK \"{}\"", escaped);

        // Default hyperlink visual style (blue + single underline)
        let mut link_fmt = CharacterFormatting::default();
        link_fmt.underline = Some(true);
        link_fmt.color = Some((0x00, 0x00, 0xFF));

        // Field begin/separator/end special chars
        let mut spec_fmt = CharacterFormatting::default();
        spec_fmt.special = Some(true);

        // Field instruction should be hidden (vanished) but not special
        let mut instr_fmt = CharacterFormatting::default();
        instr_fmt.field_vanish = Some(true);

        let runs = vec![
            ("\u{0013}".to_string(), spec_fmt.clone()), // fldBegin
            (instr, instr_fmt),                         // instruction text (hidden)
            ("\u{0014}".to_string(), spec_fmt.clone()), // fldSep
            (display_text.to_string(), link_fmt),       // display text
            ("\u{0015}".to_string(), spec_fmt),         // fldEnd
        ];

        // Keep consistent paragraph spacing defaults for hyperlink paragraph (no auto spacing)
        if para_fmt.space_before_auto.is_none() {
            para_fmt.space_before_auto = Some(false);
        }
        if para_fmt.space_after_auto.is_none() {
            para_fmt.space_after_auto = Some(false);
        }

        self.add_paragraph_runs(runs, para_fmt)
    }

    /// Set the odd-page header text (HeaderStories index 7)
    pub fn set_odd_header(&mut self, text: &str) {
        self.header_odd = Some(text.to_string());
    }
    /// Set the even-page header text (HeaderStories index 6)
    pub fn set_even_header(&mut self, text: &str) {
        self.header_even = Some(text.to_string());
    }
    /// Set the first-page header text (HeaderStories index 10)
    pub fn set_first_header(&mut self, text: &str) {
        self.header_first = Some(text.to_string());
    }
    /// Set the odd-page footer text (HeaderStories index 9)
    pub fn set_odd_footer(&mut self, text: &str) {
        self.footer_odd = Some(text.to_string());
    }
    /// Set the even-page footer text (HeaderStories index 8)
    pub fn set_even_footer(&mut self, text: &str) {
        self.footer_even = Some(text.to_string());
    }
    /// Set the first-page footer text (HeaderStories index 11)
    pub fn set_first_footer(&mut self, text: &str) {
        self.footer_first = Some(text.to_string());
    }

    /// Build header/footer story text and PlcfHdd
    ///
    /// Appends header/footer text to `text_stream`, extends CHPX/PAPX entries and pieces.
    /// Returns (plcfhdd_bytes, header_cp_length). If no header/footer set, returns None.
    #[allow(clippy::too_many_arguments)] // TODO: Refactor to reduce arguments
    fn build_header_story(
        &self,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
    ) -> Option<(Vec<u8>, u32)> {
        // TODO(stage:headers_footers): support complex content (multiple paragraphs, fields)
        // For now, each defined header/footer is one paragraph, terminated by chEop (0x0D)

        // Short-circuit if nothing set
        if self.header_even.is_none()
            && self.header_odd.is_none()
            && self.header_first.is_none()
            && self.footer_even.is_none()
            && self.footer_odd.is_none()
            && self.footer_first.is_none()
        {
            return None;
        }

        // Build index->text mapping for 12 slots defined by POI HeaderStories
        // 0..5 unused here (footnote/endnote separators)
        let mut idx_text: [Option<&str>; 12] = [None; 12];
        if let Some(ref s) = self.header_even {
            idx_text[6] = Some(s.as_str());
        }
        if let Some(ref s) = self.header_odd {
            idx_text[7] = Some(s.as_str());
        }
        if let Some(ref s) = self.header_first {
            idx_text[10] = Some(s.as_str());
        }
        if let Some(ref s) = self.footer_even {
            idx_text[8] = Some(s.as_str());
        }
        if let Some(ref s) = self.footer_odd {
            idx_text[9] = Some(s.as_str());
        }
        if let Some(ref s) = self.footer_first {
            idx_text[11] = Some(s.as_str());
        }

        // Local CP within header story (counts only header subdocument)
        let mut header_cp: u32 = 0;
        let mut cp_starts: [u32; 12] = [0; 12];

        for i in 0..12 {
            cp_starts[i] = header_cp;
            if let Some(text) = idx_text[i] {
                // Add a single paragraph for this header/footer part
                let fc_para_start = text_fc_start + text_stream.len() as u32;
                let mut para_chars: u32 = 0;
                let mut last_run_index_for_para: Option<usize> = None;

                // Character formatting default
                let char_fmt = CharacterFormatting::default();
                let grpprl = build_chpx_grpprl(&char_fmt, font_builder);
                let run_fc_start = fc_para_start;
                // Encode text to UTF-16LE
                for u in text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                para_chars += text.chars().count() as u32;
                // chpx range for run (without paragraph mark yet)
                let run_fc_end = run_fc_start + para_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                last_run_index_for_para = Some(chpx_entries.len() - 1);

                // Paragraph mark for the content paragraph
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                if let Some(last_idx) = last_run_index_for_para {
                    chpx_entries[last_idx].1 += 2; // include content paragraph mark
                }
                let fc_para_end = text_fc_start + text_stream.len() as u32;

                // PAPX for content paragraph
                let pap_grpprl = build_papx_grpprl(&ParagraphFormatting::default());
                papx_entries.push((fc_para_start, fc_para_end, pap_grpprl));

                // Guard paragraph mark between stories (required by MS-DOC for non-empty stories)
                let fc_guard_start = fc_para_end;
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                // Extend last run to cover the guard mark as well (default formatting)
                if let Some(last_idx) = last_run_index_for_para {
                    chpx_entries[last_idx].1 += 2;
                }
                let fc_guard_end = fc_guard_start + 2;
                // PAPX for guard empty paragraph (default formatting)
                let guard_pap_grpprl = build_papx_grpprl(&ParagraphFormatting::default());
                papx_entries.push((fc_guard_start, fc_guard_end, guard_pap_grpprl));

                // Add to global piece table (CLX) after main body; include content + guard marks
                let fc_offset = fc_para_start;
                pieces.push(Piece::new(
                    *current_cp_total,
                    *current_cp_total + para_chars + 2,
                    fc_offset,
                    true,
                ));
                *current_cp_total += para_chars + 2;

                // Local header story CPs include content chars + content EOP + guard EOP
                header_cp += para_chars + 2;
            }
        }

        // Trailing placeholder EOP after the last story (per MS-DOC PlcfHdd requirements):
        // Ensure that the final CP (n+2) is exactly one greater than the last story's end CP (n+1),
        // with an actual chEop stored at that file position.
        if header_cp > 0 {
            let fc_trailing_start = text_fc_start + text_stream.len() as u32;
            // write trailing chEop
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            // extend last CHPX to cover the trailing eop
            if let Some((_, last_end, _)) = chpx_entries.last_mut() {
                *last_end += 2;
            }
            let fc_trailing_end = text_fc_start + text_stream.len() as u32;
            // PAPX for trailing empty paragraph
            let trailing_pap = build_papx_grpprl(&ParagraphFormatting::default());
            papx_entries.push((fc_trailing_start, fc_trailing_end, trailing_pap));
            // Piece for trailing EOP
            pieces.push(Piece::new(
                *current_cp_total,
                *current_cp_total + 1,
                fc_trailing_start,
                true,
            ));
            *current_cp_total += 1;

            // Header CP accounts for this trailing paragraph mark
            header_cp += 1;
        }

        // Build PlcfHdd (PLCF with cbStruct=0): cp starts for 12 props + final end CP
        let mut plcfhdd = Vec::with_capacity((12 + 1) * 4);
        for cp_start in &cp_starts {
            plcfhdd.extend_from_slice(&cp_start.to_le_bytes());
        }
        plcfhdd.extend_from_slice(&header_cp.to_le_bytes());

        Some((plcfhdd, header_cp))
    }

    /// Create a new table with the specified dimensions
    ///
    /// # Arguments
    ///
    /// * `rows` - Number of rows
    /// * `cols` - Number of columns
    ///
    /// # Returns
    ///
    /// * `Result<usize, DocWriteError>` - Table index or error
    ///
    /// # Implementation Status
    ///
    /// Table creation with TAP (Table Properties) structures is deferred.
    /// Use the DOCX writer for production table support.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Result<usize, DocWriteError> {
        if rows == 0 || cols == 0 {
            return Err(DocWriteError::InvalidData(
                "Table must have at least 1 row and 1 column".to_string(),
            ));
        }

        let mut table = WritableTable { rows: Vec::new() };

        for _ in 0..rows {
            let mut row = TableRow { cells: Vec::new() };
            for _ in 0..cols {
                row.cells.push(TableCell {
                    paragraphs: vec![WritableParagraph {
                        runs: vec![TextRun {
                            text: String::new(),
                            formatting: CharacterFormatting::default(),
                        }],
                        formatting: ParagraphFormatting::default(),
                    }],
                });
            }
            table.rows.push(row);
        }

        let index = self.tables.len();
        self.tables.push(table);
        Ok(index)
    }

    /// Set text in a table cell
    ///
    /// # Arguments
    ///
    /// * `table_idx` - Table index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `text` - Cell text
    pub fn set_table_cell_text(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
        text: &str,
    ) -> Result<(), DocWriteError> {
        let table = self
            .tables
            .get_mut(table_idx)
            .ok_or_else(|| DocWriteError::InvalidData(format!("Table {} not found", table_idx)))?;

        let row_data = table
            .rows
            .get_mut(row)
            .ok_or_else(|| DocWriteError::InvalidData(format!("Row {} not found", row)))?;

        let cell = row_data
            .cells
            .get_mut(col)
            .ok_or_else(|| DocWriteError::InvalidData(format!("Column {} not found", col)))?;

        cell.paragraphs = vec![WritableParagraph {
            runs: vec![TextRun {
                text: text.to_string(),
                formatting: CharacterFormatting::default(),
            }],
            formatting: ParagraphFormatting::default(),
        }];

        Ok(())
    }

    /// Set a document property
    ///
    /// # Arguments
    ///
    /// * `name` - Property name (e.g., "Title", "Author", "Subject")
    /// * `value` - Property value
    pub fn set_property(&mut self, name: &str, value: &str) {
        self.properties.insert(name.to_string(), value.to_string());
    }

    /// Save the document to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), DocWriteError>` - Success or error
    ///
    /// # Implementation
    ///
    /// This generates a complete Word 97-2003 binary file conforming to MS-DOC specification:
    /// - FIB (File Information Block) - [MS-DOC] Section 2.5
    /// - Text stream with piece table - [MS-DOC] Section 2.8
    /// - Character and paragraph formatting via SPRMs - [MS-DOC] Section 2.6.1
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), DocWriteError> {
        // Based on Apache POI's HWPFDocument.write() implementation
        // This includes ALL mandatory structures required by Microsoft Word

        // Initialize streams for building the document
        let mut word_document_stream = Vec::new();
        let mut table_stream = Vec::new();

        // 1. Reserve space for FIB (File Information Block)
        // Word 2007+ format (nFib 0x0101) requires 1242 bytes
        // We'll write the actual FIB later after we know all the offsets
        let fib_placeholder = vec![0u8; 1242];
        word_document_stream.extend_from_slice(&fib_placeholder);

        // 2. fcMin: will be the padded start of text (set after padding below)

        // 3. Build text stream and piece table
        // Text starts immediately after FIB, with padding to 512-byte boundary
        // Per POI's TextPieceTable.writeTo() lines 427-433
        let mut text_stream = Vec::new();
        let mut current_cp = 0u32; // Character position in document
        let mut pieces = Vec::new();
        let mut chpx_entries: Vec<(u32, u32, Vec<u8>)> = Vec::new();
        let mut papx_entries: Vec<(u32, u32, Vec<u8>)> = Vec::new();
        let mut font_builder = FontTableBuilder::new();

        // Pad to 512-byte boundary before writing text (POI line 428-433)
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        let text_fc_start = word_document_stream.len() as u32;
        // Align fcMin to actual (padded) start of text
        let fc_min: u32 = text_fc_start;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = run_text.chars().count() as u32;
                // Build run grpprl
                let grpprl = build_chpx_grpprl(&run.formatting, &mut font_builder);
                // Encode run text to UTF-16LE
                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }
            // Add paragraph mark (0x0D), include in last run coverage to avoid gaps
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            if let Some(last_idx) = last_run_index_for_para {
                chpx_entries[last_idx].1 += 2; // include paragraph mark bytes
            }
            let fc_para_end = text_fc_start + text_stream.len() as u32;
            // Build PAPX grpprl for paragraph
            let pap_grpprl = build_papx_grpprl(&paragraph.formatting);
            papx_entries.push((fc_para_start, fc_para_end, pap_grpprl));

            // Piece descriptor for this paragraph (includes paragraph mark)
            let fc_offset = fc_para_start;
            pieces.push(Piece::new(
                current_cp,
                current_cp + para_chars + 1,
                fc_offset,
                true,
            ));
            current_cp += para_chars + 1;
        }

        // ccpText (main body) length in CPs
        let text_length = current_cp;

        // Stage: headers_footers - build header story and PlcfHdd
        // This appends header/footer text to the text stream and extends FKPs/pieces
        let mut header_plcfhdd: Option<(Vec<u8>, u32)> = None;
        if let Some((plcf_bytes, header_cp)) = self.build_header_story(
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        ) {
            header_plcfhdd = Some((plcf_bytes, header_cp));
        }

        // Initialize FIB builder
        let mut fib = FibBuilder::new();
        fib.set_main_text(0, text_length);
        if let Some((_, header_cp)) = &header_plcfhdd {
            fib.set_ccp_hdd(*header_cp);
        }

        let mut table_offset = 0u32;

        // 3. Write StyleSheet to table stream (MANDATORY - POI line 681-684)
        let stylesheet_data = crate::ole::doc::writer::stylesheet::generate_minimal_stylesheet();
        fib.set_stshf(table_offset, stylesheet_data.len() as u32);
        table_stream.extend_from_slice(&stylesheet_data);
        table_offset = table_stream.len() as u32;

        // 4. Write piece table (Clx) to table stream (POI line 699-702)
        let mut piece_table = PieceTableBuilder::new();
        for piece in pieces {
            piece_table.add_piece(piece);
        }
        let clx_data = piece_table.generate()?;
        fib.set_clx(table_offset, clx_data.len() as u32);
        table_stream.extend_from_slice(&clx_data);
        table_offset = table_stream.len() as u32;

        // 5. Write DocumentProperties to table stream (MANDATORY - POI line 715-718)
        // Set fFacingPages if even headers/footers are present, and set doc-level grpfIhdt mask
        let mut doc_grpf_ihdt: u8 = 0;
        if self.header_even.is_some() {
            doc_grpf_ihdt |= 0x01;
        }
        if self.header_odd.is_some() {
            doc_grpf_ihdt |= 0x02;
        }
        if self.footer_even.is_some() {
            doc_grpf_ihdt |= 0x04;
        }
        if self.footer_odd.is_some() {
            doc_grpf_ihdt |= 0x08;
        }
        if self.header_first.is_some() {
            doc_grpf_ihdt |= 0x10;
        }
        if self.footer_first.is_some() {
            doc_grpf_ihdt |= 0x20;
        }
        let facing_pages = self.header_even.is_some() || self.footer_even.is_some();
        let dop_data = crate::ole::doc::writer::dop::generate_dop(facing_pages, doc_grpf_ihdt);
        fib.set_dop(table_offset, dop_data.len() as u32);
        table_stream.extend_from_slice(&dop_data);
        table_offset = table_stream.len() as u32;

        // Write PlcfHdd if present (headers/footers PLCF)
        if let Some((plcf_bytes, _header_cp)) = &header_plcfhdd {
            fib.set_plcfhdd(table_offset, plcf_bytes.len() as u32);
            table_stream.extend_from_slice(plcf_bytes);
            table_offset = table_stream.len() as u32;
        }

        // 6. Write CHPX bin table (character formatting) (MANDATORY - POI line 753-756)
        // NOTE: Will be populated after FKPs are written
        let chpx_bin_table_offset = table_offset;
        // Reserve space for actual bin table (will be written later)
        // Bin table with 1 entry: 2 CPs (8 bytes) + 1 PN (4 bytes) = 12 bytes
        table_stream.extend_from_slice(&[0u8; 12]);
        fib.set_plcfbte_chpx(table_offset, 12);
        table_offset = table_stream.len() as u32;

        // 7. Write PAPX bin table (paragraph formatting) (MANDATORY - POI line 767-771)
        // NOTE: Will be populated after FKPs are written
        let papx_bin_table_offset = table_offset;
        // Reserve space for actual bin table
        table_stream.extend_from_slice(&[0u8; 12]);
        fib.set_plcfbte_papx(table_offset, 12);
        table_offset = table_stream.len() as u32;

        // 8. Reserve space for section table - will write after SEPX is created
        let section_table_offset = table_offset;
        let section_table_placeholder = vec![0u8; 20]; // 20 bytes: 2 CPs (8) + 1 SED (12)
        table_stream.extend_from_slice(&section_table_placeholder);
        table_offset = table_stream.len() as u32;

        // 9. Write Font Table to table stream (MANDATORY - POI line 899-903)
        let font_table = font_builder.generate();
        fib.set_sttbfffn(table_offset, font_table.len() as u32);
        table_stream.extend_from_slice(&font_table);

        // 10. Append text (main + headers/footers) to WordDocument stream
        word_document_stream.extend_from_slice(&text_stream);

        // Capture fcMac AFTER text, BEFORE FKPs (POI line 703)
        let fc_mac_value = word_document_stream.len() as u32;
        // Compute FC end of text range for PLCFBTE (bin tables use FC domain per POI)
        let text_fc_end = text_fc_start + text_stream.len() as u32;

        // 10a. Write FKPs to WordDocument stream (CRITICAL - POI line 450-492)
        // FKPs must start at 512-byte aligned offsets
        // Pad to 512-byte boundary
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        // Calculate page number for first CHPX FKP
        let chpx_fkp_page = (word_document_stream.len() / 512) as u32;

        // Create CHPX FKP from collected runs (FKP entries use FC)
        let mut chpx_builder = crate::ole::doc::writer::fkp::ChpxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_fkp = chpx_builder.generate()?;
        word_document_stream.extend_from_slice(&chpx_fkp);

        // Update CHPX bin table with actual page number (FC domain)
        let chpx_bin_table_with_fkp =
            crate::ole::doc::writer::bin_table::generate_single_entry_bin_table(
                text_fc_start,
                text_fc_end,
                chpx_fkp_page,
            );
        // Write the bin table to table_stream
        let chpx_table_start = chpx_bin_table_offset as usize;
        let chpx_table_end = chpx_table_start + 12;
        table_stream[chpx_table_start..chpx_table_end].copy_from_slice(&chpx_bin_table_with_fkp);

        // Calculate page number for first PAPX FKP (next 512-byte page)
        let papx_fkp_page = (word_document_stream.len() / 512) as u32;

        // Create PAPX FKP (one entry per paragraph, FC domain)
        let mut papx_builder = crate::ole::doc::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_fkp = papx_builder.generate()?;
        word_document_stream.extend_from_slice(&papx_fkp);

        // Update PAPX bin table with actual page number (FC domain)
        let papx_bin_table_with_fkp =
            crate::ole::doc::writer::bin_table::generate_single_entry_bin_table(
                text_fc_start,
                text_fc_end,
                papx_fkp_page,
            );
        // Write the bin table to table_stream
        let papx_table_start = papx_bin_table_offset as usize;
        let papx_table_end = papx_table_start + 12;
        table_stream[papx_table_start..papx_table_end].copy_from_slice(&papx_bin_table_with_fkp);

        // 10b. Write SEPX to WordDocument stream (Apache POI line 825)
        // SEPX is written AFTER text and FKPs, per SectionTable.writeTo()
        let sepx_offset = word_document_stream.len() as u32;
        // Compute grpfIhdt bitfield and fTitlePage based on presence of headers/footers
        let mut grpf_ihdt: u8 = 0;
        if self.header_even.is_some() {
            grpf_ihdt |= 0x01;
        }
        if self.header_odd.is_some() {
            grpf_ihdt |= 0x02;
        }
        if self.footer_even.is_some() {
            grpf_ihdt |= 0x04;
        }
        if self.footer_odd.is_some() {
            grpf_ihdt |= 0x08;
        }
        if self.header_first.is_some() {
            grpf_ihdt |= 0x10;
        }
        if self.footer_first.is_some() {
            grpf_ihdt |= 0x20;
        }
        let first_page = self.header_first.is_some() || self.footer_first.is_some();
        let sepx_data = crate::ole::doc::writer::section::generate_sepx(first_page, grpf_ihdt);
        word_document_stream.extend_from_slice(&sepx_data);

        // 10c. Now write section table with correct SEPX offset
        let section_table =
            crate::ole::doc::writer::section::generate_section_table(text_length, sepx_offset);
        table_stream[section_table_offset as usize..(section_table_offset as usize + 20)]
            .copy_from_slice(&section_table);
        fib.set_plcfsed(section_table_offset, 20);

        // 11. Set FibBase fields (Apache POI line 906-914)
        // fcMin = start of text (after FIB)
        // fcMac = end of text (captured before FKPs)
        // cbMac = total WordDocument stream size (after SEPX)
        let cb_mac = word_document_stream.len() as u32; // Total document size
        fib.set_base_fields(fc_min, fc_mac_value, cb_mac);

        // 12. Generate FIB with all offsets set
        let fib_data = fib.generate()?;

        // 13. Write FIB at the beginning of WordDocument stream
        // Word 2007+ format FIB is 1242 bytes, not 512!
        word_document_stream[0..fib_data.len()].copy_from_slice(&fib_data);

        // 14. Pad streams to 4096 bytes (Apache POI line 911-921)
        // This ensures proper sector alignment in the OLE file
        fn pad_to_4096(stream: &mut Vec<u8>) {
            let remainder = stream.len() % 4096;
            if remainder != 0 {
                let padding = 4096 - remainder;
                stream.resize(stream.len() + padding, 0);
            }
        }

        pad_to_4096(&mut word_document_stream);
        pad_to_4096(&mut table_stream);

        // 15. Create OLE compound document
        let mut ole_writer = OleWriter::new();

        // Set Word document CLSID (REQUIRED for Microsoft Word to recognize the file)
        // CLSID: {00020906-0000-0000-C000-000000000046}
        let word_clsid = [
            0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x46,
        ];
        ole_writer.set_root_clsid(word_clsid);

        // WordDocument stream FIRST to guarantee sector 0, then 1Table
        ole_writer.create_stream(&["WordDocument"], &word_document_stream)?;
        ole_writer.create_stream(&["1Table"], &table_stream)?;

        // Create OLE metadata streams (optional for type association)
        let compobj_data = crate::ole::doc::writer::ole_metadata::generate_compobj_stream();
        let ole_data = crate::ole::doc::writer::ole_metadata::generate_ole_stream();
        ole_writer.create_stream(&["\x01CompObj"], &compobj_data)?;
        ole_writer.create_stream(&["\x01Ole"], &ole_data)?;

        // 16. Save to file
        ole_writer.save(path)?;

        Ok(())
    }

    /// Write to an in-memory buffer
    pub fn write_to<W: std::io::Write + std::io::Seek>(
        &mut self,
        writer: &mut W,
    ) -> Result<(), DocWriteError> {
        // Same implementation as save() but writes to a writer
        // Based on Apache POI's HWPFDocument.write() implementation

        let mut word_document_stream = Vec::new();
        let mut table_stream = Vec::new();

        // Reserve space for FIB (Word 2007+ format = 1242 bytes)
        let fib_placeholder = vec![0u8; 1242];
        word_document_stream.extend_from_slice(&fib_placeholder);

        // fcMin will be set to padded start of text (after 512 alignment below)

        // Build text stream and piece table
        let mut text_stream = Vec::new();
        let mut current_cp = 0u32;
        let mut pieces = Vec::new();
        let mut chpx_entries: Vec<(u32, u32, Vec<u8>)> = Vec::new();
        let mut papx_entries: Vec<(u32, u32, Vec<u8>)> = Vec::new();
        let mut font_builder = FontTableBuilder::new();

        // Pad to 512-byte boundary before text
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        let text_fc_start = word_document_stream.len() as u32;
        // fcMin is the actual (padded) start of text
        let fc_min: u32 = text_fc_start;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = run_text.chars().count() as u32;
                let grpprl = build_chpx_grpprl(&run.formatting, &mut font_builder);
                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }
            // Paragraph mark (0x0D)
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            if let Some(last_idx) = last_run_index_for_para {
                chpx_entries[last_idx].1 += 2;
            }
            let fc_para_end = text_fc_start + text_stream.len() as u32;
            let pap_grpprl = build_papx_grpprl(&paragraph.formatting);
            papx_entries.push((fc_para_start, fc_para_end, pap_grpprl));

            let fc_offset = fc_para_start;
            pieces.push(Piece::new(
                current_cp,
                current_cp + para_chars + 1,
                fc_offset,
                true,
            ));
            current_cp += para_chars + 1;
        }

        let text_length = current_cp; // ccpText
        // Stage: headers_footers - build header story and PlcfHdd
        let mut header_plcfhdd: Option<(Vec<u8>, u32)> = None;
        if let Some((plcf_bytes, header_cp)) = self.build_header_story(
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        ) {
            header_plcfhdd = Some((plcf_bytes, header_cp));
        }

        let mut fib = FibBuilder::new();
        fib.set_main_text(0, text_length);
        if let Some((_, header_cp)) = &header_plcfhdd {
            fib.set_ccp_hdd(*header_cp);
        }
        let mut table_offset = 0u32;

        // Write all mandatory structures to table stream
        let stylesheet_data = crate::ole::doc::writer::stylesheet::generate_minimal_stylesheet();
        fib.set_stshf(table_offset, stylesheet_data.len() as u32);
        table_stream.extend_from_slice(&stylesheet_data);
        table_offset = table_stream.len() as u32;

        let mut piece_table = PieceTableBuilder::new();
        for piece in pieces {
            piece_table.add_piece(piece);
        }
        let clx_data = piece_table.generate()?;
        fib.set_clx(table_offset, clx_data.len() as u32);
        table_stream.extend_from_slice(&clx_data);
        table_offset = table_stream.len() as u32;

        // DocumentProperties (again in this code path): set fFacingPages and doc-level grpfIhdt
        let mut doc_grpf_ihdt: u8 = 0;
        if self.header_even.is_some() {
            doc_grpf_ihdt |= 0x01;
        }
        if self.header_odd.is_some() {
            doc_grpf_ihdt |= 0x02;
        }
        if self.footer_even.is_some() {
            doc_grpf_ihdt |= 0x04;
        }
        if self.footer_odd.is_some() {
            doc_grpf_ihdt |= 0x08;
        }
        if self.header_first.is_some() {
            doc_grpf_ihdt |= 0x10;
        }
        if self.footer_first.is_some() {
            doc_grpf_ihdt |= 0x20;
        }
        let facing_pages = self.header_even.is_some() || self.footer_even.is_some();
        let dop_data = crate::ole::doc::writer::dop::generate_dop(facing_pages, doc_grpf_ihdt);
        fib.set_dop(table_offset, dop_data.len() as u32);
        table_stream.extend_from_slice(&dop_data);
        table_offset = table_stream.len() as u32;

        let chpx_bin_table_offset = table_offset;
        table_stream.extend_from_slice(&[0u8; 12]);
        fib.set_plcfbte_chpx(table_offset, 12);
        table_offset = table_stream.len() as u32;

        let papx_bin_table_offset = table_offset;
        table_stream.extend_from_slice(&[0u8; 12]);
        fib.set_plcfbte_papx(table_offset, 12);
        table_offset = table_stream.len() as u32;

        let section_table_offset = table_offset;
        let section_table_placeholder = vec![0u8; 20]; // 20 bytes: 2 CPs (8) + 1 SED (12)
        table_stream.extend_from_slice(&section_table_placeholder);
        table_offset = table_stream.len() as u32;

        let font_table = font_builder.generate();
        fib.set_sttbfffn(table_offset, font_table.len() as u32);
        table_stream.extend_from_slice(&font_table);

        // Append text and write FKPs
        word_document_stream.extend_from_slice(&text_stream);

        // Capture fcMac AFTER text, BEFORE FKPs (POI line 703)
        let fc_mac_value = word_document_stream.len() as u32;

        // Write FKPs to WordDocument stream at 512-byte aligned offsets
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        let chpx_fkp_page = (word_document_stream.len() / 512) as u32;
        let mut chpx_builder = crate::ole::doc::writer::fkp::ChpxFkpBuilder::new();
        let text_fc_end = text_fc_start + text_stream.len() as u32; // FC end in bytes
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_fkp = chpx_builder.generate()?;
        word_document_stream.extend_from_slice(&chpx_fkp);

        // Use FC domain (byte offsets) for bin table ranges, same as in save()
        let chpx_bin_table_with_fkp =
            crate::ole::doc::writer::bin_table::generate_single_entry_bin_table(
                text_fc_start,
                text_fc_end,
                chpx_fkp_page,
            );
        let chpx_table_start = chpx_bin_table_offset as usize;
        let chpx_table_end = chpx_table_start + 12;
        table_stream[chpx_table_start..chpx_table_end].copy_from_slice(&chpx_bin_table_with_fkp);

        let papx_fkp_page = (word_document_stream.len() / 512) as u32;
        let mut papx_builder = crate::ole::doc::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_fkp = papx_builder.generate()?;
        word_document_stream.extend_from_slice(&papx_fkp);

        // Use FC domain (byte offsets) for bin table ranges, same as in save()
        let papx_bin_table_with_fkp =
            crate::ole::doc::writer::bin_table::generate_single_entry_bin_table(
                text_fc_start,
                text_fc_end,
                papx_fkp_page,
            );
        let papx_table_start = papx_bin_table_offset as usize;
        let papx_table_end = papx_table_start + 12;
        table_stream[papx_table_start..papx_table_end].copy_from_slice(&papx_bin_table_with_fkp);

        // Write SEPX to WordDocument stream (after text and FKPs)
        let sepx_offset = word_document_stream.len() as u32;
        let mut grpf_ihdt: u8 = 0;
        if self.header_even.is_some() {
            grpf_ihdt |= 0x01;
        }
        if self.header_odd.is_some() {
            grpf_ihdt |= 0x02;
        }
        if self.footer_even.is_some() {
            grpf_ihdt |= 0x04;
        }
        if self.footer_odd.is_some() {
            grpf_ihdt |= 0x08;
        }
        if self.header_first.is_some() {
            grpf_ihdt |= 0x10;
        }
        if self.footer_first.is_some() {
            grpf_ihdt |= 0x20;
        }
        let first_page = self.header_first.is_some() || self.footer_first.is_some();
        let sepx_data = crate::ole::doc::writer::section::generate_sepx(first_page, grpf_ihdt);
        word_document_stream.extend_from_slice(&sepx_data);

        // Write section table with correct SEPX offset
        let section_table =
            crate::ole::doc::writer::section::generate_section_table(text_length, sepx_offset);
        table_stream[section_table_offset as usize..(section_table_offset as usize + 20)]
            .copy_from_slice(&section_table);
        fib.set_plcfsed(section_table_offset, 20);

        // Set FibBase fields
        let cb_mac = word_document_stream.len() as u32; // Total size after SEPX
        fib.set_base_fields(fc_min, fc_mac_value, cb_mac);
        let fib_data = fib.generate()?;
        word_document_stream[0..fib_data.len()].copy_from_slice(&fib_data);

        // Ensure both streams are large (>= 4096) so WordDocument is allocated in regular FAT
        fn pad_to_4096(stream: &mut Vec<u8>) {
            let remainder = stream.len() % 4096;
            if remainder != 0 {
                let padding = 4096 - remainder;
                stream.resize(stream.len() + padding, 0);
            }
        }
        pad_to_4096(&mut word_document_stream);
        pad_to_4096(&mut table_stream);

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();

        // Set Word document CLSID (REQUIRED for Microsoft Word)
        let word_clsid = [
            0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x46,
        ];
        ole_writer.set_root_clsid(word_clsid);

        // Ensure WordDocument gets sector 0: add it first, then 1Table
        ole_writer.create_stream(&["WordDocument"], &word_document_stream)?;
        ole_writer.create_stream(&["1Table"], &table_stream)?;

        // Add metadata streams after core ones
        let compobj_data = crate::ole::doc::writer::ole_metadata::generate_compobj_stream();
        let ole_data = crate::ole::doc::writer::ole_metadata::generate_ole_stream();
        ole_writer.create_stream(&["\x01CompObj"], &compobj_data)?;
        ole_writer.create_stream(&["\x01Ole"], &ole_data)?;
        ole_writer.write_to(writer)?;

        Ok(())
    }

    // Helper methods for DOC writer:
    // The following are implemented via the modular components:
    // - Generating FIB structure (File Information Block)
    // - Building piece table for text storage
    // - Generating SPRM sequences for character formatting (CHP)
    // - Generating SPRM sequences for paragraph formatting (PAP)
    // - Building FKP (Formatted Disk Page) structures
    // - Generating table properties (TAP)
    // - Encoding text to Word's internal format
    // - Managing style definitions
    // - Font table generation
}

/// Build a CHPX grpprl (group of SPRMs) from CharacterFormatting
fn build_chpx_grpprl(fmt: &CharacterFormatting, font_builder: &mut FontTableBuilder) -> Vec<u8> {
    let mut grp = Vec::with_capacity(16);

    #[inline]
    fn push_byte(grp: &mut Vec<u8>, opcode: u16, val: u8) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.push(val);
    }

    #[inline]
    fn push_word(grp: &mut Vec<u8>, opcode: u16, val: u16) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&val.to_le_bytes());
    }

    #[inline]
    fn push_dword(grp: &mut Vec<u8>, opcode: u16, val: u32) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&val.to_le_bytes());
    }

    // Bold
    if let Some(b) = fmt.bold {
        push_byte(&mut grp, SPRM_C_F_BOLD, if b { 1 } else { 0 });
    }
    // Italic
    if let Some(i) = fmt.italic {
        push_byte(&mut grp, SPRM_C_F_ITALIC, if i { 1 } else { 0 });
    }
    // Underline (1 = single, 0 = none)
    if let Some(u) = fmt.underline {
        push_byte(&mut grp, SPRM_C_KUL, if u { 1 } else { 0 });
    }
    // Strikethrough
    if let Some(s) = fmt.strike {
        push_byte(&mut grp, SPRM_C_F_STRIKE, if s { 1 } else { 0 });
    }
    // Double strikethrough
    if let Some(ds) = fmt.double_strike {
        push_byte(&mut grp, SPRM_C_F_D_STRIKE, if ds { 1 } else { 0 });
    }
    // Superscript/Subscript via sprmCIss (0=none,1=super,2=sub)
    let mut iss: Option<u8> = None;
    if let Some(true) = fmt.superscript {
        iss = Some(1);
    } else if let Some(true) = fmt.subscript {
        iss = Some(2);
    }
    if let Some(v) = iss {
        push_byte(&mut grp, SPRM_C_ISS, v);
    }
    // Small caps / All caps / Hidden
    if let Some(sc) = fmt.small_caps {
        push_byte(&mut grp, SPRM_C_F_SMALL_CAPS, if sc { 1 } else { 0 });
    }
    if let Some(ac) = fmt.all_caps {
        push_byte(&mut grp, SPRM_C_F_CAPS, if ac { 1 } else { 0 });
    }
    if let Some(h) = fmt.hidden {
        push_byte(&mut grp, SPRM_C_F_VANISH, if h { 1 } else { 0 });
    }
    // Special/Field vanish (for field codes and control chars)
    if let Some(sp) = fmt.special {
        push_byte(&mut grp, SPRM_C_F_SPEC, if sp { 1 } else { 0 });
    }
    if let Some(vn) = fmt.field_vanish {
        push_byte(&mut grp, SPRM_C_F_FLD_VANISH, if vn { 1 } else { 0 });
    }
    // Font size (half-points)
    if let Some(hps) = fmt.font_size {
        push_word(&mut grp, SPRM_C_HPS, hps);
    }
    // Font name -> map to ftc index via FontTableBuilder and set default font
    if let Some(name) = &fmt.font_name {
        let idx = font_builder.get_or_add(name);
        push_word(&mut grp, SPRM_C_FTC_DEFAULT, idx);
    }
    // Color (RGB) -> sprmCCv expects a 4-byte value
    if let Some((r, g, b)) = fmt.color {
        let cv: u32 = (r as u32) | ((g as u32) << 8) | ((b as u32) << 16);
        push_dword(&mut grp, SPRM_C_CV, cv);
    }

    grp
}

/// Build a PAPX grpprl (group of SPRMs) from ParagraphFormatting
fn build_papx_grpprl(fmt: &ParagraphFormatting) -> Vec<u8> {
    let mut grp = Vec::with_capacity(16);

    #[inline]
    fn push_byte(grp: &mut Vec<u8>, opcode: u16, val: u8) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.push(val);
    }

    #[inline]
    fn push_i16(grp: &mut Vec<u8>, opcode: u16, val: i16) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&(val as u16).to_le_bytes());
    }

    #[inline]
    fn push_u16(grp: &mut Vec<u8>, opcode: u16, val: u16) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&val.to_le_bytes());
    }

    #[inline]
    fn push_bool(grp: &mut Vec<u8>, opcode: u16, val: bool) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.push(if val { 1 } else { 0 });
    }

    // Alignment (emit both legacy and modern; modern last to take precedence)
    if let Some(jc) = fmt.alignment {
        push_byte(&mut grp, SPRM_P_JC, jc);
        push_byte(&mut grp, SPRM_P_JC_LOGICAL, jc);
    }
    // Indents (twips). Emit legacy and modern variants. Values are signed twips.
    if let Some(dxa_left) = fmt.left_indent {
        let v = dxa_left.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
        push_i16(&mut grp, SPRM_P_DXA_LEFT, v);
        push_i16(&mut grp, SPRM_P_DXA_LEFT_2000, v);
    }
    if let Some(dxa_right) = fmt.right_indent {
        let v = dxa_right.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
        push_i16(&mut grp, SPRM_P_DXA_RIGHT, v);
        push_i16(&mut grp, SPRM_P_DXA_RIGHT_2000, v);
    }
    if let Some(dxa_first) = fmt.first_line_indent {
        let v = dxa_first.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
        push_i16(&mut grp, SPRM_P_DXA_LEFT1, v);
        push_i16(&mut grp, SPRM_P_DXA_LEFT1_2000, v);
    }
    // Spacing (twips)
    if let Some(dya_before) = fmt.space_before {
        push_u16(&mut grp, SPRM_P_DYA_BEFORE, dya_before);
    }
    if let Some(dya_after) = fmt.space_after {
        push_u16(&mut grp, SPRM_P_DYA_AFTER, dya_after);
    }

    // Auto spacing flags
    if let Some(auto) = fmt.space_before_auto {
        push_bool(&mut grp, SPRM_P_F_DYA_BEFORE_AUTO, auto);
    }
    if let Some(auto) = fmt.space_after_auto {
        push_bool(&mut grp, SPRM_P_F_DYA_AFTER_AUTO, auto);
    }

    // Keep, keep-with-next, page break before
    if let Some(keep) = fmt.keep {
        push_bool(&mut grp, SPRM_P_F_KEEP, keep);
    }
    if let Some(keep_next) = fmt.keep_with_next {
        push_bool(&mut grp, SPRM_P_F_KEEP_FOLLOW, keep_next);
    }
    if let Some(pbb) = fmt.page_break_before {
        push_bool(&mut grp, SPRM_P_F_PAGE_BREAK_BEFORE, pbb);
    }

    // Widow/orphan control
    if let Some(wc) = fmt.widow_control {
        push_bool(&mut grp, SPRM_P_F_WIDOW_CONTROL, wc);
    }

    // BiDi paragraph
    if let Some(bidi) = fmt.bidi {
        push_bool(&mut grp, SPRM_P_F_BI_DI, bidi);
    }

    // Outline level
    if let Some(lvl) = fmt.outline_level {
        grp.extend_from_slice(&SPRM_P_OUT_LVL.to_le_bytes());
        grp.push(lvl);
    }

    // Contextual spacing and mirror indents
    if let Some(cs) = fmt.contextual_spacing {
        push_bool(&mut grp, SPRM_P_F_CONTEXTUAL_SPACING, cs);
    }
    if let Some(mi) = fmt.mirror_indents {
        push_bool(&mut grp, SPRM_P_F_MIRROR_INDENTS, mi);
    }

    // Line spacing (LSPD: 4 bytes = dyaLine (i16 LE), fMulti (i16 LE))
    if let Some(ls) = fmt.line_spacing {
        let mut bytes = [0u8; 4];
        let f_multi: u16 = if ls.is_multiple { 1 } else { 0 };
        bytes[0..2].copy_from_slice(&(ls.dya_line as u16).to_le_bytes());
        bytes[2..4].copy_from_slice(&f_multi.to_le_bytes());
        grp.extend_from_slice(&SPRM_P_DYA_LINE.to_le_bytes());
        grp.extend_from_slice(&bytes);
    }

    // Line spacing is more complex (structure) -> TODO implement later to ensure spec compliance

    grp
}

impl Default for DocWriter {
    fn default() -> Self {
        Self::new()
    }
}

// Implementation deferred - DOC binary format functions:
// These would be needed for full DOC file generation:
// - FIB (File Information Block) generation
// - Piece table builder for text storage
// - SPRM generation for CHP (Character Properties)
// - SPRM generation for PAP (Paragraph Properties)
// - FKP (Formatted Disk Page) builder
// - TAP (Table Properties) builder
//
// Recommendation: Use the DOCX writer (fully implemented) for production use.

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_writer() {
        let writer = DocWriter::new();
        assert_eq!(writer.paragraphs.len(), 0);
        assert_eq!(writer.tables.len(), 0);
    }

    #[test]
    fn test_add_paragraph() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("Test").unwrap();
        assert_eq!(writer.paragraphs.len(), 1);
        assert_eq!(writer.paragraphs[0].runs[0].text, "Test");
    }

    #[test]
    fn test_add_table() {
        let mut writer = DocWriter::new();
        let idx = writer.add_table(2, 3).unwrap();
        assert_eq!(idx, 0);
        assert_eq!(writer.tables[0].rows.len(), 2);
        assert_eq!(writer.tables[0].rows[0].cells.len(), 3);
    }

    #[test]
    fn test_set_table_cell() {
        let mut writer = DocWriter::new();
        let idx = writer.add_table(2, 2).unwrap();
        writer.set_table_cell_text(idx, 0, 0, "Cell").unwrap();
        assert_eq!(
            writer.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text,
            "Cell"
        );
    }
}
