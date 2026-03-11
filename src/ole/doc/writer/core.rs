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
use super::footnotes::FootnoteEntry;
use super::numbering::{ListFormatOverride, ListStructure, NumberingWriter};
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
    /// List level index (0-based, used with `ilfo` to associate paragraph with a list)
    pub ilvl: Option<u8>,
    /// List format override index (1-based index into PlfLfo; 0 = no list)
    pub ilfo: Option<u16>,
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
    /// Footnote entries
    footnotes: Vec<FootnoteEntry>,
    /// Endnote entries
    endnotes: Vec<FootnoteEntry>,
    /// Numbering writer for list tables
    numbering: NumberingWriter,
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
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            numbering: NumberingWriter::new(),
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

    /// Add a paragraph with paragraph formatting (default character formatting)
    pub fn add_formatted_paragraph(
        &mut self,
        text: &str,
        para_fmt: ParagraphFormatting,
    ) -> Result<(), DocWriteError> {
        self.add_paragraph_with_format(text, CharacterFormatting::default(), para_fmt)
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
        let link_fmt = CharacterFormatting {
            underline: Some(true),
            color: Some((0x00, 0x00, 0xFF)),
            ..CharacterFormatting::default()
        };

        // Field begin/separator/end special chars
        let spec_fmt = CharacterFormatting {
            special: Some(true),
            ..CharacterFormatting::default()
        };

        // Field instruction should be hidden (vanished) but not special
        let instr_fmt = CharacterFormatting {
            field_vanish: Some(true),
            ..CharacterFormatting::default()
        };

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

    /// Add a footnote to the document.
    ///
    /// The `ref_position` in `FootnoteEntry` is the character position
    /// in the main document where the footnote reference marker appears.
    pub fn add_footnote(&mut self, entry: FootnoteEntry) {
        self.footnotes.push(entry);
    }

    /// Add an endnote to the document.
    pub fn add_endnote(&mut self, entry: FootnoteEntry) {
        self.endnotes.push(entry);
    }

    /// Add a list structure definition.
    pub fn add_list(&mut self, list: ListStructure) {
        self.numbering.add_list(list);
    }

    /// Add a list format override.
    pub fn add_list_override(&mut self, lfo: ListFormatOverride) {
        self.numbering.add_override(lfo);
    }

    /// Build footnote or endnote subdocument text and PLCFs.
    ///
    /// Per MS-DOC spec:
    /// - Each note text MUST begin with U+0002 (auto-numbered reference mark) with fSpec=1
    /// - PlcffndRef final CP MUST equal `ccp_text` (main document character count)
    /// - PlcffndTxt CPs are relative to the note subdocument start
    ///
    /// `actual_ref_cps`: actual CPs in main doc where U+0002 refs were injected (entry order).
    /// `ccp_text`: FibRgLw97.ccpText — needed for the mandatory final CP in PlcffndRef.
    #[allow(clippy::too_many_arguments)]
    fn build_note_story(
        entries: &[FootnoteEntry],
        actual_ref_cps: &[u32],
        ccp_text: u32,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
    ) -> Option<(Vec<u8>, Vec<u8>, u32)> {
        if entries.is_empty() {
            return None;
        }

        let mut note_cp: u32 = 0;
        // PlcffndTxt: (n+1) CPs relative to note subdocument start
        let mut txt_cps: Vec<u32> = vec![0];

        for entry in entries {
            let fc_para_start = text_fc_start + text_stream.len() as u32;

            // 1) Auto-numbered reference mark U+0002 with fSpec=1 CHPX
            //    This is what Word displays as the footnote number in the note area.
            let fc_ref = fc_para_start;
            text_stream.extend_from_slice(&0x0002u16.to_le_bytes());
            let fc_ref_end = fc_ref + 2;
            let ref_grpprl = build_chpx_grpprl(
                &CharacterFormatting {
                    special: Some(true),
                    ..Default::default()
                },
                font_builder,
            );
            chpx_entries.push((fc_ref, fc_ref_end, ref_grpprl));

            // 2) Note body text
            let text = &entry.text;
            let text_chars = text.chars().count() as u32;
            let fc_text_start = text_fc_start + text_stream.len() as u32;
            for u in text.encode_utf16() {
                text_stream.extend_from_slice(&u.to_le_bytes());
            }
            let fc_text_end = fc_text_start + text_chars * 2;
            let body_grpprl = build_chpx_grpprl(&CharacterFormatting::default(), font_builder);
            chpx_entries.push((fc_text_start, fc_text_end, body_grpprl));

            // 3) Paragraph mark (chEop 0x0D) — extends last CHPX
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            if let Some(last) = chpx_entries.last_mut() {
                last.1 += 2;
            }
            let fc_para_end = text_fc_start + text_stream.len() as u32;

            // PAPX for this note paragraph
            papx_entries.push((
                fc_para_start,
                fc_para_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));

            // Piece: 1 (auto-ref) + text_chars + 1 (para mark)
            let total_chars = 1 + text_chars + 1;
            pieces.push(Piece::new(
                *current_cp_total,
                *current_cp_total + total_chars,
                fc_para_start,
                true,
            ));
            *current_cp_total += total_chars;
            note_cp += total_chars;

            txt_cps.push(note_cp);
        }

        // Trailing guard paragraph mark — mandatory per MS-DOC spec:
        // "The entire footnote subdocument MUST end with a paragraph mark."
        // This is an EXTRA paragraph mark beyond the last footnote's own \r.
        // LibreOffice and POI both write this guard.
        {
            let fc_guard = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_guard_end = fc_guard + 2;
            chpx_entries.push((fc_guard, fc_guard_end, Vec::new()));
            papx_entries.push((
                fc_guard,
                fc_guard_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(
                *current_cp_total,
                *current_cp_total + 1,
                fc_guard,
                true,
            ));
            *current_cp_total += 1;
            note_cp += 1;
            txt_cps.push(note_cp);
        }

        // PlcffndRef: actual reference CPs + mandatory final CP = ccpText
        let mut ref_cps: Vec<u32> = actual_ref_cps.to_vec();
        ref_cps.push(ccp_text);

        // Serialize PlcffndRef: (n+1) CPs then n FRDs (2 bytes each)
        let mut plcf_ref = Vec::with_capacity(ref_cps.len() * 4 + entries.len() * 2);
        for cp in &ref_cps {
            plcf_ref.extend_from_slice(&cp.to_le_bytes());
        }
        // FRD (Footnote Reference Descriptor): nAuto MUST be 0x0000 for
        // auto-numbered references (MS-DOC 2.9.73). Non-zero = custom mark codepoint.
        for _entry in entries {
            plcf_ref.extend_from_slice(&0u16.to_le_bytes());
        }

        // Serialize PlcffndTxt: (n+2) CPs for n footnotes (n stories + 1 guard + 1 final)
        let mut plcf_txt = Vec::with_capacity(txt_cps.len() * 4);
        for cp in &txt_cps {
            plcf_txt.extend_from_slice(&cp.to_le_bytes());
        }

        Some((plcf_ref, plcf_txt, note_cp))
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

        // Build index->text mapping for 12 slots per MS-DOC PlcfHdd / Apache POI:
        //   Slots 0-5:  footnote/endnote separator/continuation stories
        //   Slot 6:     even page header (section 0)
        //   Slot 7:     odd page header (section 0) — "default" when no facing pages
        //   Slot 8:     even page footer (section 0)
        //   Slot 9:     odd page footer (section 0) — "default" when no facing pages
        //   Slot 10:    first page header (section 0)
        //   Slot 11:    first page footer (section 0)
        // PlcfHdd has 13 CPs (12 slot starts + 1 final).
        // Verified against LibreOffice DOC writer output.
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
        // Every slot (0-11) must have at least one paragraph mark.
        // Word uses the CP ranges to locate header/footer stories.
        let mut header_cp: u32 = 0;
        let mut cp_starts: [u32; 12] = [0; 12];

        for i in 0..12 {
            cp_starts[i] = header_cp;
            if let Some(text) = idx_text[i] {
                // Slot has content: write text + paragraph mark + guard paragraph mark
                let fc_para_start = text_fc_start + text_stream.len() as u32;
                let mut para_chars: u32 = 0;

                let char_fmt = CharacterFormatting::default();
                let grpprl = build_chpx_grpprl(&char_fmt, font_builder);
                let run_fc_start = fc_para_start;
                for u in text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                para_chars += text.chars().count() as u32;
                let run_fc_end = run_fc_start + para_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                let current_chpx_idx = chpx_entries.len() - 1;

                // Content paragraph mark
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                chpx_entries[current_chpx_idx].1 += 2;
                let fc_para_end = text_fc_start + text_stream.len() as u32;
                papx_entries.push((
                    fc_para_start,
                    fc_para_end,
                    build_papx_grpprl(&ParagraphFormatting::default()),
                ));

                // Guard paragraph mark (required separator between stories)
                let fc_guard_start = fc_para_end;
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                chpx_entries[current_chpx_idx].1 += 2;
                let fc_guard_end = fc_guard_start + 2;
                papx_entries.push((
                    fc_guard_start,
                    fc_guard_end,
                    build_papx_grpprl(&ParagraphFormatting::default()),
                ));

                // Piece for content + guard
                pieces.push(Piece::new(
                    *current_cp_total,
                    *current_cp_total + para_chars + 2,
                    fc_para_start,
                    true,
                ));
                *current_cp_total += para_chars + 2;
                header_cp += para_chars + 2;
            } else {
                // Empty slot: write a single paragraph mark as guard (MS-DOC requires it)
                let fc_guard = text_fc_start + text_stream.len() as u32;
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                let fc_guard_end = fc_guard + 2;

                // CHPX for the guard mark
                chpx_entries.push((fc_guard, fc_guard_end, Vec::new()));
                // PAPX for the guard mark
                papx_entries.push((
                    fc_guard,
                    fc_guard_end,
                    build_papx_grpprl(&ParagraphFormatting::default()),
                ));

                // Piece for guard
                pieces.push(Piece::new(
                    *current_cp_total,
                    *current_cp_total + 1,
                    fc_guard,
                    true,
                ));
                *current_cp_total += 1;
                header_cp += 1;
            }
        }

        // Build PlcfHdd: 13 CPs (12 slot starts + 1 final end CP)
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
        // Word 2007+ format (nFib 0x0101) requires 1248 bytes
        // (includes cswNew + nFibNew + reserved short at the end)
        let fib_placeholder = vec![0u8; 1248];
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

        // Build sorted list of note references with type tracking.
        // Each entry: (ref_position, is_footnote, original_index_in_vec)
        let mut note_refs: Vec<(u32, bool, usize)> = Vec::new();
        for (idx, entry) in self.footnotes.iter().enumerate() {
            note_refs.push((entry.ref_position, true, idx));
        }
        for (idx, entry) in self.endnotes.iter().enumerate() {
            note_refs.push((entry.ref_position, false, idx));
        }
        note_refs.sort_by_key(|r| r.0);

        // Track field character CPs for PlcfFldMom
        let mut field_char_cps: Vec<(u32, u16)> = Vec::new();

        // Track actual CPs where U+0002 was injected, keyed by (is_footnote, entry_index)
        let mut footnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut endnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut note_inject_idx: usize = 0;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = run_text.chars().count() as u32;
                let grpprl = build_chpx_grpprl(&run.formatting, &mut font_builder);

                // Track field characters in this run
                for (char_offset, ch) in run_text.chars().enumerate() {
                    let cp = current_cp + para_chars + char_offset as u32;
                    match ch as u32 {
                        0x0013 => field_char_cps.push((cp, 0x13)),
                        0x0014 => field_char_cps.push((cp, 0x14)),
                        0x0015 => field_char_cps.push((cp, 0x15)),
                        _ => {},
                    }
                }

                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }

            // Inject U+0002 reference characters for notes whose ref_position
            // falls within this paragraph's CP range
            while note_inject_idx < note_refs.len() {
                let (ref_cp, is_footnote, entry_idx) = note_refs[note_inject_idx];
                if ref_cp <= current_cp + para_chars {
                    let actual_cp = current_cp + para_chars;
                    let fc_ref = text_fc_start + text_stream.len() as u32;
                    text_stream.extend_from_slice(&0x0002u16.to_le_bytes());
                    let fc_ref_end = fc_ref + 2;
                    let ref_grpprl = build_chpx_grpprl(
                        &CharacterFormatting {
                            special: Some(true),
                            ..Default::default()
                        },
                        &mut font_builder,
                    );
                    chpx_entries.push((fc_ref, fc_ref_end, ref_grpprl));
                    para_chars += 1;
                    last_run_index_for_para = Some(chpx_entries.len() - 1);
                    // Record actual CP for PlcffndRef/PlcfendRef
                    if is_footnote {
                        footnote_actual_cps.push((entry_idx, actual_cp));
                    } else {
                        endnote_actual_cps.push((entry_idx, actual_cp));
                    }
                    note_inject_idx += 1;
                } else {
                    break;
                }
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

        let text_length = current_cp;

        // Sort actual CPs by entry index to match footnote/endnote entry order
        footnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        endnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        let ftn_ref_cps: Vec<u32> = footnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let edn_ref_cps: Vec<u32> = endnote_actual_cps.iter().map(|&(_, cp)| cp).collect();

        // Subdocument order: main text → footnotes → headers/footers → endnotes
        let footnote_plcfs = Self::build_note_story(
            &self.footnotes,
            &ftn_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        );

        // Build header/footer story
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

        // Build endnote story (appends endnote text after headers)
        let endnote_plcfs = Self::build_note_story(
            &self.endnotes,
            &edn_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        );

        // Mandatory trailing paragraph mark when ANY subdocument exists.
        // Per MS-DOC spec: "The total number of character positions is
        // ccpText + ccpFtn + ccpHdd + ... + 1 if any of ccpFtn, ccpHdd, etc. are nonzero."
        // This extra character MUST be present; Word uses it as a sentinel.
        let has_subdocs =
            footnote_plcfs.is_some() || header_plcfhdd.is_some() || endnote_plcfs.is_some();
        if has_subdocs {
            let fc_trailing = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_trailing_end = fc_trailing + 2;
            chpx_entries.push((fc_trailing, fc_trailing_end, Vec::new()));
            papx_entries.push((
                fc_trailing,
                fc_trailing_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(current_cp, current_cp + 1, fc_trailing, true));
            current_cp += 1;
        }

        // Initialize FIB builder
        let mut fib = FibBuilder::new();
        fib.set_main_text(0, text_length);
        if let Some((_, _, ftn_cp)) = &footnote_plcfs {
            fib.set_ccp_ftn(*ftn_cp);
        }
        if let Some((_, header_cp)) = &header_plcfhdd {
            fib.set_ccp_hdd(*header_cp);
        }
        if let Some((_, _, edn_cp)) = &endnote_plcfs {
            fib.set_ccp_edn(*edn_cp);
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

        // Write footnote PLCFs if present
        if let Some((ref_bytes, txt_bytes, _)) = &footnote_plcfs {
            fib.set_plcffnd_ref(table_offset, ref_bytes.len() as u32);
            table_stream.extend_from_slice(ref_bytes);
            table_offset = table_stream.len() as u32;

            fib.set_plcffnd_txt(table_offset, txt_bytes.len() as u32);
            table_stream.extend_from_slice(txt_bytes);
            table_offset = table_stream.len() as u32;
        }

        // Write endnote PLCFs if present
        if let Some((ref_bytes, txt_bytes, _)) = &endnote_plcfs {
            fib.set_plcfend_ref(table_offset, ref_bytes.len() as u32);
            table_stream.extend_from_slice(ref_bytes);
            table_offset = table_stream.len() as u32;

            fib.set_plcfend_txt(table_offset, txt_bytes.len() as u32);
            table_stream.extend_from_slice(txt_bytes);
            table_offset = table_stream.len() as u32;
        }

        // Write PlcfFldMom (main document field table) if there are field characters
        // Structure: (n+1) CPs + n FLD descriptors (2 bytes each)
        // FLD descriptor per MS-DOC 2.8.25:
        //   fldBegin (0x13): byte0 = 0x13, byte1 = flt (field type: 0x58 = HYPERLINK)
        //   fldSep   (0x14): byte0 = 0x14, byte1 = flags (0x00)
        //   fldEnd   (0x15): byte0 = 0x15, byte1 = flags (0x00)
        // Final CP MUST equal ccpText per MS-DOC spec.
        if !field_char_cps.is_empty() {
            let n = field_char_cps.len();
            let mut plcffld = Vec::with_capacity((n + 1) * 4 + n * 2);
            for (cp, _) in &field_char_cps {
                plcffld.extend_from_slice(&cp.to_le_bytes());
            }
            // Final CP = ccpText (per MS-DOC spec PlcfFld)
            plcffld.extend_from_slice(&text_length.to_le_bytes());
            // FLD descriptors
            for (_, fld_type) in &field_char_cps {
                let (fldch, flt_or_flags) = match *fld_type {
                    0x13 => (0x13u8, 0x58u8), // fldBegin, flt = HYPERLINK (88)
                    0x14 => (0x14u8, 0x00u8), // fldSep, no flags
                    0x15 => (0x15u8, 0x00u8), // fldEnd, no flags
                    _ => (0x00, 0x00),
                };
                plcffld.push(fldch);
                plcffld.push(flt_or_flags);
            }
            fib.set_plcffld_mom(table_offset, plcffld.len() as u32);
            table_stream.extend_from_slice(&plcffld);
            table_offset = table_stream.len() as u32;
        }

        // Write numbering tables (PlfLst / PlfLfo) if present
        if !self.numbering.is_empty() {
            // PlfLst: lcbPlfLst covers only cLst + LSTF array.
            // LVL data is appended immediately after but NOT counted in lcbPlfLst
            // per MS-DOC spec and Apache POI ListTables.writeListDataTo().
            let (plflst_header, lvl_data) = self.numbering.build_plflst();
            fib.set_plflst(table_offset, plflst_header.len() as u32);
            table_stream.extend_from_slice(&plflst_header);
            table_stream.extend_from_slice(&lvl_data);
            table_offset = table_stream.len() as u32;

            let plflfo = self.numbering.build_plflfo();
            fib.set_plflfo(table_offset, plflfo.len() as u32);
            table_stream.extend_from_slice(&plflfo);
            table_offset = table_stream.len() as u32;
        }

        // 6-8. Bin tables and section table are written AFTER FKPs
        // (we need FKP page numbers first).
        // Record current table_offset; bin tables will be appended later.

        // 9. Write Font Table to table stream (MANDATORY - POI line 899-903)
        let font_table = font_builder.generate();
        fib.set_sttbfffn(table_offset, font_table.len() as u32);
        table_stream.extend_from_slice(&font_table);

        // 10. Append text (main + headers/footers) to WordDocument stream
        word_document_stream.extend_from_slice(&text_stream);

        // Capture fcMac AFTER text, BEFORE FKPs (POI line 703)
        let fc_mac_value = word_document_stream.len() as u32;
        // text_fc_end is computed inside FKP page ranges; no longer needed here

        // 10a. Write FKPs to WordDocument stream (CRITICAL - POI line 450-492)
        // FKPs must start at 512-byte aligned offsets
        // Pad to 512-byte boundary
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        // ── CHPX FKPs (multi-page) ──
        let chpx_first_page = (word_document_stream.len() / 512) as u32;
        let mut chpx_builder = crate::ole::doc::writer::fkp::ChpxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_pages = chpx_builder.generate_pages()?;
        for page in &chpx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── PAPX FKPs (multi-page) ──
        let papx_first_page = (word_document_stream.len() / 512) as u32;
        let mut papx_builder = crate::ole::doc::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_pages = papx_builder.generate_pages()?;
        for page in &papx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── Write bin tables to table stream (now that we know page numbers) ──
        let chpx_bin_table = crate::ole::doc::writer::bin_table::generate_bin_table_from_pages(
            &chpx_pages.ranges,
            chpx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_chpx(table_offset, chpx_bin_table.len() as u32);
        table_stream.extend_from_slice(&chpx_bin_table);

        let papx_bin_table = crate::ole::doc::writer::bin_table::generate_bin_table_from_pages(
            &papx_pages.ranges,
            papx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_papx(table_offset, papx_bin_table.len() as u32);
        table_stream.extend_from_slice(&papx_bin_table);

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

        // 10c. Write section table to table stream with correct SEPX offset
        // Section table CP must span ALL subdocuments (main + footnotes + headers + endnotes),
        // not just ccpText. Per MS-DOC spec and Apache POI SectionTable.
        let total_cp = current_cp;
        let section_table =
            crate::ole::doc::writer::section::generate_section_table(total_cp, sepx_offset);
        table_offset = table_stream.len() as u32;
        fib.set_plcfsed(table_offset, section_table.len() as u32);
        table_stream.extend_from_slice(&section_table);

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

        // WordDocument stream FIRST to guarantee sector 0, then 1Table, then Data
        ole_writer.create_stream(&["WordDocument"], &word_document_stream)?;
        ole_writer.create_stream(&["1Table"], &table_stream)?;

        // Data stream (MANDATORY per POI - even if empty, padded to 4096)
        let data_stream = vec![0u8; 4096];
        ole_writer.create_stream(&["Data"], &data_stream)?;

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

        // Reserve space for FIB (Word 2007+ format = 1248 bytes, includes cswNew)
        let fib_placeholder = vec![0u8; 1248];
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
        let fc_min: u32 = text_fc_start;

        // Build sorted list of note references with type tracking
        let mut note_refs: Vec<(u32, bool, usize)> = Vec::new();
        for (idx, entry) in self.footnotes.iter().enumerate() {
            note_refs.push((entry.ref_position, true, idx));
        }
        for (idx, entry) in self.endnotes.iter().enumerate() {
            note_refs.push((entry.ref_position, false, idx));
        }
        note_refs.sort_by_key(|r| r.0);

        let mut field_char_cps: Vec<(u32, u16)> = Vec::new();
        let mut footnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut endnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut note_inject_idx: usize = 0;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = run_text.chars().count() as u32;
                let grpprl = build_chpx_grpprl(&run.formatting, &mut font_builder);

                for (char_offset, ch) in run_text.chars().enumerate() {
                    let cp = current_cp + para_chars + char_offset as u32;
                    match ch as u32 {
                        0x0013 => field_char_cps.push((cp, 0x13)),
                        0x0014 => field_char_cps.push((cp, 0x14)),
                        0x0015 => field_char_cps.push((cp, 0x15)),
                        _ => {},
                    }
                }

                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }

            while note_inject_idx < note_refs.len() {
                let (ref_cp, is_footnote, entry_idx) = note_refs[note_inject_idx];
                if ref_cp <= current_cp + para_chars {
                    let actual_cp = current_cp + para_chars;
                    let fc_ref = text_fc_start + text_stream.len() as u32;
                    text_stream.extend_from_slice(&0x0002u16.to_le_bytes());
                    let fc_ref_end = fc_ref + 2;
                    let ref_grpprl = build_chpx_grpprl(
                        &CharacterFormatting {
                            special: Some(true),
                            ..Default::default()
                        },
                        &mut font_builder,
                    );
                    chpx_entries.push((fc_ref, fc_ref_end, ref_grpprl));
                    para_chars += 1;
                    last_run_index_for_para = Some(chpx_entries.len() - 1);
                    if is_footnote {
                        footnote_actual_cps.push((entry_idx, actual_cp));
                    } else {
                        endnote_actual_cps.push((entry_idx, actual_cp));
                    }
                    note_inject_idx += 1;
                } else {
                    break;
                }
            }

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

        let text_length = current_cp;

        footnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        endnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        let ftn_ref_cps: Vec<u32> = footnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let edn_ref_cps: Vec<u32> = endnote_actual_cps.iter().map(|&(_, cp)| cp).collect();

        let footnote_plcfs = Self::build_note_story(
            &self.footnotes,
            &ftn_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        );

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

        let endnote_plcfs = Self::build_note_story(
            &self.endnotes,
            &edn_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        );

        // Mandatory trailing paragraph mark when ANY subdocument exists (same as save()).
        let has_subdocs =
            footnote_plcfs.is_some() || header_plcfhdd.is_some() || endnote_plcfs.is_some();
        if has_subdocs {
            let fc_trailing = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_trailing_end = fc_trailing + 2;
            chpx_entries.push((fc_trailing, fc_trailing_end, Vec::new()));
            papx_entries.push((
                fc_trailing,
                fc_trailing_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(current_cp, current_cp + 1, fc_trailing, true));
            current_cp += 1;
        }

        let mut fib = FibBuilder::new();
        fib.set_main_text(0, text_length);
        if let Some((_, _, ftn_cp)) = &footnote_plcfs {
            fib.set_ccp_ftn(*ftn_cp);
        }
        if let Some((_, header_cp)) = &header_plcfhdd {
            fib.set_ccp_hdd(*header_cp);
        }
        if let Some((_, _, edn_cp)) = &endnote_plcfs {
            fib.set_ccp_edn(*edn_cp);
        }

        let mut table_offset = 0u32;

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

        // DocumentProperties
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

        // Write PlcfHdd if present
        if let Some((plcf_bytes, _header_cp)) = &header_plcfhdd {
            fib.set_plcfhdd(table_offset, plcf_bytes.len() as u32);
            table_stream.extend_from_slice(plcf_bytes);
            table_offset = table_stream.len() as u32;
        }

        // Write footnote PLCFs if present
        if let Some((ref_bytes, txt_bytes, _)) = &footnote_plcfs {
            fib.set_plcffnd_ref(table_offset, ref_bytes.len() as u32);
            table_stream.extend_from_slice(ref_bytes);
            table_offset = table_stream.len() as u32;

            fib.set_plcffnd_txt(table_offset, txt_bytes.len() as u32);
            table_stream.extend_from_slice(txt_bytes);
            table_offset = table_stream.len() as u32;
        }

        // Write endnote PLCFs if present
        if let Some((ref_bytes, txt_bytes, _)) = &endnote_plcfs {
            fib.set_plcfend_ref(table_offset, ref_bytes.len() as u32);
            table_stream.extend_from_slice(ref_bytes);
            table_offset = table_stream.len() as u32;

            fib.set_plcfend_txt(table_offset, txt_bytes.len() as u32);
            table_stream.extend_from_slice(txt_bytes);
            table_offset = table_stream.len() as u32;
        }

        // Write PlcfFldMom if there are field characters
        if !field_char_cps.is_empty() {
            let n = field_char_cps.len();
            let mut plcffld = Vec::with_capacity((n + 1) * 4 + n * 2);
            for (cp, _) in &field_char_cps {
                plcffld.extend_from_slice(&cp.to_le_bytes());
            }
            // Final CP = ccpText (per MS-DOC spec PlcfFld)
            plcffld.extend_from_slice(&text_length.to_le_bytes());
            for (_, fld_type) in &field_char_cps {
                let (fldch, flt_or_flags) = match *fld_type {
                    0x13 => (0x13u8, 0x58u8), // fldBegin, flt = HYPERLINK (88)
                    0x14 => (0x14u8, 0x00u8), // fldSep, no flags
                    0x15 => (0x15u8, 0x00u8), // fldEnd, no flags
                    _ => (0x00, 0x00),
                };
                plcffld.push(fldch);
                plcffld.push(flt_or_flags);
            }
            fib.set_plcffld_mom(table_offset, plcffld.len() as u32);
            table_stream.extend_from_slice(&plcffld);
            table_offset = table_stream.len() as u32;
        }

        // Write numbering tables if present
        if !self.numbering.is_empty() {
            let (plflst_header, lvl_data) = self.numbering.build_plflst();
            fib.set_plflst(table_offset, plflst_header.len() as u32);
            table_stream.extend_from_slice(&plflst_header);
            table_stream.extend_from_slice(&lvl_data);
            table_offset = table_stream.len() as u32;

            let plflfo = self.numbering.build_plflfo();
            fib.set_plflfo(table_offset, plflfo.len() as u32);
            table_stream.extend_from_slice(&plflfo);
            table_offset = table_stream.len() as u32;
        }

        // 6-8. Bin tables and section table written AFTER FKPs (need page numbers).

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

        // ── CHPX FKPs (multi-page) ──
        let chpx_first_page = (word_document_stream.len() / 512) as u32;
        let mut chpx_builder = crate::ole::doc::writer::fkp::ChpxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_pages = chpx_builder.generate_pages()?;
        for page in &chpx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── PAPX FKPs (multi-page) ──
        let papx_first_page = (word_document_stream.len() / 512) as u32;
        let mut papx_builder = crate::ole::doc::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_pages = papx_builder.generate_pages()?;
        for page in &papx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── Write bin tables to table stream ──
        let chpx_bin_table = crate::ole::doc::writer::bin_table::generate_bin_table_from_pages(
            &chpx_pages.ranges,
            chpx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_chpx(table_offset, chpx_bin_table.len() as u32);
        table_stream.extend_from_slice(&chpx_bin_table);

        let papx_bin_table = crate::ole::doc::writer::bin_table::generate_bin_table_from_pages(
            &papx_pages.ranges,
            papx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_papx(table_offset, papx_bin_table.len() as u32);
        table_stream.extend_from_slice(&papx_bin_table);

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

        // Write section table to table stream
        let total_cp = current_cp;
        let section_table =
            crate::ole::doc::writer::section::generate_section_table(total_cp, sepx_offset);
        table_offset = table_stream.len() as u32;
        fib.set_plcfsed(table_offset, section_table.len() as u32);
        table_stream.extend_from_slice(&section_table);

        // Set FibBase fields
        let cb_mac = word_document_stream.len() as u32;
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

        // Ensure WordDocument gets sector 0: add it first, then 1Table, then Data
        ole_writer.create_stream(&["WordDocument"], &word_document_stream)?;
        ole_writer.create_stream(&["1Table"], &table_stream)?;

        // Data stream (MANDATORY per POI - even if empty, padded to 4096)
        let data_stream = vec![0u8; 4096];
        ole_writer.create_stream(&["Data"], &data_stream)?;

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

    // List numbering: ilvl (list level) and ilfo (list format override)
    if let Some(ilvl) = fmt.ilvl {
        push_byte(&mut grp, SPRM_P_ILVL, ilvl);
    }
    if let Some(ilfo) = fmt.ilfo {
        push_u16(&mut grp, SPRM_P_ILFO, ilfo);
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
    use std::io::Cursor;

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
    fn test_add_multiple_paragraphs() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("First paragraph").unwrap();
        writer.add_paragraph("Second paragraph").unwrap();
        writer.add_paragraph("Third paragraph").unwrap();
        assert_eq!(writer.paragraphs.len(), 3);
        assert_eq!(writer.paragraphs[0].runs[0].text, "First paragraph");
        assert_eq!(writer.paragraphs[1].runs[0].text, "Second paragraph");
        assert_eq!(writer.paragraphs[2].runs[0].text, "Third paragraph");
    }

    #[test]
    fn test_add_formatted_paragraph() {
        let mut writer = DocWriter::new();
        let para_fmt = ParagraphFormatting {
            alignment: Some(1), // Center
            space_before: Some(240),
            space_after: Some(120),
            ..Default::default()
        };
        writer
            .add_formatted_paragraph("Formatted text", para_fmt)
            .unwrap();
        assert_eq!(writer.paragraphs.len(), 1);
        assert_eq!(writer.paragraphs[0].runs[0].text, "Formatted text");
        assert_eq!(writer.paragraphs[0].formatting.alignment, Some(1));
    }

    #[test]
    fn test_add_paragraph_with_character_formatting() {
        let mut writer = DocWriter::new();
        let char_fmt = CharacterFormatting {
            bold: Some(true),
            italic: Some(true),
            font_size: Some(24),
            ..Default::default()
        };
        let para_fmt = ParagraphFormatting::default();
        writer
            .add_paragraph_with_format("Bold italic text", char_fmt, para_fmt)
            .unwrap();
        assert_eq!(writer.paragraphs.len(), 1);
        assert_eq!(writer.paragraphs[0].runs[0].text, "Bold italic text");
        assert_eq!(writer.paragraphs[0].runs[0].formatting.bold, Some(true));
        assert_eq!(writer.paragraphs[0].runs[0].formatting.italic, Some(true));
        assert_eq!(writer.paragraphs[0].runs[0].formatting.font_size, Some(24));
    }

    #[test]
    fn test_add_paragraph_runs() {
        let mut writer = DocWriter::new();
        let runs = vec![
            (
                "Bold ".to_string(),
                CharacterFormatting {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            (
                "Italic".to_string(),
                CharacterFormatting {
                    italic: Some(true),
                    ..Default::default()
                },
            ),
        ];
        writer
            .add_paragraph_runs(runs, ParagraphFormatting::default())
            .unwrap();
        assert_eq!(writer.paragraphs.len(), 1);
        assert_eq!(writer.paragraphs[0].runs.len(), 2);
        assert_eq!(writer.paragraphs[0].runs[0].text, "Bold ");
        assert_eq!(writer.paragraphs[0].runs[1].text, "Italic");
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

    #[test]
    fn test_set_table_cell_multiple() {
        let mut writer = DocWriter::new();
        let idx = writer.add_table(2, 2).unwrap();
        writer.set_table_cell_text(idx, 0, 0, "A").unwrap();
        writer.set_table_cell_text(idx, 0, 1, "B").unwrap();
        writer.set_table_cell_text(idx, 1, 0, "C").unwrap();
        writer.set_table_cell_text(idx, 1, 1, "D").unwrap();
        assert_eq!(
            writer.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text,
            "A"
        );
        assert_eq!(
            writer.tables[0].rows[0].cells[1].paragraphs[0].runs[0].text,
            "B"
        );
        assert_eq!(
            writer.tables[0].rows[1].cells[0].paragraphs[0].runs[0].text,
            "C"
        );
        assert_eq!(
            writer.tables[0].rows[1].cells[1].paragraphs[0].runs[0].text,
            "D"
        );
    }

    #[test]
    fn test_set_property() {
        let mut writer = DocWriter::new();
        writer.set_property("Title", "Test Document");
        writer.set_property("Author", "Test Author");
        assert_eq!(
            writer.properties.get("Title"),
            Some(&"Test Document".to_string())
        );
        assert_eq!(
            writer.properties.get("Author"),
            Some(&"Test Author".to_string())
        );
    }

    #[test]
    fn test_headers_and_footers() {
        let mut writer = DocWriter::new();
        writer.set_odd_header("Odd Header");
        writer.set_even_header("Even Header");
        writer.set_first_header("First Header");
        writer.set_odd_footer("Odd Footer");
        writer.set_even_footer("Even Footer");
        writer.set_first_footer("First Footer");
        assert_eq!(writer.header_odd, Some("Odd Header".to_string()));
        assert_eq!(writer.header_even, Some("Even Header".to_string()));
        assert_eq!(writer.header_first, Some("First Header".to_string()));
        assert_eq!(writer.footer_odd, Some("Odd Footer".to_string()));
        assert_eq!(writer.footer_even, Some("Even Footer".to_string()));
        assert_eq!(writer.footer_first, Some("First Footer".to_string()));
    }

    #[test]
    fn test_footnotes() {
        let mut writer = DocWriter::new();
        let entry = FootnoteEntry::new(0u32, "This is a footnote", 1u16);
        writer.add_footnote(entry);
        assert_eq!(writer.footnotes.len(), 1);
        assert_eq!(writer.footnotes[0].text, "This is a footnote");
    }

    #[test]
    fn test_endnotes() {
        let mut writer = DocWriter::new();
        let entry = FootnoteEntry::new(0u32, "This is an endnote", 1u16);
        writer.add_endnote(entry);
        assert_eq!(writer.endnotes.len(), 1);
        assert_eq!(writer.endnotes[0].text, "This is an endnote");
    }

    #[test]
    fn test_write_to_memory() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("Test paragraph").unwrap();
        let mut cursor = Cursor::new(Vec::new());
        let result = writer.write_to(&mut cursor);
        assert!(result.is_ok());
        assert!(cursor.into_inner().len() > 0);
    }

    #[test]
    fn test_empty_document_write() {
        let mut writer = DocWriter::new();
        let mut cursor = Cursor::new(Vec::new());
        let result = writer.write_to(&mut cursor);
        assert!(result.is_ok());
        let data = cursor.into_inner();
        assert!(data.len() > 0);
    }

    #[test]
    fn test_character_formatting_default() {
        let fmt = CharacterFormatting::default();
        assert!(fmt.bold.is_none());
        assert!(fmt.italic.is_none());
        assert!(fmt.underline.is_none());
        assert!(fmt.font_size.is_none());
    }

    #[test]
    fn test_paragraph_formatting_default() {
        let fmt = ParagraphFormatting::default();
        assert!(fmt.alignment.is_none());
        assert!(fmt.left_indent.is_none());
        assert!(fmt.right_indent.is_none());
        assert!(fmt.space_before.is_none());
        assert!(fmt.space_after.is_none());
    }

    #[test]
    fn test_line_spacing_default() {
        let ls = LineSpacing::default();
        assert_eq!(ls.dya_line, 0);
        assert!(!ls.is_multiple);
    }

    #[test]
    fn test_add_table_invalid_dimensions() {
        let mut writer = DocWriter::new();
        assert!(writer.add_table(0, 3).is_err());
        assert!(writer.add_table(2, 0).is_err());
        assert!(writer.add_table(0, 0).is_err());
    }

    #[test]
    fn test_set_table_cell_invalid_indices() {
        let mut writer = DocWriter::new();
        let idx = writer.add_table(2, 2).unwrap();
        assert!(writer.set_table_cell_text(idx, 2, 0, "Invalid").is_err());
        assert!(writer.set_table_cell_text(idx, 0, 2, "Invalid").is_err());
        assert!(writer.set_table_cell_text(999, 0, 0, "Invalid").is_err());
    }
}
