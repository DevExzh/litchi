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
//! use litchi_ole::doc::DocWriter;
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

use super::bookmarks::BookmarkEntry;
use super::comments::CommentEntry;
use super::fib::FibBuilder;
use super::font_table::FontTableBuilder;
use super::footnotes::FootnoteEntry;
use super::numbering::{ListFormatOverride, ListStructure, NumberingWriter};
use super::piece_table::{Piece, PieceTableBuilder};
use super::revisions::{DisplayFieldRevision, FormattingRevision, NumberingRevision, TextRevision};
use crate::doc::CommentDateTime;
use crate::sprm_operations::*;
use litchi_cfb::writer::OleWriter;
use std::collections::HashMap;

/// Error type for DOC writing
#[derive(Debug)]
pub enum DocWriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(crate::OleError),
}

impl From<std::io::Error> for DocWriteError {
    fn from(err: std::io::Error) -> Self {
        DocWriteError::Io(err)
    }
}

impl From<crate::OleError> for DocWriteError {
    fn from(err: crate::OleError) -> Self {
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

fn utf16_code_unit_len(text: &str) -> Result<u32, DocWriteError> {
    let length = u32::try_from(text.encode_utf16().count()).map_err(|_| {
        DocWriteError::InvalidData("DOC text exceeds the 32-bit CP range".to_string())
    })?;
    if length >= 0x7FFF_FFFF {
        return Err(DocWriteError::InvalidData(
            "DOC text exceeds the MS-DOC CP limit".to_string(),
        ));
    }
    Ok(length)
}

fn pack_dttm(value: Option<CommentDateTime>) -> Result<u32, DocWriteError> {
    let Some(value) = value else {
        return Ok(0);
    };
    if !(1900..=2411).contains(&value.year)
        || !(1..=12).contains(&value.month)
        || !(1..=31).contains(&value.day)
        || value.hour > 23
        || value.minute > 59
        || value.weekday > 6
    {
        return Err(DocWriteError::InvalidData(
            "DOC timestamp is outside the DTTM field ranges".to_string(),
        ));
    }
    Ok(u32::from(value.minute)
        | (u32::from(value.hour) << 6)
        | (u32::from(value.day) << 11)
        | (u32::from(value.month) << 16)
        | (u32::from(value.year - 1900) << 20)
        | (u32::from(value.weekday) << 29))
}

type NoteStoryData = (Vec<u8>, Vec<u8>, u32);
type HeaderStoryData = (Vec<u8>, u32);

struct CommentStoryData {
    owners: Vec<u8>,
    references: Vec<u8>,
    text_positions: Vec<u8>,
    bookmark_names: Vec<u8>,
    bookmark_starts: Vec<u8>,
    bookmark_ends: Vec<u8>,
    extended_metadata: Vec<u8>,
    char_count: u32,
}

struct BookmarkTableData {
    names: Vec<u8>,
    starts: Vec<u8>,
    ends: Vec<u8>,
}

struct RevisionWriterData {
    indexes: HashMap<String, u16>,
    table: Vec<u8>,
}

#[derive(Clone, Copy)]
enum MainReferenceKind {
    Footnote,
    Endnote,
    Comment,
}

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
    /// Mark this run as inserted text.
    pub insertion_revision: Option<TextRevision>,
    /// Mark this run as deleted text.
    pub deletion_revision: Option<TextRevision>,
    /// Mark the run's character formatting as a tracked change.
    pub formatting_revision: Option<FormattingRevision>,
    /// Mark a LISTNUM display-field result as revised.
    pub display_field_revision: Option<DisplayFieldRevision>,
    // Future enhancement: Additional properties (color, strikethrough, subscript, superscript, etc.)
}

/// Line spacing descriptor for paragraphs, equivalent to POI's LineSpacingDescriptor (LSPD).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LineSpacing {
    /// Line height. If `is_multiple` is false, value is in twips. If true, value is in 240ths of a line.
    pub dya_line: i16,
    /// Whether `dya_line` is a multiple of single line (value is 240ths of a line) instead of twips.
    pub is_multiple: bool,
}

impl LineSpacing {
    /// Single-line spacing (240/240 of a line).
    pub const fn single() -> Self {
        Self {
            dya_line: 240,
            is_multiple: true,
        }
    }

    /// One-and-a-half-line spacing (360/240 of a line).
    pub const fn one_and_half() -> Self {
        Self {
            dya_line: 360,
            is_multiple: true,
        }
    }

    /// Double-line spacing (480/240 of a line).
    pub const fn double() -> Self {
        Self {
            dya_line: 480,
            is_multiple: true,
        }
    }

    /// Create proportional line spacing expressed in 240ths of one line.
    pub fn multiple_240ths(value: u16) -> Result<Self, DocWriteError> {
        let dya_line = i16::try_from(value).map_err(|_| {
            DocWriteError::InvalidData(format!(
                "line-spacing multiple {value} exceeds the signed LSPD range"
            ))
        })?;
        if dya_line == 0 {
            return Err(DocWriteError::InvalidData(
                "line-spacing multiple must be greater than zero".into(),
            ));
        }
        Ok(Self {
            dya_line,
            is_multiple: true,
        })
    }

    /// Create minimum line spacing in twips.
    pub fn at_least_twips(value: u16) -> Result<Self, DocWriteError> {
        let dya_line = i16::try_from(value).map_err(|_| {
            DocWriteError::InvalidData(format!(
                "minimum line spacing {value} twips exceeds the signed LSPD range"
            ))
        })?;
        if dya_line == 0 {
            return Err(DocWriteError::InvalidData(
                "minimum line spacing must be greater than zero".into(),
            ));
        }
        Ok(Self {
            dya_line,
            is_multiple: false,
        })
    }

    /// Create exact line spacing in twips.
    pub fn exact_twips(value: u16) -> Result<Self, DocWriteError> {
        if value == 0 || value > 32_768 {
            return Err(DocWriteError::InvalidData(format!(
                "exact line spacing {value} twips is outside the LSPD range 1..=32768"
            )));
        }
        Ok(Self {
            dya_line: -(i32::from(value)) as i16,
            is_multiple: false,
        })
    }
}

impl Default for LineSpacing {
    fn default() -> Self {
        Self::single()
    }
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
    /// Mark the paragraph formatting as a tracked change.
    pub formatting_revision: Option<FormattingRevision>,
    /// Whether a numbered list was applied after the previous revision.
    pub numbering_revision_list_applied: Option<bool>,
    /// Retained numbering state for a tracked numbering change.
    pub numbering_revision: Option<NumberingRevision>,
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

fn writable_paragraph_from_runs(
    runs: Vec<(String, CharacterFormatting)>,
    formatting: ParagraphFormatting,
) -> WritableParagraph {
    let runs = if runs.is_empty() {
        vec![TextRun {
            text: String::new(),
            formatting: CharacterFormatting::default(),
        }]
    } else {
        runs.into_iter()
            .map(|(text, formatting)| TextRun { text, formatting })
            .collect()
    };
    WritableParagraph { runs, formatting }
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
    /// Row and cell layout encoded in the row mark's TAP properties.
    formatting: super::tap::TableRow,
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
    /// Comments
    comments: Vec<CommentEntry>,
    /// Standard bookmarks
    bookmarks: Vec<BookmarkEntry>,
    /// Property revision metadata for the writer's single document section
    section_formatting_revision: Option<FormattingRevision>,
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
            comments: Vec::new(),
            bookmarks: Vec::new(),
            section_formatting_revision: None,
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

    /// Add a point or ranged comment to the document.
    pub fn add_comment(&mut self, entry: CommentEntry) {
        self.comments.push(entry);
    }

    /// Add a standard bookmark to the document.
    pub fn add_bookmark(&mut self, entry: BookmarkEntry) {
        self.bookmarks.push(entry);
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
    ) -> Result<Option<NoteStoryData>, DocWriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() != actual_ref_cps.len() {
            return Err(DocWriteError::InvalidData(
                "every DOC note must have a reference in the main document".to_string(),
            ));
        }

        let mut ordered = entries
            .iter()
            .zip(actual_ref_cps.iter().copied())
            .collect::<Vec<_>>();
        ordered.sort_by_key(|(_, cp)| *cp);
        if ordered.windows(2).any(|pair| pair[0].1 == pair[1].1) {
            return Err(DocWriteError::InvalidData(
                "DOC note references must have unique character positions".to_string(),
            ));
        }
        if ordered.iter().any(|(_, cp)| *cp >= ccp_text) {
            return Err(DocWriteError::InvalidData(
                "DOC note reference lies outside the main document".to_string(),
            ));
        }

        let mut note_cp: u32 = 0;
        // PlcffndTxt: n story starts, one story terminator, and one ignored final CP.
        let mut txt_cps: Vec<u32> = vec![0];

        for (entry, _) in &ordered {
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
            let text_chars = utf16_code_unit_len(text)?;
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
        let mut ref_cps = ordered.iter().map(|(_, cp)| *cp).collect::<Vec<_>>();
        ref_cps.push(ccp_text);

        // Serialize PlcffndRef: (n+1) CPs then n FRDs (2 bytes each)
        let mut plcf_ref = Vec::with_capacity(ref_cps.len() * 4 + entries.len() * 2);
        for cp in &ref_cps {
            plcf_ref.extend_from_slice(&cp.to_le_bytes());
        }
        // FRD nAuto is nonzero for an automatically numbered note.
        for (entry, _) in &ordered {
            plcf_ref.extend_from_slice(&entry.number.max(1).to_le_bytes());
        }

        // Serialize PlcffndTxt: (n+2) CPs for n footnotes (n stories + 1 guard + 1 final)
        let mut plcf_txt = Vec::with_capacity(txt_cps.len() * 4);
        for cp in &txt_cps {
            plcf_txt.extend_from_slice(&cp.to_le_bytes());
        }

        Ok(Some((plcf_ref, plcf_txt, note_cp)))
    }

    /// Append the comment subdocument and build its owner, reference, and text tables.
    #[allow(clippy::too_many_arguments)]
    fn build_comment_story(
        entries: &[CommentEntry],
        actual_ref_cps: &[u32],
        ccp_text: u32,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
    ) -> Result<Option<CommentStoryData>, DocWriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() != actual_ref_cps.len() {
            return Err(DocWriteError::InvalidData(
                "every DOC comment must have a reference in the main document".to_string(),
            ));
        }

        let mut ordered = entries
            .iter()
            .zip(actual_ref_cps.iter().copied())
            .collect::<Vec<_>>();
        ordered.sort_by_key(|(_, cp)| *cp);
        if ordered.windows(2).any(|pair| pair[0].1 == pair[1].1) {
            return Err(DocWriteError::InvalidData(
                "DOC comment references must have unique character positions".to_string(),
            ));
        }
        if ordered.iter().any(|(_, cp)| *cp >= ccp_text) {
            return Err(DocWriteError::InvalidData(
                "DOC comment reference lies outside the main document".to_string(),
            ));
        }

        let mut owners = Vec::<String>::new();
        let mut owner_indexes = Vec::with_capacity(ordered.len());
        for (entry, _) in &ordered {
            let author_len = entry.author.encode_utf16().count();
            if author_len >= 56 {
                return Err(DocWriteError::InvalidData(
                    "DOC comment author names must contain fewer than 56 UTF-16 code units"
                        .to_string(),
                ));
            }
            let initials_len = entry.initials.encode_utf16().count();
            if initials_len > 9 {
                return Err(DocWriteError::InvalidData(
                    "DOC comment initials must contain at most nine UTF-16 code units".to_string(),
                ));
            }
            let index = if let Some(index) = owners.iter().position(|owner| owner == &entry.author)
            {
                index
            } else {
                if owners.len() >= 0x7FFF {
                    return Err(DocWriteError::InvalidData(
                        "DOC comment owner array exceeds 0x7FFF entries".to_string(),
                    ));
                }
                owners.push(entry.author.clone());
                owners.len() - 1
            };
            owner_indexes.push(index as u16);
        }

        let mut owner_bytes = Vec::new();
        for owner in &owners {
            let units = owner.encode_utf16().collect::<Vec<_>>();
            owner_bytes.extend_from_slice(&(units.len() as u16).to_le_bytes());
            owner_bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }

        let ranged_count = ordered
            .iter()
            .filter(|(entry, _)| entry.range.is_some())
            .count();
        if ranged_count > 0x3FFC {
            return Err(DocWriteError::InvalidData(
                "DOC annotation bookmark table exceeds 0x3FFC entries".to_string(),
            ));
        }
        let bookmark_sentinel = ccp_text.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC annotation bookmark sentinel overflows".to_string())
        })?;
        let mut bookmark_tags = vec![None; ordered.len()];
        let mut ranges = Vec::<(u32, u32, u32)>::with_capacity(ranged_count);
        for (index, (entry, _)) in ordered.iter().enumerate() {
            let Some((start, end)) = entry.range else {
                continue;
            };
            if start > end || end > ccp_text {
                return Err(DocWriteError::InvalidData(
                    "DOC comment range must be ordered and inside the main document".to_string(),
                ));
            }
            let tag = i32::try_from(index).map_err(|_| {
                DocWriteError::InvalidData("DOC comment bookmark tag overflows".to_string())
            })? as u32;
            bookmark_tags[index] = Some(tag);
            ranges.push((tag, start, end));
        }

        let mut bookmark_names = Vec::new();
        let mut bookmark_starts = Vec::new();
        let mut bookmark_ends = Vec::new();
        if !ranges.is_empty() {
            let mut start_order = ranges.clone();
            start_order.sort_by_key(|&(tag, start, _)| (start, tag));
            let mut end_order = ranges.clone();
            end_order.sort_by_key(|&(tag, _, end)| (end, tag));
            let end_indexes = end_order
                .iter()
                .enumerate()
                .map(|(index, &(tag, _, _))| (tag, index as u16))
                .collect::<HashMap<_, _>>();

            bookmark_names.extend_from_slice(&0xFFFFu16.to_le_bytes());
            bookmark_names.extend_from_slice(&(ranges.len() as u16).to_le_bytes());
            bookmark_names.extend_from_slice(&10u16.to_le_bytes());
            for &(tag, _, _) in &start_order {
                bookmark_names.extend_from_slice(&0u16.to_le_bytes());
                bookmark_names.extend_from_slice(&0x0100u16.to_le_bytes());
                bookmark_names.extend_from_slice(&tag.to_le_bytes());
                bookmark_names.extend_from_slice(&(-1i32).to_le_bytes());
            }

            for &(_, start, _) in &start_order {
                bookmark_starts.extend_from_slice(&start.to_le_bytes());
            }
            bookmark_starts.extend_from_slice(&bookmark_sentinel.to_le_bytes());
            for &(tag, _, _) in &start_order {
                bookmark_starts.extend_from_slice(&end_indexes[&tag].to_le_bytes());
                bookmark_starts.extend_from_slice(&0u16.to_le_bytes());
            }

            for &(_, _, end) in &end_order {
                bookmark_ends.extend_from_slice(&end.to_le_bytes());
            }
            bookmark_ends.extend_from_slice(&bookmark_sentinel.to_le_bytes());
        }

        let mut extended_metadata = Vec::with_capacity(ordered.len() * 18);
        let mut active_ancestors = Vec::<usize>::new();
        for (index, (entry, _)) in ordered.iter().enumerate() {
            let metadata = entry
                .extended_metadata
                .unwrap_or(crate::doc::CommentExtendedMetadata {
                    modified_at: None,
                    depth: 0,
                    parent_index: None,
                    is_ink: false,
                });
            let depth = usize::try_from(metadata.depth).map_err(|_| {
                DocWriteError::InvalidData("DOC comment reply depth is too large".to_string())
            })?;
            if depth > active_ancestors.len() {
                return Err(DocWriteError::InvalidData(
                    "DOC comment reply tree must be in pre-order".to_string(),
                ));
            }
            active_ancestors.truncate(depth);
            let parent_delta = match (depth, metadata.parent_index) {
                (0, None) => 0,
                (0, Some(_)) | (_, None) => {
                    return Err(DocWriteError::InvalidData(
                        "DOC comment parent and reply depth are inconsistent".to_string(),
                    ));
                },
                (_, Some(parent)) => {
                    let expected = active_ancestors.get(depth - 1).copied().ok_or_else(|| {
                        DocWriteError::InvalidData(
                            "DOC comment reply tree is malformed".to_string(),
                        )
                    })?;
                    if parent != expected {
                        return Err(DocWriteError::InvalidData(
                            "DOC comment parent does not match pre-order reply depth".to_string(),
                        ));
                    }
                    i32::try_from(parent as i64 - index as i64).map_err(|_| {
                        DocWriteError::InvalidData(
                            "DOC comment parent offset exceeds the binary format".to_string(),
                        )
                    })?
                },
            };
            extended_metadata.extend_from_slice(&pack_dttm(metadata.modified_at)?.to_le_bytes());
            extended_metadata.extend_from_slice(&0u16.to_le_bytes());
            extended_metadata.extend_from_slice(&metadata.depth.to_le_bytes());
            extended_metadata.extend_from_slice(&parent_delta.to_le_bytes());
            extended_metadata.extend_from_slice(&(u32::from(metadata.is_ink) << 1).to_le_bytes());
            active_ancestors.push(index);
        }

        let mut comment_cp = 0u32;
        let mut text_cps = vec![0u32];
        for (entry, _) in &ordered {
            let fc_story_start = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x0005u16.to_le_bytes());
            let fc_marker_end = fc_story_start + 2;
            let marker_grpprl = build_chpx_grpprl(
                &CharacterFormatting {
                    special: Some(true),
                    ..Default::default()
                },
                font_builder,
            );
            chpx_entries.push((fc_story_start, fc_marker_end, marker_grpprl));

            let body_chars = utf16_code_unit_len(&entry.text)?;
            let fc_body_start = text_fc_start + text_stream.len() as u32;
            text_stream.extend(entry.text.encode_utf16().flat_map(u16::to_le_bytes));
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_story_end = text_fc_start + text_stream.len() as u32;
            chpx_entries.push((
                fc_body_start,
                fc_story_end,
                build_chpx_grpprl(&CharacterFormatting::default(), font_builder),
            ));
            papx_entries.push((
                fc_story_start,
                fc_story_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));

            let story_chars = body_chars.checked_add(2).ok_or_else(|| {
                DocWriteError::InvalidData("DOC comment story CP overflows".to_string())
            })?;
            let story_end = current_cp_total.checked_add(story_chars).ok_or_else(|| {
                DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
            })?;
            pieces.push(Piece::new(
                *current_cp_total,
                story_end,
                fc_story_start,
                true,
            ));
            *current_cp_total = story_end;
            comment_cp = comment_cp.checked_add(story_chars).ok_or_else(|| {
                DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
            })?;
            text_cps.push(comment_cp);
        }

        let fc_guard = text_fc_start + text_stream.len() as u32;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_guard_end = fc_guard + 2;
        chpx_entries.push((fc_guard, fc_guard_end, Vec::new()));
        papx_entries.push((
            fc_guard,
            fc_guard_end,
            build_papx_grpprl(&ParagraphFormatting::default()),
        ));
        let guard_end = current_cp_total.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
        })?;
        pieces.push(Piece::new(*current_cp_total, guard_end, fc_guard, true));
        *current_cp_total = guard_end;
        comment_cp = comment_cp.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
        })?;
        text_cps.push(comment_cp);

        let mut references = Vec::with_capacity((ordered.len() + 1) * 4 + ordered.len() * 30);
        for (_, cp) in &ordered {
            references.extend_from_slice(&cp.to_le_bytes());
        }
        references.extend_from_slice(&ccp_text.to_le_bytes());
        for (index, ((entry, _), author_index)) in ordered.iter().zip(owner_indexes).enumerate() {
            let initials = entry.initials.encode_utf16().collect::<Vec<_>>();
            references.extend_from_slice(&(initials.len() as u16).to_le_bytes());
            for index in 0..9 {
                references
                    .extend_from_slice(&initials.get(index).copied().unwrap_or(0).to_le_bytes());
            }
            references.extend_from_slice(&author_index.to_le_bytes());
            references.extend_from_slice(&0u16.to_le_bytes());
            references.extend_from_slice(&0u16.to_le_bytes());
            let tag = bookmark_tags[index].map_or(-1, |tag| tag as i32);
            references.extend_from_slice(&tag.to_le_bytes());
        }

        let mut text_positions = Vec::with_capacity(text_cps.len() * 4);
        for cp in text_cps {
            text_positions.extend_from_slice(&cp.to_le_bytes());
        }

        Ok(Some(CommentStoryData {
            owners: owner_bytes,
            references,
            text_positions,
            bookmark_names,
            bookmark_starts,
            bookmark_ends,
            extended_metadata,
            char_count: comment_cp,
        }))
    }

    fn append_comment_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        comment: &CommentStoryData,
    ) {
        let mut offset = table_stream.len() as u32;
        fib.set_grp_xst_atn_owners(offset, comment.owners.len() as u32);
        table_stream.extend_from_slice(&comment.owners);

        offset = table_stream.len() as u32;
        fib.set_plcfand_ref(offset, comment.references.len() as u32);
        table_stream.extend_from_slice(&comment.references);

        offset = table_stream.len() as u32;
        fib.set_plcfand_txt(offset, comment.text_positions.len() as u32);
        table_stream.extend_from_slice(&comment.text_positions);

        if !comment.bookmark_names.is_empty() {
            offset = table_stream.len() as u32;
            fib.set_sttbf_atn_bkmk(offset, comment.bookmark_names.len() as u32);
            table_stream.extend_from_slice(&comment.bookmark_names);

            offset = table_stream.len() as u32;
            fib.set_plcf_atn_bkf(offset, comment.bookmark_starts.len() as u32);
            table_stream.extend_from_slice(&comment.bookmark_starts);

            offset = table_stream.len() as u32;
            fib.set_plcf_atn_bkl(offset, comment.bookmark_ends.len() as u32);
            table_stream.extend_from_slice(&comment.bookmark_ends);
        }

        offset = table_stream.len() as u32;
        fib.set_atrd_extra(offset, comment.extended_metadata.len() as u32);
        table_stream.extend_from_slice(&comment.extended_metadata);
    }

    fn build_bookmark_tables(
        entries: &[BookmarkEntry],
        document_end: u32,
    ) -> Result<Option<BookmarkTableData>, DocWriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() > 0x3FFB {
            return Err(DocWriteError::InvalidData(
                "DOC standard bookmark table exceeds 0x3FFB entries".to_string(),
            ));
        }
        let mut unique = std::collections::HashSet::with_capacity(entries.len());
        let mut records = Vec::with_capacity(entries.len());
        for (index, entry) in entries.iter().enumerate() {
            let units = entry.name.encode_utf16().collect::<Vec<_>>();
            if units.is_empty() || units.len() >= 40 || !unique.insert(entry.name.clone()) {
                return Err(DocWriteError::InvalidData(
                    "DOC bookmark names must be unique and contain 1 through 39 UTF-16 code units"
                        .to_string(),
                ));
            }
            if entry.start > entry.end || entry.end > document_end {
                return Err(DocWriteError::InvalidData(
                    "DOC bookmark range must be ordered and inside the document parts".to_string(),
                ));
            }
            let mut bkc = u16::from(entry.is_native) << 14;
            if let Some((first, limit)) = entry.column_range {
                if first >= limit || first > 0x7F || limit > 0x3F {
                    return Err(DocWriteError::InvalidData(
                        "DOC bookmark column range exceeds BKC limits".to_string(),
                    ));
                }
                bkc |= 0x8000 | u16::from(first) | (u16::from(limit) << 8);
            }
            records.push((index, entry, units, bkc));
        }

        let sentinel = document_end.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC bookmark sentinel CP overflows".to_string())
        })?;
        let mut start_order = records.iter().collect::<Vec<_>>();
        start_order.sort_by_key(|record| (record.1.start, record.0));
        let mut end_order = records.iter().collect::<Vec<_>>();
        end_order.sort_by_key(|record| (record.1.end, record.0));
        let end_indexes = end_order
            .iter()
            .enumerate()
            .map(|(end_index, record)| (record.0, end_index as u16))
            .collect::<HashMap<_, _>>();

        let mut names = Vec::new();
        names.extend_from_slice(&0xFFFFu16.to_le_bytes());
        names.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        names.extend_from_slice(&0u16.to_le_bytes());
        for record in &start_order {
            names.extend_from_slice(&(record.2.len() as u16).to_le_bytes());
            names.extend(record.2.iter().copied().flat_map(u16::to_le_bytes));
        }

        let mut starts = Vec::with_capacity((entries.len() + 1) * 4 + entries.len() * 4);
        for record in &start_order {
            starts.extend_from_slice(&record.1.start.to_le_bytes());
        }
        starts.extend_from_slice(&sentinel.to_le_bytes());
        for record in &start_order {
            starts.extend_from_slice(&end_indexes[&record.0].to_le_bytes());
            starts.extend_from_slice(&record.3.to_le_bytes());
        }

        let mut ends = Vec::with_capacity((entries.len() + 1) * 4);
        for record in &end_order {
            ends.extend_from_slice(&record.1.end.to_le_bytes());
        }
        ends.extend_from_slice(&sentinel.to_le_bytes());
        Ok(Some(BookmarkTableData {
            names,
            starts,
            ends,
        }))
    }

    fn build_revision_writer_data(&self) -> Result<Option<RevisionWriterData>, DocWriteError> {
        let mut authors = vec!["Unknown".to_string()];
        let mut indexes = HashMap::from([("Unknown".to_string(), 0u16)]);
        let mut has_revisions = false;
        let mut index_author = |author: &str| -> Result<(), DocWriteError> {
            has_revisions = true;
            if !indexes.contains_key(author) {
                if authors.len() >= 0x8000 {
                    return Err(DocWriteError::InvalidData(
                        "DOC revision author table exceeds the signed author-index range"
                            .to_string(),
                    ));
                }
                let index = authors.len() as u16;
                authors.push(author.to_string());
                indexes.insert(author.to_string(), index);
            }
            Ok(())
        };
        if let Some(revision) = &self.section_formatting_revision {
            index_author(&revision.author)?;
        }
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            if let Some(revision) = &paragraph.formatting.formatting_revision {
                index_author(&revision.author)?;
            }
            if let Some(revision) = &paragraph.formatting.numbering_revision {
                index_author(&revision.author)?;
            }
            for run in &paragraph.runs {
                if let Some(revision) = &run.formatting.insertion_revision {
                    index_author(&revision.author)?;
                }
                if let Some(revision) = &run.formatting.deletion_revision {
                    index_author(&revision.author)?;
                }
                if let Some(revision) = &run.formatting.formatting_revision {
                    index_author(&revision.author)?;
                }
                if let Some(revision) = &run.formatting.display_field_revision {
                    index_author(&revision.author)?;
                }
            }
        }
        if !has_revisions {
            return Ok(None);
        }

        let mut table = Vec::new();
        table.extend_from_slice(&0xFFFFu16.to_le_bytes());
        table.extend_from_slice(&(authors.len() as u16).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        for author in authors {
            let units = author.encode_utf16().collect::<Vec<_>>();
            let length = u16::try_from(units.len()).map_err(|_| {
                DocWriteError::InvalidData(
                    "DOC revision author exceeds the STTB string-length limit".to_string(),
                )
            })?;
            table.extend_from_slice(&length.to_le_bytes());
            table.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        Ok(Some(RevisionWriterData { indexes, table }))
    }

    fn append_revision_author_table(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        revisions: &RevisionWriterData,
    ) {
        let offset = table_stream.len() as u32;
        fib.set_sttbf_rmark(offset, revisions.table.len() as u32);
        table_stream.extend_from_slice(&revisions.table);
    }

    #[allow(clippy::too_many_arguments)]
    fn append_tables_to_main_story(
        &self,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        field_char_cps: &mut Vec<(u32, u16)>,
        font_builder: &mut FontTableBuilder,
        revision_data: Option<&RevisionWriterData>,
    ) -> Result<(), DocWriteError> {
        for table in &self.tables {
            let mut encountered_body_row = false;
            let mut vertical_merges = table
                .rows
                .first()
                .map(|row| vec![false; row.cells.len()])
                .unwrap_or_default();
            for row in &table.rows {
                let column_count = row.cells.len();
                if !(1..=63).contains(&column_count) {
                    return Err(DocWriteError::InvalidData(
                        "DOC table rows must contain between 1 and 63 cells".to_string(),
                    ));
                }
                if row.formatting.cells.len() != column_count {
                    return Err(DocWriteError::InvalidData(
                        "DOC table row formatting must define every cell".to_string(),
                    ));
                }
                if row.formatting.is_header && encountered_body_row {
                    return Err(DocWriteError::InvalidData(
                        "DOC header rows must form a contiguous prefix of the table".to_string(),
                    ));
                }
                encountered_body_row |= !row.formatting.is_header;
                for (index, cell) in row.formatting.cells.iter().enumerate() {
                    match cell.vertical_merge {
                        crate::doc::parts::tap::VerticalMergeStatus::None => {
                            vertical_merges[index] = false;
                        },
                        crate::doc::parts::tap::VerticalMergeStatus::First => {
                            vertical_merges[index] = true;
                        },
                        crate::doc::parts::tap::VerticalMergeStatus::Merged => {
                            if !vertical_merges[index] {
                                return Err(DocWriteError::InvalidData(format!(
                                    "DOC cell {index} continues a vertical merge that was not started"
                                )));
                            }
                        },
                    }
                }
                for cell in &row.cells {
                    if cell.paragraphs.is_empty() {
                        return Err(DocWriteError::InvalidData(
                            "DOC table cells must contain at least one paragraph".to_string(),
                        ));
                    }
                    let last_paragraph = cell.paragraphs.len() - 1;
                    for (index, paragraph) in cell.paragraphs.iter().enumerate() {
                        let terminator = if index == last_paragraph {
                            0x0007
                        } else {
                            0x000D
                        };
                        Self::append_table_paragraph(
                            paragraph,
                            terminator,
                            text_fc_start,
                            text_stream,
                            current_cp,
                            pieces,
                            chpx_entries,
                            papx_entries,
                            field_char_cps,
                            font_builder,
                            revision_data,
                        )?;
                    }
                }

                let fc_start = text_fc_start
                    .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                        DocWriteError::InvalidData(
                            "DOC text stream exceeds 32-bit FC space".to_string(),
                        )
                    })?)
                    .ok_or_else(|| {
                        DocWriteError::InvalidData("DOC table row FC overflows".to_string())
                    })?;
                text_stream.extend_from_slice(&0x0007u16.to_le_bytes());
                let fc_end = fc_start.checked_add(2).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table row FC overflows".to_string())
                })?;
                chpx_entries.push((fc_start, fc_end, Vec::new()));
                papx_entries.push((
                    fc_start,
                    fc_end,
                    build_table_row_papx_grpprl(&row.formatting)?,
                ));
                let cp_end = current_cp.checked_add(1).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table CP range overflows".to_string())
                })?;
                pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
                *current_cp = cp_end;
            }

            // The main document must end in U+000D. A non-table paragraph also
            // separates adjacent writer table objects into distinct tables.
            Self::append_empty_main_paragraph(
                text_fc_start,
                text_stream,
                current_cp,
                pieces,
                chpx_entries,
                papx_entries,
            )?;
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn append_table_paragraph(
        paragraph: &WritableParagraph,
        terminator: u16,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        field_char_cps: &mut Vec<(u32, u16)>,
        font_builder: &mut FontTableBuilder,
        revision_data: Option<&RevisionWriterData>,
    ) -> Result<(), DocWriteError> {
        let fc_start = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph FC overflows".to_string())
            })?;
        let mut paragraph_cps = 0u32;
        let mut last_chpx = None;
        for run in &paragraph.runs {
            let run_fc_start = text_fc_start
                .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                    DocWriteError::InvalidData(
                        "DOC text stream exceeds 32-bit FC space".to_string(),
                    )
                })?)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run FC overflows".to_string())
                })?;
            let run_cps = utf16_code_unit_len(&run.text)?;
            let mut offset = 0u32;
            for ch in run.text.chars() {
                let cp = current_cp
                    .checked_add(paragraph_cps)
                    .and_then(|value| value.checked_add(offset))
                    .ok_or_else(|| {
                        DocWriteError::InvalidData(
                            "DOC table field character CP overflows".to_string(),
                        )
                    })?;
                if matches!(ch as u32, 0x0013..=0x0015) {
                    field_char_cps.push((cp, ch as u16));
                }
                offset = offset.checked_add(ch.len_utf16() as u32).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run CP range overflows".to_string())
                })?;
            }
            for unit in run.text.encode_utf16() {
                text_stream.extend_from_slice(&unit.to_le_bytes());
            }
            let run_fc_end = run_fc_start
                .checked_add(run_cps.checked_mul(2).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run FC overflows".to_string())
                })?)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run FC overflows".to_string())
                })?;
            chpx_entries.push((
                run_fc_start,
                run_fc_end,
                build_revision_chpx_grpprl(&run.formatting, font_builder, revision_data)?,
            ));
            last_chpx = Some(chpx_entries.len() - 1);
            paragraph_cps = paragraph_cps.checked_add(run_cps).ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph CP range overflows".to_string())
            })?;
        }
        text_stream.extend_from_slice(&terminator.to_le_bytes());
        let fc_end = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph FC overflows".to_string())
            })?;
        if let Some(index) = last_chpx {
            chpx_entries[index].1 = fc_end;
        } else {
            chpx_entries.push((fc_start, fc_end, Vec::new()));
        }
        let mut papx = build_revision_papx_grpprl(&paragraph.formatting, revision_data)?;
        append_table_depth_sprms(&mut papx);
        papx_entries.push((fc_start, fc_end, papx));
        let cp_end = current_cp
            .checked_add(paragraph_cps)
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph CP range overflows".to_string())
            })?;
        pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
        *current_cp = cp_end;
        Ok(())
    }

    fn append_empty_main_paragraph(
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
    ) -> Result<(), DocWriteError> {
        let fc_start = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC final paragraph FC overflows".to_string())
            })?;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_end = fc_start.checked_add(2).ok_or_else(|| {
            DocWriteError::InvalidData("DOC final paragraph FC overflows".to_string())
        })?;
        chpx_entries.push((fc_start, fc_end, Vec::new()));
        papx_entries.push((fc_start, fc_end, Vec::new()));
        let cp_end = current_cp.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC final paragraph CP overflows".to_string())
        })?;
        pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
        *current_cp = cp_end;
        Ok(())
    }

    fn append_bookmark_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        bookmarks: &BookmarkTableData,
    ) {
        let mut offset = table_stream.len() as u32;
        fib.set_sttbf_bkmk(offset, bookmarks.names.len() as u32);
        table_stream.extend_from_slice(&bookmarks.names);
        offset = table_stream.len() as u32;
        fib.set_plcf_bkf(offset, bookmarks.starts.len() as u32);
        table_stream.extend_from_slice(&bookmarks.starts);
        offset = table_stream.len() as u32;
        fib.set_plcf_bkl(offset, bookmarks.ends.len() as u32);
        table_stream.extend_from_slice(&bookmarks.ends);
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
    ) -> Result<Option<HeaderStoryData>, DocWriteError> {
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
            return Ok(None);
        }

        // Build index->text mapping for 12 slots per MS-DOC PlcfHdd / Apache POI:
        //   Slots 0-5:  footnote/endnote separator/continuation stories
        //   Slot 6:     even page header (section 0)
        //   Slot 7:     odd page header (section 0) — "default" when no facing pages
        //   Slot 8:     even page footer (section 0)
        //   Slot 9:     odd page footer (section 0) — "default" when no facing pages
        //   Slot 10:    first page header (section 0)
        //   Slot 11:    first page footer (section 0)
        // PlcfHdd has 14 CPs (12 slot starts + story end + ignored final CP).
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
        // Empty slots consume no CPs. Non-empty header/footer stories contain a content paragraph
        // mark and a separate guard paragraph mark.
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
                para_chars += utf16_code_unit_len(text)?;
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
            }
        }

        // The header subdocument ends with an extra paragraph mark. The second-to-last PlcfHdd
        // CP terminates the final story at ccpHdd - 1; the last CP is ignored.
        let stories_end = header_cp;
        let fc_trailing = text_fc_start + text_stream.len() as u32;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_trailing_end = fc_trailing + 2;
        chpx_entries.push((fc_trailing, fc_trailing_end, Vec::new()));
        papx_entries.push((
            fc_trailing,
            fc_trailing_end,
            build_papx_grpprl(&ParagraphFormatting::default()),
        ));
        pieces.push(Piece::new(
            *current_cp_total,
            *current_cp_total + 1,
            fc_trailing,
            true,
        ));
        *current_cp_total += 1;
        header_cp += 1;

        let mut plcfhdd = Vec::with_capacity((12 + 2) * 4);
        for cp_start in &cp_starts {
            plcfhdd.extend_from_slice(&cp_start.to_le_bytes());
        }
        plcfhdd.extend_from_slice(&stories_end.to_le_bytes());
        plcfhdd.extend_from_slice(&header_cp.to_le_bytes());

        Ok(Some((plcfhdd, header_cp)))
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
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Result<usize, DocWriteError> {
        if rows == 0 || cols == 0 {
            return Err(DocWriteError::InvalidData(
                "Table must have at least 1 row and 1 column".to_string(),
            ));
        }
        if cols > 63 {
            return Err(DocWriteError::InvalidData(
                "DOC table rows cannot exceed 63 cells".to_string(),
            ));
        }

        let mut table = WritableTable { rows: Vec::new() };

        for _ in 0..rows {
            let mut row = TableRow {
                cells: Vec::new(),
                formatting: super::tap::TableRow {
                    cells: Vec::with_capacity(cols),
                    height: 0,
                    is_header: false,
                    allow_break: true,
                },
            };
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
            for index in 0..cols {
                const DEFAULT_TABLE_WIDTH: u32 = 8640;
                let left = DEFAULT_TABLE_WIDTH * index as u32 / cols as u32;
                let right = DEFAULT_TABLE_WIDTH * (index + 1) as u32 / cols as u32;
                row.formatting.cells.push(super::tap::TableCell {
                    width: (right - left) as u16,
                    merged: false,
                    ..super::tap::TableCell::default()
                });
            }
            table.rows.push(row);
        }

        let index = self.tables.len();
        self.tables.push(table);
        Ok(index)
    }

    /// Mark the document section's properties as a tracked formatting change.
    ///
    /// The legacy writer currently emits one section spanning the document, so
    /// this revision applies to that complete section.
    pub fn set_section_formatting_revision(&mut self, revision: FormattingRevision) {
        self.section_formatting_revision = Some(revision);
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
        self.set_table_cell_paragraph_runs(
            table_idx,
            row,
            col,
            vec![(text.to_string(), CharacterFormatting::default())],
            ParagraphFormatting::default(),
        )
    }

    /// Replace a table cell with one paragraph composed of formatted runs.
    pub fn set_table_cell_paragraph_runs(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
        runs: Vec<(String, CharacterFormatting)>,
        formatting: ParagraphFormatting,
    ) -> Result<(), DocWriteError> {
        let paragraph = writable_paragraph_from_runs(runs, formatting);
        self.table_cell_mut(table_idx, row, col)?.paragraphs = vec![paragraph];
        Ok(())
    }

    /// Append a paragraph composed of formatted runs to a table cell.
    pub fn append_table_cell_paragraph_runs(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
        runs: Vec<(String, CharacterFormatting)>,
        formatting: ParagraphFormatting,
    ) -> Result<(), DocWriteError> {
        let paragraph = writable_paragraph_from_runs(runs, formatting);
        self.table_cell_mut(table_idx, row, col)?
            .paragraphs
            .push(paragraph);
        Ok(())
    }

    fn table_cell_mut(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
    ) -> Result<&mut TableCell, DocWriteError> {
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
        Ok(cell)
    }

    /// Set the widths, horizontal merges, height, and header state for a table row.
    pub fn set_table_row_formatting(
        &mut self,
        table_idx: usize,
        row: usize,
        formatting: super::tap::TableRow,
    ) -> Result<(), DocWriteError> {
        let table = self
            .tables
            .get_mut(table_idx)
            .ok_or_else(|| DocWriteError::InvalidData(format!("Table {table_idx} not found")))?;
        let row_data = table
            .rows
            .get_mut(row)
            .ok_or_else(|| DocWriteError::InvalidData(format!("Row {row} not found")))?;
        if formatting.cells.len() != row_data.cells.len() {
            return Err(DocWriteError::InvalidData(format!(
                "Row {row} formatting has {} cells but the row contains {}",
                formatting.cells.len(),
                row_data.cells.len()
            )));
        }
        super::tap::generate_row_sprms(&formatting)
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
        row_data.formatting = formatting;
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
        let revision_data = self.build_revision_writer_data()?;

        // Pad to 512-byte boundary before writing text (POI line 428-433)
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        let text_fc_start = word_document_stream.len() as u32;
        // Align fcMin to actual (padded) start of text
        let fc_min: u32 = text_fc_start;

        // Build one sorted list for all main-story reference characters.
        let mut main_refs: Vec<(u32, MainReferenceKind, usize)> = Vec::new();
        for (idx, entry) in self.footnotes.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Footnote, idx));
        }
        for (idx, entry) in self.endnotes.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Endnote, idx));
        }
        for (idx, entry) in self.comments.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Comment, idx));
        }
        main_refs.sort_by_key(|reference| reference.0);

        // Track field character CPs for PlcfFldMom
        let mut field_char_cps: Vec<(u32, u16)> = Vec::new();

        // Track actual CPs by source-vector index for each reference kind.
        let mut footnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut endnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut comment_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut reference_inject_idx: usize = 0;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = utf16_code_unit_len(run_text)?;
                let grpprl = build_revision_chpx_grpprl(
                    &run.formatting,
                    &mut font_builder,
                    revision_data.as_ref(),
                )?;

                // Track field characters in this run
                let mut utf16_offset = 0u32;
                for ch in run_text.chars() {
                    let cp = current_cp + para_chars + utf16_offset;
                    match ch as u32 {
                        0x0013 => field_char_cps.push((cp, 0x13)),
                        0x0014 => field_char_cps.push((cp, 0x14)),
                        0x0015 => field_char_cps.push((cp, 0x15)),
                        _ => {},
                    }
                    utf16_offset += ch.len_utf16() as u32;
                }
                debug_assert_eq!(utf16_offset, run_len_chars);

                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }

            // Inject special reference characters whose requested CP falls in this paragraph.
            while reference_inject_idx < main_refs.len() {
                let (ref_cp, kind, entry_idx) = main_refs[reference_inject_idx];
                if ref_cp <= current_cp + para_chars {
                    let actual_cp = current_cp + para_chars;
                    let fc_ref = text_fc_start + text_stream.len() as u32;
                    let marker = match kind {
                        MainReferenceKind::Footnote | MainReferenceKind::Endnote => 0x0002u16,
                        MainReferenceKind::Comment => 0x0005u16,
                    };
                    text_stream.extend_from_slice(&marker.to_le_bytes());
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
                    match kind {
                        MainReferenceKind::Footnote => {
                            footnote_actual_cps.push((entry_idx, actual_cp));
                        },
                        MainReferenceKind::Endnote => {
                            endnote_actual_cps.push((entry_idx, actual_cp));
                        },
                        MainReferenceKind::Comment => {
                            comment_actual_cps.push((entry_idx, actual_cp));
                        },
                    }
                    reference_inject_idx += 1;
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
            let pap_grpprl =
                build_revision_papx_grpprl(&paragraph.formatting, revision_data.as_ref())?;
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

        self.append_tables_to_main_story(
            text_fc_start,
            &mut text_stream,
            &mut current_cp,
            &mut pieces,
            &mut chpx_entries,
            &mut papx_entries,
            &mut field_char_cps,
            &mut font_builder,
            revision_data.as_ref(),
        )?;
        if current_cp == 0 {
            Self::append_empty_main_paragraph(
                text_fc_start,
                &mut text_stream,
                &mut current_cp,
                &mut pieces,
                &mut chpx_entries,
                &mut papx_entries,
            )?;
        }

        let text_length = current_cp;

        // Sort actual CPs by entry index to match footnote/endnote entry order
        footnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        endnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        comment_actual_cps.sort_by_key(|&(idx, _)| idx);
        let ftn_ref_cps: Vec<u32> = footnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let edn_ref_cps: Vec<u32> = endnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let comment_ref_cps: Vec<u32> = comment_actual_cps.iter().map(|&(_, cp)| cp).collect();

        // Subdocument order: main text → footnotes → headers/footers → comments → endnotes.
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
        )?;

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
        )? {
            header_plcfhdd = Some((plcf_bytes, header_cp));
        }

        // Comments follow headers and precede endnotes in the concatenated CP space.
        let comment_story = Self::build_comment_story(
            &self.comments,
            &comment_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        )?;

        // Build endnote story (appends endnote text after comments)
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
        )?;
        let bookmark_tables = Self::build_bookmark_tables(&self.bookmarks, current_cp)?;

        // Mandatory trailing paragraph mark when ANY subdocument exists.
        // Per MS-DOC spec: "The total number of character positions is
        // ccpText + ccpFtn + ccpHdd + ... + 1 if any of ccpFtn, ccpHdd, etc. are nonzero."
        // This extra character MUST be present; Word uses it as a sentinel.
        let has_subdocs = footnote_plcfs.is_some()
            || header_plcfhdd.is_some()
            || comment_story.is_some()
            || endnote_plcfs.is_some();
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
        if let Some(comment) = &comment_story {
            fib.set_ccp_atn(comment.char_count);
        }
        if let Some((_, _, edn_cp)) = &endnote_plcfs {
            fib.set_ccp_edn(*edn_cp);
        }

        let mut table_offset = 0u32;

        // 3. Write StyleSheet to table stream (MANDATORY - POI line 681-684)
        let stylesheet_data = crate::doc::writer::stylesheet::generate_minimal_stylesheet();
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
        let dop_data = crate::doc::writer::dop::generate_dop(facing_pages, doc_grpf_ihdt);
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

        if let Some(comment) = &comment_story {
            Self::append_comment_tables(&mut fib, &mut table_stream, comment);
            table_offset = table_stream.len() as u32;
        }
        if let Some(bookmarks) = &bookmark_tables {
            Self::append_bookmark_tables(&mut fib, &mut table_stream, bookmarks);
            table_offset = table_stream.len() as u32;
        }
        if let Some(revisions) = &revision_data {
            Self::append_revision_author_table(&mut fib, &mut table_stream, revisions);
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
        let mut chpx_builder = crate::doc::writer::fkp::ChpxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_pages = chpx_builder.generate_pages()?;
        for page in &chpx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── PAPX FKPs (multi-page) ──
        let papx_first_page = (word_document_stream.len() / 512) as u32;
        let mut papx_builder = crate::doc::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_pages = papx_builder.generate_pages()?;
        for page in &papx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── Write bin tables to table stream (now that we know page numbers) ──
        let chpx_bin_table = crate::doc::writer::bin_table::generate_bin_table_from_pages(
            &chpx_pages.ranges,
            chpx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_chpx(table_offset, chpx_bin_table.len() as u32);
        table_stream.extend_from_slice(&chpx_bin_table);

        let papx_bin_table = crate::doc::writer::bin_table::generate_bin_table_from_pages(
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
        let section_revision = self
            .section_formatting_revision
            .as_ref()
            .map(|revision| {
                Ok::<_, DocWriteError>((
                    revision_data
                        .as_ref()
                        .expect("section revisions initialize revision writer data")
                        .indexes[&revision.author],
                    pack_dttm(revision.timestamp)?,
                ))
            })
            .transpose()?;
        let sepx_data = crate::doc::writer::section::generate_sepx_with_revision(
            first_page,
            grpf_ihdt,
            section_revision,
        );
        word_document_stream.extend_from_slice(&sepx_data);

        // 10c. Write section table to table stream with correct SEPX offset
        // Section table CP must span ALL subdocuments (main + footnotes + headers + endnotes),
        // not just ccpText. Per MS-DOC spec and Apache POI SectionTable.
        let total_cp = current_cp;
        let section_table =
            crate::doc::writer::section::generate_section_table(total_cp, sepx_offset);
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
        let compobj_data = crate::doc::writer::ole_metadata::generate_compobj_stream();
        let ole_data = crate::doc::writer::ole_metadata::generate_ole_stream();
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
        let revision_data = self.build_revision_writer_data()?;

        // Pad to 512-byte boundary before text
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        let text_fc_start = word_document_stream.len() as u32;
        let fc_min: u32 = text_fc_start;

        // Build one sorted list for all main-story reference characters.
        let mut main_refs: Vec<(u32, MainReferenceKind, usize)> = Vec::new();
        for (idx, entry) in self.footnotes.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Footnote, idx));
        }
        for (idx, entry) in self.endnotes.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Endnote, idx));
        }
        for (idx, entry) in self.comments.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Comment, idx));
        }
        main_refs.sort_by_key(|reference| reference.0);

        let mut field_char_cps: Vec<(u32, u16)> = Vec::new();
        let mut footnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut endnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut comment_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut reference_inject_idx: usize = 0;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = utf16_code_unit_len(run_text)?;
                let grpprl = build_revision_chpx_grpprl(
                    &run.formatting,
                    &mut font_builder,
                    revision_data.as_ref(),
                )?;

                let mut utf16_offset = 0u32;
                for ch in run_text.chars() {
                    let cp = current_cp + para_chars + utf16_offset;
                    match ch as u32 {
                        0x0013 => field_char_cps.push((cp, 0x13)),
                        0x0014 => field_char_cps.push((cp, 0x14)),
                        0x0015 => field_char_cps.push((cp, 0x15)),
                        _ => {},
                    }
                    utf16_offset += ch.len_utf16() as u32;
                }
                debug_assert_eq!(utf16_offset, run_len_chars);

                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }

            while reference_inject_idx < main_refs.len() {
                let (ref_cp, kind, entry_idx) = main_refs[reference_inject_idx];
                if ref_cp <= current_cp + para_chars {
                    let actual_cp = current_cp + para_chars;
                    let fc_ref = text_fc_start + text_stream.len() as u32;
                    let marker = match kind {
                        MainReferenceKind::Footnote | MainReferenceKind::Endnote => 0x0002u16,
                        MainReferenceKind::Comment => 0x0005u16,
                    };
                    text_stream.extend_from_slice(&marker.to_le_bytes());
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
                    match kind {
                        MainReferenceKind::Footnote => {
                            footnote_actual_cps.push((entry_idx, actual_cp));
                        },
                        MainReferenceKind::Endnote => {
                            endnote_actual_cps.push((entry_idx, actual_cp));
                        },
                        MainReferenceKind::Comment => {
                            comment_actual_cps.push((entry_idx, actual_cp));
                        },
                    }
                    reference_inject_idx += 1;
                } else {
                    break;
                }
            }

            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            if let Some(last_idx) = last_run_index_for_para {
                chpx_entries[last_idx].1 += 2;
            }
            let fc_para_end = text_fc_start + text_stream.len() as u32;
            let pap_grpprl =
                build_revision_papx_grpprl(&paragraph.formatting, revision_data.as_ref())?;
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

        self.append_tables_to_main_story(
            text_fc_start,
            &mut text_stream,
            &mut current_cp,
            &mut pieces,
            &mut chpx_entries,
            &mut papx_entries,
            &mut field_char_cps,
            &mut font_builder,
            revision_data.as_ref(),
        )?;
        if current_cp == 0 {
            Self::append_empty_main_paragraph(
                text_fc_start,
                &mut text_stream,
                &mut current_cp,
                &mut pieces,
                &mut chpx_entries,
                &mut papx_entries,
            )?;
        }

        let text_length = current_cp;

        footnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        endnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        comment_actual_cps.sort_by_key(|&(idx, _)| idx);
        let ftn_ref_cps: Vec<u32> = footnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let edn_ref_cps: Vec<u32> = endnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let comment_ref_cps: Vec<u32> = comment_actual_cps.iter().map(|&(_, cp)| cp).collect();

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
        )?;

        let mut header_plcfhdd: Option<(Vec<u8>, u32)> = None;
        if let Some((plcf_bytes, header_cp)) = self.build_header_story(
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        )? {
            header_plcfhdd = Some((plcf_bytes, header_cp));
        }

        let comment_story = Self::build_comment_story(
            &self.comments,
            &comment_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        )?;

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
        )?;
        let bookmark_tables = Self::build_bookmark_tables(&self.bookmarks, current_cp)?;

        // Mandatory trailing paragraph mark when ANY subdocument exists (same as save()).
        let has_subdocs = footnote_plcfs.is_some()
            || header_plcfhdd.is_some()
            || comment_story.is_some()
            || endnote_plcfs.is_some();
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
        if let Some(comment) = &comment_story {
            fib.set_ccp_atn(comment.char_count);
        }
        if let Some((_, _, edn_cp)) = &endnote_plcfs {
            fib.set_ccp_edn(*edn_cp);
        }

        let mut table_offset = 0u32;

        let stylesheet_data = crate::doc::writer::stylesheet::generate_minimal_stylesheet();
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
        let dop_data = crate::doc::writer::dop::generate_dop(facing_pages, doc_grpf_ihdt);
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

        if let Some(comment) = &comment_story {
            Self::append_comment_tables(&mut fib, &mut table_stream, comment);
            table_offset = table_stream.len() as u32;
        }
        if let Some(bookmarks) = &bookmark_tables {
            Self::append_bookmark_tables(&mut fib, &mut table_stream, bookmarks);
            table_offset = table_stream.len() as u32;
        }
        if let Some(revisions) = &revision_data {
            Self::append_revision_author_table(&mut fib, &mut table_stream, revisions);
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
        let mut chpx_builder = crate::doc::writer::fkp::ChpxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_pages = chpx_builder.generate_pages()?;
        for page in &chpx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── PAPX FKPs (multi-page) ──
        let papx_first_page = (word_document_stream.len() / 512) as u32;
        let mut papx_builder = crate::doc::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_pages = papx_builder.generate_pages()?;
        for page in &papx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── Write bin tables to table stream ──
        let chpx_bin_table = crate::doc::writer::bin_table::generate_bin_table_from_pages(
            &chpx_pages.ranges,
            chpx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_chpx(table_offset, chpx_bin_table.len() as u32);
        table_stream.extend_from_slice(&chpx_bin_table);

        let papx_bin_table = crate::doc::writer::bin_table::generate_bin_table_from_pages(
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
        let section_revision = self
            .section_formatting_revision
            .as_ref()
            .map(|revision| {
                Ok::<_, DocWriteError>((
                    revision_data
                        .as_ref()
                        .expect("section revisions initialize revision writer data")
                        .indexes[&revision.author],
                    pack_dttm(revision.timestamp)?,
                ))
            })
            .transpose()?;
        let sepx_data = crate::doc::writer::section::generate_sepx_with_revision(
            first_page,
            grpf_ihdt,
            section_revision,
        );
        word_document_stream.extend_from_slice(&sepx_data);

        // Write section table to table stream
        let total_cp = current_cp;
        let section_table =
            crate::doc::writer::section::generate_section_table(total_cp, sepx_offset);
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
        let compobj_data = crate::doc::writer::ole_metadata::generate_compobj_stream();
        let ole_data = crate::doc::writer::ole_metadata::generate_ole_stream();
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

fn build_revision_chpx_grpprl(
    fmt: &CharacterFormatting,
    font_builder: &mut FontTableBuilder,
    revisions: Option<&RevisionWriterData>,
) -> Result<Vec<u8>, DocWriteError> {
    if fmt.insertion_revision.is_some() && fmt.deletion_revision.is_some() {
        return Err(DocWriteError::InvalidData(
            "a DOC character run cannot be both an insertion and a deletion".to_string(),
        ));
    }
    let mut grp = build_chpx_grpprl(fmt, font_builder);
    let Some(revisions) = revisions else {
        return Ok(grp);
    };
    let mut append = |revision: &TextRevision,
                      flag_opcode: u16,
                      author_opcode: u16,
                      time_opcode: u16,
                      reason_opcode: u16,
                      rsid_opcode: u16|
     -> Result<(), DocWriteError> {
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            DocWriteError::InvalidData("DOC revision author was not indexed".to_string())
        })?;
        grp.extend_from_slice(&flag_opcode.to_le_bytes());
        grp.push(1);
        grp.extend_from_slice(&author_opcode.to_le_bytes());
        grp.extend_from_slice(&author_index.to_le_bytes());
        if revision.timestamp.is_some() {
            grp.extend_from_slice(&time_opcode.to_le_bytes());
            grp.extend_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        }
        let structured_reason = revision.reason.map(crate::doc::RevisionReason::raw);
        if let (Some(raw), Some(structured)) = (revision.revision_id, structured_reason)
            && raw != structured
        {
            return Err(DocWriteError::InvalidData(
                "DOC revision contains conflicting raw and structured reason codes".to_string(),
            ));
        }
        if let Some(reason) = structured_reason.or(revision.revision_id) {
            if reason > crate::doc::RevisionReason::MAX_VALUE {
                return Err(DocWriteError::InvalidData(
                    "DOC revision reason code is undefined".to_string(),
                ));
            }
            grp.extend_from_slice(&reason_opcode.to_le_bytes());
            grp.extend_from_slice(&reason.to_le_bytes());
        }
        if let Some(revision_save_id) = revision.revision_save_id {
            grp.extend_from_slice(&rsid_opcode.to_le_bytes());
            grp.extend_from_slice(&revision_save_id.to_le_bytes());
        }
        Ok(())
    };
    if let Some(revision) = &fmt.insertion_revision {
        append(
            revision,
            SPRM_C_F_RMARK,
            SPRM_C_IBST_RMARK,
            SPRM_C_DTTM_RMARK,
            SPRM_C_IDSL_RMARK,
            SPRM_C_RSID_TEXT,
        )?;
    }
    if let Some(revision) = &fmt.deletion_revision {
        append(
            revision,
            SPRM_C_F_RMARK_DEL,
            SPRM_C_IBST_RMARK_DEL,
            SPRM_C_DTTM_RMARK_DEL,
            SPRM_C_IDSL_RMARK_DEL,
            SPRM_C_RSID_RM_DEL,
        )?;
    }
    if let Some(revision) = &fmt.formatting_revision {
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            DocWriteError::InvalidData("DOC revision author was not indexed".to_string())
        })?;
        grp.extend_from_slice(&SPRM_C_PROP_RMARK_CURRENT.to_le_bytes());
        grp.push(7);
        grp.push(1);
        grp.extend_from_slice(&author_index.to_le_bytes());
        grp.extend_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        if let Some(reason) = revision.reason {
            let insertion_reason = fmt.insertion_revision.as_ref().and_then(|insertion| {
                insertion
                    .reason
                    .map(crate::doc::RevisionReason::raw)
                    .or(insertion.revision_id)
            });
            if insertion_reason.is_some_and(|value| value != reason.raw()) {
                return Err(DocWriteError::InvalidData(
                    "DOC insertion and formatting revisions have conflicting reason codes"
                        .to_string(),
                ));
            }
            grp.extend_from_slice(&SPRM_C_IDSL_RMARK.to_le_bytes());
            grp.extend_from_slice(&reason.raw().to_le_bytes());
        }
        if let Some(revision_save_id) = revision.revision_save_id {
            grp.extend_from_slice(&SPRM_C_RSID_PROP.to_le_bytes());
            grp.extend_from_slice(&revision_save_id.to_le_bytes());
        }
    }
    if let Some(revision) = &fmt.display_field_revision {
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            DocWriteError::InvalidData(
                "DOC display-field revision author was not indexed".to_string(),
            )
        })?;
        let units = revision.previous_result.encode_utf16().collect::<Vec<_>>();
        if units.len() > 15 {
            return Err(DocWriteError::InvalidData(
                "DOC LISTNUM previous result exceeds its 15-code-unit XST".to_string(),
            ));
        }
        let mut operand = [0u8; 39];
        operand[0] = 1;
        operand[1..3].copy_from_slice(&author_index.to_le_bytes());
        operand[3..7].copy_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        operand[7..9].copy_from_slice(&(units.len() as u16).to_le_bytes());
        for (index, unit) in units.into_iter().enumerate() {
            let offset = 9 + index * 2;
            operand[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
        }
        grp.extend_from_slice(&SPRM_C_DISP_FLD_RMARK.to_le_bytes());
        grp.push(39);
        grp.extend_from_slice(&operand);
    }
    Ok(grp)
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
    if let Some(applied) = fmt.numbering_revision_list_applied {
        push_bool(&mut grp, SPRM_P_F_NUM_RM_INS, applied);
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

    grp
}

fn build_revision_papx_grpprl(
    fmt: &ParagraphFormatting,
    revisions: Option<&RevisionWriterData>,
) -> Result<Vec<u8>, DocWriteError> {
    let mut grp = build_papx_grpprl(fmt);
    if let Some(revision) = &fmt.formatting_revision {
        let revisions = revisions.ok_or_else(|| {
            DocWriteError::InvalidData("DOC paragraph revision author was not indexed".to_string())
        })?;
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            DocWriteError::InvalidData("DOC paragraph revision author was not indexed".to_string())
        })?;
        grp.extend_from_slice(&SPRM_P_PROP_RMARK_CURRENT.to_le_bytes());
        grp.push(7);
        grp.push(1);
        grp.extend_from_slice(&author_index.to_le_bytes());
        grp.extend_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
    }
    if let Some(revision) = &fmt.numbering_revision {
        let revisions = revisions.ok_or_else(|| {
            DocWriteError::InvalidData("DOC numbering revision author was not indexed".to_string())
        })?;
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            DocWriteError::InvalidData("DOC numbering revision author was not indexed".to_string())
        })?;
        let units = revision.format_string.encode_utf16().collect::<Vec<_>>();
        if units.len() > 31
            || revision
                .placeholder_positions
                .iter()
                .any(|position| usize::from(*position) > units.len())
        {
            return Err(DocWriteError::InvalidData(
                "DOC numbering revision format or placeholder exceeds NumRM limits".to_string(),
            ));
        }
        let mut numrm = [0u8; 128];
        numrm[0] = u8::from(revision.was_numbered);
        numrm[2..4].copy_from_slice(&author_index.to_le_bytes());
        numrm[4..8].copy_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        numrm[8..17].copy_from_slice(&revision.placeholder_positions);
        numrm[17..26].copy_from_slice(&revision.number_formats);
        for (index, number) in revision.numbers.iter().enumerate() {
            let offset = 28 + index * 4;
            numrm[offset..offset + 4].copy_from_slice(&number.to_le_bytes());
        }
        numrm[64..66].copy_from_slice(&(units.len() as u16).to_le_bytes());
        for (index, unit) in units.into_iter().enumerate() {
            let offset = 66 + index * 2;
            numrm[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
        }
        grp.extend_from_slice(&SPRM_P_NUM_RM.to_le_bytes());
        grp.push(128);
        grp.extend_from_slice(&numrm);
    }
    Ok(grp)
}

fn append_table_depth_sprms(grp: &mut Vec<u8>) {
    grp.extend_from_slice(&SPRM_P_F_IN_TABLE.to_le_bytes());
    grp.push(1);
    grp.extend_from_slice(&SPRM_P_ITAP.to_le_bytes());
    grp.extend_from_slice(&1u32.to_le_bytes());
}

fn build_table_row_papx_grpprl(
    formatting: &super::tap::TableRow,
) -> Result<Vec<u8>, DocWriteError> {
    let mut grp = Vec::new();
    append_table_depth_sprms(&mut grp);
    grp.extend_from_slice(&SPRM_P_F_TTP.to_le_bytes());
    grp.push(1);
    grp.extend_from_slice(
        &super::tap::generate_row_sprms(formatting)
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))?,
    );
    Ok(grp)
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
    fn tables_round_trip_through_both_output_paths() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("Before table").unwrap();
        let table = writer.add_table(2, 2).unwrap();
        writer
            .set_table_cell_paragraph_runs(
                table,
                0,
                0,
                vec![(
                    "A😀".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        ..CharacterFormatting::default()
                    },
                )],
                ParagraphFormatting::default(),
            )
            .unwrap();
        writer
            .append_table_cell_paragraph_runs(
                table,
                0,
                0,
                vec![(
                    "continued".to_string(),
                    CharacterFormatting {
                        italic: Some(true),
                        ..CharacterFormatting::default()
                    },
                )],
                ParagraphFormatting::default(),
            )
            .unwrap();
        writer.set_table_cell_text(table, 0, 1, "B").unwrap();
        writer
            .set_table_row_formatting(
                table,
                0,
                crate::doc::writer::TableRow {
                    cells: vec![
                        crate::doc::writer::TableCell {
                            width: 2880,
                            merged: false,
                            vertical_merge: crate::doc::parts::tap::VerticalMergeStatus::First,
                            vertical_alignment: crate::doc::parts::tap::VerticalAlignment::Center,
                            text_direction: crate::doc::parts::tap::TextDirection::TbRl,
                            fit_text: true,
                            no_wrap: true,
                            hide_mark: true,
                            borders: crate::doc::parts::tap::CellBorders {
                                top: Some(crate::doc::parts::tap::BorderStyle {
                                    width: 8,
                                    color: Some((255, 0, 0)),
                                    border_type: crate::doc::parts::tap::BorderType::Single,
                                    spacing: 2,
                                    shadow: true,
                                    frame: false,
                                }),
                                ..crate::doc::parts::tap::CellBorders::default()
                            },
                        },
                        crate::doc::writer::TableCell {
                            width: 5760,
                            merged: true,
                            ..crate::doc::writer::TableCell::default()
                        },
                    ],
                    height: 360,
                    is_header: true,
                    allow_break: true,
                },
            )
            .unwrap();
        writer.set_table_cell_text(table, 1, 0, "C").unwrap();
        writer
            .set_table_row_formatting(
                table,
                1,
                crate::doc::writer::TableRow {
                    cells: vec![
                        crate::doc::writer::TableCell {
                            width: 4320,
                            merged: false,
                            vertical_merge: crate::doc::parts::tap::VerticalMergeStatus::Merged,
                            ..crate::doc::writer::TableCell::default()
                        },
                        crate::doc::writer::TableCell {
                            width: 4320,
                            merged: false,
                            ..crate::doc::writer::TableCell::default()
                        },
                    ],
                    height: -480,
                    is_header: false,
                    allow_break: false,
                },
            )
            .unwrap();
        let second_table = writer.add_table(1, 1).unwrap();
        writer
            .set_table_cell_text(second_table, 0, 0, "Separate")
            .unwrap();

        let assert_document = |document: crate::doc::Document| {
            let tables = document.tables().unwrap();
            assert_eq!(tables.len(), 2);
            assert_eq!(tables[0].row_count().unwrap(), 2);
            assert_eq!(tables[0].column_count().unwrap(), 2);
            let rows = tables[0].rows().unwrap();
            assert_eq!(rows[0].properties().unwrap().cell_count, 2);
            assert_eq!(
                rows[0].properties().unwrap().cell_boundaries,
                [0, 2880, 8640]
            );
            assert_eq!(rows[0].properties().unwrap().row_height, Some(360));
            assert!(rows[0].properties().unwrap().is_header_row);
            assert!(!rows[0].properties().unwrap().allow_row_break);
            assert_eq!(
                rows[0].cells().unwrap()[0].text().unwrap(),
                "A😀\ncontinued"
            );
            assert_eq!(rows[0].cells().unwrap()[0].paragraphs().unwrap().len(), 2);
            let cell_paragraphs = rows[0].cells().unwrap()[0].paragraphs().unwrap();
            assert_eq!(cell_paragraphs[0].runs().unwrap()[0].bold(), Some(true));
            assert_eq!(cell_paragraphs[1].runs().unwrap()[0].italic(), Some(true));
            assert_eq!(rows[0].cells().unwrap()[1].text().unwrap(), "B");
            let first_cell_properties = rows[0].cells().unwrap()[0].properties().unwrap().clone();
            assert_eq!(
                first_cell_properties.merge_status,
                crate::doc::parts::tap::CellMergeStatus::First
            );
            assert_eq!(first_cell_properties.preferred_width.unwrap().value, 2880);
            assert_eq!(
                first_cell_properties.vertical_merge_status,
                crate::doc::parts::tap::VerticalMergeStatus::First
            );
            assert_eq!(
                first_cell_properties.vertical_alignment,
                crate::doc::parts::tap::VerticalAlignment::Center
            );
            assert_eq!(
                first_cell_properties.text_direction,
                crate::doc::parts::tap::TextDirection::TbRl
            );
            assert!(first_cell_properties.fit_text);
            assert!(first_cell_properties.no_wrap);
            assert!(first_cell_properties.hide_mark);
            let top_border = first_cell_properties.borders.top.unwrap();
            assert_eq!(top_border.width, 8);
            assert_eq!(top_border.color, Some((255, 0, 0)));
            assert_eq!(
                top_border.border_type,
                crate::doc::parts::tap::BorderType::Single
            );
            assert_eq!(top_border.spacing, 2);
            assert!(top_border.shadow);
            assert!(!top_border.frame);
            assert_eq!(
                rows[0].cells().unwrap()[1]
                    .properties()
                    .unwrap()
                    .merge_status,
                crate::doc::parts::tap::CellMergeStatus::Merged
            );
            assert_eq!(rows[1].cells().unwrap()[0].text().unwrap(), "C");
            assert_eq!(rows[1].cells().unwrap()[1].text().unwrap(), "");
            assert_eq!(rows[1].properties().unwrap().row_height, Some(-480));
            assert!(!rows[1].properties().unwrap().allow_row_break);
            assert_eq!(
                rows[1].cells().unwrap()[0]
                    .properties()
                    .unwrap()
                    .vertical_merge_status,
                crate::doc::parts::tap::VerticalMergeStatus::Merged
            );
            assert_eq!(
                tables[1].rows().unwrap()[0].cells().unwrap()[0]
                    .text()
                    .unwrap(),
                "Separate"
            );
            assert!(document.text().unwrap().ends_with('\r'));
            let element_table = document
                .elements()
                .unwrap()
                .into_iter()
                .find_map(|element| match element {
                    crate::doc::DocElement::Table(table) => Some(table),
                    crate::doc::DocElement::Paragraph(_) => None,
                })
                .unwrap();
            assert_eq!(
                element_table.properties().unwrap().cell_boundaries,
                [0, 2880, 8640]
            );
        };

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        assert_document(package.document().unwrap());

        let path = std::env::temp_dir().join(format!(
            "litchi-doc-table-{}-{}.doc",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        writer.save(&path).unwrap();
        let mut package = crate::doc::Package::open(&path).unwrap();
        assert_document(package.document().unwrap());
        std::fs::remove_file(path).unwrap();
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
        assert!(!cursor.into_inner().is_empty());
    }

    #[test]
    fn test_empty_document_write() {
        let mut writer = DocWriter::new();
        let mut cursor = Cursor::new(Vec::new());
        let result = writer.write_to(&mut cursor);
        assert!(result.is_ok());
        let data = cursor.into_inner();
        assert!(!data.is_empty());
        let mut package = crate::doc::Package::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(package.document().unwrap().text().unwrap(), "\r");
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
        assert_eq!(ls, LineSpacing::single());
        assert_eq!(ls.dya_line, 240);
        assert!(ls.is_multiple);
    }

    #[test]
    fn test_line_spacing_constructors_and_sprm_encoding() {
        let cases = [
            (LineSpacing::single(), [0xf0, 0x00, 0x01, 0x00]),
            (LineSpacing::one_and_half(), [0x68, 0x01, 0x01, 0x00]),
            (LineSpacing::double(), [0xe0, 0x01, 0x01, 0x00]),
            (
                LineSpacing::multiple_240ths(300).unwrap(),
                [0x2c, 0x01, 0x01, 0x00],
            ),
            (
                LineSpacing::at_least_twips(240).unwrap(),
                [0xf0, 0x00, 0x00, 0x00],
            ),
            (
                LineSpacing::exact_twips(240).unwrap(),
                [0x10, 0xff, 0x00, 0x00],
            ),
        ];

        for (line_spacing, operand) in cases {
            let formatting = ParagraphFormatting {
                line_spacing: Some(line_spacing),
                ..ParagraphFormatting::default()
            };
            let mut expected = SPRM_P_DYA_LINE.to_le_bytes().to_vec();
            expected.extend_from_slice(&operand);
            assert_eq!(build_papx_grpprl(&formatting), expected);
        }

        assert_eq!(LineSpacing::exact_twips(32_768).unwrap().dya_line, i16::MIN);
        assert!(LineSpacing::multiple_240ths(0).is_err());
        assert!(LineSpacing::multiple_240ths(32_768).is_err());
        assert!(LineSpacing::at_least_twips(0).is_err());
        assert!(LineSpacing::at_least_twips(32_768).is_err());
        assert!(LineSpacing::exact_twips(0).is_err());
        assert!(LineSpacing::exact_twips(32_769).is_err());
    }

    #[test]
    fn test_paragraph_formatting_writer_reader_round_trip() {
        let mut writer = DocWriter::new();
        writer
            .add_formatted_paragraph(
                "Exactly spaced",
                ParagraphFormatting {
                    alignment: Some(1),
                    space_before: Some(120),
                    line_spacing: Some(LineSpacing::exact_twips(360).unwrap()),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();
        writer
            .add_formatted_paragraph(
                "Double spaced",
                ParagraphFormatting {
                    line_spacing: Some(LineSpacing::double()),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            super::super::super::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();

        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].text().unwrap(), "Exactly spaced");
        assert_eq!(
            paragraphs[0].properties().justification,
            crate::doc::parts::pap::Justification::Center
        );
        assert_eq!(paragraphs[0].properties().space_before, Some(120));
        assert_eq!(paragraphs[0].properties().line_spacing, Some(-360));
        assert_eq!(
            paragraphs[0].properties().line_spacing_type,
            crate::doc::parts::pap::LineSpacingType::Exactly
        );
        assert_eq!(paragraphs[1].text().unwrap(), "Double spaced");
        assert_eq!(paragraphs[1].properties().line_spacing, Some(480));
        assert_eq!(
            paragraphs[1].properties().line_spacing_type,
            crate::doc::parts::pap::LineSpacingType::Double
        );
    }

    #[test]
    fn supplementary_unicode_uses_utf16_code_unit_character_positions() {
        assert_eq!(utf16_code_unit_len("A😀𝄞").unwrap(), 5);

        let mut writer = DocWriter::new();
        writer
            .add_paragraph_runs(
                vec![
                    (
                        "A😀".to_string(),
                        CharacterFormatting {
                            bold: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        "B𝄞C".to_string(),
                        CharacterFormatting {
                            italic: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                ],
                ParagraphFormatting::default(),
            )
            .unwrap();
        writer.add_paragraph("After 🦀").unwrap();
        writer
            .add_paragraph("😀\u{13} HYPERLINK \"https://example.test\" \u{14}link\u{15}")
            .unwrap();
        writer.set_odd_header("Header 😀");
        writer.set_odd_footer("Footer 𝄞");
        writer.add_footnote(FootnoteEntry::new(1, "Footnote 🦀", 1));
        writer.add_endnote(FootnoteEntry::new(2, "Endnote 😀", 1));

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            super::super::super::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();

        let paragraphs = document.paragraphs().unwrap();
        assert_eq!(paragraphs[0].text().unwrap(), "A😀B𝄞C\u{2}\u{2}");
        assert_eq!(paragraphs[1].text().unwrap(), "After 🦀");
        let fields = document.fields_table().unwrap().main_document_fields();
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].start_cp, 21);
        assert_eq!(
            fields[0].field_type,
            crate::doc::parts::fields::FieldType::Hyperlink
        );
        let headers = document.headers().unwrap();
        assert_eq!(headers.len(), 1, "{headers:?}");
        assert!(
            headers
                .iter()
                .any(|header| header.text().contains("Header 😀")),
            "{headers:?}"
        );
        let footers = document.footers().unwrap();
        assert_eq!(footers.len(), 1, "{footers:?}");
        assert!(
            footers
                .iter()
                .any(|footer| footer.text().contains("Footer 𝄞")),
            "{footers:?}"
        );
        let footnotes = document.footnotes().unwrap();
        assert_eq!(footnotes[0].number, 1);
        assert!(footnotes[0].text().contains("Footnote 🦀"));
        let endnotes = document.endnotes().unwrap();
        assert_eq!(endnotes[0].number, 1);
        assert!(endnotes[0].text().contains("Endnote 😀"));
    }

    #[test]
    fn comments_round_trip_with_other_subdocuments() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("Main 😀").unwrap();
        writer.add_footnote(FootnoteEntry::new(0, "Footnote", 1));
        writer.add_comment(
            CommentEntry::new(1, "Review 🦀", "Alice 😀", "A😀")
                .with_range(2, 6)
                .with_extended_metadata(crate::doc::CommentExtendedMetadata {
                    modified_at: Some(crate::doc::CommentDateTime {
                        year: 2026,
                        month: 7,
                        day: 15,
                        hour: 14,
                        minute: 30,
                        weekday: 3,
                    }),
                    depth: 0,
                    parent_index: None,
                    is_ink: false,
                }),
        );
        writer.add_comment(
            CommentEntry::new(3, "Second review", "Alice 😀", "AL")
                .with_range(0, 7)
                .with_extended_metadata(crate::doc::CommentExtendedMetadata {
                    modified_at: None,
                    depth: 1,
                    parent_index: Some(0),
                    is_ink: true,
                }),
        );
        writer.add_endnote(FootnoteEntry::new(2, "Endnote", 1));
        writer.set_odd_header("Header");

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();

        assert_eq!(document.footnotes().unwrap().len(), 1);
        assert_eq!(document.headers().unwrap().len(), 1);
        assert_eq!(document.endnotes().unwrap().len(), 1);
        let comments = document.comments().unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].author, "Alice 😀");
        assert_eq!(comments[0].initials, "A😀");
        assert_eq!(comments[0].bookmark_tag, Some(0));
        assert_eq!(
            (comments[0].range_start, comments[0].range_end),
            (Some(2), Some(6))
        );
        let first_metadata = comments[0].extended_metadata.unwrap();
        assert_eq!(first_metadata.depth, 0);
        assert_eq!(first_metadata.parent_index, None);
        assert_eq!(
            first_metadata.modified_at,
            Some(crate::doc::CommentDateTime {
                year: 2026,
                month: 7,
                day: 15,
                hour: 14,
                minute: 30,
                weekday: 3,
            })
        );
        assert!(comments[0].text().contains("Review 🦀"));
        assert_eq!(comments[0].paragraphs().unwrap().len(), 1);
        assert_eq!(comments[1].author, "Alice 😀");
        assert_eq!(comments[1].initials, "AL");
        assert_eq!(
            (comments[1].range_start, comments[1].range_end),
            (Some(0), Some(7))
        );
        assert_eq!(comments[1].extended_metadata.unwrap().parent_index, Some(0));
        assert!(comments[1].extended_metadata.unwrap().is_ink);
        assert!(comments[1].text().contains("Second review"));

        let path = std::env::temp_dir().join(format!(
            "litchi-doc-comments-{}-{}.doc",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        writer.save(&path).unwrap();
        let mut package = crate::doc::Package::open(&path).unwrap();
        let comments = package.document().unwrap().comments().unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(
            (comments[0].range_start, comments[0].range_end),
            (Some(2), Some(6))
        );
        assert_eq!(comments[1].extended_metadata.unwrap().parent_index, Some(0));
        std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn rejects_comment_metadata_outside_binary_limits() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("Main").unwrap();
        writer.add_comment(CommentEntry::new(0, "Body", "Author", "0123456789"));

        let error = writer.write_to(&mut Cursor::new(Vec::new())).unwrap_err();
        assert!(error.to_string().contains("at most nine"));
    }

    #[test]
    fn rejects_invalid_comment_ranges_timestamps_and_reply_trees() {
        let write_error = |entry: CommentEntry| {
            let mut writer = DocWriter::new();
            writer.add_paragraph("Main").unwrap();
            writer.add_comment(entry);
            writer
                .write_to(&mut Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string()
        };

        let error = write_error(CommentEntry::new(0, "Body", "Author", "A").with_range(4, 2));
        assert!(error.contains("range must be ordered"));

        let error = write_error(
            CommentEntry::new(0, "Body", "Author", "A").with_extended_metadata(
                crate::doc::CommentExtendedMetadata {
                    modified_at: Some(crate::doc::CommentDateTime {
                        year: 2026,
                        month: 13,
                        day: 1,
                        hour: 0,
                        minute: 0,
                        weekday: 0,
                    }),
                    depth: 0,
                    parent_index: None,
                    is_ink: false,
                },
            ),
        );
        assert!(error.contains("DTTM"));

        let error = write_error(
            CommentEntry::new(0, "Body", "Author", "A").with_extended_metadata(
                crate::doc::CommentExtendedMetadata {
                    modified_at: None,
                    depth: 1,
                    parent_index: Some(0),
                    is_ink: false,
                },
            ),
        );
        assert!(error.contains("pre-order"));
    }

    #[test]
    fn standard_bookmarks_round_trip_through_both_output_paths() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("Main text").unwrap();
        writer.add_bookmark(BookmarkEntry::new("Outer", 2, 5));
        writer.add_bookmark(
            BookmarkEntry::new("_Cell", 0, 8)
                .with_native_export(false)
                .with_column_range(1, 3),
        );

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let bookmarks = package.document().unwrap().bookmarks().unwrap();
        assert_eq!(bookmarks.len(), 2);
        assert_eq!(bookmarks[0].name, "_Cell");
        assert_eq!((bookmarks[0].start, bookmarks[0].end), (0, 8));
        assert_eq!(bookmarks[0].column_range, Some((1, 3)));
        assert!(!bookmarks[0].is_native);
        assert_eq!(bookmarks[1].name, "Outer");
        assert_eq!((bookmarks[1].start, bookmarks[1].end), (2, 5));

        let path = std::env::temp_dir().join(format!(
            "litchi-doc-bookmarks-{}-{}.doc",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        writer.save(&path).unwrap();
        let mut package = crate::doc::Package::open(&path).unwrap();
        assert_eq!(package.document().unwrap().bookmarks().unwrap(), bookmarks);
        std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn rejects_invalid_standard_bookmarks() {
        let write_error = |entries: Vec<BookmarkEntry>| {
            let mut writer = DocWriter::new();
            writer.add_paragraph("Main").unwrap();
            for entry in entries {
                writer.add_bookmark(entry);
            }
            writer
                .write_to(&mut Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string()
        };
        assert!(write_error(vec![BookmarkEntry::new("", 0, 1)]).contains("names"));
        assert!(
            write_error(vec![
                BookmarkEntry::new("Same", 0, 1),
                BookmarkEntry::new("Same", 1, 2),
            ])
            .contains("unique")
        );
        assert!(write_error(vec![BookmarkEntry::new("Range", 4, 2)]).contains("range"));
        assert!(
            write_error(vec![
                BookmarkEntry::new("Column", 0, 1).with_column_range(3, 2)
            ])
            .contains("column")
        );
    }

    #[test]
    fn tracked_text_revisions_round_trip_through_both_output_paths() {
        let timestamp = crate::doc::CommentDateTime {
            year: 2026,
            month: 7,
            day: 15,
            hour: 14,
            minute: 30,
            weekday: 3,
        };
        let mut writer = DocWriter::new();
        writer.set_section_formatting_revision(
            FormattingRevision::new("Section Editor").with_timestamp(timestamp),
        );
        writer
            .add_paragraph_runs(
                vec![
                    (
                        "inserted ".to_string(),
                        CharacterFormatting {
                            insertion_revision: Some(
                                TextRevision::new("Alice 😀")
                                    .with_timestamp(timestamp)
                                    .with_reason(crate::doc::RevisionReason::from_raw(42).unwrap())
                                    .with_revision_save_id(0x11223344),
                            ),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        "deleted".to_string(),
                        CharacterFormatting {
                            deletion_revision: Some(
                                TextRevision::new("Bob")
                                    .with_id(7)
                                    .with_revision_save_id(0x55667788),
                            ),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        " formatted".to_string(),
                        CharacterFormatting {
                            bold: Some(true),
                            formatting_revision: Some(
                                FormattingRevision::new("张三")
                                    .with_timestamp(timestamp)
                                    .with_reason(crate::doc::RevisionReason::APPLIED_STYLE)
                                    .with_revision_save_id(0x99AABBCC),
                            ),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        "\u{13}".to_string(),
                        CharacterFormatting {
                            special: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        " LISTNUM ".to_string(),
                        CharacterFormatting {
                            field_vanish: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        "\u{14}".to_string(),
                        CharacterFormatting {
                            special: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        "12.".to_string(),
                        CharacterFormatting {
                            display_field_revision: Some(
                                DisplayFieldRevision::new("Field Editor", "11.")
                                    .with_timestamp(timestamp),
                            ),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (
                        "\u{15}".to_string(),
                        CharacterFormatting {
                            special: Some(true),
                            ..CharacterFormatting::default()
                        },
                    ),
                ],
                ParagraphFormatting {
                    alignment: Some(1),
                    formatting_revision: Some(
                        FormattingRevision::new("Paragraph Editor").with_timestamp(timestamp),
                    ),
                    numbering_revision_list_applied: Some(true),
                    numbering_revision: Some(NumberingRevision {
                        was_numbered: true,
                        placeholder_positions: [1, 0, 0, 0, 0, 0, 0, 0, 0],
                        numbers: [12, 0, 0, 0, 0, 0, 0, 0, 0],
                        ..NumberingRevision::new("Numbering Editor", "%.").with_timestamp(timestamp)
                    }),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        assert_eq!(
            document.revision_authors(),
            [
                "Unknown",
                "Section Editor",
                "Paragraph Editor",
                "Numbering Editor",
                "Alice 😀",
                "Bob",
                "张三",
                "Field Editor"
            ]
        );
        let section_revision = &document.section_revisions()[0];
        assert_eq!(section_revision.start, 0);
        assert!(section_revision.end > section_revision.start);
        assert_eq!(section_revision.author, "Section Editor");
        assert_eq!(section_revision.timestamp, Some(timestamp));
        let paragraphs = document.paragraphs().unwrap();
        let paragraph_revision = paragraphs[0].formatting_revision().unwrap();
        assert_eq!(paragraph_revision.author, "Paragraph Editor");
        assert_eq!(paragraph_revision.timestamp, Some(timestamp));
        assert_eq!(paragraphs[0].numbering_revision_list_applied(), Some(true));
        let numbering_revision = paragraphs[0].numbering_revision().unwrap();
        assert_eq!(numbering_revision.author, "Numbering Editor");
        assert_eq!(numbering_revision.timestamp, Some(timestamp));
        assert!(numbering_revision.was_numbered);
        assert_eq!(numbering_revision.placeholder_positions[0], 1);
        assert_eq!(numbering_revision.numbers[0], 12);
        assert_eq!(numbering_revision.format_string, "%.");
        let runs = paragraphs[0].runs().unwrap();
        let insertion = runs
            .iter()
            .find_map(|run| run.insertion_revision())
            .unwrap();
        assert_eq!(insertion.author, "Alice 😀");
        assert_eq!(insertion.timestamp, Some(timestamp));
        assert_eq!(insertion.reason.unwrap().raw(), 42);
        assert_eq!(insertion.revision_id, Some(42));
        assert_eq!(insertion.revision_save_id, Some(0x11223344));
        let deletion = runs.iter().find_map(|run| run.deletion_revision()).unwrap();
        assert_eq!(deletion.author, "Bob");
        assert_eq!(deletion.timestamp, None);
        assert_eq!(deletion.reason.unwrap().raw(), 7);
        assert_eq!(deletion.revision_id, Some(7));
        assert_eq!(deletion.revision_save_id, Some(0x55667788));
        let formatting = runs
            .iter()
            .find_map(|run| run.formatting_revision())
            .unwrap();
        assert_eq!(formatting.kind, crate::doc::RevisionKind::Formatting);
        assert_eq!(formatting.author, "张三");
        assert_eq!(formatting.timestamp, Some(timestamp));
        assert_eq!(
            formatting.reason,
            Some(crate::doc::RevisionReason::APPLIED_STYLE)
        );
        assert_eq!(formatting.revision_id, Some(1));
        assert_eq!(formatting.revision_save_id, Some(0x99AABBCC));
        let display_field = runs
            .iter()
            .find_map(|run| run.display_field_revision())
            .unwrap();
        assert_eq!(display_field.author, "Field Editor");
        assert_eq!(display_field.timestamp, Some(timestamp));
        assert_eq!(display_field.previous_result, "11.");

        let path = std::env::temp_dir().join(format!(
            "litchi-doc-revisions-{}-{}.doc",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        writer.save(&path).unwrap();
        let mut package = crate::doc::Package::open(&path).unwrap();
        let document = package.document().unwrap();
        assert_eq!(
            document.revision_authors(),
            [
                "Unknown",
                "Section Editor",
                "Paragraph Editor",
                "Numbering Editor",
                "Alice 😀",
                "Bob",
                "张三",
                "Field Editor"
            ]
        );
        assert_eq!(document.section_revisions()[0].author, "Section Editor");
        assert!(
            document.paragraphs().unwrap()[0]
                .formatting_revision()
                .is_some()
        );
        assert!(
            document.paragraphs().unwrap()[0]
                .numbering_revision()
                .is_some()
        );
        assert!(
            document.paragraphs().unwrap()[0]
                .runs()
                .unwrap()
                .iter()
                .any(|run| run.deletion_revision().is_some())
        );
        assert!(
            document.paragraphs().unwrap()[0]
                .runs()
                .unwrap()
                .iter()
                .any(|run| run.formatting_revision().is_some())
        );
        assert!(
            document.paragraphs().unwrap()[0]
                .runs()
                .unwrap()
                .iter()
                .any(|run| run.display_field_revision().is_some())
        );
        std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn rejects_invalid_writer_revision_metadata() {
        let error_for = |formatting: CharacterFormatting| {
            let mut writer = DocWriter::new();
            writer
                .add_paragraph_runs(
                    vec![("text".to_string(), formatting)],
                    ParagraphFormatting::default(),
                )
                .unwrap();
            writer
                .write_to(&mut Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string()
        };
        let both = CharacterFormatting {
            insertion_revision: Some(TextRevision::new("Alice")),
            deletion_revision: Some(TextRevision::new("Alice")),
            ..CharacterFormatting::default()
        };
        assert!(error_for(both).contains("both an insertion and a deletion"));

        let invalid_reason = CharacterFormatting {
            insertion_revision: Some(TextRevision::new("Alice").with_id(0x002C)),
            ..CharacterFormatting::default()
        };
        assert!(error_for(invalid_reason).contains("reason code is undefined"));

        let conflicting_reason = CharacterFormatting {
            insertion_revision: Some(
                TextRevision::new("Alice")
                    .with_id(1)
                    .with_reason(crate::doc::RevisionReason::NORMAL_EDIT),
            ),
            ..CharacterFormatting::default()
        };
        assert!(error_for(conflicting_reason).contains("conflicting"));

        let conflicting_formatting_reason = CharacterFormatting {
            insertion_revision: Some(
                TextRevision::new("Alice").with_reason(crate::doc::RevisionReason::NORMAL_EDIT),
            ),
            formatting_revision: Some(
                FormattingRevision::new("Alice")
                    .with_reason(crate::doc::RevisionReason::APPLIED_STYLE),
            ),
            ..CharacterFormatting::default()
        };
        assert!(error_for(conflicting_formatting_reason).contains("insertion and formatting"));

        let invalid_time = CharacterFormatting {
            insertion_revision: Some(TextRevision::new("Alice").with_timestamp(
                crate::doc::CommentDateTime {
                    year: 2026,
                    month: 13,
                    day: 1,
                    hour: 0,
                    minute: 0,
                    weekday: 0,
                },
            )),
            ..CharacterFormatting::default()
        };
        assert!(error_for(invalid_time).contains("timestamp"));

        let mut writer = DocWriter::new();
        writer.set_section_formatting_revision(FormattingRevision::new("Editor").with_timestamp(
            crate::doc::CommentDateTime {
                year: 2026,
                month: 0,
                day: 1,
                hour: 0,
                minute: 0,
                weekday: 0,
            },
        ));
        writer.add_paragraph("text").unwrap();
        assert!(
            writer
                .write_to(&mut Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string()
                .contains("timestamp")
        );

        let mut writer = DocWriter::new();
        writer
            .add_paragraph_runs(
                vec![("text".to_string(), CharacterFormatting::default())],
                ParagraphFormatting {
                    numbering_revision: Some(NumberingRevision::new("Alice", "x".repeat(32))),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();
        assert!(
            writer
                .write_to(&mut Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string()
                .contains("NumRM")
        );

        let invalid_display = CharacterFormatting {
            display_field_revision: Some(DisplayFieldRevision::new("Alice", "x".repeat(16))),
            ..CharacterFormatting::default()
        };
        assert!(error_for(invalid_display).contains("LISTNUM"));
    }

    #[test]
    fn list_tables_round_trip_through_fib_indices() {
        let mut writer = DocWriter::new();
        let mut list = crate::doc::writer::numbering::ListStructure::new(42);
        let mut level = crate::doc::writer::numbering::ListLevel::new(
            3,
            crate::doc::writer::numbering::NumberFormat::Decimal,
        );
        level.number_text = "%1.😀".to_string();
        list.add_level(level);
        writer.add_list(list);
        writer.add_list_override(crate::doc::writer::numbering::ListFormatOverride::new(
            42, 1,
        ));
        writer
            .add_paragraph_runs(
                vec![("List item".to_string(), CharacterFormatting::default())],
                ParagraphFormatting {
                    ilvl: Some(0),
                    ilfo: Some(1),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        let tables = document.list_tables().unwrap();
        assert_eq!(tables.structures().len(), 1);
        assert_eq!(tables.overrides().len(), 1);
        assert_eq!(tables.structures()[0].levels[0].number_text, "%1.😀");

        let paragraphs = document.paragraphs().unwrap();
        let info = document.paragraph_list_info(&paragraphs[0]).unwrap();
        assert_eq!(info.start_at, 3);
        assert_eq!(info.number_text, "%1.😀");
    }

    #[test]
    fn test_add_table_invalid_dimensions() {
        let mut writer = DocWriter::new();
        assert!(writer.add_table(0, 3).is_err());
        assert!(writer.add_table(2, 0).is_err());
        assert!(writer.add_table(0, 0).is_err());
        assert!(writer.add_table(1, 64).is_err());
    }

    #[test]
    fn test_set_table_cell_invalid_indices() {
        let mut writer = DocWriter::new();
        let idx = writer.add_table(2, 2).unwrap();
        assert!(writer.set_table_cell_text(idx, 2, 0, "Invalid").is_err());
        assert!(writer.set_table_cell_text(idx, 0, 2, "Invalid").is_err());
        assert!(writer.set_table_cell_text(999, 0, 0, "Invalid").is_err());
    }

    #[test]
    fn rejects_invalid_table_row_formatting() {
        let mut writer = DocWriter::new();
        let table = writer.add_table(2, 2).unwrap();
        let one_cell = crate::doc::writer::TableRow {
            cells: vec![crate::doc::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::doc::writer::TableCell::default()
            }],
            height: 0,
            is_header: false,
            allow_break: true,
        };
        assert!(writer.set_table_row_formatting(table, 0, one_cell).is_err());

        let invalid_merge = crate::doc::writer::TableRow {
            cells: vec![
                crate::doc::writer::TableCell {
                    width: 1000,
                    merged: true,
                    ..crate::doc::writer::TableCell::default()
                },
                crate::doc::writer::TableCell {
                    width: 1000,
                    merged: false,
                    ..crate::doc::writer::TableCell::default()
                },
            ],
            height: 0,
            is_header: false,
            allow_break: true,
        };
        assert!(
            writer
                .set_table_row_formatting(table, 0, invalid_merge)
                .is_err()
        );

        let late_header = crate::doc::writer::TableRow {
            cells: vec![
                crate::doc::writer::TableCell {
                    width: 1000,
                    merged: false,
                    ..crate::doc::writer::TableCell::default()
                },
                crate::doc::writer::TableCell {
                    width: 1000,
                    merged: false,
                    ..crate::doc::writer::TableCell::default()
                },
            ],
            height: 0,
            is_header: true,
            allow_break: true,
        };
        writer
            .set_table_row_formatting(table, 1, late_header)
            .unwrap();
        assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());

        let mut writer = DocWriter::new();
        let table = writer.add_table(1, 1).unwrap();
        writer
            .set_table_row_formatting(
                table,
                0,
                crate::doc::writer::TableRow {
                    cells: vec![crate::doc::writer::TableCell {
                        width: 1000,
                        vertical_merge: crate::doc::parts::tap::VerticalMergeStatus::Merged,
                        ..crate::doc::writer::TableCell::default()
                    }],
                    height: 0,
                    is_header: false,
                    allow_break: true,
                },
            )
            .unwrap();
        assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
    }
}
