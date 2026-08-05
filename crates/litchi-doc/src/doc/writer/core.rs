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
//! use litchi_doc::DocWriter;
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
use super::smart_tags::{DocSmartTagEntry, SmartTagTableData};
use crate::doc::CommentDateTime;
use crate::doc::encryption::{
    DocEncryptionProfile, encrypt_document_streams_for_write, validate_writer_password,
};
use crate::doc::parts::pap::{
    AutoNumberAlignment, Border as ParagraphBorder, BorderStyle as ParagraphBorderStyle,
    Borders as ParagraphBorders, DropCap, FontAlignment, FrameAnchor, FrameHeight,
    FrameHorizontalAnchor, FrameHorizontalPosition, FrameTextFlow, FrameTextWrap,
    FrameVerticalAnchor, FrameVerticalPosition, LegacyAutoNumbering, LegacyBorderPosition,
    LegacyBorderStyle, PhysicalJustification, Shading as ParagraphShading, TabAlignment, TabLeader,
    TabStop, TextBoxTightWrap,
};
use crate::doc::parts::{list_names::ListNamesTable, list_templates::ListTemplateTable};
use crate::doc::{
    AssociatedStringSlot, DocumentAssociatedStrings, GlossaryMetadata, ProofingFeature,
    ProofingStateTable, ProofingTables, SavedByTable, SmartTagRecognizerRange,
};
use crate::sprm_operations::*;
use litchi_cfb::OleError;
use litchi_cfb::writer::OleWriter;
use std::collections::HashMap;
use zeroize::Zeroizing;

const WORD_DOCUMENT_CLSID: [u8; 16] = [
    0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
const VBA_PROJECT_STORAGE_NAME: &str = "Macros";

/// Error type for DOC writing
#[derive(Debug)]
pub enum DocWriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(OleError),
    /// MS-OVBA project authoring error
    Vba(litchi_vba::Error),
}

impl From<std::io::Error> for DocWriteError {
    fn from(err: std::io::Error) -> Self {
        DocWriteError::Io(err)
    }
}

impl From<OleError> for DocWriteError {
    fn from(err: OleError) -> Self {
        DocWriteError::Ole(err)
    }
}

impl From<litchi_vba::Error> for DocWriteError {
    fn from(err: litchi_vba::Error) -> Self {
        DocWriteError::Vba(err)
    }
}

impl std::fmt::Display for DocWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DocWriteError::Io(e) => write!(f, "I/O error: {}", e),
            DocWriteError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            DocWriteError::Ole(e) => write!(f, "OLE error: {}", e),
            DocWriteError::Vba(e) => write!(f, "VBA project error: {}", e),
        }
    }
}

impl std::error::Error for DocWriteError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::Vba(error) => Some(error),
            Self::InvalidData(_) => None,
        }
    }
}

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

const MAX_HEADER_FOOTER_PARAGRAPHS: usize = 65_535;
const MAX_HEADER_FOOTER_RUNS: usize = 65_535;
const MAX_HEADER_FIELD_DEPTH: usize = 128;

#[derive(Default)]
struct HeaderFieldState {
    separator_seen: Vec<bool>,
}

impl HeaderFieldState {
    fn observe(
        &mut self,
        character: char,
        formatting: &CharacterFormatting,
    ) -> Result<bool, DocWriteError> {
        if !matches!(character as u32, 0x0013..=0x0015) {
            return Ok(false);
        }
        if formatting.special != Some(true) {
            return Err(DocWriteError::InvalidData(
                "DOC header/footer field marker requires fSpec formatting".to_string(),
            ));
        }
        match character as u32 {
            0x0013 => {
                if self.separator_seen.len() >= MAX_HEADER_FIELD_DEPTH {
                    return Err(DocWriteError::InvalidData(
                        "DOC header/footer field nesting exceeds the limit".to_string(),
                    ));
                }
                self.separator_seen.push(false);
            },
            0x0014 => {
                let seen = self.separator_seen.last_mut().ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC header/footer field separator has no begin marker".to_string(),
                    )
                })?;
                if *seen {
                    return Err(DocWriteError::InvalidData(
                        "DOC header/footer field has duplicate separators".to_string(),
                    ));
                }
                *seen = true;
            },
            0x0015 => {
                self.separator_seen.pop().ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC header/footer field end has no begin marker".to_string(),
                    )
                })?;
            },
            _ => unreachable!(),
        }
        Ok(true)
    }

    fn finish(self) -> Result<(), DocWriteError> {
        if self.separator_seen.is_empty() {
            Ok(())
        } else {
            Err(DocWriteError::InvalidData(
                "DOC header/footer field is not terminated within its story".to_string(),
            ))
        }
    }
}

fn checked_text_fc(text_fc_start: u32, stream_length: usize) -> Result<u32, DocWriteError> {
    let stream_length = u32::try_from(stream_length).map_err(|_| {
        DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
    })?;
    text_fc_start
        .checked_add(stream_length)
        .ok_or_else(|| DocWriteError::InvalidData("DOC text stream FC range overflows".to_string()))
}

fn validate_header_footer_paragraphs(
    paragraphs: &[HeaderFooterParagraph],
) -> Result<(), DocWriteError> {
    if paragraphs.is_empty() {
        return Err(DocWriteError::InvalidData(
            "DOC header/footer story requires at least one paragraph".to_string(),
        ));
    }
    if paragraphs.len() > MAX_HEADER_FOOTER_PARAGRAPHS {
        return Err(DocWriteError::InvalidData(
            "DOC header/footer story exceeds the paragraph limit".to_string(),
        ));
    }

    let mut run_count = 0usize;
    let mut character_count = 1u32; // Inter-story guard paragraph mark.
    let mut field_state = HeaderFieldState::default();
    for paragraph in paragraphs {
        run_count = run_count.checked_add(paragraph.runs.len()).ok_or_else(|| {
            DocWriteError::InvalidData("DOC header/footer run count overflows".to_string())
        })?;
        if run_count > MAX_HEADER_FOOTER_RUNS {
            return Err(DocWriteError::InvalidData(
                "DOC header/footer story exceeds the run limit".to_string(),
            ));
        }
        for (text, formatting) in &paragraph.runs {
            if text.contains('\r') {
                return Err(DocWriteError::InvalidData(
                    "DOC header/footer run contains an embedded paragraph mark".to_string(),
                ));
            }
            for character in text.chars() {
                field_state.observe(character, formatting)?;
            }
            character_count = character_count
                .checked_add(utf16_code_unit_len(text)?)
                .ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC header/footer story CP range overflows".to_string(),
                    )
                })?;
        }
        character_count = character_count.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC header/footer story CP range overflows".to_string())
        })?;
    }
    if character_count >= 0x7FFF_FFFF {
        return Err(DocWriteError::InvalidData(
            "DOC header/footer story exceeds the MS-DOC CP limit".to_string(),
        ));
    }
    field_state.finish()
}

pub(super) fn pack_dttm(value: Option<CommentDateTime>) -> Result<u32, DocWriteError> {
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
struct HeaderStoryData {
    plcfhdd: Vec<u8>,
    fields: Vec<u8>,
    char_count: u32,
    /// Story-relative anchor CPs of header floating items with their kind
    /// (in story order, which is CP-ascending by construction).
    shape_anchor_cps: Vec<(u32, FloatingAnchorKind)>,
}

/// PlcfHdd slot of the odd page header, which Word uses as the default
/// header when the document does not use facing pages.
const HEADER_SLOT_ODD: usize = 7;
/// PlcfHdd slot of the even page header.
const HEADER_SLOT_EVEN: usize = 6;
/// PlcfHdd slot of the first page header.
const HEADER_SLOT_FIRST: usize = 10;

/// Which header a floating text box or picture is anchored in.
///
/// The writer emits the section properties each kind needs automatically:
/// even headers enable DOP `fFacingPages`, first-page headers enable SEP
/// `fTitlePage`, because appending the anchor creates the corresponding
/// header story.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocHeaderKind {
    /// Odd page header; Word's default header.
    Odd,
    /// Even page header (requires facing pages, enabled automatically).
    Even,
    /// First page header (requires a different first page, enabled
    /// automatically).
    FirstPage,
}

impl DocHeaderKind {
    /// The PlcfHdd slot holding this header kind's story.
    fn slot(self) -> usize {
        match self {
            Self::Odd => HEADER_SLOT_ODD,
            Self::Even => HEADER_SLOT_EVEN,
            Self::FirstPage => HEADER_SLOT_FIRST,
        }
    }
}

/// A floating-item anchor paragraph appended to a header's paragraphs.
struct HeaderAnchor {
    /// PlcfHdd slot of the header holding the anchor.
    slot: usize,
    /// Paragraph index within that slot's paragraph list.
    paragraph_index: usize,
    /// Which floating item the anchor belongs to.
    kind: FloatingAnchorKind,
}

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
    /// Style-sheet index of the applied character style.
    pub style_index: Option<u16>,
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
    /// Vertical offset relative to the normal baseline, in signed half-points.
    pub position: Option<crate::doc::parts::chp::CharacterPosition>,
    /// Word-breaking behavior used when this run is hyphenated.
    pub hyphenation: Option<crate::doc::parts::chp::HresiOperand>,
    /// Animated text effect applied to this run.
    pub text_effect: Option<crate::doc::parts::chp::TextEffect>,
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
    /// Formatting state retained before a tracked character-property change.
    pub preserved_properties_for_revision: Option<Box<CharacterFormatting>>,
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
        if !(1..=31_680).contains(&value) {
            return Err(DocWriteError::InvalidData(format!(
                "line-spacing multiple {value} is outside the LSPD range 1..=31680"
            )));
        }
        Ok(Self {
            dya_line: value as i16,
            is_multiple: true,
        })
    }

    /// Create minimum line spacing in twips.
    pub fn at_least_twips(value: u16) -> Result<Self, DocWriteError> {
        if !(1..=31_680).contains(&value) {
            return Err(DocWriteError::InvalidData(format!(
                "minimum line spacing {value} twips is outside the LSPD range 1..=31680"
            )));
        }
        Ok(Self {
            dya_line: value as i16,
            is_multiple: false,
        })
    }

    /// Create exact line spacing in twips.
    pub fn exact_twips(value: u16) -> Result<Self, DocWriteError> {
        if !(1..=31_680).contains(&value) {
            return Err(DocWriteError::InvalidData(format!(
                "exact line spacing {value} twips is outside the LSPD range 1..=31680"
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
    /// Style-sheet index of the applied paragraph style.
    pub style_index: Option<u16>,
    /// Alignment (0=left, 1=center, 2=right, 3=justify)
    pub alignment: Option<u8>,
    /// Explicit Word 97 physical justification for compatibility readers
    pub physical_justification: Option<PhysicalJustification>,
    /// Left indent (in twips, 1440 twips = 1 inch)
    pub left_indent: Option<i32>,
    /// Right indent (in twips)
    pub right_indent: Option<i32>,
    /// First line indent (in twips)
    pub first_line_indent: Option<i32>,
    /// Logical left indent in hundredths of a character
    pub left_indent_chars: Option<i16>,
    /// Logical right indent in hundredths of a character
    pub right_indent_chars: Option<i16>,
    /// First-line indent in hundredths of a character
    pub first_line_indent_chars: Option<i16>,
    /// Space before paragraph (in twips)
    pub space_before: Option<u16>,
    /// Space after paragraph (in twips)
    pub space_after: Option<u16>,
    /// Exclude this paragraph from line numbering
    pub no_line_numbering: Option<bool>,
    /// Space before paragraph in hundredths of a line (`-20..=31680`)
    pub space_before_lines: Option<i16>,
    /// Space after paragraph in hundredths of a line (`-20..=31680`)
    pub space_after_lines: Option<i16>,
    /// Use auto spacing for space before
    pub space_before_auto: Option<bool>,
    /// Use auto spacing for space after
    pub space_after_auto: Option<bool>,
    /// Keep a cell mark visible immediately after a nested table
    pub open_table_cell_mark: Option<bool>,
    /// Widow/orphan control
    pub widow_control: Option<bool>,
    /// Lock the paragraph frame anchor
    pub frame_anchor_locked: Option<bool>,
    /// Use East Asian line-breaking rules
    pub kinsoku: Option<bool>,
    /// Prefer word-level wrapping
    pub word_wrap: Option<bool>,
    /// Permit punctuation to overflow the line extent
    pub overflow_punctuation: Option<bool>,
    /// Compress punctuation at the beginning of a line
    pub top_line_punctuation: Option<bool>,
    /// Automatically space East Asian and Latin text
    pub auto_space_east_asian_latin: Option<bool>,
    /// Automatically space East Asian text and numbers
    pub auto_space_east_asian_numbers: Option<bool>,
    /// Vertical character alignment within a line
    pub font_alignment: Option<FontAlignment>,
    /// Direction and glyph rotation of text in a frame
    pub frame_text_flow: Option<FrameTextFlow>,
    /// Horizontal paragraph-frame position
    pub frame_horizontal_position: Option<FrameHorizontalPosition>,
    /// Vertical paragraph-frame position
    pub frame_vertical_position: Option<FrameVerticalPosition>,
    /// Paragraph-frame width in twips, where zero means automatic
    pub frame_width: Option<u16>,
    /// Reference points used by paragraph-frame coordinates
    pub frame_anchor: Option<FrameAnchor>,
    /// Explicit table membership flag
    pub in_table: Option<bool>,
    /// Mark a cell mark as a table-terminating paragraph
    pub table_terminating_paragraph: Option<bool>,
    /// Wrapping of surrounding text around the paragraph frame
    pub frame_text_wrap: Option<FrameTextWrap>,
    /// Paragraph frame height
    pub frame_height: Option<FrameHeight>,
    /// Minimum horizontal distance between frame and surrounding text
    pub frame_horizontal_text_distance: Option<i16>,
    /// Minimum vertical distance between frame and surrounding text
    pub frame_vertical_text_distance: Option<i16>,
    /// Drop-cap placement and line count
    pub drop_cap: Option<DropCap>,
    /// Disable automatic hyphenation for this paragraph
    pub no_auto_hyphenation: Option<bool>,
    /// Lay this paragraph out side-by-side with adjacent paragraphs
    pub side_by_side: Option<bool>,
    /// Keep the paragraph on one page
    pub keep: Option<bool>,
    /// Keep the paragraph with the next paragraph
    pub keep_with_next: Option<bool>,
    /// Insert a page break before this paragraph
    pub page_break_before: Option<bool>,
    /// Bi-directional paragraph
    pub bidi: Option<bool>,
    /// Follow vertical document-grid settings
    pub use_page_setup_settings: Option<bool>,
    /// Automatically adjust the right indent to the document grid
    pub adjust_right_indent: Option<bool>,
    /// Outline level (0..9)
    pub outline_level: Option<u8>,
    /// Prevent overlapping floating objects anchored to the paragraph
    pub no_allow_overlap: Option<bool>,
    /// Contextual spacing (ignore spacing between same style)
    pub contextual_spacing: Option<bool>,
    /// Mirror indents (for facing pages)
    pub mirror_indents: Option<bool>,
    /// Lines in a text box whose edges permit tight wrapping
    pub text_box_tight_wrap: Option<TextBoxTightWrap>,
    /// Paragraph borders
    pub borders: ParagraphBorders,
    /// Obsolete paragraph-border line style retained for old DOC consumers
    pub legacy_border_style: Option<LegacyBorderStyle>,
    /// Obsolete paragraph-border placement retained for old DOC consumers
    pub legacy_border_position: Option<LegacyBorderPosition>,
    /// Paragraph background shading
    pub shading: Option<ParagraphShading>,
    /// Line spacing descriptor
    pub line_spacing: Option<LineSpacing>,
    /// Existing tab-stop positions to delete, in twips
    pub tab_stops_to_delete: Vec<i32>,
    /// Tab stops to add or replace
    pub tab_stops_to_add: Vec<TabStop>,
    /// List level index (0 through 8), or 12 to skip this paragraph in list numbering
    pub ilvl: Option<u8>,
    /// Raw list override encoding (positive values are 1-based; negative encodings preserve indents)
    pub ilfo: Option<u16>,
    /// Legacy autonumber descriptor for compatibility with pre-list-table documents
    pub legacy_autonumbering: Option<LegacyAutoNumbering>,
    /// Revision save ID associated with this paragraph's formatting
    pub revision_save_id: Option<u32>,
    /// Formatting state retained before a tracked paragraph-property change
    pub preserved_properties_for_revision: Option<Box<ParagraphFormatting>>,
    /// Mark the paragraph formatting as a tracked change.
    pub formatting_revision: Option<FormattingRevision>,
    /// Whether a numbered list was applied after the previous revision.
    pub numbering_revision_list_applied: Option<bool>,
    /// Retained numbering state for a tracked numbering change.
    pub numbering_revision: Option<NumberingRevision>,
}

/// Represents a text run with formatting
#[derive(Debug, Clone)]
struct TextRun {
    /// Text content
    text: String,
    /// Character formatting
    formatting: CharacterFormatting,
    /// Index into `DocWriter::pictures` when this run is a picture
    /// (a single 0x0001 inline or 0x0008 floating picture character).
    picture_index: Option<u32>,
    /// Index into `DocWriter::shapes` when this run is a floating
    /// drawing-shape anchor (a single 0x0008 character).
    shape_index: Option<u32>,
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
            picture_index: None,
            shape_index: None,
        }]
    } else {
        runs.into_iter()
            .map(|(text, formatting)| TextRun {
                text,
                formatting,
                picture_index: None,
                shape_index: None,
            })
            .collect()
    };
    WritableParagraph { runs, formatting }
}

/// One formatted paragraph in a header or footer story.
///
/// The paragraph owns its runs so callers can build content without tying the
/// writer to temporary buffers. Paragraph marks are emitted by the writer and
/// therefore MUST NOT be embedded in run text passed to [`Self::from_runs`].
#[derive(Debug, Clone)]
pub struct HeaderFooterParagraph {
    runs: Vec<(String, CharacterFormatting)>,
    formatting: ParagraphFormatting,
}

impl HeaderFooterParagraph {
    /// Construct a plain paragraph containing one run.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            runs: vec![(text.into(), CharacterFormatting::default())],
            formatting: ParagraphFormatting::default(),
        }
    }

    /// Construct a paragraph from owned character runs and paragraph formatting.
    pub fn from_runs(
        runs: Vec<(String, CharacterFormatting)>,
        formatting: ParagraphFormatting,
    ) -> Self {
        Self { runs, formatting }
    }

    /// Construct a paragraph containing one inert Word field.
    ///
    /// The instruction is transported verbatim and is never evaluated.
    pub fn field(
        instruction: impl Into<String>,
        result: impl Into<String>,
        result_formatting: CharacterFormatting,
    ) -> Result<Self, DocWriteError> {
        let instruction = instruction.into();
        let result = result.into();
        if instruction.is_empty() {
            return Err(DocWriteError::InvalidData(
                "DOC header/footer field instruction is empty".to_string(),
            ));
        }
        if instruction
            .chars()
            .chain(result.chars())
            .any(|character| character == '\r' || matches!(character as u32, 0x0013..=0x0015))
        {
            return Err(DocWriteError::InvalidData(
                "DOC header/footer field text contains a structural control character".to_string(),
            ));
        }
        let special = CharacterFormatting {
            special: Some(true),
            ..CharacterFormatting::default()
        };
        let hidden = CharacterFormatting {
            field_vanish: Some(true),
            ..CharacterFormatting::default()
        };
        Ok(Self::from_runs(
            vec![
                ("\u{0013}".to_string(), special.clone()),
                (instruction, hidden),
                ("\u{0014}".to_string(), special.clone()),
                (result, result_formatting),
                ("\u{0015}".to_string(), special),
            ],
            ParagraphFormatting::default(),
        ))
    }

    /// Character runs in document order.
    pub fn runs(&self) -> &[(String, CharacterFormatting)] {
        &self.runs
    }

    /// Paragraph-level formatting.
    pub fn formatting(&self) -> &ParagraphFormatting {
        &self.formatting
    }
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

/// A picture queued for embedding, with its placement mode.
#[derive(Debug, Clone)]
struct WriterPicture {
    /// The picture data and display dimensions.
    picture: super::images::DocPicture,
    /// Shape id allocated at insert time (shared sequence with shapes).
    shape_id: u32,
    /// Position and wrapping when the picture floats; `None` for inline.
    floating: Option<super::images::FloatingPosition>,
}

/// A primitive drawing shape queued for the drawing layer.
#[derive(Debug, Clone)]
struct WriterShape {
    /// The shape geometry, size, and colors.
    shape: super::shapes::DocDrawingShape,
    /// Shape id allocated at insert time (shared sequence with pictures).
    shape_id: u32,
    /// Position and wrapping.
    position: super::images::FloatingPosition,
    /// Textbox story text when the shape is a text box.
    text: Option<String>,
}

/// What kind of floating content a 0x0008 anchor character refers to.
#[derive(Debug, Clone, Copy)]
enum FloatingAnchorKind {
    /// Index into `DocWriter::pictures`.
    Picture(u32),
    /// Index into `DocWriter::shapes`.
    Shape(u32),
}

/// Fully assembled DOC streams before compound-file packaging.
struct DocOutputStreams {
    word_document: Vec<u8>,
    table: Vec<u8>,
    data: Vec<u8>,
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
    /// Header/footer paragraphs (`None` means the story is not set).
    /// Indices map to plcfHdd entries (following Apache POI HeaderStories indexing):
    /// 0..5: footnote/endnote separators (unused here)
    /// 6: even header, 7: odd header, 10: first header
    /// 8: even footer, 9: odd footer, 11: first footer
    header_even: Option<Vec<HeaderFooterParagraph>>,
    header_odd: Option<Vec<HeaderFooterParagraph>>,
    header_first: Option<Vec<HeaderFooterParagraph>>,
    footer_even: Option<Vec<HeaderFooterParagraph>>,
    footer_odd: Option<Vec<HeaderFooterParagraph>>,
    footer_first: Option<Vec<HeaderFooterParagraph>>,
    /// Footnote entries
    footnotes: Vec<FootnoteEntry>,
    /// Endnote entries
    endnotes: Vec<FootnoteEntry>,
    /// Comments
    comments: Vec<CommentEntry>,
    /// Standard bookmarks
    bookmarks: Vec<BookmarkEntry>,
    /// Embedded smart-tag bookmarks and property bags.
    smart_tags: Vec<DocSmartTagEntry>,
    /// Smart-tag recognizer processing-state ranges.
    smart_tag_recognizer_ranges: Vec<SmartTagRecognizerRange>,
    /// Optional spelling and grammar proofing-state PLCFs.
    proofing_tables: ProofingTables,
    /// Mandatory fixed associated-document string table.
    associated_strings: DocumentAssociatedStrings,
    /// Optional Word 97/2000 save-history table.
    saved_by_table: Option<SavedByTable>,
    /// Optional glossary-only AutoText metadata over the main story.
    glossary_metadata: Option<GlossaryMetadata>,
    /// Optional distinct AutoText-only document attached to this template.
    attached_glossary: Option<Box<DocWriter>>,
    /// Property revision metadata for the writer's single document section
    section_formatting_revision: Option<FormattingRevision>,
    /// Explicit column geometry for the writer's single document section.
    section_columns: Option<crate::doc::section::columns::Layout>,
    /// Whether section columns are populated from right to left.
    section_right_to_left: bool,
    /// Section-wide glyph and line flow.
    section_text_flow: crate::doc::SectionTextFlow,
    /// Explicit page-border edges and placement for the single section.
    section_page_borders: Option<crate::doc::section::borders::Borders>,
    /// Numbering writer for list tables
    numbering: NumberingWriter,
    /// User-defined styles appended after the fifteen fixed style slots
    styles: Vec<super::stylesheet::DocStyleDefinition>,
    /// Inline pictures embedded via [`DocWriter::insert_picture`]
    pictures: Vec<WriterPicture>,
    /// Primitive drawing shapes embedded via [`DocWriter::insert_floating_shape`]
    shapes: Vec<WriterShape>,
    /// Text boxes anchored in the header story, in insertion order.
    header_shapes: Vec<WriterShape>,
    /// Pictures anchored in the header story, in insertion order.
    header_pictures: Vec<WriterPicture>,
    /// Anchor paragraphs appended to header paragraph lists, in insertion
    /// order (one per header floating item).
    header_anchors: Vec<HeaderAnchor>,
    /// Next shape id to allocate (shared by pictures and drawing shapes).
    next_shape_id: u32,
    /// Password-to-open settings. The password is wiped when replaced, cleared, or dropped.
    encryption: Option<DocWriterEncryption>,
    /// Complete inert MS-OVBA project written under the MS-DOC `Macros` storage.
    vba_project: Option<litchi_vba::Payload>,
}

struct DocWriterEncryption {
    profile: DocEncryptionProfile,
    password: Zeroizing<String>,
}

/// Append one textbox story (main or header) to the text stream.
///
/// Per text box: its paragraphs (each `\r`-terminated, with `\n`/`\r`/`"\r\n"`
/// as input separators) plus a trailing CR; one story-final CR is included
/// in the returned story character count. Returns the story-relative start
/// CP of each text box and the total story length (a ccp value).
fn write_textbox_story_text(
    texts: &[&str],
    text_stream: &mut Vec<u8>,
    current_cp: &mut u32,
) -> Result<(Vec<u32>, u32), DocWriteError> {
    let story_start_cp = *current_cp;
    let mut start_cps = Vec::with_capacity(texts.len());
    for text in texts {
        start_cps.push(*current_cp - story_start_cp);
        for paragraph in text.replace("\r\n", "\n").replace('\r', "\n").split('\n') {
            let para_len = utf16_code_unit_len(paragraph)?;
            for unit in paragraph.encode_utf16() {
                text_stream.extend_from_slice(&unit.to_le_bytes());
            }
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            *current_cp += para_len + 1;
        }
        // Trailing CR of this text box's text, as Word writes.
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        *current_cp += 1;
    }
    // Story-final CR, included in the ccp count.
    text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
    *current_cp += 1;
    Ok((start_cps, *current_cp - story_start_cp))
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
            smart_tags: Vec::new(),
            smart_tag_recognizer_ranges: Vec::new(),
            proofing_tables: ProofingTables::default(),
            associated_strings: DocumentAssociatedStrings::default(),
            saved_by_table: None,
            glossary_metadata: None,
            attached_glossary: None,
            section_formatting_revision: None,
            section_columns: None,
            section_right_to_left: false,
            section_text_flow: crate::doc::SectionTextFlow::default(),
            section_page_borders: None,
            numbering: NumberingWriter::new(),
            styles: Vec::new(),
            pictures: Vec::new(),
            shapes: Vec::new(),
            header_shapes: Vec::new(),
            header_pictures: Vec::new(),
            header_anchors: Vec::new(),
            next_shape_id: super::images::FIRST_PICTURE_SHAPE_ID,
            encryption: None,
            vba_project: None,
        }
    }

    /// Protect the generated document with a password-to-open profile.
    ///
    /// Validation is atomic: an invalid password or profile leaves any previous
    /// password setting unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: DocEncryptionProfile,
    ) -> Result<(), DocWriteError> {
        let password = Zeroizing::new(password.into());
        validate_writer_password(profile, password.as_str()).map_err(DocWriteError::InvalidData)?;
        self.encryption = Some(DocWriterEncryption { profile, password });
        Ok(())
    }

    /// Remove password-to-open protection and wipe the stored password.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured password-to-open profile without exposing the password.
    pub fn encryption_profile(&self) -> Option<DocEncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Configure a complete inert VBA project with safe default limits.
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<(), DocWriteError> {
        self.set_vba_with(project, &litchi_vba::Limits::default())
    }

    /// Configure a complete inert VBA project using explicit resource limits.
    ///
    /// Validation and serialization complete before writer state is changed.
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<(), DocWriteError> {
        let payload = project.finish(limits)?;
        self.put_vba(payload);
        Ok(())
    }

    /// Configure an already validated and serialized inert VBA project.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) {
        self.vba_project = Some(payload);
    }

    /// Remove the configured VBA project storage.
    pub fn clear_vba(&mut self) {
        self.vba_project = None;
    }

    /// Whether a complete VBA project is configured for output.
    pub fn has_vba(&self) -> bool {
        self.vba_project.is_some()
    }

    /// Insert or replace a spelling or grammar proofing-state table.
    ///
    /// Character positions use the concatenated DOC document-part coordinate
    /// space. The final CP ceiling is validated when output is generated.
    pub fn set_proofing_table(&mut self, table: ProofingStateTable) -> Option<ProofingStateTable> {
        self.proofing_tables.set(table)
    }

    /// Replace both optional proofing tables.
    pub fn set_proofing_tables(&mut self, tables: ProofingTables) {
        self.proofing_tables = tables;
    }

    /// Access one configured proofing table.
    pub fn proofing_table(&self, feature: ProofingFeature) -> Option<&ProofingStateTable> {
        self.proofing_tables.get(feature)
    }

    /// Remove and return one configured proofing table.
    pub fn clear_proofing_table(&mut self, feature: ProofingFeature) -> Option<ProofingStateTable> {
        self.proofing_tables.remove(feature)
    }

    /// Replace all 18 associated-document string slots.
    pub fn set_associated_strings(&mut self, strings: DocumentAssociatedStrings) {
        self.associated_strings = strings;
    }

    /// Access the associated-document string table that will be written.
    pub fn associated_strings(&self) -> &DocumentAssociatedStrings {
        &self.associated_strings
    }

    /// Replace one associated-document string slot atomically.
    pub fn set_associated_string(
        &mut self,
        slot: AssociatedStringSlot,
        value: impl Into<String>,
    ) -> Result<String, DocWriteError> {
        self.associated_strings
            .set(slot, value)
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))
    }

    /// Reset all associated-document string slots to empty strings.
    ///
    /// The mandatory `SttbfAssoc` structure is still emitted.
    pub fn reset_associated_strings(&mut self) {
        self.associated_strings = DocumentAssociatedStrings::default();
    }

    /// Configure the optional Word 97/2000 save-history table.
    pub fn set_saved_by_table(&mut self, table: SavedByTable) -> Option<SavedByTable> {
        self.saved_by_table.replace(table)
    }

    /// Access the configured save-history table.
    pub fn saved_by_table(&self) -> Option<&SavedByTable> {
        self.saved_by_table.as_ref()
    }

    /// Remove and return the configured save-history table.
    pub fn clear_saved_by_table(&mut self) -> Option<SavedByTable> {
        self.saved_by_table.take()
    }

    /// Configure this output as a glossary-only DOC.
    ///
    /// Item ranges use main-story UTF-16 character positions and may cover
    /// formatted text, tables, drawings, or pictures. The metadata's `ccpText`
    /// is checked against the generated main story before output is modified.
    pub fn set_glossary_metadata(
        &mut self,
        metadata: GlossaryMetadata,
    ) -> Option<GlossaryMetadata> {
        self.glossary_metadata.replace(metadata)
    }

    /// Access the configured glossary-only metadata.
    pub fn glossary_metadata(&self) -> Option<&GlossaryMetadata> {
        self.glossary_metadata.as_ref()
    }

    /// Return this writer to ordinary-document output.
    pub fn clear_glossary_metadata(&mut self) -> Option<GlossaryMetadata> {
        self.glossary_metadata.take()
    }

    /// Attach a distinct glossary-only document to this template.
    ///
    /// The attached writer must have [`DocWriter::set_glossary_metadata`]
    /// configured. Its main story becomes the template's AutoText story.
    ///
    /// # Errors
    ///
    /// Returns an error for nested or independently encrypted glossary
    /// documents and independent VBA projects. Those configurations cannot be
    /// represented by the shared DOC stream topology.
    pub fn set_attached_glossary(
        &mut self,
        glossary: DocWriter,
    ) -> Result<Option<DocWriter>, DocWriteError> {
        glossary.validate_as_attached_glossary()?;
        Ok(self
            .attached_glossary
            .replace(Box::new(glossary))
            .map(|previous| *previous))
    }

    /// Access the attached glossary writer.
    pub fn attached_glossary(&self) -> Option<&DocWriter> {
        self.attached_glossary.as_deref()
    }

    /// Mutably access the attached glossary writer.
    pub fn attached_glossary_mut(&mut self) -> Option<&mut DocWriter> {
        self.attached_glossary.as_deref_mut()
    }

    /// Remove and return the attached glossary writer.
    pub fn clear_attached_glossary(&mut self) -> Option<DocWriter> {
        self.attached_glossary.take().map(|glossary| *glossary)
    }

    fn validate_as_attached_glossary(&self) -> Result<(), DocWriteError> {
        if self.glossary_metadata.is_none() {
            return Err(DocWriteError::InvalidData(
                "attached DOC glossary requires glossary metadata".to_string(),
            ));
        }
        if self.attached_glossary.is_some() {
            return Err(DocWriteError::InvalidData(
                "attached DOC glossaries cannot contain another attached glossary".to_string(),
            ));
        }
        if self.encryption.is_some() {
            return Err(DocWriteError::InvalidData(
                "an attached DOC glossary cannot have independent encryption".to_string(),
            ));
        }
        if self.vba_project.is_some() {
            return Err(DocWriteError::InvalidData(
                "an attached DOC glossary cannot contain an independent VBA project".to_string(),
            ));
        }
        Ok(())
    }

    fn encryption_table_header_len(&self) -> Result<usize, DocWriteError> {
        self.encryption
            .as_ref()
            .map(|value| value.profile.table_header_len())
            .transpose()
            .map(|value| value.unwrap_or(0))
            .map_err(DocWriteError::InvalidData)
    }

    fn encrypt_output_streams(
        &self,
        word_document: &mut [u8],
        table_stream: &mut [u8],
        data_stream: &mut [u8],
    ) -> Result<(), DocWriteError> {
        let Some(encryption) = &self.encryption else {
            return Ok(());
        };
        encrypt_document_streams_for_write(
            encryption.profile,
            encryption.password.as_str(),
            word_document,
            table_stream,
            data_stream,
        )
        .map_err(DocWriteError::InvalidData)
    }

    fn populate_compound_document(
        &self,
        ole_writer: &mut OleWriter,
        word_document_stream: &[u8],
        table_stream: &[u8],
        data_stream: &[u8],
    ) -> Result<(), DocWriteError> {
        ole_writer.set_root_clsid(WORD_DOCUMENT_CLSID);

        // Preserve the conventional stream order so WordDocument occupies the
        // first regular FAT sector, followed by the table and Data streams.
        ole_writer.create_stream(&["WordDocument"], word_document_stream)?;
        ole_writer.create_stream(&["1Table"], table_stream)?;
        ole_writer.create_stream(&["Data"], data_stream)?;

        let compobj_data = crate::doc::writer::ole_metadata::generate_compobj_stream();
        let ole_data = crate::doc::writer::ole_metadata::generate_ole_stream();
        ole_writer.create_stream(&["\x01CompObj"], &compobj_data)?;
        ole_writer.create_stream(&["\x01Ole"], &ole_data)?;

        if let Some(project) = &self.vba_project {
            ole_writer.create_storage(&[VBA_PROJECT_STORAGE_NAME])?;
            project.write_into(ole_writer, &[VBA_PROJECT_STORAGE_NAME])?;
        }
        Ok(())
    }

    /// Add a custom paragraph, character, table, or numbering style.
    ///
    /// Custom styles occupy consecutive indices beginning at 15. The returned
    /// index can be used by the corresponding formatting properties.
    pub fn add_style(
        &mut self,
        style: super::stylesheet::DocStyleDefinition,
    ) -> Result<u16, DocWriteError> {
        let index = 15usize
            .checked_add(self.styles.len())
            .and_then(|index| u16::try_from(index).ok())
            .filter(|index| *index <= 0x0FFC)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC stylesheet exceeds 4093 style slots".to_string())
            })?;
        self.styles.push(style);
        Ok(index)
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
                picture_index: None,
                shape_index: None,
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
                picture_index: None,
                shape_index: None,
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
            wruns.push(TextRun {
                text,
                formatting,
                picture_index: None,
                shape_index: None,
            });
        }
        self.paragraphs.push(WritableParagraph {
            runs: wruns,
            formatting: para_fmt,
        });
        Ok(())
    }

    /// Insert an inline picture as its own paragraph.
    ///
    /// The picture is written as a single 0x0001 picture character with
    /// sprmCFSpec and sprmCPicLocation applied ([MS-DOC] 1.3); the character
    /// points to an OfficeArtWordDrawing block (PICF + OfficeArtSpContainer +
    /// OfficeArtFBSE with an embedded BLIP) in the Data stream. The image
    /// bytes are stored verbatim — no re-encoding is performed.
    pub fn insert_picture(
        &mut self,
        picture: super::images::DocPicture,
    ) -> Result<(), DocWriteError> {
        self.insert_picture_run(picture, None, "\u{0001}")
    }

    /// Insert a floating picture anchored to its own paragraph.
    ///
    /// The anchor is a single 0x0008 character with sprmCFSpec and
    /// sprmCPicLocation applied ([MS-DOC] 1.3). The picture data is stored
    /// like an inline picture's, and the anchor character position is
    /// recorded in the Main Document's PlcfSpa together with an
    /// OfficeArtContent drawing group (fcDggInfo) holding the picture-frame
    /// shape, so readers can resolve the anchor to position and image.
    pub fn insert_floating_picture(
        &mut self,
        picture: super::images::DocPicture,
        position: super::images::FloatingPosition,
    ) -> Result<(), DocWriteError> {
        self.insert_picture_run(picture, Some(position), "\u{0008}")
    }

    /// Shared tail of `insert_picture`/`insert_floating_picture`: queue the
    /// picture and append a single-character anchor paragraph.
    fn insert_picture_run(
        &mut self,
        picture: super::images::DocPicture,
        floating: Option<super::images::FloatingPosition>,
        anchor: &str,
    ) -> Result<(), DocWriteError> {
        let picture_index = u32::try_from(self.pictures.len()).map_err(|_| {
            DocWriteError::InvalidData("DOC picture count exceeds the 32-bit range".to_string())
        })?;
        let shape_id = self.allocate_shape_id()?;
        self.pictures.push(WriterPicture {
            picture,
            shape_id,
            floating,
        });
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: anchor.to_string(),
                formatting: CharacterFormatting {
                    special: Some(true),
                    ..CharacterFormatting::default()
                },
                picture_index: Some(picture_index),
                shape_index: None,
            }],
            formatting: ParagraphFormatting::default(),
        });
        Ok(())
    }

    /// Shared tail of `insert_floating_shape`/`insert_floating_text_box`.
    fn insert_shape_run(
        &mut self,
        shape: super::shapes::DocDrawingShape,
        position: super::images::FloatingPosition,
        text: Option<String>,
    ) -> Result<(), DocWriteError> {
        let shape_index = u32::try_from(self.shapes.len()).map_err(|_| {
            DocWriteError::InvalidData("DOC shape count exceeds the 32-bit range".to_string())
        })?;
        let shape_id = self.allocate_shape_id()?;
        self.shapes.push(WriterShape {
            shape,
            shape_id,
            position,
            text,
        });
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: "\u{0008}".to_string(),
                formatting: CharacterFormatting {
                    special: Some(true),
                    ..CharacterFormatting::default()
                },
                picture_index: None,
                shape_index: Some(shape_index),
            }],
            formatting: ParagraphFormatting::default(),
        });
        Ok(())
    }

    /// Insert a floating primitive drawing shape anchored to its own paragraph.
    ///
    /// The anchor is a single 0x0008 character with sprmCFSpec applied, and
    /// the shape is emitted into the document's drawing group (fcDggInfo
    /// OfficeArtContent) with its position recorded in the Main Document's
    /// PlcfSpa — the same mechanism as floating pictures ([MS-DOC] 1.3).
    ///
    /// Shape text (text boxes) is not supported; see
    /// [`super::shapes::DocDrawingShape`].
    pub fn insert_floating_shape(
        &mut self,
        shape: super::shapes::DocDrawingShape,
        position: super::images::FloatingPosition,
    ) -> Result<(), DocWriteError> {
        self.insert_shape_run(shape, position, None)
    }

    /// Insert a floating text box anchored to its own paragraph.
    ///
    /// Anchoring and positioning work like [`Self::insert_floating_shape`],
    /// but the shape is emitted as an msosptTextBox with an
    /// OfficeArtClientTextbox record whose TXID links it to an entry in the
    /// textbox story ([MS-DOC] PlcftxbxTxt). The story text is appended to
    /// the WordDocument stream after the endnote story and counted in
    /// ccpTxbx. The text is plain: `\n` (or `\r` / `"\r\n"`) separates
    /// paragraphs; no character or paragraph formatting is applied.
    pub fn insert_floating_text_box(
        &mut self,
        shape: super::shapes::DocDrawingShape,
        position: super::images::FloatingPosition,
        text: impl Into<String>,
    ) -> Result<(), DocWriteError> {
        self.insert_shape_run(shape, position, Some(text.into()))
    }

    /// Insert a text box anchored in the given header.
    ///
    /// The anchor is a single 0x0008 paragraph appended to that header's
    /// paragraphs (created when absent); position and wrapping work like
    /// [`Self::insert_floating_text_box`], but the shape position is recorded
    /// in the Header Document's PlcfSpaHdr, the text goes to the header
    /// textbox story (counted in ccpHdrTxbx, linked through PlcfHdrtxbxTxt),
    /// and the shape joins the Header Document drawing of the fcDggInfo
    /// OfficeArtContent. Header floating items use their own shape-id cluster
    /// starting at 2049, so they never collide with main-story shapes.
    ///
    /// Set or replace header paragraphs BEFORE calling this method: the
    /// anchor lives in paragraphs this method appends, and replacing the
    /// header's paragraph list afterwards drops the anchor.
    pub fn insert_header_text_box(
        &mut self,
        kind: DocHeaderKind,
        shape: super::shapes::DocDrawingShape,
        position: super::images::FloatingPosition,
        text: impl Into<String>,
    ) -> Result<(), DocWriteError> {
        let shape_id = self.allocate_header_shape_id()?;
        let item_index = u32::try_from(self.header_shapes.len()).map_err(|_| {
            DocWriteError::InvalidData(
                "DOC header text box count exceeds the 32-bit range".to_string(),
            )
        })?;
        self.header_shapes.push(WriterShape {
            shape,
            shape_id,
            position,
            text: Some(text.into()),
        });
        self.append_header_anchor(kind, FloatingAnchorKind::Shape(item_index))
    }

    /// Insert a floating picture anchored in the given header (the classic
    /// letterhead logo / watermark pattern).
    ///
    /// Anchoring works like [`Self::insert_header_text_box`]: the picture is
    /// written as a PICF block with an embedded BLIP in the Data stream
    /// (bytes stored verbatim), referenced by sprmCPicLocation on the 0x0008
    /// anchor character, positioned through the PlcfSpaHdr, and rendered as a
    /// picture-frame shape in the Header Document drawing.
    pub fn insert_header_picture(
        &mut self,
        kind: DocHeaderKind,
        picture: super::images::DocPicture,
        position: super::images::FloatingPosition,
    ) -> Result<(), DocWriteError> {
        let shape_id = self.allocate_header_shape_id()?;
        let item_index = u32::try_from(self.header_pictures.len()).map_err(|_| {
            DocWriteError::InvalidData(
                "DOC header picture count exceeds the 32-bit range".to_string(),
            )
        })?;
        self.header_pictures.push(WriterPicture {
            picture,
            shape_id,
            floating: Some(position),
        });
        self.append_header_anchor(kind, FloatingAnchorKind::Picture(item_index))
    }

    /// Allocate the next header-drawing shape id from the header cluster.
    fn allocate_header_shape_id(&mut self) -> Result<u32, DocWriteError> {
        let count = self.header_shapes.len() + self.header_pictures.len();
        let index = u32::try_from(count).map_err(|_| {
            DocWriteError::InvalidData(
                "DOC header floating item count exceeds the 32-bit range".to_string(),
            )
        })?;
        Ok(super::images::HEADER_FIRST_SHAPE_ID + index)
    }

    /// Append a 0x0008 anchor paragraph to the given header and record it.
    fn append_header_anchor(
        &mut self,
        kind: DocHeaderKind,
        anchor_kind: FloatingAnchorKind,
    ) -> Result<(), DocWriteError> {
        let paragraphs = match kind {
            DocHeaderKind::Odd => &mut self.header_odd,
            DocHeaderKind::Even => &mut self.header_even,
            DocHeaderKind::FirstPage => &mut self.header_first,
        };
        let paragraphs = paragraphs.get_or_insert_with(Vec::new);
        let paragraph_index = paragraphs.len();
        paragraphs.push(HeaderFooterParagraph::from_runs(
            vec![(
                "\u{0008}".to_string(),
                CharacterFormatting {
                    special: Some(true),
                    ..CharacterFormatting::default()
                },
            )],
            ParagraphFormatting::default(),
        ));
        self.header_anchors.push(HeaderAnchor {
            slot: kind.slot(),
            paragraph_index,
            kind: anchor_kind,
        });
        Ok(())
    }

    /// Allocate the next shape id from the sequence shared by pictures and
    /// drawing shapes (group shape ids start one below the first picture id).
    fn allocate_shape_id(&mut self) -> Result<u32, DocWriteError> {
        let shape_id = self.next_shape_id;
        self.next_shape_id = self
            .next_shape_id
            .checked_add(1)
            .ok_or_else(|| DocWriteError::InvalidData("DOC shape ids exhausted".to_string()))?;
        Ok(shape_id)
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
        self.header_odd = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the even-page header text (HeaderStories index 6)
    pub fn set_even_header(&mut self, text: &str) {
        self.header_even = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the first-page header text (HeaderStories index 10)
    pub fn set_first_header(&mut self, text: &str) {
        self.header_first = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the odd-page footer text (HeaderStories index 9)
    pub fn set_odd_footer(&mut self, text: &str) {
        self.footer_odd = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the even-page footer text (HeaderStories index 8)
    pub fn set_even_footer(&mut self, text: &str) {
        self.footer_even = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the first-page footer text (HeaderStories index 11)
    pub fn set_first_footer(&mut self, text: &str) {
        self.footer_first = Some(vec![HeaderFooterParagraph::plain(text)]);
    }

    /// Set formatted odd-page header paragraphs (HeaderStories index 7).
    pub fn set_odd_header_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), DocWriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.header_odd = Some(paragraphs);
        Ok(())
    }

    /// Set formatted even-page header paragraphs (HeaderStories index 6).
    pub fn set_even_header_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), DocWriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.header_even = Some(paragraphs);
        Ok(())
    }

    /// Set formatted first-page header paragraphs (HeaderStories index 10).
    pub fn set_first_header_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), DocWriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.header_first = Some(paragraphs);
        Ok(())
    }

    /// Set formatted odd-page footer paragraphs (HeaderStories index 9).
    pub fn set_odd_footer_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), DocWriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.footer_odd = Some(paragraphs);
        Ok(())
    }

    /// Set formatted even-page footer paragraphs (HeaderStories index 8).
    pub fn set_even_footer_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), DocWriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.footer_even = Some(paragraphs);
        Ok(())
    }

    /// Set formatted first-page footer paragraphs (HeaderStories index 11).
    pub fn set_first_footer_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), DocWriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.footer_first = Some(paragraphs);
        Ok(())
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

    /// Add an inert smart-tag bookmark and property bag.
    pub fn add_smart_tag(&mut self, entry: DocSmartTagEntry) {
        self.smart_tags.push(entry);
    }

    /// Add one contiguous smart-tag recognizer-state range.
    ///
    /// Ranges are serialized in insertion order and must form a contiguous CP
    /// sequence when the document is saved.
    pub fn add_smart_tag_recognizer_range(&mut self, range: SmartTagRecognizerRange) {
        self.smart_tag_recognizer_ranges.push(range);
    }

    /// Add a list structure definition.
    pub fn add_list(&mut self, list: ListStructure) {
        self.numbering.add_list(list);
    }

    /// Add a list format override.
    pub fn add_list_override(&mut self, lfo: ListFormatOverride) {
        self.numbering.add_override(lfo);
    }

    /// Set names parallel to the document's list definitions.
    pub fn set_list_names(&mut self, table: ListNamesTable) {
        self.numbering.set_list_names(table);
    }

    /// Set template codes parallel to the document's list definitions.
    pub fn set_list_templates(&mut self, table: ListTemplateTable) {
        self.numbering.set_list_templates(table);
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
        for style in &self.styles {
            if let Some(revision) = &style.revision {
                index_author(&revision.author)?;
            }
        }
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            let mut formatting = Some(&paragraph.formatting);
            while let Some(current) = formatting {
                if let Some(revision) = &current.formatting_revision {
                    index_author(&revision.author)?;
                }
                if let Some(revision) = &current.numbering_revision {
                    index_author(&revision.author)?;
                }
                formatting = current.preserved_properties_for_revision.as_deref();
            }
            for run in &paragraph.runs {
                let mut formatting = Some(&run.formatting);
                while let Some(current) = formatting {
                    if let Some(revision) = &current.insertion_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.deletion_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.formatting_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.display_field_revision {
                        index_author(&revision.author)?;
                    }
                    formatting = current.preserved_properties_for_revision.as_deref();
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

    fn validate_style_reference(
        &self,
        index: u16,
        expected_kind: crate::doc::StyleKind,
        context: &str,
    ) -> Result<(), DocWriteError> {
        let actual_kind = match index {
            0 => Some(crate::doc::StyleKind::Paragraph),
            10 => Some(crate::doc::StyleKind::Character),
            15..=0x0FFC => self
                .styles
                .get(usize::from(index - 15))
                .map(|style| style.kind),
            _ => None,
        };
        let Some(actual_kind) = actual_kind else {
            return Err(DocWriteError::InvalidData(format!(
                "{context} references undefined DOC style index {index}"
            )));
        };
        if actual_kind != expected_kind {
            return Err(DocWriteError::InvalidData(format!(
                "{context} references {actual_kind:?} DOC style {index}, expected {expected_kind:?}"
            )));
        }
        Ok(())
    }

    fn validate_character_style_references(
        &self,
        formatting: &CharacterFormatting,
        context: &str,
    ) -> Result<(), DocWriteError> {
        if let Some(index) = formatting.style_index {
            self.validate_style_reference(index, crate::doc::StyleKind::Character, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_character_style_references(previous, context)?;
        }
        Ok(())
    }

    fn validate_paragraph_style_references(
        &self,
        formatting: &ParagraphFormatting,
        context: &str,
    ) -> Result<(), DocWriteError> {
        if let Some(index) = formatting.style_index {
            self.validate_style_reference(index, crate::doc::StyleKind::Paragraph, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_paragraph_style_references(previous, context)?;
        }
        Ok(())
    }

    fn validate_table_style_references(
        &self,
        formatting: &super::tap::TableRow,
        context: &str,
    ) -> Result<(), DocWriteError> {
        if let Some(index) = formatting.table_style_index {
            self.validate_style_reference(index, crate::doc::StyleKind::Table, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_table_style_references(previous, context)?;
        }
        Ok(())
    }

    fn validate_style_references(&self) -> Result<(), DocWriteError> {
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            self.validate_paragraph_style_references(
                &paragraph.formatting,
                "DOC paragraph formatting",
            )?;
            for run in &paragraph.runs {
                self.validate_character_style_references(
                    &run.formatting,
                    "DOC character formatting",
                )?;
            }
        }
        for table in &self.tables {
            for row in &table.rows {
                self.validate_table_style_references(&row.formatting, "DOC table row formatting")?;
            }
        }
        Ok(())
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

    fn append_smart_tag_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        smart_tags: &SmartTagTableData,
    ) {
        if let Some(data) = &smart_tags.infos {
            let offset = table_stream.len() as u32;
            fib.set_sttbf_bkmk_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.starts {
            let offset = table_stream.len() as u32;
            fib.set_plcf_bkf_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.ends {
            let offset = table_stream.len() as u32;
            fib.set_plcf_bkl_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.factoid_data {
            let offset = table_stream.len() as u32;
            fib.set_factoid_data(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.recognizer_ranges {
            let offset = table_stream.len() as u32;
            fib.set_plcf_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
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
        header_pic_offsets: &[u32],
    ) -> Result<Option<HeaderStoryData>, DocWriteError> {
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
        let story_text_start = text_stream.len();

        // Build index->paragraph mapping for 12 slots per MS-DOC PlcfHdd / Apache POI:
        //   Slots 0-5:  footnote/endnote separator/continuation stories
        //   Slot 6:     even page header (section 0)
        //   Slot 7:     odd page header (section 0) — "default" when no facing pages
        //   Slot 8:     even page footer (section 0)
        //   Slot 9:     odd page footer (section 0) — "default" when no facing pages
        //   Slot 10:    first page header (section 0)
        //   Slot 11:    first page footer (section 0)
        // PlcfHdd has 14 CPs (12 slot starts + story end + ignored final CP).
        // Verified against LibreOffice DOC writer output.
        let mut idx_paragraphs: [Option<&[HeaderFooterParagraph]>; 12] = [None; 12];
        if let Some(ref paragraphs) = self.header_even {
            idx_paragraphs[6] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.header_odd {
            idx_paragraphs[7] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.header_first {
            idx_paragraphs[10] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.footer_even {
            idx_paragraphs[8] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.footer_odd {
            idx_paragraphs[9] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.footer_first {
            idx_paragraphs[11] = Some(paragraphs);
        }

        // Local CP within header story (counts only header subdocument)
        // Empty slots consume no CPs. Non-empty header/footer stories contain a content paragraph
        // mark and a separate guard paragraph mark.
        let mut header_cp: u32 = 0;
        let mut cp_starts: [u32; 12] = [0; 12];
        let mut field_char_cps = Vec::new();
        let mut shape_anchor_cps: Vec<(u32, FloatingAnchorKind)> = Vec::new();

        for i in 0..12 {
            cp_starts[i] = header_cp;
            if let Some(paragraphs) = idx_paragraphs[i] {
                let mut field_state = HeaderFieldState::default();
                let fc_story_start = checked_text_fc(text_fc_start, text_stream.len())?;
                let mut story_chars = 0u32;

                for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
                    // Paragraphs appended by insert_header_text_box /
                    // insert_header_picture hold 0x0008 anchors; record their
                    // story CPs and the anchored item kind.
                    let anchor_kind = self
                        .header_anchors
                        .iter()
                        .find(|anchor| {
                            anchor.slot == i && anchor.paragraph_index == paragraph_index
                        })
                        .map(|anchor| anchor.kind);
                    if let Some(kind) = anchor_kind {
                        shape_anchor_cps.push((header_cp + story_chars, kind));
                    }
                    let fc_para_start = checked_text_fc(text_fc_start, text_stream.len())?;
                    let mut paragraph_chars = 0u32;
                    let mut last_chpx = None;

                    for (text, formatting) in &paragraph.runs {
                        let run_chars = utf16_code_unit_len(text)?;
                        let mut marker_cp = header_cp
                            .checked_add(story_chars)
                            .and_then(|value| value.checked_add(paragraph_chars))
                            .ok_or_else(|| {
                                DocWriteError::InvalidData(
                                    "DOC header/footer field CP range overflows".to_string(),
                                )
                            })?;
                        for character in text.chars() {
                            if field_state.observe(character, formatting)? {
                                field_char_cps.push((marker_cp, character as u16));
                            }
                            marker_cp = marker_cp
                                .checked_add(character.len_utf16() as u32)
                                .ok_or_else(|| {
                                    DocWriteError::InvalidData(
                                        "DOC header/footer field CP range overflows".to_string(),
                                    )
                                })?;
                        }
                        if run_chars == 0 {
                            continue;
                        }
                        let run_fc_start = checked_text_fc(text_fc_start, text_stream.len())?;
                        for unit in text.encode_utf16() {
                            text_stream.extend_from_slice(&unit.to_le_bytes());
                        }
                        let run_fc_end = checked_text_fc(text_fc_start, text_stream.len())?;
                        // Header picture anchors also carry sprmCPicLocation
                        // pointing at the picture's Data-stream block.
                        let mut grpprl = build_chpx_grpprl(formatting, font_builder);
                        if let Some(FloatingAnchorKind::Picture(pic_index)) = anchor_kind {
                            let pic_offset =
                                header_pic_offsets.get(pic_index as usize).ok_or_else(|| {
                                    DocWriteError::InvalidData(format!(
                                        "DOC header picture index {pic_index} is out of range"
                                    ))
                                })?;
                            grpprl.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
                            grpprl.extend_from_slice(&pic_offset.to_le_bytes());
                        }
                        chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                        last_chpx = Some(chpx_entries.len() - 1);
                        paragraph_chars =
                            paragraph_chars.checked_add(run_chars).ok_or_else(|| {
                                DocWriteError::InvalidData(
                                    "DOC header/footer paragraph CP range overflows".to_string(),
                                )
                            })?;
                    }

                    text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                    let fc_para_end = checked_text_fc(text_fc_start, text_stream.len())?;
                    if let Some(index) = last_chpx {
                        chpx_entries[index].1 = fc_para_end;
                    } else {
                        chpx_entries.push((fc_para_start, fc_para_end, Vec::new()));
                    }
                    papx_entries.push((
                        fc_para_start,
                        fc_para_end,
                        build_papx_grpprl(&paragraph.formatting),
                    ));
                    story_chars = story_chars
                        .checked_add(paragraph_chars)
                        .and_then(|value| value.checked_add(1))
                        .ok_or_else(|| {
                            DocWriteError::InvalidData(
                                "DOC header/footer story CP range overflows".to_string(),
                            )
                        })?;
                }

                // Guard paragraph mark required between stories.
                let fc_guard_start = checked_text_fc(text_fc_start, text_stream.len())?;
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                let fc_guard_end = checked_text_fc(text_fc_start, text_stream.len())?;
                chpx_entries.push((fc_guard_start, fc_guard_end, Vec::new()));
                papx_entries.push((
                    fc_guard_start,
                    fc_guard_end,
                    build_papx_grpprl(&ParagraphFormatting::default()),
                ));
                story_chars = story_chars.checked_add(1).ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC header/footer story CP range overflows".to_string(),
                    )
                })?;

                let cp_story_end = current_cp_total.checked_add(story_chars).ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC header/footer total CP range overflows".to_string(),
                    )
                })?;
                pieces.push(Piece::new(
                    *current_cp_total,
                    cp_story_end,
                    fc_story_start,
                    true,
                ));
                *current_cp_total = cp_story_end;
                header_cp = header_cp.checked_add(story_chars).ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC header/footer subdocument CP range overflows".to_string(),
                    )
                })?;
                field_state.finish()?;
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

        let fields = if field_char_cps.is_empty() {
            Vec::new()
        } else {
            super::fields::build_plcffld(
                &field_char_cps,
                header_cp,
                &text_stream[story_text_start..],
            )?
        };
        Ok(Some(HeaderStoryData {
            plcfhdd,
            fields,
            char_count: header_cp,
            shape_anchor_cps,
        }))
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
                    ..super::tap::TableRow::default()
                },
            };
            for _ in 0..cols {
                row.cells.push(TableCell {
                    paragraphs: vec![WritableParagraph {
                        runs: vec![TextRun {
                            text: String::new(),
                            formatting: CharacterFormatting::default(),
                            picture_index: None,
                            shape_index: None,
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

    /// Set validated column geometry for the writer's single section.
    pub fn set_section_columns(
        &mut self,
        columns: crate::doc::section::columns::Layout,
    ) -> Result<(), DocWriteError> {
        columns
            .validate()
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
        self.section_columns = Some(columns);
        Ok(())
    }

    /// Return explicit section column geometry, or `None` for the file-format default.
    pub fn section_columns(&self) -> Option<&crate::doc::section::columns::Layout> {
        self.section_columns.as_ref()
    }

    /// Remove the explicit column override and restore the single-column default.
    pub fn clear_section_columns(&mut self) {
        self.section_columns = None;
    }

    /// Select left-to-right or right-to-left section column population order.
    pub fn set_section_right_to_left(&mut self, value: bool) {
        self.section_right_to_left = value;
    }

    pub fn section_right_to_left(&self) -> bool {
        self.section_right_to_left
    }

    /// Set the section-wide text-flow mode.
    pub fn set_section_text_flow(&mut self, value: crate::doc::SectionTextFlow) {
        self.section_text_flow = value;
    }

    pub fn section_text_flow(&self) -> crate::doc::SectionTextFlow {
        self.section_text_flow
    }

    /// Set validated page borders for the writer's single section.
    pub fn set_section_page_borders(
        &mut self,
        borders: crate::doc::section::borders::Borders,
    ) -> Result<(), DocWriteError> {
        borders
            .validate()
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
        self.section_page_borders = Some(borders);
        Ok(())
    }

    /// Return explicit page borders, or `None` for the file-format default.
    pub fn section_page_borders(&self) -> Option<&crate::doc::section::borders::Borders> {
        self.section_page_borders.as_ref()
    }

    /// Remove all explicit page-border edges and placement controls.
    pub fn clear_section_page_borders(&mut self) {
        self.section_page_borders = None;
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
        self.build_ole_writer()?.save(path)?;
        Ok(())
    }

    /// Build and validate the three core DOC streams.
    fn build_output_streams(&mut self) -> Result<DocOutputStreams, DocWriteError> {
        self.build_output_streams_with_data_prefix(Vec::new())
    }

    /// Build the DOC streams while retaining an existing shared Data prefix.
    fn build_output_streams_with_data_prefix(
        &mut self,
        data_prefix: Vec<u8>,
    ) -> Result<DocOutputStreams, DocWriteError> {
        if self.attached_glossary.is_some() && self.glossary_metadata.is_some() {
            return Err(DocWriteError::InvalidData(
                "a DOC template cannot be both glossary-only and contain an attached glossary"
                    .to_string(),
            ));
        }
        self.validate_style_references()?;
        let table_header_len = self.encryption_table_header_len()?;

        // Based on Apache POI's HWPFDocument.write() implementation

        let mut word_document_stream = Vec::new();
        let mut table_stream = vec![0u8; table_header_len];

        // Reserve space for FIB (Word 2007+ format = 1248 bytes, includes cswNew)
        let fib_placeholder = vec![0u8; 1248];
        word_document_stream.extend_from_slice(&fib_placeholder);

        // fcMin will be set to padded start of text (after 512 alignment below)

        // Build text stream and piece table
        let mut text_stream = Vec::new();
        let mut data_stream = data_prefix;
        let mut floating_anchors: Vec<(u32, FloatingAnchorKind)> = Vec::new();
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
                // Pictures: append the OfficeArtWordDrawing block to the
                // Data stream and point sprmCPicLocation at it. Floating
                // pictures and shapes also record their anchor CP for the
                // PlcfSpa.
                let grpprl = if let Some(picture_index) = run.picture_index {
                    let entry = self.pictures.get(picture_index as usize).ok_or_else(|| {
                        DocWriteError::InvalidData(format!(
                            "DOC picture index {picture_index} is out of range"
                        ))
                    })?;
                    let pic_offset = u32::try_from(data_stream.len()).map_err(|_| {
                        DocWriteError::InvalidData(
                            "DOC Data stream exceeds 32-bit FC space".to_string(),
                        )
                    })?;
                    super::images::write_picture_block(
                        &entry.picture,
                        entry.shape_id,
                        &mut data_stream,
                    )?;
                    if entry.floating.is_some() {
                        floating_anchors.push((
                            current_cp + para_chars,
                            FloatingAnchorKind::Picture(picture_index),
                        ));
                    }
                    let mut grpprl = grpprl;
                    grpprl.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
                    grpprl.extend_from_slice(&pic_offset.to_le_bytes());
                    grpprl
                } else {
                    grpprl
                };
                if let Some(shape_index) = run.shape_index {
                    floating_anchors.push((
                        current_cp + para_chars,
                        FloatingAnchorKind::Shape(shape_index),
                    ));
                }

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

        // Header pictures: append their OfficeArtWordDrawing blocks to the
        // Data stream so the header story can point sprmCPicLocation at them.
        let mut header_pic_offsets: Vec<u32> = Vec::with_capacity(self.header_pictures.len());
        for entry in &self.header_pictures {
            let pic_offset = u32::try_from(data_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC Data stream exceeds 32-bit FC space".to_string())
            })?;
            super::images::write_picture_block(&entry.picture, entry.shape_id, &mut data_stream)?;
            header_pic_offsets.push(pic_offset);
        }

        let header_plcfhdd = self.build_header_story(
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
            &header_pic_offsets,
        )?;

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
        // Build textbox story (appends textbox text after the endnote story).
        // Entry order follows the anchor CPs so the FTXBXS indices match the
        // ClientTextbox TXIDs emitted into the drawing group below.
        floating_anchors.sort_by_key(|&(anchor_cp, _)| anchor_cp);
        let textbox_shapes: Vec<&WriterShape> = floating_anchors
            .iter()
            .filter_map(|&(_, kind)| match kind {
                FloatingAnchorKind::Shape(index) => {
                    let entry = &self.shapes[index as usize];
                    entry.text.as_ref().map(|_| entry)
                },
                FloatingAnchorKind::Picture(_) => None,
            })
            .collect();
        let mut txbx_start_cps: Vec<u32> = Vec::new();
        let mut ccp_txbx = 0u32;
        if !textbox_shapes.is_empty() {
            let txbx_story_start_cp = current_cp;
            let fc_story_start = text_fc_start + text_stream.len() as u32;
            for entry in &textbox_shapes {
                let text = entry.text.as_deref().expect("filtered on text presence");
                txbx_start_cps.push(current_cp - txbx_story_start_cp);
                // '\n' (and '\r' / "\r\n") separate plain-text paragraphs.
                for paragraph in text.replace("\r\n", "\n").replace('\r', "\n").split('\n') {
                    let para_len = utf16_code_unit_len(paragraph)?;
                    for unit in paragraph.encode_utf16() {
                        text_stream.extend_from_slice(&unit.to_le_bytes());
                    }
                    text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                    current_cp += para_len + 1;
                }
                // Trailing CR of this text box's text, as Word writes.
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                current_cp += 1;
            }
            // Story-final CR, included in ccpTxbx.
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            current_cp += 1;
            ccp_txbx = current_cp - txbx_story_start_cp;
            let fc_story_end = text_fc_start + text_stream.len() as u32;
            chpx_entries.push((fc_story_start, fc_story_end, Vec::new()));
            papx_entries.push((
                fc_story_start,
                fc_story_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(
                txbx_story_start_cp,
                current_cp,
                fc_story_start,
                true,
            ));
        }

        // Build header textbox story (after the main textbox story). Entry
        // order follows the header-story anchors so the FTXBXS indices match
        // the ClientTextbox TXIDs emitted into the header drawing below.
        let header_textbox_ids: Vec<u32> = header_plcfhdd
            .as_ref()
            .map(|header| {
                header
                    .shape_anchor_cps
                    .iter()
                    .filter_map(|&(_, kind)| match kind {
                        FloatingAnchorKind::Shape(index) => {
                            Some(self.header_shapes[index as usize].shape_id)
                        },
                        FloatingAnchorKind::Picture(_) => None,
                    })
                    .collect()
            })
            .unwrap_or_default();
        let header_texts: Vec<&str> = header_plcfhdd
            .as_ref()
            .map(|header| {
                header
                    .shape_anchor_cps
                    .iter()
                    .filter_map(|&(_, kind)| match kind {
                        FloatingAnchorKind::Shape(index) => {
                            self.header_shapes[index as usize].text.as_deref()
                        },
                        FloatingAnchorKind::Picture(_) => None,
                    })
                    .collect()
            })
            .unwrap_or_default();
        let mut hdr_txbx_start_cps: Vec<u32> = Vec::new();
        let mut ccp_hdr_txbx = 0u32;
        if !header_texts.is_empty() {
            let hdr_story_start_cp = current_cp;
            let fc_story_start = text_fc_start + text_stream.len() as u32;
            let (start_cps, ccp) =
                write_textbox_story_text(&header_texts, &mut text_stream, &mut current_cp)?;
            hdr_txbx_start_cps = start_cps;
            ccp_hdr_txbx = ccp;
            let fc_story_end = text_fc_start + text_stream.len() as u32;
            chpx_entries.push((fc_story_start, fc_story_end, Vec::new()));
            papx_entries.push((
                fc_story_start,
                fc_story_end,
                build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(
                hdr_story_start_cp,
                current_cp,
                fc_story_start,
                true,
            ));
        }

        let bookmark_tables = Self::build_bookmark_tables(&self.bookmarks, current_cp)?;
        let smart_tag_tables = super::smart_tags::build_tables(
            &self.smart_tags,
            &self.smart_tag_recognizer_ranges,
            current_cp,
        )?;

        // Mandatory trailing paragraph mark when ANY subdocument exists (same as save()).
        let has_subdocs = footnote_plcfs.is_some()
            || header_plcfhdd.is_some()
            || comment_story.is_some()
            || endnote_plcfs.is_some()
            || ccp_txbx > 0;
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
        let proofing_maximum_cp = current_cp
            .checked_add(if has_subdocs { 1 } else { 2 })
            .ok_or_else(|| {
                DocWriteError::InvalidData("document-parts proofing CP ceiling overflows".into())
            })?;

        let mut fib = FibBuilder::new();
        fib.set_main_text(0, text_length);
        if let Some((_, _, ftn_cp)) = &footnote_plcfs {
            fib.set_ccp_ftn(*ftn_cp);
        }
        if let Some(header) = &header_plcfhdd {
            fib.set_ccp_hdd(header.char_count);
        }
        if let Some(comment) = &comment_story {
            fib.set_ccp_atn(comment.char_count);
        }
        if let Some((_, _, edn_cp)) = &endnote_plcfs {
            fib.set_ccp_edn(*edn_cp);
        }
        if ccp_txbx > 0 {
            fib.set_ccp_txbx(ccp_txbx);
        }
        if ccp_hdr_txbx > 0 {
            fib.set_ccp_hdr_txbx(ccp_hdr_txbx);
        }

        let mut table_offset = table_stream.len() as u32;

        let stylesheet_data = crate::doc::writer::stylesheet::generate_stylesheet(
            &self.styles,
            revision_data.as_ref().map(|data| &data.indexes),
        )
        .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
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
        let dop_data = crate::doc::writer::dop::generate_dop(
            facing_pages,
            doc_grpf_ihdt,
            !smart_tag_tables.is_empty(),
        );
        fib.set_dop(table_offset, dop_data.len() as u32);
        table_stream.extend_from_slice(&dop_data);
        table_offset = table_stream.len() as u32;
        table_offset = super::auxiliary_strings::append_auxiliary_string_tables(
            &mut fib,
            &mut table_stream,
            &self.associated_strings,
            self.saved_by_table.as_ref(),
            table_offset,
        )?;
        table_offset = super::glossary::append_glossary_tables(
            &mut fib,
            &mut table_stream,
            self.glossary_metadata.as_ref(),
            table_offset,
            text_length,
            &text_stream,
        )?;

        // Write PlcfHdd if present
        if let Some(header) = &header_plcfhdd {
            fib.set_plcfhdd(table_offset, header.plcfhdd.len() as u32);
            table_stream.extend_from_slice(&header.plcfhdd);
            table_offset = table_stream.len() as u32;
            if !header.fields.is_empty() {
                fib.set_plcffld_hdr(table_offset, header.fields.len() as u32);
                table_stream.extend_from_slice(&header.fields);
                table_offset = table_stream.len() as u32;
            }
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
        if !smart_tag_tables.is_empty() {
            Self::append_smart_tag_tables(&mut fib, &mut table_stream, &smart_tag_tables);
            table_offset = table_stream.len() as u32;
        }
        if let Some(revisions) = &revision_data {
            Self::append_revision_author_table(&mut fib, &mut table_stream, revisions);
            table_offset = table_stream.len() as u32;
        }
        table_offset = super::proofing::append_proofing_tables(
            &mut fib,
            &mut table_stream,
            &self.proofing_tables,
            table_offset,
            proofing_maximum_cp,
        )?;

        // Write PlcfFldMom if there are field characters
        if !field_char_cps.is_empty() {
            let main_text_bytes = usize::try_from(text_length)
                .ok()
                .and_then(|value| value.checked_mul(2))
                .and_then(|length| text_stream.get(..length))
                .ok_or_else(|| {
                    DocWriteError::InvalidData(
                        "DOC main field story exceeds the text stream".to_string(),
                    )
                })?;
            let plcffld =
                super::fields::build_plcffld(&field_char_cps, text_length, main_text_bytes)?;
            if !plcffld.is_empty() {
                fib.set_plcffld_mom(table_offset, plcffld.len() as u32);
                table_stream.extend_from_slice(&plcffld);
                table_offset = table_stream.len() as u32;
            }
        }

        // Write numbering tables if present
        if !self.numbering.is_empty() {
            let (plflst_header, lvl_data) = self.numbering.build_plflst()?;
            fib.set_plflst(table_offset, plflst_header.len() as u32);
            table_stream.extend_from_slice(&plflst_header);
            table_stream.extend_from_slice(&lvl_data);
            table_offset = table_stream.len() as u32;

            let plflfo = self.numbering.build_plflfo();
            fib.set_plflfo(table_offset, plflfo.len() as u32);
            table_stream.extend_from_slice(&plflfo);
            table_offset = table_stream.len() as u32;

            if let Some(list_names) = self.numbering.build_sttb_list_names()? {
                fib.set_sttb_list_names(table_offset, list_names.len() as u32);
                table_stream.extend_from_slice(&list_names);
                table_offset = table_stream.len() as u32;
            }
            if let Some(list_templates) = self.numbering.build_sttb_rgtplc()? {
                fib.set_sttb_rgtplc(table_offset, list_templates.len() as u32);
                table_stream.extend_from_slice(&list_templates);
                table_offset = table_stream.len() as u32;
            }
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
        let sepx_data = crate::doc::writer::section::generate_sepx_with_properties(
            first_page,
            grpf_ihdt,
            section_revision,
            self.section_columns.as_ref(),
            self.section_right_to_left,
            self.section_text_flow,
            self.section_page_borders.as_ref(),
        )
        .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
        word_document_stream.extend_from_slice(&sepx_data);

        // Write section table to table stream
        let total_cp = current_cp;
        let section_table =
            crate::doc::writer::section::generate_section_table(total_cp, sepx_offset);
        table_offset = table_stream.len() as u32;
        fib.set_plcfsed(table_offset, section_table.len() as u32);
        table_stream.extend_from_slice(&section_table);

        // Floating pictures and shapes: shape position tables (PlcfSpaMom /
        // PlcfSpaHdr), the textbox story PLCs, and the drawing group
        // (fcDggInfo OfficeArtContent) that anchors the shapes to the
        // document's drawing layer.
        let header_anchor_cps: &[(u32, FloatingAnchorKind)] = header_plcfhdd
            .as_ref()
            .map(|header| header.shape_anchor_cps.as_slice())
            .unwrap_or(&[]);
        if !floating_anchors.is_empty() || !header_anchor_cps.is_empty() {
            table_offset = table_stream.len() as u32;
            let floating_shapes: Vec<super::images::FloatingShapeInfo<'_>> = floating_anchors
                .iter()
                .map(|&(anchor_cp, kind)| match kind {
                    FloatingAnchorKind::Picture(picture_index) => {
                        let entry = &self.pictures[picture_index as usize];
                        super::images::FloatingShapeInfo {
                            anchor_cp,
                            shape_id: entry.shape_id,
                            content: super::images::FloatingShapeContent::Picture(&entry.picture),
                            width_twips: entry.picture.width_twips(),
                            height_twips: entry.picture.height_twips(),
                            position: entry
                                .floating
                                .as_ref()
                                .expect("floating anchors are only recorded for floating pictures"),
                            text: None,
                        }
                    },
                    FloatingAnchorKind::Shape(shape_index) => {
                        let entry = &self.shapes[shape_index as usize];
                        super::images::FloatingShapeInfo {
                            anchor_cp,
                            shape_id: entry.shape_id,
                            content: super::images::FloatingShapeContent::Primitive(&entry.shape),
                            width_twips: entry.shape.width_twips(),
                            height_twips: entry.shape.height_twips(),
                            position: &entry.position,
                            text: entry.text.as_deref(),
                        }
                    },
                })
                .collect();
            let header_floating_shapes: Vec<super::images::FloatingShapeInfo<'_>> =
                header_anchor_cps
                    .iter()
                    .map(|&(anchor_cp, kind)| match kind {
                        FloatingAnchorKind::Shape(shape_index) => {
                            let entry = &self.header_shapes[shape_index as usize];
                            super::images::FloatingShapeInfo {
                                anchor_cp,
                                shape_id: entry.shape_id,
                                content: super::images::FloatingShapeContent::Primitive(
                                    &entry.shape,
                                ),
                                width_twips: entry.shape.width_twips(),
                                height_twips: entry.shape.height_twips(),
                                position: &entry.position,
                                text: entry.text.as_deref(),
                            }
                        },
                        FloatingAnchorKind::Picture(picture_index) => {
                            let entry = &self.header_pictures[picture_index as usize];
                            super::images::FloatingShapeInfo {
                                anchor_cp,
                                shape_id: entry.shape_id,
                                content: super::images::FloatingShapeContent::Picture(
                                    &entry.picture,
                                ),
                                width_twips: entry.picture.width_twips(),
                                height_twips: entry.picture.height_twips(),
                                position: entry
                                    .floating
                                    .as_ref()
                                    .expect("header pictures always have a floating position"),
                                text: None,
                            }
                        },
                    })
                    .collect();
            if !txbx_start_cps.is_empty() {
                let txbx_shape_ids: Vec<u32> =
                    textbox_shapes.iter().map(|entry| entry.shape_id).collect();
                let plcf_txbx =
                    super::shapes::build_plcf_txbx_txt(&txbx_shape_ids, &txbx_start_cps, ccp_txbx);
                fib.set_plcftxbx_txt(table_offset, plcf_txbx.len() as u32);
                table_stream.extend_from_slice(&plcf_txbx);
                table_offset = table_stream.len() as u32;
            }
            if !hdr_txbx_start_cps.is_empty() {
                let plcf_hdr_txbx = super::shapes::build_plcf_txbx_txt(
                    &header_textbox_ids,
                    &hdr_txbx_start_cps,
                    ccp_hdr_txbx,
                );
                fib.set_plcf_hdr_txbx_txt(table_offset, plcf_hdr_txbx.len() as u32);
                table_stream.extend_from_slice(&plcf_hdr_txbx);
                table_offset = table_stream.len() as u32;
            }
            if !floating_shapes.is_empty() {
                let plcf_spa = super::images::build_plcf_spa(&floating_shapes, text_length);
                fib.set_plc_spa_mom(table_offset, plcf_spa.len() as u32);
                table_stream.extend_from_slice(&plcf_spa);
                table_offset = table_stream.len() as u32;
            }
            if !header_floating_shapes.is_empty() {
                let header_char_count = header_plcfhdd
                    .as_ref()
                    .map(|header| header.char_count)
                    .unwrap_or(0);
                let plcf_spa_hdr =
                    super::images::build_plcf_spa(&header_floating_shapes, header_char_count);
                fib.set_plc_spa_hdr(table_offset, plcf_spa_hdr.len() as u32);
                table_stream.extend_from_slice(&plcf_spa_hdr);
                table_offset = table_stream.len() as u32;
            }

            let total_shapes = (self.pictures.len() + self.shapes.len()) as u32;
            let dgg_info = super::images::build_dgg_info(
                &floating_shapes,
                &header_floating_shapes,
                total_shapes,
            )?;
            fib.set_dgg_info(table_offset, dgg_info.len() as u32);
            table_stream.extend_from_slice(&dgg_info);
        }

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

        // POI writes a zero-filled Data stream when the document has no pictures.
        let mut data_stream = if data_stream.is_empty() {
            vec![0u8; 4096]
        } else {
            pad_to_4096(&mut data_stream);
            data_stream
        };
        if let Some(glossary) = self.attached_glossary.as_mut() {
            glossary.validate_as_attached_glossary()?;
            let data_prefix = std::mem::take(&mut data_stream);
            let mut glossary_streams =
                glossary.build_output_streams_with_data_prefix(data_prefix)?;
            super::attached_glossary::merge_attached_glossary(
                &mut word_document_stream,
                &mut table_stream,
                &mut glossary_streams.word_document,
                &mut glossary_streams.table,
            )?;
            data_stream = glossary_streams.data;
        }
        self.encrypt_output_streams(
            &mut word_document_stream,
            &mut table_stream,
            &mut data_stream,
        )?;

        Ok(DocOutputStreams {
            word_document: word_document_stream,
            table: table_stream,
            data: data_stream,
        })
    }

    /// Build the complete compound document after validating every staged structure.
    fn build_ole_writer(&mut self) -> Result<OleWriter, DocWriteError> {
        let streams = self.build_output_streams()?;
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(
            &mut ole_writer,
            &streams.word_document,
            &streams.table,
            &streams.data,
        )?;
        Ok(ole_writer)
    }

    /// Write the document to a seekable output.
    pub fn write_to<W: std::io::Write + std::io::Seek>(
        &mut self,
        writer: &mut W,
    ) -> Result<(), DocWriteError> {
        self.build_ole_writer()?.write_to(writer)?;
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

    if let Some(style_index) = fmt.style_index {
        push_word(&mut grp, SPRM_C_ISTD, style_index);
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
    if let Some(position) = fmt.position {
        grp.extend_from_slice(&SPRM_C_HPS_POS.to_le_bytes());
        grp.extend_from_slice(&position.half_points().to_le_bytes());
    }
    if let Some(hyphenation) = fmt.hyphenation {
        grp.extend_from_slice(&SPRM_C_HRESI.to_le_bytes());
        grp.extend_from_slice(&hyphenation.bytes());
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
    if let Some(effect) = fmt.text_effect {
        push_byte(&mut grp, SPRM_C_SFXT_TEXT, effect.into());
    }

    grp
}

fn build_revision_chpx_grpprl(
    fmt: &CharacterFormatting,
    font_builder: &mut FontTableBuilder,
    revisions: Option<&RevisionWriterData>,
) -> Result<Vec<u8>, DocWriteError> {
    if fmt
        .preserved_properties_for_revision
        .as_ref()
        .is_some_and(|previous| previous.preserved_properties_for_revision.is_some())
    {
        return Err(DocWriteError::InvalidData(
            "DOC character property revisions cannot contain nested preserved states".to_string(),
        ));
    }
    if fmt.insertion_revision.is_some() && fmt.deletion_revision.is_some() {
        return Err(DocWriteError::InvalidData(
            "a DOC character run cannot be both an insertion and a deletion".to_string(),
        ));
    }
    let mut grp = if let Some(previous) = &fmt.preserved_properties_for_revision {
        let mut grp = build_revision_chpx_grpprl(previous, font_builder, revisions)?;
        grp.extend_from_slice(&SPRM_C_WALL.to_le_bytes());
        grp.push(1);
        grp.extend_from_slice(&build_chpx_grpprl(fmt, font_builder));
        grp
    } else {
        build_chpx_grpprl(fmt, font_builder)
    };
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

    if let Some(style_index) = fmt.style_index {
        push_u16(&mut grp, SPRM_P_ISTD, style_index);
    }
    // Alignment. Emit a compatible physical value before the authoritative logical value.
    if let Some(jc) = fmt.alignment {
        let physical = match jc {
            0..=3 => Some(jc),
            4 | 5 => Some(4),
            7 | 8 => Some(5),
            9 => Some(3),
            _ => None,
        };
        if let Some(physical) = physical {
            push_byte(&mut grp, SPRM_P_JC, physical);
        }
        push_byte(&mut grp, SPRM_P_JC_LOGICAL, jc);
    } else if let Some(physical) = fmt.physical_justification {
        let code = match physical {
            PhysicalJustification::Left => 0,
            PhysicalJustification::Center => 1,
            PhysicalJustification::Right => 2,
            PhysicalJustification::LowCompression => 3,
            PhysicalJustification::MediumCompression => 4,
            PhysicalJustification::HighCompression => 5,
        };
        push_byte(&mut grp, SPRM_P_JC, code);
    }
    if let Some(style) = fmt.legacy_border_style {
        push_byte(&mut grp, SPRM_P_BRCL, style as u8);
    }
    if let Some(position) = fmt.legacy_border_position {
        push_byte(&mut grp, SPRM_P_BRCP, position as u8);
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
    if let Some(dxc_left) = fmt.left_indent_chars {
        push_i16(&mut grp, SPRM_P_DXC_LEFT, dxc_left);
    }
    if let Some(dxc_right) = fmt.right_indent_chars {
        push_i16(&mut grp, SPRM_P_DXC_RIGHT, dxc_right);
    }
    if let Some(dxc_first) = fmt.first_line_indent_chars {
        push_i16(&mut grp, SPRM_P_DXC_LEFT1, dxc_first);
    }
    // Spacing (twips)
    if let Some(dya_before) = fmt.space_before {
        push_u16(&mut grp, SPRM_P_DYA_BEFORE, dya_before);
    }
    if let Some(dya_after) = fmt.space_after {
        push_u16(&mut grp, SPRM_P_DYA_AFTER, dya_after);
    }
    if let Some(disabled) = fmt.no_line_numbering {
        push_bool(&mut grp, SPRM_P_F_NO_LINE_NUMB, disabled);
    }
    if let Some(dyl_before) = fmt.space_before_lines {
        push_i16(&mut grp, SPRM_P_DYL_BEFORE, dyl_before);
    }
    if let Some(dyl_after) = fmt.space_after_lines {
        push_i16(&mut grp, SPRM_P_DYL_AFTER, dyl_after);
    }

    // Auto spacing flags
    if let Some(auto) = fmt.space_before_auto {
        push_bool(&mut grp, SPRM_P_F_DYA_BEFORE_AUTO, auto);
    }
    if let Some(auto) = fmt.space_after_auto {
        push_bool(&mut grp, SPRM_P_F_DYA_AFTER_AUTO, auto);
    }
    if let Some(open) = fmt.open_table_cell_mark {
        push_bool(&mut grp, SPRM_P_F_OPEN_TCH, open);
    }

    // Side-by-side and pagination controls
    if let Some(side_by_side) = fmt.side_by_side {
        push_bool(&mut grp, SPRM_P_F_SIDE_BY_SIDE, side_by_side);
    }
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
    for (opcode, value) in [
        (SPRM_P_F_LOCKED, fmt.frame_anchor_locked),
        (SPRM_P_F_KINSOKU, fmt.kinsoku),
        (SPRM_P_F_WORD_WRAP, fmt.word_wrap),
        (SPRM_P_F_OVERFLOW_PUNCT, fmt.overflow_punctuation),
        (SPRM_P_F_TOP_LINE_PUNCT, fmt.top_line_punctuation),
        (SPRM_P_F_AUTO_SPACE_DE, fmt.auto_space_east_asian_latin),
        (SPRM_P_F_AUTO_SPACE_DN, fmt.auto_space_east_asian_numbers),
    ] {
        if let Some(value) = value {
            push_bool(&mut grp, opcode, value);
        }
    }
    if let Some(alignment) = fmt.font_alignment {
        push_u16(&mut grp, SPRM_P_W_ALIGN_FONT, alignment as u16);
    }
    if let Some(flow) = fmt.frame_text_flow {
        let value = u16::from(flow.vertical)
            | (u16::from(flow.backwards) << 1)
            | (u16::from(flow.rotate_font) << 2);
        push_u16(&mut grp, SPRM_P_FRAME_TEXT_FLOW, value);
    }
    if let Some(position) = fmt.frame_horizontal_position {
        let value = match position {
            FrameHorizontalPosition::Left => 0,
            FrameHorizontalPosition::Center => -4,
            FrameHorizontalPosition::Right => -8,
            FrameHorizontalPosition::Inside => -12,
            FrameHorizontalPosition::Outside => -16,
            FrameHorizontalPosition::Offset(offset) => offset + 1,
        };
        push_i16(&mut grp, SPRM_P_DXA_ABS, value);
    }
    if let Some(position) = fmt.frame_vertical_position {
        let value = match position {
            FrameVerticalPosition::Inline => 0,
            FrameVerticalPosition::Top => -4,
            FrameVerticalPosition::Center => -8,
            FrameVerticalPosition::Bottom => -12,
            FrameVerticalPosition::Inside => -16,
            FrameVerticalPosition::Outside => -20,
            FrameVerticalPosition::Offset(offset) => offset + 1,
        };
        push_i16(&mut grp, SPRM_P_DYA_ABS, value);
    }
    if let Some(width) = fmt.frame_width {
        push_u16(&mut grp, SPRM_P_DXA_WIDTH, width);
    }
    if let Some(anchor) = fmt.frame_anchor {
        let vertical = match anchor.vertical {
            FrameVerticalAnchor::Margin => 0,
            FrameVerticalAnchor::Page => 1,
            FrameVerticalAnchor::Paragraph => 2,
            FrameVerticalAnchor::None => 3,
        };
        let horizontal = match anchor.horizontal {
            FrameHorizontalAnchor::Column => 0,
            FrameHorizontalAnchor::Margin => 1,
            FrameHorizontalAnchor::Page => 2,
            FrameHorizontalAnchor::None => 3,
        };
        push_byte(&mut grp, SPRM_P_PC, (vertical << 4) | (horizontal << 6));
    }
    if let Some(in_table) = fmt.in_table {
        push_bool(&mut grp, SPRM_P_F_IN_TABLE, in_table);
    }
    if let Some(terminating) = fmt.table_terminating_paragraph {
        push_bool(&mut grp, SPRM_P_F_TTP, terminating);
    }
    if let Some(wrap) = fmt.frame_text_wrap {
        push_byte(&mut grp, SPRM_P_WR, wrap as u8);
    }
    if let Some(height) = fmt.frame_height {
        push_u16(
            &mut grp,
            SPRM_P_W_HEIGHT_ABS,
            height.height_twips | (u16::from(height.minimum) << 15),
        );
    }
    if let Some(distance) = fmt.frame_horizontal_text_distance {
        push_i16(&mut grp, SPRM_P_DXA_FROM_TEXT, distance);
    }
    if let Some(distance) = fmt.frame_vertical_text_distance {
        push_i16(&mut grp, SPRM_P_DYA_FROM_TEXT, distance);
    }
    if let Some(drop_cap) = fmt.drop_cap {
        let kind = match drop_cap.kind {
            crate::doc::parts::pap::DropCapType::Regular => 1u16,
            crate::doc::parts::pap::DropCapType::Margin => 2,
        };
        push_u16(
            &mut grp,
            SPRM_P_DCS,
            kind | (u16::from(drop_cap.lines) << 3),
        );
    }
    if let Some(disabled) = fmt.no_auto_hyphenation {
        push_bool(&mut grp, SPRM_P_F_NO_AUTO_HYPH, disabled);
    }

    // BiDi paragraph
    if let Some(bidi) = fmt.bidi {
        push_bool(&mut grp, SPRM_P_F_BI_DI, bidi);
    }
    if let Some(use_grid) = fmt.use_page_setup_settings {
        push_bool(&mut grp, SPRM_P_F_USE_PGSU_SETTINGS, use_grid);
    }
    if let Some(adjust) = fmt.adjust_right_indent {
        push_bool(&mut grp, SPRM_P_F_ADJUST_RIGHT, adjust);
    }

    // Outline level
    if let Some(lvl) = fmt.outline_level {
        grp.extend_from_slice(&SPRM_P_OUT_LVL.to_le_bytes());
        grp.push(lvl);
    }

    // Floating-object overlap and text-box wrapping behavior
    if let Some(no_overlap) = fmt.no_allow_overlap {
        push_bool(&mut grp, SPRM_P_F_NO_ALLOW_OVERLAP, no_overlap);
    }
    if let Some(cs) = fmt.contextual_spacing {
        push_bool(&mut grp, SPRM_P_F_CONTEXTUAL_SPACING, cs);
    }
    if let Some(mi) = fmt.mirror_indents {
        push_bool(&mut grp, SPRM_P_F_MIRROR_INDENTS, mi);
    }
    if let Some(tight_wrap) = fmt.text_box_tight_wrap {
        push_byte(&mut grp, SPRM_P_TTWO, tight_wrap as u8);
    }
    for (opcode, border) in [
        (SPRM_P_BRC_TOP, fmt.borders.top),
        (SPRM_P_BRC_LEFT, fmt.borders.left),
        (SPRM_P_BRC_BOTTOM, fmt.borders.bottom),
        (SPRM_P_BRC_RIGHT, fmt.borders.right),
        (SPRM_P_BRC_BETWEEN, fmt.borders.between),
        (SPRM_P_BRC_BAR, fmt.borders.bar),
    ] {
        if let Some(border) = border {
            append_paragraph_border(&mut grp, opcode, border);
        }
    }
    if let Some(shading) = fmt.shading {
        grp.extend_from_slice(&SPRM_P_SHD.to_le_bytes());
        grp.push(10);
        for color in [shading.foreground_color, shading.background_color] {
            match color {
                Some((red, green, blue)) => grp.extend_from_slice(&[red, green, blue, 0]),
                None => grp.extend_from_slice(&[0, 0, 0, 0xFF]),
            }
        }
        grp.extend_from_slice(&(shading.pattern as u16).to_le_bytes());
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
    if let Some(autonumbering) = &fmt.legacy_autonumbering {
        append_legacy_autonumbering(&mut grp, autonumbering);
    }
    if let Some(revision_save_id) = fmt.revision_save_id {
        grp.extend_from_slice(&SPRM_P_RSID.to_le_bytes());
        grp.extend_from_slice(&revision_save_id.to_le_bytes());
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
    append_tab_changes(&mut grp, &fmt.tab_stops_to_delete, &fmt.tab_stops_to_add);

    grp
}

fn append_tab_changes(output: &mut Vec<u8>, deletes: &[i32], additions: &[TabStop]) {
    let mut deletes = deletes.to_vec();
    deletes.sort_unstable();
    for chunk in deletes.chunks(64) {
        output.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
        output.push((2 + chunk.len() * 2) as u8);
        output.push(chunk.len() as u8);
        for position in chunk {
            output.extend_from_slice(&(*position as i16).to_le_bytes());
        }
        output.push(0);
    }

    let mut additions = additions.to_vec();
    additions.sort_unstable_by_key(|tab| tab.position);
    for chunk in additions.chunks(64) {
        output.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
        output.push((2 + chunk.len() * 3) as u8);
        output.push(0);
        output.push(chunk.len() as u8);
        for tab in chunk {
            output.extend_from_slice(&(tab.position as i16).to_le_bytes());
        }
        for tab in chunk {
            let alignment = match tab.alignment {
                TabAlignment::Left => 0,
                TabAlignment::Center => 1,
                TabAlignment::Right => 2,
                TabAlignment::Decimal => 3,
                TabAlignment::Bar => 4,
                TabAlignment::List => 6,
            };
            let leader = if tab.alignment == TabAlignment::Bar {
                0
            } else {
                match tab.leader {
                    TabLeader::None => 0,
                    TabLeader::Dots => 1,
                    TabLeader::Hyphens => 2,
                    TabLeader::Underline => 3,
                    TabLeader::Heavy => 4,
                    TabLeader::MiddleDot => 5,
                    TabLeader::DefaultLeader => 7,
                }
            };
            output.push(alignment | (leader << 3));
        }
    }
}

fn append_legacy_autonumbering(output: &mut Vec<u8>, value: &LegacyAutoNumbering) {
    let mut operand = [0u8; 84];
    operand[0] = value.number_format as u8;
    let prefix = value.prefix.encode_utf16().collect::<Vec<_>>();
    let suffix = value.suffix.encode_utf16().collect::<Vec<_>>();
    operand[1] = prefix.len() as u8;
    operand[2] = suffix.len() as u8;
    operand[3] = match value.alignment {
        AutoNumberAlignment::Left => 0,
        AutoNumberAlignment::Center => 1,
        AutoNumberAlignment::Right => 2,
        AutoNumberAlignment::Justified => 3,
    } | (u8::from(value.include_previous_levels) << 2)
        | (u8::from(value.hanging_indent) << 3)
        | (u8::from(value.set_bold) << 4)
        | (u8::from(value.set_italic) << 5)
        | (u8::from(value.set_small_caps) << 6)
        | (u8::from(value.set_caps) << 7);
    operand[4] = u8::from(value.set_strike)
        | (u8::from(value.set_underline) << 1)
        | (u8::from(value.prefix_space) << 2)
        | (u8::from(value.bold) << 3)
        | (u8::from(value.italic) << 4)
        | (u8::from(value.small_caps) << 5)
        | (u8::from(value.caps) << 6)
        | (u8::from(value.strike) << 7);
    operand[5] = value.underline | (value.color_index << 3);
    operand[6..8].copy_from_slice(&value.font_index.to_le_bytes());
    operand[8..10].copy_from_slice(&value.font_size_half_points.to_le_bytes());
    operand[10..12].copy_from_slice(&value.start_at.to_le_bytes());
    operand[12..14].copy_from_slice(&value.indent_twips.to_le_bytes());
    operand[14..16].copy_from_slice(&value.space_twips.to_le_bytes());
    operand[16] = u8::from(value.number_once_per_cell);
    operand[17] = u8::from(value.number_across_cells);
    operand[18] = u8::from(value.restart_each_section);
    for (index, unit) in prefix.into_iter().chain(suffix).enumerate() {
        let offset = 20 + index * 2;
        operand[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
    }
    output.extend_from_slice(&SPRM_P_ANLD.to_le_bytes());
    output.push(operand.len() as u8);
    output.extend_from_slice(&operand);
}

fn append_paragraph_border(output: &mut Vec<u8>, opcode: u16, border: ParagraphBorder) {
    output.extend_from_slice(&opcode.to_le_bytes());
    output.push(8);
    match border.color {
        Some((red, green, blue)) => output.extend_from_slice(&[red, green, blue, 0]),
        None => output.extend_from_slice(&[0, 0, 0, 0xFF]),
    }
    output.push(border.width);
    output.push(match border.style {
        ParagraphBorderStyle::None => 0,
        ParagraphBorderStyle::Single => 1,
        ParagraphBorderStyle::Double => 3,
        ParagraphBorderStyle::Thick => 5,
        ParagraphBorderStyle::Dotted => 6,
        ParagraphBorderStyle::Dashed => 7,
        ParagraphBorderStyle::DotDash => 8,
        ParagraphBorderStyle::DotDotDash => 9,
        ParagraphBorderStyle::Triple => 10,
        ParagraphBorderStyle::ThinThickSmallGap => 11,
        ParagraphBorderStyle::ThickThinSmallGap => 12,
        ParagraphBorderStyle::ThinThickThinSmallGap => 13,
        ParagraphBorderStyle::ThinThickMediumGap => 14,
        ParagraphBorderStyle::ThickThinMediumGap => 15,
        ParagraphBorderStyle::ThinThickThinMediumGap => 16,
        ParagraphBorderStyle::ThinThickLargeGap => 17,
        ParagraphBorderStyle::ThickThinLargeGap => 18,
        ParagraphBorderStyle::ThinThickThinLargeGap => 19,
        ParagraphBorderStyle::Wave => 20,
        ParagraphBorderStyle::DoubleWave => 21,
        ParagraphBorderStyle::DashSmallGap => 22,
        ParagraphBorderStyle::DashDotStroked => 23,
        ParagraphBorderStyle::ThreeDEmboss => 24,
        ParagraphBorderStyle::ThreeDEngrave => 25,
        ParagraphBorderStyle::Outset => 26,
        ParagraphBorderStyle::Inset => 27,
    });
    output.push(border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6));
    output.push(0);
}

fn build_revision_papx_grpprl(
    fmt: &ParagraphFormatting,
    revisions: Option<&RevisionWriterData>,
) -> Result<Vec<u8>, DocWriteError> {
    if fmt
        .preserved_properties_for_revision
        .as_ref()
        .is_some_and(|previous| previous.preserved_properties_for_revision.is_some())
    {
        return Err(DocWriteError::InvalidData(
            "DOC paragraph property revisions cannot contain nested preserved states".to_string(),
        ));
    }
    if let Some(alignment) = fmt.alignment
        && alignment > 9
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC paragraph alignment {alignment} is outside 0..=9"
        )));
    }
    if let Some(outline_level) = fmt.outline_level
        && outline_level > 9
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC paragraph outline level {outline_level} is outside 0..=9"
        )));
    }
    if let Some(level) = fmt.ilvl
        && level > 8
        && level != 0x0C
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC paragraph list level {level} is neither 0..=8 nor the skip value 12"
        )));
    }
    if let Some(ilfo) = fmt.ilfo
        && (0x07FF..=0xF800).contains(&ilfo)
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC paragraph list override {ilfo:#06x} is reserved"
        )));
    }
    if let Some(value) = &fmt.legacy_autonumbering {
        let prefix_units = value.prefix.encode_utf16().count();
        let suffix_units = value.suffix.encode_utf16().count();
        if prefix_units + suffix_units > 32 {
            return Err(DocWriteError::InvalidData(format!(
                "DOC legacy autonumber label uses {} UTF-16 units; maximum is 32",
                prefix_units + suffix_units
            )));
        }
        if value.underline > 7 {
            return Err(DocWriteError::InvalidData(format!(
                "DOC legacy autonumber underline {} exceeds 7",
                value.underline
            )));
        }
        if value.color_index > 16 {
            return Err(DocWriteError::InvalidData(format!(
                "DOC legacy autonumber color index {} exceeds 16",
                value.color_index
            )));
        }
        if !(-31_680..=31_680).contains(&value.indent_twips) {
            return Err(DocWriteError::InvalidData(format!(
                "DOC legacy autonumber indent {} is outside -31680..=31680",
                value.indent_twips
            )));
        }
        if value.space_twips > 31_680 {
            return Err(DocWriteError::InvalidData(format!(
                "DOC legacy autonumber spacing {} exceeds 31680",
                value.space_twips
            )));
        }
    }
    for (name, value) in [
        ("left_indent", fmt.left_indent),
        ("right_indent", fmt.right_indent),
        ("first_line_indent", fmt.first_line_indent),
    ] {
        if let Some(value) = value
            && !(-31_680..=31_680).contains(&value)
        {
            return Err(DocWriteError::InvalidData(format!(
                "DOC paragraph {name} value {value} is outside -31680..=31680"
            )));
        }
    }
    for (name, value) in [
        ("space_before", fmt.space_before),
        ("space_after", fmt.space_after),
    ] {
        if let Some(value) = value
            && value > 31_680
        {
            return Err(DocWriteError::InvalidData(format!(
                "DOC paragraph {name} value {value} exceeds 31680"
            )));
        }
    }
    if let Some(spacing) = fmt.line_spacing
        && !(-31_680..=31_680).contains(&spacing.dya_line)
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC paragraph line spacing {} is outside the LSPD range",
            spacing.dya_line
        )));
    }
    let added_tab_positions = fmt
        .tab_stops_to_add
        .iter()
        .map(|tab| tab.position)
        .collect::<Vec<_>>();
    for (kind, positions) in [
        ("deleted", fmt.tab_stops_to_delete.as_slice()),
        ("added", added_tab_positions.as_slice()),
    ] {
        let mut sorted = positions.to_vec();
        sorted.sort_unstable();
        if sorted
            .iter()
            .any(|position| !(-31_680..=31_680).contains(position))
        {
            return Err(DocWriteError::InvalidData(format!(
                "DOC {kind} tab position is outside -31680..=31680"
            )));
        }
        if sorted.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(DocWriteError::InvalidData(format!(
                "DOC {kind} tab positions contain a duplicate"
            )));
        }
    }
    if let Some(flow) = fmt.frame_text_flow
        && flow.backwards
        && !flow.vertical
    {
        return Err(DocWriteError::InvalidData(
            "DOC backwards frame text flow requires vertical flow".to_string(),
        ));
    }
    if let Some(height) = fmt.frame_height
        && (height.height_twips > 0x7FFF || (height.minimum && height.height_twips == 0))
    {
        return Err(DocWriteError::InvalidData(
            "DOC paragraph frame height is outside the WHeightAbs range".to_string(),
        ));
    }
    if let Some(drop_cap) = fmt.drop_cap
        && !(1..=10).contains(&drop_cap.lines)
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC drop-cap line count {} is outside 1..=10",
            drop_cap.lines
        )));
    }
    for (name, distance) in [
        ("horizontal", fmt.frame_horizontal_text_distance),
        ("vertical", fmt.frame_vertical_text_distance),
    ] {
        if let Some(distance) = distance
            && !(0..=31_680).contains(&distance)
        {
            return Err(DocWriteError::InvalidData(format!(
                "DOC {name} frame text distance {distance} is outside 0..=31680"
            )));
        }
    }
    for (name, offset) in [
        (
            "horizontal",
            match fmt.frame_horizontal_position {
                Some(FrameHorizontalPosition::Offset(value)) => Some(value),
                _ => None,
            },
        ),
        (
            "vertical",
            match fmt.frame_vertical_position {
                Some(FrameVerticalPosition::Offset(value)) => Some(value),
                _ => None,
            },
        ),
    ] {
        if let Some(offset) = offset
            && !(-31_679..=31_681).contains(&offset)
        {
            return Err(DocWriteError::InvalidData(format!(
                "DOC {name} frame offset {offset} is outside the plus-one range"
            )));
        }
        if let Some(offset) = offset {
            let stored = offset + 1;
            let is_special =
                matches!(stored, 0 | -4 | -8 | -12 | -16) || (name == "vertical" && stored == -20);
            if is_special {
                return Err(DocWriteError::InvalidData(format!(
                    "DOC {name} frame offset {offset} encodes a reserved alignment value"
                )));
            }
        }
    }
    if let Some(width) = fmt.frame_width
        && width > 31_680
    {
        return Err(DocWriteError::InvalidData(format!(
            "DOC paragraph frame width {width} exceeds 31680"
        )));
    }
    if fmt.table_terminating_paragraph == Some(true) && fmt.in_table != Some(true) {
        return Err(DocWriteError::InvalidData(
            "DOC table-terminating paragraph requires in_table=true".to_string(),
        ));
    }
    if fmt.frame_text_flow.is_some()
        && !matches!(fmt.frame_text_wrap, Some(wrap) if wrap != FrameTextWrap::Auto)
        && !matches!(fmt.frame_height, Some(height) if height.height_twips != 0)
        && fmt.frame_horizontal_position.is_none()
        && fmt.frame_vertical_position.is_none()
        && fmt.frame_width.is_none()
        && fmt.frame_anchor.is_none()
    {
        return Err(DocWriteError::InvalidData(
            "DOC frame text flow requires a non-default frame property".to_string(),
        ));
    }
    for (name, value) in [
        ("space_before_lines", fmt.space_before_lines),
        ("space_after_lines", fmt.space_after_lines),
    ] {
        if let Some(value) = value
            && !(-20..=31_680).contains(&value)
        {
            return Err(DocWriteError::InvalidData(format!(
                "DOC paragraph {name} value {value} is outside -20..=31680"
            )));
        }
    }
    for border in [
        fmt.borders.top,
        fmt.borders.left,
        fmt.borders.bottom,
        fmt.borders.right,
        fmt.borders.between,
        fmt.borders.bar,
    ]
    .into_iter()
    .flatten()
    {
        if border.spacing > 31 {
            return Err(DocWriteError::InvalidData(format!(
                "DOC paragraph border spacing {} exceeds 31 points",
                border.spacing
            )));
        }
    }

    let mut grp = if let Some(previous) = &fmt.preserved_properties_for_revision {
        let mut grp = build_revision_papx_grpprl(previous, revisions)?;
        grp.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
        grp.push(1);
        grp.extend_from_slice(&build_papx_grpprl(fmt));
        grp
    } else {
        build_papx_grpprl(fmt)
    };
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
    use crate::doc::parts::numbering::NumberFormat;
    use std::io::Cursor;

    #[test]
    fn test_create_writer() {
        let writer = DocWriter::new();
        assert_eq!(writer.paragraphs.len(), 0);
        assert_eq!(writer.tables.len(), 0);
    }

    #[test]
    fn writes_custom_styles_into_document_stylesheet() {
        let mut writer = DocWriter::new();
        let paragraph_style = writer
            .add_style(super::super::stylesheet::DocStyleDefinition::new(
                crate::doc::StyleKind::Paragraph,
                "Custom Body",
            ))
            .unwrap();
        let character_style = writer
            .add_style(super::super::stylesheet::DocStyleDefinition::new(
                crate::doc::StyleKind::Character,
                "Custom Emphasis",
            ))
            .unwrap();
        let table_style = writer
            .add_style(super::super::stylesheet::DocStyleDefinition::new(
                crate::doc::StyleKind::Table,
                "Custom Grid",
            ))
            .unwrap();
        assert_eq!(
            (paragraph_style, character_style, table_style),
            (15, 16, 17)
        );
        writer
            .add_paragraph_with_format(
                "Styled document",
                CharacterFormatting {
                    style_index: Some(character_style),
                    ..CharacterFormatting::default()
                },
                ParagraphFormatting {
                    style_index: Some(paragraph_style),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();
        let table = writer.add_table(1, 1).unwrap();
        writer
            .set_table_row_formatting(
                table,
                0,
                super::super::tap::TableRow {
                    cells: vec![super::super::tap::TableCell::default()],
                    table_style_index: Some(table_style),
                    ..super::super::tap::TableRow::default()
                },
            )
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        let stylesheet = document.stylesheet().unwrap();
        assert_eq!(stylesheet.styles().len(), 18);
        assert_eq!(stylesheet.get(paragraph_style).unwrap().name, "Custom Body");
        assert_eq!(
            stylesheet.get(character_style).unwrap().name,
            "Custom Emphasis"
        );
        assert_eq!(stylesheet.get(table_style).unwrap().name, "Custom Grid");
        assert_eq!(
            stylesheet.get(table_style).unwrap().kind,
            crate::doc::StyleKind::Table
        );
        let paragraphs = document.paragraphs().unwrap();
        assert_eq!(
            paragraphs[0].properties().style_index,
            Some(paragraph_style)
        );
        assert_eq!(
            paragraphs[0].runs().unwrap()[0].properties().style_index,
            Some(character_style)
        );
        assert_eq!(
            document.tables().unwrap()[0].rows().unwrap()[0]
                .properties()
                .unwrap()
                .table_style_index,
            Some(table_style)
        );
    }

    #[test]
    fn writes_revision_marked_style_and_author_table() {
        let timestamp = CommentDateTime {
            year: 2026,
            month: 7,
            day: 16,
            hour: 11,
            minute: 45,
            weekday: 4,
        };
        let previous_papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[0]].concat();
        let previous_chpx = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
        let mut writer = DocWriter::new();
        let style_index = writer
            .add_style(
                super::super::stylesheet::DocStyleDefinition::new(
                    crate::doc::StyleKind::Paragraph,
                    "Tracked Body",
                )
                .with_revision(
                    super::super::stylesheet::DocStyleRevision::paragraph(
                        "Style Editor",
                        previous_papx.clone(),
                        previous_chpx.clone(),
                    )
                    .with_timestamp(timestamp),
                ),
            )
            .unwrap();
        writer
            .add_formatted_paragraph(
                "Tracked style",
                ParagraphFormatting {
                    style_index: Some(style_index),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        assert_eq!(document.revision_authors(), ["Unknown", "Style Editor"]);
        let stylesheet = document.stylesheet().unwrap();
        let revision = stylesheet
            .get(style_index)
            .unwrap()
            .revision
            .as_ref()
            .unwrap();
        assert_eq!(revision.author_index, 1);
        assert_eq!(revision.author.as_deref(), Some("Style Editor"));
        assert_eq!(revision.timestamp, Some(timestamp));
        assert_eq!(
            revision.paragraph_properties.as_deref(),
            Some(previous_papx.as_slice())
        );
        assert_eq!(revision.character_properties, previous_chpx);
        assert_eq!(
            document.paragraphs().unwrap()[0].properties().style_index,
            Some(style_index)
        );
    }

    #[test]
    fn rejects_undefined_or_wrong_kind_style_references() {
        let error_for_paragraph_style = |style_index| {
            let mut writer = DocWriter::new();
            writer
                .add_formatted_paragraph(
                    "text",
                    ParagraphFormatting {
                        style_index: Some(style_index),
                        ..ParagraphFormatting::default()
                    },
                )
                .unwrap();
            writer
                .write_to(&mut Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string()
        };
        assert!(error_for_paragraph_style(14).contains("undefined DOC style index 14"));

        let mut writer = DocWriter::new();
        let character_style = writer
            .add_style(super::super::stylesheet::DocStyleDefinition::new(
                crate::doc::StyleKind::Character,
                "Wrong Kind",
            ))
            .unwrap();
        writer
            .add_formatted_paragraph(
                "text",
                ParagraphFormatting {
                    style_index: Some(character_style),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();
        let error = writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string();
        assert!(error.contains("Character DOC style 15, expected Paragraph"));
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
                                    color: Some((1, 2, 3)),
                                    border_type: crate::doc::parts::tap::BorderType::Single,
                                    spacing: 2,
                                    shadow: true,
                                    frame: false,
                                }),
                                diagonal_down: Some(crate::doc::parts::tap::BorderStyle {
                                    width: 4,
                                    color: Some((10, 20, 30)),
                                    border_type: crate::doc::parts::tap::BorderType::Outset,
                                    spacing: 1,
                                    shadow: false,
                                    frame: true,
                                }),
                                ..crate::doc::parts::tap::CellBorders::default()
                            },
                            border_type_overrides: crate::doc::parts::tap::CellBorderTypes::default(
                            ),
                            shading: Some(crate::doc::parts::tap::CellShading {
                                foreground_color: Some((1, 2, 3)),
                                background_color: Some((250, 240, 230)),
                                pattern: crate::doc::parts::tap::ShadingPattern::DarkCross,
                            }),
                            padding_top: Some(120),
                            padding_left: Some(240),
                            padding_bottom: Some(120),
                            padding_right: Some(240),
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
                    borders: crate::doc::writer::TableBorders {
                        vertical: Some(crate::doc::parts::tap::BorderStyle {
                            width: 6,
                            color: Some((40, 50, 60)),
                            border_type: crate::doc::parts::tap::BorderType::Double,
                            spacing: 0,
                            shadow: false,
                            frame: false,
                        }),
                        ..crate::doc::writer::TableBorders::default()
                    },
                    ..crate::doc::writer::TableRow::default()
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
                    allow_break: false,
                    ..crate::doc::writer::TableRow::default()
                },
            )
            .unwrap();
        let second_table = writer.add_table(1, 1).unwrap();
        writer
            .set_table_cell_text(second_table, 0, 0, "Separate")
            .unwrap();

        let assert_document = |document: crate::doc::Document| {
            let stylesheet = document.stylesheet().unwrap();
            assert_eq!(stylesheet.styles().len(), 15);
            assert_eq!(stylesheet.get(0).unwrap().name, "Normal");
            assert_eq!(stylesheet.get(10).unwrap().invariant_id, 65);
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
            assert_eq!(top_border.color, Some((1, 2, 3)));
            assert_eq!(
                top_border.border_type,
                crate::doc::parts::tap::BorderType::Single
            );
            assert_eq!(top_border.spacing, 2);
            assert!(top_border.shadow);
            assert!(!top_border.frame);
            let diagonal = first_cell_properties.borders.diagonal_down.unwrap();
            assert_eq!(diagonal.color, Some((10, 20, 30)));
            assert_eq!(
                diagonal.border_type,
                crate::doc::parts::tap::BorderType::Outset
            );
            assert!(diagonal.frame);
            assert_eq!(
                rows[0].properties().unwrap().border_vertical.unwrap().color,
                Some((40, 50, 60))
            );
            assert_eq!(
                first_cell_properties.shading,
                Some(crate::doc::parts::tap::CellShading {
                    foreground_color: Some((1, 2, 3)),
                    background_color: Some((250, 240, 230)),
                    pattern: crate::doc::parts::tap::ShadingPattern::DarkCross,
                })
            );
            assert_eq!(
                first_cell_properties.background_color,
                Some((250, 240, 230))
            );
            assert_eq!(first_cell_properties.padding_top, Some(120));
            assert_eq!(first_cell_properties.padding_left, Some(240));
            assert_eq!(first_cell_properties.padding_bottom, Some(120));
            assert_eq!(first_cell_properties.padding_right, Some(240));
            let first_cell = &rows[0].cells().unwrap()[0];
            assert_eq!(first_cell.shading(), first_cell_properties.shading);
            assert_eq!(first_cell.shading_inherits_from_style(), Some(false));
            assert_eq!(first_cell.background_color(), Some((250, 240, 230)));
            assert_eq!(first_cell.padding_top(), Some(120));
            assert_eq!(first_cell.padding_left(), Some(240));
            assert_eq!(first_cell.padding_bottom(), Some(120));
            assert_eq!(first_cell.padding_right(), Some(240));
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
    fn file_and_seekable_outputs_are_byte_identical() {
        let mut writer = DocWriter::new();
        writer.set_property("Title", "Canonical output");
        writer.add_paragraph("One output assembly path").unwrap();

        let mut memory = Cursor::new(Vec::new());
        writer.write_to(&mut memory).unwrap();

        let path = std::env::temp_dir().join(format!(
            "litchi-doc-output-equivalence-{}-{}.doc",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        writer.save(&path).unwrap();
        let file = std::fs::read(&path).unwrap();
        std::fs::remove_file(path).unwrap();

        assert_eq!(file, memory.into_inner());
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
        assert_eq!(
            writer.header_odd.as_ref().unwrap()[0].runs[0].0,
            "Odd Header"
        );
        assert_eq!(
            writer.header_even.as_ref().unwrap()[0].runs[0].0,
            "Even Header"
        );
        assert_eq!(
            writer.header_first.as_ref().unwrap()[0].runs[0].0,
            "First Header"
        );
        assert_eq!(
            writer.footer_odd.as_ref().unwrap()[0].runs[0].0,
            "Odd Footer"
        );
        assert_eq!(
            writer.footer_even.as_ref().unwrap()[0].runs[0].0,
            "Even Footer"
        );
        assert_eq!(
            writer.footer_first.as_ref().unwrap()[0].runs[0].0,
            "First Footer"
        );
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
    fn encodes_physical_and_logical_paragraph_justification_compatibly() {
        let physical = build_papx_grpprl(&ParagraphFormatting {
            physical_justification: Some(PhysicalJustification::HighCompression),
            ..ParagraphFormatting::default()
        });
        assert_eq!(
            physical,
            [SPRM_P_JC.to_le_bytes().as_slice(), &[5]].concat()
        );

        let logical = build_papx_grpprl(&ParagraphFormatting {
            alignment: Some(7),
            ..ParagraphFormatting::default()
        });
        assert_eq!(
            logical,
            [
                SPRM_P_JC.to_le_bytes().as_slice(),
                &[5],
                SPRM_P_JC_LOGICAL.to_le_bytes().as_slice(),
                &[7],
            ]
            .concat()
        );
    }

    #[test]
    fn encodes_paragraph_revision_save_id() {
        let encoded = build_papx_grpprl(&ParagraphFormatting {
            revision_save_id: Some(0x1122_3344),
            ..ParagraphFormatting::default()
        });
        assert_eq!(
            encoded,
            [
                SPRM_P_RSID.to_le_bytes().as_slice(),
                0x1122_3344u32.to_le_bytes().as_slice(),
            ]
            .concat()
        );
    }

    #[test]
    fn preserves_ordered_paragraph_property_revision_state() {
        let formatting = ParagraphFormatting {
            right_indent: Some(200),
            preserved_properties_for_revision: Some(Box::new(ParagraphFormatting {
                left_indent: Some(100),
                ..ParagraphFormatting::default()
            })),
            ..ParagraphFormatting::default()
        };
        let grpprl = build_revision_papx_grpprl(&formatting, None).unwrap();
        let properties = crate::doc::parts::pap::ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.indent_left, Some(100));
        assert_eq!(properties.indent_right, Some(200));
        let previous = properties.preserved_properties_for_revision.unwrap();
        assert_eq!(previous.indent_left, Some(100));
        assert_eq!(previous.indent_right, None);

        let mut writer = DocWriter::new();
        writer
            .add_formatted_paragraph("Tracked formatting", formatting)
            .unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();
        let properties = paragraphs[0].properties();
        assert_eq!(properties.indent_left, Some(100));
        assert_eq!(properties.indent_right, Some(200));
        let previous = properties
            .preserved_properties_for_revision
            .as_ref()
            .unwrap();
        assert_eq!(previous.indent_left, Some(100));
        assert_eq!(previous.indent_right, None);
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

        assert!(LineSpacing::multiple_240ths(0).is_err());
        assert!(LineSpacing::multiple_240ths(31_681).is_err());
        assert!(LineSpacing::at_least_twips(0).is_err());
        assert!(LineSpacing::at_least_twips(31_681).is_err());
        assert!(LineSpacing::exact_twips(0).is_err());
        assert!(LineSpacing::exact_twips(31_681).is_err());
    }

    #[test]
    fn test_paragraph_formatting_writer_reader_round_trip() {
        let legacy_autonumbering = LegacyAutoNumbering {
            number_format: NumberFormat::RussianUpper,
            alignment: AutoNumberAlignment::Justified,
            include_previous_levels: true,
            hanging_indent: true,
            set_bold: true,
            set_italic: true,
            set_small_caps: true,
            set_caps: true,
            set_strike: true,
            set_underline: true,
            prefix_space: true,
            bold: true,
            italic: true,
            small_caps: true,
            caps: true,
            strike: true,
            underline: 3,
            color_index: 6,
            font_index: 4,
            font_size_half_points: 24,
            start_at: 3,
            indent_twips: -360,
            space_twips: 180,
            number_once_per_cell: true,
            number_across_cells: false,
            restart_each_section: true,
            prefix: "§(".to_string(),
            suffix: ")".to_string(),
        };
        let mut writer = DocWriter::new();
        writer
            .add_formatted_paragraph(
                "Exactly spaced",
                ParagraphFormatting {
                    alignment: Some(1),
                    left_indent_chars: Some(250),
                    right_indent_chars: Some(-125),
                    first_line_indent_chars: Some(-50),
                    space_before: Some(120),
                    space_after: Some(240),
                    no_line_numbering: Some(true),
                    space_before_lines: Some(-20),
                    space_after_lines: Some(31_680),
                    space_before_auto: Some(true),
                    space_after_auto: Some(true),
                    open_table_cell_mark: Some(true),
                    frame_anchor_locked: Some(true),
                    kinsoku: Some(true),
                    word_wrap: Some(false),
                    overflow_punctuation: Some(true),
                    top_line_punctuation: Some(true),
                    auto_space_east_asian_latin: Some(true),
                    auto_space_east_asian_numbers: Some(false),
                    font_alignment: Some(FontAlignment::Bottom),
                    frame_text_flow: Some(FrameTextFlow {
                        vertical: true,
                        backwards: true,
                        rotate_font: false,
                    }),
                    frame_horizontal_position: Some(FrameHorizontalPosition::Right),
                    frame_vertical_position: Some(FrameVerticalPosition::Offset(300)),
                    frame_width: Some(1_440),
                    frame_anchor: Some(FrameAnchor {
                        vertical: FrameVerticalAnchor::Paragraph,
                        horizontal: FrameHorizontalAnchor::Margin,
                    }),
                    in_table: Some(false),
                    table_terminating_paragraph: Some(false),
                    frame_text_wrap: Some(FrameTextWrap::Through),
                    frame_height: Some(FrameHeight {
                        height_twips: 720,
                        minimum: true,
                    }),
                    frame_horizontal_text_distance: Some(480),
                    frame_vertical_text_distance: Some(240),
                    drop_cap: Some(DropCap {
                        kind: crate::doc::parts::pap::DropCapType::Margin,
                        lines: 3,
                    }),
                    no_auto_hyphenation: Some(true),
                    side_by_side: Some(true),
                    use_page_setup_settings: Some(true),
                    adjust_right_indent: Some(false),
                    no_allow_overlap: Some(true),
                    contextual_spacing: Some(true),
                    mirror_indents: Some(true),
                    text_box_tight_wrap: Some(TextBoxTightWrap::FirstAndLastLine),
                    borders: ParagraphBorders {
                        top: Some(ParagraphBorder {
                            style: ParagraphBorderStyle::Inset,
                            width: 12,
                            color: Some((0x11, 0x22, 0x33)),
                            spacing: 7,
                            shadow: true,
                            frame: true,
                        }),
                        left: Some(ParagraphBorder {
                            style: ParagraphBorderStyle::DoubleWave,
                            width: 8,
                            color: None,
                            spacing: 4,
                            shadow: false,
                            frame: false,
                        }),
                        ..ParagraphBorders::default()
                    },
                    legacy_border_style: Some(LegacyBorderStyle::Shadow),
                    legacy_border_position: Some(LegacyBorderPosition::LeftBar),
                    shading: Some(ParagraphShading {
                        foreground_color: Some((1, 2, 3)),
                        background_color: None,
                        pattern: crate::doc::parts::tap::ShadingPattern::DiagonalCross,
                    }),
                    line_spacing: Some(LineSpacing::exact_twips(360).unwrap()),
                    tab_stops_to_delete: vec![720],
                    tab_stops_to_add: vec![
                        TabStop {
                            position: 1_440,
                            alignment: TabAlignment::List,
                            leader: TabLeader::DefaultLeader,
                        },
                        TabStop {
                            position: 720,
                            alignment: TabAlignment::Decimal,
                            leader: TabLeader::Dots,
                        },
                    ],
                    ilvl: Some(8),
                    ilfo: Some(1),
                    legacy_autonumbering: Some(legacy_autonumbering.clone()),
                    revision_save_id: Some(0x1122_3344),
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
        writer
            .add_formatted_paragraph(
                "Thai distributed",
                ParagraphFormatting {
                    alignment: Some(9),
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

        assert_eq!(paragraphs.len(), 3);
        assert_eq!(paragraphs[0].text().unwrap(), "Exactly spaced");
        assert_eq!(
            paragraphs[0].properties().justification,
            crate::doc::parts::pap::Justification::Center
        );
        assert_eq!(paragraphs[0].properties().space_before, Some(120));
        assert_eq!(paragraphs[0].properties().space_after, Some(240));
        assert!(paragraphs[0].properties().no_line_numbering);
        assert_eq!(paragraphs[0].properties().list_level, Some(8));
        assert_eq!(paragraphs[0].properties().list_format_override, Some(1));
        assert_eq!(
            paragraphs[0].properties().legacy_autonumbering,
            Some(legacy_autonumbering)
        );
        assert_eq!(
            paragraphs[0].properties().revision_save_id,
            Some(0x1122_3344)
        );
        assert_eq!(
            paragraphs[0].properties().tab_stops,
            vec![
                TabStop {
                    position: 720,
                    alignment: TabAlignment::Decimal,
                    leader: TabLeader::Dots,
                },
                TabStop {
                    position: 1_440,
                    alignment: TabAlignment::List,
                    leader: TabLeader::DefaultLeader,
                },
            ]
        );
        assert_eq!(paragraphs[0].properties().indent_left_chars, Some(250));
        assert_eq!(paragraphs[0].properties().indent_right_chars, Some(-125));
        assert_eq!(
            paragraphs[0].properties().indent_first_line_chars,
            Some(-50)
        );
        assert_eq!(paragraphs[0].properties().space_before_lines, Some(-20));
        assert_eq!(paragraphs[0].properties().space_after_lines, Some(31_680));
        assert!(paragraphs[0].properties().space_before_auto);
        assert!(paragraphs[0].properties().space_after_auto);
        assert!(paragraphs[0].properties().open_table_cell_mark);
        assert!(paragraphs[0].properties().locked);
        assert!(paragraphs[0].properties().kinsoku);
        assert!(!paragraphs[0].properties().word_wrap);
        assert!(paragraphs[0].properties().overflow_punct);
        assert!(paragraphs[0].properties().top_line_punct);
        assert!(paragraphs[0].properties().auto_space_de);
        assert!(!paragraphs[0].properties().auto_space_dn);
        assert_eq!(
            paragraphs[0].properties().font_align,
            Some(FontAlignment::Bottom)
        );
        assert_eq!(
            paragraphs[0].properties().frame_text_flow,
            Some(FrameTextFlow {
                vertical: true,
                backwards: true,
                rotate_font: false,
            })
        );
        assert_eq!(
            paragraphs[0].properties().frame_horizontal_position,
            Some(FrameHorizontalPosition::Right)
        );
        assert_eq!(
            paragraphs[0].properties().frame_vertical_position,
            Some(FrameVerticalPosition::Offset(300))
        );
        assert_eq!(paragraphs[0].properties().frame_width, Some(1_440));
        assert_eq!(
            paragraphs[0].properties().frame_anchor,
            Some(FrameAnchor {
                vertical: FrameVerticalAnchor::Paragraph,
                horizontal: FrameHorizontalAnchor::Margin,
            })
        );
        assert!(!paragraphs[0].properties().in_table);
        assert!(!paragraphs[0].properties().is_table_row_end);
        assert_eq!(
            paragraphs[0].properties().text_wrap,
            Some(FrameTextWrap::Through)
        );
        assert_eq!(
            paragraphs[0].properties().frame_height,
            Some(FrameHeight {
                height_twips: 720,
                minimum: true,
            })
        );
        assert_eq!(paragraphs[0].properties().dxa_from_text, Some(480));
        assert_eq!(paragraphs[0].properties().dya_from_text, Some(240));
        assert_eq!(
            paragraphs[0].properties().drop_cap,
            Some(DropCap {
                kind: crate::doc::parts::pap::DropCapType::Margin,
                lines: 3,
            })
        );
        assert!(paragraphs[0].properties().no_auto_hyph);
        assert!(paragraphs[0].properties().side_by_side);
        assert_eq!(
            paragraphs[0].properties().use_page_setup_settings,
            Some(true)
        );
        assert_eq!(paragraphs[0].properties().adjust_right_indent, Some(false));
        assert!(paragraphs[0].properties().no_allow_overlap);
        assert!(paragraphs[0].properties().contextual_spacing);
        assert!(paragraphs[0].properties().mirror_indents);
        assert_eq!(
            paragraphs[0].properties().text_box_tight_wrap,
            Some(TextBoxTightWrap::FirstAndLastLine)
        );
        assert_eq!(
            paragraphs[0].properties().borders,
            ParagraphBorders {
                top: Some(ParagraphBorder {
                    style: ParagraphBorderStyle::Inset,
                    width: 12,
                    color: Some((0x11, 0x22, 0x33)),
                    spacing: 7,
                    shadow: true,
                    frame: true,
                }),
                left: Some(ParagraphBorder {
                    style: ParagraphBorderStyle::DoubleWave,
                    width: 8,
                    color: None,
                    spacing: 4,
                    shadow: false,
                    frame: false,
                }),
                ..ParagraphBorders::default()
            }
        );
        assert_eq!(
            paragraphs[0].properties().legacy_border_style,
            Some(LegacyBorderStyle::Shadow)
        );
        assert_eq!(
            paragraphs[0].properties().legacy_border_position,
            Some(LegacyBorderPosition::LeftBar)
        );
        assert_eq!(
            paragraphs[0].properties().shading,
            Some(ParagraphShading {
                foreground_color: Some((1, 2, 3)),
                background_color: None,
                pattern: crate::doc::parts::tap::ShadingPattern::DiagonalCross,
            })
        );
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
        assert_eq!(
            paragraphs[2].properties().justification,
            crate::doc::parts::pap::Justification::ThaiDistributed
        );
    }

    #[test]
    fn rejects_invalid_current_paragraph_layout_values() {
        for formatting in [
            ParagraphFormatting {
                alignment: Some(10),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                outline_level: Some(10),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                ilvl: Some(9),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                ilfo: Some(0x07FF),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                preserved_properties_for_revision: Some(Box::new(ParagraphFormatting {
                    left_indent: Some(31_681),
                    ..ParagraphFormatting::default()
                })),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                preserved_properties_for_revision: Some(Box::new(ParagraphFormatting {
                    preserved_properties_for_revision: Some(Box::new(
                        ParagraphFormatting::default(),
                    )),
                    ..ParagraphFormatting::default()
                })),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                legacy_autonumbering: Some(LegacyAutoNumbering {
                    prefix: "x".repeat(33),
                    ..LegacyAutoNumbering::default()
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                legacy_autonumbering: Some(LegacyAutoNumbering {
                    underline: 8,
                    ..LegacyAutoNumbering::default()
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                legacy_autonumbering: Some(LegacyAutoNumbering {
                    color_index: 17,
                    ..LegacyAutoNumbering::default()
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                legacy_autonumbering: Some(LegacyAutoNumbering {
                    indent_twips: i16::MIN,
                    ..LegacyAutoNumbering::default()
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                legacy_autonumbering: Some(LegacyAutoNumbering {
                    space_twips: 31_681,
                    ..LegacyAutoNumbering::default()
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                left_indent: Some(31_681),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                space_after: Some(31_681),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                line_spacing: Some(LineSpacing {
                    dya_line: i16::MIN,
                    is_multiple: false,
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                tab_stops_to_add: vec![TabStop {
                    position: 31_681,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                }],
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                tab_stops_to_delete: vec![720, 720],
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_text_flow: Some(FrameTextFlow {
                    vertical: false,
                    backwards: true,
                    rotate_font: false,
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_height: Some(FrameHeight {
                    height_twips: 32_768,
                    minimum: false,
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                drop_cap: Some(DropCap {
                    kind: crate::doc::parts::pap::DropCapType::Regular,
                    lines: 0,
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_horizontal_text_distance: Some(-1),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_horizontal_position: Some(FrameHorizontalPosition::Offset(i16::MAX)),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_horizontal_position: Some(FrameHorizontalPosition::Offset(-5)),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_width: Some(31_681),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                table_terminating_paragraph: Some(true),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                frame_text_flow: Some(FrameTextFlow {
                    vertical: true,
                    backwards: false,
                    rotate_font: false,
                }),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                space_before_lines: Some(-21),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                space_after_lines: Some(31_681),
                ..ParagraphFormatting::default()
            },
            ParagraphFormatting {
                borders: ParagraphBorders {
                    top: Some(ParagraphBorder {
                        style: ParagraphBorderStyle::Single,
                        width: 8,
                        color: None,
                        spacing: 32,
                        shadow: false,
                        frame: false,
                    }),
                    ..ParagraphBorders::default()
                },
                ..ParagraphFormatting::default()
            },
        ] {
            assert!(build_revision_papx_grpprl(&formatting, None).is_err());
        }
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
        let field_text = document.fields().unwrap();
        assert_eq!(field_text.len(), 1);
        assert_eq!(
            field_text[0].instruction.trim(),
            r#"HYPERLINK "https://example.test""#
        );
        assert_eq!(field_text[0].result.as_deref(), Some("link"));
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
                    modified_at: Some(CommentDateTime {
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
            Some(CommentDateTime {
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
                    modified_at: Some(CommentDateTime {
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
    fn smart_tags_round_trip_through_both_output_paths() {
        let mut writer = DocWriter::new();
        writer.add_paragraph("abcdefghijklmnopqrst").unwrap();
        writer.add_smart_tag(
            DocSmartTagEntry::new(0, 10, "urn:example:geo", "place")
                .with_origin(crate::doc::SmartTagOrigin::ExternalRecognizer)
                .with_native_export(true)
                .with_property("city", "東京"),
        );
        writer.add_smart_tag(
            DocSmartTagEntry::new(5, 15, "urn:example:geo", "place")
                .with_sub_entity(true)
                .with_property("city", "Paris"),
        );
        writer.add_smart_tag(DocSmartTagEntry::new(5, 5, "urn:example:point", "cursor"));
        writer.add_smart_tag_recognizer_range(SmartTagRecognizerRange {
            start: 0,
            end: 5,
            state: crate::doc::SmartTagRecognizerState::Dirty,
        });
        writer.add_smart_tag_recognizer_range(SmartTagRecognizerRange {
            start: 5,
            end: 20,
            state: crate::doc::SmartTagRecognizerState::Clean,
        });

        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        for index in [114usize, 115, 117, 118, 132] {
            assert!(document.fib().get_table_pointer(index).unwrap().1 > 0);
        }
        let smart_tags = document.smart_tags().unwrap().clone();
        assert_eq!(smart_tags.tags.len(), 3);
        assert_eq!(smart_tags.store.as_ref().unwrap().types.len(), 2);
        assert_eq!(
            smart_tags.tags[0].info.origin,
            crate::doc::SmartTagOrigin::ExternalRecognizer
        );
        assert!(smart_tags.tags[0].is_native);
        assert_eq!(
            smart_tags
                .store
                .as_ref()
                .unwrap()
                .resolve_property(smart_tags.tags[0].property_bag.properties[0]),
            Some(("city", "東京"))
        );
        assert_eq!(
            (smart_tags.tags[1].start_depth, smart_tags.tags[1].end_depth),
            (3, 0)
        );
        assert_eq!(smart_tags.recognizer_ranges.len(), 2);

        let path = std::env::temp_dir().join(format!(
            "litchi-doc-smart-tags-{}-{}.doc",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        writer.save(&path).unwrap();
        let mut package = crate::doc::Package::open(&path).unwrap();
        assert_eq!(package.document().unwrap().smart_tags(), Some(&smart_tags));
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
        let timestamp = CommentDateTime {
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
    fn preserves_ordered_character_property_revision_state() {
        let formatting = CharacterFormatting {
            italic: Some(true),
            preserved_properties_for_revision: Some(Box::new(CharacterFormatting {
                bold: Some(true),
                ..CharacterFormatting::default()
            })),
            ..CharacterFormatting::default()
        };
        let mut fonts = FontTableBuilder::new();
        let grpprl = build_revision_chpx_grpprl(&formatting, &mut fonts, None).unwrap();
        let properties = crate::doc::parts::chp::CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.is_bold, Some(true));
        assert_eq!(properties.is_italic, Some(true));
        let previous = properties.preserved_properties_for_revision.unwrap();
        assert_eq!(previous.is_bold, Some(true));
        assert_eq!(previous.is_italic, None);

        let mut writer = DocWriter::new();
        writer
            .add_paragraph_runs(
                vec![("Tracked".to_string(), formatting)],
                ParagraphFormatting::default(),
            )
            .unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();
        let runs = paragraphs[0].runs().unwrap();
        let properties = runs[0].properties();
        assert_eq!(properties.is_bold, Some(true));
        assert_eq!(properties.is_italic, Some(true));
        let previous = properties
            .preserved_properties_for_revision
            .as_ref()
            .unwrap();
        assert_eq!(previous.is_bold, Some(true));
        assert_eq!(previous.is_italic, None);
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

        let nested = CharacterFormatting {
            preserved_properties_for_revision: Some(Box::new(CharacterFormatting {
                preserved_properties_for_revision: Some(Box::new(CharacterFormatting::default())),
                ..CharacterFormatting::default()
            })),
            ..CharacterFormatting::default()
        };
        assert!(error_for(nested).contains("nested preserved states"));

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
            insertion_revision: Some(TextRevision::new("Alice").with_timestamp(CommentDateTime {
                year: 2026,
                month: 13,
                day: 1,
                hour: 0,
                minute: 0,
                weekday: 0,
            })),
            ..CharacterFormatting::default()
        };
        assert!(error_for(invalid_time).contains("timestamp"));

        let mut writer = DocWriter::new();
        writer.set_section_formatting_revision(FormattingRevision::new("Editor").with_timestamp(
            CommentDateTime {
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
        let mut list = ListStructure::new(42);
        let mut level = crate::doc::writer::numbering::ListLevel::new(3, NumberFormat::Decimal);
        level.number_text = "%1.😀".to_string();
        list.add_level(level);
        writer.add_list(list);
        writer.add_list_override(ListFormatOverride::new(42, 1));
        writer.set_list_names(ListNamesTable::try_new(vec!["Outline".to_string()]).unwrap());
        let template = crate::doc::ListTemplateCode::BuiltIn {
            format: crate::doc::BuiltInListTemplate::ArabicPeriod,
            language: crate::doc::ListTemplateLanguageId::new(0x0409),
        };
        writer.set_list_templates(ListTemplateTable::try_new(vec![Some([template; 9])]).unwrap());
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
        assert_eq!(document.list_names().unwrap().name(0), Some("Outline"));
        assert_eq!(
            document.list_templates().unwrap().get(0).unwrap(),
            &[template; 9]
        );

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
            ..crate::doc::writer::TableRow::default()
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
            ..crate::doc::writer::TableRow::default()
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
            is_header: true,
            ..crate::doc::writer::TableRow::default()
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
                    ..crate::doc::writer::TableRow::default()
                },
            )
            .unwrap();
        assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
    }
}

#[cfg(test)]
mod header_kind_tests {
    use super::*;

    #[test]
    fn header_kinds_map_to_plcfhdd_slots() {
        assert_eq!(DocHeaderKind::Odd.slot(), HEADER_SLOT_ODD);
        assert_eq!(DocHeaderKind::Even.slot(), HEADER_SLOT_EVEN);
        assert_eq!(DocHeaderKind::FirstPage.slot(), HEADER_SLOT_FIRST);
        // The writer's slot assignment matches the MS-DOC PlcfHdd layout:
        // even header 6, odd header 7, first-page header 10.
        assert_eq!(
            (HEADER_SLOT_EVEN, HEADER_SLOT_ODD, HEADER_SLOT_FIRST),
            (6, 7, 10)
        );
    }

    #[test]
    fn header_shape_ids_use_the_header_cluster() {
        let mut writer = DocWriter::new();
        writer
            .insert_header_picture(
                DocHeaderKind::Odd,
                crate::doc::writer::images::DocPicture::from_parts(
                    vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
                    480,
                    240,
                )
                .unwrap(),
                crate::doc::writer::images::FloatingPosition::new(0, 0),
            )
            .unwrap();
        writer
            .insert_header_text_box(
                DocHeaderKind::Even,
                crate::doc::writer::shapes::DocDrawingShape::new(
                    crate::doc::writer::shapes::DocShapeKind::Rectangle,
                    1440,
                    720,
                )
                .unwrap(),
                crate::doc::writer::images::FloatingPosition::new(0, 0),
                "box",
            )
            .unwrap();
        // One shared cluster for both kinds, in insertion order.
        assert_eq!(writer.header_pictures[0].shape_id, 2049);
        assert_eq!(writer.header_shapes[0].shape_id, 2050);
        // Anchors landed in the right header paragraph lists.
        assert_eq!(writer.header_odd.as_ref().unwrap().len(), 1);
        assert_eq!(writer.header_even.as_ref().unwrap().len(), 1);
        assert!(writer.header_first.is_none());
    }
}

#[cfg(test)]
mod chpx_position_hresi_effect_writer_tests {
    use super::*;
    use crate::doc::parts::chp::{CharacterPosition, HresiOperand, HyphenationMode, TextEffect};
    use std::io::Cursor;

    #[test]
    fn emits_canonical_typed_sprms_and_round_trips_package() {
        let position = CharacterPosition::new(-3168).unwrap();
        let hyphenation =
            HresiOperand::with_character(HyphenationMode::DeleteAndChange, b'Z').unwrap();
        let formatting = CharacterFormatting {
            position: Some(position),
            hyphenation: Some(hyphenation),
            text_effect: Some(TextEffect::Shimmer),
            ..CharacterFormatting::default()
        };
        let mut fonts = FontTableBuilder::new();
        let grpprl = build_chpx_grpprl(&formatting, &mut fonts);
        let mut expected = Vec::new();
        expected.extend_from_slice(&SPRM_C_HPS_POS.to_le_bytes());
        expected.extend_from_slice(&(-3168i16).to_le_bytes());
        expected.extend_from_slice(&SPRM_C_HRESI.to_le_bytes());
        expected.extend_from_slice(&[6, b'Z']);
        expected.extend_from_slice(&SPRM_C_SFXT_TEXT.to_le_bytes());
        expected.push(6);
        assert_eq!(grpprl, expected);

        let properties = crate::doc::parts::chp::CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.position, position);
        assert_eq!(properties.hyphenation, hyphenation);
        assert_eq!(properties.text_effect, TextEffect::Shimmer);

        let mut writer = DocWriter::new();
        writer
            .add_paragraph_runs(
                vec![("effects".to_string(), formatting)],
                ParagraphFormatting::default(),
            )
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let mut package =
            crate::doc::Package::from_reader(Cursor::new(output.into_inner())).unwrap();
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();
        let runs = paragraphs[0].runs().unwrap();
        let properties = runs[0].properties();
        assert_eq!(properties.position, position);
        assert_eq!(properties.hyphenation, hyphenation);
        assert_eq!(properties.text_effect, TextEffect::Shimmer);
    }
}
