//! Writer-facing document state, input models, and authoring methods.

use crate::CommentDateTime;
use crate::encryption::{DocEncryptionProfile, validate_writer_password};
use crate::parts::pap::{
    Borders as ParagraphBorders, DropCap, FontAlignment, FrameAnchor, FrameHeight,
    FrameHorizontalPosition, FrameTextFlow, FrameTextWrap, FrameVerticalPosition,
    LegacyAutoNumbering, LegacyBorderPosition, LegacyBorderStyle, PhysicalJustification,
    Shading as ParagraphShading, TabStop, TextBoxTightWrap,
};
use crate::parts::{list_names::ListNamesTable, list_templates::ListTemplateTable};
use crate::writer::bookmarks::BookmarkEntry;
use crate::writer::comments::CommentEntry;
use crate::writer::footnotes::FootnoteEntry;
use crate::writer::numbering::{ListFormatOverride, ListStructure, NumberingWriter};
use crate::writer::revisions::{
    DisplayFieldRevision, FormattingRevision, NumberingRevision, TextRevision,
};
use crate::writer::smart_tags::DocSmartTagEntry;
use crate::{
    AssociatedStringSlot, DocumentAssociatedStrings, GlossaryMetadata, ProofingFeature,
    ProofingStateTable, ProofingTables, SavedByTable, SmartTagRecognizerRange,
};
use litchi_cfb::OleError;
use std::collections::HashMap;
use zeroize::Zeroizing;

pub(super) const WORD_DOCUMENT_CLSID: [u8; 16] = [
    0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
pub(super) const VBA_PROJECT_STORAGE_NAME: &str = "Macros";

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

pub(super) fn utf16_code_unit_len(text: &str) -> Result<u32, DocWriteError> {
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

pub(super) const MAX_HEADER_FOOTER_PARAGRAPHS: usize = 65_535;
pub(super) const MAX_HEADER_FOOTER_RUNS: usize = 65_535;
pub(super) const MAX_HEADER_FIELD_DEPTH: usize = 128;

#[derive(Default)]
pub(super) struct HeaderFieldState {
    pub(super) separator_seen: Vec<bool>,
}

impl HeaderFieldState {
    pub(super) fn observe(
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

    pub(super) fn finish(self) -> Result<(), DocWriteError> {
        if self.separator_seen.is_empty() {
            Ok(())
        } else {
            Err(DocWriteError::InvalidData(
                "DOC header/footer field is not terminated within its story".to_string(),
            ))
        }
    }
}

pub(super) fn checked_text_fc(
    text_fc_start: u32,
    stream_length: usize,
) -> Result<u32, DocWriteError> {
    let stream_length = u32::try_from(stream_length).map_err(|_| {
        DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
    })?;
    text_fc_start
        .checked_add(stream_length)
        .ok_or_else(|| DocWriteError::InvalidData("DOC text stream FC range overflows".to_string()))
}

pub(super) fn validate_header_footer_paragraphs(
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

pub(crate) fn pack_dttm(value: Option<CommentDateTime>) -> Result<u32, DocWriteError> {
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

pub(super) type NoteStoryData = (Vec<u8>, Vec<u8>, u32);
pub(super) struct HeaderStoryData {
    pub(super) plcfhdd: Vec<u8>,
    pub(super) fields: Vec<u8>,
    pub(super) char_count: u32,
    /// Story-relative anchor CPs of header floating items with their kind
    /// (in story order, which is CP-ascending by construction).
    pub(super) shape_anchor_cps: Vec<(u32, FloatingAnchorKind)>,
}

/// PlcfHdd slot of the odd page header, which Word uses as the default
/// header when the document does not use facing pages.
pub(super) const HEADER_SLOT_ODD: usize = 7;
/// PlcfHdd slot of the even page header.
pub(super) const HEADER_SLOT_EVEN: usize = 6;
/// PlcfHdd slot of the first page header.
pub(super) const HEADER_SLOT_FIRST: usize = 10;

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
    pub(super) fn slot(self) -> usize {
        match self {
            Self::Odd => HEADER_SLOT_ODD,
            Self::Even => HEADER_SLOT_EVEN,
            Self::FirstPage => HEADER_SLOT_FIRST,
        }
    }
}

/// A floating-item anchor paragraph appended to a header's paragraphs.
pub(super) struct HeaderAnchor {
    /// PlcfHdd slot of the header holding the anchor.
    pub(super) slot: usize,
    /// Paragraph index within that slot's paragraph list.
    pub(super) paragraph_index: usize,
    /// Which floating item the anchor belongs to.
    pub(super) kind: FloatingAnchorKind,
}

pub(super) struct CommentStoryData {
    pub(super) owners: Vec<u8>,
    pub(super) references: Vec<u8>,
    pub(super) text_positions: Vec<u8>,
    pub(super) bookmark_names: Vec<u8>,
    pub(super) bookmark_starts: Vec<u8>,
    pub(super) bookmark_ends: Vec<u8>,
    pub(super) extended_metadata: Vec<u8>,
    pub(super) char_count: u32,
}

pub(super) struct BookmarkTableData {
    pub(super) names: Vec<u8>,
    pub(super) starts: Vec<u8>,
    pub(super) ends: Vec<u8>,
}

pub(super) struct RevisionWriterData {
    pub(super) indexes: HashMap<String, u16>,
    pub(super) table: Vec<u8>,
}

#[derive(Clone, Copy)]
pub(super) enum MainReferenceKind {
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
    pub position: Option<crate::parts::chp::CharacterPosition>,
    /// Word-breaking behavior used when this run is hyphenated.
    pub hyphenation: Option<crate::parts::chp::HresiOperand>,
    /// Animated text effect applied to this run.
    pub text_effect: Option<crate::parts::chp::TextEffect>,
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
pub(super) struct TextRun {
    /// Text content
    pub(super) text: String,
    /// Character formatting
    pub(super) formatting: CharacterFormatting,
    /// Index into `DocWriter::pictures` when this run is a picture
    /// (a single 0x0001 inline or 0x0008 floating picture character).
    pub(super) picture_index: Option<u32>,
    /// Index into `DocWriter::shapes` when this run is a floating
    /// drawing-shape anchor (a single 0x0008 character).
    pub(super) shape_index: Option<u32>,
}

/// Represents a paragraph
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
pub(super) struct WritableParagraph {
    /// Text runs in this paragraph
    pub(super) runs: Vec<TextRun>,
    /// Paragraph formatting
    pub(super) formatting: ParagraphFormatting,
}

pub(super) fn writable_paragraph_from_runs(
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
    pub(super) runs: Vec<(String, CharacterFormatting)>,
    pub(super) formatting: ParagraphFormatting,
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
pub(super) struct TableCell {
    /// Paragraphs in the cell
    pub(super) paragraphs: Vec<WritableParagraph>,
}

/// Represents a table row
#[derive(Debug, Clone)]
pub(super) struct TableRow {
    /// Cells in the row
    pub(super) cells: Vec<TableCell>,
    /// Row and cell layout encoded in the row mark's TAP properties.
    pub(super) formatting: crate::writer::tap::TableRow,
}

/// Represents a table
#[derive(Debug, Clone)]
pub(super) struct WritableTable {
    /// Rows in the table
    pub(super) rows: Vec<TableRow>,
}

/// A picture queued for embedding, with its placement mode.
#[derive(Debug, Clone)]
pub(super) struct WriterPicture {
    /// The picture data and display dimensions.
    pub(super) picture: crate::writer::images::DocPicture,
    /// Shape id allocated at insert time (shared sequence with shapes).
    pub(super) shape_id: u32,
    /// Position and wrapping when the picture floats; `None` for inline.
    pub(super) floating: Option<crate::writer::images::FloatingPosition>,
}

/// A primitive drawing shape queued for the drawing layer.
#[derive(Debug, Clone)]
pub(super) struct WriterShape {
    /// The shape geometry, size, and colors.
    pub(super) shape: crate::writer::shapes::Shape,
    /// Shape id allocated at insert time (shared sequence with pictures).
    pub(super) shape_id: u32,
    /// Position and wrapping.
    pub(super) position: crate::writer::images::FloatingPosition,
    /// Textbox story text when the shape is a text box.
    pub(super) text: Option<String>,
}

/// What kind of floating content a 0x0008 anchor character refers to.
#[derive(Debug, Clone, Copy)]
pub(super) enum FloatingAnchorKind {
    /// Index into `DocWriter::pictures`.
    Picture(u32),
    /// Index into `DocWriter::shapes`.
    Shape(u32),
}

/// DOC file writer
///
/// Provides methods to create and modify DOC files.
pub struct DocWriter {
    /// Paragraphs in the document
    pub(super) paragraphs: Vec<WritableParagraph>,
    /// Tables in the document
    pub(super) tables: Vec<WritableTable>,
    /// Document properties
    pub(super) properties: HashMap<String, String>,
    /// Header/footer paragraphs (`None` means the story is not set).
    /// Indices map to plcfHdd entries (following Apache POI HeaderStories indexing):
    /// 0..5: footnote/endnote separators (unused here)
    /// 6: even header, 7: odd header, 10: first header
    /// 8: even footer, 9: odd footer, 11: first footer
    pub(super) header_even: Option<Vec<HeaderFooterParagraph>>,
    pub(super) header_odd: Option<Vec<HeaderFooterParagraph>>,
    pub(super) header_first: Option<Vec<HeaderFooterParagraph>>,
    pub(super) footer_even: Option<Vec<HeaderFooterParagraph>>,
    pub(super) footer_odd: Option<Vec<HeaderFooterParagraph>>,
    pub(super) footer_first: Option<Vec<HeaderFooterParagraph>>,
    /// Footnote entries
    pub(super) footnotes: Vec<FootnoteEntry>,
    /// Endnote entries
    pub(super) endnotes: Vec<FootnoteEntry>,
    /// Comments
    pub(super) comments: Vec<CommentEntry>,
    /// Standard bookmarks
    pub(super) bookmarks: Vec<BookmarkEntry>,
    /// Embedded smart-tag bookmarks and property bags.
    pub(super) smart_tags: Vec<DocSmartTagEntry>,
    /// Smart-tag recognizer processing-state ranges.
    pub(super) smart_tag_recognizer_ranges: Vec<SmartTagRecognizerRange>,
    /// Optional spelling and grammar proofing-state PLCFs.
    pub(super) proofing_tables: ProofingTables,
    /// Mandatory fixed associated-document string table.
    pub(super) associated_strings: DocumentAssociatedStrings,
    /// Optional Word 97/2000 save-history table.
    pub(super) saved_by_table: Option<SavedByTable>,
    /// Optional glossary-only AutoText metadata over the main story.
    pub(super) glossary_metadata: Option<GlossaryMetadata>,
    /// Optional distinct AutoText-only document attached to this template.
    pub(super) attached_glossary: Option<Box<DocWriter>>,
    /// Property revision metadata for the writer's single document section
    pub(super) section_formatting_revision: Option<FormattingRevision>,
    /// Explicit column geometry for the writer's single document section.
    pub(super) section_columns: Option<crate::section::columns::Layout>,
    /// Whether section columns are populated from right to left.
    pub(super) section_right_to_left: bool,
    /// Section-wide glyph and line flow.
    pub(super) section_text_flow: crate::TextFlow,
    /// Explicit page-border edges and placement for the single section.
    pub(super) section_page_borders: Option<crate::section::borders::Borders>,
    /// Numbering writer for list tables
    pub(super) numbering: NumberingWriter,
    /// User-defined styles appended after the fifteen fixed style slots
    pub(super) styles: Vec<crate::writer::stylesheet::DocStyleDefinition>,
    /// Inline pictures embedded via [`DocWriter::insert_picture`]
    pub(super) pictures: Vec<WriterPicture>,
    /// Primitive drawing shapes embedded via [`DocWriter::insert_floating_shape`]
    pub(super) shapes: Vec<WriterShape>,
    /// Text boxes anchored in the header story, in insertion order.
    pub(super) header_shapes: Vec<WriterShape>,
    /// Pictures anchored in the header story, in insertion order.
    pub(super) header_pictures: Vec<WriterPicture>,
    /// Anchor paragraphs appended to header paragraph lists, in insertion
    /// order (one per header floating item).
    pub(super) header_anchors: Vec<HeaderAnchor>,
    /// Next shape id to allocate (shared by pictures and drawing shapes).
    pub(super) next_shape_id: u32,
    /// Password-to-open settings. The password is wiped when replaced, cleared, or dropped.
    pub(super) encryption: Option<DocWriterEncryption>,
    /// Complete inert MS-OVBA project written under the MS-DOC `Macros` storage.
    pub(super) vba_project: Option<litchi_vba::Payload>,
}

pub(super) struct DocWriterEncryption {
    pub(super) profile: DocEncryptionProfile,
    pub(super) password: Zeroizing<String>,
}

/// Append one textbox story (main or header) to the text stream.
///
/// Per text box: its paragraphs (each `\r`-terminated, with `\n`/`\r`/`"\r\n"`
/// as input separators) plus a trailing CR; one story-final CR is included
/// in the returned story character count. Returns the story-relative start
/// CP of each text box and the total story length (a ccp value).

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
            section_text_flow: crate::TextFlow::default(),
            section_page_borders: None,
            numbering: NumberingWriter::new(),
            styles: Vec::new(),
            pictures: Vec::new(),
            shapes: Vec::new(),
            header_shapes: Vec::new(),
            header_pictures: Vec::new(),
            header_anchors: Vec::new(),
            next_shape_id: crate::writer::images::FIRST_PICTURE_SHAPE_ID,
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
}

impl DocWriter {
    /// Add a custom paragraph, character, table, or numbering style.
    ///
    /// Custom styles occupy consecutive indices beginning at 15. The returned
    /// index can be used by the corresponding formatting properties.
    pub fn add_style(
        &mut self,
        style: crate::writer::stylesheet::DocStyleDefinition,
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
        picture: crate::writer::images::DocPicture,
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
        picture: crate::writer::images::DocPicture,
        position: crate::writer::images::FloatingPosition,
    ) -> Result<(), DocWriteError> {
        self.insert_picture_run(picture, Some(position), "\u{0008}")
    }

    /// Shared tail of `insert_picture`/`insert_floating_picture`: queue the
    /// picture and append a single-character anchor paragraph.
    fn insert_picture_run(
        &mut self,
        picture: crate::writer::images::DocPicture,
        floating: Option<crate::writer::images::FloatingPosition>,
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
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
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
    /// [`crate::writer::shapes::Shape`].
    pub fn insert_floating_shape(
        &mut self,
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
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
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
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
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
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
        picture: crate::writer::images::DocPicture,
        position: crate::writer::images::FloatingPosition,
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
        Ok(crate::writer::images::HEADER_FIRST_SHAPE_ID + index)
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
}

impl DocWriter {
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
                formatting: crate::writer::tap::TableRow {
                    cells: Vec::with_capacity(cols),
                    ..crate::writer::tap::TableRow::default()
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
                row.formatting.cells.push(crate::writer::tap::TableCell {
                    width: (right - left) as u16,
                    merged: false,
                    ..crate::writer::tap::TableCell::default()
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
        columns: crate::section::columns::Layout,
    ) -> Result<(), DocWriteError> {
        columns
            .validate()
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
        self.section_columns = Some(columns);
        Ok(())
    }

    /// Return explicit section column geometry, or `None` for the file-format default.
    pub fn section_columns(&self) -> Option<&crate::section::columns::Layout> {
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
    pub fn set_section_text_flow(&mut self, value: crate::TextFlow) {
        self.section_text_flow = value;
    }

    pub fn section_text_flow(&self) -> crate::TextFlow {
        self.section_text_flow
    }

    /// Set validated page borders for the writer's single section.
    pub fn set_section_page_borders(
        &mut self,
        borders: crate::section::borders::Borders,
    ) -> Result<(), DocWriteError> {
        borders
            .validate()
            .map_err(|error| DocWriteError::InvalidData(error.to_string()))?;
        self.section_page_borders = Some(borders);
        Ok(())
    }

    /// Return explicit page borders, or `None` for the file-format default.
    pub fn section_page_borders(&self) -> Option<&crate::section::borders::Borders> {
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
        formatting: crate::writer::tap::TableRow,
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
        crate::writer::tap::generate_row_sprms(&formatting)
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
