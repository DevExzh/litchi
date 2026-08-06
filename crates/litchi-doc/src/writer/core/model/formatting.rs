use super::{
    DisplayFieldRevision, DropCap, FontAlignment, FormattingRevision, FrameAnchor, FrameHeight,
    FrameHorizontalPosition, FrameTextFlow, FrameTextWrap, FrameVerticalPosition,
    LegacyAutoNumbering, LegacyBorderPosition, LegacyBorderStyle, NumberingRevision,
    ParagraphBorders, ParagraphShading, PhysicalJustification, TabStop, TextBoxTightWrap,
    TextRevision, WriteError,
};

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
    pub fn multiple_240ths(value: u16) -> Result<Self, WriteError> {
        if !(1..=31_680).contains(&value) {
            return Err(WriteError::InvalidData(format!(
                "line-spacing multiple {value} is outside the LSPD range 1..=31680"
            )));
        }
        Ok(Self {
            dya_line: value as i16,
            is_multiple: true,
        })
    }

    /// Create minimum line spacing in twips.
    pub fn at_least_twips(value: u16) -> Result<Self, WriteError> {
        if !(1..=31_680).contains(&value) {
            return Err(WriteError::InvalidData(format!(
                "minimum line spacing {value} twips is outside the LSPD range 1..=31680"
            )));
        }
        Ok(Self {
            dya_line: value as i16,
            is_multiple: false,
        })
    }

    /// Create exact line spacing in twips.
    pub fn exact_twips(value: u16) -> Result<Self, WriteError> {
        if !(1..=31_680).contains(&value) {
            return Err(WriteError::InvalidData(format!(
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
