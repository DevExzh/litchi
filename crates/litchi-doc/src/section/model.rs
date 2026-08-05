//! Typed semantic section-layout values for Word 97+ documents.

use super::{borders, columns};
use crate::NumberFormat;

/// A section and the character-position range to which its properties apply.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocSection {
    /// Inclusive start character position in the main document story.
    pub start_cp: u32,
    /// Exclusive end character position in the main document story.
    pub end_cp: u32,
    /// The break that terminates this section.
    pub break_kind: BreakKind,
    /// Page geometry for this section.
    pub page: PageLayout,
    /// Column geometry for this section.
    pub columns: columns::Layout,
    /// Page-number field behavior for this section.
    pub page_numbering: PageNumbering,
    /// Printed line-number behavior for this section.
    pub line_numbering: LineNumbering,
    /// Footnote and endnote placement and numbering overrides.
    pub notes: NoteSettings,
    /// Section-level protection, direction, title-page, and revision behavior.
    pub behavior: Behavior,
    /// Printer-specific paper-source and paper-kind selections.
    pub paper: PaperSettings,
    /// Page-border edges and their section-wide placement controls.
    pub page_borders: borders::Borders,
    /// East Asian document-grid settings for this section.
    pub page_grid: PageGrid,
    /// Orientation and sequencing of glyphs and lines in this section.
    pub text_flow: TextFlow,
}

/// The kind of break that terminates a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakKind {
    Continuous,
    NewColumn,
    NewPage,
    EvenPage,
    OddPage,
}

/// Page orientation explicitly stored in section properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// Separator between chapter and page numbers in page-number fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChapterNumberSeparator {
    Hyphen,
    Period,
    Colon,
    EmDash,
    EnDash,
}

/// Page-number field behavior for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageNumbering {
    pub chapter_separator: ChapterNumberSeparator,
    /// Heading level 1 through 9, or `None` when chapter numbers are hidden.
    pub chapter_heading_level: Option<u8>,
    pub number_format: NumberFormat,
    pub restart: bool,
    pub start_at: u32,
}

impl Default for PageNumbering {
    fn default() -> Self {
        Self {
            chapter_separator: ChapterNumberSeparator::Hyphen,
            chapter_heading_level: None,
            number_format: NumberFormat::Arabic,
            restart: false,
            start_at: 0,
        }
    }
}

/// Point at which line numbering restarts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineNumberRestart {
    EachPage,
    EachSection,
    Continuous,
}

/// Printed line-number behavior for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LineNumbering {
    /// Distance in lines between labels; zero disables line numbering.
    pub interval: u16,
    pub restart: LineNumberRestart,
    /// Distance from the text in twips; zero requests automatic positioning.
    pub distance_twips: u16,
    pub start_at: u16,
}

impl Default for LineNumbering {
    fn default() -> Self {
        Self {
            interval: 0,
            restart: LineNumberRestart::EachPage,
            distance_twips: 0,
            start_at: 1,
        }
    }
}

/// Vertical alignment of the section contents between the page margins.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalJustification {
    Top,
    Center,
    Justified,
    Bottom,
}

/// Position of footnotes on a section page.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FootnotePosition {
    BottomOfPage,
    BeneathText,
}

/// Point at which note numbering restarts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteNumberRestart {
    Continuous,
    EachSection,
    EachPage,
}

/// Footnote and endnote overrides stored in one section's SEPX.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NoteSettings {
    pub show_endnotes_at_section_end: bool,
    pub footnote_position: FootnotePosition,
    pub footnote_restart: NoteNumberRestart,
    pub endnote_restart: NoteNumberRestart,
    /// Stored `sprmSNFtn` value. For continuous numbering, the added offset is
    /// this value minus one.
    pub footnote_offset_operand: u16,
    pub footnote_number_format: NumberFormat,
    /// Stored `sprmSNEdn` value. For continuous numbering, the added offset is
    /// this value minus one.
    pub endnote_offset_operand: u16,
    pub endnote_number_format: NumberFormat,
}

impl Default for NoteSettings {
    fn default() -> Self {
        Self {
            show_endnotes_at_section_end: true,
            footnote_position: FootnotePosition::BottomOfPage,
            footnote_restart: NoteNumberRestart::Continuous,
            endnote_restart: NoteNumberRestart::Continuous,
            footnote_offset_operand: 1,
            footnote_number_format: NumberFormat::Arabic,
            endnote_offset_operand: 1,
            endnote_number_format: NumberFormat::LowerRoman,
        }
    }
}

/// Section protection relative to the document's form-field protection mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Protection {
    /// No SEPX override; use the document-wide protection setting.
    DocumentDefault,
    Protected,
    Unprotected,
}

/// Section-level behavior that is independent of page geometry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Behavior {
    pub protection: Protection,
    pub different_first_page: bool,
    pub right_to_left: bool,
    pub right_to_left_gutter: bool,
    pub preserve_properties_for_revision: bool,
    pub revision_save_id: Option<u32>,
}

impl Default for Behavior {
    fn default() -> Self {
        Self {
            protection: Protection::DocumentDefault,
            different_first_page: false,
            right_to_left: false,
            right_to_left_gutter: false,
            preserve_properties_for_revision: false,
            revision_save_id: None,
        }
    }
}

/// Printer-specific paper selections retained without platform interpretation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PaperSettings {
    pub first_page_source: Option<u16>,
    pub other_page_source: Option<u16>,
    pub requested_paper_kind: Option<u16>,
}

/// Document-grid mode applied to a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageGridMode {
    Disabled,
    CharactersAndLines,
    LinesOnly,
    EnforceCharacterGrid,
}

/// Character and line pitch used by a section's document grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageGrid {
    pub mode: PageGridMode,
    /// Difference from the Normal-style font pitch, in 1/4096-point units.
    pub character_pitch_adjustment: i32,
    /// Explicit grid line height in twips. Required whenever `mode` is enabled.
    pub line_pitch_twips: Option<u16>,
}

impl Default for PageGrid {
    fn default() -> Self {
        Self {
            mode: PageGridMode::Disabled,
            character_pitch_adjustment: 0,
            line_pitch_twips: None,
        }
    }
}

/// Text-flow rules applied to all text in a section.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TextFlow {
    #[default]
    HorizontalNonAsian,
    TopToBottomAsian,
    BottomToTop,
    TopToBottomNonAsian,
    HorizontalAsian,
    VerticalNonAsian,
}

/// A vertical page margin as defined by `sprmSDyaTop` and `sprmSDyaBottom`.
///
/// Positive values are minimum margins that can grow to accommodate headers,
/// footers, or footnotes. Negative values are fixed distances whose absolute
/// value must be used.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalMargin {
    Minimum(u16),
    Fixed(u16),
}

impl VerticalMargin {
    /// Return the lossless signed-twips representation used by the file format.
    pub fn signed_twips(self) -> i16 {
        match self {
            Self::Minimum(value) => value as i16,
            Self::Fixed(value) => -(value as i16),
        }
    }

    /// Return the physical distance in twips, independent of margin behavior.
    pub fn distance_twips(self) -> u16 {
        match self {
            Self::Minimum(value) | Self::Fixed(value) => value,
        }
    }

    pub(crate) fn from_signed_twips(value: i16) -> Self {
        if value < 0 {
            Self::Fixed(value.unsigned_abs())
        } else {
            Self::Minimum(value as u16)
        }
    }
}

/// Page margins for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Margins {
    pub left_twips: u16,
    pub right_twips: u16,
    pub top: VerticalMargin,
    pub bottom: VerticalMargin,
    pub gutter_twips: u16,
}

/// Page geometry and header/footer distances for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageLayout {
    pub width_twips: u16,
    pub height_twips: u16,
    pub orientation: PageOrientation,
    pub margins: Margins,
    pub header_distance_twips: u16,
    pub footer_distance_twips: u16,
    pub vertical_justification: VerticalJustification,
}
