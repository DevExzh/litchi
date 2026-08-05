//! Public section-layout model for Word 97+ documents.

use crate::NumberFormat;

pub mod borders;
pub mod columns;

/// A section and the character-position range to which its properties apply.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocSection {
    /// Inclusive start character position in the main document story.
    pub start_cp: u32,
    /// Exclusive end character position in the main document story.
    pub end_cp: u32,
    /// The break that terminates this section.
    pub break_kind: SectionBreakKind,
    /// Page geometry for this section.
    pub page: SectionPageLayout,
    /// Column geometry for this section.
    pub columns: columns::Layout,
    /// Page-number field behavior for this section.
    pub page_numbering: SectionPageNumbering,
    /// Printed line-number behavior for this section.
    pub line_numbering: SectionLineNumbering,
    /// Footnote and endnote placement and numbering overrides.
    pub notes: SectionNoteSettings,
    /// Section-level protection, direction, title-page, and revision behavior.
    pub behavior: SectionBehavior,
    /// Printer-specific paper-source and paper-kind selections.
    pub paper: SectionPaperSettings,
    /// Page-border edges and their section-wide placement controls.
    pub page_borders: borders::Borders,
    /// East Asian document-grid settings for this section.
    pub page_grid: SectionPageGrid,
    /// Orientation and sequencing of glyphs and lines in this section.
    pub text_flow: SectionTextFlow,
}

/// The kind of break that terminates a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionBreakKind {
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
pub struct SectionPageNumbering {
    pub chapter_separator: ChapterNumberSeparator,
    /// Heading level 1 through 9, or `None` when chapter numbers are hidden.
    pub chapter_heading_level: Option<u8>,
    pub number_format: NumberFormat,
    pub restart: bool,
    pub start_at: u32,
}

impl Default for SectionPageNumbering {
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
pub struct SectionLineNumbering {
    /// Distance in lines between labels; zero disables line numbering.
    pub interval: u16,
    pub restart: LineNumberRestart,
    /// Distance from the text in twips; zero requests automatic positioning.
    pub distance_twips: u16,
    pub start_at: u16,
}

impl Default for SectionLineNumbering {
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
pub enum SectionVerticalJustification {
    Top,
    Center,
    Justified,
    Bottom,
}

/// Position of footnotes on a section page.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionFootnotePosition {
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
pub struct SectionNoteSettings {
    pub show_endnotes_at_section_end: bool,
    pub footnote_position: SectionFootnotePosition,
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

impl Default for SectionNoteSettings {
    fn default() -> Self {
        Self {
            show_endnotes_at_section_end: true,
            footnote_position: SectionFootnotePosition::BottomOfPage,
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
pub enum SectionProtection {
    /// No SEPX override; use the document-wide protection setting.
    DocumentDefault,
    Protected,
    Unprotected,
}

/// Section-level behavior that is independent of page geometry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionBehavior {
    pub protection: SectionProtection,
    pub different_first_page: bool,
    pub right_to_left: bool,
    pub right_to_left_gutter: bool,
    pub preserve_properties_for_revision: bool,
    pub revision_save_id: Option<u32>,
}

impl Default for SectionBehavior {
    fn default() -> Self {
        Self {
            protection: SectionProtection::DocumentDefault,
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
pub struct SectionPaperSettings {
    pub first_page_source: Option<u16>,
    pub other_page_source: Option<u16>,
    pub requested_paper_kind: Option<u16>,
}

/// Document-grid mode applied to a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageGridMode {
    Disabled,
    CharactersAndLines,
    LinesOnly,
    EnforceCharacterGrid,
}

/// Character and line pitch used by a section's document grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageGrid {
    pub mode: SectionPageGridMode,
    /// Difference from the Normal-style font pitch, in 1/4096-point units.
    pub character_pitch_adjustment: i32,
    /// Explicit grid line height in twips. Required whenever `mode` is enabled.
    pub line_pitch_twips: Option<u16>,
}

impl Default for SectionPageGrid {
    fn default() -> Self {
        Self {
            mode: SectionPageGridMode::Disabled,
            character_pitch_adjustment: 0,
            line_pitch_twips: None,
        }
    }
}

/// Text-flow rules applied to all text in a section.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum SectionTextFlow {
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
pub struct SectionMargins {
    pub left_twips: u16,
    pub right_twips: u16,
    pub top: VerticalMargin,
    pub bottom: VerticalMargin,
    pub gutter_twips: u16,
}

/// Page geometry and header/footer distances for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageLayout {
    pub width_twips: u16,
    pub height_twips: u16,
    pub orientation: PageOrientation,
    pub margins: SectionMargins,
    pub header_distance_twips: u16,
    pub footer_distance_twips: u16,
    pub vertical_justification: SectionVerticalJustification,
}
