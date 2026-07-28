//! Public section-layout model for Word 97+ documents.

use crate::doc::NumberFormat;
use std::fmt;

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
    pub columns: SectionColumnLayout,
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
    pub page_borders: SectionPageBorders,
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

/// Pages in a section to which its page borders apply.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageBorderApplyTo {
    AllPages,
    FirstPage,
    AllButFirstPage,
}

/// Z-order of page borders relative to the section contents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageBorderDepth {
    InFront,
    Behind,
}

/// Reference from which a page border's spacing is measured.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageBorderOffsetFrom {
    Text,
    PageEdge,
}

/// A validated Word page-border art code.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageBorderArt(u8);

impl SectionPageBorderArt {
    /// Return the `BrcType` art code in the inclusive range `0x40..=0xE3`.
    pub fn code(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for SectionPageBorderArt {
    type Error = u8;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        if (0x40..=0xE3).contains(&value) {
            Ok(Self(value))
        } else {
            Err(value)
        }
    }
}

/// Line or image style of a Word 97 `Brc80` page border.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageBorderStyle {
    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Art(SectionPageBorderArt),
}

/// Palette color selected by a Word `Ico` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageBorderColor {
    Automatic,
    Black,
    Blue,
    Cyan,
    Green,
    Magenta,
    Red,
    Yellow,
    White,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
}

/// One section page-border edge decoded from `Brc80`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageBorder {
    pub style: SectionPageBorderStyle,
    /// Width in eighths of a point. Values below two render as two.
    pub width_eighth_points: u8,
    pub color: SectionPageBorderColor,
    /// Distance from text or the page edge, in points.
    pub spacing_points: u8,
    pub shadow: bool,
    pub frame: bool,
}

/// Invalid caller-supplied `Brc80` page-border data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionPageBorderError {
    /// `Brc80.dptSpace` is a five-bit value.
    InvalidSpacing(u8),
}

impl fmt::Display for SectionPageBorderError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSpacing(value) => write!(
                formatter,
                "section page-border spacing {value} exceeds the 31-point Brc80 limit"
            ),
        }
    }
}

impl std::error::Error for SectionPageBorderError {}

impl SectionPageBorder {
    /// Validate fields whose domains are wider in Rust than in `Brc80`.
    pub fn validate(self) -> Result<(), SectionPageBorderError> {
        if self.spacing_points > 31 {
            return Err(SectionPageBorderError::InvalidSpacing(
                self.spacing_points,
            ));
        }
        Ok(())
    }
}

/// Page borders and shared placement controls for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageBorders {
    pub top: Option<SectionPageBorder>,
    pub left: Option<SectionPageBorder>,
    pub bottom: Option<SectionPageBorder>,
    pub right: Option<SectionPageBorder>,
    pub apply_to: SectionPageBorderApplyTo,
    pub depth: SectionPageBorderDepth,
    pub offset_from: SectionPageBorderOffsetFrom,
}

impl Default for SectionPageBorders {
    fn default() -> Self {
        Self {
            top: None,
            left: None,
            bottom: None,
            right: None,
            apply_to: SectionPageBorderApplyTo::AllPages,
            depth: SectionPageBorderDepth::InFront,
            offset_from: SectionPageBorderOffsetFrom::Text,
        }
    }
}

impl SectionPageBorders {
    /// Validate every present border edge.
    pub fn validate(self) -> Result<(), SectionPageBorderError> {
        for border in [self.top, self.left, self.bottom, self.right]
            .into_iter()
            .flatten()
        {
            border.validate()?;
        }
        Ok(())
    }
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

/// Column geometry for a section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SectionColumnLayout {
    /// Equal-width columns separated by a common spacing value.
    Even {
        count: u8,
        spacing_twips: u16,
        line_between: bool,
    },
    /// Individually sized columns and their following spacing.
    Unequal {
        columns: Vec<SectionColumn>,
        line_between: bool,
    },
}

/// Validation failure for a section column layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SectionColumnError {
    InvalidCount(usize),
    InvalidWidth { index: usize, width_twips: u16 },
    InvalidSpacing { index: usize, spacing_twips: u16 },
    MissingSpacing { index: usize },
    FinalColumnHasSpacing,
}

impl fmt::Display for SectionColumnError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidCount(count) => {
                write!(formatter, "section column count {count} is outside 1..=44")
            },
            Self::InvalidWidth { index, width_twips } => write!(
                formatter,
                "section column {index} width {width_twips} is outside 718..=31680 twips"
            ),
            Self::InvalidSpacing {
                index,
                spacing_twips,
            } => write!(
                formatter,
                "section column {index} spacing {spacing_twips} exceeds 31680 twips"
            ),
            Self::MissingSpacing { index } => {
                write!(formatter, "section column {index} is missing following spacing")
            },
            Self::FinalColumnHasSpacing => {
                formatter.write_str("the final section column cannot have following spacing")
            },
        }
    }
}

impl std::error::Error for SectionColumnError {}

impl SectionColumnLayout {
    pub const MAX_COLUMNS: usize = 44;
    pub const MAX_TWIPS: u16 = 31_680;
    pub const MIN_UNEQUAL_WIDTH_TWIPS: u16 = 718;

    /// Construct and validate an equal-width layout.
    pub fn even(
        count: u8,
        spacing_twips: u16,
        line_between: bool,
    ) -> Result<Self, SectionColumnError> {
        let layout = Self::Even {
            count,
            spacing_twips,
            line_between,
        };
        layout.validate()?;
        Ok(layout)
    }

    /// Construct and validate an unequal-width layout.
    pub fn unequal(
        columns: Vec<SectionColumn>,
        line_between: bool,
    ) -> Result<Self, SectionColumnError> {
        let layout = Self::Unequal {
            columns,
            line_between,
        };
        layout.validate()?;
        Ok(layout)
    }

    /// Validate all cross-field constraints without depending on SPRM order.
    pub fn validate(&self) -> Result<(), SectionColumnError> {
        let count = self.count();
        if !(1..=Self::MAX_COLUMNS).contains(&count) {
            return Err(SectionColumnError::InvalidCount(count));
        }
        match self {
            Self::Even { spacing_twips, .. } => {
                if *spacing_twips > Self::MAX_TWIPS {
                    return Err(SectionColumnError::InvalidSpacing {
                        index: 0,
                        spacing_twips: *spacing_twips,
                    });
                }
            },
            Self::Unequal { columns, .. } => {
                for (index, column) in columns.iter().enumerate() {
                    if !(Self::MIN_UNEQUAL_WIDTH_TWIPS..=Self::MAX_TWIPS)
                        .contains(&column.width_twips)
                    {
                        return Err(SectionColumnError::InvalidWidth {
                            index,
                            width_twips: column.width_twips,
                        });
                    }
                    if index + 1 == columns.len() {
                        if column.spacing_after_twips.is_some() {
                            return Err(SectionColumnError::FinalColumnHasSpacing);
                        }
                    } else {
                        let spacing_twips = column
                            .spacing_after_twips
                            .ok_or(SectionColumnError::MissingSpacing { index })?;
                        if spacing_twips > Self::MAX_TWIPS {
                            return Err(SectionColumnError::InvalidSpacing {
                                index,
                                spacing_twips,
                            });
                        }
                    }
                }
            },
        }
        Ok(())
    }

    /// Replace this layout with a validated equal-width layout.
    pub fn set_even(
        &mut self,
        count: u8,
        spacing_twips: u16,
        line_between: bool,
    ) -> Result<(), SectionColumnError> {
        *self = Self::even(count, spacing_twips, line_between)?;
        Ok(())
    }

    /// Replace this layout with a validated unequal-width layout.
    pub fn set_unequal(
        &mut self,
        columns: Vec<SectionColumn>,
        line_between: bool,
    ) -> Result<(), SectionColumnError> {
        *self = Self::unequal(columns, line_between)?;
        Ok(())
    }

    /// Change only the line-between flag without affecting column geometry.
    pub fn set_line_between(&mut self, value: bool) {
        match self {
            Self::Even { line_between, .. } | Self::Unequal { line_between, .. } => {
                *line_between = value;
            },
        }
    }

    /// Number of columns in this section.
    pub fn count(&self) -> usize {
        match self {
            Self::Even { count, .. } => usize::from(*count),
            Self::Unequal { columns, .. } => columns.len(),
        }
    }

    /// Whether a vertical line is drawn between columns.
    pub fn line_between(&self) -> bool {
        match self {
            Self::Even { line_between, .. } | Self::Unequal { line_between, .. } => *line_between,
        }
    }
}

/// Width and following spacing for one unequal-width column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionColumn {
    pub width_twips: u16,
    /// Space after this column. The final column has no following spacing.
    pub spacing_after_twips: Option<u16>,
}
