//! RTF section support.
//!
//! This module provides support for document sections, headers, footers,
//! page breaks, and section formatting in RTF documents.

use super::types::{Formatting, Paragraph, TextDirection};
use std::borrow::Cow;

/// Maximum number of columns retained for one RTF section.
pub const MAX_SECTION_COLUMNS: u16 = 64;

/// Maximum accepted column width or inter-column spacing, in twips.
pub const MAX_SECTION_COLUMN_TWIPS: i32 = 31_680;

/// Maximum accepted line-number increment.
pub const MAX_SECTION_LINE_INCREMENT: u16 = u16::MAX;

/// Maximum accepted line-number distance from text, in twips.
pub const MAX_SECTION_LINE_DISTANCE: i32 = 31_680;

/// Maximum retained starting line number.
pub const MAX_SECTION_LINE_START: u32 = 1_000_000;

/// Maximum heading level usable as a page-number prefix (`pgnhn`).
pub const MAX_PAGE_NUMBER_HEADING_LEVEL: i32 = 9;

/// Maximum accepted section line-grid pitch, in twips.
pub const MAX_SECTION_LINE_GRID_TWIPS: i32 = 31_680;

/// One explicitly sized section column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionColumn {
    /// Column width in twips (`colw`).
    pub width: i32,
    /// Space to the right in twips (`colsr`).
    pub space_after: Option<i32>,
}

impl SectionColumn {
    /// Construct an explicitly sized column.
    pub fn new(width: i32, space_after: Option<i32>) -> crate::RtfResult<Self> {
        let column = Self { width, space_after };
        column.validate()?;
        Ok(column)
    }

    /// Validate this column against the implementation safety bounds.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if !(1..=MAX_SECTION_COLUMN_TWIPS).contains(&self.width) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF section-column width must be in 1..=31680 twips".to_string(),
            ));
        }
        if self
            .space_after
            .is_some_and(|value| !(0..=MAX_SECTION_COLUMN_TWIPS).contains(&value))
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF section-column spacing must be in 0..=31680 twips".to_string(),
            ));
        }
        Ok(())
    }
}

/// Equal- or variable-width column layout for an RTF section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionColumns {
    /// Number of columns (`cols`), bounded by [`MAX_SECTION_COLUMNS`].
    pub count: u16,
    /// Default spacing for equal-width columns (`colsx`), in twips.
    pub default_spacing: i32,
    /// Whether RTF requests a rule between columns (`linebetcol`).
    pub separator: bool,
    /// Ordered explicit column geometry. Empty means equal-width columns.
    pub explicit: Vec<SectionColumn>,
}

impl SectionColumns {
    /// Construct an equal-width column layout.
    pub fn equal(count: u16, spacing: i32, separator: bool) -> crate::RtfResult<Self> {
        let columns = Self {
            count,
            default_spacing: spacing,
            separator,
            explicit: Vec::new(),
        };
        columns.validate()?;
        Ok(columns)
    }

    /// Construct a variable-width column layout.
    pub fn variable(
        columns: Vec<SectionColumn>,
        default_spacing: i32,
        separator: bool,
    ) -> crate::RtfResult<Self> {
        let count = u16::try_from(columns.len()).map_err(|_| {
            crate::RtfError::MalformedDocument(
                "RTF section-column count exceeds the safety limit".to_string(),
            )
        })?;
        let layout = Self {
            count,
            default_spacing,
            separator,
            explicit: columns,
        };
        layout.validate()?;
        Ok(layout)
    }

    /// Whether the layout contains explicit variable-width geometry.
    pub fn is_variable(&self) -> bool {
        !self.explicit.is_empty()
    }

    /// Validate cardinality and numeric safety bounds.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if !(1..=MAX_SECTION_COLUMNS).contains(&self.count) {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF section-column count must be in 1..={MAX_SECTION_COLUMNS}"
            )));
        }
        if !(0..=MAX_SECTION_COLUMN_TWIPS).contains(&self.default_spacing) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF section-column default spacing must be in 0..=31680 twips".to_string(),
            ));
        }
        if !self.explicit.is_empty() && self.explicit.len() != usize::from(self.count) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF explicit section-column count does not match cols".to_string(),
            ));
        }
        for column in &self.explicit {
            column.validate()?;
        }
        Ok(())
    }
}

impl Default for SectionColumns {
    fn default() -> Self {
        Self {
            count: 1,
            default_spacing: 720,
            separator: false,
            explicit: Vec::new(),
        }
    }
}

/// Section break type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SectionBreakType {
    /// Continuous section (no page break)
    Continuous,
    /// New column
    Column,
    /// New page
    #[default]
    Page,
    /// New even page
    EvenPage,
    /// New odd page
    OddPage,
}

/// Page orientation
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageOrientation {
    /// Portrait orientation
    #[default]
    Portrait,
    /// Landscape orientation
    Landscape,
}

/// Page numbering format
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageNumberFormat {
    /// Arabic numerals (1, 2, 3...)
    #[default]
    Decimal,
    /// Uppercase Roman (I, II, III...)
    UpperRoman,
    /// Lowercase Roman (i, ii, iii...)
    LowerRoman,
    /// Uppercase letters (A, B, C...)
    UpperLetter,
    /// Lowercase letters (a, b, c...)
    LowerLetter,
    /// Bidi Arabic alphabetic (`\pgnbidia`)
    BidiAlphabetic,
    /// Bidi Arabic abjad (`\pgnbidib`)
    BidiAbjad,
    /// Korean Chosung (`\pgnchosung`)
    KoreanChosung,
    /// Enclosed circle numbers (`\pgncnum`)
    Circle,
    /// Kanji without the digit character (`\pgndbnum`)
    KanjiDigitless,
    /// Kanji with the digit character (`\pgndbnumd`)
    KanjiWithDigit,
    /// Kanji numbering 3 (`\pgndbnumt`)
    KanjiThree,
    /// Kanji numbering 4 (`\pgndbnumk`)
    KanjiFour,
    /// Double decimal numbering (`\pgndecd`)
    DoubleDecimal,
    /// Korean Ganada (`\pgnganada`)
    KoreanGanada,
    /// Chinese numbering 1 (`\pgngbnum`)
    ChineseOne,
    /// Chinese numbering 2 (`\pgngbnumd`)
    ChineseTwo,
    /// Chinese numbering 3 (`\pgngbnuml`)
    ChineseThree,
    /// Chinese numbering 4 (`\pgngbnumk`)
    ChineseFour,
    /// Hindi vowels (`\pgnhindia`)
    HindiVowels,
    /// Hindi consonants (`\pgnhindib`)
    HindiConsonants,
    /// Hindi numbers (`\pgnhindic`)
    HindiNumbers,
    /// Hindi descriptive (`\pgnhindid`)
    HindiDescriptive,
    /// Thai letters (`\pgnthaia`)
    ThaiLetters,
    /// Thai numbers (`\pgnthaib`)
    ThaiNumbers,
    /// Thai descriptive (`\pgnthaic`)
    ThaiDescriptive,
    /// Vietnamese cardinal (`\pgnvieta`)
    VietnameseCardinal,
    /// Zodiac numbering 1 (`\pgnzodiac`)
    ZodiacOne,
    /// Zodiac numbering 2 (`\pgnzodiacd`)
    ZodiacTwo,
    /// Zodiac numbering 3 (`\pgnzodiacl`)
    ZodiacThree,
}

impl PageNumberFormat {
    /// The RTF control word that selects this page-number format.
    pub const fn control_word(self) -> &'static str {
        match self {
            Self::Decimal => "pgndec",
            Self::UpperRoman => "pgnucrm",
            Self::LowerRoman => "pgnlcrm",
            Self::UpperLetter => "pgnucltr",
            Self::LowerLetter => "pgnlcltr",
            Self::BidiAlphabetic => "pgnbidia",
            Self::BidiAbjad => "pgnbidib",
            Self::KoreanChosung => "pgnchosung",
            Self::Circle => "pgncnum",
            Self::KanjiDigitless => "pgndbnum",
            Self::KanjiWithDigit => "pgndbnumd",
            Self::KanjiThree => "pgndbnumt",
            Self::KanjiFour => "pgndbnumk",
            Self::DoubleDecimal => "pgndecd",
            Self::KoreanGanada => "pgnganada",
            Self::ChineseOne => "pgngbnum",
            Self::ChineseTwo => "pgngbnumd",
            Self::ChineseThree => "pgngbnuml",
            Self::ChineseFour => "pgngbnumk",
            Self::HindiVowels => "pgnhindia",
            Self::HindiConsonants => "pgnhindib",
            Self::HindiNumbers => "pgnhindic",
            Self::HindiDescriptive => "pgnhindid",
            Self::ThaiLetters => "pgnthaia",
            Self::ThaiNumbers => "pgnthaib",
            Self::ThaiDescriptive => "pgnthaic",
            Self::VietnameseCardinal => "pgnvieta",
            Self::ZodiacOne => "pgnzodiac",
            Self::ZodiacTwo => "pgnzodiacd",
            Self::ZodiacThree => "pgnzodiacl",
        }
    }
}

/// Whether page numbering restarts at this section or continues from the
/// preceding one (`\pgnrestart` / `\pgncont`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageNumberRestart {
    /// Page numbers restart at the `\pgnstartsN` value.
    Restart,
    /// Page numbers continue from the preceding section.
    Continuous,
}

impl PageNumberRestart {
    /// The RTF control word that selects this restart behavior.
    pub const fn control_word(self) -> &'static str {
        match self {
            Self::Restart => "pgnrestart",
            Self::Continuous => "pgncont",
        }
    }
}

/// Vertical alignment of text within a section
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VerticalAlignment {
    /// Top-aligned
    #[default]
    Top,
    /// Centered
    Center,
    /// Justified (distributed)
    Justify,
    /// Bottom-aligned
    Bottom,
}

/// Restart behavior for section line numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionLineNumberRestart {
    /// Restart numbering at each section (`linerestart`).
    Section,
    /// Restart numbering on each page (`lineppage`).
    Page,
    /// Continue numbering from the previous section (`linecont`).
    Continuous,
}

/// Complete explicit section line-numbering state.
///
/// `increment == None` means numbering is disabled. Other fields retain
/// independently authored section controls such as the ubiquitous `linex0`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SectionLineNumbering {
    /// Amount added to each displayed line number (`linemod`).
    pub increment: Option<u16>,
    /// Distance from the text column to the line number (`linex`), in twips.
    pub distance: Option<i32>,
    /// One-based starting line number (`linestarts`).
    pub start: Option<u32>,
    /// Explicit restart mode.
    pub restart: Option<SectionLineNumberRestart>,
}

impl SectionLineNumbering {
    /// Construct enabled line numbering with the given increment.
    pub fn new(increment: u16) -> crate::RtfResult<Self> {
        let value = Self {
            increment: Some(increment),
            ..Self::default()
        };
        value.validate()?;
        Ok(value)
    }

    /// Whether this section explicitly enables line numbering.
    pub fn is_enabled(&self) -> bool {
        self.increment.is_some()
    }

    /// Whether no line-numbering controls were authored.
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Validate line-numbering values against implementation safety bounds.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.increment == Some(0) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF line-number increment must be in 1..=65535".to_string(),
            ));
        }
        if self
            .distance
            .is_some_and(|value| !(0..=MAX_SECTION_LINE_DISTANCE).contains(&value))
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF line-number distance must be in 0..=31680 twips".to_string(),
            ));
        }
        if self
            .start
            .is_some_and(|value| !(1..=MAX_SECTION_LINE_START).contains(&value))
        {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF starting line number must be in 1..={MAX_SECTION_LINE_START}"
            )));
        }
        Ok(())
    }
}

/// Section-level footnote placement override.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionFootnotePlacement {
    BeneathText,
    BottomOfPage,
}

/// Explicit section-level footnote and endnote overrides.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SectionNoteOptions {
    pub footnote_placement: Option<SectionFootnotePlacement>,
    pub footnote_start: Option<i32>,
    pub endnote_start: Option<i32>,
    pub footnote_restart: Option<crate::FootnoteRestart>,
    pub endnote_restart: Option<crate::EndnoteRestart>,
    pub footnote_numbering: Option<crate::NoteNumberingStyle>,
    pub endnote_numbering: Option<crate::NoteNumberingStyle>,
    /// Place endnotes at the end of the section rather than the document
    /// (`\endnhere`).
    pub endnote_here: bool,
}

impl SectionNoteOptions {
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.footnote_start.is_some_and(|value| value <= 0)
            || self.endnote_start.is_some_and(|value| value <= 0)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF section note starting numbers must be positive".to_string(),
            ));
        }
        Ok(())
    }

    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

/// Separator drawn between the heading number and the page number.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageNumberHeadingSeparator {
    /// Hyphen separator (`pgnhnsh`).
    Hyphen,
    /// Period separator (`pgnhnsp`).
    Period,
    /// Colon separator (`pgnhnsc`).
    Colon,
    /// Em-dash separator (`pgnhnsm`).
    EmDash,
    /// En-dash separator (`pgnhnsn`).
    EnDash,
}

impl PageNumberHeadingSeparator {
    /// The canonical RTF control word for this separator.
    pub fn control_word(&self) -> &'static str {
        match self {
            Self::Hyphen => "pgnhnsh",
            Self::Period => "pgnhnsp",
            Self::Colon => "pgnhnsc",
            Self::EmDash => "pgnhnsm",
            Self::EnDash => "pgnhnsn",
        }
    }
}

/// Heading-number prefix applied to section page numbers (`pgnhn` family).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SectionPageNumberHeading {
    /// Heading level used to prefix page numbers (`pgnhnN`). `Some(0)`
    /// explicitly disables the prefix; `None` means the control was not
    /// authored. Values are bounded by [`MAX_PAGE_NUMBER_HEADING_LEVEL`].
    pub level: Option<u8>,
    /// Separator drawn between the heading number and the page number.
    pub separator: Option<PageNumberHeadingSeparator>,
}

impl SectionPageNumberHeading {
    /// Whether no heading-number controls were authored.
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Validate heading-number values against implementation safety bounds.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self
            .level
            .is_some_and(|value| i32::from(value) > MAX_PAGE_NUMBER_HEADING_LEVEL)
        {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF page-number heading level must be in 0..={MAX_PAGE_NUMBER_HEADING_LEVEL}"
            )));
        }
        Ok(())
    }
}

/// Document grid used to lay out a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionDocumentGridType {
    /// Line and character grid (`sectspecifyl`).
    LinesAndCharacters,
    /// Character grid only (`sectspecifycl`).
    CharactersOnly,
    /// Default grid marker (`sectspecifygen`): neither `sectspecifycl` nor
    /// `sectspecifyl` was authored by the producing application.
    Default,
}

impl SectionDocumentGridType {
    /// The canonical RTF control word for this grid type.
    pub fn control_word(&self) -> &'static str {
        match self {
            Self::LinesAndCharacters => "sectspecifyl",
            Self::CharactersOnly => "sectspecifycl",
            Self::Default => "sectspecifygen",
        }
    }
}

/// Explicit section document-grid settings.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SectionDocumentGrid {
    /// Line-grid pitch for the section in twips (`sectlinegridN`).
    pub line_grid: Option<i32>,
    /// Explicit grid type (`sectspecifyl` / `sectspecifycl` /
    /// `sectspecifygen`).
    pub grid_type: Option<SectionDocumentGridType>,
}

impl SectionDocumentGrid {
    /// Whether no document-grid controls were authored.
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Validate grid values against implementation safety bounds.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self
            .line_grid
            .is_some_and(|value| !(0..=MAX_SECTION_LINE_GRID_TWIPS).contains(&value))
        {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF section line-grid pitch must be in 0..={MAX_SECTION_LINE_GRID_TWIPS} twips"
            )));
        }
        Ok(())
    }
}

/// Section properties
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionProperties {
    /// Optional section-style handle referenced by this section.
    pub section_style: Option<u16>,
    /// RSID attached to the section formatting (`\sectrsidN`).
    pub section_rsid: Option<u32>,
    /// Different first page: the first page uses its own header/footer
    /// (`\titlepg`).
    pub title_page: bool,
    /// Explicit direction used to thread section columns.
    pub direction: Option<TextDirection>,
    /// Section break type
    pub break_type: SectionBreakType,
    /// Page width (in twips)
    pub page_width: i32,
    /// Page height (in twips)
    pub page_height: i32,
    /// Left margin (in twips)
    pub margin_left: i32,
    /// Right margin (in twips)
    pub margin_right: i32,
    /// Top margin (in twips)
    pub margin_top: i32,
    /// Bottom margin (in twips)
    pub margin_bottom: i32,
    /// Gutter margin (in twips)
    pub margin_gutter: i32,
    /// Header distance from top (in twips)
    pub header_distance: i32,
    /// Footer distance from bottom (in twips)
    pub footer_distance: i32,
    /// Page orientation
    pub orientation: PageOrientation,
    /// Equal- or variable-width section column layout.
    pub columns: SectionColumns,
    /// Page number start
    pub page_number_start: i32,
    /// Page number format
    pub page_number_format: PageNumberFormat,
    /// Explicit page-number restart behavior (`\pgnrestart` / `\pgncont`).
    pub page_number_restart: Option<PageNumberRestart>,
    /// Horizontal page-number position offset in twips (`\pgnxN`).
    pub page_number_offset_x: Option<i32>,
    /// Vertical page-number position offset in twips (`\pgnyN`).
    pub page_number_offset_y: Option<i32>,
    /// Heading-number prefix applied to page numbers (`\pgnhn` family).
    pub page_number_heading: SectionPageNumberHeading,
    /// Document-grid settings for the section (`\sectlinegridN`,
    /// `\sectspecifyl` / `\sectspecifycl` / `\sectspecifygen`).
    pub document_grid: SectionDocumentGrid,
    /// Author/date metadata for the revision that changed this section's
    /// properties (`\srauthN`, `\srdateN`).
    pub revision: crate::RevisionMetadata,
    /// Vertical alignment
    pub vertical_alignment: VerticalAlignment,
    /// Typed section line-numbering properties.
    pub line_numbering: SectionLineNumbering,
    /// Explicit section-level footnote and endnote overrides.
    pub note_options: SectionNoteOptions,
    /// Page-border edges and placement for this section.
    pub page_borders: crate::PageBorders,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            section_style: None,
            section_rsid: None,
            title_page: false,
            direction: None,
            break_type: SectionBreakType::default(),
            page_width: 12240,  // 8.5 inches at 1440 twips/inch
            page_height: 15840, // 11 inches
            margin_left: 1800,  // 1.25 inches
            margin_right: 1800,
            margin_top: 1440, // 1 inch
            margin_bottom: 1440,
            margin_gutter: 0,
            header_distance: 720, // 0.5 inches
            footer_distance: 720,
            orientation: PageOrientation::default(),
            columns: SectionColumns::default(),
            page_number_start: 1,
            page_number_format: PageNumberFormat::default(),
            page_number_restart: None,
            page_number_offset_x: None,
            page_number_offset_y: None,
            page_number_heading: SectionPageNumberHeading::default(),
            document_grid: SectionDocumentGrid::default(),
            revision: crate::RevisionMetadata::default(),
            vertical_alignment: VerticalAlignment::default(),
            line_numbering: SectionLineNumbering::default(),
            note_options: SectionNoteOptions::default(),
            page_borders: crate::PageBorders::default(),
        }
    }
}

impl SectionProperties {
    /// Set or clear the section-style handle referenced by this section.
    #[inline]
    pub fn set_section_style(&mut self, section_style: Option<u16>) {
        self.section_style = section_style;
    }
}

/// Header/footer type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterType {
    /// Header for all pages
    Header,
    /// Footer for all pages
    Footer,
    /// Header for first page
    HeaderFirst,
    /// Footer for first page
    FooterFirst,
    /// Header for left (even) pages
    HeaderLeft,
    /// Footer for left (even) pages
    FooterLeft,
    /// Header for right (odd) pages
    HeaderRight,
    /// Footer for right (odd) pages
    FooterRight,
}

/// A header or footer content
#[derive(Debug, Clone)]
pub struct HeaderFooter<'a> {
    /// Type of header/footer
    pub header_type: HeaderFooterType,
    /// Content paragraphs
    pub paragraphs: Vec<HeaderFooterParagraph<'a>>,
    /// Positional root shapes owned by this header/footer story.
    pub shapes: Vec<crate::Shape<'a>>,
    /// Positional root shape groups owned by this header/footer story.
    pub shape_groups: Vec<crate::ShapeGroup<'a>>,
    /// Exact source order of drawings in this header/footer story.
    pub drawing_order: Vec<crate::StoryDrawing>,
    /// Exact source order of drawings and generic fields in this story.
    pub story_events: Vec<crate::StoryEvent>,
}

impl<'a> HeaderFooter<'a> {
    /// Create a new header/footer
    #[inline]
    pub fn new(header_type: HeaderFooterType) -> Self {
        Self {
            header_type,
            paragraphs: Vec::new(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
        }
    }

    /// Add a paragraph to the header/footer
    #[inline]
    pub fn add_paragraph(&mut self, paragraph: HeaderFooterParagraph<'a>) {
        self.paragraphs.push(paragraph);
    }

    /// Get the text content
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text.as_ref())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Append a validated positional shape to this story.
    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> crate::RtfResult<()> {
        let mut shapes = self.shapes.clone();
        shapes.push(shape);
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::Shape(self.shapes.len()));
        let text = self.text();
        crate::shape::validate_story_drawings(
            &text,
            &shapes,
            &self.shape_groups,
            &order,
            "header/footer",
        )?;
        self.shapes = shapes;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(
                self.shapes.len() - 1,
            )));
        Ok(())
    }

    /// Append a validated positional root shape group to this story.
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> crate::RtfResult<()> {
        let mut groups = self.shape_groups.clone();
        groups.push(group);
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::ShapeGroup(self.shape_groups.len()));
        let text = self.text();
        crate::shape::validate_story_drawings(
            &text,
            &self.shapes,
            &groups,
            &order,
            "header/footer",
        )?;
        self.shape_groups = groups;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(
                self.shape_groups.len() - 1,
            )));
        Ok(())
    }

    /// Clear all drawings owned by this header/footer story.
    pub fn clear_drawings(&mut self) {
        self.shapes.clear();
        self.shape_groups.clear();
        self.drawing_order.clear();
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::Drawing(_)));
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.story_events.iter().filter_map(|event| match event {
            crate::StoryEvent::PageBreak(page_break) => Some(page_break),
            _ => None,
        })
    }

    pub fn push_page_break(&mut self, position: usize) -> crate::RtfResult<()> {
        let text = self.text();
        crate::field::push_story_page_break(
            &mut self.story_events,
            &text,
            position,
            "header/footer",
        )
    }

    pub fn clear_page_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::PageBreak(_)));
    }
}

/// A paragraph in a header or footer
#[derive(Debug, Clone)]
pub struct HeaderFooterParagraph<'a> {
    /// Text content
    pub text: Cow<'a, str>,
    /// Character formatting
    pub formatting: Formatting,
    /// Paragraph properties
    pub paragraph: Paragraph,
}

impl<'a> HeaderFooterParagraph<'a> {
    /// Create a new header/footer paragraph
    #[inline]
    pub fn new(text: Cow<'a, str>, formatting: Formatting, paragraph: Paragraph) -> Self {
        Self {
            text,
            formatting,
            paragraph,
        }
    }
}

/// RTF section
#[derive(Debug, Clone)]
pub struct Section<'a> {
    /// Section properties
    pub properties: SectionProperties,
    /// Headers and footers for this section
    pub headers_footers: Vec<HeaderFooter<'a>>,
}

impl<'a> Section<'a> {
    /// Create a new section
    #[inline]
    pub fn new() -> Self {
        Self {
            properties: SectionProperties::default(),
            headers_footers: Vec::new(),
        }
    }

    /// Add a header or footer
    #[inline]
    pub fn add_header_footer(&mut self, hf: HeaderFooter<'a>) {
        self.headers_footers.push(hf);
    }

    /// Get header by type
    pub fn get_header(&self, htype: HeaderFooterType) -> Option<&HeaderFooter<'a>> {
        self.headers_footers.iter().find(|h| h.header_type == htype)
    }
}

impl<'a> Default for Section<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// Footnote or endnote
#[derive(Debug, Clone)]
pub struct Note<'a> {
    /// UTF-8 byte offset where the note destination occurs in the main story.
    pub position: usize,
    /// Whether this is a footnote (true) or endnote (false)
    pub is_footnote: bool,
    /// Reference mark (number or symbol)
    pub reference: Cow<'a, str>,
    /// Note content
    pub content: Cow<'a, str>,
    /// Character formatting for the note
    pub formatting: Formatting,
    /// Positional root shapes owned by the note story.
    pub shapes: Vec<crate::Shape<'a>>,
    /// Positional root shape groups owned by the note story.
    pub shape_groups: Vec<crate::ShapeGroup<'a>>,
    /// Exact source order of drawings in this note story.
    pub drawing_order: Vec<crate::StoryDrawing>,
    /// Exact source order of drawings and generic fields in this note story.
    pub story_events: Vec<crate::StoryEvent>,
}

pub(crate) const MAX_NOTES: usize = 65_536;
pub(crate) const MAX_NOTE_BODY_BYTES: usize = 4 * 1_048_576;
pub(crate) const MAX_NOTE_TEXT_TOTAL_BYTES: usize = 16 * 1_048_576;

impl<'a> Note<'a> {
    /// Create a new footnote
    #[inline]
    pub fn footnote(reference: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            position: 0,
            is_footnote: true,
            reference,
            content,
            formatting: Formatting::default(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
        }
    }

    /// Create a new endnote
    #[inline]
    pub fn endnote(reference: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            position: 0,
            is_footnote: false,
            reference,
            content,
            formatting: Formatting::default(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
        }
    }

    /// Validate this note story independently of its main-story anchor.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.content.len() > MAX_NOTE_BODY_BYTES || self.reference.len() > 65_536 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF note text exceeds the safety limit".to_string(),
            ));
        }
        crate::field::validate_story_events(
            self.content.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            &self.story_events,
            "note",
        )
    }

    pub(crate) fn text_bytes(&self) -> Option<usize> {
        self.content.len().checked_add(self.reference.len())
    }

    /// Append a validated positional root shape to this note story.
    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> crate::RtfResult<()> {
        let mut shapes = self.shapes.clone();
        shapes.push(shape);
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::Shape(self.shapes.len()));
        crate::shape::validate_story_drawings(
            self.content.as_ref(),
            &shapes,
            &self.shape_groups,
            &order,
            "note",
        )?;
        self.shapes = shapes;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(
                self.shapes.len() - 1,
            )));
        Ok(())
    }

    /// Append a validated positional root shape group to this note story.
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> crate::RtfResult<()> {
        let mut groups = self.shape_groups.clone();
        groups.push(group);
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::ShapeGroup(self.shape_groups.len()));
        crate::shape::validate_story_drawings(
            self.content.as_ref(),
            &self.shapes,
            &groups,
            &order,
            "note",
        )?;
        self.shape_groups = groups;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(
                self.shape_groups.len() - 1,
            )));
        Ok(())
    }

    /// Clear all drawings owned by this note story.
    pub fn clear_drawings(&mut self) {
        self.shapes.clear();
        self.shape_groups.clear();
        self.drawing_order.clear();
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::Drawing(_)));
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.story_events.iter().filter_map(|event| match event {
            crate::StoryEvent::PageBreak(page_break) => Some(page_break),
            _ => None,
        })
    }

    pub fn push_page_break(&mut self, position: usize) -> crate::RtfResult<()> {
        crate::field::push_story_page_break(
            &mut self.story_events,
            self.content.as_ref(),
            position,
            "note",
        )
    }

    pub fn clear_page_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::PageBreak(_)));
    }
}
