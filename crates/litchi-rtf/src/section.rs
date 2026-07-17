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

/// Section properties
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionProperties {
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
    /// Vertical alignment
    pub vertical_alignment: VerticalAlignment,
    /// Line numbering enabled
    pub line_numbering: bool,
    /// Line number restart on each page
    pub line_number_restart: bool,
    /// Explicit section-level footnote and endnote overrides.
    pub note_options: SectionNoteOptions,
    /// Page-border edges and placement for this section.
    pub page_borders: crate::PageBorders,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
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
            vertical_alignment: VerticalAlignment::default(),
            line_numbering: false,
            line_number_restart: false,
            note_options: SectionNoteOptions::default(),
            page_borders: crate::PageBorders::default(),
        }
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
}

impl<'a> HeaderFooter<'a> {
    /// Create a new header/footer
    #[inline]
    pub fn new(header_type: HeaderFooterType) -> Self {
        Self {
            header_type,
            paragraphs: Vec::new(),
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
    /// Whether this is a footnote (true) or endnote (false)
    pub is_footnote: bool,
    /// Reference mark (number or symbol)
    pub reference: Cow<'a, str>,
    /// Note content
    pub content: Cow<'a, str>,
    /// Character formatting for the note
    pub formatting: Formatting,
}

impl<'a> Note<'a> {
    /// Create a new footnote
    #[inline]
    pub fn footnote(reference: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            is_footnote: true,
            reference,
            content,
            formatting: Formatting::default(),
        }
    }

    /// Create a new endnote
    #[inline]
    pub fn endnote(reference: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            is_footnote: false,
            reference,
            content,
            formatting: Formatting::default(),
        }
    }
}
