//! Public semantic value types used by the Pages editor.

use std::num::NonZeroU32;

use crate::protobuf::tp::DocumentArchive;
use crate::text::TextStorageInfo;
use crate::{Error, Result};

const RAW_SECTION_START_NEXT_PAGE: u32 = 0;
const RAW_SECTION_START_RIGHT_PAGE: u32 = 1;
const RAW_SECTION_START_LEFT_PAGE: u32 = 2;
const RAW_PAGE_NUMBERING_CONTINUE: u32 = 0;
const RAW_PAGE_NUMBERING_RESTART: u32 = 1;
const RAW_PAGE_ORIENTATION_PORTRAIT: u32 = 0;
const RAW_PAGE_ORIENTATION_LANDSCAPE: u32 = 1;
const RAW_FOOTNOTE_KIND_FOOTNOTES: i32 = 0;
const RAW_FOOTNOTE_KIND_DOCUMENT_ENDNOTES: i32 = 1;
const RAW_FOOTNOTE_KIND_SECTION_ENDNOTES: i32 = 2;
const RAW_FOOTNOTE_FORMAT_NUMERIC: i32 = 0;
const RAW_FOOTNOTE_FORMAT_ROMAN: i32 = 1;
const RAW_FOOTNOTE_FORMAT_SYMBOLIC: i32 = 2;
const RAW_FOOTNOTE_FORMAT_JAPANESE_NUMERIC: i32 = 3;
const RAW_FOOTNOTE_FORMAT_JAPANESE_IDEOGRAPHIC: i32 = 4;
const RAW_FOOTNOTE_FORMAT_ARABIC_NUMERIC: i32 = 5;
const RAW_FOOTNOTE_NUMBERING_CONTINUOUS: i32 = 0;
const RAW_FOOTNOTE_NUMBERING_RESTART_EACH_PAGE: i32 = 1;
const RAW_FOOTNOTE_NUMBERING_RESTART_EACH_SECTION: i32 = 2;

/// Which page variant owns a Pages header/footer storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesTemplateKind {
    First,
    Even,
    Odd,
}

/// Whether a reachable Pages text region is a header or a footer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesHeaderFooterKind {
    Header,
    Footer,
}

/// A reachable header/footer slot and its current writable text storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesHeaderFooterInfo {
    pub section_id: u64,
    pub section_name: Option<String>,
    /// UTF-16 position where the section begins in the body storage.
    pub section_character_index: u32,
    pub template_id: u64,
    pub template: PagesTemplateKind,
    pub kind: PagesHeaderFooterKind,
    /// Archive order within the header/footer list, normally left/center/right.
    pub slot: usize,
    pub storage: TextStorageInfo,
}

/// A writable text storage owned by a drawable reachable from a Pages document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesDrawableTextInfo {
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// Result of removing a body-anchored Pages text box.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedPagesTextBox {
    pub text: PagesDrawableTextInfo,
    /// UTF-16 body position formerly occupied by the object-replacement character.
    pub anchor_character_index: u32,
}

/// A section boundary reachable from the main Pages body storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesSectionInfo {
    pub object_id: u64,
    /// UTF-16 position where the section begins in the body storage.
    pub character_index: u32,
    pub name: Option<String>,
    pub first_template_id: Option<u64>,
    pub even_template_id: Option<u64>,
    pub odd_template_id: Option<u64>,
}

/// Page on which a Pages section begins when facing pages are enabled.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesSectionStart {
    NextPage,
    RightPage,
    LeftPage,
    /// A value written by a newer Pages version.
    Unknown(u32),
}

impl PagesSectionStart {
    /// Decode the lossless protobuf representation.
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            RAW_SECTION_START_NEXT_PAGE => Self::NextPage,
            RAW_SECTION_START_RIGHT_PAGE => Self::RightPage,
            RAW_SECTION_START_LEFT_PAGE => Self::LeftPage,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the lossless protobuf representation.
    pub const fn as_raw(self) -> u32 {
        match self {
            Self::NextPage => RAW_SECTION_START_NEXT_PAGE,
            Self::RightPage => RAW_SECTION_START_RIGHT_PAGE,
            Self::LeftPage => RAW_SECTION_START_LEFT_PAGE,
            Self::Unknown(raw) => raw,
        }
    }

    pub(super) const fn is_canonical(self) -> bool {
        match self {
            Self::Unknown(raw) => matches!(Self::from_raw(raw), Self::Unknown(_)),
            _ => true,
        }
    }
}

/// Whether a Pages section continues or restarts page numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesSectionPageNumbering {
    ContinueFromPrevious,
    Restart,
    /// A value written by a newer Pages version.
    Unknown(u32),
}

impl PagesSectionPageNumbering {
    /// Decode the lossless protobuf representation.
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            RAW_PAGE_NUMBERING_CONTINUE => Self::ContinueFromPrevious,
            RAW_PAGE_NUMBERING_RESTART => Self::Restart,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the lossless protobuf representation.
    pub const fn as_raw(self) -> u32 {
        match self {
            Self::ContinueFromPrevious => RAW_PAGE_NUMBERING_CONTINUE,
            Self::Restart => RAW_PAGE_NUMBERING_RESTART,
            Self::Unknown(raw) => raw,
        }
    }

    pub(super) const fn is_canonical(self) -> bool {
        match self {
            Self::Unknown(raw) => matches!(Self::from_raw(raw), Self::Unknown(_)),
            _ => true,
        }
    }
}

/// A validated, non-zero Pages page number.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PagesPageNumber(NonZeroU32);

impl PagesPageNumber {
    /// Validate and construct a page number.
    pub fn new(value: u32) -> Result<Self> {
        NonZeroU32::new(value).map(Self).ok_or_else(|| {
            Error::ParseError("Pages page numbers must be greater than zero".to_owned())
        })
    }

    /// Return the numeric page number.
    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

impl TryFrom<u32> for PagesPageNumber {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<PagesPageNumber> for u32 {
    fn from(value: PagesPageNumber) -> Self {
        value.get()
    }
}

/// Where Pages places notes belonging to the document body.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesFootnoteKind {
    Footnotes,
    DocumentEndnotes,
    SectionEndnotes,
    /// A value written by a newer Pages version.
    Unknown(i32),
}

impl PagesFootnoteKind {
    /// Decode the lossless protobuf representation.
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            RAW_FOOTNOTE_KIND_FOOTNOTES => Self::Footnotes,
            RAW_FOOTNOTE_KIND_DOCUMENT_ENDNOTES => Self::DocumentEndnotes,
            RAW_FOOTNOTE_KIND_SECTION_ENDNOTES => Self::SectionEndnotes,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the lossless protobuf representation.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Footnotes => RAW_FOOTNOTE_KIND_FOOTNOTES,
            Self::DocumentEndnotes => RAW_FOOTNOTE_KIND_DOCUMENT_ENDNOTES,
            Self::SectionEndnotes => RAW_FOOTNOTE_KIND_SECTION_ENDNOTES,
            Self::Unknown(raw) => raw,
        }
    }

    pub(super) const fn is_canonical(self) -> bool {
        match self {
            Self::Unknown(raw) => matches!(Self::from_raw(raw), Self::Unknown(_)),
            _ => true,
        }
    }
}

/// Marker sequence used for Pages footnotes and endnotes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesFootnoteFormat {
    Numeric,
    Roman,
    Symbolic,
    JapaneseNumeric,
    JapaneseIdeographic,
    ArabicNumeric,
    /// A value written by a newer Pages version.
    Unknown(i32),
}

impl PagesFootnoteFormat {
    /// Decode the lossless protobuf representation.
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            RAW_FOOTNOTE_FORMAT_NUMERIC => Self::Numeric,
            RAW_FOOTNOTE_FORMAT_ROMAN => Self::Roman,
            RAW_FOOTNOTE_FORMAT_SYMBOLIC => Self::Symbolic,
            RAW_FOOTNOTE_FORMAT_JAPANESE_NUMERIC => Self::JapaneseNumeric,
            RAW_FOOTNOTE_FORMAT_JAPANESE_IDEOGRAPHIC => Self::JapaneseIdeographic,
            RAW_FOOTNOTE_FORMAT_ARABIC_NUMERIC => Self::ArabicNumeric,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the lossless protobuf representation.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Numeric => RAW_FOOTNOTE_FORMAT_NUMERIC,
            Self::Roman => RAW_FOOTNOTE_FORMAT_ROMAN,
            Self::Symbolic => RAW_FOOTNOTE_FORMAT_SYMBOLIC,
            Self::JapaneseNumeric => RAW_FOOTNOTE_FORMAT_JAPANESE_NUMERIC,
            Self::JapaneseIdeographic => RAW_FOOTNOTE_FORMAT_JAPANESE_IDEOGRAPHIC,
            Self::ArabicNumeric => RAW_FOOTNOTE_FORMAT_ARABIC_NUMERIC,
            Self::Unknown(raw) => raw,
        }
    }

    pub(super) const fn is_canonical(self) -> bool {
        match self {
            Self::Unknown(raw) => matches!(Self::from_raw(raw), Self::Unknown(_)),
            _ => true,
        }
    }
}

/// How Pages restarts footnote or endnote numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesFootnoteNumbering {
    Continuous,
    RestartEachPage,
    RestartEachSection,
    /// A value written by a newer Pages version.
    Unknown(i32),
}

impl PagesFootnoteNumbering {
    /// Decode the lossless protobuf representation.
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            RAW_FOOTNOTE_NUMBERING_CONTINUOUS => Self::Continuous,
            RAW_FOOTNOTE_NUMBERING_RESTART_EACH_PAGE => Self::RestartEachPage,
            RAW_FOOTNOTE_NUMBERING_RESTART_EACH_SECTION => Self::RestartEachSection,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the lossless protobuf representation.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Continuous => RAW_FOOTNOTE_NUMBERING_CONTINUOUS,
            Self::RestartEachPage => RAW_FOOTNOTE_NUMBERING_RESTART_EACH_PAGE,
            Self::RestartEachSection => RAW_FOOTNOTE_NUMBERING_RESTART_EACH_SECTION,
            Self::Unknown(raw) => raw,
        }
    }

    pub(super) const fn is_canonical(self) -> bool {
        match self {
            Self::Unknown(raw) => matches!(Self::from_raw(raw), Self::Unknown(_)),
            _ => true,
        }
    }
}

/// Validated spacing between Pages footnotes or endnotes, in whole points.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PagesFootnoteGap(u32);

impl PagesFootnoteGap {
    /// Validate and construct a note gap.
    pub fn new(points: u32) -> Result<Self> {
        i32::try_from(points).map(|_| Self(points)).map_err(|_| {
            Error::ParseError(
                "Pages footnote gap exceeds the native signed integer range".to_owned(),
            )
        })
    }

    /// Return the gap in whole points.
    pub const fn points(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for PagesFootnoteGap {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<PagesFootnoteGap> for u32 {
    fn from(value: PagesFootnoteGap) -> Self {
        value.points()
    }
}

/// Lossless settings shown by Pages' Footnotes formatter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PagesFootnoteSettings {
    pub kind: Option<PagesFootnoteKind>,
    pub format: Option<PagesFootnoteFormat>,
    pub numbering: Option<PagesFootnoteNumbering>,
    pub gap: Option<PagesFootnoteGap>,
}

/// Writable settings stored directly on a Pages section.
///
/// Unknown native discriminants remain lossless through their typed `Unknown`
/// variants. `background_fill_payload`, when present, is the exact encoded
/// `TSD.FillArchive` payload.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PagesSectionSettings {
    pub name: Option<String>,
    pub inherit_previous_header_footer: Option<bool>,
    pub first_page_different: Option<bool>,
    pub even_odd_pages_different: Option<bool>,
    pub start: Option<PagesSectionStart>,
    pub page_numbering: Option<PagesSectionPageNumbering>,
    pub starting_page_number: Option<PagesPageNumber>,
    pub first_page_hides_header_footer: Option<bool>,
    pub background_fill_payload: Option<Vec<u8>>,
}

/// RGB color space used by a semantic Pages section background.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PagesRgbColorSpace {
    Srgb,
    DisplayP3,
}

/// Normalized RGB color components in the inclusive `0.0..=1.0` range.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PagesRgbaColor {
    pub red: f32,
    pub green: f32,
    pub blue: f32,
    pub alpha: f32,
    pub color_space: PagesRgbColorSpace,
}

/// Semantic Pages section background.
///
/// Gradient, image, extension, and future fills are exposed as `Opaque` so
/// callers can round-trip them losslessly through the same API.
#[derive(Debug, Clone, PartialEq)]
pub enum PagesSectionBackground {
    None,
    Solid(PagesRgbaColor),
    Opaque(Vec<u8>),
}

/// Physical orientation of a Pages document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesPageOrientation {
    Portrait,
    Landscape,
    /// A value written by a newer Pages version.
    Unknown(u32),
}

impl PagesPageOrientation {
    /// Decode the lossless protobuf representation.
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            RAW_PAGE_ORIENTATION_PORTRAIT => Self::Portrait,
            RAW_PAGE_ORIENTATION_LANDSCAPE => Self::Landscape,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the lossless protobuf representation.
    pub const fn as_raw(self) -> u32 {
        match self {
            Self::Portrait => RAW_PAGE_ORIENTATION_PORTRAIT,
            Self::Landscape => RAW_PAGE_ORIENTATION_LANDSCAPE,
            Self::Unknown(raw) => raw,
        }
    }

    pub(super) const fn is_canonical(self) -> bool {
        match self {
            Self::Unknown(raw) => matches!(Self::from_raw(raw), Self::Unknown(_)),
            _ => true,
        }
    }
}

/// Writable page geometry stored on the Pages document root.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesPageLayout {
    pub page_width: Option<f32>,
    pub page_height: Option<f32>,
    pub left_margin: Option<f32>,
    pub right_margin: Option<f32>,
    pub top_margin: Option<f32>,
    pub bottom_margin: Option<f32>,
    pub header_margin: Option<f32>,
    pub footer_margin: Option<f32>,
    pub page_scale: Option<f32>,
    pub orientation: Option<PagesPageOrientation>,
    pub lays_out_body_vertically: Option<bool>,
}

impl From<&DocumentArchive> for PagesPageLayout {
    fn from(document: &DocumentArchive) -> Self {
        Self {
            page_width: document.page_width,
            page_height: document.page_height,
            left_margin: document.left_margin,
            right_margin: document.right_margin,
            top_margin: document.top_margin,
            bottom_margin: document.bottom_margin,
            header_margin: document.header_margin,
            footer_margin: document.footer_margin,
            page_scale: document.page_scale,
            orientation: document.orientation.map(PagesPageOrientation::from_raw),
            lays_out_body_vertically: document.lays_out_body_vertically,
        }
    }
}
