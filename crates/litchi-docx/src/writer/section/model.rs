#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::cast_possible_truncation,
    reason = "OOXML numeric values are bounded before conversion"
)]
#![expect(
    clippy::cast_sign_loss,
    reason = "the value is validated as nonnegative before conversion"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
use crate::error::{Error, Result};
use crate::header_footer::Kind;
use crate::section::Start;

use super::borders;
use super::validation;

/// Page number format used by page, footnote, and endnote numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageNumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
}

impl PageNumberFormat {
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "decimal" => Ok(Self::Decimal),
            "upperRoman" => Ok(Self::UpperRoman),
            "lowerRoman" => Ok(Self::LowerRoman),
            "upperLetter" => Ok(Self::UpperLetter),
            "lowerLetter" => Ok(Self::LowerLetter),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported section numbering format '{value}'"
            ))),
        }
    }
}

/// Page orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

impl PageOrientation {
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(Error::InvalidFormat(format!(
                "invalid section page orientation '{value}'"
            ))),
        }
    }
}

/// One header or footer relationship used by a section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionHeaderFooterReference {
    pub kind: Kind,
    /// Existing relationship ID. `None` binds a default reference to content
    /// created through [`crate::writer::MutableDocument::header`] or `footer`.
    pub relationship_id: Option<String>,
    /// New or replacement part owned by this section. References with the same
    /// non-empty key are deliberately shared and must have identical XML.
    pub part: Option<SectionHeaderFooterPart>,
}

/// Header/footer XML staged as a package part during save.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionHeaderFooterPart {
    pub key: String,
    pub xml: String,
}

#[derive(Debug, Clone)]
pub(super) enum AuthoredChild {
    Header(Kind),
    Footer(Kind),
    FootnotePr,
    EndnotePr,
    Type,
    PageSize,
    PageMargins,
    PaperSource,
    PageBorders,
    LineNumbering,
    PageNumbering,
    Columns,
    FormProtection,
    VerticalAlignment,
    NoEndnote(String),
    TitlePage,
    TextDirection,
    Bidirectional,
    RtlGutter,
    DocumentGrid,
    PrinterSettings,
    SectionChange(String),
    Unknown(String),
}

#[derive(Debug, Clone)]
pub(super) struct NamespaceBinding {
    pub(super) prefix: Option<String>,
    pub(super) uri: String,
}

#[derive(Debug, Clone, Default)]
pub(super) struct AuthoredChildRaw {
    pub(super) raw: String,
    pub(super) preserved_attributes: Vec<String>,
    pub(super) original_on_off: Option<bool>,
}

impl SectionHeaderFooterReference {
    pub fn existing(kind: Kind, relationship_id: impl Into<String>) -> Self {
        Self {
            kind,
            relationship_id: Some(relationship_id.into()),
            part: None,
        }
    }

    pub fn owned(kind: Kind, key: impl Into<String>, xml: impl Into<String>) -> Self {
        Self {
            kind,
            relationship_id: None,
            part: Some(SectionHeaderFooterPart {
                key: key.into(),
                xml: xml.into(),
            }),
        }
    }

    #[must_use]
    pub fn managed_default(kind: Kind) -> Self {
        Self {
            kind,
            relationship_id: None,
            part: None,
        }
    }
}

/// Page-numbering settings for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionPageNumbering {
    pub format: PageNumberFormat,
    pub start: Option<u32>,
    pub chapter_style: Option<u8>,
    pub chapter_separator: Option<ChapterSep>,
}

/// Separator rendered between chapter and page numbers (`ST_ChapterSep`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChapterSep {
    Hyphen,
    Period,
    Colon,
    EmDash,
    EnDash,
}

impl ChapterSep {
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Hyphen => "hyphen",
            Self::Period => "period",
            Self::Colon => "colon",
            Self::EmDash => "emDash",
            Self::EnDash => "enDash",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "hyphen" => Ok(Self::Hyphen),
            "period" => Ok(Self::Period),
            "colon" => Ok(Self::Colon),
            "emDash" => Ok(Self::EmDash),
            "enDash" => Ok(Self::EnDash),
            _ => Err(Error::InvalidFormat(format!(
                "invalid chapter separator '{value}'"
            ))),
        }
    }
}

/// One explicitly sized newspaper-style column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionColumn {
    pub width: u32,
    pub space: Option<u32>,
}

/// Section column layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionColumns {
    pub equal_width: bool,
    pub count: u16,
    pub space: Option<u32>,
    pub separator: bool,
    pub columns: Vec<SectionColumn>,
}

impl Default for SectionColumns {
    fn default() -> Self {
        Self {
            equal_width: true,
            count: 1,
            space: None,
            separator: false,
            columns: Vec::new(),
        }
    }
}

/// Restart policy for note numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteNumberRestart {
    Continuous,
    EachSection,
    EachPage,
}

impl NoteNumberRestart {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::EachSection => "eachSect",
            Self::EachPage => "eachPage",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "continuous" => Ok(Self::Continuous),
            "eachSect" => Ok(Self::EachSection),
            "eachPage" => Ok(Self::EachPage),
            _ => Err(Error::InvalidFormat(format!(
                "invalid note-number restart '{value}'"
            ))),
        }
    }
}

/// Footnote positioning location (`ST_FtnPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FootnotePos {
    PageBottom,
    BeneathText,
    SectionEnd,
    DocumentEnd,
}

/// Endnote positioning location (`ST_EdnPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EndnotePos {
    SectionEnd,
    DocumentEnd,
}

pub(super) trait NotePos: Copy {
    fn parse(value: &str) -> Result<Self>;
    fn as_str(self) -> &'static str;
}

impl NotePos for FootnotePos {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "pageBottom" => Ok(Self::PageBottom),
            "beneathText" => Ok(Self::BeneathText),
            "sectEnd" => Ok(Self::SectionEnd),
            "docEnd" => Ok(Self::DocumentEnd),
            _ => Err(Error::InvalidFormat(format!(
                "invalid footnote position '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::PageBottom => "pageBottom",
            Self::BeneathText => "beneathText",
            Self::SectionEnd => "sectEnd",
            Self::DocumentEnd => "docEnd",
        }
    }
}

impl NotePos for EndnotePos {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "sectEnd" => Ok(Self::SectionEnd),
            "docEnd" => Ok(Self::DocumentEnd),
            _ => Err(Error::InvalidFormat(format!(
                "invalid endnote position '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::SectionEnd => "sectEnd",
            Self::DocumentEnd => "docEnd",
        }
    }
}

/// Footnote numbering and placement within a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Footnotes {
    pub format: PageNumberFormat,
    pub start: Option<u32>,
    pub restart: Option<NoteNumberRestart>,
    pub position: Option<FootnotePos>,
}

impl Default for Footnotes {
    fn default() -> Self {
        Self {
            format: PageNumberFormat::Decimal,
            start: None,
            restart: None,
            position: None,
        }
    }
}

/// Endnote numbering and placement within a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Endnotes {
    pub format: PageNumberFormat,
    pub start: Option<u32>,
    pub restart: Option<NoteNumberRestart>,
    pub position: Option<EndnotePos>,
}

impl Default for Endnotes {
    fn default() -> Self {
        Self {
            format: PageNumberFormat::Decimal,
            start: None,
            restart: None,
            position: None,
        }
    }
}

/// Text flow for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionTextDirection {
    LeftToRightTopToBottom,
    TopToBottomRightToLeft,
    BottomToTopLeftToRight,
    LeftToRightTopToBottomRotated,
    TopToBottomRightToLeftRotated,
    TopToBottomLeftToRightRotated,
}

impl SectionTextDirection {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::LeftToRightTopToBottom => "lrTb",
            Self::TopToBottomRightToLeft => "tbRl",
            Self::BottomToTopLeftToRight => "btLr",
            Self::LeftToRightTopToBottomRotated => "lrTbV",
            Self::TopToBottomRightToLeftRotated => "tbRlV",
            Self::TopToBottomLeftToRightRotated => "tbLrV",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "lrTb" => Ok(Self::LeftToRightTopToBottom),
            "tbRl" => Ok(Self::TopToBottomRightToLeft),
            "btLr" => Ok(Self::BottomToTopLeftToRight),
            "lrTbV" => Ok(Self::LeftToRightTopToBottomRotated),
            "tbRlV" => Ok(Self::TopToBottomRightToLeftRotated),
            "tbLrV" => Ok(Self::TopToBottomLeftToRightRotated),
            _ => Err(Error::InvalidFormat(format!(
                "invalid section text direction '{value}'"
            ))),
        }
    }
}

/// Document-grid mode for East Asian layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GridType {
    Default,
    Lines,
    LinesAndChars,
    SnapToChars,
}

impl GridType {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Lines => "lines",
            Self::LinesAndChars => "linesAndChars",
            Self::SnapToChars => "snapToChars",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "default" => Ok(Self::Default),
            "lines" => Ok(Self::Lines),
            "linesAndChars" => Ok(Self::LinesAndChars),
            "snapToChars" => Ok(Self::SnapToChars),
            _ => Err(Error::InvalidFormat(format!(
                "invalid document grid type '{value}'"
            ))),
        }
    }
}

/// Section document-grid settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionDocumentGrid {
    pub grid_type: GridType,
    pub line_pitch: Option<u32>,
    pub char_space: Option<i32>,
}

/// Maximum page-border size in eighths of a point for line-style borders
/// (`ST_EighthPointMeasure`, ECMA-376 §17.18.11).
pub(super) const MAX_PAGE_BORDER_LINE_SIZE: u32 = 96;
/// Maximum page-border size in points for art borders
/// (`ST_PointMeasure`, ECMA-376 §17.18.68).
pub(super) const MAX_PAGE_BORDER_ART_SIZE: u32 = 1638;
/// Maximum page-border spacing from text or page edge in points
/// (`w:space` on `CT_Border` is limited to 31, ECMA-376 §17.6.16).
pub(super) const MAX_PAGE_BORDER_SPACE: u32 = 31;

macro_rules! define_page_border_art {
    ($($variant:ident => $token:literal),+ $(,)?) => {
        /// Fixed page-border artwork domain from `ST_Border`.
        ///
        /// The 164 variants follow Word's complete `0x40..=0xE3` page-art
        /// range in `[MS-DOC]` section 2.9.22. The schema's separate `custom`
        /// sentinel is intentionally excluded because Word treats it as a
        /// corrupt document (`[MS-OI29500]` section 2.1.528(d)).
        #[repr(u8)]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        pub enum Art {
            $($variant),+
        }

        impl Art {
            /// Every page-border artwork value in schema order.
            pub const ALL: &'static [Self] = &[$(Self::$variant),+];

            /// Return the exact `ST_Border` XML token.
            pub const fn token(self) -> &'static str {
                match self {
                    $(Self::$variant => $token),+
                }
            }

            /// Return Word's `BrcType` page-art code (`0x40..=0xE3`).
            pub const fn code(self) -> u8 {
                self as u8 + 0x40
            }
        }

        impl std::str::FromStr for Art {
            type Err = Error;

            fn from_str(value: &str) -> Result<Self> {
                match value {
                    $($token => Ok(Self::$variant)),+,
                    _ => Err(Error::InvalidFormat(format!(
                        "invalid page border artwork '{value}'"
                    ))),
                }
            }
        }

        impl std::fmt::Display for Art {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter.write_str(self.token())
            }
        }

        impl TryFrom<u8> for Art {
            type Error = Error;

            fn try_from(value: u8) -> Result<Self> {
                value
                    .checked_sub(0x40)
                    .and_then(|index| Self::ALL.get(usize::from(index)))
                    .copied()
                    .ok_or_else(|| Error::InvalidFormat(format!(
                        "invalid page border artwork code {value:#04X}"
                    )))
            }
        }
    };
}

define_page_border_art! {
    Apples => "apples",
    ArchedScallops => "archedScallops",
    BabyPacifier => "babyPacifier",
    BabyRattle => "babyRattle",
    Balloons3Colors => "balloons3Colors",
    BalloonsHotAir => "balloonsHotAir",
    BasicBlackDashes => "basicBlackDashes",
    BasicBlackDots => "basicBlackDots",
    BasicBlackSquares => "basicBlackSquares",
    BasicThinLines => "basicThinLines",
    BasicWhiteDashes => "basicWhiteDashes",
    BasicWhiteDots => "basicWhiteDots",
    BasicWhiteSquares => "basicWhiteSquares",
    BasicWideInline => "basicWideInline",
    BasicWideMidline => "basicWideMidline",
    BasicWideOutline => "basicWideOutline",
    Bats => "bats",
    Birds => "birds",
    BirdsFlight => "birdsFlight",
    Cabins => "cabins",
    CakeSlice => "cakeSlice",
    CandyCorn => "candyCorn",
    CelticKnotwork => "celticKnotwork",
    CertificateBanner => "certificateBanner",
    ChainLink => "chainLink",
    ChampagneBottle => "champagneBottle",
    CheckedBarBlack => "checkedBarBlack",
    CheckedBarColor => "checkedBarColor",
    Checkered => "checkered",
    ChristmasTree => "christmasTree",
    CirclesLines => "circlesLines",
    CirclesRectangles => "circlesRectangles",
    ClassicalWave => "classicalWave",
    Clocks => "clocks",
    Compass => "compass",
    Confetti => "confetti",
    ConfettiGrays => "confettiGrays",
    ConfettiOutline => "confettiOutline",
    ConfettiStreamers => "confettiStreamers",
    ConfettiWhite => "confettiWhite",
    CornerTriangles => "cornerTriangles",
    CouponCutoutDashes => "couponCutoutDashes",
    CouponCutoutDots => "couponCutoutDots",
    CrazyMaze => "crazyMaze",
    CreaturesButterfly => "creaturesButterfly",
    CreaturesFish => "creaturesFish",
    CreaturesInsects => "creaturesInsects",
    CreaturesLadyBug => "creaturesLadyBug",
    CrossStitch => "crossStitch",
    Cup => "cup",
    DecoArch => "decoArch",
    DecoArchColor => "decoArchColor",
    DecoBlocks => "decoBlocks",
    DiamondsGray => "diamondsGray",
    DoubleD => "doubleD",
    DoubleDiamonds => "doubleDiamonds",
    Earth1 => "earth1",
    Earth2 => "earth2",
    EclipsingSquares1 => "eclipsingSquares1",
    EclipsingSquares2 => "eclipsingSquares2",
    EggsBlack => "eggsBlack",
    Fans => "fans",
    Film => "film",
    Firecrackers => "firecrackers",
    FlowersBlockPrint => "flowersBlockPrint",
    FlowersDaisies => "flowersDaisies",
    FlowersModern1 => "flowersModern1",
    FlowersModern2 => "flowersModern2",
    FlowersPansy => "flowersPansy",
    FlowersRedRose => "flowersRedRose",
    FlowersRoses => "flowersRoses",
    FlowersTeacup => "flowersTeacup",
    FlowersTiny => "flowersTiny",
    Gems => "gems",
    GingerbreadMan => "gingerbreadMan",
    Gradient => "gradient",
    Handmade1 => "handmade1",
    Handmade2 => "handmade2",
    HeartBalloon => "heartBalloon",
    HeartGray => "heartGray",
    Hearts => "hearts",
    HeebieJeebies => "heebieJeebies",
    Holly => "holly",
    HouseFunky => "houseFunky",
    Hypnotic => "hypnotic",
    IceCreamCones => "iceCreamCones",
    LightBulb => "lightBulb",
    Lightning1 => "lightning1",
    Lightning2 => "lightning2",
    MapPins => "mapPins",
    MapleLeaf => "mapleLeaf",
    MapleMuffins => "mapleMuffins",
    Marquee => "marquee",
    MarqueeToothed => "marqueeToothed",
    Moons => "moons",
    Mosaic => "mosaic",
    MusicNotes => "musicNotes",
    Northwest => "northwest",
    Ovals => "ovals",
    Packages => "packages",
    PalmsBlack => "palmsBlack",
    PalmsColor => "palmsColor",
    PaperClips => "paperClips",
    Papyrus => "papyrus",
    PartyFavor => "partyFavor",
    PartyGlass => "partyGlass",
    Pencils => "pencils",
    People => "people",
    PeopleWaving => "peopleWaving",
    PeopleHats => "peopleHats",
    Poinsettias => "poinsettias",
    PostageStamp => "postageStamp",
    Pumpkin1 => "pumpkin1",
    PushPinNote2 => "pushPinNote2",
    PushPinNote1 => "pushPinNote1",
    Pyramids => "pyramids",
    PyramidsAbove => "pyramidsAbove",
    Quadrants => "quadrants",
    Rings => "rings",
    Safari => "safari",
    Sawtooth => "sawtooth",
    SawtoothGray => "sawtoothGray",
    ScaredCat => "scaredCat",
    Seattle => "seattle",
    ShadowedSquares => "shadowedSquares",
    SharksTeeth => "sharksTeeth",
    ShorebirdTracks => "shorebirdTracks",
    Skyrocket => "skyrocket",
    SnowflakeFancy => "snowflakeFancy",
    Snowflakes => "snowflakes",
    Sombrero => "sombrero",
    Southwest => "southwest",
    Stars => "stars",
    StarsTop => "starsTop",
    Stars3d => "stars3d",
    StarsBlack => "starsBlack",
    StarsShadowed => "starsShadowed",
    Sun => "sun",
    Swirligig => "swirligig",
    TornPaper => "tornPaper",
    TornPaperBlack => "tornPaperBlack",
    Trees => "trees",
    TriangleParty => "triangleParty",
    Triangles => "triangles",
    Tribal1 => "tribal1",
    Tribal2 => "tribal2",
    Tribal3 => "tribal3",
    Tribal4 => "tribal4",
    Tribal5 => "tribal5",
    Tribal6 => "tribal6",
    TwistedLines1 => "twistedLines1",
    TwistedLines2 => "twistedLines2",
    Vine => "vine",
    Waveline => "waveline",
    WeavingAngles => "weavingAngles",
    WeavingBraid => "weavingBraid",
    WeavingRibbon => "weavingRibbon",
    WeavingStrips => "weavingStrips",
    WhiteFlowers => "whiteFlowers",
    Woodwork => "woodwork",
    XIllusions => "xIllusions",
    ZanyTriangles => "zanyTriangles",
    ZigZag => "zigZag",
    ZigZagStitch => "zigZagStitch",
}

/// Page-border style (`ST_Border`, ECMA-376 §17.18.2).
///
/// Line styles map directly to fixed tokens; artwork is closed over
/// [`Art`] and cannot carry arbitrary strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Style {
    Nil,
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThinThickMediumGap,
    ThinThickLargeGap,
    ThickThinSmallGap,
    ThickThinMediumGap,
    ThickThinLargeGap,
    ThinThickThinSmallGap,
    ThinThickThinMediumGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
    Art(Art),
}

impl Style {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Nil => "nil",
            Self::None => "none",
            Self::Single => "single",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::DotDash => "dotDash",
            Self::DotDotDash => "dotDotDash",
            Self::Triple => "triple",
            Self::ThinThickSmallGap => "thinThickSmallGap",
            Self::ThinThickMediumGap => "thinThickMediumGap",
            Self::ThinThickLargeGap => "thinThickLargeGap",
            Self::ThickThinSmallGap => "thickThinSmallGap",
            Self::ThickThinMediumGap => "thickThinMediumGap",
            Self::ThickThinLargeGap => "thickThinLargeGap",
            Self::ThinThickThinSmallGap => "thinThickThinSmallGap",
            Self::ThinThickThinMediumGap => "thinThickThinMediumGap",
            Self::ThinThickThinLargeGap => "thinThickThinLargeGap",
            Self::Wave => "wave",
            Self::DoubleWave => "doubleWave",
            Self::DashSmallGap => "dashSmallGap",
            Self::DashDotStroked => "dashDotStroked",
            Self::ThreeDEmboss => "threeDEmboss",
            Self::ThreeDEngrave => "threeDEngrave",
            Self::Outset => "outset",
            Self::Inset => "inset",
            Self::Art(art) => art.token(),
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "nil" => Self::Nil,
            "none" => Self::None,
            "single" => Self::Single,
            "thick" => Self::Thick,
            "double" => Self::Double,
            "dotted" => Self::Dotted,
            "dashed" => Self::Dashed,
            "dotDash" => Self::DotDash,
            "dotDotDash" => Self::DotDotDash,
            "triple" => Self::Triple,
            "thinThickSmallGap" => Self::ThinThickSmallGap,
            "thinThickMediumGap" => Self::ThinThickMediumGap,
            "thinThickLargeGap" => Self::ThinThickLargeGap,
            "thickThinSmallGap" => Self::ThickThinSmallGap,
            "thickThinMediumGap" => Self::ThickThinMediumGap,
            "thickThinLargeGap" => Self::ThickThinLargeGap,
            "thinThickThinSmallGap" => Self::ThinThickThinSmallGap,
            "thinThickThinMediumGap" => Self::ThinThickThinMediumGap,
            "thinThickThinLargeGap" => Self::ThinThickThinLargeGap,
            "wave" => Self::Wave,
            "doubleWave" => Self::DoubleWave,
            "dashSmallGap" => Self::DashSmallGap,
            "dashDotStroked" => Self::DashDotStroked,
            "threeDEmboss" => Self::ThreeDEmboss,
            "threeDEngrave" => Self::ThreeDEngrave,
            "outset" => Self::Outset,
            "inset" => Self::Inset,
            art => Self::Art(art.parse()?),
        })
    }
}

impl From<Art> for Style {
    fn from(value: Art) -> Self {
        Self::Art(value)
    }
}

/// Which pages display the section page border (`w:display`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Display {
    AllPages,
    FirstPage,
    NotFirstPage,
}

impl Display {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::AllPages => "allPages",
            Self::FirstPage => "firstPage",
            Self::NotFirstPage => "notFirstPage",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "allPages" => Ok(Self::AllPages),
            "firstPage" => Ok(Self::FirstPage),
            "notFirstPage" => Ok(Self::NotFirstPage),
            _ => Err(Error::InvalidFormat(format!(
                "invalid page border display '{value}'"
            ))),
        }
    }
}

/// Page-border measurement origin (`w:offsetFrom`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OffsetFrom {
    Page,
    Text,
}

impl OffsetFrom {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Text => "text",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "text" => Ok(Self::Text),
            _ => Err(Error::InvalidFormat(format!(
                "invalid page border offset '{value}'"
            ))),
        }
    }
}

/// Whether the page border renders in front of or behind text (`w:zOrder`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ZOrder {
    Front,
    Back,
}

impl ZOrder {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Front => "front",
            Self::Back => "back",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "front" => Ok(Self::Front),
            "back" => Ok(Self::Back),
            _ => Err(Error::InvalidFormat(format!(
                "invalid page border z-order '{value}'"
            ))),
        }
    }
}

/// Page-border color (`ST_HexColor`), represented without heap allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Color {
    Auto,
    Rgb([u8; 3]),
}

impl Color {
    /// Creates an explicit red-green-blue color.
    #[must_use]
    pub const fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self::Rgb([red, green, blue])
    }

    /// Returns the RGB components, or `None` for automatic color selection.
    #[must_use]
    pub const fn components(self) -> Option<[u8; 3]> {
        match self {
            Self::Auto => None,
            Self::Rgb(components) => Some(components),
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        if value == "auto" {
            return Ok(Self::Auto);
        }
        if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(Error::InvalidFormat(format!(
                "invalid page border color '{value}'"
            )));
        }
        let component = |range| {
            u8::from_str_radix(&value[range], 16).map_err(|_source_error| {
                Error::InvalidFormat(format!("invalid page border color '{value}'"))
            })
        };
        Ok(Self::rgb(
            component(0..2)?,
            component(2..4)?,
            component(4..6)?,
        ))
    }
}

/// Paper-source tray codes (`CT_PaperSource`, ECMA-376 §17.6.19).
///
/// Values are printer-driver tray codes; `first` applies to the first page,
/// `other` to the remaining pages of the section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SectionPaperSource {
    pub first: Option<u32>,
    pub other: Option<u32>,
}

/// Line-number restart policy (`ST_LineNumberRestart`, ECMA-376 §17.18.49).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineNumberRestart {
    NewPage,
    NewSection,
    Continuous,
}

impl LineNumberRestart {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::NewPage => "newPage",
            Self::NewSection => "newSection",
            Self::Continuous => "continuous",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "newPage" => Ok(Self::NewPage),
            "newSection" => Ok(Self::NewSection),
            "continuous" => Ok(Self::Continuous),
            _ => Err(Error::InvalidFormat(format!(
                "invalid line-number restart '{value}'"
            ))),
        }
    }
}

/// Section line-numbering settings (`CT_LineNumber`, ECMA-376 §17.6.12).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SectionLineNumbering {
    /// Line-number increment; lines are numbered when divisible by this value.
    pub count_by: Option<u32>,
    /// First line number used by the numbering scheme.
    pub start: Option<u32>,
    /// Distance between line numbers and text in twips.
    pub distance: Option<u32>,
    pub restart: Option<LineNumberRestart>,
}

/// Vertical content alignment (`ST_VerticalJc`, ECMA-376 §17.18.100).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionVerticalAlignment {
    Top,
    Center,
    Justified,
    Bottom,
}

impl SectionVerticalAlignment {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Justified => "both",
            Self::Bottom => "bottom",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "both" => Ok(Self::Justified),
            "bottom" => Ok(Self::Bottom),
            _ => Err(Error::InvalidFormat(format!(
                "invalid section vertical alignment '{value}'"
            ))),
        }
    }
}

/// Typed `WordprocessingML` `sectPr` properties.
#[derive(Debug, Clone)]
pub struct SectionProperties {
    pub page_width: u32,
    pub page_height: u32,
    pub orientation: PageOrientation,
    pub margin_top: u32,
    pub margin_bottom: u32,
    pub margin_left: u32,
    pub margin_right: u32,
    pub header_distance: u32,
    pub footer_distance: u32,
    pub gutter: u32,
    pub start_type: Option<Start>,
    pub headers: Vec<SectionHeaderFooterReference>,
    pub footers: Vec<SectionHeaderFooterReference>,
    pub page_numbering: Option<SectionPageNumbering>,
    pub columns: Option<SectionColumns>,
    pub footnotes: Option<Footnotes>,
    pub endnotes: Option<Endnotes>,
    pub text_direction: Option<SectionTextDirection>,
    pub document_grid: Option<SectionDocumentGrid>,
    pub form_protection: bool,
    pub paper_source: Option<SectionPaperSource>,
    pub page_borders: Option<borders::Borders>,
    pub line_numbering: Option<SectionLineNumbering>,
    pub vertical_alignment: Option<SectionVerticalAlignment>,
    /// Different first-page header/footer (`w:titlePg`).
    pub title_page: bool,
    /// Right-to-left section layout (`w:bidi`).
    pub bidirectional: bool,
    /// Gutter positioned on the right side (`w:rtlGutter`).
    pub rtl_gutter: bool,
    /// Relationship ID of the printer-settings part (`w:printerSettings`).
    pub printer_settings_relationship_id: Option<String>,
    pub(super) preserved_unknown_children: Vec<String>,
    pub(super) authored_children: Vec<AuthoredChild>,
    pub(super) authored_child_raw: Vec<AuthoredChildRaw>,
    pub(super) namespace_bindings: Vec<NamespaceBinding>,
    pub(super) root_prefix: Option<String>,
    pub(super) root_default_namespace: bool,
    pub(super) root_namespace_resolved: bool,
    pub(super) root_attributes: Vec<String>,
    pub(super) suppress_managed_headers: bool,
    pub(super) suppress_managed_footers: bool,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            page_width: 12240,
            page_height: 15840,
            orientation: PageOrientation::Portrait,
            margin_top: 1440,
            margin_bottom: 1440,
            margin_left: 1440,
            margin_right: 1440,
            header_distance: 720,
            footer_distance: 720,
            gutter: 0,
            start_type: None,
            headers: Vec::new(),
            footers: Vec::new(),
            page_numbering: None,
            columns: None,
            footnotes: None,
            endnotes: None,
            text_direction: None,
            document_grid: None,
            form_protection: false,
            paper_source: None,
            page_borders: None,
            line_numbering: None,
            vertical_alignment: None,
            title_page: false,
            bidirectional: false,
            rtl_gutter: false,
            printer_settings_relationship_id: None,
            preserved_unknown_children: Vec::new(),
            authored_children: Vec::new(),
            authored_child_raw: Vec::new(),
            namespace_bindings: Vec::new(),
            root_prefix: None,
            root_default_namespace: false,
            root_namespace_resolved: true,
            root_attributes: Vec::new(),
            suppress_managed_headers: false,
            suppress_managed_footers: false,
        }
    }
}

impl SectionProperties {
    #[must_use]
    pub fn a4() -> Self {
        Self {
            page_width: 11906,
            page_height: 16838,
            ..Self::default()
        }
    }

    #[must_use]
    pub fn letter() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn legal() -> Self {
        Self {
            page_height: 20160,
            ..Self::default()
        }
    }

    #[must_use]
    pub fn landscape(mut self) -> Self {
        self.orientation = PageOrientation::Landscape;
        std::mem::swap(&mut self.page_width, &mut self.page_height);
        self
    }

    #[must_use]
    pub fn margins(mut self, top: f64, bottom: f64, left: f64, right: f64) -> Self {
        self.margin_top = inches_to_twips(top);
        self.margin_bottom = inches_to_twips(bottom);
        self.margin_left = inches_to_twips(left);
        self.margin_right = inches_to_twips(right);
        self
    }

    #[must_use]
    pub fn with_start_type(mut self, start_type: Start) -> Self {
        self.start_type = Some(start_type);
        self
    }

    /// Create or replace a section-owned header part of the selected kind.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_header_part(
        &mut self,
        kind: Kind,
        key: impl Into<String>,
        xml: impl Into<String>,
    ) -> Result<()> {
        self.set_header_footer_part(true, kind, key.into(), xml.into())
    }

    /// Create or replace a section-owned footer part of the selected kind.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_footer_part(
        &mut self,
        kind: Kind,
        key: impl Into<String>,
        xml: impl Into<String>,
    ) -> Result<()> {
        self.set_header_footer_part(false, kind, key.into(), xml.into())
    }

    /// Remove a header reference and any section-owned replacement of that kind.
    pub fn remove_header(&mut self, kind: Kind) -> bool {
        self.suppress_managed_headers = true;
        remove_reference(&mut self.headers, kind)
    }

    /// Remove a footer reference and any section-owned replacement of that kind.
    pub fn remove_footer(&mut self, kind: Kind) -> bool {
        self.suppress_managed_footers = true;
        remove_reference(&mut self.footers, kind)
    }

    fn set_header_footer_part(
        &mut self,
        header: bool,
        kind: Kind,
        key: String,
        xml: String,
    ) -> Result<()> {
        validation::validate_header_footer_xml(&xml, header)?;
        if key.is_empty() {
            return Err(Error::InvalidFormat(
                "section header/footer part key is empty".to_string(),
            ));
        }
        let references = if header {
            self.suppress_managed_headers = false;
            &mut self.headers
        } else {
            self.suppress_managed_footers = false;
            &mut self.footers
        };
        remove_reference(references, kind);
        references.push(SectionHeaderFooterReference::owned(kind, key, xml));
        Ok(())
    }
}

fn remove_reference(references: &mut Vec<SectionHeaderFooterReference>, kind: Kind) -> bool {
    let length = references.len();
    references.retain(|reference| reference.kind != kind);
    references.len() != length
}

fn inches_to_twips(inches: f64) -> u32 {
    (inches * 1440.0).round().max(0.0) as u32
}
