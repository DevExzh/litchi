//! Typed PresentationML models for slide masters, layouts, and placeholders.

use litchi_opc::packuri::PackURI;

/// ECMA-376 Part 1: slide master and slide layout IDs start at 2^31.
pub const MIN_MASTER_OR_LAYOUT_ID: u32 = 2_147_483_648;

// ============================================================================
// Typed enums
// ============================================================================

/// Slide layout type (`ST_SlideLayoutType`, ECMA-376 Part 1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SlideLayoutKind {
    Title,
    Text,
    TwoColumnText,
    Table,
    TextAndChart,
    ChartAndText,
    Diagram,
    Chart,
    TextAndClipArt,
    ClipArtAndText,
    TitleOnly,
    Blank,
    TextAndObject,
    ObjectAndText,
    ObjectOnly,
    Object,
    TextAndMedia,
    MediaAndText,
    ObjectOverText,
    TextOverObject,
    TextAndTwoObjects,
    TwoObjectsAndText,
    TwoObjectsOverText,
    FourObjects,
    VerticalText,
    ClipArtAndVerticalText,
    VerticalTitleAndText,
    VerticalTitleAndTextOverChart,
    TwoObjects,
    ObjectAndTwoObjects,
    TwoObjectsAndObject,
    Custom,
    SectionHeader,
    TwoTextAndTwoObjects,
    ObjectText,
    PictureWithText,
}

impl SlideLayoutKind {
    /// The spec token written to the `type` attribute of `p:sldLayout`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Text => "tx",
            Self::TwoColumnText => "twoColTx",
            Self::Table => "tbl",
            Self::TextAndChart => "txAndChart",
            Self::ChartAndText => "chartAndTx",
            Self::Diagram => "dgm",
            Self::Chart => "chart",
            Self::TextAndClipArt => "txAndClipArt",
            Self::ClipArtAndText => "clipArtAndTx",
            Self::TitleOnly => "titleOnly",
            Self::Blank => "blank",
            Self::TextAndObject => "txAndObj",
            Self::ObjectAndText => "objAndTx",
            Self::ObjectOnly => "objOnly",
            Self::Object => "obj",
            Self::TextAndMedia => "txAndMedia",
            Self::MediaAndText => "mediaAndTx",
            Self::ObjectOverText => "objOverTx",
            Self::TextOverObject => "txOverObj",
            Self::TextAndTwoObjects => "txAndTwoObj",
            Self::TwoObjectsAndText => "twoObjAndTx",
            Self::TwoObjectsOverText => "twoObjOverTx",
            Self::FourObjects => "fourObj",
            Self::VerticalText => "vertTx",
            Self::ClipArtAndVerticalText => "clipArtAndVertTx",
            Self::VerticalTitleAndText => "vertTitleAndTx",
            Self::VerticalTitleAndTextOverChart => "vertTitleAndTxOverChart",
            Self::TwoObjects => "twoObj",
            Self::ObjectAndTwoObjects => "objAndTwoObj",
            Self::TwoObjectsAndObject => "twoObjAndObj",
            Self::Custom => "cust",
            Self::SectionHeader => "secHead",
            Self::TwoTextAndTwoObjects => "twoTxTwoObj",
            Self::ObjectText => "objTx",
            Self::PictureWithText => "picTx",
        }
    }
}

/// Placeholder type (`ST_PlaceholderType`, ECMA-376 Part 1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PlaceholderKind {
    Title,
    Body,
    CenteredTitle,
    Subtitle,
    DateTime,
    SlideNumber,
    Footer,
    Header,
    Object,
    Chart,
    Table,
    ClipArt,
    Diagram,
    Media,
    SlideImage,
    Picture,
}

impl PlaceholderKind {
    /// The spec token written to the `type` attribute of `p:ph`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Body => "body",
            Self::CenteredTitle => "ctrTitle",
            Self::Subtitle => "subTitle",
            Self::DateTime => "dt",
            Self::SlideNumber => "sldNum",
            Self::Footer => "ftr",
            Self::Header => "hdr",
            Self::Object => "obj",
            Self::Chart => "chart",
            Self::Table => "tbl",
            Self::ClipArt => "clipArt",
            Self::Diagram => "dgm",
            Self::Media => "media",
            Self::SlideImage => "sldImg",
            Self::Picture => "pic",
        }
    }

    /// Human-readable label used to build default shape names.
    pub(super) const fn label(self) -> &'static str {
        match self {
            Self::Title => "Title",
            Self::Body => "Body",
            Self::CenteredTitle => "Centered Title",
            Self::Subtitle => "Subtitle",
            Self::DateTime => "Date",
            Self::SlideNumber => "Slide Number",
            Self::Footer => "Footer",
            Self::Header => "Header",
            Self::Object => "Object",
            Self::Chart => "Chart",
            Self::Table => "Table",
            Self::ClipArt => "Clip Art",
            Self::Diagram => "Diagram",
            Self::Media => "Media",
            Self::SlideImage => "Slide Image",
            Self::Picture => "Picture",
        }
    }
}

/// A placeholder shape to author on a slide master or slide layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PlaceholderSpec {
    /// Placeholder type written to `p:ph/@type`.
    pub kind: PlaceholderKind,
    /// Placeholder index written to `p:ph/@idx`; omitted when `None`.
    pub index: Option<u32>,
    /// Shape name written to `p:cNvPr/@name`; a default is generated when `None`.
    pub name: Option<String>,
    /// Prompt text written as a single run into the placeholder text body.
    pub text: Option<String>,
}

impl PlaceholderSpec {
    /// Create a placeholder of the given kind with no index, name, or text.
    pub const fn new(kind: PlaceholderKind) -> Self {
        Self {
            kind,
            index: None,
            name: None,
            text: None,
        }
    }

    /// Set the placeholder index (`p:ph/@idx`).
    pub fn with_index(mut self, index: u32) -> Self {
        self.index = Some(index);
        self
    }

    /// Set the shape name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the prompt text of the placeholder.
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }

    /// The index used for identity matching; ECMA defaults `idx` to zero.
    pub(super) fn effective_index(&self) -> u32 {
        self.index.unwrap_or(0)
    }
}

/// Identity of a slide master created by [`add_slide_master`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideMaster {
    /// The `p:sldMasterId/@id` value (always ≥ [`MIN_MASTER_OR_LAYOUT_ID`]).
    pub master_id: u32,
    /// Relationship ID from the presentation part to the master part.
    pub relationship_id: String,
    /// Part name of the new master, e.g. `/ppt/slideMasters/slideMaster2.xml`.
    pub part_name: PackURI,
}

/// Identity of a slide layout created by [`add_slide_layout`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideLayout {
    /// The `p:sldLayoutId/@id` value (always ≥ [`MIN_MASTER_OR_LAYOUT_ID`]).
    pub layout_id: u32,
    /// Relationship ID from the owning master part to the layout part.
    pub relationship_id: String,
    /// Part name of the new layout, e.g. `/ppt/slideLayouts/slideLayout12.xml`.
    pub part_name: PackURI,
    /// Part name of the owning slide master.
    pub master_part_name: PackURI,
}
