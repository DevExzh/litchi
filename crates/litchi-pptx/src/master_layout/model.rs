//! Typed `PresentationML` models for slide masters, layouts, and placeholders.

use litchi_opc::packuri::PackURI;

use crate::shape::PlaceholderTypeExtension;

/// ECMA-376 Part 1: slide master and slide layout IDs start at 2^31.
pub const MIN_MASTER_OR_LAYOUT_ID: u32 = 2_147_483_648;

/// Strip the XML whitespace (`#x20`, tab, carriage return, line feed) that the
/// `xsd:token` whitespace facet removes from both ends of a token.
fn trim_xml_whitespace(value: &str) -> &str {
    value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'))
}

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
    /// Every layout type, in `ST_SlideLayoutType` schema order.
    pub const ALL: [Self; 36] = [
        Self::Title,
        Self::Text,
        Self::TwoColumnText,
        Self::Table,
        Self::TextAndChart,
        Self::ChartAndText,
        Self::Diagram,
        Self::Chart,
        Self::TextAndClipArt,
        Self::ClipArtAndText,
        Self::TitleOnly,
        Self::Blank,
        Self::TextAndObject,
        Self::ObjectAndText,
        Self::ObjectOnly,
        Self::Object,
        Self::TextAndMedia,
        Self::MediaAndText,
        Self::ObjectOverText,
        Self::TextOverObject,
        Self::TextAndTwoObjects,
        Self::TwoObjectsAndText,
        Self::TwoObjectsOverText,
        Self::FourObjects,
        Self::VerticalText,
        Self::ClipArtAndVerticalText,
        Self::VerticalTitleAndText,
        Self::VerticalTitleAndTextOverChart,
        Self::TwoObjects,
        Self::ObjectAndTwoObjects,
        Self::TwoObjectsAndObject,
        Self::Custom,
        Self::SectionHeader,
        Self::TwoTextAndTwoObjects,
        Self::ObjectText,
        Self::PictureWithText,
    ];

    /// Classify a `p:sldLayout@type` token, the inverse of [`Self::as_str`].
    ///
    /// `ST_SlideLayoutType` restricts `xsd:token`, so leading and trailing
    /// XML whitespace is ignored; the comparison is otherwise exact and
    /// case-sensitive, as in the schema. An unknown token is `None`. A layout
    /// whose `type` attribute is absent has the schema default `cust`
    /// ([`Self::Custom`]); that default is the caller's to apply, because
    /// [`crate::SlideLayout::kind`] reports an absent attribute as `None`.
    ///
    /// ```
    /// use litchi_pptx::SlideLayoutKind;
    ///
    /// assert_eq!(SlideLayoutKind::from_token("twoColTx"), Some(SlideLayoutKind::TwoColumnText));
    /// assert_eq!(SlideLayoutKind::from_token(" secHead\n"), Some(SlideLayoutKind::SectionHeader));
    /// assert_eq!(SlideLayoutKind::from_token("TwoColTx"), None);
    /// ```
    #[must_use]
    pub fn from_token(token: &str) -> Option<Self> {
        Some(match trim_xml_whitespace(token) {
            "title" => Self::Title,
            "tx" => Self::Text,
            "twoColTx" => Self::TwoColumnText,
            "tbl" => Self::Table,
            "txAndChart" => Self::TextAndChart,
            "chartAndTx" => Self::ChartAndText,
            "dgm" => Self::Diagram,
            "chart" => Self::Chart,
            "txAndClipArt" => Self::TextAndClipArt,
            "clipArtAndTx" => Self::ClipArtAndText,
            "titleOnly" => Self::TitleOnly,
            "blank" => Self::Blank,
            "txAndObj" => Self::TextAndObject,
            "objAndTx" => Self::ObjectAndText,
            "objOnly" => Self::ObjectOnly,
            "obj" => Self::Object,
            "txAndMedia" => Self::TextAndMedia,
            "mediaAndTx" => Self::MediaAndText,
            "objOverTx" => Self::ObjectOverText,
            "txOverObj" => Self::TextOverObject,
            "txAndTwoObj" => Self::TextAndTwoObjects,
            "twoObjAndTx" => Self::TwoObjectsAndText,
            "twoObjOverTx" => Self::TwoObjectsOverText,
            "fourObj" => Self::FourObjects,
            "vertTx" => Self::VerticalText,
            "clipArtAndVertTx" => Self::ClipArtAndVerticalText,
            "vertTitleAndTx" => Self::VerticalTitleAndText,
            "vertTitleAndTxOverChart" => Self::VerticalTitleAndTextOverChart,
            "twoObj" => Self::TwoObjects,
            "objAndTwoObj" => Self::ObjectAndTwoObjects,
            "twoObjAndObj" => Self::TwoObjectsAndObject,
            "cust" => Self::Custom,
            "secHead" => Self::SectionHeader,
            "twoTxTwoObj" => Self::TwoTextAndTwoObjects,
            "objTx" => Self::ObjectText,
            "picTx" => Self::PictureWithText,
            _ => return None,
        })
    }

    /// The spec token written to the `type` attribute of `p:sldLayout`.
    #[must_use]
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
    /// Every placeholder type, in `ST_PlaceholderType` schema order.
    pub const ALL: [Self; 16] = [
        Self::Title,
        Self::Body,
        Self::CenteredTitle,
        Self::Subtitle,
        Self::DateTime,
        Self::SlideNumber,
        Self::Footer,
        Self::Header,
        Self::Object,
        Self::Chart,
        Self::Table,
        Self::ClipArt,
        Self::Diagram,
        Self::Media,
        Self::SlideImage,
        Self::Picture,
    ];

    /// Classify a `p:ph@type` token, the inverse of [`Self::as_str`].
    ///
    /// `ST_PlaceholderType` restricts `xsd:token`, so leading and trailing
    /// XML whitespace is ignored; the comparison is otherwise exact and
    /// case-sensitive, as in the schema. An unknown token is `None`. A
    /// placeholder whose `type` attribute is absent has the schema default
    /// `obj` ([`Self::Object`]); that default is the caller's to apply,
    /// because [`crate::shape::Placeholder::kind`] reports an absent
    /// attribute as `None`.
    ///
    /// ```
    /// use litchi_pptx::PlaceholderKind;
    ///
    /// assert_eq!(PlaceholderKind::from_token("ctrTitle"), Some(PlaceholderKind::CenteredTitle));
    /// assert_eq!(PlaceholderKind::from_token("\tsldNum "), Some(PlaceholderKind::SlideNumber));
    /// assert_eq!(PlaceholderKind::from_token("Title"), None);
    /// ```
    #[must_use]
    pub fn from_token(token: &str) -> Option<Self> {
        Some(match trim_xml_whitespace(token) {
            "title" => Self::Title,
            "body" => Self::Body,
            "ctrTitle" => Self::CenteredTitle,
            "subTitle" => Self::Subtitle,
            "dt" => Self::DateTime,
            "sldNum" => Self::SlideNumber,
            "ftr" => Self::Footer,
            "hdr" => Self::Header,
            "obj" => Self::Object,
            "chart" => Self::Chart,
            "tbl" => Self::Table,
            "clipArt" => Self::ClipArt,
            "dgm" => Self::Diagram,
            "media" => Self::Media,
            "sldImg" => Self::SlideImage,
            "pic" => Self::Picture,
            _ => return None,
        })
    }

    /// The spec token written to the `type` attribute of `p:ph`.
    #[must_use]
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
    /// Optional PowerPoint 2023 placeholder type extension.
    pub type_extension: Option<PlaceholderTypeExtension>,
}

impl PlaceholderSpec {
    /// Create a placeholder of the given kind with no index, name, or text.
    #[must_use]
    pub const fn new(kind: PlaceholderKind) -> Self {
        Self {
            kind,
            index: None,
            name: None,
            text: None,
            type_extension: None,
        }
    }

    /// Set the placeholder index (`p:ph/@idx`).
    #[must_use]
    pub fn with_index(mut self, index: u32) -> Self {
        self.index = Some(index);
        self
    }

    /// Set the shape name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the prompt text of the placeholder.
    #[must_use]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }

    /// Set the PowerPoint 2023 placeholder type extension.
    #[must_use]
    pub const fn with_type_extension(mut self, extension: PlaceholderTypeExtension) -> Self {
        self.type_extension = Some(extension);
        self
    }

    /// The index used for identity matching; ECMA defaults `idx` to zero.
    pub(super) fn effective_index(&self) -> u32 {
        self.index.unwrap_or(0)
    }
}

/// Identity of a slide master created by [`Edit::add_slide_master`](crate::master_layout::Edit::add_slide_master).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideMaster {
    /// The `p:sldMasterId/@id` value (always ≥ [`MIN_MASTER_OR_LAYOUT_ID`]).
    pub master_id: u32,
    /// Relationship ID from the presentation part to the master part.
    pub relationship_id: String,
    /// Part name of the new master, e.g. `/ppt/slideMasters/slideMaster2.xml`.
    pub part_name: PackURI,
}

/// Identity of a slide layout created by [`Edit::add_slide_layout`](crate::master_layout::Edit::add_slide_layout).
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

#[cfg(test)]
mod token_tests {
    use super::{PlaceholderKind, SlideLayoutKind};

    #[test]
    fn layout_tokens_round_trip_every_schema_value() {
        for (position, kind) in SlideLayoutKind::ALL.into_iter().enumerate() {
            assert_eq!(SlideLayoutKind::from_token(kind.as_str()), Some(kind));
            // Each schema token names exactly one variant.
            assert!(
                SlideLayoutKind::ALL[..position]
                    .iter()
                    .all(|earlier| earlier.as_str() != kind.as_str()),
                "{kind:?}"
            );
        }
    }

    #[test]
    fn placeholder_tokens_round_trip_every_schema_value() {
        for (position, kind) in PlaceholderKind::ALL.into_iter().enumerate() {
            assert_eq!(PlaceholderKind::from_token(kind.as_str()), Some(kind));
            assert!(
                PlaceholderKind::ALL[..position]
                    .iter()
                    .all(|earlier| earlier.as_str() != kind.as_str()),
                "{kind:?}"
            );
        }
    }

    #[test]
    fn tokens_apply_the_xsd_token_facet_and_stay_case_sensitive() {
        assert_eq!(
            PlaceholderKind::from_token(" \t\r\nsubTitle\n"),
            Some(PlaceholderKind::Subtitle)
        );
        assert_eq!(
            SlideLayoutKind::from_token("\tvertTitleAndTxOverChart "),
            Some(SlideLayoutKind::VerticalTitleAndTextOverChart)
        );
        for unknown in [
            "",
            " ",
            "Title",
            "TITLE",
            "sub Title",
            "subtitle",
            "\u{a0}title",
        ] {
            assert_eq!(PlaceholderKind::from_token(unknown), None, "{unknown:?}");
        }
        for unknown in [
            "",
            "Blank",
            "two ColTx",
            "custom",
            "titleonly",
            "\u{a0}blank",
        ] {
            assert_eq!(SlideLayoutKind::from_token(unknown), None, "{unknown:?}");
        }
    }
}
