//! Package-independent WebVTT track values.

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header {
    pub name: String,
    pub value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CueSettingKind {
    Vertical,
    Line,
    Position,
    Size,
    Align,
    Region,
}

impl CueSettingKind {
    pub(crate) fn from(value: &str) -> Option<Self> {
        Some(match value {
            "vertical" => Self::Vertical,
            "line" => Self::Line,
            "position" => Self::Position,
            "size" => Self::Size,
            "align" => Self::Align,
            "region" => Self::Region,
            _ => return None,
        })
    }

    pub(crate) const fn token(self) -> &'static str {
        match self {
            Self::Vertical => "vertical",
            Self::Line => "line",
            Self::Position => "position",
            Self::Size => "size",
            Self::Align => "align",
            Self::Region => "region",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CueSetting {
    pub kind: CueSettingKind,
    pub value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RegionSettingKind {
    Id,
    Width,
    Lines,
    RegionAnchor,
    ViewportAnchor,
    Scroll,
}

impl RegionSettingKind {
    pub(crate) fn from(value: &str) -> Option<Self> {
        Some(match value {
            "id" => Self::Id,
            "width" => Self::Width,
            "lines" => Self::Lines,
            "regionanchor" => Self::RegionAnchor,
            "viewportanchor" => Self::ViewportAnchor,
            "scroll" => Self::Scroll,
            _ => return None,
        })
    }

    pub(crate) const fn token(self) -> &'static str {
        match self {
            Self::Id => "id",
            Self::Width => "width",
            Self::Lines => "lines",
            Self::RegionAnchor => "regionanchor",
            Self::ViewportAnchor => "viewportanchor",
            Self::Scroll => "scroll",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RegionSetting {
    pub kind: RegionSettingKind,
    pub value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cue {
    pub identifier: Option<String>,
    pub start_milliseconds: u64,
    pub end_milliseconds: u64,
    pub settings: Vec<CueSetting>,
    /// Cue content remains inert text.
    pub payload: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Block {
    Cue(Cue),
    Note { header: String, lines: Vec<String> },
    Style { lines: Vec<String> },
    Region { settings: Vec<RegionSetting> },
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct File {
    pub header_text: Option<String>,
    pub headers: Vec<Header>,
    pub blocks: Vec<Block>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    Internal { part_name: String, track: File },
    External { target: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Track {
    pub source_part_name: String,
    pub relationship_id: String,
    pub target: Target,
}

/// The location where PowerPoint renders a caption track.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DisplayLocation {
    /// Render the track over the media object.
    Media,
    /// Render the track over the slide.
    Slide,
}

impl DisplayLocation {
    pub(crate) fn from_token(value: &str) -> crate::Result<Self> {
        match value {
            "media" => Ok(Self::Media),
            "slide" => Ok(Self::Slide),
            _ => Err(crate::Error::Invalid(format!(
                "invalid [MS-PPTX] track display location '{value}'"
            ))),
        }
    }

    pub(crate) const fn token(self) -> &'static str {
        match self {
            Self::Media => "media",
            Self::Slide => "slide",
        }
    }
}

/// The relationship target selected by a typed caption reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CaptionTarget {
    /// An inert internal WebVTT part. Its payload is never decoded by the
    /// metadata owner and remains owned by the OPC package.
    Internal {
        part_name: String,
        content_type: String,
    },
    /// An inert external target. No network access is performed.
    External { target: String },
}

/// One `[MS-PPTX]` `p17:track` caption descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Caption {
    /// Required stable track GUID as authored in `p17:track@id`.
    pub id: String,
    /// Required human-readable label.
    pub label: String,
    /// Optional DrawingML text-language identifier.
    pub language: Option<String>,
    /// Effective relationship target. Both authored relationship attributes
    /// remain available through the package-bound source state.
    pub target: CaptionTarget,
}

/// Typed `[MS-PPTX]` `CT_TracksInfo` metadata for one media object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TracksInfo {
    pub display_location: DisplayLocation,
    pub captions: Vec<Caption>,
}

impl TracksInfo {
    /// Construct an empty bounded track list for a new semantic owner.
    #[must_use]
    pub const fn new(display_location: DisplayLocation) -> Self {
        Self {
            display_location,
            captions: Vec::new(),
        }
    }
}

/// Stable contextual identity of a media picture in a PresentationML slide.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MediaKey {
    pub slide_part_name: String,
    pub shape_id: u32,
}

/// Media caption and narration metadata attached to one media picture.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaMetadata {
    pub key: MediaKey,
    /// The inert `p14:media` relationship ID, when present.
    pub media_relationship_id: Option<String>,
    pub tracks_info: Option<TracksInfo>,
    /// Authored `p15:isNarration@val`, preserving absence as `None`.
    pub narration: Option<bool>,
}
