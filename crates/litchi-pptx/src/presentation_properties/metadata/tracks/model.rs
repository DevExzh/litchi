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
