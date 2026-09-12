//! Read-only projection of ODF change tracking metadata.

use crate::paragraph::Paragraph;

/// Presence of a `text:tracked-changes` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChangeTracking {
    track_changes: Option<bool>,
}

impl ChangeTracking {
    pub(crate) const fn projected(track_changes: Option<bool>) -> Self {
        Self { track_changes }
    }

    /// Optional `text:track-changes` flag.
    #[must_use]
    pub const fn track_changes(&self) -> Option<bool> {
        self.track_changes
    }
}

/// Exact metadata and cached paragraphs from one `office:change-info`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChangeInfo {
    creator: String,
    date: String,
    paragraphs: Vec<Paragraph>,
}

impl ChangeInfo {
    pub(crate) const fn projected(
        creator: String,
        date: String,
        paragraphs: Vec<Paragraph>,
    ) -> Self {
        Self {
            creator,
            date,
            paragraphs,
        }
    }

    /// Required creator lexical value.
    #[must_use]
    pub fn creator(&self) -> &str {
        &self.creator
    }
    /// Required validated date-time lexical value.
    #[must_use]
    pub fn date(&self) -> &str {
        &self.date
    }
    /// Direct change-info paragraphs in source order.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }
}

/// Kind of tracked change or review marker.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// A `text:insertion` region.
    Insertion,
    /// A `text:deletion` region.
    Deletion,
    /// A `text:format-change` region.
    Format,
    /// A producer move/change extension.
    Move,
    /// A `text:change-start` marker.
    Start,
    /// A `text:change-end` marker.
    End,
    /// A valid but unclassified producer region.
    Other(String),
}

impl Kind {
    /// Lexical ODF element or marker name represented by this kind.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Insertion => "insertion",
            Self::Deletion => "deletion",
            Self::Format => "format-change",
            Self::Move => "move",
            Self::Start => "change-start",
            Self::End => "change-end",
            Self::Other(value) => value,
        }
    }
}

/// One inert tracked-change region or marker.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Change {
    author: Option<String>,
    date: Option<String>,
    id: Option<String>,
    info: Option<ChangeInfo>,
    kind: Kind,
    marker_change_id: Option<String>,
    region_id: Option<String>,
    text: String,
    xml_id: Option<String>,
}

impl Change {
    pub(crate) const fn projected(
        id: Option<String>,
        kind: Kind,
        author: Option<String>,
        date: Option<String>,
        xml_id: Option<String>,
        region_id: Option<String>,
        marker_change_id: Option<String>,
        info: Option<ChangeInfo>,
        text: String,
    ) -> Self {
        Self {
            author,
            date,
            id,
            info,
            kind,
            marker_change_id,
            region_id,
            text,
            xml_id,
        }
    }

    /// Producer change identity.
    #[must_use]
    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }

    /// Required `xml:id` on a changed region, when this is a region.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Deprecated `text:id` region identity, when present.
    #[must_use]
    pub fn region_id(&self) -> Option<&str> {
        self.region_id.as_deref()
    }

    /// Marker `text:change-id` identity, when this is a marker.
    #[must_use]
    pub fn marker_change_id(&self) -> Option<&str> {
        self.marker_change_id.as_deref()
    }

    /// Exact change-info metadata for a changed region.
    #[must_use]
    pub const fn info(&self) -> Option<&ChangeInfo> {
        self.info.as_ref()
    }

    /// Tracked change kind.
    #[must_use]
    pub const fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Review author, if supplied by the producer.
    #[must_use]
    pub fn author(&self) -> Option<&str> {
        self.author.as_deref()
    }

    /// Review timestamp, if supplied by the producer.
    #[must_use]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Inert text represented by this changed region.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}
