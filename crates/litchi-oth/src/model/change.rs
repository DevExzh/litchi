//! Read-only projection of ODF change tracking metadata.

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
    kind: Kind,
    text: String,
}

impl Change {
    pub(crate) const fn projected(
        id: Option<String>,
        kind: Kind,
        author: Option<String>,
        date: Option<String>,
        text: String,
    ) -> Self {
        Self {
            author,
            date,
            id,
            kind,
            text,
        }
    }

    /// Producer change identity.
    #[must_use]
    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
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
