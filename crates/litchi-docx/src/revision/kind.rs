//! Typed revision budgets (record 0355a).

/// The bounded resources of revision reads and tracked-revision authoring, carried by [`crate::Error::RevisionLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Attribute bytes.
    AttributeBytes,
    /// Attributes.
    Attributes,
    /// Changes.
    Changes,
    /// Depth.
    Depth,
    /// Events.
    Events,
    /// Inherited namespace bytes.
    InheritedNamespaceBytes,
    /// Inherited namespaces.
    InheritedNamespaces,
    /// Metadata bytes.
    MetadataBytes,
    /// Move ranges.
    MoveRanges,
    /// Output bytes.
    OutputBytes,
    /// Records.
    Records,
    /// Source bytes.
    SourceBytes,
    /// Text bytes.
    TextBytes,
    /// Value bytes.
    ValueBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AttributeBytes => "attribute bytes",
            Self::Attributes => "attributes",
            Self::Changes => "changes",
            Self::Depth => "depth",
            Self::Events => "events",
            Self::InheritedNamespaceBytes => "inherited namespace bytes",
            Self::InheritedNamespaces => "inherited namespaces",
            Self::MetadataBytes => "metadata bytes",
            Self::MoveRanges => "move ranges",
            Self::OutputBytes => "output bytes",
            Self::Records => "records",
            Self::SourceBytes => "source bytes",
            Self::TextBytes => "text bytes",
            Self::ValueBytes => "value bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
