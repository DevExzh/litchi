//! Typed budgets of the conflict-markup family (record 0355a).

/// The bounded resources of the conflict-markup family, carried by [`crate::Error::ConflictLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Conflict attribute bytes.
    AttributeBytes,
    /// Conflicts.
    Conflicts,
    /// XML depth.
    Depth,
    /// XML events.
    Events,
    /// MCE candidates.
    MceCandidates,
    /// MCE prepass depth.
    MceDepth,
    /// MCE prepass events.
    MceEvents,
    /// MCE marked source bytes.
    MceMarkedBytes,
    /// MCE prepass namespace declarations.
    MceNamespaceDeclarations,
    /// Conflict metadata bytes.
    MetadataBytes,
    /// Output bytes.
    OutputBytes,
    /// Mc:ProcessContent active targets.
    ProcessContentTargets,
    /// Mc:ProcessContent tokens.
    ProcessContentTokens,
    /// Conflict range-end attribute bytes.
    RangeEndAttributeBytes,
    /// Conflict ranges.
    Ranges,
    /// Source bytes.
    SourceBytes,
    /// Conflict text bytes.
    TextBytes,
    /// Aggregate conflict elements.
    TotalConflicts,
    /// Aggregate metadata bytes.
    TotalMetadataBytes,
    /// Aggregate conflict ranges.
    TotalRanges,
    /// Aggregate story bytes.
    TotalStoryBytes,
    /// Aggregate text segments.
    TotalTextSegments,
    /// XML attribute bytes.
    XmlAttributeBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AttributeBytes => "conflict attribute bytes",
            Self::Conflicts => "conflicts",
            Self::Depth => "XML depth",
            Self::Events => "XML events",
            Self::MceCandidates => "MCE candidates",
            Self::MceDepth => "MCE prepass depth",
            Self::MceEvents => "MCE prepass events",
            Self::MceMarkedBytes => "MCE marked source bytes",
            Self::MceNamespaceDeclarations => "MCE prepass namespace declarations",
            Self::MetadataBytes => "conflict metadata bytes",
            Self::OutputBytes => "output bytes",
            Self::ProcessContentTargets => "mc:ProcessContent active targets",
            Self::ProcessContentTokens => "mc:ProcessContent tokens",
            Self::RangeEndAttributeBytes => "conflict range-end attribute bytes",
            Self::Ranges => "conflict ranges",
            Self::SourceBytes => "source bytes",
            Self::TextBytes => "conflict text bytes",
            Self::TotalConflicts => "aggregate conflict elements",
            Self::TotalMetadataBytes => "aggregate metadata bytes",
            Self::TotalRanges => "aggregate conflict ranges",
            Self::TotalStoryBytes => "aggregate story bytes",
            Self::TotalTextSegments => "aggregate text segments",
            Self::XmlAttributeBytes => "XML attribute bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
