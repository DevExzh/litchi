//! Typed budgets of the story family (record 0355a).

/// The bounded resources of the story family, carried by [`crate::Error::StoryLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Package parts.
    PackageParts,
    /// Story XML prolog events.
    PrologEvents,
    /// Story relationships.
    Relationships,
    /// Relationships per story owner.
    RelationshipsPerOwner,
    /// Stories.
    Stories,
    /// Story bytes.
    StoryBytes,
    /// Story topology bytes.
    TopologyBytes,
    /// Aggregate story bytes.
    TotalStoryBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::PackageParts => "package parts",
            Self::PrologEvents => "story XML prolog events",
            Self::Relationships => "story relationships",
            Self::RelationshipsPerOwner => "relationships per story owner",
            Self::Stories => "stories",
            Self::StoryBytes => "story bytes",
            Self::TopologyBytes => "story topology bytes",
            Self::TotalStoryBytes => "aggregate story bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
