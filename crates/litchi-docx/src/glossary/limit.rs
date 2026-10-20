//! Typed budgets of the glossary family (record 0355a).

/// The bounded resources of the glossary family, carried by [`crate::Error::GlossaryLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Glossary auxiliary part.
    AuxiliaryPart,
    /// Behavior.
    Behavior,
    /// Glossary entry.
    Entry,
    /// Glossary graph-wide relationship.
    GraphWideRelationship,
    /// Glossary namespace binding.
    NamespaceBinding,
    /// Glossary owned-part.
    OwnedPart,
    /// Glossary producer entry.
    ProducerEntry,
    /// Glossary relationship.
    Relationship,
    /// Type.
    Type,
    /// Glossary XML aggregate attribute.
    XmlAggregateAttribute,
    /// Glossary XML attribute.
    XmlAttribute,
    /// Glossary XML content token.
    XmlContentToken,
    /// Glossary XML depth.
    XmlDepth,
    /// Glossary XML node.
    XmlNode,
    /// Glossary XML token.
    XmlToken,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AuxiliaryPart => "glossary auxiliary part",
            Self::Behavior => "behavior",
            Self::Entry => "glossary entry",
            Self::GraphWideRelationship => "glossary graph-wide relationship",
            Self::NamespaceBinding => "glossary namespace binding",
            Self::OwnedPart => "glossary owned-part",
            Self::ProducerEntry => "glossary producer entry",
            Self::Relationship => "glossary relationship",
            Self::Type => "type",
            Self::XmlAggregateAttribute => "glossary XML aggregate attribute",
            Self::XmlAttribute => "glossary XML attribute",
            Self::XmlContentToken => "glossary XML content token",
            Self::XmlDepth => "glossary XML depth",
            Self::XmlNode => "glossary XML node",
            Self::XmlToken => "glossary XML token",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
