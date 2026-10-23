//! Typed budgets of the glossary family (record 0355a).

use super::{
    MAX, MAX_DEPTH, MAX_DOM_ATTRIBUTES, MAX_DOM_BYTES, MAX_DOM_CONTENT, MAX_DOM_TOKENS,
    MAX_GRAPH_BYTES, MAX_GRAPH_METADATA_BYTES, MAX_NAME_KEY, MAX_NODES, MAX_PARTS, MAX_STRING,
    MAX_VALUES,
};

/// Caller-selected admission limits for the raw glossary XML and graph codec.
///
/// The format ceilings remain hard invariants. The raw entry points require a
/// policy and reject input before retaining a DOM or graph payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    xml_bytes: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
    content: usize,
    tokens: usize,
    dom_bytes: usize,
    parts: usize,
    values: usize,
    string_bytes: usize,
    name_bytes: usize,
    graph_bytes: usize,
    metadata_bytes: usize,
}

/// The two aggregate byte budgets for the auxiliary glossary graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GraphLimits {
    graph_bytes: usize,
    metadata_bytes: usize,
}

impl GraphLimits {
    /// Group the graph and graph-metadata byte budgets.
    #[must_use]
    pub const fn new(graph_bytes: usize, metadata_bytes: usize) -> Self {
        Self {
            graph_bytes,
            metadata_bytes,
        }
    }
}

impl Limits {
    /// The standard bounded raw glossary policy.
    pub const DEFAULT: Self = Self {
        xml_bytes: MAX,
        depth: MAX_DEPTH,
        nodes: MAX_NODES,
        attributes: MAX_DOM_ATTRIBUTES,
        content: MAX_DOM_CONTENT,
        tokens: MAX_DOM_TOKENS,
        dom_bytes: MAX_DOM_BYTES,
        parts: MAX_PARTS,
        values: MAX_VALUES,
        string_bytes: MAX_STRING,
        name_bytes: MAX_NAME_KEY,
        graph_bytes: MAX_GRAPH_BYTES,
        metadata_bytes: MAX_GRAPH_METADATA_BYTES,
    };

    /// Construct a finite raw glossary policy no wider than the format.
    #[must_use]
    pub const fn new(
        xml_bytes: usize,
        depth: usize,
        nodes: usize,
        attributes: usize,
        content: usize,
        tokens: usize,
        dom_bytes: usize,
        parts: usize,
        values: usize,
        string_bytes: usize,
        name_bytes: usize,
        graph: GraphLimits,
    ) -> Option<Self> {
        if xml_bytes == 0
            || xml_bytes > MAX
            || depth == 0
            || depth > MAX_DEPTH
            || nodes == 0
            || nodes > MAX_NODES
            || attributes == 0
            || attributes > MAX_DOM_ATTRIBUTES
            || content == 0
            || content > MAX_DOM_CONTENT
            || tokens == 0
            || tokens > MAX_DOM_TOKENS
            || dom_bytes == 0
            || dom_bytes > MAX_DOM_BYTES
            || parts == 0
            || parts > MAX_PARTS
            || values == 0
            || values > MAX_VALUES
            || string_bytes == 0
            || string_bytes > MAX_STRING
            || name_bytes == 0
            || name_bytes > MAX_NAME_KEY
            || graph.graph_bytes == 0
            || graph.graph_bytes > MAX_GRAPH_BYTES
            || graph.metadata_bytes == 0
            || graph.metadata_bytes > MAX_GRAPH_METADATA_BYTES
        {
            None
        } else {
            Some(Self {
                xml_bytes,
                depth,
                nodes,
                attributes,
                content,
                tokens,
                dom_bytes,
                parts,
                values,
                string_bytes,
                name_bytes,
                graph_bytes: graph.graph_bytes,
                metadata_bytes: graph.metadata_bytes,
            })
        }
    }

    #[must_use]
    pub const fn xml_bytes(self) -> usize {
        self.xml_bytes
    }
    #[must_use]
    pub const fn depth(self) -> usize {
        self.depth
    }
    #[must_use]
    pub const fn nodes(self) -> usize {
        self.nodes
    }
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }
    #[must_use]
    pub const fn content(self) -> usize {
        self.content
    }
    #[must_use]
    pub const fn tokens(self) -> usize {
        self.tokens
    }
    #[must_use]
    pub const fn dom_bytes(self) -> usize {
        self.dom_bytes
    }
    #[must_use]
    pub const fn parts(self) -> usize {
        self.parts
    }
    #[must_use]
    pub const fn values(self) -> usize {
        self.values
    }
    #[must_use]
    pub const fn string_bytes(self) -> usize {
        self.string_bytes
    }
    #[must_use]
    pub const fn name_bytes(self) -> usize {
        self.name_bytes
    }
    #[must_use]
    pub const fn graph_bytes(self) -> usize {
        self.graph_bytes
    }
    #[must_use]
    pub const fn metadata_bytes(self) -> usize {
        self.metadata_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The bounded resources of the glossary family, carried by [`crate::Error::GlossaryLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Glossary XML bytes.
    XmlBytes,
    /// Glossary DOM bytes.
    DomBytes,
    /// Glossary graph bytes.
    GraphBytes,
    /// Glossary graph metadata bytes.
    GraphMetadataBytes,
    /// Glossary string bytes.
    StringBytes,
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
            Self::XmlBytes => "glossary XML bytes",
            Self::DomBytes => "glossary DOM bytes",
            Self::GraphBytes => "glossary graph bytes",
            Self::GraphMetadataBytes => "glossary graph metadata bytes",
            Self::StringBytes => "glossary string bytes",
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
