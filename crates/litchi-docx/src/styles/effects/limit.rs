//! Typed budgets of the stylesWithEffects family (record 0355a).

use super::{MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_EVENTS};

/// Caller-selected admission limits for detached stylesWithEffects XML and
/// direct OPC graph operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    xml_bytes: usize,
    events: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
}

impl Limits {
    /// The standard bounded raw stylesWithEffects policy.
    pub const DEFAULT: Self = Self {
        xml_bytes: MAX_XML_BYTES,
        events: MAX_XML_EVENTS,
        depth: MAX_XML_DEPTH,
        nodes: MAX_XML_EVENTS,
        attributes: MAX_XML_EVENTS,
        attribute_bytes: MAX_XML_BYTES,
    };

    /// Construct a finite policy no wider than the format ceilings.
    #[must_use]
    pub const fn new(
        xml_bytes: usize,
        events: usize,
        depth: usize,
        nodes: usize,
        attributes: usize,
        attribute_bytes: usize,
    ) -> Option<Self> {
        if xml_bytes == 0
            || xml_bytes > MAX_XML_BYTES
            || events == 0
            || events > MAX_XML_EVENTS
            || depth == 0
            || depth > MAX_XML_DEPTH
            || nodes == 0
            || nodes > MAX_XML_EVENTS
            || attributes == 0
            || attributes > MAX_XML_EVENTS
            || attribute_bytes == 0
            || attribute_bytes > MAX_XML_BYTES
        {
            None
        } else {
            Some(Self {
                xml_bytes,
                events,
                depth,
                nodes,
                attributes,
                attribute_bytes,
            })
        }
    }

    #[must_use]
    pub const fn xml_bytes(self) -> usize {
        self.xml_bytes
    }
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
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
    pub const fn attribute_bytes(self) -> usize {
        self.attribute_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The bounded resources of the stylesWithEffects family, carried by [`crate::Error::StylesWithEffectsLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// StylesWithEffects XML bytes.
    XmlBytes,
    /// StylesWithEffects XML depth.
    XmlDepth,
    /// StylesWithEffects XML events.
    XmlEvents,
    /// StylesWithEffects XML node count.
    XmlNodes,
    /// StylesWithEffects XML attribute count.
    XmlAttributes,
    /// StylesWithEffects XML attribute bytes.
    XmlAttributeBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "stylesWithEffects XML bytes",
            Self::XmlDepth => "stylesWithEffects XML depth",
            Self::XmlEvents => "stylesWithEffects XML events",
            Self::XmlNodes => "stylesWithEffects XML nodes",
            Self::XmlAttributes => "stylesWithEffects XML attributes",
            Self::XmlAttributeBytes => "stylesWithEffects XML attribute bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
