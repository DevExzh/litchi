//! Typed budgets of the web-settings family (record 0355a).

use super::{MAX_FRAMESET_NESTING, MAX_TEXT_BYTES, MAX_XML_BYTES, MAX_XML_EVENTS};

/// Caller-selected admission limits for detached web-settings XML and raw
/// package operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    xml_bytes: usize,
    events: usize,
    text_bytes: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
}

impl Limits {
    /// The standard bounded raw web-settings policy.
    pub const DEFAULT: Self = Self {
        xml_bytes: MAX_XML_BYTES,
        events: MAX_XML_EVENTS,
        text_bytes: MAX_TEXT_BYTES,
        depth: MAX_FRAMESET_NESTING,
        nodes: MAX_XML_EVENTS,
        attributes: MAX_XML_EVENTS,
    };

    /// Construct a finite raw web-settings policy no wider than the format.
    #[must_use]
    pub const fn new(
        xml_bytes: usize,
        events: usize,
        text_bytes: usize,
        depth: usize,
        nodes: usize,
        attributes: usize,
    ) -> Option<Self> {
        if xml_bytes == 0
            || xml_bytes > MAX_XML_BYTES
            || events == 0
            || events > MAX_XML_EVENTS
            || text_bytes == 0
            || text_bytes > MAX_TEXT_BYTES
            || depth == 0
            || depth > MAX_FRAMESET_NESTING
            || nodes == 0
            || nodes > MAX_XML_EVENTS
            || attributes == 0
            || attributes > MAX_XML_EVENTS
        {
            None
        } else {
            Some(Self {
                xml_bytes,
                events,
                text_bytes,
                depth,
                nodes,
                attributes,
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
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
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
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The bounded resources of the web-settings family, carried by [`crate::Error::WebLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Web-settings XML bytes.
    XmlBytes,
    /// Web-settings XML events.
    XmlEvents,
    /// Web-settings text bytes.
    TextBytes,
    /// Web-settings XML node count.
    XmlNodes,
    /// Web-settings XML attribute count.
    XmlAttributes,
    /// Web frameset depth.
    FramesetDepth,
    /// HTML division depth.
    HtmlDivisionDepth,
    /// Web-settings XML depth.
    SettingsXmlDepth,
    /// Web XML depth.
    XmlDepth,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "web XML bytes",
            Self::XmlEvents => "web XML events",
            Self::TextBytes => "web text bytes",
            Self::XmlNodes => "web XML nodes",
            Self::XmlAttributes => "web XML attributes",
            Self::FramesetDepth => "web frameset depth",
            Self::HtmlDivisionDepth => "HTML division depth",
            Self::SettingsXmlDepth => "web-settings XML depth",
            Self::XmlDepth => "web XML depth",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
