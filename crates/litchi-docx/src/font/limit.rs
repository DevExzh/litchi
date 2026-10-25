//! Typed budgets of the font-table family (record 0355a).

use super::model::{MAX_ALL_FONTS, MAX_DEPTH, MAX_FONT, MAX_FONTS, MAX_NODES, MAX_TEXT, MAX_XML};

/// Caller-selected admission limits for the raw font-table codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    xml_bytes: usize,
    embedded_font_bytes: usize,
    aggregate_font_bytes: usize,
    fonts: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
    text_bytes: usize,
}

impl Limits {
    /// The standard bounded raw font-table policy.
    pub const DEFAULT: Self = Self {
        xml_bytes: MAX_XML,
        embedded_font_bytes: MAX_FONT,
        aggregate_font_bytes: MAX_ALL_FONTS,
        fonts: MAX_FONTS,
        depth: MAX_DEPTH,
        nodes: MAX_NODES,
        attributes: MAX_NODES,
        attribute_bytes: MAX_TEXT,
        text_bytes: MAX_TEXT,
    };

    /// Construct a finite raw font-table policy.
    #[must_use]
    pub const fn new(
        xml_bytes: usize,
        embedded_font_bytes: usize,
        aggregate_font_bytes: usize,
        fonts: usize,
        depth: usize,
        nodes: usize,
        attributes: usize,
        attribute_bytes: usize,
        text_bytes: usize,
    ) -> Option<Self> {
        if xml_bytes == 0
            || xml_bytes > MAX_XML
            || embedded_font_bytes == 0
            || embedded_font_bytes > MAX_FONT
            || aggregate_font_bytes == 0
            || aggregate_font_bytes > MAX_ALL_FONTS
            || fonts == 0
            || fonts > MAX_FONTS
            || depth == 0
            || depth > MAX_DEPTH
            || nodes == 0
            || nodes > MAX_NODES
            || attributes == 0
            || attributes > MAX_NODES
            || attribute_bytes == 0
            || attribute_bytes > MAX_TEXT
            || text_bytes == 0
            || text_bytes > MAX_TEXT
        {
            None
        } else {
            Some(Self {
                xml_bytes,
                embedded_font_bytes,
                aggregate_font_bytes,
                fonts,
                depth,
                nodes,
                attributes,
                attribute_bytes,
                text_bytes,
            })
        }
    }

    #[must_use]
    pub const fn xml_bytes(self) -> usize {
        self.xml_bytes
    }
    #[must_use]
    pub const fn embedded_font_bytes(self) -> usize {
        self.embedded_font_bytes
    }
    #[must_use]
    pub const fn aggregate_font_bytes(self) -> usize {
        self.aggregate_font_bytes
    }
    #[must_use]
    pub const fn fonts(self) -> usize {
        self.fonts
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
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The bounded resources of the font-table family, carried by [`crate::Error::FontLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Font-table XML bytes.
    XmlBytes,
    /// Font-table XML depth.
    XmlDepth,
    /// Font-table XML attribute count.
    XmlAttributeCount,
    /// Font-table XML attribute bytes.
    XmlAttributeBytes,
    /// Font-table font count.
    FontCount,
    /// Aggregate embedded-font bytes.
    AggregateBytes,
    /// Embedded font bytes.
    EmbeddedFontBytes,
    /// Font-table string bytes.
    StringBytes,
    /// Font-table text bytes.
    Text,
    /// Font-usage XML node.
    XmlNode,
    /// Font-table XML structure.
    XmlStructure,
    /// Word 2010 OpenType XML bytes.
    OpenTypeXmlBytes,
    /// Word 2010 OpenType XML depth.
    OpenTypeXmlDepth,
    /// Word 2010 OpenType XML node count.
    OpenTypeXmlNodes,
    /// Word 2010 OpenType XML attribute count.
    OpenTypeXmlAttributes,
    /// Word 2010 OpenType XML attribute bytes.
    OpenTypeXmlAttributeBytes,
    /// Word 2010 OpenType stylistic-set count.
    OpenTypeStyleSets,
    /// Format-independent automatic font preparation.
    FontPreparation,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "font-table XML bytes",
            Self::XmlDepth => "font-table XML depth",
            Self::XmlAttributeCount => "font-table XML attribute count",
            Self::XmlAttributeBytes => "font-table XML attribute bytes",
            Self::FontCount => "font-table count",
            Self::AggregateBytes => "aggregate embedded-font bytes",
            Self::EmbeddedFontBytes => "embedded font bytes",
            Self::StringBytes => "font-table string bytes",
            Self::Text => "font-table text bytes",
            Self::XmlNode => "font-usage XML node",
            Self::XmlStructure => "font-table XML structure",
            Self::OpenTypeXmlBytes => "OpenType XML bytes",
            Self::OpenTypeXmlDepth => "OpenType XML depth",
            Self::OpenTypeXmlNodes => "OpenType XML node count",
            Self::OpenTypeXmlAttributes => "OpenType XML attribute count",
            Self::OpenTypeXmlAttributeBytes => "OpenType XML attribute bytes",
            Self::OpenTypeStyleSets => "OpenType stylistic-set count",
            Self::FontPreparation => "automatic font preparation",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
