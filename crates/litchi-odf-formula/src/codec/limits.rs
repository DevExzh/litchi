//! Caller-selectable resource limits for inert `MathML` processing.

use std::fmt;

/// Absolute package-byte safety ceiling.
pub const HARD_MAX_PACKAGE_BYTES: usize = 512 * 1024 * 1024;
/// Absolute `content.xml` byte safety ceiling.
pub const HARD_MAX_XML_BYTES: usize = 64 * 1024 * 1024;
/// Absolute element-nesting safety ceiling.
pub const HARD_MAX_DEPTH: usize = 4_096;
/// Absolute element-count safety ceiling.
pub const HARD_MAX_NODES: usize = 1_000_000;
/// Absolute per-element attribute-count safety ceiling.
pub const HARD_MAX_ATTRIBUTES: usize = 4_096;
/// Absolute decoded byte ceiling for one attribute value.
pub const HARD_MAX_ATTRIBUTE_BYTES: usize = 16 * 1024 * 1024;
/// Absolute aggregate decoded text safety ceiling.
pub const HARD_MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// The resource controlled by a [`Limits`] field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    PackageBytes,
    XmlBytes,
    Depth,
    Nodes,
    Attributes,
    AttributeBytes,
    TextBytes,
}

/// An invalid caller-selected resource limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LimitError {
    kind: LimitKind,
    requested: usize,
    hard_maximum: usize,
}

impl LimitError {
    const fn new(kind: LimitKind, requested: usize, hard_maximum: usize) -> Self {
        Self {
            kind,
            requested,
            hard_maximum,
        }
    }

    /// The resource whose requested limit is invalid.
    #[must_use]
    pub const fn kind(&self) -> LimitKind {
        self.kind
    }

    /// The rejected value.
    #[must_use]
    pub const fn requested(&self) -> usize {
        self.requested
    }

    /// The immutable safety ceiling for this resource.
    #[must_use]
    pub const fn hard_maximum(&self) -> usize {
        self.hard_maximum
    }
}

impl fmt::Display for LimitError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "invalid {:?} limit {}; expected 1..={}",
            self.kind, self.requested, self.hard_maximum
        )
    }
}

impl std::error::Error for LimitError {}

/// Hierarchical resource limits for Formula package and `MathML` ingress.
///
/// Defaults retain the crate's established parser ceilings. Every field can
/// be lowered or raised through checked builders, but never beyond its hard
/// safety ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    package_bytes: usize,
    xml_bytes: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
    text_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            package_bytes: 256 * 1024 * 1024,
            xml_bytes: HARD_MAX_XML_BYTES,
            depth: 128,
            nodes: 65_536,
            attributes: 256,
            attribute_bytes: 1024 * 1024,
            text_bytes: 32 * 1024 * 1024,
        }
    }
}

impl Limits {
    /// Return the production-safe default profile.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Maximum encoded package bytes.
    #[must_use]
    pub const fn package_bytes(self) -> usize {
        self.package_bytes
    }

    /// Maximum encoded `content.xml` bytes.
    #[must_use]
    pub const fn xml_bytes(self) -> usize {
        self.xml_bytes
    }

    /// Maximum simultaneously open `MathML` elements.
    #[must_use]
    pub const fn depth(self) -> usize {
        self.depth
    }

    /// Maximum total `MathML` elements.
    #[must_use]
    pub const fn nodes(self) -> usize {
        self.nodes
    }

    /// Maximum attributes on one element.
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }

    /// Maximum decoded bytes in one attribute value.
    #[must_use]
    pub const fn attribute_bytes(self) -> usize {
        self.attribute_bytes
    }

    /// Maximum aggregate decoded text bytes.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Return a copy with a checked package-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_package_bytes(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::PackageBytes, value, HARD_MAX_PACKAGE_BYTES)?;
        self.package_bytes = value;
        Ok(self)
    }

    /// Return a copy with a checked `content.xml` byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_xml_bytes(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::XmlBytes, value, HARD_MAX_XML_BYTES)?;
        self.xml_bytes = value;
        Ok(self)
    }

    /// Return a copy with a checked element-depth ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_depth(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::Depth, value, HARD_MAX_DEPTH)?;
        self.depth = value;
        Ok(self)
    }

    /// Return a copy with a checked element-count ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_nodes(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::Nodes, value, HARD_MAX_NODES)?;
        self.nodes = value;
        Ok(self)
    }

    /// Return a copy with a checked per-element attribute ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_attributes(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::Attributes, value, HARD_MAX_ATTRIBUTES)?;
        self.attributes = value;
        Ok(self)
    }

    /// Return a copy with a checked single-attribute byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_attribute_bytes(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::AttributeBytes, value, HARD_MAX_ATTRIBUTE_BYTES)?;
        self.attribute_bytes = value;
        Ok(self)
    }

    /// Return a copy with a checked aggregate text-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_text_bytes(mut self, value: usize) -> Result<Self, LimitError> {
        check(LimitKind::TextBytes, value, HARD_MAX_TEXT_BYTES)?;
        self.text_bytes = value;
        Ok(self)
    }
}

fn check(kind: LimitKind, value: usize, hard_maximum: usize) -> Result<(), LimitError> {
    if value == 0 || value > hard_maximum {
        Err(LimitError::new(kind, value, hard_maximum))
    } else {
        Ok(())
    }
}
