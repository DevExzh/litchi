use super::super::codec::require_nonempty;
use super::super::{
    IMAGE_RELATIONSHIP_TYPE, Result, STANDARD_DEPTH, STANDARD_IMAGE_BYTES, STANDARD_ITEMS,
    STANDARD_NODES, STANDARD_PACKAGE_PARTS, STANDARD_PACKAGE_RELATIONSHIPS,
    STANDARD_PART_ALLOCATIONS, STANDARD_PART_DELETIONS, STANDARD_STRING_BYTES,
    STANDARD_TOTAL_IMAGE_BYTES, STANDARD_TOTAL_STRING_BYTES, STANDARD_TOTAL_XML_BYTES,
    STANDARD_XML_BYTES, STRICT_IMAGE_RELATIONSHIP_TYPE, STRICT_RELATIONSHIPS_NAMESPACE,
    TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
};
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in one source or authored XML part.
    pub xml_bytes: usize,
    /// Maximum aggregate source or authored XML bytes in one operation.
    pub total_xml_bytes: usize,
    /// Maximum XML element nesting depth.
    pub depth: usize,
    /// Maximum element count in one XML part or retained fragment.
    pub nodes: usize,
    /// Maximum panes or items in any schema collection.
    pub items: usize,
    /// Maximum aggregate decoded string bytes in one XML part.
    pub string_bytes: usize,
    /// Maximum aggregate retained XML, decoded strings, and indexed package metadata.
    pub total_string_bytes: usize,
    /// Maximum bytes in one embedded snapshot image.
    pub image_bytes: usize,
    /// Maximum unique embedded snapshot bytes in one package graph.
    pub total_image_bytes: usize,
    /// Maximum number of package parts inspected by one operation.
    pub package_parts: usize,
    /// Maximum aggregate package-level and part-level relationships inspected.
    pub package_relationships: usize,
    /// Maximum new parts or deterministic part-name allocation attempts.
    pub part_allocations: usize,
    /// Maximum old graph parts that one operation may delete.
    pub part_deletions: usize,
}

impl Limits {
    /// Conservative defaults for untrusted packages.
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            xml_bytes: STANDARD_XML_BYTES,
            total_xml_bytes: STANDARD_TOTAL_XML_BYTES,
            depth: STANDARD_DEPTH,
            nodes: STANDARD_NODES,
            items: STANDARD_ITEMS,
            string_bytes: STANDARD_STRING_BYTES,
            total_string_bytes: STANDARD_TOTAL_STRING_BYTES,
            image_bytes: STANDARD_IMAGE_BYTES,
            total_image_bytes: STANDARD_TOTAL_IMAGE_BYTES,
            package_parts: STANDARD_PACKAGE_PARTS,
            package_relationships: STANDARD_PACKAGE_RELATIONSHIPS,
            part_allocations: STANDARD_PART_ALLOCATIONS,
            part_deletions: STANDARD_PART_DELETIONS,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::standard()
    }
}

/// Task-pane docking state with forward-compatible retention of newer values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Dock {
    Left,
    Right,
    Top,
    Bottom,
    Floating,
    Other(String),
}

impl Dock {
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Left => "left",
            Self::Right => "right",
            Self::Top => "top",
            Self::Bottom => "bottom",
            Self::Floating => "float",
            Self::Other(value) => value,
        }
    }

    pub(in crate::web) fn parse(value: &str) -> Result<Self> {
        require_nonempty("dock state", value)?;
        Ok(match value {
            "left" => Self::Left,
            "right" => Self::Right,
            "top" => Self::Top,
            "bottom" => Self::Bottom,
            "float" | "floating" => Self::Floating,
            value => Self::Other(value.to_owned()),
        })
    }
}

impl AsRef<str> for Dock {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

// Internal defaults used by standalone fragment constructors. Package-level
// entry points thread caller-provided limits explicitly.
pub(in crate::web) const MAX_WEB_EXTENSION_XML_BYTES: usize = STANDARD_XML_BYTES;
#[cfg(test)]
pub(in crate::web) const MAX_WEB_EXTENSION_XML_NODES: usize = STANDARD_NODES;
pub(in crate::web) const MAX_WEB_EXTENSION_ITEMS: usize = STANDARD_ITEMS;
pub(in crate::web) const MAX_WEB_EXTENSION_SNAPSHOT_BYTES: usize = STANDARD_IMAGE_BYTES;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(in crate::web) fn relationships_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
            Self::Strict => STRICT_RELATIONSHIPS_NAMESPACE,
        }
    }

    pub(in crate::web) fn image_relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => IMAGE_RELATIONSHIP_TYPE,
            Self::Strict => STRICT_IMAGE_RELATIONSHIP_TYPE,
        }
    }
}
