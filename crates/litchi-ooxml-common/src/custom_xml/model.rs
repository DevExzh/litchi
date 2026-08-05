//! Typed Custom XML Data Storage values and bounded authoring requests.

use crate::mce::Name;
use litchi_opc::PackURI;
use std::sync::Arc;

/// Transitional Custom XML Data Storage namespace.
pub const TRANSITIONAL_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
/// Strict Custom XML Data Storage namespace.
pub const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/officeDocument/customXml";
/// Transitional relationship from a host part to a Custom XML data part.
pub const TRANSITIONAL_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
/// Strict relationship from a host part to a Custom XML data part.
pub const STRICT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
/// Transitional relationship from a data part to its properties part.
pub const TRANSITIONAL_PROPS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
/// Strict relationship from a data part to its properties part.
pub const STRICT_PROPS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps";
/// Content type required for a Custom XML properties part.
pub const PROPS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

/// Maximum bytes accepted in one Custom XML payload part.
pub const MAX_PART_BYTES: usize = 16 * 1024 * 1024;
/// Maximum bytes accepted in one Custom XML properties part.
pub const MAX_PROPS_BYTES: usize = 1024 * 1024;
/// Maximum XML element nesting in payloads and properties.
pub const MAX_DEPTH: usize = 256;
/// Maximum Custom XML relationship occurrences in a package.
pub const MAX_ITEMS: usize = 4096;
/// Maximum schema references in one properties part.
pub const MAX_SCHEMA_REFS: usize = 4096;
/// Maximum elements scanned in one payload.
pub const MAX_ELEMENTS: usize = 1_000_000;
/// Maximum aggregate bytes in property strings.
pub const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;

/// OOXML namespace and relationship family used for newly-authored parts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Conformance {
    /// ECMA-376 transitional namespace family.
    Transitional,
    /// ISO/IEC 29500 strict namespace family.
    Strict,
}

impl Conformance {
    /// Namespace used by the properties vocabulary.
    #[must_use]
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_NAMESPACE,
            Self::Strict => STRICT_NAMESPACE,
        }
    }

    /// Relationship used from a host part to the data part.
    #[must_use]
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIP,
            Self::Strict => STRICT_RELATIONSHIP,
        }
    }

    /// Relationship used from the data part to its properties part.
    #[must_use]
    pub const fn props_relationship(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_PROPS_RELATIONSHIP,
            Self::Strict => STRICT_PROPS_RELATIONSHIP,
        }
    }
}

/// Contents of a Custom XML Data Storage Properties part (ISO/IEC 29500 §22.5).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Props {
    /// Globally unique `ST_Guid` associated with the data part.
    pub id: String,
    /// Target namespaces of associated schemas; these URIs are never resolved.
    pub schemas: Vec<String>,
}

/// One relationship occurrence targeting a Custom XML Data Storage part.
#[derive(Debug, PartialEq, Eq)]
pub struct Item {
    /// Part that owns the data relationship.
    source: PackURI,
    /// Relationship ID on [`Self::source`].
    rel_id: String,
    /// Canonical package name of the data part.
    part: PackURI,
    /// Declared XML-based content type.
    content_type: String,
    /// Expanded name of the payload document element.
    root: Name,
    /// Exact, uninterpreted payload bytes.
    xml: Arc<Vec<u8>>,
    /// Canonical package name of the optional properties part.
    props_part: Option<PackURI>,
    /// Parsed optional properties.
    props: Option<Props>,
}

impl Item {
    pub(super) fn new(
        source: PackURI,
        rel_id: String,
        part: PackURI,
        content_type: String,
        root: Name,
        xml: Arc<Vec<u8>>,
        props_part: Option<PackURI>,
        props: Option<Props>,
    ) -> Self {
        Self {
            source,
            rel_id,
            part,
            content_type,
            root,
            xml,
            props_part,
            props,
        }
    }

    /// Part that owns this data relationship.
    #[must_use]
    pub fn source(&self) -> &PackURI {
        &self.source
    }

    /// Relationship ID on [`Self::source`].
    #[must_use]
    pub fn rel_id(&self) -> &str {
        &self.rel_id
    }

    /// Canonical package name of the data part.
    #[must_use]
    pub fn part(&self) -> &PackURI {
        &self.part
    }

    /// Declared XML-based content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Expanded name of the payload document element.
    #[must_use]
    pub fn root(&self) -> &Name {
        &self.root
    }

    /// Exact, uninterpreted payload bytes borrowed from shared OPC storage.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Canonical package name of the optional properties part.
    #[must_use]
    pub fn props_part(&self) -> Option<&PackURI> {
        self.props_part.as_ref()
    }

    /// Parsed optional properties.
    #[must_use]
    pub fn props(&self) -> Option<&Props> {
        self.props.as_ref()
    }
}

/// Properties authoring request.
///
/// Grouping these fields makes an incomplete properties request
/// unrepresentable: the part name, relationship ID, and value are all present
/// or all absent through [`Option`].
#[derive(Debug, PartialEq, Eq)]
pub struct NewProps {
    /// Package name for the new properties part.
    pub part: PackURI,
    /// Relationship ID from the new data part to the properties part.
    pub rel_id: String,
    /// Typed properties value to serialize.
    pub value: Props,
}

/// Deterministic package authoring request.
#[derive(Debug, PartialEq, Eq)]
pub struct NewItem {
    /// Existing part that will own the new data relationship.
    pub source: PackURI,
    /// Relationship ID from [`Self::source`] to [`Self::part`].
    pub rel_id: String,
    /// Package name for the new data part.
    pub part: PackURI,
    /// XML-based content type for the payload.
    pub content_type: String,
    /// Exact payload bytes to store after validation.
    pub xml: Vec<u8>,
    /// Optional, type-safe properties authoring request.
    pub props: Option<NewProps>,
    /// Namespace and relationship family to author.
    pub conformance: Conformance,
}
