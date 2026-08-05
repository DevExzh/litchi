//! Chart-part relationship and inert external-resource vocabulary.

/// Storage target for a chart's external-data relationship.
#[derive(Debug, Clone)]
pub enum ExternalDataTarget {
    /// A part embedded in the containing OOXML package
    Embedded {
        /// Complete bytes of the embedded object
        data: Vec<u8>,
        /// OPC content type for the embedded part
        content_type: String,
        /// Filename extension without a leading dot
        extension: String,
    },
    /// An externally linked object
    Linked {
        /// External relationship target
        target: String,
    },
}

/// Package payload and relationship type for chart external data.
#[derive(Debug, Clone)]
pub struct ExternalDataPart {
    /// Relationship type, normally package or OLE object
    pub relationship_type: String,
    /// Embedded or linked relationship target
    pub target: ExternalDataTarget,
}

/// Target of a relationship owned by a chart or chart user-shapes part.
#[derive(Debug, Clone)]
pub enum RelationshipTarget {
    /// A directly related part embedded in the containing package
    Embedded {
        /// Complete target-part bytes
        data: Vec<u8>,
        /// OPC content type for the target part
        content_type: String,
        /// Filename extension without a leading dot
        extension: String,
    },
    /// An external relationship target
    External {
        /// External target URI
        target: String,
    },
}

/// One relationship owned by a chart or chart user-shapes part.
#[derive(Debug, Clone)]
pub struct Relationship {
    /// Relationship identifier referenced by the owning part's XML
    pub relationship_id: String,
    /// Relationship type URI
    pub relationship_type: String,
    /// Internal payload or external target
    pub target: RelationshipTarget,
}

/// Lossless chart user-shapes XML and its direct relationship targets.
#[derive(Debug, Clone)]
pub struct UserShapesPart {
    /// Complete chart user-shapes XML document
    pub xml: Vec<u8>,
    /// Relationships owned by the chart user-shapes part
    pub relationships: Vec<Relationship>,
}

impl UserShapesPart {
    /// Create a relationship-free user-shapes drawing part.
    pub fn new(xml: Vec<u8>) -> Self {
        Self {
            xml,
            relationships: Vec::new(),
        }
    }
}

impl ExternalDataPart {
    /// Create an embedded OOXML spreadsheet payload.
    pub fn embedded_workbook(data: Vec<u8>) -> Self {
        Self {
            relationship_type: litchi_opc::constants::relationship_type::PACKAGE.to_string(),
            target: ExternalDataTarget::Embedded {
                data,
                content_type: litchi_opc::constants::content_type::OFC_PACKAGE.to_string(),
                extension: "xlsx".to_string(),
            },
        }
    }

    /// Create a linked OOXML package relationship.
    pub fn linked_package(target: impl Into<String>) -> Self {
        Self {
            relationship_type: litchi_opc::constants::relationship_type::PACKAGE.to_string(),
            target: ExternalDataTarget::Linked {
                target: target.into(),
            },
        }
    }
}

pub(crate) fn is_chart_external_data_relationship_type(relationship_type: &str) -> bool {
    chart_external_data_content_type(relationship_type).is_some()
}

pub(crate) fn chart_external_data_content_type(relationship_type: &str) -> Option<&'static str> {
    match relationship_type {
        litchi_opc::constants::relationship_type::PACKAGE
        | "http://purl.oclc.org/ooxml/officeDocument/relationships/package" => {
            Some(litchi_opc::constants::content_type::OFC_PACKAGE)
        },
        litchi_opc::constants::relationship_type::OLE_OBJECT
        | "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject" => {
            Some(litchi_opc::constants::content_type::OFC_OLE_OBJECT)
        },
        _ => None,
    }
}

pub(crate) fn is_chart_user_shapes_relationship_type(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        litchi_opc::constants::relationship_type::CHART_USER_SHAPES
            | "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes"
    )
}
