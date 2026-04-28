//! Error types for OPC package operations
use thiserror::Error;

#[derive(Error, Debug)]
pub enum OpcError {
    #[error("Package not found: {0}")]
    PackageNotFound(String),

    #[error("Invalid pack URI: {0}")]
    InvalidPackUri(String),

    #[error("Part not found: {0}")]
    PartNotFound(String),

    #[error("Relationship not found: {0}")]
    RelationshipNotFound(String),

    #[error("Content type not found for partname: {0}")]
    ContentTypeNotFound(String),

    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    #[error("XML parsing error: {0}")]
    XmlError(String),

    #[error("ZIP error: {0}")]
    ZipError(String),

    #[error("IO error: {0}")]
    IoError(#[from] std::io::Error),

    #[error("Quick-XML error: {0}")]
    QuickXmlError(#[from] quick_xml::Error),

    #[error("UTF-8 conversion error: {0}")]
    Utf8Error(#[from] std::str::Utf8Error),

    #[error("Integer parse error: {0}")]
    ParseIntError(#[from] std::num::ParseIntError),

    #[error("Attribute error: {0}")]
    AttrError(String),
}

impl From<soapberry_zip::Error> for OpcError {
    fn from(err: soapberry_zip::Error) -> Self {
        OpcError::ZipError(err.to_string())
    }
}

impl From<quick_xml::events::attributes::AttrError> for OpcError {
    fn from(err: quick_xml::events::attributes::AttrError) -> Self {
        OpcError::AttrError(err.to_string())
    }
}

pub type Result<T> = std::result::Result<T, OpcError>;

// ---------------------------------------------------------------------------
// Bridge to the umbrella's unified `litchi_core::Error` type.
//
// This impl previously lived in `src/error_ext.rs` (umbrella crate). After the
// litchi-opc carve-out (P3b), both `OpcError` and `litchi_core::Error` are
// external to the umbrella crate, so the orphan rule (E0117) forbids the impl
// at that location. We therefore relocate it here, where the target type's
// crate (`litchi-core`) is a direct dependency of `litchi-opc`. The mapping
// body is preserved verbatim from the original `from_opc_error` helper in
// src/error_ext.rs lines 57-65.
impl From<OpcError> for litchi_core::Error {
    fn from(err: OpcError) -> Self {
        match err {
            OpcError::IoError(e) => litchi_core::Error::Io(e),
            OpcError::ZipError(e) => litchi_core::Error::ZipError(e.to_string()),
            OpcError::XmlError(s) => litchi_core::Error::XmlError(s),
            OpcError::PartNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            _ => litchi_core::Error::Other(err.to_string()),
        }
    }
}
