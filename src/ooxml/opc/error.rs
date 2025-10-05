/// Error types for OPC package operations
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
    ZipError(#[from] zip::result::ZipError),

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

impl From<quick_xml::events::attributes::AttrError> for OpcError {
    fn from(err: quick_xml::events::attributes::AttrError) -> Self {
        OpcError::AttrError(err.to_string())
    }
}

pub type Result<T> = std::result::Result<T, OpcError>;

