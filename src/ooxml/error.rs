/// Error types for OOXML operations.
use thiserror::Error;

/// Result type for OOXML operations.
pub type Result<T> = std::result::Result<T, OoxmlError>;

/// Error types for OOXML operations.
#[derive(Error, Debug)]
pub enum OoxmlError {
    /// OPC package error
    #[error("OPC error: {0}")]
    Opc(#[from] crate::ooxml::opc::error::OpcError),

    /// XML parsing error
    #[error("XML error: {0}")]
    Xml(String),

    /// Part not found
    #[error("Part not found: {0}")]
    PartNotFound(String),

    /// Invalid content type
    #[error("Invalid content type: expected {expected}, got {got}")]
    InvalidContentType { expected: String, got: String },

    /// Invalid relationship
    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    /// Invalid format
    #[error("Invalid format: {0}")]
    InvalidFormat(String),

    /// IO error
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// Generic error
    #[error("{0}")]
    Other(String),
}

impl From<quick_xml::Error> for OoxmlError {
    fn from(err: quick_xml::Error) -> Self {
        OoxmlError::Xml(err.to_string())
    }
}
