/// Error types for OOXML operations.
use thiserror::Error;

/// Result type for OOXML operations.
pub type Result<T> = std::result::Result<T, OoxmlError>;

/// Error types for OOXML operations.
#[derive(Error, Debug)]
#[non_exhaustive]
pub enum OoxmlError {
    /// OPC package error
    #[error("OPC error: {0}")]
    Opc(#[from] litchi_opc::error::OpcError),

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

    /// Shared DrawingML parsing error.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::DrawingError),

    #[error("markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::MceError),

    /// IO error
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// IO error (alternative form for compatibility)
    #[error("IO error: {0}")]
    IoError(std::io::Error),

    /// Invalid URI
    #[error("Invalid URI: {0}")]
    InvalidUri(String),

    /// The requested legacy mutation path cannot preserve an opened artifact.
    #[error("unsafe {format} edit '{operation}' rejected: {reason}")]
    UnsafeEdit {
        format: &'static str,
        operation: &'static str,
        reason: &'static str,
    },

    /// Generic error
    #[error("{0}")]
    Other(String),
}

impl From<quick_xml::Error> for OoxmlError {
    fn from(err: quick_xml::Error) -> Self {
        OoxmlError::Xml(err.to_string())
    }
}

impl From<std::fmt::Error> for OoxmlError {
    fn from(err: std::fmt::Error) -> Self {
        OoxmlError::Other(err.to_string())
    }
}

// `From<OoxmlError> for litchi_core::Error` lives here (not in the umbrella)
// so the orphan rule is satisfied — both source and target crates are
// external to the umbrella. Lets `?` propagate `OoxmlError` across the
// umbrella seam.
impl From<OoxmlError> for litchi_core::Error {
    fn from(err: OoxmlError) -> Self {
        match err {
            OoxmlError::Io(e) => litchi_core::Error::Io(e),
            OoxmlError::Xml(s) => litchi_core::Error::XmlError(s),
            OoxmlError::PartNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            OoxmlError::InvalidContentType { expected, got } => {
                litchi_core::Error::InvalidContentType { expected, got }
            },
            OoxmlError::InvalidRelationship(s) => litchi_core::Error::Other(s),
            OoxmlError::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            OoxmlError::Drawing(e) => litchi_core::Error::InvalidFormat(e.to_string()),
            OoxmlError::MarkupCompatibility(e) => litchi_core::Error::InvalidFormat(e.to_string()),
            OoxmlError::Opc(e) => litchi_core::Error::from(e),
            OoxmlError::IoError(e) => litchi_core::Error::Io(e),
            OoxmlError::InvalidUri(s) => litchi_core::Error::Other(s),
            OoxmlError::UnsafeEdit {
                format,
                operation,
                reason,
            } => litchi_core::Error::Unsupported(format!(
                "safe {format} operation '{operation}' is unavailable: {reason}"
            )),
            OoxmlError::Other(s) => litchi_core::Error::Other(s),
        }
    }
}
