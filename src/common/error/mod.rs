/// Unified error types for Litchi library.
///
/// This module provides a unified error type that encompasses errors from both
/// OLE2 and OOXML parsing, presenting a consistent API to users.
use thiserror::Error;

/// Main error type for Litchi operations.
#[derive(Error, Debug)]
pub enum Error {
    /// IO error
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// Parse error occurred
    #[error("Parse error: {0}")]
    ParseError(String),

    /// Invalid file format
    #[error("Invalid format: {0}")]
    InvalidFormat(String),

    /// File is not a recognized Office format
    #[error("Not a valid Office file")]
    NotOfficeFile,

    /// Corrupted or malformed file
    #[error("Corrupted file: {0}")]
    CorruptedFile(String),

    /// Stream or part not found
    #[error("Component not found: {0}")]
    ComponentNotFound(String),

    /// XML parsing error
    #[error("XML error: {0}")]
    XmlError(String),

    /// Invalid content type
    #[error("Invalid content type: expected {expected}, got {got}")]
    InvalidContentType { expected: String, got: String },

    /// ZIP archive error
    #[error("ZIP error: {0}")]
    ZipError(String),

    /// Unsupported feature
    #[error("Unsupported feature: {0}")]
    Unsupported(String),

    /// Generic error
    #[error("{0}")]
    Other(String),
}

/// Result type for Litchi operations.
pub type Result<T> = std::result::Result<T, Error>;

// Conversions from internal error types
impl From<crate::ole::OleError> for Error {
    fn from(err: crate::ole::OleError) -> Self {
        match err {
            crate::ole::OleError::Io(e) => Error::Io(e),
            crate::ole::OleError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ole::OleError::NotOleFile => Error::NotOfficeFile,
            crate::ole::OleError::CorruptedFile(s) => Error::CorruptedFile(s),
            crate::ole::OleError::StreamNotFound => Error::ComponentNotFound("Stream not found".to_string()),
        }
    }
}

impl From<crate::ole::doc::package::DocError> for Error {
    fn from(err: crate::ole::doc::package::DocError) -> Self {
        match err {
            crate::ole::doc::package::DocError::Io(e) => Error::Io(e),
            crate::ole::doc::package::DocError::Ole(ole_err) => Error::from(ole_err),
            crate::ole::doc::package::DocError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ole::doc::package::DocError::StreamNotFound(s) => Error::ComponentNotFound(s),
            crate::ole::doc::package::DocError::Corrupted(s) => Error::CorruptedFile(s),
        }
    }
}

impl From<crate::ole::ppt::package::PptError> for Error {
    fn from(err: crate::ole::ppt::package::PptError) -> Self {
        match err {
            crate::ole::ppt::package::PptError::Io(e) => Error::Io(e),
            crate::ole::ppt::package::PptError::Ole(ole_err) => Error::from(ole_err),
            crate::ole::ppt::package::PptError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ole::ppt::package::PptError::StreamNotFound(s) => Error::ComponentNotFound(s),
            crate::ole::ppt::package::PptError::Corrupted(s) => Error::CorruptedFile(s),
        }
    }
}

impl From<crate::ooxml::opc::error::OpcError> for Error {
    fn from(err: crate::ooxml::opc::error::OpcError) -> Self {
        Error::from_opc_error(err)
    }
}

impl From<crate::ooxml::error::OoxmlError> for Error {
    fn from(err: crate::ooxml::error::OoxmlError) -> Self {
        match err {
            crate::ooxml::error::OoxmlError::Io(e) => Error::Io(e),
            crate::ooxml::error::OoxmlError::Xml(s) => Error::XmlError(s),
            crate::ooxml::error::OoxmlError::PartNotFound(s) => Error::ComponentNotFound(s),
            crate::ooxml::error::OoxmlError::InvalidContentType { expected, got } => {
                Error::InvalidContentType { expected, got }
            }
            crate::ooxml::error::OoxmlError::InvalidRelationship(s) => Error::Other(s),
            crate::ooxml::error::OoxmlError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ooxml::error::OoxmlError::Opc(e) => Error::from_opc_error(e),
            crate::ooxml::error::OoxmlError::Other(s) => Error::Other(s),
        }
    }
}

impl Error {
    fn from_opc_error(err: crate::ooxml::opc::error::OpcError) -> Self {
        match err {
            crate::ooxml::opc::error::OpcError::IoError(e) => Error::Io(e),
            crate::ooxml::opc::error::OpcError::ZipError(e) => Error::ZipError(e.to_string()),
            crate::ooxml::opc::error::OpcError::XmlError(s) => Error::XmlError(s),
            crate::ooxml::opc::error::OpcError::PartNotFound(s) => Error::ComponentNotFound(s),
            _ => Error::Other(err.to_string()),
        }
    }
}

impl From<quick_xml::Error> for Error {
    fn from(err: quick_xml::Error) -> Self {
        Error::XmlError(err.to_string())
    }
}

impl From<zip::result::ZipError> for Error {
    fn from(err: zip::result::ZipError) -> Self {
        Error::ZipError(err.to_string())
    }
}

