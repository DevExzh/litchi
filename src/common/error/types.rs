//! Unified error types for Litchi library.
//!
//! This module provides a unified error type that encompasses errors from both
//! OLE2 and OOXML parsing, presenting a consistent API to users.
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

    /// Feature disabled at compile time
    #[error("Feature '{0}' is disabled. Enable it with --features {0}")]
    FeatureDisabled(String),

    /// Generic error
    #[error("{0}")]
    Other(String),
}

/// Result type for Litchi operations.
pub type Result<T> = std::result::Result<T, Error>;

