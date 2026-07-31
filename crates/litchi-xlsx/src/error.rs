//! Typed XLSX failures.

use thiserror::Error;

/// Result of an XLSX operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to open, inspect, or edit an XLSX document.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The OPC package is malformed or inaccessible.
    #[error("OPC package error: {0}")]
    Package(#[from] litchi_opc::OpcError),
    /// Markup-compatibility preprocessing failed.
    #[error("markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::MceError),
    /// Shared OOXML decoding failed.
    #[error("OOXML decoding error: {0}")]
    Xml(#[from] litchi_ooxml_common::XmlError),
    /// An XLSX structural invariant is invalid.
    #[error("invalid XLSX structure: {0}")]
    Invalid(String),
    /// A requested name matches more than one sheet.
    #[error("sheet name '{name}' is ambiguous ({matches} matches)")]
    AmbiguousSheetName { name: String, matches: usize },
    /// A selector variant is not supported by this API version.
    #[error("unsupported sheet selector")]
    UnsupportedSelector,
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
