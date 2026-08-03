use std::collections::TryReserveError;
use thiserror::Error;

/// Result returned by canonical DOCX operations.
pub type Result<T> = std::result::Result<T, Error>;

/// A bounded parsing, validation, or package-graph failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying OPC graph is malformed or could not be mutated safely.
    #[error("OPC error: {0}")]
    Opc(#[from] litchi_opc::error::OpcError),

    /// XML syntax or encoding is invalid.
    #[error("invalid DOCX XML: {0}")]
    Xml(String),

    /// A part has a content type forbidden by the WordprocessingML relation.
    #[error("invalid DOCX content type: expected {expected}, got {actual}")]
    ContentType { expected: String, actual: String },

    /// Parsed or requested data violates a WordprocessingML invariant.
    #[error("invalid DOCX data: {0}")]
    Invalid(String),

    /// A semantic selector matched more than one producer object.
    #[error("{object} selector '{key}' is ambiguous")]
    Ambiguous {
        /// Kind of object being selected.
        object: &'static str,
        /// User-facing semantic selector.
        key: String,
    },

    /// A checked numeric selector was outside the current collection.
    #[error("{object} index {index} is out of bounds for length {len}")]
    OutOfBounds {
        /// Kind of object being selected.
        object: &'static str,
        /// Requested zero-based index.
        index: usize,
        /// Collection length at validation time.
        len: usize,
    },

    /// An OPC part URI is invalid.
    #[error("invalid DOCX part URI: {0}")]
    Uri(String),

    /// Markup-compatibility preprocessing failed.
    #[error("DOCX markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::MceError),

    /// A bounded authoring operation could not reserve its planned buffer.
    #[error("DOCX allocation failed for {resource}: {source}")]
    Allocation {
        /// Buffer or collection being reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}

impl From<std::fmt::Error> for Error {
    fn from(error: std::fmt::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
