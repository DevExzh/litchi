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

    /// An OPC part URI is invalid.
    #[error("invalid DOCX part URI: {0}")]
    Uri(String),

    /// Markup-compatibility preprocessing failed.
    #[error("DOCX markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::MceError),
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
