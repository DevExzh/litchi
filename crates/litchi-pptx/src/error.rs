//! Typed PPTX failures.

use thiserror::Error;

/// Result of a PPTX operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to decode or encode a PresentationML capability.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The XML stream is not well formed or cannot be decoded safely.
    #[error("invalid PresentationML XML: {0}")]
    Xml(String),

    /// The document violates a PresentationML structural or value invariant.
    #[error("invalid PresentationML: {0}")]
    Invalid(String),

    /// A bounded decoder resource was exhausted.
    #[error("PresentationML {resource} exceeds the limit of {limit}")]
    Limit {
        /// Resource that exceeded its configured limit.
        resource: &'static str,
        /// Active upper bound.
        limit: usize,
    },

    /// Markup-compatibility processing failed.
    #[error("PresentationML markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::MceError),

    /// Shared OOXML attribute decoding failed.
    #[error("PresentationML attribute decoding error: {0}")]
    Decode(#[from] litchi_ooxml_common::XmlError),

    /// Writing into the requested text sink failed.
    #[error("could not encode PresentationML text")]
    Write,
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
