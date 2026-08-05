//! Shared failures for host-neutral OOXML package services.

use thiserror::Error;

/// Result returned by shared OOXML package services.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to decode, validate, or traverse host-neutral OOXML package data.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying OPC graph is malformed or could not be mutated safely.
    #[error("OPC error: {0}")]
    Opc(#[from] litchi_opc::error::OpcError),

    /// XML syntax or encoding is invalid.
    #[error("invalid OOXML XML: {0}")]
    Xml(String),

    /// A required package part is absent.
    #[error("OOXML part not found: {0}")]
    Missing(String),

    /// A package part has a different content type than its relationship requires.
    #[error("invalid OOXML content type: expected {expected}, got {actual}")]
    ContentType {
        /// Content type required by the package vocabulary.
        expected: String,
        /// Content type declared by the package part.
        actual: String,
    },

    /// A package relationship violates the shared vocabulary.
    #[error("invalid OOXML relationship: {0}")]
    Relationship(String),

    /// Parsed data violates a structural, cardinality, or value constraint.
    #[error("invalid OOXML data: {0}")]
    Invalid(String),

    /// A bounded resource exceeds the service's declared ceiling.
    #[error("OOXML {resource} exceeds limit {max} (actual {actual})")]
    Limit {
        /// Resource whose bounded representation was exceeded.
        resource: &'static str,
        /// Maximum accepted value.
        max: usize,
        /// Observed value, or `usize::MAX` when arithmetic overflowed.
        actual: usize,
    },

    /// A package URI is invalid for the requested operation.
    #[error("invalid OOXML URI: {0}")]
    Uri(String),

    /// A bounded inert MS-OVBA project payload failed to decode or validate.
    #[error("VBA error: {0}")]
    Vba(#[from] litchi_vba::Error),

    /// Markup-compatibility preprocessing failed.
    #[error("OOXML markup compatibility error: {0}")]
    Mce(#[from] crate::MceError),

    /// Shared entity, namespace, or attribute decoding failed.
    #[error("OOXML decoding error: {0}")]
    Decode(#[from] crate::XmlError),

    /// Reading or writing a caller-provided stream failed.
    #[error("OOXML I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// A formatting operation failed.
    #[error("OOXML formatting error")]
    Fmt(#[from] std::fmt::Error),
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
