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

    /// A requested OPC part is absent.
    #[error("DOCX part not found: {0}")]
    PartNotFound(String),

    /// XML syntax or encoding is invalid.
    #[error("invalid DOCX XML: {0}")]
    Xml(String),

    /// A part has a content type forbidden by the `WordprocessingML` relation.
    #[error("invalid DOCX content type: expected {expected}, got {actual}")]
    ContentType { expected: String, actual: String },

    /// A package part has an invalid content type for its relationship.
    #[error("invalid DOCX content type: expected {expected}, got {got}")]
    InvalidContentType { expected: String, got: String },

    /// A package relationship is malformed or inconsistent.
    #[error("invalid DOCX relationship: {0}")]
    InvalidRelationship(String),

    /// Parsed or requested data violates a `WordprocessingML` invariant.
    #[error("invalid DOCX data: {0}")]
    Invalid(String),

    /// A host-level DOCX operation rejected the package as invalid.
    #[error("invalid DOCX format: {0}")]
    InvalidFormat(String),

    /// `DrawingML` parsing or authoring failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),

    /// Bounded, inert VBA parsing or authoring failed.
    #[cfg(feature = "vba-inspection")]
    #[error("VBA error: {0}")]
    Vba(#[from] litchi_vba::Error),

    /// Runtime-neutral OOXML encryption failed.
    #[cfg(feature = "encryption")]
    #[error("OOXML encryption error: {0}")]
    Encryption(#[from] litchi_crypto::ooxml::Error),

    /// A shared host-neutral OOXML service failed.
    #[error("shared OOXML error: {0}")]
    Common(#[from] litchi_ooxml_common::Error),

    /// File or stream I/O failed.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// A package edit cannot preserve the opened artifact safely.
    #[error("unsafe DOCX edit '{operation}' rejected: {reason}")]
    UnsafeEdit {
        /// Format whose package graph would be damaged.
        format: &'static str,
        /// Operation that was rejected.
        operation: &'static str,
        /// Reason the operation cannot be lossless.
        reason: &'static str,
    },

    /// A generic bounded DOCX operation failed.
    #[error("{0}")]
    Other(String),

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

    /// An invalid URI was supplied to a package operation.
    #[error("invalid DOCX URI: {0}")]
    InvalidUri(String),

    /// Markup-compatibility preprocessing failed.
    #[error("DOCX markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::mce::Error),

    /// A bounded authoring operation could not reserve its planned buffer.
    #[error("DOCX allocation failed for {resource}: {source}")]
    Allocation {
        /// Buffer or collection being reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },

    /// A bounded external hyperlink-wrapper detachment ceiling was exceeded.
    #[error("DOCX external hyperlink detachment {resource} limit exceeded: {actual} > {maximum}")]
    ExternalHyperlinkDetachmentLimit {
        /// Bounded resource.
        resource: &'static str,
        /// Maximum accepted value.
        maximum: usize,
        /// Observed value.
        actual: usize,
    },

    /// A bounded main-document section inventory ceiling was exceeded.
    #[error("DOCX section inventory {resource} limit exceeded: {actual} > {maximum}")]
    SectionInventoryLimit {
        /// Bounded resource.
        resource: &'static str,
        /// Maximum accepted value.
        maximum: usize,
        /// Observed value.
        actual: usize,
    },

    /// A source-bound external hyperlink patch was applied to a foreign closure.
    #[error("DOCX external hyperlink detachment patch conflicts with the supplied source")]
    ExternalHyperlinkDetachmentConflict,

    /// A source-backed document-variable patch does not match its exact source.
    #[error("DOCX document-variable patch conflicts with the supplied source")]
    DocumentVariablesConflict,

    /// A changed document-variable publication cannot preserve selected markup.
    #[error("unsafe DOCX document-variable edit rejected: {0}")]
    DocumentVariablesPreservation(&'static str),
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

impl From<litchi_ooxml_common::XmlError> for Error {
    fn from(error: litchi_ooxml_common::XmlError) -> Self {
        Self::Common(litchi_ooxml_common::Error::Decode(error))
    }
}
