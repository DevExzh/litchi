//! Error types for OPC package operations
use std::collections::TryReserveError;

use crate::ReadResource;
use thiserror::Error;

#[derive(Error, Debug)]
#[non_exhaustive]
#[allow(
    clippy::module_name_repetitions,
    reason = "the error type is named for the OPC domain it covers; renaming the public `OpcError` type would break the crate's public API"
)]
pub enum OpcError {
    #[error("invalid OPC read limit for {resource}: {value}")]
    InvalidReadLimit {
        /// Limit that was invalid.
        resource: ReadResource,
        /// Invalid configured value.
        value: u64,
    },

    #[error("OPC read limit exceeded for {resource}: {actual} > {maximum}")]
    ReadLimit {
        /// Resource whose configured maximum was exceeded.
        resource: ReadResource,
        /// Observed resource use.
        actual: u64,
        /// Configured maximum.
        maximum: u64,
    },

    #[error("Package not found: {0}")]
    PackageNotFound(String),

    #[error("Invalid pack URI: {0}")]
    InvalidPackUri(String),

    #[error("Part not found: {0}")]
    PartNotFound(String),

    #[error("Duplicate OPC part name: {0}")]
    DuplicatePartName(String),

    #[error("ASCII-case-equivalent OPC part names coexist: '{existing}' and '{candidate}'")]
    EquivalentPartNames { existing: String, candidate: String },

    #[error("Derived OPC part names coexist: '{existing}' and '{candidate}'")]
    DerivedPartNames { existing: String, candidate: String },

    #[error("Relationship not found: {0}")]
    RelationshipNotFound(String),

    #[error("Content type not found for partname: {0}")]
    ContentTypeNotFound(String),

    #[error("Invalid content type '{value}': {reason}")]
    InvalidContentType { value: String, reason: String },

    #[error("Invalid [Content_Types].xml manifest: {0}")]
    InvalidContentTypesManifest(String),

    #[error("Duplicate default content type mapping for extension: {0}")]
    DuplicateContentTypeDefault(String),

    #[error(
        "Duplicate or ASCII-case-equivalent content type overrides: '{existing}' and '{candidate}'"
    )]
    DuplicateContentTypeOverride { existing: String, candidate: String },

    #[error("Invalid content type extension: {0}")]
    InvalidContentTypeExtension(String),

    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    #[error("Invalid relationships manifest: {0}")]
    InvalidRelationshipsManifest(String),

    #[error("Duplicate relationship ID: {0}")]
    DuplicateRelationshipId(String),

    #[error("Invalid relationship TargetMode: {0}")]
    InvalidRelationshipTargetMode(String),

    #[error("A relationships part cannot be a relationship source: {0}")]
    RelationshipPartCannotBeSource(String),

    #[error("A package cannot contain more than one core-properties relationship")]
    MultipleCorePropertiesRelationships,

    #[error("XML parsing error: {0}")]
    XmlError(String),

    /// Authored or changed XML did not meet the package publication contract.
    #[error("XML publication rejected for '{part}': {source}")]
    XmlPublication {
        /// Package-relative part name.
        part: String,
        /// Bounded XML audit failure.
        #[source]
        source: xml_minifier::audit::Error,
    },

    #[error("ZIP error: {0}")]
    ZipError(String),

    #[error("IO error: {0}")]
    IoError(#[from] std::io::Error),

    /// The destination was atomically replaced, but its parent directory
    /// could not be synchronized. Callers must not blindly retry as if the
    /// old destination were still present.
    #[error(
        "package destination was replaced but directory durability could not be confirmed: {source}"
    )]
    Committed {
        #[source]
        source: std::io::Error,
    },

    #[error("incomplete OPC output after {written} byte(s): {source}")]
    IncompleteOutput {
        written: u64,
        #[source]
        source: Box<OpcError>,
    },

    #[error("Quick-XML error: {0}")]
    QuickXmlError(#[from] quick_xml::Error),

    #[error("UTF-8 conversion error: {0}")]
    Utf8Error(#[from] std::str::Utf8Error),

    #[error("Integer parse error: {0}")]
    ParseIntError(#[from] std::num::ParseIntError),

    #[error("Attribute error: {0}")]
    AttrError(String),

    /// A bounded package operation could not reserve its required memory.
    #[error("OPC allocation failed for {resource}: {source}")]
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },
}

impl From<soapberry_zip::Error> for OpcError {
    fn from(err: soapberry_zip::Error) -> Self {
        OpcError::ZipError(err.to_string())
    }
}

impl From<quick_xml::events::attributes::AttrError> for OpcError {
    fn from(err: quick_xml::events::attributes::AttrError) -> Self {
        OpcError::AttrError(err.to_string())
    }
}

pub type Result<T> = std::result::Result<T, OpcError>;

// `From<OpcError> for litchi_core::Error` lives here (not in the umbrella)
// so the orphan rule is satisfied — both source and target crates are
// external to the umbrella.
impl From<OpcError> for litchi_core::Error {
    fn from(err: OpcError) -> Self {
        match err {
            OpcError::IoError(e) => litchi_core::Error::Io(e),
            OpcError::ZipError(e) => litchi_core::Error::ZipError(e),
            OpcError::XmlError(s) => litchi_core::Error::XmlError(s),
            OpcError::PartNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            OpcError::Allocation { resource, source } => {
                litchi_core::Error::Allocation { resource, source }
            },
            OpcError::InvalidReadLimit { .. }
            | OpcError::ReadLimit { .. }
            | OpcError::PackageNotFound(_)
            | OpcError::InvalidPackUri(_)
            | OpcError::DuplicatePartName(_)
            | OpcError::EquivalentPartNames { .. }
            | OpcError::DerivedPartNames { .. }
            | OpcError::RelationshipNotFound(_)
            | OpcError::ContentTypeNotFound(_)
            | OpcError::InvalidContentType { .. }
            | OpcError::InvalidContentTypesManifest(_)
            | OpcError::DuplicateContentTypeDefault(_)
            | OpcError::DuplicateContentTypeOverride { .. }
            | OpcError::InvalidContentTypeExtension(_)
            | OpcError::InvalidRelationship(_)
            | OpcError::InvalidRelationshipsManifest(_)
            | OpcError::DuplicateRelationshipId(_)
            | OpcError::InvalidRelationshipTargetMode(_)
            | OpcError::RelationshipPartCannotBeSource(_)
            | OpcError::MultipleCorePropertiesRelationships
            | OpcError::XmlPublication { .. }
            | OpcError::Committed { .. }
            | OpcError::IncompleteOutput { .. }
            | OpcError::QuickXmlError(_)
            | OpcError::Utf8Error(_)
            | OpcError::ParseIntError(_)
            | OpcError::AttrError(_) => litchi_core::Error::Other(err.to_string()),
        }
    }
}
