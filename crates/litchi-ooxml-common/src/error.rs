//! Shared failures for host-neutral OOXML package services.

use std::collections::TryReserveError;

use thiserror::Error;

/// Result returned by shared OOXML package services.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to decode, validate, or traverse host-neutral OOXML package data.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying OPC graph is malformed or could not be mutated safely.
    #[error("OPC error: {0}")]
    Opc(#[source] litchi_opc::error::OpcError),

    /// A bounded durable semantic patch failed decoding or validation.
    #[error("invalid durable OOXML patch: {0}")]
    Patch(#[from] litchi_core::patch::PatchError),

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

    /// A semantic patch was built from a different selected graph base.
    #[error("OOXML patch base does not match recipient")]
    PatchConflict,

    /// A well-formed operation was refused without exposing package content.
    #[error("OOXML operation refused: {0}")]
    Refused(litchi_core::edit::Refusal),

    /// A changed package still carries signature infrastructure.
    #[error("signed OOXML package requires an explicit signature disposition")]
    Signed,

    /// A source-backed package changed while an operation was in progress.
    #[error("OOXML source changed while the operation was in progress")]
    SourceChanged {
        /// Version captured before the operation.
        expected: litchi_core::SourceVersion,
        /// Version observed when the operation completed.
        observed: litchi_core::SourceVersion,
    },

    /// A generated staging or publication invariant was not satisfied.
    #[error("OOXML internal invariant failed: {0}")]
    Internal(&'static str),

    /// Parsed data violates a structural, cardinality, or value constraint.
    #[error("invalid OOXML data: {0}")]
    Invalid(String),

    /// A `SpreadsheetML` Custom XML Maps part violates its bounded vocabulary.
    #[error("{0}")]
    SpreadsheetXmlMaps(String),

    /// A bounded resource exceeds its ceiling (record 0351 §2.4).
    #[error("OOXML {kind} exceeds limit {maximum} (actual {observed})")]
    Limit {
        /// The exceeded bound.
        kind: LimitKind,
        /// Observed value, or `usize::MAX` when arithmetic overflowed.
        observed: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A bounded inventory collection could not reserve its next entry.
    #[error("OOXML allocation failed for {resource}: {source}")]
    Allocation {
        /// Collection whose bounded reservation failed.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },

    /// A package URI is invalid for the requested operation.
    #[error("invalid OOXML URI: {0}")]
    Uri(String),

    /// A bounded inert MS-OVBA project payload failed to decode or validate.
    #[cfg(feature = "vba-inspection")]
    #[error("VBA error: {0}")]
    Vba(#[from] litchi_vba::Error),

    /// Markup-compatibility preprocessing failed.
    #[error("OOXML markup compatibility error: {0}")]
    Mce(#[from] crate::mce::Error),

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

impl Error {
    /// Creates a content-free typed refusal for a well-formed request.
    pub(crate) const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
    }
}

impl From<litchi_opc::error::OpcError> for Error {
    fn from(error: litchi_opc::error::OpcError) -> Self {
        match error {
            litchi_opc::error::OpcError::SourceChanged { expected, actual } => {
                Self::SourceChanged {
                    expected,
                    observed: actual,
                }
            },
            litchi_opc::error::OpcError::SignedSourceRequiresExplicitPolicy => Self::Signed,
            other => Self::Opc(other),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn opc_errors_retain_their_source_chain() {
        let error = Error::from(litchi_opc::error::OpcError::PartNotFound(
            "missing.xml".to_owned(),
        ));
        let source = std::error::Error::source(&error).expect("OPC source");
        assert!(
            source
                .downcast_ref::<litchi_opc::error::OpcError>()
                .is_some()
        );
    }
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}

impl From<litchi_core::metadata::keywords::Error> for Error {
    fn from(error: litchi_core::metadata::keywords::Error) -> Self {
        match error {
            litchi_core::metadata::keywords::Error::Limit { observed, maximum } => Self::Limit {
                kind: LimitKind::CorePropertyTextBytes,
                maximum,
                observed,
            },
            other => Self::Invalid(other.to_string()),
        }
    }
}

/// A bounded dimension of a shared OOXML service, as reported by
/// [`Error::Limit`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// A Custom XML budget ([`crate::custom_xml::LimitKind`]).
    CustomXml(crate::custom_xml::LimitKind),
    /// A Web Extensions (task-pane) budget ([`crate::web::LimitKind`]).
    Web(crate::web::LimitKind),
    /// A Ribbon (customUI) budget ([`crate::ribbon::LimitKind`]).
    Ribbon(crate::ribbon::LimitKind),
    /// A document-properties transaction budget
    /// ([`crate::properties::LimitKind`]).
    Properties(crate::properties::LimitKind),
    /// A VBA graph publication budget ([`crate::vba::LimitKind`]).
    #[cfg(feature = "vba-inspection")]
    Vba(crate::vba::LimitKind),
    /// Core keyword items.
    CoreKeywordItems,
    /// Core-properties part-name candidates.
    CorePropertiesPartNameCandidates,
    /// Core-properties XML bytes.
    CorePropertiesXmlBytes,
    /// Core-properties XML events.
    CorePropertiesXmlEvents,
    /// Core property text bytes.
    CorePropertyTextBytes,
    /// Custom Data storage id UTF-16 units.
    CustomDataStorageIdUtf16Units,
    /// Custom Data UID UTF-16 units.
    CustomDataUidUtf16Units,
    /// Custom Data XML namespace bytes.
    CustomDataXmlNamespaceBytes,
    /// Custom Data XML namespace prefix.
    CustomDataXmlNamespacePrefix,
    /// Custom Data XML namespace URI.
    CustomDataXmlNamespaceUri,
    /// Custom properties.
    CustomProperties,
    /// Custom-properties XML attribute bytes.
    CustomPropertiesXmlAttributeBytes,
    /// Custom-properties XML attributes.
    CustomPropertiesXmlAttributes,
    /// Custom-properties XML bytes.
    CustomPropertiesXmlBytes,
    /// Custom-properties XML depth.
    CustomPropertiesXmlDepth,
    /// Custom-properties XML nodes.
    CustomPropertiesXmlNodes,
    /// Custom-property name bytes.
    CustomPropertyNameBytes,
    /// Custom-property name characters.
    CustomPropertyNameCharacters,
    /// Custom-property text bytes.
    CustomPropertyTextBytes,
    /// Document-property part names.
    DocumentPropertyPartNames,
    /// Embedded payload relationships.
    EmbeddedPayloadRelationships,
    /// Embedded relationships.
    EmbeddedRelationships,
    /// Escaped XML attribute length.
    EscapedXmlAttributeLength,
    /// Extended-properties part names.
    ExtendedPropertiesPartNames,
    /// Extended-properties XML bytes.
    ExtendedPropertiesXmlBytes,
    /// Extended-properties XML depth.
    ExtendedPropertiesXmlDepth,
    /// Extended-properties XML events.
    ExtendedPropertiesXmlEvents,
    /// Extended-property text bytes.
    ExtendedPropertyTextBytes,
    /// Extension URI.
    ExtensionUri,
    /// Extension XML bytes.
    ExtensionXmlBytes,
    /// Package relationship identifiers.
    PackageRelationshipIdentifiers,
    /// Properties XML bytes.
    PropertiesXmlBytes,
    /// Properties XML output.
    PropertiesXmlOutput,
    /// Serialized properties XML bytes.
    SerializedPropertiesXmlBytes,
    /// Source-backed embedded scan progress.
    SourceBackedEmbeddedScanProgress,
    /// Storage id bytes.
    StorageIdBytes,
    /// XML attribute count.
    XmlAttributeCount,
    /// XML depth.
    XmlDepth,
    /// XML event count.
    XmlEventCount,
    /// XML namespace binding count.
    XmlNamespaceBindingCount,
    /// XML node count.
    XmlNodeCount,
    /// XML string bytes.
    XmlStringBytes,
}

impl LimitKind {
    /// A short structural description of the dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::CustomXml(kind) => kind.as_str(),
            Self::Web(kind) => kind.as_str(),
            Self::Ribbon(kind) => kind.as_str(),
            Self::Properties(kind) => kind.as_str(),
            #[cfg(feature = "vba-inspection")]
            Self::Vba(kind) => kind.as_str(),
            Self::CoreKeywordItems => "core keyword items",
            Self::CorePropertiesPartNameCandidates => "core-properties part-name candidates",
            Self::CorePropertiesXmlBytes => "core-properties XML bytes",
            Self::CorePropertiesXmlEvents => "core-properties XML events",
            Self::CorePropertyTextBytes => "core property text bytes",
            Self::CustomDataStorageIdUtf16Units => "Custom Data storage id UTF-16 units",
            Self::CustomDataUidUtf16Units => "Custom Data UID UTF-16 units",
            Self::CustomDataXmlNamespaceBytes => "Custom Data XML namespace bytes",
            Self::CustomDataXmlNamespacePrefix => "Custom Data XML namespace prefix",
            Self::CustomDataXmlNamespaceUri => "Custom Data XML namespace URI",
            Self::CustomProperties => "custom properties",
            Self::CustomPropertiesXmlAttributeBytes => "custom-properties XML attribute bytes",
            Self::CustomPropertiesXmlAttributes => "custom-properties XML attributes",
            Self::CustomPropertiesXmlBytes => "custom-properties XML bytes",
            Self::CustomPropertiesXmlDepth => "custom-properties XML depth",
            Self::CustomPropertiesXmlNodes => "custom-properties XML nodes",
            Self::CustomPropertyNameBytes => "custom-property name bytes",
            Self::CustomPropertyNameCharacters => "custom-property name characters",
            Self::CustomPropertyTextBytes => "custom-property text bytes",
            Self::DocumentPropertyPartNames => "document-property part names",
            Self::EmbeddedPayloadRelationships => "embedded payload relationships",
            Self::EmbeddedRelationships => "embedded relationships",
            Self::EscapedXmlAttributeLength => "escaped XML attribute length",
            Self::ExtendedPropertiesPartNames => "extended-properties part names",
            Self::ExtendedPropertiesXmlBytes => "extended-properties XML bytes",
            Self::ExtendedPropertiesXmlDepth => "extended-properties XML depth",
            Self::ExtendedPropertiesXmlEvents => "extended-properties XML events",
            Self::ExtendedPropertyTextBytes => "extended-property text bytes",
            Self::ExtensionUri => "extension URI",
            Self::ExtensionXmlBytes => "extension XML bytes",
            Self::PackageRelationshipIdentifiers => "package relationship identifiers",
            Self::PropertiesXmlBytes => "properties XML bytes",
            Self::PropertiesXmlOutput => "properties XML output",
            Self::SerializedPropertiesXmlBytes => "serialized properties XML bytes",
            Self::SourceBackedEmbeddedScanProgress => "source-backed embedded scan progress",
            Self::StorageIdBytes => "storage id bytes",
            Self::XmlAttributeCount => "XML attribute count",
            Self::XmlDepth => "XML depth",
            Self::XmlEventCount => "XML event count",
            Self::XmlNamespaceBindingCount => "XML namespace binding count",
            Self::XmlNodeCount => "XML node count",
            Self::XmlStringBytes => "XML string bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl From<crate::custom_xml::LimitKind> for LimitKind {
    fn from(kind: crate::custom_xml::LimitKind) -> Self {
        Self::CustomXml(kind)
    }
}

impl From<crate::web::LimitKind> for LimitKind {
    fn from(kind: crate::web::LimitKind) -> Self {
        Self::Web(kind)
    }
}

impl From<crate::ribbon::LimitKind> for LimitKind {
    fn from(kind: crate::ribbon::LimitKind) -> Self {
        Self::Ribbon(kind)
    }
}

impl From<crate::properties::LimitKind> for LimitKind {
    fn from(kind: crate::properties::LimitKind) -> Self {
        Self::Properties(kind)
    }
}
