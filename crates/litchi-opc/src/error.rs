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

    /// A decoded source-part splice exceeded one of its operation-local
    /// workspace ceilings. This is kept distinct from package input bytes so
    /// callers can report parser, preservation, and replay memory budgets
    /// accurately.
    #[error("OPC source-part splice limit exceeded for {resource}: {actual} > {maximum}")]
    SourcePartSpliceLimit {
        /// Workspace resource whose finite ceiling was exceeded.
        resource: SpliceResource,
        /// Observed or conservatively reserved bytes.
        actual: u64,
        /// Configured finite ceiling.
        maximum: u64,
    },

    /// A source-part splice workspace ceiling was configured to an invalid
    /// value. This remains distinct from the package input-byte limits so a
    /// caller can identify the rejected operation-local quota.
    #[error("invalid OPC source-part splice limit for {resource}: {value}")]
    InvalidSourcePartSpliceLimit {
        /// Workspace resource whose configured ceiling was invalid.
        resource: SpliceResource,
        /// Invalid configured value.
        value: u64,
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

    #[error(
        "Duplicate or ASCII-case-equivalent reserved content-types members: '{existing}' and '{candidate}'"
    )]
    DuplicateContentTypesMember { existing: String, candidate: String },

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

    /// An explicitly requested OPC open was cooperatively cancelled.
    #[error("OPC open cancelled")]
    Cancelled,

    /// An explicit execution context rejected a policy, cancellation check, or
    /// hierarchical resource charge.
    #[error("OPC execution failed: {0}")]
    Execution(#[source] litchi_core::ExecutionError),

    /// The archive-local scheduler used by an explicit open session failed.
    #[error("OPC local parallel read failed: {0}")]
    ParallelRead(#[source] soapberry_zip::Error),

    /// The requested neutral affinity policy has no archive-local adapter.
    #[error("OPC open does not support the requested worker-affinity policy")]
    UnsupportedExecutionAffinity,

    #[error("IO error: {0}")]
    IoError(#[from] std::io::Error),

    /// A caller-provided central-directory spool failed during an explicit
    /// storage operation. The source remains available to callers that need
    /// the provider's underlying diagnostic.
    #[error("OPC central-directory spool {operation} failed")]
    CentralDirectorySpool {
        /// Content-free operation performed on the caller-provided store.
        operation: &'static str,
        /// Original provider error.
        #[source]
        source: std::io::Error,
    },

    /// The caller-provided central-directory spool exceeded its explicit
    /// serialized-byte ceiling.
    #[error("OPC central-directory spool limit exceeded: {actual} > {maximum}")]
    CentralDirectorySpoolLimitExceeded { actual: u64, maximum: u64 },

    /// A positional source changed after this package captured its snapshot.
    #[error("OPC source changed from {expected:?} to {actual:?}")]
    SourceChanged {
        /// Version captured at package open.
        expected: litchi_core::SourceVersion,
        /// Version observed during a later operation.
        actual: litchi_core::SourceVersion,
    },

    /// A source-backed bounded-Part publisher cannot prove that the physical
    /// source can be preserved without a full materializing rewrite.
    #[error("source-backed OPC Part overlay is unavailable: {reason}")]
    SourceBackedOverlayUnavailable {
        /// Content-free reason the conservative raw-copy path refused.
        reason: String,
    },

    /// A caller-owned OPC operation report could not represent another
    /// accepted or observed byte count without overflowing.
    #[error("OPC operation accounting overflow: {counter}")]
    OperationAccountingOverflow {
        /// Counter whose checked merge or update overflowed.
        counter: &'static str,
    },

    /// A source-backed Part change would invalidate an existing signature
    /// without an explicit strip-or-resign policy.
    #[error("signed OPC source requires an explicit signature edit policy")]
    SignedSourceRequiresExplicitPolicy,

    /// An explicitly supplied source artifact did not match its authenticated
    /// length or digest proof.  Keeping the artifact role and field typed lets
    /// durable inverse callers distinguish current-candidate and original
    /// provider mismatches without exposing provider implementation details.
    #[error("source artifact {artifact} does not match its {field} proof")]
    SourceArtifactMismatch {
        /// Artifact role being authenticated, such as `current` or
        /// `original`.
        artifact: &'static str,
        /// Authenticated field that differed.
        field: &'static str,
    },

    /// An owned source package cannot preserve its physical ZIP layout after
    /// an exact-source authorization was revoked. Falling back to the normal
    /// writer would silently discard opaque source members or framing bytes.
    #[error("owned OPC source preservation is unavailable: {reason}")]
    PreservationUnavailable {
        /// Content-free reason the source-preserving writer refused.
        reason: String,
    },

    /// A managed source-backed payload cannot be detached from the cache's
    /// hierarchical memory reservation. Keep the returned [`crate::PartData`] handle
    /// (or borrow it with `as_bytes`) until the operation is complete.
    #[error("managed source-backed OPC PartData cannot escape its budgeted handle")]
    ManagedPartDataArcEscape,

    /// A managed source-backed package cannot be converted into an owning
    /// package without detaching materialized payloads from their budgeted
    /// cache handles. Use the source-backed view while the execution context
    /// is active, or materialize only after reopening through an unmanaged
    /// compatibility constructor.
    #[error("managed source-backed OPC package cannot be materialized into an owning package")]
    ManagedPackageMaterialization,

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

    /// A fallible inline collection could not grow.
    #[error("OPC allocation failed for {resource}")]
    CollectionAllocation {
        /// Resource whose bounded collection could not grow.
        resource: &'static str,
    },

    /// A bounded, format-neutral validation report could not be constructed.
    #[error("OPC validation report construction failed: {0}")]
    ValidationReport(#[from] litchi_core::ValidationReportError),
}

/// Operation-local resource used by a bounded decoded source-part splice.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SpliceResource {
    /// Accepted publication output bytes.
    OutputBytes,
    /// XML parser state and the fixed logical candidate adapter buffer.
    XmlWorkspaceBytes,
    /// Preservation index and retained central-directory planning state.
    PreservationMemoryBytes,
    /// Replay payload windows and compressor state.
    ReplayMemoryBytes,
    /// The bounded adapter window used to replay caller-authored bytes.
    ///
    /// This is separate from [`Self::ReplayMemoryBytes`], which remains the
    /// ZIP replay/compressor workspace.  A caller-owned replay store may have
    /// additional storage that is transferred through its replay handle; it
    /// is not charged once per reader.
    AuthoredReplayMemoryBytes,
}

impl std::fmt::Display for SpliceResource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::OutputBytes => "splice output bytes",
            Self::XmlWorkspaceBytes => "XML splice workspace bytes",
            Self::PreservationMemoryBytes => "preservation splice workspace bytes",
            Self::ReplayMemoryBytes => "replay splice workspace bytes",
            Self::AuthoredReplayMemoryBytes => "authored replay window bytes",
        })
    }
}

impl From<soapberry_zip::Error> for OpcError {
    fn from(err: soapberry_zip::Error) -> Self {
        match err.into_kind() {
            soapberry_zip::ErrorKind::Cancelled => Self::Cancelled,
            soapberry_zip::ErrorKind::IO(error) | soapberry_zip::ErrorKind::Io(error) => {
                map_io_error(error)
            },
            soapberry_zip::ErrorKind::LimitExceeded {
                resource,
                actual,
                maximum,
            } => Self::ReadLimit {
                resource: match resource {
                    soapberry_zip::LimitResource::FileCount => ReadResource::ArchiveMembers,
                    soapberry_zip::LimitResource::MemberNameBytes => {
                        ReadResource::ArchiveMemberNameBytes
                    },
                    soapberry_zip::LimitResource::MetadataBytes => {
                        ReadResource::ArchiveMetadataBytes
                    },
                    soapberry_zip::LimitResource::CompressedSize => {
                        ReadResource::ArchiveCompressedBytes
                    },
                    soapberry_zip::LimitResource::EntrySize => ReadResource::ArchiveEntryBytes,
                    soapberry_zip::LimitResource::TotalSize => ReadResource::ArchiveTotalBytes,
                },
                actual,
                maximum,
            },
            soapberry_zip::ErrorKind::Allocation { resource, source } => {
                Self::Allocation { resource, source }
            },
            soapberry_zip::ErrorKind::CentralDirectorySpool { operation, source } => {
                Self::CentralDirectorySpool { operation, source }
            },
            soapberry_zip::ErrorKind::CentralDirectorySpoolLimitExceeded { actual, maximum } => {
                Self::CentralDirectorySpoolLimitExceeded { actual, maximum }
            },
            kind => Self::ZipError(kind.to_string()),
        }
    }
}

impl From<quick_xml::events::attributes::AttrError> for OpcError {
    fn from(err: quick_xml::events::attributes::AttrError) -> Self {
        OpcError::AttrError(err.to_string())
    }
}

pub type Result<T> = std::result::Result<T, OpcError>;

#[derive(Debug)]
pub(crate) struct ExecutionIoError(pub(crate) litchi_core::ExecutionError);

impl std::fmt::Display for ExecutionIoError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Display::fmt(&self.0, formatter)
    }
}

impl std::error::Error for ExecutionIoError {}

pub(crate) fn execution_io_error(error: litchi_core::ExecutionError) -> std::io::Error {
    std::io::Error::other(ExecutionIoError(error))
}

pub(crate) fn map_io_error(error: std::io::Error) -> OpcError {
    // Format-owned cooperative adapters can carry the public execution error
    // directly through Write/Read without depending on OPC's private marker.
    if let Some(execution) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<litchi_core::ExecutionError>())
    {
        return execution_to_opc_error(execution.clone());
    }
    if let Some(execution) = error
        .get_ref()
        .and_then(|source| source.downcast_ref::<ExecutionIoError>())
    {
        return execution_to_opc_error(execution.0.clone());
    }
    if error
        .get_ref()
        .is_some_and(|source| source.is::<soapberry_zip::Error>())
    {
        let source = error
            .into_inner()
            .expect("a checked ZIP I/O error must retain its source");
        let error = source
            .downcast::<soapberry_zip::Error>()
            .expect("the checked I/O error source must remain a ZIP error");
        return OpcError::from(*error);
    }
    OpcError::IoError(error)
}

fn execution_to_opc_error(error: litchi_core::ExecutionError) -> OpcError {
    match error {
        litchi_core::ExecutionError::Cancelled => OpcError::Cancelled,
        error => OpcError::Execution(error),
    }
}

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
            OpcError::CollectionAllocation { resource } => {
                litchi_core::Error::Other(format!("allocation failed for {resource}"))
            },
            OpcError::ValidationReport(error) => litchi_core::Error::Other(error.to_string()),
            error @ (OpcError::SourceBackedOverlayUnavailable { .. }
            | OpcError::PreservationUnavailable { .. }
            | OpcError::SignedSourceRequiresExplicitPolicy
            | OpcError::SourceArtifactMismatch { .. }) => {
                litchi_core::Error::Unsupported(error.to_string())
            },
            OpcError::InvalidReadLimit { .. }
            | OpcError::ReadLimit { .. }
            | OpcError::Cancelled
            | OpcError::Execution(_)
            | OpcError::ParallelRead(_)
            | OpcError::UnsupportedExecutionAffinity
            | OpcError::ManagedPartDataArcEscape
            | OpcError::ManagedPackageMaterialization
            | OpcError::OperationAccountingOverflow { .. }
            | OpcError::CentralDirectorySpool { .. }
            | OpcError::CentralDirectorySpoolLimitExceeded { .. }
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
            | OpcError::DuplicateContentTypesMember { .. }
            | OpcError::InvalidContentTypeExtension(_)
            | OpcError::InvalidRelationship(_)
            | OpcError::InvalidRelationshipsManifest(_)
            | OpcError::DuplicateRelationshipId(_)
            | OpcError::InvalidRelationshipTargetMode(_)
            | OpcError::RelationshipPartCannotBeSource(_)
            | OpcError::MultipleCorePropertiesRelationships
            | OpcError::XmlPublication { .. }
            | OpcError::SourceChanged { .. }
            | OpcError::Committed { .. }
            | OpcError::IncompleteOutput { .. }
            | OpcError::QuickXmlError(_)
            | OpcError::Utf8Error(_)
            | OpcError::ParseIntError(_)
            | OpcError::AttrError(_) => litchi_core::Error::Other(err.to_string()),
            OpcError::SourcePartSpliceLimit {
                resource,
                actual,
                maximum,
            } => litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: match resource {
                    SpliceResource::OutputBytes => litchi_core::Resource::OutputBytes,
                    SpliceResource::XmlWorkspaceBytes
                    | SpliceResource::PreservationMemoryBytes
                    | SpliceResource::ReplayMemoryBytes
                    | SpliceResource::AuthoredReplayMemoryBytes => litchi_core::Resource::Memory,
                },
                observed: actual,
                limit: maximum,
                scope: std::sync::Arc::from(format!("OPC source-part splice ({resource})")),
            }),
            error @ OpcError::InvalidSourcePartSpliceLimit { .. } => {
                litchi_core::Error::Unsupported(error.to_string())
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{OpcError, execution_io_error};

    #[test]
    fn spool_error_display_is_content_free_and_retains_the_provider_source() {
        let error = OpcError::from(soapberry_zip::Error::from(
            soapberry_zip::ErrorKind::CentralDirectorySpool {
                operation: "read for replay",
                source: std::io::Error::other("private/provider/path"),
            },
        ));
        assert_eq!(
            error.to_string(),
            "OPC central-directory spool read for replay failed"
        );
        assert_eq!(
            std::error::Error::source(&error).unwrap().to_string(),
            "private/provider/path"
        );
    }

    #[test]
    fn publication_capability_errors_retain_typed_core_classification() {
        for error in [
            OpcError::PreservationUnavailable {
                reason: "unsupported framing".to_owned(),
            },
            OpcError::SourceBackedOverlayUnavailable {
                reason: "opaque member cannot be patched".to_owned(),
            },
            OpcError::SignedSourceRequiresExplicitPolicy,
        ] {
            assert!(matches!(
                litchi_core::Error::from(error),
                litchi_core::Error::Unsupported(_)
            ));
        }
    }

    #[test]
    fn accounting_overflow_is_not_misclassified_as_unsupported() {
        assert!(matches!(
            litchi_core::Error::from(OpcError::OperationAccountingOverflow {
                counter: "OPC output bytes accepted"
            }),
            litchi_core::Error::Other(message) if message.contains("accounting overflow")
        ));
    }

    #[test]
    fn execution_io_errors_survive_zip_error_conversion() {
        let error = soapberry_zip::Error::from(soapberry_zip::ErrorKind::IO(execution_io_error(
            litchi_core::ExecutionError::Cancelled,
        )));
        assert!(matches!(OpcError::from(error), OpcError::Cancelled));
    }

    #[test]
    fn public_execution_error_from_a_format_sink_remains_typed() {
        let error = soapberry_zip::Error::from(soapberry_zip::ErrorKind::IO(
            std::io::Error::other(litchi_core::ExecutionError::Cancelled),
        ));
        assert!(matches!(OpcError::from(error), OpcError::Cancelled));
    }

    #[test]
    fn allocation_zip_errors_remain_typed() {
        let source = Vec::<u8>::new()
            .try_reserve_exact(usize::MAX)
            .expect_err("the test allocation must fail without allocating");
        let error = soapberry_zip::Error::from(soapberry_zip::ErrorKind::Allocation {
            resource: "test allocation",
            source,
        });

        assert!(matches!(
            OpcError::from(error),
            OpcError::Allocation {
                resource: "test allocation",
                ..
            }
        ));
    }
}
