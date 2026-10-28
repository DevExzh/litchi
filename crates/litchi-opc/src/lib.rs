//! Open Packaging Conventions (OPC) implementation.
//!
//! This module provides a complete implementation of the OPC specification,
//! which defines the structure and packaging format for Office Open XML documents.
//! It includes support for:
//!
//! - Package structure (parts, relationships)
//! - Content type management
//! - ZIP-based physical packaging
//! - Efficient parsing and minimal memory allocation
//!
//! # Performance Features
//!
//! - Uses `memchr` for fast string searching in XML
//! - Uses `atoi_simd` for fast integer parsing
//! - Uses `quick-xml` for efficient zero-copy XML parsing
//! - Minimizes allocations by borrowing data where possible
//! - Uses hash maps for O(1) lookups
//!
//! # Bounded ingestion
//!
//! Every package ingress takes a caller-selected [`SourceLimits`] profile.
//! Build a checked, tighter profile with [`SourceLimits::builder`]. The policy
//! bounds package input; ZIP members,
//! names, metadata, and declared sizes; materialized OPC parts;
//! `[Content_Types].xml`; and relationship XML, attributes, targets, events,
//! depth, and graph traversal. It is a Litchi safety policy, not a set of
//! specification maxima, and provides a defensive-consumer boundary for
//! ECMA-376 Part 2 sections 7.3.6 and 10 and MS-OI29500 sections 2.1.1749-1752.
//!
//! Macros, VBA, `ActiveX`, controls, OLE objects, and embedded code are retained
//! only as inert blobs when exposed or preserved. This crate never executes or
//! activates them.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; see
// `xml_attributes` and the workspace `clippy.toml` (record 0770).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

pub mod atomic;

pub mod accounting;
pub mod constants;
pub mod content_type;
pub mod error;
pub mod execution;
pub mod leniency;
pub mod limits;
pub mod members;
pub mod package;
pub mod packuri;
pub mod part;
mod payload;
pub mod phys_pkg;
pub mod pkgreader;
pub mod pkgwriter;
pub mod prepared;
pub mod rel;
#[cfg(feature = "sign")]
pub mod sign;
pub mod source_backed;
/// Read-only, bounded OPC package validation.
pub mod validation;
pub mod xml_attributes;
pub mod xml_splice;

// Re-export commonly used types
pub use accounting::OpcOperationAccounting;
pub use content_type::ContentType;
pub use error::{OpcError, Result, SpliceResource};
pub use execution::OpenSession;
pub use limits::{SourceLimitKind, SourceLimits, SourceLimitsBuilder};
pub use members::{NonPartMember, NonPartReason};
pub use package::{
    CanonicalRelationshipsPlan, CompressedPartTransfer, ContentTypeEdit, ContentTypeFragment,
    ContentTypeFragmentEdit, ContentTypeFragmentKey, ContentTypeFragmentPlan, ContentTypeFragments,
    ContentTypesEditPlan, FontEmbedding, OpcPackage, OwnedContentTypes, OwnedRelationships,
    PackageIdentity, RelationshipEdit, RelationshipFragment, RelationshipFragmentEdit,
    RelationshipFragmentPlan, RelationshipFragments, RelationshipSourcePlan, RelationshipsEditPlan,
    RetainedArchiveMetrics, SaveOptions,
};
pub use packuri::PackURI;
pub use part::{BlobPart, Part, PartMetadata, XmlPart};
pub use pkgreader::{PackageCatalog, probe_package_catalog_from_reader};
pub use pkgwriter::PackageWriter;
pub use prepared::PreparedOoxmlSource;
pub use rel::{Relationship, Relationships, TargetMode};
pub use source_backed::{
    ArtifactStream, ArtifactStreamError, ArtifactStreamGuard, ArtifactStreamPolicy,
    ArtifactStreamReport, ArtifactStreamTarget, AuthorizedPrecompressedPart, EffectiveContentType,
    EffectivePart, EffectivePartData, EffectiveTopology, PartBatch, PartData, PartReadSession,
    PartView, PreparedMetadataData, PreparedPartData, PreparedPartView, PreparedTopology,
    ReplacementSpan, RetainedPrecompressedPart, SourceArtifact, SourceArtifactFingerprint,
    SourceArtifactRestoreProof, SourceBackedPackage, SourceCacheCounterDelta,
    SourceCacheDiagnostics, SourceCacheDiagnosticsError, SourceCacheLimitError, SourceCacheLimits,
    SourceLineage, SourcePartSpliceFragment, SourcePartSpliceLimits, SourcePartSplicePlan,
    SourcePartSpliceProof, SourcePartSplicePublication, SourcePartSpliceReplay,
    SourcePartSpliceReplayError, SourcePartSpliceReplayHandle, SourcePartSpliceReplayProof,
    SourceReadDiagnostics, SourceReadPolicy, SourceReadPolicyError, SourceRelationshipTarget,
    SourceTopologyPlan, VerifiedDecodedReaderError,
};
#[cfg(feature = "performance-diagnostics")]
pub use source_backed::{CacheState, DiagnosticSnapshot, FlightState, Operation};
pub use validation::{validate_parts, validate_read_at};
pub use xml_splice::{
    AuthoredXmlFragment, OwnedAttributeUpdate, OwnedChildElement, OwnedElementEdit,
    OwnedElementUpdate, OwnedXmlPart, SourceXmlPart, XmlSourceRange, XmlSplicePublication,
    authored_xml_requires_source_proof, validate_source_xml_bytes,
};
