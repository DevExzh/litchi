//! Bounded physical-layer primitives for Apple iWork IWA components.
//!
//! This crate owns the checksum-free Snappy framing used by `.iwa` members and
//! the copyable resource budgets that higher-level archive parsers can carry
//! across API boundaries. It deliberately does not open packages or interpret
//! application-specific protobuf messages.

#![forbid(unsafe_code)]

pub mod archive;
mod error;
mod limits;
mod snappy;

pub use archive::{
    Archive, ArchiveCloneInventory, ArchiveCloneInventoryChunk, ArchiveCloneInventoryError,
    ArchiveCloneInventoryKind, ArchiveCloneInventoryResult, ArchiveCloneInventoryVisitor,
    ArchiveEncodingPlan, ArchiveEncodingPreparationChunk, ArchiveEncodingPreparationError,
    ArchiveEncodingPreparationKind, ArchiveEncodingPreparationResult,
    ArchiveEncodingPreparationVisitor, ArchiveEncodingRequirements, ArchiveInfo,
    ArchiveIngressRequirements, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor,
    CanonicalFieldDataReferenceOperation, CanonicalFieldDataReferenceTransition,
    CanonicalObjectReferenceField, CanonicalObjectReferenceFields, DataReferenceTransition,
    FieldDataReferenceTransition, FieldInfo, FieldPath, FieldType, KnownFieldRule, MessageInfo,
    MessageInfoCollection, RawMessage, UnknownFieldRule,
};
pub use error::{Error, HeaderKind, HeaderOperation, LimitKind, Result};
pub use limits::Limits;
pub use snappy::{
    SnappyCompressionPlan, SnappyCompressionPreparationChunk, SnappyCompressionPreparationError,
    SnappyCompressionPreparationKind, SnappyCompressionPreparationResult,
    SnappyCompressionPreparationVisitor, SnappyCompressionRequirements, SnappyLimits, SnappyStream,
};
