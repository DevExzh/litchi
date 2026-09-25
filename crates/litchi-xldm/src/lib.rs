//! Safe structural inspection of MS-XLDM storage streams.
//!
//! This owner implements the outer storage described by MS-XLDM 2.1 and an
//! explicit, fixture-backed version-150 tabular compatibility profile:
//! header, partition marker, serial file allocations, CRC markers, page
//! padding, and virtual directory. Member payloads are never decompressed,
//! decrypted, evaluated, or used for I/O. The nested metadata owner is
//! responsible for the typed section 2.5 model.

// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

mod error;

pub mod compression;
pub mod crypt;
pub mod generated;
pub mod identity;
pub mod metadata;
pub mod native;
pub mod olap;
pub mod olapproof;

mod codec;
mod model;
mod seen_names;
mod semantic;
mod tabular_paths;
mod validation;
mod xml_attributes;

#[cfg(test)]
mod tests;

pub use codec::{inspect, write};
pub use error::{Error, Result};
pub use identity::{
    Xldm140Closure, Xldm140ClosureMember, Xldm140ColumnBinding, Xldm140ColumnIdentity,
    Xldm140FileReplacement, Xldm140IdentityProjection, Xldm140InversePatch, Xldm140MemberSection,
    Xldm140Patch, Xldm140PatchBytes, Xldm140RelationshipIdentity, Xldm140RenameError,
    Xldm140RenameErrorKind, Xldm140TableIdentity, Xldm140TimeGroupingBinding,
    Xldm140TimeGroupingContentType, project_xldm140_identity,
    project_xldm140_identity_with_closure, prove_xldm140_closure,
    validate_xldm140_identity_closure,
};
pub use model::{
    BackupLog, Compression, FileEntry, FileGroup, FileGroupClass, FileKind, GeneratedNameKind,
    GeneratedPath, Header, LoggedFile, Offset, PartitionMarker, Size, Storage, StorageProfile,
    WriteAccess, XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE, XmlEncoding,
};
pub use olapproof::{
    MeasureGroupDimensionKind, OlapCubeBinding, OlapCubeDimensionBinding, OlapFileGroupBinding,
    OlapMeasureGroupBinding, OlapMeasureGroupDimensionBinding, OlapPartitionBinding,
    OlapProofError, OlapProofLimits, OlapProofResult, OlapReference, OlapReferenceField,
    OlapRelationshipBinding, OlapTableBinding, OlapUnknownMember, Xldm140OlapProof,
    is_relationship_index_path, prove_xldm140_olap,
};
pub use semantic::classify_generated_path;

/// Inspect a borrowed source buffer and retain it for all typed views.
///
/// This name makes the source-sharing lifetime explicit at call sites; it is
/// equivalent to [`inspect`]. No member payload is copied by inspection.
#[must_use = "inspection results carry the borrowed source lifetime"]
pub fn inspect_shared(bytes: &[u8]) -> Result<Storage<'_>> {
    inspect(bytes)
}
