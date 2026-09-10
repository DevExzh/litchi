//! Safe structural inspection of MS-XLDM storage streams.
//!
//! This owner implements the outer storage described by MS-XLDM 2.1 and an
//! explicit, fixture-backed version-150 tabular compatibility profile:
//! header, partition marker, serial file allocations, CRC markers, page
//! padding, and virtual directory. Member payloads are never decompressed,
//! decrypted, evaluated, or used for I/O. The nested metadata owner is
//! responsible for the typed section 2.5 model.

mod error;

pub mod compression;
pub mod crypt;
pub mod generated;
pub mod metadata;
pub mod native;
pub mod olap;

mod codec;
mod model;
mod semantic;
mod tabular_paths;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{inspect, write};
pub use error::{Error, Result};
pub use model::{
    BackupLog, Compression, FileEntry, FileGroup, FileGroupClass, FileKind, GeneratedNameKind,
    GeneratedPath, Header, LoggedFile, Offset, PartitionMarker, Size, Storage, StorageProfile,
    WriteAccess, XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE, XmlEncoding,
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
