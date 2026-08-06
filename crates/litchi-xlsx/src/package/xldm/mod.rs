//! Safe structural inspection of MS-XLDM storage streams.
//!
//! This owner implements only the outer storage described by MS-XLDM 2.1:
//! header, partition marker, serial file allocations, CRC markers, page
//! padding, and virtual directory. Member payloads are never decompressed,
//! decrypted, evaluated, or used for I/O. The nested metadata owner is
//! responsible for the typed section 2.5 model.

pub mod compression;
pub mod crypt;
pub mod generated;
pub mod metadata;
pub mod native;
pub mod olap;

mod codec;
mod model;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{inspect, write};
pub use model::{
    BackupLog, Compression, FileEntry, FileGroup, FileGroupClass, FileKind, GeneratedNameKind,
    GeneratedPath, Header, LoggedFile, Offset, PartitionMarker, Size, Storage, WriteAccess,
    XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE, XmlEncoding,
};
pub use semantic::classify_generated_path;

#[cfg(test)]
pub(crate) use tests::test_xldm_bytes;
