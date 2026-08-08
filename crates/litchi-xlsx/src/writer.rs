//! XLSX package publication over the OPC-owned streaming writer.
//!
//! This layer deliberately contains no SpreadsheetML model. It is the narrow
//! publication boundary used by both [`crate::package::Package`] and the
//! immutable workbook snapshot, so serialization never falls back to a host
//! archive implementation or an archive-sized intermediate buffer.

use std::io::Write;
use std::path::Path;

use litchi_opc::{OpcPackage, PackageWriter};

use crate::error::Result;

#[path = "writer/shape.rs"]
#[allow(dead_code)]
pub mod shape;

/// Serialize a validated XLSX package into owned bytes.
pub(crate) fn to_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
    Ok(PackageWriter::to_bytes(package)?)
}

/// Stream a validated XLSX package into a sequential sink.
pub(crate) fn write_to(package: &OpcPackage, writer: impl Write) -> Result<()> {
    Ok(PackageWriter::write_to_stream(writer, package)?)
}

/// Atomically publish a validated XLSX package to a filesystem path.
pub(crate) fn save(package: &OpcPackage, path: impl AsRef<Path>) -> Result<()> {
    Ok(PackageWriter::write(path, package)?)
}

/// Atomically publish already-encrypted managed-package bytes.
#[cfg(feature = "encryption")]
pub(crate) fn save_encrypted(bytes: &[u8], path: impl AsRef<Path>) -> Result<()> {
    litchi_opc::atomic::replace_with::<crate::Error>(path.as_ref(), |temporary| {
        temporary
            .write_all(bytes)
            .map_err(|source| crate::Error::Encryption(litchi_crypto::ooxml::Error::Io(source)))
    })
}
