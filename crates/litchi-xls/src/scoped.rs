//! Shared plumbing of the scoped XLS edit families (record 0359).
//!
//! Every package-level family `x` follows the record 0351 §2.3 shape:
//! `Package::edit_x(&self) -> x::Edit`, `x::Edit::commit(self) -> x::Commit`
//! publishing a new [`Package`], and `Package::apply_x(&self, &x::Patch) ->
//! x::Commit`. Dropping an `Edit` discards it. The families keep their
//! record-level staging (the exact bytes they render) and this module turns
//! the rendered artifact into a published package: an exact no-op returns
//! the source package unchanged (signatures included), a changed artifact is
//! republished under the source package's open policy, and a changed write
//! of a signed source is refused with [`crate::Error::Signed`].

use std::io::Cursor;

use crate::Package;
use crate::error::{Error, Result};

/// The names of the compound-file entries that carry a document signature.
const SIGNATURE_ENTRIES: [&str; 4] = [
    "_xmlsignatures",
    "_signatures",
    "DigitalSignature",
    "\u{0005}DigitalSignature",
];

/// Whether the compound file `bytes` carries a document signature entry at
/// its root.
pub(crate) fn is_signed(bytes: &[u8]) -> Result<bool> {
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes))?;
    Ok(SIGNATURE_ENTRIES.iter().any(|name| ole.exists(&[name])))
}

/// The source a family edit started from: the package whose open policy a
/// commit republishes under, or nothing for a codec-level edit of bare
/// bytes (whose commit opens the result with the default policy).
pub(crate) type Base = Option<Package>;

/// Capture `package` as the base of a family edit, with its exact bytes.
pub(crate) fn capture(package: &Package) -> Result<(Base, Vec<u8>)> {
    if package.is_encrypted() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::Protected,
            "edit workbook",
            "an encrypted workbook cannot be republished by a scoped edit",
        ));
    }
    let bytes = package.to_bytes().map_err(package_error)?;
    Ok((Some(package.clone()), bytes))
}

/// Publish the artifact a family rendered: `before` are the source bytes,
/// `after` the rendered ones.
pub(crate) fn publish(base: Base, before: &[u8], after: &[u8]) -> Result<Package> {
    if before == after {
        if let Some(package) = base {
            return Ok(package);
        }
        return Package::from_bytes(after).map_err(package_error);
    }
    if is_signed(before)? {
        return Err(Error::Signed);
    }
    match base {
        Some(package) => crate::package::republish(&package, after.to_vec()).map_err(package_error),
        None => Package::from_bytes(after).map_err(package_error),
    }
}

/// The exact source bytes of `package` for a patch replay, or a patch
/// conflict when they are not `expected`.
pub(crate) fn replay_source(
    package: &Package,
    expected: &[u8],
    operation: &'static str,
) -> Result<Vec<u8>> {
    let bytes = package.to_bytes().map_err(package_error)?;
    if bytes != expected {
        return Err(Error::PatchConflict {
            operation,
            reason: "the patch was made from another package",
        });
    }
    Ok(bytes)
}

/// The record-level error behind a package error.
pub(crate) fn package_error(error: crate::package::Error) -> Error {
    use crate::package::Error as PackageError;
    match error {
        PackageError::Reader(error) => error,
        PackageError::Cfb(error) => Error::Cfb(error),
        PackageError::Io(error) => Error::Io(error),
        PackageError::Allocation { resource, .. } => Error::Allocation(resource),
        PackageError::EncryptionPolicy => Error::refused(
            litchi_core::edit::RefusalKind::Protected,
            "edit workbook",
            "an encrypted workbook cannot be republished by a scoped edit",
        ),
        _ => Error::Internal("the edited workbook could not be republished"),
    }
}
