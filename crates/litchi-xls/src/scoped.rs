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
use litchi_ole_common::object::{Edit as ObjectEditor, Package as ObjectPackage};

/// Collects a common OLE package through its bounded streaming output API.
pub(crate) fn object_package_bytes(package: &ObjectPackage) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).map_err(Error::from)?;
    Ok(bytes)
}

/// Commits a cloned common OLE editor and collects its immutable package.
pub(crate) fn object_editor_bytes(editor: &ObjectEditor) -> Result<Vec<u8>> {
    let commit = editor.clone().commit().map_err(Error::from)?;
    object_package_bytes(commit.package())
}

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
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
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
        PackageError::Object(error) => Error::Object(error),
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

/// Declare the whole-package `Patch`, `Commit` and `Diagnostics` of one
/// byte-rendering family in the calling module, with `published` (a commit
/// from the source and rendered bytes) and `replay` (`apply_x`) helpers.
macro_rules! byte_family {
    ($noun:literal, $operation:literal) => {
        #[doc = concat!(
                            "A published ", $noun, " edit: the new package, its exact whole-package ",
                            "patch and diagnostics."
                        )]
        pub type Commit = litchi_core::edit::Commit<$crate::Package, Patch>;

        #[doc = concat!("Content-free diagnostics of one ", $noun, " commit.")]
        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!("An exact, reversible whole-package ", $noun, " patch.")]
        #[derive(Clone, PartialEq, Eq)]
        pub struct Patch {
            before: ::std::sync::Arc<[u8]>,
            after: ::std::sync::Arc<[u8]>,
        }

        impl ::std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter
                    .debug_struct("Patch")
                    .field("before_bytes", &self.before.len())
                    .field("after_bytes", &self.after.len())
                    .finish()
            }
        }

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.before == self.after
            }

            /// The patch that undoes this one.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    before: ::std::sync::Arc::clone(&self.after),
                    after: ::std::sync::Arc::clone(&self.before),
                }
            }

            /// The exact package bytes the patch applies to.
            #[must_use]
            pub fn before(&self) -> &[u8] {
                &self.before
            }

            /// The exact package bytes the patch produces.
            #[must_use]
            pub fn after(&self) -> &[u8] {
                &self.after
            }
        }

        fn published(
            base: $crate::scoped::Base,
            before: &[u8],
            after: Vec<u8>,
        ) -> $crate::Result<Commit> {
            let package = $crate::scoped::publish(base, before, &after)?;
            let patch = Patch {
                before: ::std::sync::Arc::from(before),
                after: ::std::sync::Arc::from(after),
            };
            let diagnostics = Diagnostics::new(!patch.is_noop());
            Ok(Commit::new(package, patch, diagnostics))
        }

        fn replay(package: &$crate::Package, patch: &Patch) -> $crate::Result<Commit> {
            $crate::scoped::replay_source(package, patch.before(), $operation)?;
            let target =
                $crate::scoped::publish(Some(package.clone()), patch.before(), patch.after())?;
            Ok(Commit::new(
                target,
                patch.clone(),
                Diagnostics::new(!patch.is_noop()),
            ))
        }
    };
}

pub(crate) use byte_family;
