//! Shared plumbing of the scoped PPT edit families (record 0357).
//!
//! Every scoped family `x` follows the record 0351 §2.3 shape:
//! `Package::edit_x(&self, ..) -> x::Edit`, `x::Edit::commit(self) ->
//! x::Commit` publishing a new [`Package`], and `Package::apply_x(&self,
//! &x::Patch) -> x::Commit`. Dropping an `Edit` discards it. The families
//! keep their record-level codecs and staging owners crate-private; this
//! module captures the source package, publishes a family's candidate bytes
//! as a new package, and replays byte patches.
//!
//! A family's patch is exact and whole-artifact: it carries the source and
//! target package bytes, so `apply_x` refuses a package that is not the
//! patch's source ([`Failure::PatchConflict`]) and `inverse` swaps them. An
//! exact no-op commit publishes the source package unchanged (signatures
//! included); a changed commit of a signed package is refused
//! ([`Failure::Signed`]).

use std::io::Cursor;
use std::sync::Arc;

use litchi_core::edit::{Diagnostics, Refusal, RefusalKind};

use crate::package::{Error as PackageError, OpenOptions, Package};
use crate::slide_order;

/// Exact source and target package bytes of one scoped publication.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct BytePatch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl BytePatch {
    pub(crate) fn new(before: Arc<[u8]>, after: Arc<[u8]>) -> Self {
        Self { before, after }
    }

    pub(crate) fn is_noop(&self) -> bool {
        self.before == self.after
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }

    pub(crate) fn before(&self) -> &[u8] {
        &self.before
    }

    pub(crate) fn after(&self) -> &[u8] {
        &self.after
    }
}

impl std::fmt::Debug for BytePatch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

/// A scoped-family failure before it is wrapped in the family's own `Error`.
#[derive(Debug)]
pub(crate) enum Failure {
    Package(PackageError),
    Refused(Refusal),
    PatchConflict,
    Signed,
}

impl From<PackageError> for Failure {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<slide_order::Error> for Failure {
    fn from(error: slide_order::Error) -> Self {
        match error {
            slide_order::Error::Package(error) => Self::Package(error),
            slide_order::Error::Refused(slide_order::Refusal::Signed) => Self::Signed,
            slide_order::Error::Refused(refusal) => Self::Refused(Refusal::new(
                refusal.kind(),
                "publish",
                "the presentation cannot publish this scoped edit",
            )),
            slide_order::Error::Text(error) => match *error {
                crate::text_edit::Error::Package(error) => Self::Package(error),
                crate::text_edit::Error::Refused(refusal) => Self::Refused(Refusal::new(
                    refusal.kind(),
                    "publish",
                    "the presentation text cannot publish this scoped edit",
                )),
                crate::text_edit::Error::Source(_) => Self::Package(PackageError::Corrupted(
                    "scoped PPT publication failed in the source overlay".into(),
                )),
            },
            slide_order::Error::Composition(_) => Self::Package(PackageError::Corrupted(
                "scoped PPT publication failed to compose".into(),
            )),
        }
    }
}

/// Map a staging failure of `operation`: a value the record grammar refuses
/// becomes a content-free [`RefusalKind::InvalidValue`] refusal; limits and
/// I/O keep their package error.
pub(crate) fn staging(operation: &'static str) -> impl FnOnce(PackageError) -> Failure {
    move |error| match error {
        PackageError::InvalidFormat(_) => Failure::Refused(Refusal::new(
            RefusalKind::InvalidValue,
            operation,
            "the record grammar cannot hold the requested value",
        )),
        error => Failure::Package(error),
    }
}

/// Map a publication failure: content the record publisher cannot preserve
/// becomes a content-free [`RefusalKind::UnsupportedContent`] refusal.
pub(crate) fn publication(operation: &'static str) -> impl FnOnce(PackageError) -> Failure {
    move |error| match error {
        PackageError::InvalidFormat(_) => Failure::Refused(Refusal::new(
            RefusalKind::UnsupportedContent,
            operation,
            "the package holds content this publication cannot preserve",
        )),
        error => Failure::Package(error),
    }
}

/// A refusal of a target that does not exist.
pub(crate) fn not_found(operation: &'static str, reason: &'static str) -> Failure {
    Failure::Refused(Refusal::new(RefusalKind::NotFound, operation, reason))
}

/// A refusal of a target that cannot carry the edit.
pub(crate) fn unsupported_target(operation: &'static str, reason: &'static str) -> Failure {
    Failure::Refused(Refusal::new(
        RefusalKind::UnsupportedTarget,
        operation,
        reason,
    ))
}

/// One published scoped commit before it is wrapped in the family's types.
pub(crate) struct Published {
    pub(crate) package: Package,
    pub(crate) patch: BytePatch,
    pub(crate) diagnostics: Diagnostics,
}

/// The immutable source package a scoped edit started from, with its exact
/// bytes.
#[derive(Clone)]
pub(crate) struct Base {
    package: Package,
    bytes: Arc<[u8]>,
}

impl std::fmt::Debug for Base {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Base")
            .field("package", &self.package)
            .finish_non_exhaustive()
    }
}

impl Base {
    /// Capture `package` for a scoped edit. Encrypted packages are refused:
    /// their content cannot be republished without storing it in plaintext.
    pub(crate) fn capture(package: &Package) -> Result<Self, Failure> {
        if package.is_encrypted() {
            return Err(Failure::Package(PackageError::EncryptionPolicy));
        }
        let bytes: Arc<[u8]> = Arc::from(package.to_encrypted_bytes()?);
        Ok(Self {
            package: package.clone(),
            bytes,
        })
    }

    /// The exact source package bytes.
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Whether the source carries a digital-signature component.
    pub(crate) fn is_signed(&self) -> Result<bool, Failure> {
        contains_signature(&self.bytes)
    }

    /// A shared handle on the exact source package bytes.
    pub(crate) fn shared_bytes(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// The options every republished package is opened with: the source
    /// package's limits and leniency.
    pub(crate) fn open_options(&self) -> OpenOptions<'static> {
        open_options(&self.package)
    }

    /// The source as the native slide-order read view, for families whose
    /// records live in the live document or a slide.
    pub(crate) fn native(&self) -> Result<slide_order::Snapshot, Failure> {
        Ok(slide_order::Snapshot::from_bytes_with_options(
            self.bytes.to_vec(),
            self.open_options(),
        )?)
    }

    /// The exact current bytes of one persisted owner record.
    pub(crate) fn owner_record(
        &self,
        native: &slide_order::Snapshot,
        owner: slide_order::records::Owner,
    ) -> Result<Vec<u8>, Failure> {
        Ok(slide_order::records::owner_record(native, owner)?)
    }

    /// Publish whole candidate package bytes.
    pub(crate) fn publish(self, after: Vec<u8>) -> Result<Published, Failure> {
        if after.as_slice() == self.bytes.as_ref() {
            return Ok(self.unchanged());
        }
        if contains_signature(&self.bytes)? {
            return Err(Failure::Signed);
        }
        let package = Package::from_bytes_with_options(&after, self.open_options())?;
        Ok(Published {
            package,
            patch: BytePatch::new(self.bytes, Arc::from(after)),
            diagnostics: Diagnostics::CHANGED,
        })
    }

    /// Publish one persisted owner record replacement through the
    /// append-only native publisher; every other record keeps its bytes.
    pub(crate) fn publish_owner(
        self,
        native: &slide_order::Snapshot,
        owner: slide_order::records::Owner,
        before: Vec<u8>,
        after: Vec<u8>,
    ) -> Result<Published, Failure> {
        if before == after {
            return Ok(self.unchanged());
        }
        if contains_signature(&self.bytes)? {
            return Err(Failure::Signed);
        }
        let published = slide_order::records::publish(
            native,
            &[slide_order::records::Rewrite {
                owner,
                before: Arc::from(before),
                after: Arc::from(after),
            }],
        )?;
        self.publish(published.bytes().to_vec())
    }

    /// The exact no-op publication: the source package unchanged.
    pub(crate) fn unchanged(self) -> Published {
        Published {
            patch: BytePatch::new(Arc::clone(&self.bytes), self.bytes),
            package: self.package,
            diagnostics: Diagnostics::UNCHANGED,
        }
    }
}

/// Replay an exact byte patch onto the package it was made from.
pub(crate) fn apply(package: &Package, patch: &BytePatch) -> Result<Published, Failure> {
    let base = Base::capture(package)?;
    if base.bytes() != patch.before() {
        return Err(Failure::PatchConflict);
    }
    base.publish(patch.after().to_vec())
}

fn open_options(package: &Package) -> OpenOptions<'static> {
    OpenOptions::new()
        .source_limits(package.source_limits())
        .leniency(package.leniency())
}

fn contains_signature(bytes: &[u8]) -> Result<bool, Failure> {
    Ok(crate::raw::Reader::from_reader(Cursor::new(bytes))?.contains_signature_component()?)
}

/// Declare the canonical family types of one scoped family in the calling
/// module: `Diagnostics`, `LimitKind`, `Error`, `Result`, `Patch` and
/// `Commit`.
macro_rules! family {
    ($noun:literal) => {
        #[doc = concat!("Content-free diagnostics of one ", $noun, " commit.")]
        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!(
                            "A published ", $noun, " edit: the new package, its exact patch and ",
                            "diagnostics."
                        )]
        pub type Commit = litchi_core::edit::Commit<$crate::Package, Patch>;

        #[doc = concat!(
                            "A ", $noun, " edit limit dimension.\n\n",
                            "The records bound their own sizes and the republished package is ",
                            "admitted under the source package's [`SourceLimits`]",
                            "(crate::SourceLimits), so no caller-selected dimension exists."
                        )]
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {}

        #[doc = concat!(
                            "A ", $noun, " edit failure. Messages carry no authored content."
                        )]
        #[derive(Debug)]
        #[non_exhaustive]
        pub enum Error {
            /// The source package could not be read, or the edited bytes
            /// could not be republished as a PPT package.
            Package($crate::package::Error),
            /// A well-formed request this route refuses, classified by kind.
            Refused(litchi_core::edit::Refusal),
            /// The patch was made from a different source package.
            PatchConflict,
            /// A changed write would invalidate the package signature.
            Signed,
        }

        impl ::std::fmt::Display for Error {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                match self {
                    Self::Package(_) => {
                        formatter.write_str(concat!("PPT ", $noun, " edit failed in the package"))
                    },
                    Self::Refused(refusal) => ::std::fmt::Display::fmt(refusal, formatter),
                    Self::PatchConflict => formatter.write_str(concat!(
                        "PPT ",
                        $noun,
                        " patch source does not match the package receiving it"
                    )),
                    Self::Signed => formatter.write_str(concat!(
                        "changed PPT ",
                        $noun,
                        " writes require explicit signature removal"
                    )),
                }
            }
        }

        impl ::std::error::Error for Error {
            fn source(&self) -> Option<&(dyn ::std::error::Error + 'static)> {
                match self {
                    Self::Package(error) => Some(error),
                    Self::Refused(_) | Self::PatchConflict | Self::Signed => None,
                }
            }
        }

        impl From<$crate::package::Error> for Error {
            fn from(error: $crate::package::Error) -> Self {
                Self::Package(error)
            }
        }

        impl From<$crate::scoped::Failure> for Error {
            fn from(failure: $crate::scoped::Failure) -> Self {
                match failure {
                    $crate::scoped::Failure::Package(error) => Self::Package(error),
                    $crate::scoped::Failure::Refused(refusal) => Self::Refused(refusal),
                    $crate::scoped::Failure::PatchConflict => Self::PatchConflict,
                    $crate::scoped::Failure::Signed => Self::Signed,
                }
            }
        }

        #[doc = concat!("Result type of the ", $noun, " family.")]
        pub type Result<T> = ::std::result::Result<T, Error>;

        #[doc = concat!(
                            "An exact, reversible ", $noun, " patch bound to the package bytes it ",
                            "was made from."
                        )]
        #[derive(Clone, PartialEq, Eq)]
        pub struct Patch($crate::scoped::BytePatch);

        impl ::std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                ::std::fmt::Debug::fmt(&self.0, formatter)
            }
        }

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.0.is_noop()
            }

            /// The patch that undoes this one, applicable to its result.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(self.0.inverse())
            }

            /// The exact package bytes the patch applies to.
            #[must_use]
            pub fn before_bytes(&self) -> &[u8] {
                self.0.before()
            }

            /// The exact package bytes the patch produces.
            #[must_use]
            pub fn after_bytes(&self) -> &[u8] {
                self.0.after()
            }
        }

        #[allow(
            dead_code,
            reason = "every family publishes through these helpers; a family may use a subset"
        )]
        fn published(published: $crate::scoped::Published) -> Commit {
            Commit::new(
                published.package,
                Patch(published.patch),
                published.diagnostics,
            )
        }

        #[allow(
            dead_code,
            reason = "every family replays through this helper; a family may use a subset"
        )]
        fn apply_patch(package: &$crate::Package, patch: &Patch) -> Result<Commit> {
            Ok(published($crate::scoped::apply(package, &patch.0)?))
        }
    };
}

pub(crate) use family;
