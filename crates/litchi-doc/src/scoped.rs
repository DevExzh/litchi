//! Shared plumbing of the scoped DOC edit families (record 0359).
//!
//! Every scoped family `x` follows the record 0351 §2.3 shape:
//! `Package::edit_x(&self) -> x::Edit` (and `edit_x_with_options`),
//! `x::Edit::commit(self) -> x::Commit` publishing a new [`Package`], and
//! `Package::apply_x(&self, &x::Patch) -> x::Commit`. Dropping an `Edit`
//! discards it. The families keep their part-level codecs (`x::record`) and
//! byte-level editors crate-private; this module captures the source
//! package, republishes a family's rendered bytes as a new package under the
//! source's open policy, and replays exact byte patches.
//!
//! A family's patch is exact and whole-artifact: it carries the source and
//! target package bytes, so `apply_x` refuses a package that is not the
//! patch's source ([`Failure::PatchConflict`]) and `inverse` swaps them. An
//! exact no-op commit publishes the source package unchanged (signatures
//! included); a changed commit of a signed package is refused
//! ([`Failure::Signed`]).

use std::io::Cursor;
use std::sync::Arc;

use litchi_core::edit::{Diagnostics, Refusal};

use crate::package::{EditOptions, Error as PackageError, OpenOptions, Package};

/// The names of the compound-file entries that carry a document signature.
const SIGNATURE_ENTRIES: [&str; 4] = [
    "_xmlsignatures",
    "_signatures",
    "DigitalSignature",
    "\u{0005}DigitalSignature",
];

/// Exact source and target package bytes of one scoped publication.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct BytePatch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl BytePatch {
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
    #[allow(
        dead_code,
        reason = "families that refuse before the part editor construct it"
    )]
    Refused(Refusal),
    PatchConflict,
    Signed,
}

impl From<PackageError> for Failure {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

/// One published scoped commit before it is wrapped in the family's types.
pub(crate) struct Published {
    pub(crate) package: Package,
    pub(crate) patch: BytePatch,
    pub(crate) diagnostics: Diagnostics,
}

/// The immutable source package a scoped edit started from, with its exact
/// bytes and the edit options it grants.
#[derive(Clone)]
pub(crate) struct Base {
    package: Package,
    bytes: Arc<std::sync::Mutex<Option<Arc<[u8]>>>>,
    options: EditOptions,
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
    /// Capture `package` for a scoped edit under its retained edit options.
    pub(crate) fn capture(package: &Package) -> Result<Self, Failure> {
        Self::capture_with_options(package, package.retained_edit_options())
    }

    /// Capture `package` for a scoped edit under explicit edit options,
    /// constrained by the profile the package retains.
    pub(crate) fn capture_with_options(
        package: &Package,
        options: EditOptions,
    ) -> Result<Self, Failure> {
        Ok(Self {
            package: package.clone(),
            bytes: Arc::new(std::sync::Mutex::new(None)),
            options: package.constrained_edit_options(options),
        })
    }

    /// Materialize the exact source package bytes on first use.
    ///
    /// Capturing an edit only retains the immutable package and its effective
    /// options. The source fence and bounded copy are deliberately kept at the
    /// first operation that needs bytes, so a dropped or untouched edit is
    /// constant-time and allocation-free in the artifact size.
    pub(crate) fn bytes(&self) -> Result<Arc<[u8]>, PackageError> {
        let mut memo = self
            .bytes
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(bytes) = memo.as_ref() {
            return Ok(Arc::clone(bytes));
        }
        let bytes: Arc<[u8]> = Arc::from(self.package.to_bytes()?);
        *memo = Some(Arc::clone(&bytes));
        Ok(bytes)
    }

    /// The exact source package bytes as an owned buffer.
    pub(crate) fn to_vec(&self) -> Result<Vec<u8>, PackageError> {
        Ok(self.bytes()?.to_vec())
    }

    /// The edit options granted to this edit.
    pub(crate) const fn options(&self) -> &EditOptions {
        &self.options
    }

    /// The physical source policy retained by the immutable package captured
    /// for this edit.
    pub(crate) const fn source_limits(&self) -> crate::package::SourceLimits {
        self.package.source_limits()
    }

    /// Publish whole candidate package bytes.
    pub(crate) fn publish(self, after: Vec<u8>) -> Result<Published, Failure> {
        let before = self.bytes().map_err(Failure::Package)?;
        if after.as_slice() == before.as_ref() {
            return self.unchanged();
        }
        if is_signed(&before)? {
            return Err(Failure::Signed);
        }
        let package = Package::from_bytes_with_options(&after, open_options(&self.package))?
            .with_edit_options(&self.options);
        Ok(Published {
            package,
            patch: BytePatch {
                before,
                after: Arc::from(after),
            },
            diagnostics: Diagnostics::CHANGED,
        })
    }

    /// The exact no-op publication: the source package unchanged.
    pub(crate) fn unchanged(self) -> Result<Published, Failure> {
        let bytes = self.bytes().map_err(Failure::Package)?;
        Ok(Published {
            patch: BytePatch {
                before: Arc::clone(&bytes),
                after: bytes,
            },
            package: self.package,
            diagnostics: Diagnostics::UNCHANGED,
        })
    }
}

/// Replay an exact byte patch onto the package it was made from.
pub(crate) fn apply(package: &Package, patch: &BytePatch) -> Result<Published, Failure> {
    let base = Base::capture(package)?;
    let bytes = base.bytes().map_err(Failure::Package)?;
    if bytes.as_ref() != patch.before() {
        return Err(Failure::PatchConflict);
    }
    if !patch.is_noop() {
        // A replayed change is a write: the receiving package's protection
        // policy authorizes it exactly as it authorizes a staged edit.
        let protection = crate::package::property_set::classify_host(&bytes)?;
        base.options()
            .protection_policy_value()
            .authorize(protection)?;
    }
    base.publish(patch.after().to_vec())
}

fn open_options(package: &Package) -> OpenOptions {
    OpenOptions::default()
        .source_limits(package.source_limits())
        .leniency(package.leniency())
}

fn is_signed(bytes: &[u8]) -> Result<bool, Failure> {
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .map_err(PackageError::from)?;
    Ok(SIGNATURE_ENTRIES.iter().any(|name| ole.exists(&[name])))
}

/// Declare the canonical family types of one scoped family in the calling
/// module: `Diagnostics`, `LimitKind`, `Error`, `Result`, `Patch` and
/// `Commit`, plus the `published` and `apply_patch` helpers.
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
                            "The part bounds its own records under the edit options' revision ",
                            "limits and the republished package is admitted under the source ",
                            "package's [`SourceLimits`](crate::SourceLimits), so no ",
                            "caller-selected dimension exists."
                        )]
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {}

        #[doc = concat!("A ", $noun, " edit failure. Messages carry no authored content.")]
        #[derive(Debug)]
        #[non_exhaustive]
        pub enum Error {
            /// The source package could not be read, the part failed its
            /// record validation, or the edited bytes could not be
            /// republished as a DOC package.
            Package($crate::Error),
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
                    Self::Package(error) => ::std::fmt::Display::fmt(error, formatter),
                    Self::Refused(refusal) => ::std::fmt::Display::fmt(refusal, formatter),
                    Self::PatchConflict => formatter.write_str(concat!(
                        "DOC ",
                        $noun,
                        " patch source does not match the package receiving it"
                    )),
                    Self::Signed => formatter.write_str(concat!(
                        "changed DOC ",
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

        impl From<$crate::Error> for Error {
            fn from(error: $crate::Error) -> Self {
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
