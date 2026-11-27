//! The equation family's commit, patch and error types: the scoped-family
//! types (see `crate::scoped::family`) with one more failure, an equation
//! that does not convert or render.

use crate::scoped::{self, Failure};

pub use litchi_core::edit::Diagnostics;

/// A published equation edit: the new package, its exact patch and
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

/// An equation edit failure. Messages carry no authored content.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The source package could not be read, or the edited bytes could not
    /// be republished as a PPT package.
    Package(crate::package::Error),
    /// A well-formed request this route refuses, classified by kind.
    Refused(litchi_core::edit::Refusal),
    /// The equation does not convert to MTEF or does not render within the
    /// caller's limits.
    Equation(litchi_equation::Error),
    /// The patch was made from a different source package.
    PatchConflict,
    /// A changed write would invalidate the package signature.
    Signed,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Package(_) => formatter.write_str("PPT equation edit failed in the package"),
            Self::Refused(refusal) => std::fmt::Display::fmt(refusal, formatter),
            Self::Equation(error) => std::fmt::Display::fmt(error, formatter),
            Self::PatchConflict => formatter
                .write_str("PPT equation patch source does not match the package receiving it"),
            Self::Signed => formatter
                .write_str("changed PPT equation writes require explicit signature removal"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Equation(error) => Some(error),
            Self::Refused(_) | Self::PatchConflict | Self::Signed => None,
        }
    }
}

impl From<crate::package::Error> for Error {
    fn from(error: crate::package::Error) -> Self {
        Self::Package(error)
    }
}

impl From<litchi_equation::Error> for Error {
    fn from(error: litchi_equation::Error) -> Self {
        Self::Equation(error)
    }
}

impl From<Failure> for Error {
    fn from(failure: Failure) -> Self {
        match failure {
            Failure::Package(error) => Self::Package(error),
            Failure::Refused(refusal) => Self::Refused(refusal),
            Failure::PatchConflict => Self::PatchConflict,
            Failure::Signed => Self::Signed,
        }
    }
}

/// Result type of the equation family.
pub type Result<T> = std::result::Result<T, Error>;

/// An exact, reversible equation patch: the changed byte ranges between two
/// packages, bound to the exact source package (length and block root) it
/// was made from.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch(scoped::BytePatch);

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Debug::fmt(&self.0, formatter)
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

    /// The byte length of the package the patch applies to.
    #[must_use]
    pub fn source_len(&self) -> usize {
        self.0.source_len()
    }

    /// The byte length of the package the patch produces.
    #[must_use]
    pub fn target_len(&self) -> usize {
        self.0.target_len()
    }

    /// The number of package bytes the patch's changed ranges span; zero
    /// for a no-op.
    #[must_use]
    pub fn changed_bytes(&self) -> usize {
        self.0.changed_bytes()
    }
}

pub(super) fn published(published: scoped::Published) -> Commit {
    Commit::new(
        published.package,
        Patch(published.patch),
        published.diagnostics,
    )
}

pub(super) fn apply_patch(package: &crate::Package, patch: &Patch) -> Result<Commit> {
    Ok(published(scoped::apply(package, &patch.0)?))
}
