//! The publication side of the record families (record 0360).
//!
//! A record family (Custom Data, Surveys, calculation metadata, page breaks,
//! the Data Model) stages a typed draft against a [`Source`], an owned
//! reference-count clone of the immutable [`Package`] it started from, and
//! commits by publishing a candidate OPC graph as a new package. Its replay
//! payload is its own exact record patch, which implements [`Record`];
//! [`replay`] applies one to another package. The source package never
//! changes.

use std::ops::Deref;

use litchi_opc::OpcPackage;

use crate::Package;
use crate::error::Result;

/// The package a record-family edit started from. It reads as the
/// package's OPC graph.
#[derive(Debug, Clone)]
pub(crate) struct Source {
    package: Package,
}

impl Source {
    /// Start an edit of `package`; refuses packages whose mutations are not
    /// allowed (a package opened with a password).
    pub(crate) fn new(package: &Package, operation: &'static str) -> Result<Self> {
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        Ok(Self {
            package: package.clone(),
        })
    }

    /// Consume the source into its package.
    pub(crate) fn into_package(self) -> Package {
        self.package
    }

    /// Validate `staged` as a complete workbook and publish it as a new
    /// package under the source's policies.
    pub(crate) fn publish(&self, staged: OpcPackage) -> Result<Package> {
        self.package.derive(staged)
    }
}

impl Deref for Source {
    type Target = OpcPackage;

    fn deref(&self) -> &OpcPackage {
        self.package.opc()
    }
}

/// The exact replay payload of a record family.
pub(crate) trait Record: Clone {
    /// Whether replaying changes nothing.
    fn is_noop(&self) -> bool;

    /// Replay onto `package` after checking its source state; a stale
    /// package is refused with `PatchConflict` and left unchanged.
    fn apply_to(&self, package: &mut OpcPackage) -> Result<()>;
}

/// Replay `patch` onto `package` and publish the result as a new package;
/// `package` is unchanged. A no-op patch whose source state matches returns
/// `package` itself.
pub(crate) fn replay<P: Record>(
    package: &Package,
    patch: &P,
    operation: &'static str,
) -> Result<litchi_core::edit::Commit<Package, P>> {
    package.workbook_ref().ensure_mutation_allowed(operation)?;
    let mut candidate = package.opc().clone();
    patch.apply_to(&mut candidate)?;
    let changed = !patch.is_noop();
    let after = if changed {
        package.derive(candidate)?
    } else {
        package.clone()
    };
    Ok(litchi_core::edit::Commit::new(
        after,
        patch.clone(),
        litchi_core::edit::Diagnostics::new(changed),
    ))
}

/// The `Commit`, `Diagnostics` and `Error` members of a record family whose
/// `Patch` implements [`Record`] (ADR 0034 §3).
macro_rules! record_family {
    ($what:literal) => {
        pub use crate::Error;
        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!("The result of a ", $what, " edit's `commit`: the new package, the patch and the diagnostics.")]
        pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;
    };
}

pub(crate) use record_family;

/// A commit of a record edit run directly over a raw OPC graph, for the
/// unit tests that author bare graphs: the published graph is written back
/// to the caller's graph, and the commit keeps the published snapshot.
#[cfg(test)]
#[derive(Debug)]
pub(crate) struct RawCommit<S, P> {
    pub(crate) snapshot: S,
    pub(crate) patch: P,
    pub(crate) changed: bool,
}

#[cfg(test)]
impl<S, P> RawCommit<S, P> {
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }

    pub(crate) const fn snapshot(&self) -> &S {
        &self.snapshot
    }

    pub(crate) const fn patch(&self) -> &P {
        &self.patch
    }
}

/// The caller's graph or package a test route writes a commit back to.
#[cfg(test)]
pub(crate) enum TestTarget<'a> {
    Opc(&'a mut OpcPackage),
    Package(&'a mut Package),
}

#[cfg(test)]
impl TestTarget<'_> {
    /// The package to edit: a bare graph is adopted under strict validation.
    pub(crate) fn package(&self) -> Result<Package> {
        match self {
            Self::Opc(package) => Package::from_opc((**package).clone()),
            Self::Package(package) => Ok((**package).clone()),
        }
    }

    /// Write the published package back.
    pub(crate) fn store(self, published: Package) {
        match self {
            Self::Opc(package) => *package = published.opc().clone(),
            Self::Package(package) => *package = published,
        }
    }
}

#[cfg(test)]
impl<'a> From<&'a mut OpcPackage> for TestTarget<'a> {
    fn from(package: &'a mut OpcPackage) -> Self {
        Self::Opc(package)
    }
}

#[cfg(test)]
impl<'a> From<&'a mut Package> for TestTarget<'a> {
    fn from(package: &'a mut Package) -> Self {
        Self::Package(package)
    }
}
