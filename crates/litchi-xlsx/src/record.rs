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

    /// Conservative retained-data charge for the source endpoint kept by a
    /// record-family edit.
    pub(crate) fn retained_bytes(&self) -> usize {
        self.package.retained_bytes()
    }

    /// Validate `staged` as a complete workbook and publish it as a new
    /// package under the source's policies.
    /// A changed graph of a signed package is refused with
    /// [`crate::Error::Signed`] (ADR 0006).
    pub(crate) fn publish(&self, staged: OpcPackage) -> Result<Package> {
        crate::signature_policy::ensure_unsigned(self.package.opc())?;
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
/// `package` itself; a changed patch of a signed package is refused with
/// [`crate::Error::Signed`] before its source state is checked.
pub(crate) fn replay<P: Record>(
    package: &Package,
    patch: &P,
    operation: &'static str,
) -> Result<litchi_core::edit::Commit<Package, P>> {
    package.workbook_ref().ensure_mutation_allowed(operation)?;
    let changed = !patch.is_noop();
    if changed {
        crate::signature_policy::ensure_unsigned(package.opc())?;
    }
    let mut candidate = package.opc().clone();
    patch.apply_to(&mut candidate)?;
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

/// The typed bounds of a record family: each variant names one bound and its
/// structural description, reported by `Error::Limit` through the crate's
/// `LimitKind::$crate_variant`.
macro_rules! limit_kinds {
    ($what:literal, $crate_variant:ident { $($(#[$doc:meta])* $variant:ident => $text:literal,)+ }) => {
        #[doc = concat!("A bounded dimension of the ", $what, " family.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {
            $($(#[$doc])* $variant,)+
        }

        impl LimitKind {
            /// A short structural description of the dimension.
            #[must_use]
            pub const fn as_str(self) -> &'static str {
                match self {
                    $(Self::$variant => $text,)+
                }
            }
        }

        impl std::fmt::Display for LimitKind {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter.write_str(self.as_str())
            }
        }

        impl From<LimitKind> for crate::LimitKind {
            fn from(kind: LimitKind) -> Self {
                Self::$crate_variant(kind)
            }
        }
    };
}

pub(crate) use limit_kinds;

/// A raw-graph route over a record family's `Edit`, for the unit tests that
/// author bare graphs: `Transaction::new(&mut graph, ..)` adopts the graph as
/// a package, starts `Package::$entry(..)`, and `commit` writes the
/// published graph back. The family's `Edit` provides
/// `fn publish(self) -> Result<(Commit, Snapshot)>`.
macro_rules! raw_test_route {
    ($entry:ident $(, $arg:ident : $ty:ty => $map:expr)*) => {
        #[cfg(test)]
        pub(crate) struct Transaction<'a> {
            target: crate::record::TestTarget<'a>,
            edit: Edit,
        }

        #[cfg(test)]
        impl<'a> Transaction<'a> {
            pub(crate) fn new(
                target: impl Into<crate::record::TestTarget<'a>>,
                $($arg: $ty,)*
            ) -> crate::error::Result<Self> {
                let target = target.into();
                let package = target.package()?;
                let edit = package.$entry($(($map)(&package, $arg)?),*)?;
                Ok(Self { target, edit })
            }

            pub(crate) fn commit(
                self,
            ) -> crate::error::Result<crate::record::RawCommit<Snapshot, Patch>> {
                let (commit, snapshot) = self.edit.publish()?;
                let changed = commit.diagnostics().changed();
                let (package, patch) = commit.into_parts();
                self.target.store(package);
                Ok(crate::record::RawCommit {
                    snapshot,
                    patch,
                    changed,
                })
            }
        }

        #[cfg(test)]
        impl std::ops::Deref for Transaction<'_> {
            type Target = Edit;

            fn deref(&self) -> &Edit {
                &self.edit
            }
        }

        #[cfg(test)]
        impl std::ops::DerefMut for Transaction<'_> {
            fn deref_mut(&mut self) -> &mut Edit {
                &mut self.edit
            }
        }
    };
}

pub(crate) use raw_test_route;

/// The catalog position of the sheet whose part is `part`, for the raw test
/// routes of worksheet-scoped families.
#[cfg(test)]
pub(crate) fn sheet_of(package: &Package, part: &litchi_opc::PackURI) -> Result<usize> {
    package
        .workbook_ref()
        .sheets()
        .position(|sheet| sheet.part_uri().is_equivalent_to(part))
        .ok_or_else(|| crate::error::invalid("no sheet owns the worksheet part"))
}

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

/// Make a bare graph holding the worksheet part `worksheet` a complete
/// workbook for the raw test routes: add a workbook part (in the Strict
/// namespace when `strict`) whose only sheet is that worksheet, and the
/// package's main-document relationship.
#[cfg(test)]
pub(crate) fn adopt_worksheet(
    package: &mut OpcPackage,
    worksheet: &litchi_opc::PackURI,
    strict: bool,
) {
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    let (main, relationships, office, sheet) = if strict {
        (
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
        )
    } else {
        (
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            rt::OFFICE_DOCUMENT,
            rt::WORKSHEET,
        )
    };
    let workbook = litchi_opc::PackURI::new("/xl/workbook.xml").expect("static part name");
    let mut part = litchi_opc::BlobPart::new(
        workbook.clone(),
        ct::SML_SHEET_MAIN.into(),
        format!(
            "<workbook xmlns=\"{main}\" xmlns:r=\"{relationships}\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets></workbook>"
        )
        .into_bytes(),
    );
    litchi_opc::Part::rels_mut(&mut part).add_relationship(
        sheet.into(),
        worksheet.relative_ref(workbook.base_uri()),
        "rIdSheet1".into(),
        false,
    );
    package.add_part(Box::new(part));
    package.rels_mut().add_relationship(
        office.into(),
        "xl/workbook.xml".into(),
        "rIdWorkbook".into(),
        false,
    );
}
