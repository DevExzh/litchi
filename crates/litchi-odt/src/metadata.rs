//! Document properties of `meta.xml` as a scoped transaction (ADR 0034 §7).
//!
//! [`Package::edit_metadata`](crate::Package::edit_metadata) stages an
//! [`Edit`]; [`Edit::commit`] publishes a new [`Package`](crate::Package)
//! and leaves the source untouched;
//! [`Package::apply_metadata`](crate::Package::apply_metadata) replays the
//! [`Patch`] onto the exact source snapshot. A change splices only the
//! touched `office:meta` children of `meta.xml`; generator, dates and
//! statistics are never rewritten, so a commit consults no clock.
//!
//! ```
//! use litchi_odt::Package;
//! use litchi_odt::metadata::Field;
//!
//! # fn main() -> Result<(), litchi_odt::metadata::Error> {
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata();
//! edit.set(Field::Title, Some("Quarterly report"))?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let next = commit.into_package();
//! let title = next.document()?.metadata()?.title;
//! assert_eq!(title.as_deref(), Some("Quarterly report"));
//! # Ok(())
//! # }
//! ```

/// The family error: typed limits plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;

use crate::body::xml::Ns;

/// A document property the metadata transaction writes.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Field {
    /// `dc:title`.
    Title,
    /// `dc:subject`.
    Subject,
    /// `dc:description`.
    Description,
    /// `meta:initial-creator`, the original author.
    Author,
    /// `dc:creator`, who saved the document last.
    LastModifiedBy,
    /// `dc:language`.
    Language,
    /// The `meta:keyword` elements, written from a comma-separated list.
    Keywords,
}

impl Field {
    pub(crate) const fn qualified(self) -> (Ns, &'static str) {
        match self {
            Self::Title => (Ns::Dc, "title"),
            Self::Subject => (Ns::Dc, "subject"),
            Self::Description => (Ns::Dc, "description"),
            Self::Author => (Ns::Meta, "initial-creator"),
            Self::LastModifiedBy => (Ns::Dc, "creator"),
            Self::Language => (Ns::Dc, "language"),
            Self::Keywords => (Ns::Meta, "keyword"),
        }
    }
}

/// A detached metadata edit of one source package.
pub struct Edit {
    inner: crate::transaction::Edit,
    fields: usize,
}

impl Edit {
    pub(crate) fn new(package: &crate::Package) -> Self {
        Self {
            inner: package.edit(),
            fields: 0,
        }
    }

    /// Stage setting (`Some`) or removing (`None`) one property. A later
    /// assignment of the same field replaces an earlier one.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for text XML cannot carry.
    pub fn set(&mut self, field: Field, value: Option<&str>) -> Result<&mut Self, Error> {
        self.inner.set_metadata(field, value)?;
        self.fields += 1;
        Ok(self)
    }

    /// Publish the staged properties as one new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] when the package has no `meta.xml`, a
    /// namespace the edit writes is undeclared, or the package is signed or
    /// encrypted (changed writes there are refused).
    pub fn commit(self) -> Result<Commit, Error> {
        let fields = self.fields;
        let (package, patch) = self.inner.commit()?.into_parts();
        let changed = !patch.is_noop();
        Ok(Commit {
            package,
            patch: Patch(patch),
            diagnostics: Diagnostics { changed, fields },
        })
    }
}

/// The result of one metadata commit.
pub struct Commit {
    package: crate::Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The committed package.
    #[must_use]
    pub const fn package(&self) -> &crate::Package {
        &self.package
    }

    /// The exact source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free facts about the commit.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// The committed package.
    #[must_use]
    pub fn into_package(self) -> crate::Package {
        self.package
    }

    /// The committed package and its patch.
    #[must_use]
    pub fn into_parts(self) -> (crate::Package, Patch) {
        (self.package, self.patch)
    }
}

/// Content-free facts about one metadata commit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    fields: usize,
}

impl Diagnostics {
    /// Whether the package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// The number of field assignments staged.
    #[must_use]
    pub const fn field_count(&self) -> usize {
        self.fields
    }
}

/// An exact-byte metadata patch, applicable only to its source snapshot.
#[derive(Clone)]
pub struct Patch(crate::transaction::Patch);

impl Patch {
    /// Whether applying the patch preserves the source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    /// The patch that restores the source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    pub(crate) const fn primary(&self) -> &crate::transaction::Patch {
        &self.0
    }
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("noop", &self.is_noop())
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("fields", &self.fields)
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}
