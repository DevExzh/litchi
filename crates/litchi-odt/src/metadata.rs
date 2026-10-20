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

/// A document property the metadata transaction writes (shared by every
/// ODF leaf, record 0288).
pub use litchi_odf_common::metadata::Field;

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

    /// Stage the keyword list, one `meta:keyword` element per keyword; an
    /// empty list removes every keyword.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for text XML cannot carry.
    pub fn set_keywords(&mut self, keywords: &[String]) -> Result<&mut Self, Error> {
        self.inner.set_metadata_keywords(keywords)?;
        self.fields += 1;
        Ok(self)
    }

    /// Publish the staged properties as one new package.
    ///
    /// A package without `meta.xml` gains one declaring the content's
    /// `office:version`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] when a namespace the edit writes is
    /// undeclared, or the package is signed or encrypted (changed writes
    /// there are refused).
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
    pub(crate) fn replayed(package: crate::Package, primary: crate::transaction::Patch) -> Self {
        let changed = !primary.is_noop();
        Self {
            package,
            patch: Patch(primary),
            diagnostics: Diagnostics { changed, fields: 0 },
        }
    }

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
