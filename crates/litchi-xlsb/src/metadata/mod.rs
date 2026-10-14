//! Document metadata of an XLSB package: the core, application and custom
//! property families, edited as one exact and reversible transaction.
//!
//! [`Package::metadata`] reads the unified [`Metadata`];
//! [`Package::edit_metadata`] starts an [`Edit`] whose [`Edit::commit`]
//! returns a [`Commit`] holding the new snapshot, the [`Patch`] and the
//! [`Diagnostics`]. [`Package::apply_metadata`] replays a patch onto an
//! equal snapshot. The source snapshot never changes. [`source`] edits the
//! same families of a [`crate::SourceBackedWorkbook`] without materializing
//! it.
//!
//! The shared OOXML engine ([`litchi_ooxml_common::properties`]) routes each
//! field to the part that stores it and keeps the lexical form of every
//! unchanged value. A field no OOXML family stores is refused with
//! [`Error::Unsupported`]; a changed edit of a signed workbook is refused
//! with [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! ```
//! use litchi_xlsb::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata()?;
//! edit.update(|metadata| metadata.title = Some("Budget".to_owned()))?;
//! let after = edit.commit()?.into_package();
//! assert_eq!(after.metadata()?.title.as_deref(), Some("Budget"));
//! assert_eq!(package.metadata()?.title, None);
//! # Ok(())
//! # }
//! ```

use litchi_core::Metadata;
use litchi_ooxml_common::custom::{self, Host};
use litchi_ooxml_common::properties::{self, Snapshot, app};

pub use litchi_ooxml_common::properties::{Diagnostics, Error, LimitKind, Patch};

use crate::Workbook;
use crate::package::Package;

pub mod source;

type Result<T> = std::result::Result<T, Error>;

/// A staged metadata edit of one [`Package`] or [`Workbook`].
///
/// Dropping the edit discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    owner: Owner<'a>,
    inner: properties::Edit,
}

/// The snapshot an [`Edit`] stages over.
#[derive(Debug, Clone, Copy)]
enum Owner<'a> {
    Package(&'a Package),
    Workbook(&'a Workbook),
}

impl Edit<'_> {
    /// The staged metadata.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of the application-properties part.
    pub fn metadata(&self) -> Result<Metadata> {
        self.inner.metadata()
    }

    /// Stage the complete metadata.
    pub fn set(&mut self, metadata: Metadata) -> &mut Self {
        self.inner.set_metadata(metadata);
        self
    }

    /// Stage a change of the metadata, starting from the staged value.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of the application-properties part.
    pub fn update(&mut self, update: impl FnOnce(&mut Metadata)) -> Result<&mut Self> {
        self.inner.update_metadata(update)?;
        Ok(self)
    }

    /// Stage the complete application properties; `None` removes the part.
    pub fn set_app_props(&mut self, props: Option<app::Props>) -> &mut Self {
        self.inner.set_app_props(props);
        self
    }

    /// The staged custom properties.
    #[must_use]
    pub fn custom_props(&self) -> &custom::Props {
        self.inner.custom_props()
    }

    /// Stage the complete custom properties; an empty collection removes the
    /// part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.inner.set_custom_props(props);
        self
    }

    /// Compute the patch and publish it to a new snapshot.
    ///
    /// An exact no-op returns the same snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, and [`Error::Signed`] for a changed edit of a signed
    /// workbook.
    pub fn commit(self) -> Result<Commit> {
        let (_, patch, diagnostics) = self.inner.commit()?.into_parts();
        let package = match self.owner {
            Owner::Package(package) => package.apply_metadata(&patch)?,
            Owner::Workbook(workbook) => {
                if patch.is_noop() {
                    workbook.to_package()
                } else {
                    let mut candidate = workbook.package.clone();
                    patch.apply(&mut candidate)?;
                    Package::reparse_workbook_candidate(workbook, candidate)
                        .map_err(|error| Error::Invalid(error.to_string()))?
                }
            },
        };
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

/// The result of [`Edit::commit`].
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The snapshot after the edit.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit into the snapshot after the edit.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The exact, reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the edit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit into the snapshot and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    fn metadata_snapshot(&self) -> Result<Snapshot> {
        Snapshot::read(self.opc_package(), Host::Excel)
    }

    /// The document metadata: the core and application properties.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the property parts.
    pub fn metadata(&self) -> Result<Metadata> {
        self.metadata_snapshot()?.metadata()
    }

    /// The OOXML application properties (`docProps/app.xml`), when present.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the part.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        self.metadata_snapshot()?.app_props()
    }

    /// Start a metadata edit.
    ///
    /// # Errors
    ///
    /// Returns the strict graph errors of the property families.
    pub fn edit_metadata(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            owner: Owner::Package(self),
            inner: self.metadata_snapshot()?.edit(),
        })
    }

    /// Apply a metadata patch to a new package; this one is unchanged.
    ///
    /// An empty patch returns this package, a signed one included. The new
    /// package keeps this one's limits and repair policy.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Self> {
        if patch.is_noop() {
            return Ok(self.clone());
        }
        let mut candidate = self.opc_package().clone();
        patch.apply(&mut candidate)?;
        self.reparse_candidate_retaining_projection(candidate)
            .map_err(|error| Error::Invalid(error.to_string()))
    }
}

impl Workbook {
    /// Start a metadata edit of this workbook; [`Edit::commit`] returns the
    /// edited [`Package`], as [`Package::edit_metadata`] does for the
    /// workbook's package, without first copying the workbook into a
    /// package. The workbook is unchanged.
    ///
    /// ```
    /// use litchi_xlsb::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let workbook = Package::new()?.into_workbook()?;
    /// let mut edit = workbook.edit_metadata()?;
    /// edit.update(|metadata| metadata.title = Some("Budget".to_owned()))?;
    /// let after = edit.commit()?.into_package();
    /// assert_eq!(after.metadata()?.title.as_deref(), Some("Budget"));
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the strict graph errors of the property families.
    pub fn edit_metadata(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            owner: Owner::Workbook(self),
            inner: Snapshot::read(&self.package, Host::Excel)?.edit(),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn workbook_edits_equal_package_edits() {
        let package = Package::new().unwrap();
        let workbook = package.clone().into_workbook().unwrap();
        let stage = |edit: &mut Edit<'_>| {
            edit.update(|metadata| {
                metadata.title = Some("Budget".to_owned());
                metadata.company = Some("Litchi".to_owned());
            })
            .unwrap();
        };
        let mut from_package = package.edit_metadata().unwrap();
        stage(&mut from_package);
        let mut from_workbook = workbook.edit_metadata().unwrap();
        stage(&mut from_workbook);
        let (expected, expected_patch) = from_package.commit().unwrap().into_parts();
        let (actual, actual_patch) = from_workbook.commit().unwrap().into_parts();
        assert_eq!(actual.to_bytes().unwrap(), expected.to_bytes().unwrap());
        assert_eq!(actual.metadata().unwrap(), expected.metadata().unwrap());
        assert_eq!(actual_patch.changes().len(), expected_patch.changes().len());
        // The committed package carries its workbook projection.
        let reopened = actual.into_workbook().unwrap();
        assert_eq!(
            reopened.to_package().metadata().unwrap().title.as_deref(),
            Some("Budget")
        );
        // A no-op keeps every byte.
        let edit = workbook.edit_metadata().unwrap();
        let (same, patch) = edit.commit().unwrap().into_parts();
        assert!(patch.is_noop());
        assert_eq!(same.to_bytes().unwrap(), package.to_bytes().unwrap());
    }

    #[test]
    fn edits_publish_a_new_package_and_replay_exactly() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.update(|metadata| {
            metadata.author = Some("Ada".to_owned());
            metadata.company = Some("Litchi".to_owned());
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());
        let (after, patch) = commit.into_parts();
        assert_eq!(after.metadata().unwrap().author.as_deref(), Some("Ada"));
        assert_eq!(
            after.app_props().unwrap().unwrap().company.as_deref(),
            Some("Litchi")
        );
        let replayed = package.apply_metadata(&patch).unwrap();
        assert_eq!(replayed.metadata().unwrap(), after.metadata().unwrap());
        let restored = after.apply_metadata(&patch.inverse()).unwrap();
        assert_eq!(restored.metadata().unwrap(), package.metadata().unwrap());
    }
}
