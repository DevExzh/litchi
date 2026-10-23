//! Document metadata of a workbook: the core, application and custom
//! property families, edited as one exact and reversible transaction.
//!
//! [`Package::metadata`] reads the unified [`Metadata`];
//! [`Package::edit_metadata`] starts an [`Edit`] whose [`Edit::commit`]
//! returns a [`Commit`] holding the new package, the [`Patch`] and the
//! [`Diagnostics`]. [`Package::apply_metadata`] replays a patch onto an
//! equal package. The source package never changes.
//!
//! The shared OOXML engine ([`litchi_ooxml_common::properties`]) routes each
//! field to the part that stores it and keeps the lexical form of every
//! unchanged value. A field no OOXML family stores is refused with
//! [`Error::Unsupported`]; a changed edit of a signed workbook is refused
//! with [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! ```
//! use litchi_xlsx::Package;
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

use crate::Package;

type Result<T> = std::result::Result<T, Error>;

/// Bounds for one metadata edit's retained exact patch.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    max_retained_bytes: Option<usize>,
}

impl EditOptions {
    /// Creates an unbounded metadata edit.
    pub const fn new() -> Self {
        Self {
            max_retained_bytes: None,
        }
    }

    /// Sets the exact retained-patch ceiling for this edit.
    pub const fn max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = Some(maximum);
        self
    }

    /// Returns the selected retained-patch ceiling.
    #[must_use]
    pub const fn max_retained_bytes_value(self) -> Option<usize> {
        self.max_retained_bytes
    }
}

/// A staged metadata edit of one [`Package`].
///
/// Dropping the edit discards it.
#[derive(Debug)]
pub struct Edit {
    package: Package,
    inner: properties::Edit,
}

impl Edit {
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
        let commit = self.inner.commit()?;
        let diagnostics = commit.diagnostics().clone();
        let (opc, patch) = commit.into_parts();
        let package = if patch.is_noop() {
            self.package
        } else {
            self.package.adopt_metadata(opc)?
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
    /// The package after the edit.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit into the package after the edit.
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

    /// Consume the commit into the package and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    fn metadata_snapshot(&self) -> Result<Snapshot> {
        Snapshot::read(self.opc(), Host::Excel)
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
    pub fn edit_metadata(&self) -> Result<Edit> {
        self.edit_metadata_with_options(EditOptions::default())
    }

    /// Start a metadata edit under an explicit retained-patch ceiling.
    pub fn edit_metadata_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            package: self.clone(),
            inner: properties::Edit::new_with_options(
                self.opc(),
                Host::Excel,
                options
                    .max_retained_bytes_value()
                    .map_or_else(properties::EditOptions::default, |maximum| {
                        properties::EditOptions::new().max_retained_bytes(maximum)
                    }),
            )?,
        })
    }

    /// Apply a metadata patch to a new package; this one is unchanged.
    ///
    /// An empty patch returns this package, a signed one included. The new
    /// package keeps this one's encryption provenance and repair policy.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit> {
        if patch.is_noop() {
            // An exact no-op still authenticates every recorded property
            // family, but it must not clone the complete OPC graph merely to
            // discover that no bytes will change.
            patch.check(self.opc())?;
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: patch.diagnostics(),
            });
        }
        let mut candidate = self.opc().clone();
        patch.apply(&mut candidate)?;
        let package = self.adopt_metadata(candidate)?;
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: patch.diagnostics(),
        })
    }

    /// Wrap a package whose document properties changed as a new package.
    fn adopt_metadata(&self, candidate: litchi_opc::OpcPackage) -> Result<Self> {
        self.workbook_ref()
            .adopt_published_package(candidate)
            .map(Self::from_workbook)
            .map_err(|error| Error::Invalid(error.to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn edits_publish_a_new_snapshot_and_replay_exactly() {
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
        assert!(!package.same_content(&after).unwrap());

        let replayed = package.apply_metadata(&patch).unwrap();
        assert!(replayed.diagnostics().changed());
        let replayed = replayed.into_package();
        assert!(replayed.same_content(&after).unwrap());
        let restored = after
            .apply_metadata(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(restored.same_content(&package).unwrap());
    }

    #[test]
    fn no_op_edits_return_the_same_snapshot() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.set(package.metadata().unwrap());
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert!(commit.package().same_content(&package).unwrap());
    }

    #[test]
    fn no_op_patch_replay_still_requires_its_recorded_property_scope() {
        let source = Package::new().unwrap();
        let mut edit = source.edit_metadata().unwrap();
        edit.set(source.metadata().unwrap());
        let no_op = edit.commit().unwrap();
        assert!(no_op.patch().is_noop());

        let mut other_edit = source.edit_metadata().unwrap();
        other_edit
            .update(|metadata| metadata.title = Some("Different owner".to_owned()))
            .unwrap();
        let other = other_edit.commit().unwrap().into_package();

        let error = other
            .apply_metadata(no_op.patch())
            .expect_err("a no-op patch must authenticate its recorded base");
        assert!(matches!(error, Error::Conflict(_)), "{error:?}");

        let replayed = source.apply_metadata(no_op.patch()).unwrap();
        assert!(replayed.patch().is_noop());
        assert!(replayed.package().same_content(&source).unwrap());
    }
}
