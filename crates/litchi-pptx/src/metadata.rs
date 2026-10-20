//! Document metadata of a presentation: the core, application and custom
//! property families, edited as one exact and reversible transaction.
//!
//! [`Package::metadata`] reads the unified [`Metadata`];
//! [`Package::edit_metadata`] starts an [`Edit`] whose [`Edit::commit`]
//! returns a [`Commit`] holding the new [`Package`], the [`Patch`] and the
//! [`Diagnostics`]. [`Package::apply_metadata`] replays a patch onto an
//! equal package. The source package never changes.
//!
//! The shared OOXML engine ([`litchi_ooxml_common::properties`]) routes each
//! field to the part that stores it and keeps the lexical form of every
//! unchanged value. A field no OOXML family stores is refused with
//! [`Error::Unsupported`]; a changed edit of a signed package is refused
//! with [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! ```
//! use litchi_core::Metadata;
//! use litchi_pptx::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata()?;
//! edit.update(|metadata| {
//!     metadata.title = Some("Quarterly review".to_owned());
//!     metadata.company = Some("Litchi".to_owned());
//! })?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let metadata: Metadata = commit.package().metadata()?;
//! assert_eq!(metadata.title.as_deref(), Some("Quarterly review"));
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

/// A staged metadata edit of one [`Package`].
///
/// Dropping the edit discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    package: &'a Package,
    inner: properties::Edit,
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

    /// Compute the patch and publish it to a new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, and [`Error::Signed`] for a changed edit of a signed
    /// package.
    pub fn commit(self) -> Result<Commit> {
        let commit = self.inner.commit()?;
        let diagnostics = commit.diagnostics().clone();
        let (opc, patch) = commit.into_parts();
        // An exact no-op publishes the source graph itself.
        let package = self.package.with_metadata_opc(opc);
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
        Snapshot::read(&self.opc, Host::PowerPoint)
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
    /// Returns the strict graph errors of the property families, and refuses
    /// a package with pending authoring state.
    pub fn edit_metadata(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            package: self,
            inner: properties::Edit::new(&self.opc, Host::PowerPoint)?,
        })
    }

    /// Apply a metadata patch to a new package; this package is unchanged.
    ///
    /// An empty patch returns an equal package, a signed one included.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// package.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Self> {
        let mut opc = self.opc.clone();
        patch.apply(&mut opc)?;
        Ok(self.with_metadata_opc(opc))
    }

    /// This package with `opc`, whose document properties changed.
    fn with_metadata_opc(&self, opc: litchi_opc::OpcPackage) -> Self {
        Self {
            opc,
            part_digests: std::sync::OnceLock::new(),
            physical_source_provenance: self.physical_source_provenance,
            tolerance: self.tolerance.clone(),
            #[cfg(feature = "encryption")]
            encryption: self.encryption,
            #[cfg(feature = "encryption")]
            encryption_limits: self.encryption_limits,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn edits_publish_a_new_package_and_replay_exactly() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.update(|metadata| {
            metadata.author = Some("Ada".to_owned());
            metadata.manager = Some("Lin".to_owned());
        })
        .unwrap();
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(after.metadata().unwrap().author.as_deref(), Some("Ada"));
        assert_eq!(
            after.app_props().unwrap().unwrap().manager.as_deref(),
            Some("Lin")
        );
        assert_eq!(package.metadata().unwrap().author, None);

        let replayed = package.apply_metadata(&patch).unwrap();
        assert_eq!(replayed.metadata().unwrap(), after.metadata().unwrap());
        let restored = after.apply_metadata(&patch.inverse()).unwrap();
        assert_eq!(restored.metadata().unwrap(), package.metadata().unwrap());
        assert!(matches!(
            package.apply_metadata(&patch.inverse()),
            Err(Error::Conflict(_))
        ));
    }

    #[test]
    fn unsupported_fields_are_typed_refusals() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.update(|metadata| metadata.codepage = Some(1252))
            .unwrap();
        assert!(matches!(
            edit.commit(),
            Err(Error::Unsupported {
                field: litchi_core::metadata::Field::Codepage
            })
        ));
    }
}
