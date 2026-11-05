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
//! [`Error::Refused`]; a changed edit of a signed package is refused
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
    /// part when it changes the semantic collection. Setting an equal empty
    /// collection is an exact no-op, preserving a present raw part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.inner.set_custom_props(props);
        self
    }

    /// Stage explicit removal of the custom-properties part and relationship.
    ///
    /// This is distinct from setting an equal empty collection: it removes a
    /// present raw part even when its tolerant semantic projection is empty.
    /// An absent graph remains an exact no-op.
    pub fn remove_custom_props(&mut self) -> &mut Self {
        self.inner.remove_custom_props();
        self
    }

    /// Compute the patch and publish it to a new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, and [`Error::Signed`] for a changed edit of a signed
    /// package.
    pub fn commit(self) -> Result<Commit> {
        let commit = self.inner.commit()?;
        let diagnostics = commit.diagnostics().clone();
        let (opc, patch) = commit.into_parts();
        // An exact no-op publishes the source graph itself.
        let package = if patch.is_noop() {
            self.package.duplicate()
        } else {
            self.package.with_metadata_opc(opc)?
        };
        Ok(Commit::new(package, patch, diagnostics))
    }
}

/// The result of [`Edit::commit`].
pub type Commit = litchi_core::edit::Commit<Package, Patch, Diagnostics>;

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
        self.edit_metadata_with_options(EditOptions::default())
    }

    /// Start a metadata edit under an explicit retained-patch ceiling.
    pub fn edit_metadata_with_options(&self, options: EditOptions) -> Result<Edit<'_>> {
        Ok(Edit {
            package: self,
            inner: properties::Edit::new_with_options(
                &self.opc,
                Host::PowerPoint,
                options
                    .max_retained_bytes_value()
                    .map_or_else(properties::EditOptions::default, |maximum| {
                        properties::EditOptions::new().max_retained_bytes(maximum)
                    }),
            )?,
        })
    }

    /// Apply a metadata patch and return its commit; this package is unchanged.
    ///
    /// An empty patch returns an equal package, a signed one included.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// package.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit> {
        let package = if patch.is_noop() {
            // The shared patch scope authenticates an exact no-op without
            // cloning the package graph. This also preserves signed bytes.
            patch.check(self.opc.as_ref())?;
            self.duplicate()
        } else {
            let opc = patch.apply(self.opc.as_ref())?.into_package();
            self.with_metadata_opc(opc)?
        };
        Ok(Commit::new(package, patch.clone(), patch.diagnostics()))
    }

    /// This package with `opc`, whose document properties changed.
    fn with_metadata_opc(&self, opc: litchi_opc::OpcPackage) -> Result<Self> {
        let part_bytes = Self::part_bytes_for(&opc).map_err(|error| match error {
            crate::Error::Opc(source) => Error::from(litchi_ooxml_common::Error::Opc(source)),
            crate::Error::Internal { reason, .. } => Error::Internal(reason),
            _ => Error::Internal("unexpected PPTX metadata owner byte-accounting failure"),
        })?;
        Ok(self.derive_with_part_bytes(std::sync::Arc::new(opc), part_bytes))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn edits_publish_a_new_package_and_replay_exactly() {
        let package = Package::new().unwrap();
        let original = package.to_bytes().unwrap();
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
        assert!(replayed.diagnostics().changed());
        let replayed = replayed.into_package();
        assert_eq!(replayed.metadata().unwrap(), after.metadata().unwrap());
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        let restored = after
            .apply_metadata(&patch.inverse())
            .unwrap()
            .into_package();
        assert_eq!(restored.metadata().unwrap(), package.metadata().unwrap());
        assert_eq!(restored.to_bytes().unwrap(), original);
        assert!(matches!(
            package.apply_metadata(&patch.inverse()),
            Err(Error::PatchConflict)
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
            Err(Error::Refused(refusal))
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedProperty
        ));
    }

    #[test]
    fn an_unchanged_metadata_patch_replays_as_an_exact_noop() {
        let package = Package::new().unwrap();
        let commit = package.edit_metadata().unwrap().commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        let replayed = package.apply_metadata(commit.patch()).unwrap();
        assert!(!replayed.diagnostics().changed());
        assert!(replayed.patch().is_noop());
        assert_eq!(
            replayed.package().to_bytes().unwrap(),
            package.to_bytes().unwrap()
        );
    }

    #[test]
    fn retained_patch_limit_is_checked_before_metadata_publication() {
        let package = Package::new().unwrap();
        let mut edit = package
            .edit_metadata_with_options(EditOptions::default().max_retained_bytes(1))
            .unwrap();
        edit.update(|metadata| metadata.title = Some("bounded metadata".to_owned()))
            .unwrap();
        assert!(matches!(
            edit.commit(),
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed,
                maximum: 1,
            }) if observed > 1
        ));
        assert_eq!(package.metadata().unwrap().title, None);
    }
}
