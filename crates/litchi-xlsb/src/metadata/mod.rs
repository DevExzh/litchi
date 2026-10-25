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
//! [`Error::Refused`]; a changed edit of a signed workbook is refused
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

use crate::package::Package;

pub mod source;

type Result<T> = std::result::Result<T, Error>;

/// The native retained-snapshot policy of one metadata edit.
///
/// The default leaves the package's native metadata policy unconstrained;
/// facade callers select the transaction's remaining semantic budget here,
/// before the edit is committed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    max_snapshot_bytes: Option<usize>,
}

impl EditOptions {
    /// Bound the exact before/after property payload retained by this edit.
    #[must_use]
    pub const fn max_snapshot_bytes(mut self, maximum: usize) -> Self {
        self.max_snapshot_bytes = Some(maximum);
        self
    }

    /// The configured retained-snapshot ceiling, if any.
    #[must_use]
    pub const fn max_snapshot_bytes_value(self) -> Option<usize> {
        self.max_snapshot_bytes
    }
}

fn enforce_snapshot_limit(patch: &Patch, options: EditOptions) -> Result<()> {
    let Some(maximum) = options.max_snapshot_bytes_value() else {
        return Ok(());
    };
    let retained = patch.retained_bytes();
    if retained > maximum {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
            observed: retained,
            maximum,
        });
    }
    Ok(())
}

/// A staged metadata edit of one [`Package`].
///
/// Dropping the edit discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Package,
    inner: properties::Plan,
    options: EditOptions,
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

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.inner = self.inner.base().plan();
    }

    /// Compute the patch and publish it to a new snapshot.
    ///
    /// An exact no-op returns the same snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, and [`Error::Signed`] for a changed edit of a signed
    /// workbook.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            source,
            inner,
            options,
        } = self;
        let (_snapshot, patch, diagnostics) = inner.finish()?;
        enforce_snapshot_limit(&patch, options)?;
        let package = if patch.is_noop() {
            patch.check(source.opc_package())?;
            source
        } else {
            let candidate = patch.apply(source.opc_package())?.into_package();
            source
                .reparse_candidate_retaining_projection(candidate)
                .map_err(|error| Error::Invalid(error.to_string()))?
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
    pub fn edit_metadata(&self) -> Result<Edit> {
        self.edit_metadata_with_options(EditOptions::default())
    }

    /// Start a metadata edit with an explicit retained-snapshot ceiling.
    pub fn edit_metadata_with_options(&self, options: EditOptions) -> Result<Edit> {
        let snapshot = self.metadata_snapshot()?;
        Ok(Edit {
            source: self.clone(),
            inner: snapshot.plan(),
            options,
        })
    }

    /// Apply a metadata patch to a new package; this one is unchanged.
    ///
    /// An empty patch returns this package, a signed one included. The new
    /// package keeps this one's limits and repair policy.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit> {
        let package = if patch.is_noop() {
            patch.check(self.opc_package())?;
            self.clone()
        } else {
            let candidate = patch.apply(self.opc_package())?.into_package();
            self.reparse_candidate_retaining_projection(candidate)
                .map_err(|error| Error::Invalid(error.to_string()))?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: patch.diagnostics(),
        })
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
        let replayed = package.apply_metadata(&patch).unwrap().into_package();
        assert_eq!(replayed.metadata().unwrap(), after.metadata().unwrap());
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        assert!(matches!(
            package.apply_metadata(&patch.inverse()),
            Err(Error::PatchConflict)
        ));
        let restored = after
            .apply_metadata(&patch.inverse())
            .unwrap()
            .into_package();
        assert_eq!(restored.metadata().unwrap(), package.metadata().unwrap());
        assert_eq!(restored.to_bytes().unwrap(), original);
    }

    #[test]
    fn an_equal_metadata_patch_replays_as_an_exact_package_copy() {
        let package = Package::new().unwrap();
        let original = package.to_bytes().unwrap();
        let commit = package.edit_metadata().unwrap().commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());

        let replayed = package.apply_metadata(commit.patch()).unwrap();
        assert!(!replayed.diagnostics().changed());
        assert!(replayed.patch().is_noop());
        assert_eq!(replayed.package().to_bytes().unwrap(), original);
    }

    #[test]
    fn a_foreign_noop_metadata_patch_conflicts_instead_of_bypassing_scope() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.update(|metadata| metadata.title = Some("foreign".to_owned()))
            .unwrap();
        let foreign = edit.commit().unwrap().into_package();
        let foreign_noop = foreign.edit_metadata().unwrap().commit().unwrap();

        assert!(matches!(
            package.apply_metadata(foreign_noop.patch()),
            Err(Error::PatchConflict)
        ));
    }

    #[test]
    fn retained_snapshot_limit_is_checked_before_candidate_publication() {
        let package = Package::new().unwrap();
        let before = package.metadata().unwrap();
        let mut edit = package
            .edit_metadata_with_options(EditOptions::default().max_snapshot_bytes(1))
            .unwrap();
        edit.update(|metadata| metadata.title = Some("Budget".to_owned()))
            .unwrap();
        assert!(matches!(edit.commit(), Err(Error::Limit { .. })));
        assert_eq!(package.metadata().unwrap(), before);
    }
}
