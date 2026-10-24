//! Immutable, source-bound calculation-metadata edits of a
//! [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_calculation_metadata`] starts an [`Edit`] of the
//! workbook part's `calcPr` and calculation features;
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_calculation_metadata`] replays a patch onto an
//! equal package. The source package never changes; an exact no-op returns
//! it, a signed one included, and a changed edit of a signed workbook is
//! refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::calculation_properties::Properties;
//!
//! # fn main() -> litchi_xlsx::Result<()> {
//! let package = Package::new()?;
//! let mut edit = package.edit_calculation_metadata()?;
//! edit.set_properties(Properties::new().with_calculation_id(Some(191_029)))?;
//! let (after, patch) = edit.commit()?.into_parts();
//! let calculation_id = after
//!     .calculation_metadata()?
//!     .properties()
//!     .map(|properties| properties.calculation_id());
//! assert_eq!(calculation_id, Some(191_029));
//! let restored = after.apply_calculation_metadata(&patch.inverse())?.into_package();
//! assert!(restored.calculation_metadata()?.properties().is_none());
//! # Ok(())
//! # }
//! ```

use std::borrow::Cow;

use litchi_opc::OpcPackage;

pub use super::LimitKind;
use super::snapshot::{Snapshot, same_properties};
use super::{Features, Limits, Properties, inspect, rewrite};
use crate::error::{Result, invalid};

crate::record::record_family!("calculation-metadata");

/// The exact, reversible change of a calculation-metadata edit: the
/// workbook part's before and after states.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch(pub(crate) super::Patch);

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.0.is_empty()
    }

    /// The reverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// The calculation metadata the patch requires before application.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The calculation metadata the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        self.0.after()
    }
}

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.0.apply(package)
    }
}

/// The configuration of a calculation-metadata edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Bound the edit by `limits` instead of the package's semantic
    /// calculation-metadata limits
    /// ([`crate::SemanticLimits::calculation_properties`]).
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }
}

/// A staged calculation-metadata edit of one [`crate::Package`]. Dropping it
/// discards it. Starting one retains only the source package and limits; the
/// source graph and semantic draft are materialized by the first getter,
/// mutation or commit.
pub struct Edit {
    target: crate::record::Source,
    loaded: Option<Loaded>,
    limits: Limits,
}

/// The source graph and semantic draft are loaded together on the first
/// operation that needs them. Keeping this as one unit is important: a
/// failed source load leaves `Edit::loaded` empty, so a caller can retry
/// after a transient source/managed-part failure without retaining a partial
/// snapshot or a partially cloned draft.
struct Loaded {
    before: Snapshot,
    properties: Option<Properties>,
    features: Option<Features>,
}

impl crate::Package {
    /// Start a calculation-metadata edit under the package's semantic
    /// limits.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Workbook XML and source-limit errors are deferred to the
    /// first getter, mutation or commit.
    pub fn edit_calculation_metadata(&self) -> Result<Edit> {
        self.edit_calculation_metadata_with_options(EditOptions::default())
    }

    /// Start a calculation-metadata edit with explicit options.
    ///
    /// # Errors
    ///
    /// See [`Self::edit_calculation_metadata`].
    pub fn edit_calculation_metadata_with_options(&self, options: EditOptions) -> Result<Edit> {
        let limits = options
            .limits
            .unwrap_or_else(|| self.semantic_limits().calculation_properties_value());
        let target = crate::record::Source::new(self, "edit_calculation_metadata")?;
        Ok(Edit {
            target,
            loaded: None,
            limits,
        })
    }

    /// Apply an exact calculation-metadata patch to a new package; this one
    /// is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_calculation_metadata(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_calculation_metadata")
    }
}

impl Edit {
    fn ensure_loaded(&mut self) -> Result<&mut Loaded> {
        if self.loaded.is_none() {
            let before = Snapshot::load_with_limits(&self.target, &self.limits)?;
            let properties = before.properties().cloned();
            let features = before.features().cloned();
            self.loaded = Some(Loaded {
                before,
                properties,
                features,
            });
        }
        Ok(self.loaded.as_mut().expect("calculation metadata loaded"))
    }

    /// Materialize and return the exact immutable source used by this edit.
    pub fn before(&mut self) -> Result<&Snapshot> {
        Ok(&self.ensure_loaded()?.before)
    }

    /// Currently staged exact-authored `calcPr` state.
    pub fn properties(&mut self) -> Result<Option<&Properties>> {
        Ok(self.ensure_loaded()?.properties.as_ref())
    }

    /// Currently staged ordered calculation features.
    pub fn features(&mut self) -> Result<Option<&Features>> {
        Ok(self.ensure_loaded()?.features.as_ref())
    }

    /// Replace the staged `calcPr` state.
    pub fn set_properties(&mut self, properties: Properties) -> Result<bool> {
        let loaded = self.ensure_loaded()?;
        if same_properties(loaded.properties.as_ref(), Some(&properties)) {
            return Ok(false);
        }
        loaded.properties = Some(properties);
        Ok(true)
    }

    /// Clone-edit `calcPr`, creating an empty exact-authored value if absent.
    pub fn edit_properties(
        &mut self,
        edit: impl FnOnce(&mut Properties) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self.ensure_loaded()?.properties.clone().unwrap_or_default();
        edit(&mut draft)?;
        self.set_properties(draft)
    }

    /// Remove `calcPr` from the staged workbook.
    pub fn remove_properties(&mut self) -> Result<bool> {
        Ok(self.ensure_loaded()?.properties.take().is_some())
    }

    /// Replace the staged calculation-feature collection.
    pub fn set_features(&mut self, features: Features) -> Result<bool> {
        let loaded = self.ensure_loaded()?;
        if loaded.features.as_ref() == Some(&features) {
            return Ok(false);
        }
        loaded.features = Some(features);
        Ok(true)
    }

    /// Clone-edit the existing calculation-feature collection.
    pub fn edit_features(
        &mut self,
        edit: impl FnOnce(&mut Features) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .ensure_loaded()?
            .features
            .clone()
            .ok_or_else(|| invalid("workbook has no calculation features to edit"))?;
        edit(&mut draft)?;
        self.set_features(draft)
    }

    /// Remove calculation features from the staged workbook.
    pub fn remove_features(&mut self) -> Result<bool> {
        Ok(self.ensure_loaded()?.features.take().is_some())
    }

    /// Whether the exact authored semantic state differs from the source.
    fn changed(loaded: &Loaded) -> bool {
        !same_properties(loaded.before.properties(), loaded.properties.as_ref())
            || loaded.before.features() != loaded.features.as_ref()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        let loaded = self.ensure_loaded()?;
        loaded.properties = loaded.before.properties().cloned();
        loaded.features = loaded.before.features().cloned();
        Ok(())
    }

    /// Validate the staged metadata and publish it as a new package. An
    /// unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook,
    /// [`Error::PatchConflict`] when the source changed, and the rewrite and
    /// readback errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    pub(crate) fn publish(mut self) -> Result<(Commit, Snapshot)> {
        let unchanged = {
            let loaded = self.ensure_loaded()?;
            !Self::changed(loaded)
        };
        if unchanged {
            let loaded = self.loaded.take().expect("calculation metadata loaded");
            let patch = Patch(super::Patch::new(
                loaded.before.clone(),
                loaded.before.clone(),
            ));
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                loaded.before,
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }

        let loaded = self.loaded.take().expect("calculation metadata loaded");
        let Loaded {
            before,
            properties,
            features,
        } = loaded;
        if !before.matches_current_source(&self.target) {
            return Err(Error::PatchConflict {
                part: before.workbook_part_name().to_string(),
            });
        }

        let inspection = inspect(before.source_xml(), &self.limits)?;
        let output = match rewrite(
            &inspection,
            properties.as_ref(),
            features.as_ref(),
            &self.limits,
        )? {
            Cow::Owned(output) => output,
            Cow::Borrowed(_) => {
                return Err(invalid(
                    "changed calculation metadata rewrite produced no output",
                ));
            },
        };

        let mut candidate = OpcPackage::clone(&self.target);
        candidate
            .get_part_mut(before.workbook_part_name())?
            .set_blob(output);
        let snapshot = Snapshot::load_with_limits(&candidate, &self.limits)?;
        if !snapshot.same_semantics(properties.as_ref(), features.as_ref()) {
            return Err(invalid(
                "calculation metadata publication changed the staged semantics",
            ));
        }
        let patch = Patch(super::Patch::new(before, snapshot.clone()));
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
    }
}

/// A raw-graph route over [`Edit`] for the unit tests that author bare
/// graphs.
#[cfg(test)]
pub(crate) struct Transaction<'a> {
    target: crate::record::TestTarget<'a>,
    edit: Edit,
}

#[cfg(test)]
impl<'a> Transaction<'a> {
    pub(crate) fn new(target: impl Into<crate::record::TestTarget<'a>>) -> Result<Self> {
        Self::with_limits(target, &Limits::default())
    }

    pub(crate) fn with_limits(
        target: impl Into<crate::record::TestTarget<'a>>,
        limits: &Limits,
    ) -> Result<Self> {
        let target = target.into();
        let package = target.package()?;
        let edit = package
            .edit_calculation_metadata_with_options(EditOptions::default().limits(*limits))?;
        Ok(Self { target, edit })
    }

    pub(crate) fn commit(self) -> Result<super::Commit> {
        let (commit, snapshot) = self.edit.publish()?;
        let changed = commit.diagnostics().changed();
        let (package, patch) = commit.into_parts();
        self.target.store(package);
        Ok(super::Commit::new(snapshot, patch.0, changed))
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

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::TargetMode;
    use litchi_opc::constants::relationship_type as rt;

    fn properties(id: u32) -> Properties {
        Properties::new().with_calculation_id(Some(id))
    }

    #[test]
    fn no_op_preserves_exact_source_even_when_signed() {
        let mut package = crate::package::build_minimal_package().unwrap();
        package
            .rels_mut()
            .try_add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rIdSignature".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let source = package.main_document_part().unwrap().blob_arc();
        let commit = Transaction::new(&mut package).unwrap().commit().unwrap();
        assert!(!commit.changed());
        assert!(std::sync::Arc::ptr_eq(
            &source,
            &commit.snapshot().source_arc().unwrap()
        ));
        assert!(package.is_signed());
    }

    #[test]
    fn set_and_remove_publish_exact_authored_state() {
        let mut package = crate::package::build_minimal_package().unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.set_properties(properties(91)).unwrap());
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(commit.snapshot().properties().unwrap().calculation_id(), 91);

        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.remove_properties().unwrap());
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert!(commit.snapshot().properties().is_none());
    }

    #[test]
    fn changed_signed_package_is_rejected_without_mutation() {
        let mut package = crate::package::build_minimal_package().unwrap();
        package
            .rels_mut()
            .try_add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rIdSignature".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let source = package.main_document_part().unwrap().blob_arc();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_properties(properties(4)).unwrap();
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert!(std::sync::Arc::ptr_eq(
            &source,
            &package.main_document_part().unwrap().blob_arc()
        ));
        assert!(package.is_signed());
    }
}
