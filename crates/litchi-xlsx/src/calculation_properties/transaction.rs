//! Failure-atomic source-bound calculation metadata edits.

use std::borrow::Cow;

use litchi_opc::OpcPackage;

use super::patch::{Commit, Patch};
use super::snapshot::{Snapshot, same_properties};
use super::{Features, Limits, Properties, inspect, rewrite};
use crate::error::{Error, Result, invalid};

/// A staged calculation-metadata edit over one OPC package.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    properties: Option<Properties>,
    features: Option<Features>,
    limits: Limits,
}

impl<'a> Transaction<'a> {
    /// Start a transaction with default calculation-metadata limits.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        Self::with_limits(target, &Limits::default())
    }

    /// Start a transaction with a caller-supplied resource policy.
    pub fn with_limits(target: &'a mut OpcPackage, limits: &Limits) -> Result<Self> {
        let before = Snapshot::load_with_limits(target, limits)?;
        Ok(Self {
            target,
            properties: before.properties().cloned(),
            features: before.features().cloned(),
            before,
            limits: *limits,
        })
    }

    /// Exact immutable source captured when the transaction began.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged exact-authored `calcPr` state.
    #[must_use]
    pub fn properties(&self) -> Option<&Properties> {
        self.properties.as_ref()
    }

    /// Currently staged ordered calculation features.
    #[must_use]
    pub fn features(&self) -> Option<&Features> {
        self.features.as_ref()
    }

    /// Replace the staged `calcPr` state.
    pub fn set_properties(&mut self, properties: Properties) -> bool {
        if same_properties(self.properties(), Some(&properties)) {
            return false;
        }
        self.properties = Some(properties);
        true
    }

    /// Clone-edit `calcPr`, creating an empty exact-authored value if absent.
    pub fn edit_properties(
        &mut self,
        edit: impl FnOnce(&mut Properties) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self.properties.clone().unwrap_or_default();
        edit(&mut draft)?;
        Ok(self.set_properties(draft))
    }

    /// Remove `calcPr` from the staged workbook.
    pub fn remove_properties(&mut self) -> bool {
        self.properties.take().is_some()
    }

    /// Replace the staged calculation-feature collection.
    pub fn set_features(&mut self, features: Features) -> bool {
        if self.features.as_ref() == Some(&features) {
            return false;
        }
        self.features = Some(features);
        true
    }

    /// Clone-edit the existing calculation-feature collection.
    pub fn edit_features(
        &mut self,
        edit: impl FnOnce(&mut Features) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .features
            .clone()
            .ok_or_else(|| invalid("workbook has no calculation features to edit"))?;
        edit(&mut draft)?;
        Ok(self.set_features(draft))
    }

    /// Remove calculation features from the staged workbook.
    pub fn remove_features(&mut self) -> bool {
        self.features.take().is_some()
    }

    /// Whether the exact authored semantic state differs from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !same_properties(self.before.properties(), self.properties())
            || self.before.features() != self.features()
    }

    /// Validate and atomically publish the staged metadata.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }

        if !self.before.matches_current_source(self.target) {
            return Err(Error::PatchConflict {
                part: self.before.workbook_part_name().to_string(),
            });
        }

        let inspection = inspect(self.before.source_xml(), &self.limits)?;
        let output = match rewrite(
            &inspection,
            self.properties.as_ref(),
            self.features.as_ref(),
            &self.limits,
        )? {
            Cow::Owned(output) => output,
            Cow::Borrowed(_) => {
                return Err(invalid(
                    "changed calculation metadata rewrite produced no output",
                ));
            },
        };

        let mut candidate = self.target.clone();
        candidate
            .get_part_mut(self.before.workbook_part_name())?
            .set_blob(output);
        let snapshot = Snapshot::load_with_limits(&candidate, &self.limits)?;
        if !snapshot.same_semantics(self.properties.as_ref(), self.features.as_ref()) {
            return Err(invalid(
                "calculation metadata publication changed the staged semantics",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
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
        assert!(transaction.set_properties(properties(91)));
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(commit.snapshot().properties().unwrap().calculation_id(), 91);

        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.remove_properties());
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
        transaction.set_properties(properties(4));
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert!(std::sync::Arc::ptr_eq(
            &source,
            &package.main_document_part().unwrap().blob_arc()
        ));
        assert!(package.is_signed());
    }
}
