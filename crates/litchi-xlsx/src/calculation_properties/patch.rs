//! Exact reversible patches for workbook calculation metadata.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::{Error, Result};

/// An exact source-checked workbook XML replacement.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before application.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch preserves the source byte-for-byte.
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply after validating the complete workbook owner state.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        if !self.before.matches_current_source(target) {
            return Err(Error::PatchConflict {
                part: self.before.workbook_part_name().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() {
            return Err(Error::Signed);
        }

        let source = self.after.source_xml();
        let mut output = Vec::new();
        output
            .try_reserve_exact(source.len())
            .map_err(|source| Error::Allocation {
                resource: "calculation metadata patch output",
                source,
            })?;
        output.extend_from_slice(source);

        let mut candidate = target.clone();
        candidate
            .get_part_mut(self.before.workbook_part_name())?
            .set_blob(output);
        let resulting = Snapshot::load_with_limits(&candidate, &self.after.limits())?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: self.after.workbook_part_name().to_string(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful calculation-metadata transaction publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Whether exact authored calculation metadata changed.
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting immutable source-bound state.
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch produced by the transaction.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume this result into its snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calculation_properties::{Properties, Transaction};
    use litchi_opc::TargetMode;
    use litchi_opc::constants::relationship_type as rt;

    fn authored() -> Properties {
        Properties::new().with_calculation_id(Some(17))
    }

    #[test]
    fn inverse_restores_the_exact_source() {
        let mut package = crate::package::build_minimal_package().unwrap();
        let original = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_properties(authored());
        let commit = transaction.commit().unwrap();
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(Snapshot::load(&package).unwrap().same_source(&original));
    }

    #[test]
    fn stale_owner_graph_is_rejected_atomically() {
        let mut package = crate::package::build_minimal_package().unwrap();
        let original = package.clone();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_properties(authored());
        let patch = transaction.commit().unwrap().patch().clone();

        let mut stale = original;
        stale.main_document_part().unwrap();
        stale.rels_mut().remove("rId1").unwrap();
        stale
            .rels_mut()
            .try_add_relationship(
                rt::OFFICE_DOCUMENT.to_owned(),
                "xl/workbook.xml".to_owned(),
                "rIdStale".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let stale_source = stale.main_document_part().unwrap().blob_arc();
        assert!(matches!(
            patch.apply(&mut stale),
            Err(Error::PatchConflict { .. })
        ));
        assert!(std::sync::Arc::ptr_eq(
            &stale_source,
            &stale.main_document_part().unwrap().blob_arc()
        ));
    }
}
