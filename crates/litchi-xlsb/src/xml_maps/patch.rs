//! Reversible, source-checked publication for XLSB XML Maps.

use litchi_opc::{BlobPart, OpcPackage, Part, TargetMode};

use super::snapshot::{SourcePart, SourceRelationship};
use super::{SemanticLimits, Snapshot};
use crate::package::error::{Error, Result};

/// A reversible replacement of the exact XML Maps-owned package graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before publication.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by publication.
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether applying this patch changes no owned source byte or topology.
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return an inverse that restores owned bytes and topology, never signatures.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone().without_signatures(),
        }
    }

    pub(crate) fn limits(&self) -> SemanticLimits {
        self.before.limits()
    }

    pub(crate) fn check_source(
        &self,
        package: &OpcPackage,
        worksheets: Vec<litchi_opc::PackURI>,
    ) -> Result<Snapshot> {
        let current = Snapshot::read_for_worksheets(package, worksheets, self.before.limits())?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict { part: "XML Maps" });
        }
        Ok(current)
    }

    pub(crate) fn materialize(&self, package: &mut OpcPackage) -> Result<()> {
        if self.is_noop() {
            return Ok(());
        }
        let before = self.before.source();
        let after = self.after.source();

        for dependency in &after.dependencies {
            if !before
                .dependencies
                .iter()
                .any(|candidate| candidate.part_name.is_equivalent_to(&dependency.part_name))
            {
                refuse_inbound(package, &dependency.part_name)?;
            }
        }

        restore_part(package, &after.workbook)?;
        for worksheet in &after.worksheets {
            restore_part(package, worksheet)?;
        }
        for part in &before.dependencies {
            if !after
                .dependencies
                .iter()
                .any(|candidate| candidate.part_name == part.part_name)
            {
                refuse_inbound(package, &part.part_name)?;
                package.remove_part(&part.part_name);
            }
        }
        for dependency in &after.dependencies {
            restore_part(package, dependency)?;
        }
        Ok(())
    }

    #[cfg(test)]
    pub(crate) fn with_after_fixture(mut self, after: Snapshot) -> Self {
        self.after = after;
        self
    }
}

/// The result of a Custom XML Maps edit's `commit`: the new package, the
/// patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

fn restore_part(package: &mut OpcPackage, source: &SourcePart) -> Result<()> {
    if package.contains_part(&source.part_name) {
        if source.content_type == super::package::WORKSHEET_CONTENT_TYPE {
            let current = package.get_part(&source.part_name)?.blob().to_vec();
            crate::worksheet_index::maintain(
                package,
                &source.part_name,
                &current,
                source.bytes.as_slice(),
            )?;
        }
        let part = package.get_part_mut(&source.part_name)?;
        part.set_content_type(source.content_type.clone())?;
        part.set_blob_shared(source.bytes.clone());
        restore_relationships(part, &source.relationships)?;
    } else {
        let mut part = BlobPart::new_shared(
            source.part_name.clone(),
            source.content_type.clone(),
            source.bytes.clone(),
        );
        restore_relationships(&mut part, &source.relationships)?;
        package.try_add_part(Box::new(part))?;
    }
    Ok(())
}

fn restore_relationships(part: &mut dyn Part, source: &[SourceRelationship]) -> Result<()> {
    let ids = part
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_string())
        .collect::<Vec<_>>();
    for id in ids {
        part.rels_mut().remove(&id);
    }
    for relationship in source {
        part.rels_mut().try_add_relationship(
            relationship.relationship_type.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            relationship.mode,
        )?;
    }
    Ok(())
}

fn refuse_inbound(package: &OpcPackage, target_part: &litchi_opc::PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship_targets(relationship, target_part)? {
            return Err(invalid(
                "cannot remove XML Maps part while another part references it",
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship_targets(relationship, target_part)? {
                return Err(invalid(
                    "cannot remove XML Maps part while another part references it",
                ));
            }
        }
    }
    Ok(())
}

fn relationship_targets(
    relationship: &litchi_opc::Relationship,
    target_part: &litchi_opc::PackURI,
) -> Result<bool> {
    if relationship.target_mode() != TargetMode::Internal {
        return Ok(false);
    }
    let target = relationship
        .target_partname()
        .map_err(|_error| invalid("an XML maps relationship target is not a valid part name"))?;
    Ok(target.is_equivalent_to(target_part))
}

fn invalid(message: &'static str) -> Error {
    Error::InvalidFormat(message)
}
