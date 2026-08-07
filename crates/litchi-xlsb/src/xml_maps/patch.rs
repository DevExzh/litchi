//! Reversible, source-checked publication for XLSB XML Maps.

use litchi_opc::{BlobPart, OpcPackage, Part, TargetMode};

use super::snapshot::{SourcePart, SourceRelationship};
use super::{ReadLimits, Snapshot};
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
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return an inverse that restores owned bytes and topology, never signatures.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone().without_signatures(),
        }
    }

    pub(crate) fn limits(&self) -> ReadLimits {
        self.before.limits()
    }

    pub(crate) fn check_source(
        &self,
        package: &OpcPackage,
        worksheets: Vec<litchi_opc::PackURI>,
    ) -> Result<Snapshot> {
        let current = Snapshot::read_for_worksheets(package, worksheets, self.before.limits())?;
        if !current.same_source(&self.before) {
            return Err(invalid("XML Maps patch source is stale"));
        }
        Ok(current)
    }

    pub(crate) fn materialize(&self, package: &mut OpcPackage) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }
        let before = self.before.source();
        let after = self.after.source();

        for part in &before.dependencies {
            if !after
                .dependencies
                .iter()
                .any(|candidate| candidate.part_name == part.part_name)
            {
                refuse_other_inbound(package, &part.part_name, &before.workbook.part_name)?;
                package.remove_part(&part.part_name);
            }
        }

        restore_part(package, &after.workbook)?;
        for worksheet in &after.worksheets {
            restore_part(package, worksheet)?;
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

/// A detached, source-bound transaction result.
#[derive(Clone, Debug)]
pub struct Commit {
    patch: Patch,
    changed: bool,
}

impl Commit {
    pub(crate) fn new(patch: Patch, changed: bool) -> Self {
        Self { patch, changed }
    }

    /// Whether the transaction changes owned XML Maps source.
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Planned resulting snapshot.
    pub fn snapshot(&self) -> &Snapshot {
        self.patch.after()
    }

    /// Reversible source-checked patch.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume this result into its planned snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        let snapshot = self.patch.after().clone();
        (snapshot, self.patch)
    }
}

fn restore_part(package: &mut OpcPackage, source: &SourcePart) -> Result<()> {
    if package.contains_part(&source.part_name) {
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

fn refuse_other_inbound(
    package: &OpcPackage,
    removed: &litchi_opc::PackURI,
    workbook: &litchi_opc::PackURI,
) -> Result<()> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.target_mode() == TargetMode::Internal
                && relationship.target_partname().ok().as_ref() == Some(removed)
                && !(part.partname() == workbook
                    && matches!(
                        relationship.reltype(),
                        litchi_ooxml_common::spreadsheet_xml_maps::REL
                            | litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL
                    ))
            {
                return Err(invalid(format!(
                    "cannot remove XML Maps part '{}' while another part references it",
                    removed.as_str()
                )));
            }
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
