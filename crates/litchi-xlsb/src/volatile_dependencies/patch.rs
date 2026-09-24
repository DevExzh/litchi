//! Reversible, source-checked Volatile Dependencies publication.

use litchi_opc::{BlobPart, OpcPackage};

use super::codec;
use super::model::{Dependencies, ReadLimits};
use super::package::{self, CONTENT_TYPE, Graph, RELATIONSHIP_TYPE};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

/// A reversible replacement of the Workbook-owned Volatile Dependencies part.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    changed: bool,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        let changed = !before.same_source(&after);
        Self {
            before,
            after,
            changed,
        }
    }

    /// Source state required before publication.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by publication.
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch changes owner bytes or graph topology.
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }

    /// Return an exact inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }

    pub(crate) fn limits(&self) -> ReadLimits {
        self.before.limits()
    }

    /// Apply atomically after checking the complete owner source closure.
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        let current = Snapshot::read_with_limits(package, self.limits())?;
        if !current.same_source(&self.before) {
            return Err(Error::InvalidFormat(
                "Volatile Dependencies patch source is stale".to_string(),
            ));
        }
        if self.is_empty() {
            return Ok(current);
        }
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = package.clone();
        materialize_source(&mut candidate, &self.before, &self.after)?;
        let resulting = Snapshot::read_with_limits(&candidate, self.limits())?;
        if !resulting.same_state(&self.after) {
            return Err(Error::InvalidFormat(
                "Volatile Dependencies publication changed the planned semantic or owned graph"
                    .to_string(),
            ));
        }
        *package = candidate;
        Ok(resulting)
    }
}

/// A detached transaction result containing a planned snapshot and patch.
#[derive(Clone, Debug)]
pub struct Commit {
    patch: Patch,
    changed: bool,
}

impl Commit {
    pub(crate) fn new(patch: Patch, changed: bool) -> Self {
        Self { patch, changed }
    }

    /// Whether the transaction changes owner bytes or topology.
    pub const fn changed(&self) -> bool {
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

    /// Consume the commit into its snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        let snapshot = self.patch.after().clone();
        (snapshot, self.patch)
    }
}

pub(crate) fn materialize(
    package: &mut OpcPackage,
    before: &Snapshot,
    dependencies: Option<&Dependencies>,
) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    if workbook_name.as_str() != before.source().workbook.name.as_str() {
        return Err(Error::InvalidFormat(
            "Volatile Dependencies Workbook owner changed".to_string(),
        ));
    }
    let current_graph = package::discover_graph(package, &workbook_name, before.limits())?;
    let expected_graph = graph_from_source(before.source())?;
    if current_graph != expected_graph {
        return Err(Error::InvalidFormat(
            "Volatile Dependencies owner graph changed while staging".to_string(),
        ));
    }
    match (current_graph, dependencies) {
        (Some(graph), Some(dependencies)) => {
            let payload = codec::write(dependencies, before.limits(), before.sheet_count())?;
            let part = package.get_part_mut(&graph.part_name)?;
            if part.content_type() != CONTENT_TYPE || !part.rels().is_empty() {
                return Err(Error::InvalidFormat(
                    "Volatile Dependencies target is no longer relationship-free".to_string(),
                ));
            }
            part.set_blob(payload);
        },
        (Some(graph), None) => {
            package
                .get_part_mut(&workbook_name)?
                .rels_mut()
                .remove(&graph.relationship_id);
            package.remove_part(&graph.part_name);
        },
        (None, Some(dependencies)) => {
            let uri = package::default_part_uri()?;
            package.validate_new_part_name(&uri)?;
            package::reject_inbound_relationships(package, &uri, before.limits())?;
            let payload = codec::write(dependencies, before.limits(), before.sheet_count())?;
            package.try_add_part(Box::new(BlobPart::new(
                uri.clone(),
                CONTENT_TYPE.to_string(),
                payload,
            )))?;
            let target = uri.relative_ref(workbook_name.base_uri());
            package
                .get_part_mut(&workbook_name)?
                .rels_mut()
                .get_or_add(RELATIONSHIP_TYPE, &target);
        },
        (None, None) => {},
    }
    Ok(())
}

fn graph_from_source(source: &super::snapshot::SourceState) -> Result<Option<Graph>> {
    Ok(source.graph.clone())
}

fn materialize_source(package: &mut OpcPackage, before: &Snapshot, after: &Snapshot) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    if workbook_name.as_str() != before.source().workbook.name.as_str() {
        return Err(Error::InvalidFormat(
            "Volatile Dependencies Workbook owner changed".to_string(),
        ));
    }
    let current_graph = package::discover_graph(package, &workbook_name, before.limits())?;
    if current_graph != graph_from_source(before.source())? {
        return Err(Error::InvalidFormat(
            "Volatile Dependencies owner graph changed while applying patch".to_string(),
        ));
    }
    let after_graph = graph_from_source(after.source())?;
    match (current_graph, after_graph.as_ref()) {
        (Some(current), Some(target)) if current.part_name == target.part_name => {
            let after_part = after.source().owner.as_ref().ok_or_else(|| {
                Error::InvalidFormat("Volatile Dependencies target source is missing".to_string())
            })?;
            let part = package.get_part_mut(&current.part_name)?;
            part.set_content_type(after_part.content_type.clone())?;
            part.set_blob_shared(after_part.bytes.clone());
        },
        (Some(current), Some(target)) => {
            remove_owner(package, &workbook_name, &current)?;
            add_owner(package, &workbook_name, target, after, before.limits())?;
        },
        (Some(current), None) => remove_owner(package, &workbook_name, &current)?,
        (None, Some(target)) => add_owner(package, &workbook_name, target, after, before.limits())?,
        (None, None) => {},
    }
    restore_lexical_tokens(package, after.source())?;
    Ok(())
}

fn remove_owner(
    package: &mut OpcPackage,
    workbook: &litchi_opc::PackURI,
    graph: &Graph,
) -> Result<()> {
    package
        .get_part_mut(workbook)?
        .rels_mut()
        .remove(&graph.relationship_id);
    package.remove_part(&graph.part_name);
    Ok(())
}

fn add_owner(
    package: &mut OpcPackage,
    workbook: &litchi_opc::PackURI,
    graph: &Graph,
    after: &Snapshot,
    limits: ReadLimits,
) -> Result<()> {
    let owner = after.source().owner.as_ref().ok_or_else(|| {
        Error::InvalidFormat("Volatile Dependencies target source is missing".to_string())
    })?;
    package.validate_new_part_name(&graph.part_name)?;
    package::reject_inbound_relationships(package, &graph.part_name, limits)?;
    package.try_add_part(Box::new(BlobPart::new_shared(
        graph.part_name.clone(),
        owner.content_type.clone(),
        owner.bytes.clone(),
    )))?;
    package.get_part_mut(workbook)?.rels_mut().add_relationship(
        graph.relationship_type.clone(),
        graph.target_ref.clone(),
        graph.relationship_id.clone(),
        false,
    );
    Ok(())
}

fn restore_lexical_tokens(
    package: &mut OpcPackage,
    source: &super::snapshot::SourceState,
) -> Result<()> {
    let root_name =
        litchi_opc::PackURI::new("/").map_err(|error| Error::InvalidUri(error.to_string()))?;
    let current_root = package.source_relationships(&root_name)?;
    if current_root != source.root_relationships {
        package.try_replace_relationships(&current_root, &source.root_relationships)?;
    }
    let workbook_name = litchi_opc::PackURI::new(&source.workbook.name)
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    let current_workbook = package.source_relationships(&workbook_name)?;
    if current_workbook != source.workbook_relationships {
        package.try_replace_relationships(&current_workbook, &source.workbook_relationships)?;
    }
    if let Some(owner) = source.owner.as_ref() {
        let owner_name = litchi_opc::PackURI::new(&owner.name)
            .map_err(|error| Error::InvalidUri(error.to_string()))?;
        let current_owner = package.source_relationships(&owner_name)?;
        if current_owner != owner.relationships {
            package.try_replace_relationships(&current_owner, &owner.relationships)?;
        }
    }
    let current_content_types = package.source_content_types()?;
    if current_content_types != source.content_types {
        package.try_replace_content_types(current_content_types.bytes(), &source.content_types)?;
    }
    Ok(())
}
