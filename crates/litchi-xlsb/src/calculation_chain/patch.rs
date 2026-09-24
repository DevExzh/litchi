//! Reversible source-bound Calculation Chain publication.

use litchi_opc::{BlobPart, OpcPackage, TargetMode};
use std::sync::Arc;

use super::package::{CONTENT_TYPE, Graph};
use super::snapshot::{Snapshot, SourcePart, SourceState};
use crate::package::error::{Error, Result};

/// A reversible source-bound Calculation Chain removal.
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

    /// Whether publication changes the owner bytes or graph topology.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }

    /// Return a patch that restores the source state represented by `before`.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }

    pub(crate) fn limits(&self) -> super::ReadLimits {
        self.before.limits()
    }

    /// Check the current package against this patch's exact owner source.
    pub(crate) fn check_source(&self, package: &OpcPackage) -> Result<Snapshot> {
        let probe_target = self
            .after
            .source()
            .graph
            .as_ref()
            .map(|graph| &graph.part_name);
        let current = Snapshot::read_with_probe(package, self.limits(), probe_target)?;
        if !current.same_source(&self.before) {
            return Err(Error::InvalidFormat(
                "Calculation Chain patch source is stale".to_string(),
            ));
        }
        Ok(current)
    }

    pub(crate) fn check_publication_policy(&self, package: &OpcPackage) -> Result<()> {
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        Ok(())
    }

    /// Apply a source-checked patch to a caller-owned candidate.
    ///
    /// The caller must have checked the original package with [`Self::check_source`]
    /// and must pass the resulting snapshot unchanged. The candidate is already
    /// detached from the published package, so this seam performs no second
    /// source capture and no package clone. It still validates the candidate's
    /// complete resulting source state before returning it to the caller.
    pub(crate) fn apply_checked(
        &self,
        candidate: &mut OpcPackage,
        current: Snapshot,
    ) -> Result<Snapshot> {
        if self.is_empty() {
            return Ok(current);
        }
        self.check_publication_policy(candidate)?;

        self.materialize(candidate)?;
        let resulting = Snapshot::read_with_limits(candidate, self.limits())?;
        if !resulting.same_state(&self.after) {
            return Err(Error::InvalidFormat(
                "Calculation Chain publication changed the planned owned graph or source"
                    .to_string(),
            ));
        }
        Ok(resulting)
    }

    pub(crate) fn materialize(&self, package: &mut OpcPackage) -> Result<()> {
        let workbook_name = package.main_document_part()?.partname().clone();
        if !workbook_name.is_equivalent_to(&PackURIRef::new(&self.before.source().workbook.name)?) {
            return Err(Error::InvalidFormat(
                "Calculation Chain Workbook owner changed".to_string(),
            ));
        }
        match (
            self.before.source().graph.as_ref(),
            self.after.source().graph.as_ref(),
        ) {
            (Some(graph), None) => remove_graph(package, &workbook_name, graph)?,
            (None, Some(graph)) => add_graph(
                package,
                &workbook_name,
                graph,
                self.after.source().owner.as_ref(),
            )?,
            (None, None) => {},
            (Some(_), Some(_)) => {
                return Err(Error::UnsupportedFeature(
                    "Calculation Chain payload replacement is unsupported without its binary grammar"
                        .to_string(),
                ));
            },
        }
        restore_lexical_tokens(package, self.after.source(), self.limits())
    }
}

/// A detached result containing a planned snapshot and reversible patch.
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
    #[must_use]
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

fn remove_graph(
    package: &mut OpcPackage,
    workbook_name: &litchi_opc::PackURI,
    graph: &Graph,
) -> Result<()> {
    let removed_relationship = package
        .get_part_mut(workbook_name)?
        .rels_mut()
        .remove(&graph.relationship_id);
    if removed_relationship.is_none() {
        return Err(Error::InvalidFormat(
            "Calculation Chain Workbook relationship is missing during removal".to_string(),
        ));
    }
    if !package.remove_part(&graph.part_name) {
        return Err(Error::InvalidFormat(
            "Calculation Chain physical owner is missing during removal".to_string(),
        ));
    }
    Ok(())
}

fn add_graph(
    package: &mut OpcPackage,
    workbook_name: &litchi_opc::PackURI,
    graph: &Graph,
    owner: Option<&SourcePart>,
) -> Result<()> {
    let owner = owner.ok_or_else(|| {
        Error::InvalidFormat("Calculation Chain source owner is missing during inverse".to_string())
    })?;
    if owner.content_type != CONTENT_TYPE
        || graph.content_type != CONTENT_TYPE
        || owner.content_type != graph.content_type
    {
        return Err(Error::InvalidFormat(
            "Calculation Chain inverse owner is not a relationship-free typed part".to_string(),
        ));
    }
    package.validate_new_part_name(&graph.part_name)?;
    package.try_add_part(Box::new(BlobPart::new_shared(
        graph.part_name.clone(),
        owner.content_type.clone(),
        Arc::clone(&owner.bytes),
    )))?;
    let workbook = package.get_part_mut(workbook_name)?;
    if workbook.rels().get(&graph.relationship_id).is_some() {
        return Err(Error::InvalidFormat(
            "Calculation Chain inverse relationship ID is already present".to_string(),
        ));
    }
    workbook.rels_mut().try_add_relationship(
        graph.relationship_type.clone(),
        graph.target_ref.clone(),
        graph.relationship_id.clone(),
        TargetMode::Internal,
    )?;
    Ok(())
}

pub(crate) fn remove_owner(
    package: &mut OpcPackage,
    source: &SourceState,
    limits: super::ReadLimits,
) -> Result<()> {
    let graph = source.graph.as_ref().ok_or_else(|| {
        Error::InvalidFormat("Calculation Chain owner is already absent".to_string())
    })?;
    let workbook_name = package.main_document_part()?.partname().clone();
    if !workbook_name.is_equivalent_to(&PackURIRef::new(&source.workbook.name)?) {
        return Err(Error::InvalidFormat(
            "Calculation Chain Workbook owner changed".to_string(),
        ));
    }
    let workbook = package.get_part(&workbook_name)?;
    let relationship = workbook.rels().get(&graph.relationship_id).ok_or_else(|| {
        Error::InvalidFormat(
            "Calculation Chain Workbook relationship is missing while removing".to_string(),
        )
    })?;
    if relationship.reltype() != graph.relationship_type
        || relationship.target_ref() != graph.target_ref
        || relationship.target_mode() != TargetMode::Internal
        || !relationship
            .target_partname()?
            .is_equivalent_to(&graph.part_name)
        || package.get_part(&graph.part_name)?.content_type() != CONTENT_TYPE
    {
        return Err(Error::InvalidFormat(
            "Calculation Chain owner graph changed while removing".to_string(),
        ));
    }
    let workbook_relationships = source
        .workbook_relationships
        .without_relationship(&graph.relationship_id, limits.max_part_bytes)?;
    let content_types = source.content_types.without_parts(
        std::slice::from_ref(&graph.part_name),
        limits.max_part_bytes,
    )?;
    let opc_limits = super::snapshot::opc_capture_limits(limits)?;
    remove_graph(package, &workbook_name, graph)?;
    restore_lexical_tokens_from_parts(
        package,
        &source.root_relationships,
        &workbook_relationships,
        None,
        &content_types,
        opc_limits,
    )
}

fn restore_lexical_tokens(
    package: &mut OpcPackage,
    source: &SourceState,
    limits: super::ReadLimits,
) -> Result<()> {
    let opc_limits = super::snapshot::opc_capture_limits(limits)?;
    restore_lexical_tokens_from_parts(
        package,
        &source.root_relationships,
        &source.workbook_relationships,
        source.owner.as_ref().map(|owner| &owner.relationships),
        &source.content_types,
        opc_limits,
    )
}

fn restore_lexical_tokens_from_parts(
    package: &mut OpcPackage,
    root_relationships: &litchi_opc::OwnedRelationships,
    workbook_relationships: &litchi_opc::OwnedRelationships,
    owner_relationships: Option<&litchi_opc::OwnedRelationships>,
    content_types: &litchi_opc::OwnedContentTypes,
    limits: litchi_opc::ReadLimits,
) -> Result<()> {
    let root_name = PackURIRef::new("/")?;
    let current_root = package.source_relationships_with_limits(&root_name, limits)?;
    if current_root != *root_relationships {
        package.try_replace_relationships(&current_root, root_relationships)?;
    }

    let workbook_name = package.main_document_part()?.partname().clone();
    let current_workbook = package.source_relationships_with_limits(&workbook_name, limits)?;
    if current_workbook != *workbook_relationships {
        package.try_replace_relationships(&current_workbook, workbook_relationships)?;
    }

    if let Some(owner_relationships) = owner_relationships {
        let owner_name = owner_relationships.owner().clone();
        let current_owner = package.source_relationships_with_limits(&owner_name, limits)?;
        if current_owner != *owner_relationships {
            package.try_replace_relationships(&current_owner, owner_relationships)?;
        }
    }

    let current_content_types = package.source_content_types_with_limits(limits)?;
    if current_content_types != *content_types {
        package.try_replace_content_types(current_content_types.bytes(), content_types)?;
    }
    Ok(())
}

type PackURIRef = litchi_opc::PackURI;
