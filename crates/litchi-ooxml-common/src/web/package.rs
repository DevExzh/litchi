//! OPC graph discovery, planning, protection, and transactional application.

use super::codec::*;
use super::model::*;
use super::*;

pub(super) fn has_task_panes_relationship(package: &OpcPackage, limits: &Limits) -> Result<bool> {
    let relationships = package.rels().len();
    if relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            relationships,
        );
    }
    Ok(package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP))
}

/// Resolve and validate the complete package graph with safe default limits.
pub fn load(package: &OpcPackage) -> Result<Option<Panes>> {
    load_with(package, &Limits::standard())
}

/// Resolve and validate the complete package graph with explicit limits.
pub fn load_with(package: &OpcPackage, limits: &Limits) -> Result<Option<Panes>> {
    if !has_task_panes_relationship(package, limits)? {
        return Ok(None);
    }
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    load_with_index_budget(package, limits, &index, &mut budget)
}

pub(super) fn load_with_index_budget(
    package: &OpcPackage,
    limits: &Limits,
    index: &PackageGraphIndex,
    budget: &mut OperationBudget,
) -> Result<Option<Panes>> {
    let relationships: Vec<_> = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .collect();
    if relationships.is_empty() {
        return Ok(None);
    }
    if relationships.len() != 1 {
        return Err(Error::Relationship(
            "package has multiple web extension task-pane relationships".into(),
        ));
    }
    let relationship = relationships[0];
    if relationship.is_external() {
        return Err(Error::Relationship(
            "task-pane relationship must be internal".into(),
        ));
    }
    let target = checked_internal_target(relationship, "task-pane")?;
    let part_name = index
        .canonical(&target)
        .ok_or_else(|| Error::Missing(format!("task-pane part '{}'", target.as_str())))?;
    let task_panes_part = package.get_part(part_name).map_err(|error| {
        Error::Missing(format!("task-pane part '{}': {error}", part_name.as_str()))
    })?;
    require_content_type(task_panes_part, TASK_PANES_CONTENT_TYPE)?;
    let parsed_panes = parse_panes_with_budget(task_panes_part.blob(), limits, budget)?;

    let referenced_ids: HashSet<&str> = parsed_panes
        .iter()
        .map(|pane| pane.relationship_id.as_str())
        .collect();
    if referenced_ids.len() != parsed_panes.len() {
        return invalid("task panes contain duplicate relationship IDs".into());
    }
    for child_relationship in task_panes_part.rels().iter() {
        if child_relationship.reltype() != ADD_IN_RELATIONSHIP {
            return Err(Error::Relationship(format!(
                "task-pane part has forbidden relationship '{}' of type '{}'",
                child_relationship.r_id(),
                child_relationship.reltype()
            )));
        }
        if !referenced_ids.contains(child_relationship.r_id()) {
            return Err(Error::Relationship(format!(
                "task-pane part has unreferenced relationship '{}'",
                child_relationship.r_id()
            )));
        }
    }

    let mut panes = Vec::with_capacity(parsed_panes.len());
    let mut total_snapshot_bytes = 0usize;
    let mut snapshot_names = HashSet::new();
    let mut extension_names = HashSet::new();
    for pane in parsed_panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        if child_relationship.is_external() {
            return Err(Error::Relationship(format!(
                "web extension relationship '{}' must be internal",
                pane.relationship_id
            )));
        }
        let extension_target = checked_internal_target(child_relationship, "add-in")?;
        let extension_name = index.canonical(&extension_target).ok_or_else(|| {
            Error::Missing(format!(
                "web extension part '{}'",
                extension_target.as_str()
            ))
        })?;
        let extension_part = package.get_part(extension_name).map_err(|error| {
            Error::Missing(format!(
                "web extension part '{}': {error}",
                extension_name.as_str()
            ))
        })?;
        let extension_name = extension_part.partname().clone();
        if !extension_names.insert(fold_part_name(&extension_name)) {
            return Err(Error::Relationship(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            )));
        }
        require_content_type(extension_part, ADD_IN_CONTENT_TYPE)?;
        let add_in = parse_add_in_with_budget(extension_part.blob(), limits, budget)?;
        let snapshot_resources = load_snapshot_resources(
            package,
            extension_part,
            &add_in,
            &mut total_snapshot_bytes,
            &mut snapshot_names,
            limits,
            index,
        )?;
        panes.push(Pane {
            dock_state: pane.dock_state,
            visible: pane.visible,
            width: pane.width,
            row: pane.row,
            locked: pane.locked,
            relationship_id: pane.relationship_id,
            add_in,
            snapshot_resources,
            extension_list: pane.extension_list,
        });
    }
    let panes = Panes { panes };
    validate_panes(&panes, limits)?;
    Ok(Some(panes))
}

/// Create or replace the package-level persisted task-pane graph.
///
/// Add-in references, bindings, properties, and snapshot resources are stored
/// as inert data. External snapshot links are never contacted.
pub fn put(package: &mut OpcPackage, panes: Panes, conformance: Conformance) -> Result<()> {
    put_with(package, panes, conformance, &Limits::standard())
}

/// Create or replace the task-pane graph with explicit resource limits.
pub fn put_with(
    package: &mut OpcPackage,
    task_panes: Panes,
    conformance: Conformance,
    limits: &Limits,
) -> Result<()> {
    plan_put_with(package, task_panes, conformance, limits)?
        .apply(package)
        .map(|_| ())
}

/// Plan an exact, source-checked replacement of the persisted task-pane graph.
///
/// The returned patch is opaque: physical part names and relationship IDs stay
/// private to the graph owner. Payload allocations are shared with the patch,
/// and [`Patch::inverse`] does not copy them.
pub fn plan_put(package: &OpcPackage, panes: Panes, conformance: Conformance) -> Result<Patch> {
    plan_put_with(package, panes, conformance, &Limits::standard())
}

/// Plan a task-pane graph replacement with explicit resource limits.
pub fn plan_put_with(
    package: &OpcPackage,
    task_panes: Panes,
    conformance: Conformance,
    limits: &Limits,
) -> Result<Patch> {
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    charge_authored_metadata(&task_panes, &mut budget, limits)?;
    validate_panes(&task_panes, limits)?;
    let task_panes_xml = write_panes_with(&task_panes, conformance, limits)?;
    budget.charge_authored(&task_panes_xml, limits)?;
    let existing = existing_web_extension_graph(package, limits, &index, &mut budget)?;
    let mut allocation_probes = 0usize;
    let task_panes_name = match existing.as_ref() {
        Some(graph) => graph.task_panes_name.clone(),
        None => next_task_panes_part_name(&index, limits, &mut allocation_probes)?,
    };
    let mut reserved_names = BTreeSet::new();
    reserved_names.insert(fold_part_name(&task_panes_name));
    let mut planned = Vec::with_capacity(task_panes.panes.len() + 1);
    let mut planned_by_name = HashMap::with_capacity(task_panes.panes.len() + 1);
    let mut task_relationships = Vec::with_capacity(task_panes.panes.len());
    let mut total_snapshot_bytes = 0usize;
    let mut counted_snapshot_parts = HashSet::new();
    let existing_extensions = existing
        .as_ref()
        .map(|graph| &graph.extensions_by_relationship);

    for (pane_index, pane) in task_panes.panes.iter().enumerate() {
        let extension_name = match existing_extensions
            .and_then(|extensions| extensions.get(&pane.relationship_id))
        {
            Some(name) => name.clone(),
            None => next_web_extension_part_name(
                &index,
                &reserved_names,
                pane_index + 1,
                limits,
                &mut allocation_probes,
            )?,
        };
        let extension_key = fold_part_name(&extension_name);
        if folded_name_conflicts(&reserved_names, &extension_key) {
            return invalid(format!(
                "multiple task panes target web extension part '{}'",
                extension_name.as_str()
            ));
        }
        reserved_names.insert(extension_key);
        let extension_xml = write_add_in_with(&pane.add_in, conformance, limits)?;
        budget.charge_authored(&extension_xml, limits)?;
        let mut relationships = Vec::with_capacity(pane.snapshot_resources.len());
        for resource in &pane.snapshot_resources {
            let (target, external) = match &resource.target {
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => {
                    let part_key = fold_part_name(part_name);
                    let already_counted = counted_snapshot_parts.contains(&part_key);
                    if folded_name_conflicts(&reserved_names, &part_key) && !already_counted {
                        return invalid(format!(
                            "snapshot part '{}' conflicts with another authored part",
                            part_name.as_str()
                        ));
                    }
                    reserved_names.insert(part_key.clone());
                    if counted_snapshot_parts.insert(part_key) {
                        total_snapshot_bytes = total_snapshot_bytes
                            .checked_add(data.len())
                            .ok_or_else(|| {
                                Error::Invalid("aggregate snapshot byte count overflow".into())
                            })?;
                        if total_snapshot_bytes > limits.total_image_bytes {
                            return limit(
                                "aggregate web extension snapshot bytes",
                                limits.total_image_bytes,
                                total_snapshot_bytes,
                            );
                        }
                    }
                    add_or_match_planned_part(
                        &mut planned,
                        &mut planned_by_name,
                        PlannedPart {
                            name: part_name.clone(),
                            content_type: content_type.clone(),
                            data: data.clone(),
                            relationships: Vec::new(),
                        },
                    )?;
                    (part_name.relative_ref(extension_name.base_uri()), false)
                },
                SnapshotTarget::External { target } => (target.clone(), true),
            };
            relationships.push(PlannedRelationship {
                id: resource.relationship_id.clone(),
                relationship_type: conformance.image_relationship_type().into(),
                target,
                external,
            });
        }
        add_or_match_planned_part(
            &mut planned,
            &mut planned_by_name,
            PlannedPart {
                name: extension_name.clone(),
                content_type: ADD_IN_CONTENT_TYPE.into(),
                data: Arc::new(extension_xml),
                relationships,
            },
        )?;
        task_relationships.push(PlannedRelationship {
            id: pane.relationship_id.clone(),
            relationship_type: ADD_IN_RELATIONSHIP.into(),
            target: extension_name.relative_ref(task_panes_name.base_uri()),
            external: false,
        });
    }
    add_or_match_planned_part(
        &mut planned,
        &mut planned_by_name,
        PlannedPart {
            name: task_panes_name.clone(),
            content_type: TASK_PANES_CONTENT_TYPE.into(),
            data: Arc::new(task_panes_xml),
            relationships: task_relationships,
        },
    )?;

    let old_parts = existing
        .as_ref()
        .map_or(&[][..], |graph| graph.owned_parts.as_slice());
    let protected = existing.as_ref().map_or_else(HashSet::new, |graph| {
        index.protected_closure(&graph.owned_parts, &graph.root_relationship_id)
    });
    preflight_planned_parts(package, &index, &planned, old_parts, &protected)?;
    if existing
        .as_ref()
        .is_some_and(|graph| graph_matches_plan(package, &planned, graph, &reserved_names))
    {
        return Ok(Patch::default());
    }
    let deletions = planned_deletions(old_parts, &reserved_names, &protected, limits)?;
    validate_plan_counts(package, &index, &planned, existing.as_ref(), limits)?;
    let root_relationship_id = existing
        .as_ref()
        .map(|graph| graph.root_relationship_id.clone())
        .map_or_else(|| next_package_relationship_id(package, limits), Ok)?;
    let root_before = existing
        .as_ref()
        .map(|graph| {
            package
                .rels()
                .get(&graph.root_relationship_id)
                .map(RelationshipState::capture)
                .ok_or_else(|| {
                    Error::Relationship(
                        "task-pane root relationship disappeared while planning".into(),
                    )
                })
        })
        .transpose()?;
    let root_after = Some(RelationshipState {
        id: root_relationship_id,
        relationship_type: TASK_PANES_RELATIONSHIP.into(),
        target: task_panes_name.as_str().trim_start_matches('/').into(),
        external: false,
    });
    let destination_parts = planned.iter().map(|part| part.name.clone()).collect();
    Patch::planned(
        package,
        PatchPlan {
            before: PlannedGraph {
                root: root_before,
                owned_parts: old_parts.to_vec(),
            },
            after: PlannedGraph {
                root: root_after,
                owned_parts: destination_parts,
            },
            parts: planned,
            deletions,
            limits: *limits,
        },
    )
}

/// Remove the package-level task-pane relationship and graph.
///
/// Parts still referenced elsewhere remain in the package.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    remove_with(package, &Limits::standard())
}

/// Remove the task-pane graph with explicit package graph and deletion ceilings.
pub fn remove_with(package: &mut OpcPackage, limits: &Limits) -> Result<bool> {
    plan_remove_with(package, limits)?.apply(package)
}

/// Plan removal of the package-level task-pane relationship and owned graph.
///
/// An absent graph produces an empty patch, so applying it is a
/// signature-preserving no-op.
pub fn plan_remove(package: &OpcPackage) -> Result<Patch> {
    plan_remove_with(package, &Limits::standard())
}

/// Plan task-pane graph removal with explicit graph and deletion ceilings.
pub fn plan_remove_with(package: &OpcPackage, limits: &Limits) -> Result<Patch> {
    if !has_task_panes_relationship(package, limits)? {
        return Ok(Patch::default());
    }
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    let Some(existing) = existing_web_extension_graph(package, limits, &index, &mut budget)? else {
        return Ok(Patch::default());
    };
    let protected = index.protected_closure(&existing.owned_parts, &existing.root_relationship_id);
    let deletions = planned_deletions(&existing.owned_parts, &BTreeSet::new(), &protected, limits)?;
    let root_before = package
        .rels()
        .get(&existing.root_relationship_id)
        .map(RelationshipState::capture)
        .ok_or_else(|| {
            Error::Relationship("task-pane root relationship disappeared while planning".into())
        })?;
    Patch::planned(
        package,
        PatchPlan {
            before: PlannedGraph {
                root: Some(root_before),
                owned_parts: existing.owned_parts,
            },
            after: PlannedGraph {
                root: None,
                owned_parts: Vec::new(),
            },
            parts: Vec::new(),
            deletions,
            limits: *limits,
        },
    )
}

/// Opaque, exact, reversible task-pane graph transaction.
///
/// A patch records the precise source and destination state of every affected
/// relationship and part. Applying a stale patch fails before mutation. Empty
/// patches preserve package signatures; a changed package is always unsigned.
#[must_use = "a planned Web Extensions patch has no effect until it is applied"]
#[derive(Clone, Default)]
pub struct Patch {
    pub(super) root: Option<RootChange>,
    pub(super) parts: Box<[PartChange]>,
    pub(super) protection: Option<ProtectionGuard>,
}

pub(super) struct PatchPlan {
    pub(super) before: PlannedGraph,
    pub(super) after: PlannedGraph,
    pub(super) parts: Vec<PlannedPart>,
    pub(super) deletions: Vec<PackURI>,
    pub(super) limits: Limits,
}

pub(super) struct PlannedGraph {
    pub(super) root: Option<RelationshipState>,
    pub(super) owned_parts: Vec<PackURI>,
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("empty", &self.is_empty())
            .field("affected_parts", &self.parts.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    pub(super) fn planned(package: &OpcPackage, plan: PatchPlan) -> Result<Self> {
        let PatchPlan {
            before,
            after,
            parts: planned,
            deletions,
            limits,
        } = plan;
        let affected_parts = planned
            .len()
            .checked_add(deletions.len())
            .ok_or(Error::Limit {
                resource: "Web Extensions patch parts",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let mut parts = Vec::with_capacity(affected_parts);
        let mut planned_names = HashSet::with_capacity(planned.len());
        for part in planned {
            let name = part.name.clone();
            planned_names.insert(fold_part_name(&name));
            let before = package.get_part(&name).ok().map(PartState::capture);
            parts.push(PartChange {
                name,
                before,
                after: Some(PartState::from_planned(part)),
            });
        }
        for name in deletions {
            if planned_names.contains(&fold_part_name(&name)) {
                return invalid("planned Web Extensions part is also marked for deletion".into());
            }
            let before = package
                .get_part(&name)
                .ok()
                .map(PartState::capture)
                .ok_or_else(|| {
                    Error::Missing("Web Extensions patch source part disappeared".into())
                })?;
            parts.push(PartChange {
                name,
                before: Some(before),
                after: None,
            });
        }
        parts.sort_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        let protection = ProtectionGuard {
            source: GraphScope {
                root_relationship_id: before.root.as_ref().map(|root| root.id.clone()),
                owned_parts: before.owned_parts.into_boxed_slice(),
            },
            destination: GraphScope {
                root_relationship_id: after.root.as_ref().map(|root| root.id.clone()),
                owned_parts: after.owned_parts.into_boxed_slice(),
            },
            limits,
        };
        let patch = Self {
            root: Some(RootChange {
                before: before.root,
                after: after.root,
            }),
            parts: parts.into_boxed_slice(),
            protection: Some(protection),
        };
        if patch.has_changes() {
            Ok(patch)
        } else {
            Ok(Self::default())
        }
    }

    /// Return whether this patch makes no changes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.has_changes()
    }

    /// Build the exact inverse without copying part payloads.
    #[must_use = "the inverse has no effect until it is applied"]
    pub fn inverse(&self) -> Self {
        Self {
            root: self.root.as_ref().map(|root| RootChange {
                before: root.after.clone(),
                after: root.before.clone(),
            }),
            parts: self
                .parts
                .iter()
                .map(|part| PartChange {
                    name: part.name.clone(),
                    before: part.after.clone(),
                    after: part.before.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            protection: self.protection.as_ref().map(ProtectionGuard::inverse),
        }
    }

    /// Apply this patch after checking its exact source graph.
    ///
    /// Returns `true` when the package changed. Source checks, destination-name
    /// checks, and relationship staging all finish before the first mutation.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<bool> {
        if self.is_empty() {
            return Ok(false);
        }
        self.validate_source(package)?;
        self.validate_protection(package)?;
        self.validate_destination_names(package)?;
        let staged = self
            .parts
            .iter()
            .filter(|part| part.before != part.after)
            .filter_map(|part| {
                part.after
                    .as_ref()
                    .map(|after| after.blob_part(part.name.clone()))
            })
            .collect::<Result<Vec<_>>>()?;

        for part in staged {
            package.add_part(Box::new(part));
        }
        if let Some(root) = self.root.as_ref().filter(|root| root.before != root.after) {
            if let Some(before) = &root.before {
                package.rels_mut().remove(&before.id);
            }
            if let Some(after) = &root.after {
                package.rels_mut().add_relationship(
                    after.relationship_type.clone(),
                    after.target.clone(),
                    after.id.clone(),
                    after.external,
                );
            }
        }
        for part in self
            .parts
            .iter()
            .filter(|part| part.before != part.after && part.after.is_none())
        {
            package.remove_part(&part.name);
        }
        package.unsign();
        Ok(true)
    }

    pub(super) fn has_changes(&self) -> bool {
        self.root
            .as_ref()
            .is_some_and(|root| root.before != root.after)
            || self.parts.iter().any(|part| part.before != part.after)
    }

    pub(super) fn validate_source(&self, package: &OpcPackage) -> Result<()> {
        if let Some(root) = &self.root {
            let mut task_relationships = package
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP);
            let first = task_relationships.next();
            let unique = task_relationships.next().is_none();
            let root_matches = match &root.before {
                Some(before) => first.is_some_and(|actual| before.matches(actual)) && unique,
                None => first.is_none(),
            };
            if !root_matches {
                return Err(Error::Relationship(
                    "Web Extensions patch source relationship changed".into(),
                ));
            }
            if let Some(after) = &root.after {
                let reuses_source_id = root
                    .before
                    .as_ref()
                    .is_some_and(|before| before.id == after.id);
                if !reuses_source_id && package.rels().get(&after.id).is_some() {
                    return Err(Error::Relationship(
                        "Web Extensions patch destination relationship is occupied".into(),
                    ));
                }
            }
        }
        for change in &self.parts {
            let actual = package.get_part(&change.name).ok();
            let matches = match (&change.before, actual) {
                (Some(before), Some(actual)) => before.matches(actual),
                (None, None) => true,
                _ => false,
            };
            if !matches {
                return Err(Error::Relationship(
                    "Web Extensions patch source part changed".into(),
                ));
            }
        }
        Ok(())
    }

    pub(super) fn validate_protection(&self, package: &OpcPackage) -> Result<()> {
        let Some(guard) = &self.protection else {
            return Ok(());
        };
        let changed_existing: HashSet<_> = self
            .parts
            .iter()
            .filter(|part| part.before.is_some() && part.before != part.after)
            .map(|part| fold_part_name(&part.name))
            .collect();
        if !changed_existing.is_empty() {
            let protected = protected_parts(package, &guard.source, &guard.limits)?;
            if changed_existing.iter().any(|name| protected.contains(name)) {
                return Err(Error::Relationship(
                    "Web Extensions patch would change a newly shared source part".into(),
                ));
            }
        }

        let changed_destinations: HashSet<_> = self
            .parts
            .iter()
            .filter(|part| part.after.is_some() && part.before != part.after)
            .map(|part| fold_part_name(&part.name))
            .collect();
        if !changed_destinations.is_empty() {
            let protected = protected_parts(package, &guard.destination, &guard.limits)?;
            if changed_destinations
                .iter()
                .any(|name| protected.contains(name))
            {
                return Err(Error::Relationship(
                    "Web Extensions patch would create a newly shared destination part".into(),
                ));
            }
        }
        Ok(())
    }

    pub(super) fn validate_destination_names(&self, package: &OpcPackage) -> Result<()> {
        let changed = self
            .parts
            .iter()
            .filter(|change| change.before != change.after && change.after.is_some());
        let mut destinations = BTreeSet::new();
        let mut replacements = HashMap::new();
        for change in changed {
            let folded = fold_part_name(&change.name);
            if folded_name_conflicts(&destinations, &folded) {
                return invalid("Web Extensions patch destination parts conflict".into());
            }
            destinations.insert(folded.clone());
            if change.before.is_some() {
                replacements.insert(folded, &change.name);
            }
        }
        if destinations.is_empty() {
            return Ok(());
        }
        for existing in package.iter_parts() {
            let folded = fold_part_name(existing.partname());
            if !folded_name_conflicts(&destinations, &folded) {
                continue;
            }
            let is_replaced_source = replacements
                .get(&folded)
                .is_some_and(|name| *name == existing.partname());
            if !is_replaced_source {
                return invalid("Web Extensions patch destination part is occupied".into());
            }
        }
        Ok(())
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(super) struct RootChange {
    pub(super) before: Option<RelationshipState>,
    pub(super) after: Option<RelationshipState>,
}

#[derive(Clone)]
pub(super) struct ProtectionGuard {
    pub(super) source: GraphScope,
    pub(super) destination: GraphScope,
    pub(super) limits: Limits,
}

impl ProtectionGuard {
    pub(super) fn inverse(&self) -> Self {
        Self {
            source: self.destination.clone(),
            destination: self.source.clone(),
            limits: self.limits,
        }
    }
}

#[derive(Clone)]
pub(super) struct GraphScope {
    pub(super) root_relationship_id: Option<String>,
    pub(super) owned_parts: Box<[PackURI]>,
}

pub(super) fn protected_parts(
    package: &OpcPackage,
    scope: &GraphScope,
    limits: &Limits,
) -> Result<HashSet<String>> {
    let part_count = package.part_count();
    if part_count > limits.package_parts {
        return limit("package parts", limits.package_parts, part_count);
    }
    if scope.owned_parts.is_empty() {
        return Ok(HashSet::new());
    }

    let mut names = Vec::with_capacity(scope.owned_parts.len());
    let mut by_name = HashMap::with_capacity(scope.owned_parts.len());
    for name in &scope.owned_parts {
        let folded = fold_part_name(name);
        if by_name.insert(folded.clone(), names.len()).is_some() {
            return invalid("duplicate part in Web Extensions patch protection scope".into());
        }
        names.push(folded);
    }

    let mut outbound = vec![Vec::new(); names.len()];
    let mut protected = vec![false; names.len()];
    let mut queue = VecDeque::new();
    let mut relationships = 0usize;

    for relationship in package.rels().iter() {
        charge_patch_relationship(&mut relationships, limits)?;
        if relationship.is_external() {
            continue;
        }
        let Ok(target) = relationship.target_partname() else {
            continue;
        };
        let Some(&target) = by_name.get(&fold_part_name(&target)) else {
            continue;
        };
        if scope.root_relationship_id.as_deref() != Some(relationship.r_id()) && !protected[target]
        {
            protected[target] = true;
            queue.push_back(target);
        }
    }

    for part in package.iter_parts() {
        let source = by_name.get(&fold_part_name(part.partname())).copied();
        for relationship in part.rels().iter() {
            charge_patch_relationship(&mut relationships, limits)?;
            if relationship.is_external() {
                continue;
            }
            let Ok(target_name) = relationship.target_partname() else {
                continue;
            };
            let Some(&target) = by_name.get(&fold_part_name(&target_name)) else {
                continue;
            };
            if let Some(source) = source {
                outbound[source].push(target);
            } else if !protected[target] {
                protected[target] = true;
                queue.push_back(target);
            }
        }
    }

    while let Some(source) = queue.pop_front() {
        for &target in &outbound[source] {
            if !protected[target] {
                protected[target] = true;
                queue.push_back(target);
            }
        }
    }

    Ok(names
        .into_iter()
        .enumerate()
        .filter_map(|(index, name)| protected[index].then_some(name))
        .collect())
}

pub(super) fn charge_patch_relationship(relationships: &mut usize, limits: &Limits) -> Result<()> {
    *relationships = relationships.checked_add(1).ok_or(Error::Limit {
        resource: "package relationships",
        max: limits.package_relationships,
        actual: usize::MAX,
    })?;
    if *relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            *relationships,
        );
    }
    Ok(())
}

#[derive(Clone, PartialEq, Eq)]
pub(super) struct PartChange {
    pub(super) name: PackURI,
    pub(super) before: Option<PartState>,
    pub(super) after: Option<PartState>,
}

#[derive(Clone, PartialEq, Eq)]
pub(super) struct PartState {
    pub(super) content_type: String,
    pub(super) data: Arc<Vec<u8>>,
    pub(super) relationships: Box<[RelationshipState]>,
}

impl PartState {
    pub(super) fn capture(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(RelationshipState::capture)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships: relationships.into_boxed_slice(),
        }
    }

    pub(super) fn from_planned(part: PlannedPart) -> Self {
        let mut relationships = part
            .relationships
            .into_iter()
            .map(RelationshipState::from_planned)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type,
            data: part.data,
            relationships: relationships.into_boxed_slice(),
        }
    }

    pub(super) fn matches(&self, part: &dyn Part) -> bool {
        self.content_type == part.content_type()
            && self.data.as_slice() == part.blob()
            && self.relationships.len() == part.rels().len()
            && self.relationships.iter().all(|expected| {
                part.rels()
                    .get(&expected.id)
                    .is_some_and(|actual| expected.matches(actual))
            })
    }

    pub(super) fn blob_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.data));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        Ok(part)
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(super) struct RelationshipState {
    pub(super) id: String,
    pub(super) relationship_type: String,
    pub(super) target: String,
    pub(super) external: bool,
}

impl RelationshipState {
    pub(super) fn capture(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        }
    }

    pub(super) fn from_planned(relationship: PlannedRelationship) -> Self {
        Self {
            id: relationship.id,
            relationship_type: relationship.relationship_type,
            target: relationship.target,
            external: relationship.external,
        }
    }

    pub(super) fn matches(&self, relationship: &litchi_opc::Relationship) -> bool {
        self.id == relationship.r_id()
            && self.relationship_type == relationship.reltype()
            && self.target == relationship.target_ref()
            && self.external == relationship.is_external()
    }
}

#[derive(Debug)]
pub(super) struct ExistingAddInGraph {
    pub(super) root_relationship_id: String,
    pub(super) task_panes_name: PackURI,
    pub(super) extensions_by_relationship: HashMap<String, PackURI>,
    pub(super) owned_parts: Vec<PackURI>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct PlannedRelationship {
    pub(super) id: String,
    pub(super) relationship_type: String,
    pub(super) target: String,
    pub(super) external: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct PlannedPart {
    pub(super) name: PackURI,
    pub(super) content_type: String,
    pub(super) data: Arc<Vec<u8>>,
    pub(super) relationships: Vec<PlannedRelationship>,
}

#[derive(Debug)]
pub(super) struct IndexedInbound {
    pub(super) source: Option<usize>,
    pub(super) relationship_id: String,
}

#[derive(Debug)]
pub(super) struct IndexedPart {
    pub(super) name: PackURI,
    pub(super) outbound: Vec<usize>,
    pub(super) inbound: Vec<IndexedInbound>,
}

/// One bounded, ASCII-case-folded view of package membership and internal edges.
#[derive(Debug)]
pub(super) struct PackageGraphIndex {
    pub(super) parts: Vec<IndexedPart>,
    pub(super) by_folded: HashMap<String, usize>,
    pub(super) occupied: BTreeSet<String>,
    pub(super) relationships: usize,
}

impl PackageGraphIndex {
    pub(super) fn build(
        package: &OpcPackage,
        limits: &Limits,
        budget: &mut OperationBudget,
    ) -> Result<Self> {
        let part_count = package.part_count();
        if part_count > limits.package_parts {
            return limit("package parts", limits.package_parts, part_count);
        }
        let mut parts: Vec<IndexedPart> = Vec::with_capacity(part_count);
        let mut by_folded = HashMap::with_capacity(part_count);
        let mut occupied = BTreeSet::new();
        for part in package.iter_parts() {
            let metadata_bytes = part
                .partname()
                .as_str()
                .len()
                .checked_add(part.content_type().len())
                .ok_or(Error::Limit {
                    resource: "indexed web extension package metadata bytes",
                    max: limits.total_string_bytes,
                    actual: usize::MAX,
                })?;
            budget.charge_metadata(metadata_bytes, 4, limits)?;
            let folded = fold_part_name(part.partname());
            if let Some(index) = by_folded.insert(folded.clone(), parts.len()) {
                return invalid(format!(
                    "ASCII-case-equivalent package parts '{}' and '{}' coexist",
                    parts[index].name.as_str(),
                    part.partname().as_str()
                ));
            }
            occupied.insert(folded);
            parts.push(IndexedPart {
                name: part.partname().clone(),
                outbound: Vec::new(),
                inbound: Vec::new(),
            });
        }

        let mut value = Self {
            parts,
            by_folded,
            occupied,
            relationships: 0,
        };
        for relationship in package.rels().iter() {
            value.record_relationship(None, relationship, limits, budget)?;
        }
        for part in package.iter_parts() {
            let source = value
                .index_of(part.partname())
                .ok_or_else(|| Error::Missing(part.partname().to_string()))?;
            for relationship in part.rels().iter() {
                value.record_relationship(Some(source), relationship, limits, budget)?;
            }
        }
        Ok(value)
    }

    pub(super) fn record_relationship(
        &mut self,
        source: Option<usize>,
        relationship: &litchi_opc::Relationship,
        limits: &Limits,
        budget: &mut OperationBudget,
    ) -> Result<()> {
        let metadata_bytes = relationship
            .r_id()
            .len()
            .checked_add(relationship.reltype().len())
            .and_then(|bytes| bytes.checked_add(relationship.target_ref().len()))
            .ok_or(Error::Limit {
                resource: "indexed web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
        budget.charge_metadata(metadata_bytes, 3, limits)?;
        self.relationships = self.relationships.checked_add(1).ok_or(Error::Limit {
            resource: "package relationships",
            max: limits.package_relationships,
            actual: usize::MAX,
        })?;
        if self.relationships > limits.package_relationships {
            return limit(
                "package relationships",
                limits.package_relationships,
                self.relationships,
            );
        }
        if relationship.is_external() {
            return Ok(());
        }
        let Ok(target) = relationship.target_partname() else {
            // Web graph relationships are rejected with context by their callers.
            return Ok(());
        };
        let Some(target) = self.index_of(&target) else {
            return Ok(());
        };
        if let Some(source) = source {
            self.parts[source].outbound.push(target);
        }
        self.parts[target].inbound.push(IndexedInbound {
            source,
            relationship_id: relationship.r_id().to_owned(),
        });
        Ok(())
    }

    pub(super) fn index_of(&self, name: &PackURI) -> Option<usize> {
        self.by_folded.get(&fold_part_name(name)).copied()
    }

    pub(super) fn canonical(&self, name: &PackURI) -> Option<&PackURI> {
        self.index_of(name).map(|index| &self.parts[index].name)
    }

    pub(super) fn contains(&self, name: &PackURI) -> bool {
        self.index_of(name).is_some()
    }

    pub(super) fn conflicts(&self, candidate: &PackURI) -> bool {
        let folded = fold_part_name(candidate);
        if self.occupied.contains(&folded) {
            return true;
        }
        let mut ancestor = folded.as_str();
        while let Some(index) = ancestor.rfind('/') {
            if index == 0 {
                break;
            }
            ancestor = &ancestor[..index];
            if self.occupied.contains(ancestor) {
                return true;
            }
        }
        let descendant_prefix = format!("{folded}/");
        self.occupied
            .range(descendant_prefix.clone()..)
            .next()
            .is_some_and(|name| name.starts_with(&descendant_prefix))
    }

    pub(super) fn protected_closure(
        &self,
        owned_parts: &[PackURI],
        root_relationship_id: &str,
    ) -> HashSet<String> {
        let owned: HashSet<_> = owned_parts
            .iter()
            .filter_map(|name| self.index_of(name))
            .collect();
        let mut queue = VecDeque::new();
        let mut protected = HashSet::new();
        for &index in &owned {
            let has_external_ingress =
                self.parts[index]
                    .inbound
                    .iter()
                    .any(|inbound| match inbound.source {
                        None => inbound.relationship_id != root_relationship_id,
                        Some(source) => !owned.contains(&source),
                    });
            if has_external_ingress && protected.insert(index) {
                queue.push_back(index);
            }
        }
        while let Some(source) = queue.pop_front() {
            for &target in &self.parts[source].outbound {
                if owned.contains(&target) && protected.insert(target) {
                    queue.push_back(target);
                }
            }
        }
        protected
            .into_iter()
            .map(|index| fold_part_name(&self.parts[index].name))
            .collect()
    }
}

pub(super) fn fold_part_name(name: &PackURI) -> String {
    name.as_str().to_ascii_lowercase()
}

pub(super) fn existing_web_extension_graph(
    package: &OpcPackage,
    limits: &Limits,
    index: &PackageGraphIndex,
    budget: &mut OperationBudget,
) -> Result<Option<ExistingAddInGraph>> {
    let Some(loaded) = load_with_index_budget(package, limits, index, budget)? else {
        return Ok(None);
    };
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .ok_or_else(|| {
            Error::Relationship("loaded task panes have no package relationship".into())
        })?;
    let task_panes_target = checked_internal_target(relationship, "task-pane")?;
    let task_panes_name = index
        .canonical(&task_panes_target)
        .ok_or_else(|| Error::Missing(task_panes_target.to_string()))?
        .clone();
    let task_panes_part = package.get_part(&task_panes_name)?;
    let mut extensions_by_relationship = HashMap::with_capacity(loaded.panes.len());
    let mut owned = HashSet::new();
    owned.insert(task_panes_name.clone());
    for pane in loaded.panes {
        let child_relationship = task_panes_part
            .rels()
            .get(&pane.relationship_id)
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "task pane references missing relationship '{}'",
                    pane.relationship_id
                ))
            })?;
        let extension_target = checked_internal_target(child_relationship, "add-in")?;
        let extension_name = index
            .canonical(&extension_target)
            .ok_or_else(|| Error::Missing(extension_target.to_string()))?
            .clone();
        if extensions_by_relationship
            .insert(pane.relationship_id.clone(), extension_name.clone())
            .is_some()
        {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                pane.relationship_id
            ));
        }
        owned.insert(extension_name);
        for resource in pane.snapshot_resources {
            if let SnapshotTarget::Internal { part_name, .. } = resource.target {
                owned.insert(part_name);
            }
        }
    }
    let mut owned_parts: Vec<_> = owned.into_iter().collect();
    owned_parts.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(Some(ExistingAddInGraph {
        root_relationship_id: relationship.r_id().to_owned(),
        task_panes_name,
        extensions_by_relationship,
        owned_parts,
    }))
}

pub(super) fn next_web_extension_part_name(
    index: &PackageGraphIndex,
    reserved: &BTreeSet<String>,
    preferred_index: usize,
    limits: &Limits,
    allocation_probes: &mut usize,
) -> Result<PackURI> {
    let mut offset = 0usize;
    loop {
        charge_allocation_probe(allocation_probes, limits)?;
        let part_number = preferred_index
            .checked_add(offset)
            .ok_or_else(|| Error::Invalid("web extension part index overflow".into()))?;
        let candidate = PackURI::new(format!("/webextensions/webextension{part_number}.xml"))
            .map_err(Error::Uri)?;
        if !folded_name_conflicts(reserved, &fold_part_name(&candidate))
            && !index.conflicts(&candidate)
        {
            return Ok(candidate);
        }
        offset = offset
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("web extension part index overflow".into()))?;
    }
}

pub(super) fn next_task_panes_part_name(
    index: &PackageGraphIndex,
    limits: &Limits,
    allocation_probes: &mut usize,
) -> Result<PackURI> {
    let mut attempt = 1usize;
    loop {
        charge_allocation_probe(allocation_probes, limits)?;
        let suffix = if attempt == 1 {
            String::new()
        } else {
            attempt.to_string()
        };
        let candidate =
            PackURI::new(format!("/webextensions/taskpanes{suffix}.xml")).map_err(Error::Uri)?;
        if !index.conflicts(&candidate) {
            return Ok(candidate);
        }
        attempt = attempt
            .checked_add(1)
            .ok_or_else(|| Error::Invalid("task-pane part index overflow".into()))?;
    }
}

pub(super) fn charge_allocation_probe(probes: &mut usize, limits: &Limits) -> Result<()> {
    let actual = probes.saturating_add(1);
    if *probes >= limits.part_allocations {
        return limit(
            "web extension part-name allocation probes",
            limits.part_allocations,
            actual,
        );
    }
    *probes = actual;
    Ok(())
}

pub(super) fn next_package_relationship_id(
    package: &OpcPackage,
    limits: &Limits,
) -> Result<String> {
    let attempts = package.rels().len().checked_add(1).ok_or(Error::Limit {
        resource: "web extension relationship IDs",
        max: limits.part_allocations,
        actual: usize::MAX,
    })?;
    let attempts = attempts.min(limits.part_allocations);
    for index in 1..=attempts {
        let candidate = format!("rIdPanes{index}");
        if package.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Relationship(
        "unable to allocate a task-pane relationship ID".into(),
    ))
}

pub(super) fn add_or_match_planned_part(
    planned: &mut Vec<PlannedPart>,
    by_name: &mut HashMap<String, usize>,
    part: PlannedPart,
) -> Result<()> {
    let key = fold_part_name(&part.name);
    if let Some(existing) = by_name.get(&key).and_then(|index| planned.get(*index)) {
        if existing == &part {
            return Ok(());
        }
        return invalid(format!(
            "conflicting authored resources target '{}'",
            part.name.as_str()
        ));
    }
    by_name.insert(key, planned.len());
    planned.push(part);
    Ok(())
}

pub(super) fn preflight_planned_parts(
    package: &OpcPackage,
    index: &PackageGraphIndex,
    planned: &[PlannedPart],
    old_parts: &[PackURI],
    protected: &HashSet<String>,
) -> Result<()> {
    let old_names: HashSet<_> = old_parts.iter().map(fold_part_name).collect();
    let mut planned_names = BTreeSet::new();
    for part in planned {
        let folded = fold_part_name(&part.name);
        if folded_name_conflicts(&planned_names, &folded) {
            return invalid(format!(
                "authored part '{}' conflicts with another planned part",
                part.name.as_str()
            ));
        }
        planned_names.insert(folded.clone());
        if let Some(canonical) = index.canonical(&part.name) {
            if canonical != &part.name {
                return invalid(format!(
                    "authored part name '{}' differs in case from canonical package part '{}'",
                    part.name.as_str(),
                    canonical.as_str()
                ));
            }
            if !old_names.contains(&folded) {
                return invalid(format!(
                    "authored web extension part '{}' already exists outside the replaced graph",
                    part.name.as_str()
                ));
            }
            let existing = package.get_part(canonical)?;
            if existing.content_type() != part.content_type {
                return invalid(format!(
                    "cannot change content type of existing part '{}'",
                    part.name.as_str()
                ));
            }
            if !planned_part_matches_existing(existing, part) && protected.contains(&folded) {
                return Err(Error::Relationship(format!(
                    "cannot replace protected shared web extension part '{}'",
                    part.name.as_str()
                )));
            }
        } else if index.conflicts(&part.name) {
            return invalid(format!(
                "authored part '{}' conflicts with an existing package part",
                part.name.as_str()
            ));
        }
    }
    Ok(())
}

pub(super) fn folded_name_conflicts(names: &BTreeSet<String>, candidate: &str) -> bool {
    if names.contains(candidate) {
        return true;
    }
    let mut ancestor = candidate;
    while let Some(index) = ancestor.rfind('/') {
        if index == 0 {
            break;
        }
        ancestor = &ancestor[..index];
        if names.contains(ancestor) {
            return true;
        }
    }
    let prefix = format!("{candidate}/");
    names
        .range(prefix.clone()..)
        .next()
        .is_some_and(|name| name.starts_with(&prefix))
}

pub(super) fn part_names_conflict(left: &PackURI, right: &PackURI) -> bool {
    let left = fold_part_name(left);
    let right = fold_part_name(right);
    left == right
        || left
            .strip_prefix(&right)
            .is_some_and(|suffix| suffix.starts_with('/'))
        || right
            .strip_prefix(&left)
            .is_some_and(|suffix| suffix.starts_with('/'))
}

pub(super) fn planned_deletions(
    old_parts: &[PackURI],
    retained: &BTreeSet<String>,
    protected: &HashSet<String>,
    limits: &Limits,
) -> Result<Vec<PackURI>> {
    let mut deletions = Vec::new();
    for name in old_parts {
        let folded = fold_part_name(name);
        if protected.contains(&folded) || retained.contains(&folded) {
            continue;
        }
        deletions.push(name.clone());
    }
    if deletions.len() > limits.part_deletions {
        return limit(
            "web extension part deletions",
            limits.part_deletions,
            deletions.len(),
        );
    }
    Ok(deletions)
}

pub(super) fn validate_plan_counts(
    package: &OpcPackage,
    index: &PackageGraphIndex,
    planned: &[PlannedPart],
    existing_graph: Option<&ExistingAddInGraph>,
    limits: &Limits,
) -> Result<()> {
    let new_parts = planned
        .iter()
        .filter(|part| !index.contains(&part.name))
        .count();
    if new_parts > limits.part_allocations {
        return limit(
            "web extension part allocations",
            limits.part_allocations,
            new_parts,
        );
    }
    let peak_parts = index
        .parts
        .len()
        .checked_add(new_parts)
        .ok_or(Error::Limit {
            resource: "package parts",
            max: limits.package_parts,
            actual: usize::MAX,
        })?;
    if peak_parts > limits.package_parts {
        return limit("package parts", limits.package_parts, peak_parts);
    }

    let mut relationships = index.relationships;
    for part in planned {
        if let Some(canonical) = index.canonical(&part.name) {
            relationships = relationships
                .checked_sub(package.get_part(canonical)?.rels().len())
                .ok_or_else(|| Error::Invalid("relationship count underflow".into()))?;
        }
        relationships =
            relationships
                .checked_add(part.relationships.len())
                .ok_or(Error::Limit {
                    resource: "package relationships",
                    max: limits.package_relationships,
                    actual: usize::MAX,
                })?;
    }
    if existing_graph.is_none() {
        relationships = relationships.checked_add(1).ok_or(Error::Limit {
            resource: "package relationships",
            max: limits.package_relationships,
            actual: usize::MAX,
        })?;
    }
    if relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            relationships,
        );
    }
    Ok(())
}

pub(super) fn graph_matches_plan(
    package: &OpcPackage,
    planned: &[PlannedPart],
    existing: &ExistingAddInGraph,
    retained: &BTreeSet<String>,
) -> bool {
    existing.owned_parts.len() == retained.len()
        && existing
            .owned_parts
            .iter()
            .all(|name| retained.contains(&fold_part_name(name)))
        && planned.iter().all(|part| {
            package
                .get_part(&part.name)
                .is_ok_and(|existing_part| planned_part_matches_existing(existing_part, part))
        })
}

pub(super) fn planned_part_matches_existing(existing: &dyn Part, planned: &PlannedPart) -> bool {
    existing.content_type() == planned.content_type
        && existing.blob() == planned.data.as_slice()
        && existing.rels().len() == planned.relationships.len()
        && planned.relationships.iter().all(|planned_relationship| {
            existing
                .rels()
                .get(&planned_relationship.id)
                .is_some_and(|relationship| {
                    relationship.reltype() == planned_relationship.relationship_type
                        && relationship.target_ref() == planned_relationship.target
                        && relationship.is_external() == planned_relationship.external
                })
        })
}
