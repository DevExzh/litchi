use super::super::codec::{
    charge_authored_metadata, invalid, limit, validate_panes, write_add_in_with, write_panes_with,
};
use super::super::model::{Conformance, Limits, OperationBudget, Panes, SnapshotTarget};
use super::super::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, Arc, BTreeSet, Error, HashMap, HashSet, OpcPackage,
    Result, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};
use super::{
    PackageGraphIndex, Patch, PatchPlan, PlannedGraph, PlannedPart, PlannedRelationship,
    RelationshipState, add_or_match_planned_part, existing_web_extension_graph, fold_part_name,
    folded_name_conflicts, graph_matches_plan, has_task_panes_relationship,
    next_package_relationship_id, next_task_panes_part_name, next_web_extension_part_name,
    planned_deletions, preflight_planned_parts, validate_plan_counts,
};
/// Create or replace the package-level persisted task-pane graph.
///
/// Add-in references, bindings, properties, and snapshot resources are stored
/// as inert data. External snapshot links are never contacted.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn put(package: &mut OpcPackage, panes: Panes, conformance: Conformance) -> Result<()> {
    put_with(package, panes, conformance, &Limits::standard())
}

/// Create or replace the task-pane graph with explicit resource limits.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
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
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn plan_put(package: &OpcPackage, panes: Panes, conformance: Conformance) -> Result<Patch> {
    plan_put_with(package, panes, conformance, &Limits::standard())
}

/// Plan a task-pane graph replacement with explicit resource limits.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
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
    drop(task_panes);
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
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    remove_with(package, &Limits::standard())
}

/// Remove the task-pane graph with explicit package graph and deletion ceilings.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn remove_with(package: &mut OpcPackage, limits: &Limits) -> Result<bool> {
    plan_remove_with(package, limits)?.apply(package)
}

/// Plan removal of the package-level task-pane relationship and owned graph.
///
/// An absent graph produces an empty patch, so applying it is a
/// signature-preserving no-op.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn plan_remove(package: &OpcPackage) -> Result<Patch> {
    plan_remove_with(package, &Limits::standard())
}

/// Plan task-pane graph removal with explicit graph and deletion ceilings.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
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
