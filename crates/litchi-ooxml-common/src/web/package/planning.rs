use super::super::codec::{
    Node, charge_authored_metadata, invalid, limit, parse_mce_xml,
    splice_add_in_extension_list_with, validate_panes, write_add_in_with, write_panes_with,
};
use super::super::model::{AddIn, Conformance, Limits, OperationBudget, Panes, SnapshotTarget};
use super::super::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, Arc, BTreeSet, Error, HashMap, HashSet,
    IMAGE_RELATIONSHIP_TYPE, OpcPackage, Result, STRICT_IMAGE_RELATIONSHIP_TYPE,
    STRICT_RELATIONSHIPS_NAMESPACE, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
    TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
};
use super::durable::{WebIntent, bind_intent, custom_edit_for_graph};
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
    let existing = existing_web_extension_graph(package, limits, &index, &mut budget)?;
    // Compare the bounded semantic graph before invoking a canonical writer.
    // A source package may differ only in prefixes, quoting, comments, or
    // processing instructions; rewriting those bytes would turn an exact
    // semantic no-op into a signed-package mutation.
    if let Some(graph) = existing.as_ref()
        && graph.panes == task_panes
        && source_matches_conformance(package, graph, conformance, limits)?
    {
        return Ok(bind_intent(
            source_bound_noop_patch(package, graph, limits)?,
            WebIntent::Put {
                panes: task_panes,
                conformance,
            },
        ));
    }
    let source_profile_matches = existing
        .as_ref()
        .map(|graph| source_matches_conformance(package, graph, conformance, limits))
        .transpose()?
        .unwrap_or(false);
    let custom_intent = existing.as_ref().and_then(|graph| {
        if !source_profile_matches
            || !custom_function_graph_only_changed(&graph.panes, &task_panes, limits).ok()?
        {
            return None;
        }
        custom_edit_for_graph(&graph.panes, &task_panes).ok()
    });
    let task_panes_xml = if let Some(graph) = existing.as_ref()
        && source_profile_matches
        && task_pane_xml_fields_equal(&graph.panes, &task_panes)
    {
        package.get_part(&graph.task_panes_name)?.blob_arc()
    } else {
        if let Some(graph) = existing.as_ref()
            && !task_panes_source_rewrite_safe(package, graph, limits)?
        {
            return invalid(
                "cannot rewrite task-pane XML without discarding noncanonical or opaque source markup"
                    .into(),
            );
        }
        Arc::new(write_panes_with(&task_panes, conformance, limits)?)
    };
    budget.charge_authored(task_panes_xml.as_slice(), limits)?;
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
        let extension_data = if let Some(graph) = existing.as_ref()
            && let Some(old_pane) = graph
                .panes
                .panes
                .iter()
                .find(|old| old.relationship_id == pane.relationship_id)
        {
            let source_name = graph
                .extensions_by_relationship
                .get(&pane.relationship_id)
                .ok_or_else(|| {
                    Error::Missing(format!(
                        "web extension relationship '{}'",
                        pane.relationship_id
                    ))
                })?;
            let source_part = package.get_part(source_name)?;
            if old_pane.add_in == pane.add_in && source_profile_matches {
                source_part.blob_arc()
            } else if source_profile_matches
                && custom_function_add_in_only_changed(&old_pane.add_in, &pane.add_in, limits)?
            {
                let extension_list = pane.add_in.extension_list.as_ref().ok_or_else(|| {
                    Error::Invalid("custom-function update lost the add-in extLst".into())
                })?;
                Arc::new(splice_add_in_extension_list_with(
                    source_part.blob(),
                    extension_list,
                    limits,
                )?)
            } else {
                if !add_in_source_rewrite_safe(source_part.blob(), &old_pane.add_in, limits)? {
                    return invalid(
                        "cannot rewrite add-in XML without discarding noncanonical or opaque source markup"
                            .into(),
                    );
                }
                Arc::new(write_add_in_with(&pane.add_in, conformance, limits)?)
            }
        } else {
            Arc::new(write_add_in_with(&pane.add_in, conformance, limits)?)
        };
        budget.charge_authored(extension_data.as_slice(), limits)?;
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
                data: extension_data,
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
            data: task_panes_xml,
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
        return Ok(bind_intent(
            source_bound_noop_patch(
                package,
                existing
                    .as_ref()
                    .ok_or_else(|| Error::Invalid("no-op plan lost existing graph".into()))?,
                limits,
            )?,
            WebIntent::Put {
                panes: task_panes,
                conformance,
            },
        ));
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
    let patch = Patch::planned(
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
    )?;
    let intent = custom_intent.map_or(
        WebIntent::Put {
            panes: task_panes,
            conformance,
        },
        WebIntent::CustomFunctions,
    );
    Ok(bind_intent(patch, intent))
}

/// Check only relationship namespace/profile tokens that are semantically
/// authored by this graph. MCE is resolved only to determine the active
/// conformance; no edit offsets are taken from that effective tree.
pub(in crate::web::package) fn source_matches_conformance(
    package: &OpcPackage,
    graph: &super::ExistingAddInGraph,
    conformance: Conformance,
    limits: &Limits,
) -> Result<bool> {
    let expected_namespace = conformance.relationships_namespace();
    let task_panes = package.get_part(&graph.task_panes_name)?;
    if !relationship_attributes_use_namespace(task_panes.blob(), expected_namespace, limits)? {
        return Ok(false);
    }
    for pane in graph.panes.iter() {
        let Some(extension_name) = graph.extensions_by_relationship.get(&pane.relationship_id)
        else {
            return Ok(false);
        };
        let extension = package.get_part(extension_name)?;
        if !relationship_attributes_use_namespace(extension.blob(), expected_namespace, limits)? {
            return Ok(false);
        }
        for relationship in extension.rels().iter() {
            let image_type = relationship.reltype();
            if (image_type == IMAGE_RELATIONSHIP_TYPE
                || image_type == STRICT_IMAGE_RELATIONSHIP_TYPE)
                && image_type != conformance.image_relationship_type()
            {
                return Ok(false);
            }
        }
    }
    Ok(true)
}

/// Build a source-bound no-op patch with complete graph-part snapshots.  The
/// semantic state is unchanged, but `Patch::apply` still validates every owned
/// part and the package root before returning `false`; callers can therefore
/// detect a stale source without mutating or unsigning the package.
pub(in crate::web::package) fn source_bound_noop_patch(
    package: &OpcPackage,
    graph: &super::ExistingAddInGraph,
    limits: &Limits,
) -> Result<Patch> {
    let mut parts = Vec::new();
    parts
        .try_reserve(graph.owned_parts.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions no-op part snapshots",
            source,
        })?;
    for name in &graph.owned_parts {
        let part = package.get_part(name)?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve(part.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions no-op relationship snapshots",
                source,
            })?;
        for relationship in part.rels().iter() {
            relationships.push(PlannedRelationship {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            });
        }
        parts.push(PlannedPart {
            name: name.clone(),
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships,
        });
    }
    let root = package
        .rels()
        .get(&graph.root_relationship_id)
        .map(RelationshipState::capture)
        .ok_or_else(|| {
            Error::Relationship("task-pane root relationship disappeared while planning".into())
        })?;
    Patch::planned(
        package,
        PatchPlan {
            before: PlannedGraph {
                root: Some(root.clone()),
                owned_parts: graph.owned_parts.clone(),
            },
            after: PlannedGraph {
                root: Some(root),
                owned_parts: graph.owned_parts.clone(),
            },
            parts,
            deletions: Vec::new(),
            limits: *limits,
        },
    )
}

pub(in crate::web::package) fn empty_source_bound_noop_patch(
    package: &OpcPackage,
    limits: &Limits,
) -> Result<Patch> {
    Patch::planned(
        package,
        PatchPlan {
            before: PlannedGraph {
                root: None,
                owned_parts: Vec::new(),
            },
            after: PlannedGraph {
                root: None,
                owned_parts: Vec::new(),
            },
            parts: Vec::new(),
            deletions: Vec::new(),
            limits: *limits,
        },
    )
}

fn task_pane_xml_fields_equal(existing: &Panes, incoming: &Panes) -> bool {
    existing.panes.len() == incoming.panes.len()
        && existing
            .panes
            .iter()
            .zip(&incoming.panes)
            .all(|(old, new)| {
                old.dock_state == new.dock_state
                    && old.visible == new.visible
                    && old.width == new.width
                    && old.row == new.row
                    && old.locked == new.locked
                    && old.relationship_id == new.relationship_id
                    && old.extension_list == new.extension_list
            })
}

fn task_panes_source_rewrite_safe(
    package: &OpcPackage,
    graph: &super::ExistingAddInGraph,
    limits: &Limits,
) -> Result<bool> {
    let source = package.get_part(&graph.task_panes_name)?.blob();
    Ok(
        write_panes_with(&graph.panes, Conformance::Transitional, limits)?.as_slice() == source
            || write_panes_with(&graph.panes, Conformance::Strict, limits)?.as_slice() == source,
    )
}

fn add_in_source_rewrite_safe(source: &[u8], add_in: &AddIn, limits: &Limits) -> Result<bool> {
    Ok(
        write_add_in_with(add_in, Conformance::Transitional, limits)?.as_slice() == source
            || write_add_in_with(add_in, Conformance::Strict, limits)?.as_slice() == source,
    )
}

/// Determine whether only typed custom-function metadata changed on one
/// existing add-in.  Unknown source markup is admitted only when the retained
/// custom-payload skeleton is compatible; otherwise the caller gets an atomic
/// refusal instead of a canonical rewrite that drops it.
fn custom_function_add_in_only_changed(
    existing: &AddIn,
    incoming: &AddIn,
    limits: &Limits,
) -> Result<bool> {
    let mut old_without_extension = existing.clone();
    old_without_extension.extension_list = None;
    let mut new_without_extension = incoming.clone();
    new_without_extension.extension_list = None;
    if old_without_extension != new_without_extension {
        return Ok(false);
    }
    match (
        existing.extension_list.as_ref(),
        incoming.extension_list.as_ref(),
    ) {
        (Some(old_extension), Some(new_extension)) => {
            if old_extension.custom_functions() == new_extension.custom_functions() {
                return Ok(old_extension == new_extension);
            }
            if !old_extension.source_compatible_with(new_extension, limits)? {
                return invalid(
                    "cannot source-splice custom-function metadata while retaining opaque known-payload markup"
                        .into(),
                );
            }
            Ok(true)
        },
        _ => Ok(false),
    }
}

fn custom_function_graph_only_changed(
    existing: &Panes,
    incoming: &Panes,
    limits: &Limits,
) -> Result<bool> {
    if existing.panes.len() != incoming.panes.len() {
        return Ok(false);
    }
    let mut changed = false;
    for (old, new) in existing.panes.iter().zip(&incoming.panes) {
        let mut old_without_extension = old.clone();
        old_without_extension.add_in.extension_list = None;
        let mut new_without_extension = new.clone();
        new_without_extension.add_in.extension_list = None;
        if old_without_extension != new_without_extension {
            return Ok(false);
        }
        if old.add_in == new.add_in {
            continue;
        }
        if !custom_function_add_in_only_changed(&old.add_in, &new.add_in, limits)? {
            return Ok(false);
        }
        changed = true;
    }
    Ok(changed)
}

fn relationship_attributes_use_namespace(
    xml: &[u8],
    expected: &str,
    limits: &Limits,
) -> Result<bool> {
    if xml.len() > limits.xml_bytes {
        return limit("web extension XML bytes", limits.xml_bytes, xml.len());
    }
    // Resolve MCE only for this profile check. No source offsets are derived
    // from the effective tree; any actual edit still uses a raw source span
    // or refuses the operation.
    let effective = parse_mce_xml(
        xml,
        &[
            super::super::TASK_PANES_NAMESPACE,
            super::super::WEB_EXTENSION_NAMESPACE,
        ],
        limits,
    )?;
    let root = effective.root()?;
    let mut opposite_seen = false;
    inspect_relationship_attributes(root, expected, &mut opposite_seen);
    // A source with no relationship attributes has no conformance profile to
    // rewrite.  Only an actual expanded-name attribute from the opposite
    // profile requires canonicalization.
    Ok(!opposite_seen)
}

fn inspect_relationship_attributes(node: &Node, expected: &str, opposite_seen: &mut bool) {
    for attribute in &node.attributes {
        if !matches!(attribute.local_name.as_str(), "id" | "embed" | "link") {
            continue;
        }
        if attribute.namespace != expected
            && ((expected == TRANSITIONAL_RELATIONSHIPS_NAMESPACE
                && attribute.namespace == STRICT_RELATIONSHIPS_NAMESPACE)
                || (expected == STRICT_RELATIONSHIPS_NAMESPACE
                    && attribute.namespace == TRANSITIONAL_RELATIONSHIPS_NAMESPACE))
        {
            *opposite_seen = true;
        }
    }
    for child in &node.children {
        inspect_relationship_attributes(child, expected, opposite_seen);
    }
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
        return Ok(bind_intent(
            empty_source_bound_noop_patch(package, limits)?,
            WebIntent::Remove,
        ));
    }
    let mut budget = OperationBudget::default();
    let index = PackageGraphIndex::build(package, limits, &mut budget)?;
    let Some(existing) = existing_web_extension_graph(package, limits, &index, &mut budget)? else {
        return Ok(bind_intent(
            empty_source_bound_noop_patch(package, limits)?,
            WebIntent::Remove,
        ));
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
    let patch = Patch::planned(
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
    )?;
    Ok(bind_intent(patch, WebIntent::Remove))
}
