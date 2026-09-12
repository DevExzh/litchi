//! Bounded ownership planning for Pages body-table deletion.
//!
//! This module is deliberately a topology owner.  It discovers the selected
//! table's physical graph, proves which storage records are private after
//! comparing every rooted sibling table, and emits physical object removals
//! for the later archive/metadata phases.  Formula payload rewrites and
//! package-component registration remain in their sibling phases.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The graph proof keeps its source-index, topology, and census helpers together."
)]

use std::collections::{HashMap, HashSet};
use std::hash::Hash;
use std::num::NonZeroU64;

use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::comment_storage_codec;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;
use litchi_iwa_protos::tst::table_data_list::ListType;
use litchi_numbers_wire::table_data_list;
use litchi_numbers_wire::table_sidecars;

use super::{
    BodyTableDeletionError, DeletionObjectIndex, DeletionObjectLocation, DeletionRequest,
    GraphPlan, ObjectRemoval, Package, RemovalPlan, body_table_catalog, table_lock,
};

const COMMENT_STORAGE_MESSAGE_KIND: u32 = table_sidecars::COMMENT_STORAGE_MESSAGE_KIND;
const RICH_TEXT_PAYLOAD_MESSAGE_KIND: u32 = 6_218;

/// A canonical role in one table's private graph.
///
/// Styles and other optional model references are intentionally absent. They
/// are dependency/context edges, not table-owned storage, and are retained by
/// the package-wide inbound proof unless a future focused owner proves them
/// private explicitly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum OwnershipRole {
    Attachment,
    Drawable,
    Model,
    Tile,
    Header,
    Data,
    UidMap,
    StrokeSidecar,
    Sidecar,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SidecarHint {
    Generic,
    Comment,
    RichText,
}

#[derive(Debug, Clone)]
struct TableGraph {
    target: table_lock::BodyTableTarget,
    roles: HashMap<NonZeroU64, OwnershipRole>,
    sidecar_hints: HashMap<NonZeroU64, SidecarHint>,
    formula_contexts: Vec<NonZeroU64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct PlannedObjectReferenceRemoval {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    target: NonZeroU64,
}

impl TableGraph {
    fn new(target: table_lock::BodyTableTarget) -> Self {
        Self {
            target,
            roles: HashMap::new(),
            sidecar_hints: HashMap::new(),
            formula_contexts: Vec::new(),
        }
    }
}

fn insert_set_with_reserve<T: Eq + Hash>(
    set: &mut HashSet<T>,
    value: T,
) -> Result<bool, BodyTableDeletionError> {
    if set.contains(&value) {
        return Ok(false);
    }
    set.try_reserve(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    Ok(set.insert(value))
}

fn try_clone_identifiers(values: &[NonZeroU64]) -> Result<Vec<NonZeroU64>, BodyTableDeletionError> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(values.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        cloned.push(*value);
    }
    Ok(cloned)
}

/// Prepare the graph half of one deletion transaction.
///
/// The returned [`GraphPlan`] owns only compact identifiers and locations.
/// Every payload projection borrows the immutable package source and charges
/// the caller's cumulative [`table_lock::WireBudget`].  No archive mutation
/// or metadata rewrite occurs here.
pub(super) fn prepare(
    source: &Package,
    request: DeletionRequest,
    budget: &mut table_lock::WireBudget,
) -> Result<GraphPlan, BodyTableDeletionError> {
    // Recheck the selector-owned evidence before building any graph vectors.
    // This keeps the source witness used by the transaction aligned with the
    // one that admitted the request.
    table_lock::validate_body_table_target(source, &request.target, budget)
        .map_err(super::map_lock_error)?;

    let index = build_object_index(source, budget)?;
    let targets = table_lock::body_table_catalog_with_budget(source, budget)
        .map_err(super::map_lock_error)?;
    if targets.is_empty() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    let root_capacity = targets
        .len()
        .checked_mul(4)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let mut protected_roots = HashSet::new();
    protected_roots
        .try_reserve(root_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: root_capacity,
        })?;
    for target in &targets {
        protected_roots.insert(target.body_identifier);
        protected_roots.insert(target.attachment_identifier);
        protected_roots.insert(target.drawable_identifier);
        protected_roots.insert(target.model_identifier);
    }

    let mut source_tables = Vec::new();
    source_tables
        .try_reserve_exact(targets.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: targets.len(),
        })?;
    let mut graphs = Vec::new();
    graphs
        .try_reserve_exact(targets.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: targets.len(),
        })?;

    for target in targets {
        source_tables.push(
            body_table_catalog::BodyTableSnapshot::from_target(target.clone())
                .map_err(super::map_catalog_error)?,
        );
        graphs.push(discover_table_graph(
            source,
            &index,
            target,
            &protected_roots,
            budget,
        )?);
    }

    let selected_position = graphs
        .iter()
        .position(|graph| graph.target == request.target)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let selected = graphs
        .get(selected_position)
        .ok_or(BodyTableDeletionError::InvalidSource)?;

    // Every selected root must remain private to this body-table attachment.
    // A sibling root alias is an invalid source even when the payload bytes
    // happen to be identical: removing the object would redirect that table.
    let selected_roots = [
        request.target.attachment_identifier,
        request.target.drawable_identifier,
        request.target.model_identifier,
    ];
    for (sibling_index, sibling) in graphs.iter().enumerate() {
        if sibling_index == selected_position {
            continue;
        }
        if selected_roots
            .iter()
            .any(|identifier| sibling.roles.contains_key(identifier))
        {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    let mut shared = HashSet::new();
    shared
        .try_reserve(selected.roles.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: selected.roles.len(),
        })?;
    for (sibling_index, sibling) in graphs.iter().enumerate() {
        if sibling_index == selected_position {
            continue;
        }
        for (identifier, role) in &sibling.roles {
            if let Some(selected_role) = selected.roles.get(identifier) {
                if selected_role != role {
                    // A physical object cannot be both canonical table data
                    // and a root/sidecar under two roles.  Fail closed rather
                    // than guessing which ownership edge should win.
                    return Err(BodyTableDeletionError::InvalidSource);
                }
                shared.insert(*identifier);
            }
        }
    }
    let mut private_ids = Vec::new();
    private_ids
        .try_reserve_exact(selected.roles.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: selected.roles.len(),
        })?;
    for (identifier, role) in &selected.roles {
        if shared.contains(identifier) {
            continue;
        }
        if matches!(
            role,
            OwnershipRole::Attachment
                | OwnershipRole::Drawable
                | OwnershipRole::Model
                | OwnershipRole::Tile
                | OwnershipRole::Header
                | OwnershipRole::Data
                | OwnershipRole::UidMap
                | OwnershipRole::StrokeSidecar
                | OwnershipRole::Sidecar
        ) {
            private_ids.push(*identifier);
        }
    }
    budget
        .charge_sort_work(private_ids.len())
        .map_err(super::map_lock_error)?;
    private_ids.sort_unstable();
    private_ids.dedup();

    let mut removed_ids = HashSet::new();
    removed_ids
        .try_reserve(private_ids.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: private_ids.len(),
        })?;
    for identifier in &private_ids {
        removed_ids.insert(*identifier);
    }
    retain_externally_referenced_storage(
        source,
        &index,
        selected,
        &mut private_ids,
        &mut removed_ids,
        budget,
    )?;
    // Formula contexts are a bounded closure over archive-header object
    // references. Exclude the body and every table root while walking so a
    // selected table cannot accidentally acquire a sibling's formula family
    // merely because the two roots share an appearance or parent edge.
    let mut excluded_contexts = HashSet::new();
    let root_capacity = graphs
        .len()
        .checked_mul(3)
        .and_then(|value| value.checked_add(1))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    excluded_contexts.try_reserve(root_capacity).map_err(|_| {
        BodyTableDeletionError::Allocation {
            amount: root_capacity,
        }
    })?;
    for graph in &graphs {
        excluded_contexts.insert(graph.target.attachment_identifier);
        excluded_contexts.insert(graph.target.drawable_identifier);
        excluded_contexts.insert(graph.target.model_identifier);
    }
    excluded_contexts.insert(request.target.body_identifier);
    for root in selected_roots {
        excluded_contexts.remove(&root);
    }
    let mut formula_contexts = try_clone_identifiers(&selected.formula_contexts)?;
    expand_formula_contexts(
        source,
        &index,
        &mut formula_contexts,
        &excluded_contexts,
        budget,
    )?;
    prove_inbound_references(source, &index, selected, &removed_ids, budget)?;

    let mut removals = RemovalPlan::default();
    removals
        .object_removals
        .try_reserve_exact(private_ids.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: private_ids.len(),
        })?;
    for identifier in private_ids {
        let location = index
            .get(identifier)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        removals.object_removals.push(ObjectRemoval {
            component_index: location.component_index,
            object_index: location.object_index,
            identifier,
        });
    }

    let mut contexts = formula_contexts;
    budget
        .charge_sort_work(contexts.len())
        .map_err(super::map_lock_error)?;
    contexts.sort_unstable();
    contexts.dedup();
    removals
        .formula_contexts
        .try_reserve_exact(contexts.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: contexts.len(),
        })?;
    removals.formula_contexts = contexts;

    Ok(GraphPlan {
        request,
        index,
        source_tables,
        removals,
    })
}

/// Keep a selected non-root storage object when a surviving package object
/// still points at it. A table graph can be shared by non-table features (for
/// example a rich-text style or storage payload), so sibling-table
/// intersection alone is not a complete ownership proof. Once one such
/// object is retained, retain every selected child reachable from it as well.
fn retain_externally_referenced_storage(
    source: &Package,
    index: &DeletionObjectIndex,
    selected: &TableGraph,
    private_ids: &mut Vec<NonZeroU64>,
    removed: &mut HashSet<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let externally_referenced = collect_surviving_inbound_ids(source, index, removed, budget)?;
    let mut retained = HashSet::new();
    retained
        .try_reserve(private_ids.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: private_ids.len(),
        })?;
    let mut pending = Vec::new();
    pending
        .try_reserve(private_ids.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: private_ids.len(),
        })?;

    for identifier in private_ids.iter().copied() {
        let Some(role) = selected.roles.get(&identifier).copied() else {
            return Err(BodyTableDeletionError::InvalidSource);
        };
        if is_mandatory_root(role) {
            continue;
        }
        if externally_referenced.contains(&identifier) {
            insert_set_with_reserve(&mut retained, identifier)?;
            pending.push(identifier);
        }
    }

    while let Some(identifier) = pending.pop() {
        let object = object_at_index(source, index, identifier)?;
        let mut children = Vec::new();
        collect_object_references(object, &mut children, budget)?;
        for child in children {
            if !removed.contains(&child) || retained.contains(&child) {
                continue;
            }
            let Some(role) = selected.roles.get(&child).copied() else {
                return Err(BodyTableDeletionError::InvalidSource);
            };
            if is_mandatory_root(role) {
                continue;
            }
            insert_set_with_reserve(&mut retained, child)?;
            pending
                .try_reserve(1)
                .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
            pending.push(child);
        }
    }

    if retained.is_empty() {
        return Ok(());
    }
    private_ids.retain(|identifier| !retained.contains(identifier));
    for identifier in retained {
        removed.remove(&identifier);
    }
    Ok(())
}

const fn is_mandatory_root(role: OwnershipRole) -> bool {
    matches!(
        role,
        OwnershipRole::Attachment | OwnershipRole::Drawable | OwnershipRole::Model
    )
}

fn collect_surviving_inbound_ids(
    source: &Package,
    index: &DeletionObjectIndex,
    removed: &HashSet<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<HashSet<NonZeroU64>, BodyTableDeletionError> {
    let mut referenced = HashSet::new();
    referenced.try_reserve(index.ordered.len()).map_err(|_| {
        BodyTableDeletionError::Allocation {
            amount: index.ordered.len(),
        }
    })?;
    for location in index.iter() {
        if removed.contains(&location.identifier) {
            continue;
        }
        let object = object_at_index(source, index, location.identifier)?;
        for info in &object.archive_info.message_infos {
            charge_metadata_items(info, budget)?;
            let reference_count = info
                .object_references
                .len()
                .checked_add(
                    info.field_infos
                        .iter()
                        .try_fold(0usize, |count, field| {
                            count.checked_add(field.object_references.len())
                        })
                        .ok_or(BodyTableDeletionError::InvalidSource)?,
                )
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            budget
                .charge_payload_references(reference_count)
                .and_then(|_| budget.charge_payload_work(reference_count))
                .map_err(super::map_lock_error)?;
            for reference in info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter()),
            ) {
                let reference =
                    NonZeroU64::new(*reference).ok_or(BodyTableDeletionError::InvalidSource)?;
                insert_set_with_reserve(&mut referenced, reference)?;
            }
        }
    }
    Ok(referenced)
}

fn build_object_index(
    source: &Package,
    budget: &mut table_lock::WireBudget,
) -> Result<DeletionObjectIndex, BodyTableDeletionError> {
    let components = source.state.source.components();
    let mut object_count = 0usize;
    let mut message_count = 0usize;
    for component in components.iter() {
        object_count = object_count
            .checked_add(component.archive().objects.len())
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        for object in &component.archive().objects {
            message_count = message_count
                .checked_add(object.messages.len())
                .ok_or(BodyTableDeletionError::InvalidSource)?;
        }
    }
    budget
        .charge_payload_objects(object_count)
        .and_then(|_| budget.charge_payload_messages(message_count))
        .and_then(|_| budget.charge_payload_work(object_count.saturating_add(message_count)))
        .map_err(super::map_lock_error)?;

    let mut by_identifier = HashMap::new();
    by_identifier
        .try_reserve(object_count)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: object_count,
        })?;
    let mut ordered = Vec::new();
    ordered
        .try_reserve_exact(object_count)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: object_count,
        })?;

    for (component_index, component) in components.iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            let identifier = object
                .archive_info
                .identifier
                .and_then(NonZeroU64::new)
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            let component_name = clone_component_name(component.name(), budget)?;
            let location = DeletionObjectLocation {
                identifier,
                component_index,
                object_index,
                component_name,
            };
            let ordered_index = ordered.len();
            if by_identifier.insert(identifier, ordered_index).is_some() {
                return Err(BodyTableDeletionError::InvalidSource);
            }
            ordered.push(location);
        }
    }
    Ok(DeletionObjectIndex {
        by_identifier,
        ordered,
    })
}

fn clone_component_name(
    name: &str,
    budget: &mut table_lock::WireBudget,
) -> Result<Box<str>, BodyTableDeletionError> {
    budget
        .charge_payload_work(name.len())
        .map_err(super::map_lock_error)?;
    let mut value = String::new();
    value
        .try_reserve_exact(name.len())
        .map_err(|_| BodyTableDeletionError::Allocation { amount: name.len() })?;
    value.push_str(name);
    Ok(value.into_boxed_str())
}

fn discover_table_graph(
    source: &Package,
    index: &DeletionObjectIndex,
    target: table_lock::BodyTableTarget,
    protected_roots: &HashSet<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<TableGraph, BodyTableDeletionError> {
    let mut graph = TableGraph::new(target.clone());
    insert_role(
        &mut graph,
        index,
        target.attachment_identifier,
        OwnershipRole::Attachment,
        budget,
    )?;
    insert_role(
        &mut graph,
        index,
        target.drawable_identifier,
        OwnershipRole::Drawable,
        budget,
    )?;
    insert_role(
        &mut graph,
        index,
        target.model_identifier,
        OwnershipRole::Model,
        budget,
    )?;

    let info_object = object_at_route(
        source,
        target.component_index,
        target.object_index,
        target.drawable_identifier,
    )?;
    let info = info_object
        .archive_info
        .message_infos
        .get(target.info_message_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    collect_context_references(info, &mut graph.formula_contexts, budget)?;

    let model_object = object_at_route(
        source,
        target.model_component_index,
        target.model_object_index,
        target.model_identifier,
    )?;
    let model_message = model_object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let model_info = model_object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    collect_context_references(model_info, &mut graph.formula_contexts, budget)?;
    graph
        .formula_contexts
        .try_reserve(2)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 2 })?;
    graph.formula_contexts.push(target.drawable_identifier);
    graph.formula_contexts.push(target.model_identifier);

    let mut storage_visitor = StorageRootVisitor::default();
    let model_source = model_message.data.as_slice();
    let options = storage_options(budget, model_source)?;
    let (model_store, report) = match storage::decode_table_model_with_data_store_and_visitor(
        model_source,
        options,
        &mut storage_visitor,
    ) {
        Ok(value) => value,
        Err(error) if error.resource_limit().is_none() => {
            // Buffa-backed native archives can contain historical
            // metadata-only fields that the strict model envelope cannot
            // project. Retry only structural failures through the
            // additive compatibility route; typed resource failures stay
            // visible to the transaction budget. Reset the visitor so a
            // failed speculative decode cannot publish partial roots.
            budget
                .charge_payload_work(model_source.len().max(1))
                .map_err(super::map_lock_error)?;
            storage_visitor = StorageRootVisitor::default();
            let compatibility_options = storage_options(budget, model_source)?;
            storage::decode_table_model_compatibility_with_strict_storage_and_visitor(
                model_source,
                compatibility_options,
                &mut storage_visitor,
            )
            .map_err(map_storage_error)?
        },
        Err(error) => return Err(map_storage_error(error)),
    };
    charge_storage_report(budget, report)?;

    for (identifier, role) in storage_visitor.references {
        insert_role(&mut graph, index, identifier, role, budget)?;
    }
    let store = model_store.data_store();
    insert_storage_reference(
        &mut graph,
        index,
        store.column_headers(),
        OwnershipRole::Header,
        budget,
    )?;
    for reference in [store.string_table(), store.style_table()] {
        insert_storage_reference(&mut graph, index, reference, OwnershipRole::Data, budget)?;
    }
    for reference in [store.formula_table(), store.format_table_pre_bnc()] {
        insert_storage_reference(&mut graph, index, reference, OwnershipRole::Data, budget)?;
    }
    for reference in [
        store.formula_error_table(),
        store.merge_region_map(),
        store.deprecated_custom_format_table(),
        store.multiple_choice_list_format_table(),
        store.rich_text_table(),
        store.conditional_style_table(),
        store.comment_storage_table(),
        store.import_warning_set_table(),
        store.control_cell_spec_table(),
        store.format_table(),
    ]
    .into_iter()
    .flatten()
    {
        insert_storage_reference(&mut graph, index, reference, OwnershipRole::Data, budget)?;
    }
    if let Some(reference) = model_store.model().base_column_row_uids() {
        insert_storage_reference(&mut graph, index, reference, OwnershipRole::UidMap, budget)?;
    }
    if let Some(reference) = model_store.model().stroke_sidecar() {
        insert_storage_reference(
            &mut graph,
            index,
            reference,
            OwnershipRole::StrokeSidecar,
            budget,
        )?;
    }

    walk_data_lists(source, index, &mut graph, budget)?;
    walk_comment_sidecars(source, index, &mut graph, budget)?;
    expand_sidecar_references(source, index, &mut graph, protected_roots, budget)?;
    Ok(graph)
}

fn insert_storage_reference(
    graph: &mut TableGraph,
    index: &DeletionObjectIndex,
    reference: storage::ReferenceSnapshot,
    role: OwnershipRole,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let Some(identifier) = local_storage_identifier(reference)? else {
        // Native Numbers uses the generated proto2 zero reference as the
        // sentinel for an omitted optional route. The host ownership helper
        // deliberately ignores that value before checking its role.
        return Ok(());
    };
    insert_role(graph, index, identifier, role, budget)
}

fn local_storage_identifier(
    reference: storage::ReferenceSnapshot,
) -> Result<Option<NonZeroU64>, BodyTableDeletionError> {
    if reference.deprecated_is_external() == Some(true) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(NonZeroU64::new(reference.identifier()))
}

fn insert_role(
    graph: &mut TableGraph,
    index: &DeletionObjectIndex,
    identifier: NonZeroU64,
    role: OwnershipRole,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    if index.get(identifier).is_none() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    if let Some(existing) = graph.roles.get(&identifier) {
        if *existing != role {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        return Ok(());
    }
    graph
        .roles
        .try_reserve(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    budget
        .charge_payload_items(1)
        .and_then(|_| budget.charge_payload_work(1))
        .map_err(super::map_lock_error)?;
    graph.roles.insert(identifier, role);
    Ok(())
}

fn collect_context_references(
    info: &litchi_iwa_core::MessageInfo,
    contexts: &mut Vec<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    charge_metadata_items(info, budget)?;
    let field_count = info.field_infos.iter().try_fold(0usize, |count, field| {
        count
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
    });
    let field_count = field_count.ok_or(BodyTableDeletionError::InvalidSource)?;
    let references = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|value| value.checked_add(field_count))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_references(references)
        .and_then(|_| budget.charge_payload_work(references))
        .map_err(super::map_lock_error)?;
    contexts
        .try_reserve(references)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: references })?;
    for identifier in &info.object_references {
        contexts.push(NonZeroU64::new(*identifier).ok_or(BodyTableDeletionError::InvalidSource)?);
    }
    for identifier in &info.data_references {
        // Data references are metadata-owned edges, so they do not become
        // formula roots here.  They still participate in source validation;
        // an all-zero or malformed edge must never be silently treated as a
        // missing optional field.
        NonZeroU64::new(*identifier).ok_or(BodyTableDeletionError::InvalidSource)?;
    }
    for field in &info.field_infos {
        for identifier in &field.object_references {
            contexts
                .push(NonZeroU64::new(*identifier).ok_or(BodyTableDeletionError::InvalidSource)?);
        }
        for identifier in &field.data_references {
            NonZeroU64::new(*identifier).ok_or(BodyTableDeletionError::InvalidSource)?;
        }
    }
    Ok(())
}

fn charge_metadata_items(
    info: &litchi_iwa_core::MessageInfo,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let items = 1usize
        .checked_add(info.field_infos.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_items(items)
        .and_then(|_| budget.charge_payload_work(items))
        .map_err(super::map_lock_error)
}

#[derive(Default)]
struct StorageRootVisitor {
    references: Vec<(NonZeroU64, OwnershipRole)>,
}

impl storage::StorageVisitor for StorageRootVisitor {
    fn visit_tile_reference(
        &mut self,
        record: storage::TileReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if record.reference().deprecated_is_external() == Some(true) {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        let Some(identifier) = NonZeroU64::new(record.reference().identifier()) else {
            return Ok(());
        };
        self.references
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(1))?;
        self.references.push((identifier, OwnershipRole::Tile));
        Ok(())
    }

    fn visit_header_bucket(
        &mut self,
        record: storage::ReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if record.reference().deprecated_is_external() == Some(true) {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        let Some(identifier) = NonZeroU64::new(record.reference().identifier()) else {
            return Ok(());
        };
        self.references
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(1))?;
        self.references.push((identifier, OwnershipRole::Header));
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        record: storage::ReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if record.reference().deprecated_is_external() == Some(true) {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        let Some(identifier) = NonZeroU64::new(record.reference().identifier()) else {
            return Ok(());
        };
        self.references
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(1))?;
        self.references.push((identifier, OwnershipRole::Data));
        Ok(())
    }
}

#[derive(Default)]
struct ListEdgeVisitor {
    segments: Vec<NonZeroU64>,
    sidecars: Vec<(NonZeroU64, SidecarHint)>,
}

impl storage::StorageVisitor for ListEdgeVisitor {
    fn visit_list_segment(
        &mut self,
        record: storage::ReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if record.reference().deprecated_is_external() == Some(true) {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        let Some(identifier) = NonZeroU64::new(record.reference().identifier()) else {
            return Ok(());
        };
        self.segments
            .try_reserve(1)
            .map_err(|_| storage::DecodeError::allocation(1))?;
        self.segments.push(identifier);
        Ok(())
    }

    fn visit_list_entry_record(
        &mut self,
        record: storage::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        let entry = record.snapshot();
        if entry.comment_storage().is_some() && entry.ref_count() == 0 {
            return Err(storage::DecodeError::invalid_visitor_result());
        }
        let mut references = [None, None, None];
        let mut count = 0usize;
        if let Some(reference) = entry.reference() {
            references[count] = Some((reference, SidecarHint::Generic));
            count += 1;
        }
        if let Some(reference) = entry.rich_text_payload() {
            references[count] = Some((reference, SidecarHint::RichText));
            count += 1;
        }
        if let Some(reference) = entry.comment_storage() {
            references[count] = Some((reference, SidecarHint::Comment));
            count += 1;
        }
        for (reference, hint) in references.into_iter().take(count).flatten() {
            if reference.deprecated_is_external() == Some(true) {
                return Err(storage::DecodeError::invalid_visitor_result());
            }
            let Some(identifier) = NonZeroU64::new(reference.identifier()) else {
                continue;
            };
            self.sidecars
                .try_reserve(1)
                .map_err(|_| storage::DecodeError::allocation(1))?;
            self.sidecars.push((identifier, hint));
        }
        Ok(())
    }
}

fn walk_data_lists(
    source: &Package,
    index: &DeletionObjectIndex,
    graph: &mut TableGraph,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut pending = Vec::new();
    pending
        .try_reserve(graph.roles.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: graph.roles.len(),
        })?;
    for (identifier, role) in &graph.roles {
        if *role == OwnershipRole::Data {
            pending.push(*identifier);
        }
    }
    let mut visited = HashSet::new();
    visited
        .try_reserve(pending.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: pending.len(),
        })?;
    while let Some(identifier) = pending.pop() {
        if !insert_set_with_reserve(&mut visited, identifier)? {
            continue;
        }
        let object = object_at_index(source, index, identifier)?;
        let mut root_seen = false;
        for message in &object.messages {
            if !matches!(
                message.type_,
                table_data_list::TABLE_DATA_LIST_MESSAGE_KIND
                    | table_data_list::NATIVE_TABLE_DATA_LIST_MESSAGE_KIND
            ) {
                continue;
            }
            if root_seen {
                return Err(BodyTableDeletionError::InvalidSource);
            }
            root_seen = true;
            let mut visitor = ListEdgeVisitor::default();
            let options = storage_options(budget, message.data.as_slice())?;
            let (snapshot, report) = storage::decode_table_data_list_with_visitor(
                message.data.as_slice(),
                options,
                &mut visitor,
            )
            .map_err(map_storage_error)?;
            charge_storage_report(budget, report)?;
            if ListType::try_from(snapshot.list_type()).is_err() {
                // The host graph intentionally ignores unknown list kinds:
                // they are metadata extensions, not proven table-owned
                // segment routes.
                continue;
            }
            for segment in visitor.segments {
                if segment == identifier {
                    return Err(BodyTableDeletionError::InvalidSource);
                }
                insert_role(graph, index, segment, OwnershipRole::Data, budget)?;
                pending
                    .try_reserve(1)
                    .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
                pending.push(segment);
                let segment_visitor =
                    validate_list_segment(source, index, segment, snapshot.list_type(), budget)?;
                for (sidecar, hint) in segment_visitor.sidecars {
                    insert_role(graph, index, sidecar, OwnershipRole::Sidecar, budget)?;
                    merge_sidecar_hint(graph, sidecar, hint)?;
                }
            }
            for (sidecar, hint) in visitor.sidecars {
                insert_role(graph, index, sidecar, OwnershipRole::Sidecar, budget)?;
                merge_sidecar_hint(graph, sidecar, hint)?;
            }
        }
    }
    Ok(())
}

fn validate_list_segment(
    source: &Package,
    index: &DeletionObjectIndex,
    identifier: NonZeroU64,
    expected_type: i32,
    budget: &mut table_lock::WireBudget,
) -> Result<ListEdgeVisitor, BodyTableDeletionError> {
    let object = object_at_index(source, index, identifier)?;
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != table_data_list::TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND {
            continue;
        }
        if selected.is_some() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        selected = Some(message);
    }
    let message = selected.ok_or(BodyTableDeletionError::InvalidSource)?;
    let mut visitor = ListEdgeVisitor::default();
    let options = storage_options(budget, message.data.as_slice())?;
    let (snapshot, report) = storage::decode_table_data_list_segment_with_visitor(
        message.data.as_slice(),
        options,
        &mut visitor,
    )
    .map_err(map_storage_error)?;
    charge_storage_report(budget, report)?;
    if snapshot.list_type() != expected_type || !visitor.segments.is_empty() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(visitor)
}

fn walk_comment_sidecars(
    source: &Package,
    index: &DeletionObjectIndex,
    graph: &mut TableGraph,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut rich_text_ids = Vec::new();
    rich_text_ids
        .try_reserve(graph.sidecar_hints.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: graph.sidecar_hints.len(),
        })?;
    for (identifier, hint) in &graph.sidecar_hints {
        if *hint == SidecarHint::RichText {
            rich_text_ids.push(*identifier);
        }
    }
    for identifier in rich_text_ids {
        let object = object_at_index(source, index, identifier)?;
        let message = unique_message(object, RICH_TEXT_PAYLOAD_MESSAGE_KIND)?;
        let options = storage_options(budget, message.data.as_slice())?;
        let (payload, report) =
            storage::decode_rich_text_payload_with_report(message.data.as_slice(), options)
                .map_err(map_storage_error)?;
        charge_storage_report(budget, report)?;
        let storage_identifier = local_storage_identifier(payload.storage())?
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        insert_role(
            graph,
            index,
            storage_identifier,
            OwnershipRole::Sidecar,
            budget,
        )?;
        merge_sidecar_hint(graph, storage_identifier, SidecarHint::Generic)?;
    }

    let mut pending = Vec::new();
    for (identifier, hint) in &graph.sidecar_hints {
        match hint {
            SidecarHint::Comment => {
                pending
                    .try_reserve(1)
                    .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
                pending.push(*identifier);
            },
            SidecarHint::RichText => {
                // Rich-text payload envelopes are not comment archives. A
                // strict message-kind/cardinality check keeps a comment
                // object from being misclassified as a rich payload while
                // leaving its nested text-storage projection to the focused
                // rich-text codec owner.
                let object = object_at_index(source, index, *identifier)?;
                unique_message(object, RICH_TEXT_PAYLOAD_MESSAGE_KIND)?;
            },
            SidecarHint::Generic => {},
        }
    }
    let mut processed = HashSet::new();
    let mut queued = HashSet::new();
    processed
        .try_reserve(pending.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: pending.len(),
        })?;
    queued
        .try_reserve(pending.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: pending.len(),
        })?;
    for identifier in &pending {
        if !insert_set_with_reserve(&mut queued, *identifier)? {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    while let Some(identifier) = pending.pop() {
        queued.remove(&identifier);
        if !insert_set_with_reserve(&mut processed, identifier)? {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        let object = object_at_index(source, index, identifier)?;
        if object.messages.len() != 1 {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        let message = unique_message(object, COMMENT_STORAGE_MESSAGE_KIND)?;
        let mut visitor = ReplyVisitor::default();
        let options = comment_options(budget, message.data.as_slice())?;
        let (snapshot, report) =
            comment_storage_codec::decode_comment_storage_archive_with_visitor(
                message.data.as_slice(),
                options,
                &mut visitor,
            )
            .map_err(map_comment_error)?;
        charge_comment_report(budget, report)?;
        if visitor.allocation_failed {
            return Err(BodyTableDeletionError::Allocation {
                amount: report.replies().max(1),
            });
        }

        if let Some(author) = snapshot.author() {
            let author = local_comment_identifier(author)?;
            insert_role(graph, index, author, OwnershipRole::Sidecar, budget)?;
            merge_sidecar_hint(graph, author, SidecarHint::Generic)?;
        }
        for reply in visitor.replies {
            let reply = local_comment_identifier(reply)?;
            if reply == identifier || processed.contains(&reply) || queued.contains(&reply) {
                return Err(BodyTableDeletionError::InvalidSource);
            }
            insert_role(graph, index, reply, OwnershipRole::Sidecar, budget)?;
            merge_sidecar_hint(graph, reply, SidecarHint::Comment)?;
            if !insert_set_with_reserve(&mut queued, reply)? {
                return Err(BodyTableDeletionError::InvalidSource);
            }
            pending
                .try_reserve(1)
                .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
            pending.push(reply);
        }
    }
    Ok(())
}

/// Follow archive-header edges emitted by sidecar payloads after the storage
/// lists have identified their roots. This closes rich-text storage and
/// comment-author/style subgraphs without treating a body/table root as a
/// private sidecar merely because it is referenced from one payload.
fn expand_sidecar_references(
    source: &Package,
    index: &DeletionObjectIndex,
    graph: &mut TableGraph,
    protected_roots: &HashSet<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut pending = Vec::new();
    pending
        .try_reserve(graph.sidecar_hints.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: graph.sidecar_hints.len(),
        })?;
    for identifier in graph.sidecar_hints.keys().copied() {
        pending.push(identifier);
    }
    let mut visited = HashSet::new();
    visited
        .try_reserve(pending.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: pending.len(),
        })?;
    while let Some(identifier) = pending.pop() {
        if !insert_set_with_reserve(&mut visited, identifier)? {
            continue;
        }
        let object = object_at_index(source, index, identifier)?;
        for info in &object.archive_info.message_infos {
            charge_metadata_items(info, budget)?;
            let reference_count = info
                .object_references
                .len()
                .checked_add(
                    info.field_infos
                        .iter()
                        .try_fold(0usize, |count, field| {
                            count.checked_add(field.object_references.len())
                        })
                        .ok_or(BodyTableDeletionError::InvalidSource)?,
                )
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            budget
                .charge_payload_references(reference_count)
                .and_then(|_| budget.charge_payload_work(reference_count))
                .map_err(super::map_lock_error)?;
            for reference in info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter()),
            ) {
                let reference =
                    NonZeroU64::new(*reference).ok_or(BodyTableDeletionError::InvalidSource)?;
                if protected_roots.contains(&reference) || graph.roles.contains_key(&reference) {
                    continue;
                }
                insert_role(graph, index, reference, OwnershipRole::Sidecar, budget)?;
                merge_sidecar_hint(graph, reference, SidecarHint::Generic)?;
                pending
                    .try_reserve(1)
                    .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
                pending.push(reference);
            }
        }
    }
    Ok(())
}

#[derive(Default)]
struct ReplyVisitor {
    replies: Vec<comment_storage_codec::ReferenceSnapshot>,
    allocation_failed: bool,
}

impl comment_storage_codec::CommentStorageVisitor for ReplyVisitor {
    fn visit_reply(
        &mut self,
        record: comment_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        if self.replies.try_reserve(1).is_err() {
            self.allocation_failed = true;
            return Ok(());
        }
        self.replies.push(record.reference());
        Ok(())
    }
}

fn merge_sidecar_hint(
    graph: &mut TableGraph,
    identifier: NonZeroU64,
    hint: SidecarHint,
) -> Result<(), BodyTableDeletionError> {
    match graph.sidecar_hints.get(&identifier).copied() {
        None => {
            graph
                .sidecar_hints
                .try_reserve(1)
                .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
            graph.sidecar_hints.insert(identifier, hint);
        },
        Some(existing) if existing == hint || existing == SidecarHint::Generic => {
            if existing == SidecarHint::Generic && hint != SidecarHint::Generic {
                graph.sidecar_hints.insert(identifier, hint);
            }
        },
        Some(SidecarHint::Comment) if hint == SidecarHint::Generic => {},
        Some(SidecarHint::RichText) if hint == SidecarHint::Generic => {},
        Some(_) => return Err(BodyTableDeletionError::InvalidSource),
    }
    Ok(())
}

/// Expand the source-authorized formula context closure using only archive
/// header facts. This is the focused Pages counterpart of the legacy
/// `expand_formula_contexts` helper: body and sibling table roots are kept out
/// of the closure, while ordinary context objects contribute their declared
/// object edges until a formula-owner object is reached.
fn expand_formula_contexts(
    source: &Package,
    index: &DeletionObjectIndex,
    contexts: &mut Vec<NonZeroU64>,
    excluded: &HashSet<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    contexts.retain(|identifier| !excluded.contains(identifier));
    let mut seen = HashSet::new();
    seen.try_reserve(contexts.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: contexts.len(),
        })?;
    for identifier in contexts.iter().copied() {
        if !insert_set_with_reserve(&mut seen, identifier)? {
            continue;
        }
    }

    let mut cursor = 0usize;
    while cursor < contexts.len() {
        let identifier = contexts[cursor];
        cursor = cursor
            .checked_add(1)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        // The legacy helper deliberately leaves an unresolved context for
        // later formula validation. Preserve that behavior for references
        // outside the local archive index; selected owned storage itself has
        // already been required by `insert_role`.
        let Some(_) = index.get(identifier) else {
            continue;
        };
        let object = object_at_index(source, index, identifier)?;
        if object
            .messages
            .iter()
            .any(|message| matches!(message.type_, 4_000 | 4_008 | 4_009))
        {
            continue;
        }
        for info in &object.archive_info.message_infos {
            charge_metadata_items(info, budget)?;
            let edge_count = info
                .object_references
                .len()
                .checked_add(
                    info.field_infos
                        .iter()
                        .try_fold(0usize, |count, field| {
                            count.checked_add(field.object_references.len())
                        })
                        .ok_or(BodyTableDeletionError::InvalidSource)?,
                )
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            budget
                .charge_payload_references(edge_count)
                .and_then(|_| budget.charge_payload_work(edge_count))
                .map_err(super::map_lock_error)?;
            for reference in info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter()),
            ) {
                let reference =
                    NonZeroU64::new(*reference).ok_or(BodyTableDeletionError::InvalidSource)?;
                if excluded.contains(&reference) || !insert_set_with_reserve(&mut seen, reference)?
                {
                    continue;
                }
                contexts
                    .try_reserve(1)
                    .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
                contexts.push(reference);
            }
        }
    }
    Ok(())
}

fn build_planned_reference_removals(
    graph: &RemovalPlan,
    formula: &RemovalPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<HashSet<PlannedObjectReferenceRemoval>, BodyTableDeletionError> {
    let edits = graph
        .message_edits
        .len()
        .checked_add(formula.message_edits.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let mut references = 0usize;
    for edit in graph.message_edits.iter().chain(&formula.message_edits) {
        references = references
            .checked_add(edit.remove_object_references.len())
            .ok_or(BodyTableDeletionError::InvalidSource)?;
    }
    let work = edits
        .checked_add(references)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_items(edits)
        .and_then(|_| budget.charge_payload_references(references))
        .and_then(|_| budget.charge_payload_work(work))
        .map_err(super::map_lock_error)?;

    let mut planned = HashSet::new();
    planned
        .try_reserve(references)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: references })?;
    for edit in graph.message_edits.iter().chain(&formula.message_edits) {
        for target in &edit.remove_object_references {
            insert_set_with_reserve(
                &mut planned,
                PlannedObjectReferenceRemoval {
                    component_index: edit.component_index,
                    object_index: edit.object_index,
                    message_index: edit.message_index,
                    target: *target,
                },
            )?;
        }
    }
    Ok(planned)
}

/// Complete the formula phase with the legacy orphan-context closure.
///
/// Formula owner and dependency removal may make ordinary context objects
/// unreachable even though the table storage graph itself was private. Walk
/// those contexts against the immutable source, treating every planned
/// object removal and every planned object-reference edit as already applied.
/// Only an object with no surviving inbound object reference is appended to
/// the formula removal worklist; its own declared context edges then become
/// the next bounded candidates. Metadata preparation runs after this closure
/// and therefore sees the complete object set.
pub(super) fn complete_removals(
    source: &Package,
    graph: &GraphPlan,
    formula: &mut super::FormulaPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut removed = HashSet::new();
    let removal_capacity = graph
        .removals
        .object_removals
        .len()
        .checked_add(formula.removals.object_removals.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    removed
        .try_reserve(removal_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removal_capacity,
        })?;
    for removal in graph
        .removals
        .object_removals
        .iter()
        .chain(&formula.removals.object_removals)
    {
        if !insert_set_with_reserve(&mut removed, removal.identifier)? {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }

    let planned_references =
        build_planned_reference_removals(&graph.removals, &formula.removals, budget)?;

    let mut protected = HashSet::new();
    protected
        .try_reserve(1usize.saturating_add(graph.source_tables.len().saturating_mul(3)))
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: 1usize.saturating_add(graph.source_tables.len().saturating_mul(3)),
        })?;
    protected.insert(graph.request.target.body_identifier);
    for target in source_table_targets(source, graph, budget)? {
        protected.insert(target.attachment_identifier);
        protected.insert(target.drawable_identifier);
        protected.insert(target.model_identifier);
    }

    let mut pending = try_clone_identifiers(&graph.removals.formula_contexts)?;
    pending
        .try_reserve(formula.removals.formula_contexts.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: formula.removals.formula_contexts.len(),
        })?;
    pending.extend(formula.removals.formula_contexts.iter().copied());
    let mut queued = HashSet::new();
    queued
        .try_reserve(pending.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: pending.len(),
        })?;
    for identifier in pending.iter().copied() {
        if NonZeroU64::new(identifier.get()).is_none() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        insert_set_with_reserve(&mut queued, identifier)?;
    }
    let mut examined = HashSet::new();
    examined
        .try_reserve(pending.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: pending.len(),
        })?;

    while let Some(identifier) = pending.pop() {
        queued.remove(&identifier);
        if !insert_set_with_reserve(&mut examined, identifier)?
            || removed.contains(&identifier)
            || protected.contains(&identifier)
        {
            continue;
        }
        if object_has_surviving_inbound_reference(
            source,
            &graph.index,
            identifier,
            &removed,
            &planned_references,
            budget,
        )? {
            continue;
        }
        let object = object_at_index(source, &graph.index, identifier)?;
        let mut child_references = Vec::new();
        collect_object_references(object, &mut child_references, budget)?;

        let location = graph
            .index
            .get(identifier)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        formula
            .removals
            .object_removals
            .try_reserve(1)
            .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
        budget
            .charge_payload_items(1)
            .and_then(|_| budget.charge_payload_work(1))
            .map_err(super::map_lock_error)?;
        formula.removals.object_removals.push(ObjectRemoval {
            component_index: location.component_index,
            object_index: location.object_index,
            identifier,
        });
        if !insert_set_with_reserve(&mut removed, identifier)? {
            return Err(BodyTableDeletionError::InvalidSource);
        }

        for child in child_references {
            if protected.contains(&child) || removed.contains(&child) {
                continue;
            }
            if !insert_set_with_reserve(&mut queued, child)? {
                continue;
            }
            pending
                .try_reserve(1)
                .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
            pending.push(child);
        }
    }
    Ok(())
}

/// Reconstruct the rooted target set from the graph's source snapshots. The
/// target resolver remains the authority for exact routes; this helper only
/// needs the three protected root identifiers while closing formula orphans.
fn source_table_targets(
    source: &Package,
    graph: &GraphPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<table_lock::BodyTableTarget>, BodyTableDeletionError> {
    let targets = table_lock::body_table_catalog_with_budget(source, budget)
        .map_err(super::map_lock_error)?;
    if targets.len() != graph.source_tables.len() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    if targets
        .iter()
        .zip(&graph.source_tables)
        .any(|(target, snapshot)| {
            target.table_position != snapshot.index()
                || target.table_name.as_ref() != snapshot.name()
                || target.table_rows != snapshot.rows()
                || target.table_columns != snapshot.columns()
        })
    {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(targets)
}

fn object_has_surviving_inbound_reference(
    source: &Package,
    index: &DeletionObjectIndex,
    target: NonZeroU64,
    removed: &HashSet<NonZeroU64>,
    planned_references: &HashSet<PlannedObjectReferenceRemoval>,
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableDeletionError> {
    for location in index.iter() {
        if removed.contains(&location.identifier) {
            continue;
        }
        let object = object_at_index(source, index, location.identifier)?;
        for (message_index, info) in object.archive_info.message_infos.iter().enumerate() {
            charge_metadata_items(info, budget)?;
            let reference_count = info
                .object_references
                .len()
                .checked_add(
                    info.field_infos
                        .iter()
                        .try_fold(0usize, |count, field| {
                            count.checked_add(field.object_references.len())
                        })
                        .ok_or(BodyTableDeletionError::InvalidSource)?,
                )
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            budget
                .charge_payload_references(reference_count)
                .and_then(|_| budget.charge_payload_work(reference_count))
                .map_err(super::map_lock_error)?;
            for reference in info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter()),
            ) {
                let reference =
                    NonZeroU64::new(*reference).ok_or(BodyTableDeletionError::InvalidSource)?;
                let planned = PlannedObjectReferenceRemoval {
                    component_index: location.component_index,
                    object_index: location.object_index,
                    message_index,
                    target,
                };
                if reference == target && !planned_references.contains(&planned) {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

fn collect_object_references(
    object: &ArchiveObject,
    references: &mut Vec<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    for info in &object.archive_info.message_infos {
        charge_metadata_items(info, budget)?;
        let count = info
            .object_references
            .len()
            .checked_add(
                info.field_infos
                    .iter()
                    .try_fold(0usize, |count, field| {
                        count.checked_add(field.object_references.len())
                    })
                    .ok_or(BodyTableDeletionError::InvalidSource)?,
            )
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        budget
            .charge_payload_references(count)
            .and_then(|_| budget.charge_payload_work(count))
            .map_err(super::map_lock_error)?;
        references
            .try_reserve(count)
            .map_err(|_| BodyTableDeletionError::Allocation { amount: count })?;
        for reference in info.object_references.iter().chain(
            info.field_infos
                .iter()
                .flat_map(|field| field.object_references.iter()),
        ) {
            references
                .push(NonZeroU64::new(*reference).ok_or(BodyTableDeletionError::InvalidSource)?);
        }
    }
    Ok(())
}

fn prove_inbound_references(
    source: &Package,
    index: &DeletionObjectIndex,
    selected: &TableGraph,
    removed: &HashSet<NonZeroU64>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let selected_roots = [
        selected.target.attachment_identifier,
        selected.target.drawable_identifier,
        selected.target.model_identifier,
    ];
    let selected_body_identifier = selected.target.body_identifier;
    let selected_body_message_index = selected.target.body_message_index;
    let selected_attachment_identifier = selected.target.attachment_identifier;
    let mut body_aggregate_edges = 0usize;
    let mut body_field_edges = 0usize;
    let mut body_field_declarations = 0usize;
    for location in index.iter() {
        let object = object_at_index(source, index, location.identifier)?;
        let object_removed = removed.contains(&location.identifier);
        let formula_source = object
            .messages
            .iter()
            .any(|message| matches!(message.type_, 4_000 | 4_008 | 4_009 | 4_010));
        for (message_index, info) in object.archive_info.message_infos.iter().enumerate() {
            charge_metadata_items(info, budget)?;
            let metadata_count = info
                .object_references
                .len()
                .checked_add(info.data_references.len())
                .and_then(|count| {
                    info.field_infos.iter().try_fold(count, |count, field| {
                        count
                            .checked_add(field.object_references.len())
                            .and_then(|value| value.checked_add(field.data_references.len()))
                    })
                })
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            budget
                .charge_payload_references(metadata_count)
                .and_then(|_| budget.charge_payload_work(metadata_count))
                .map_err(super::map_lock_error)?;
            let body_message = location.identifier == selected_body_identifier
                && message_index == selected_body_message_index;
            for identifier in &info.object_references {
                if body_message && *identifier == selected_attachment_identifier.get() {
                    body_aggregate_edges = body_aggregate_edges
                        .checked_add(1)
                        .ok_or(BodyTableDeletionError::InvalidSource)?;
                }
                check_inbound(
                    location.identifier,
                    *identifier,
                    object_removed,
                    formula_source,
                    message_index,
                    body_message,
                    selected_body_identifier,
                    selected_body_message_index,
                    selected_attachment_identifier,
                    &selected_roots,
                    removed,
                )?;
            }
            for identifier in &info.data_references {
                check_data_inbound(location.identifier, *identifier, object_removed, removed)?;
            }
            for field in &info.field_infos {
                let body_table_field = body_message && field.path.as_slice() == [9];
                if body_table_field {
                    body_field_declarations = body_field_declarations
                        .checked_add(1)
                        .ok_or(BodyTableDeletionError::InvalidSource)?;
                }
                for identifier in &field.object_references {
                    if body_table_field && *identifier == selected_attachment_identifier.get() {
                        body_field_edges = body_field_edges
                            .checked_add(1)
                            .ok_or(BodyTableDeletionError::InvalidSource)?;
                    }
                    check_inbound(
                        location.identifier,
                        *identifier,
                        object_removed,
                        formula_source,
                        message_index,
                        body_table_field,
                        selected_body_identifier,
                        selected_body_message_index,
                        selected_attachment_identifier,
                        &selected_roots,
                        removed,
                    )?;
                }
                for identifier in &field.data_references {
                    check_data_inbound(location.identifier, *identifier, object_removed, removed)?;
                }
            }
        }
    }
    if body_aggregate_edges != 1 || (body_field_declarations != 0 && body_field_edges != 1) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(())
}

fn check_inbound(
    source_identifier: NonZeroU64,
    target_identifier: u64,
    source_removed: bool,
    formula_source: bool,
    message_index: usize,
    body_edge_allowed: bool,
    selected_body_identifier: NonZeroU64,
    selected_body_message_index: usize,
    selected_attachment_identifier: NonZeroU64,
    selected_roots: &[NonZeroU64; 3],
    removed: &HashSet<NonZeroU64>,
) -> Result<(), BodyTableDeletionError> {
    let target = NonZeroU64::new(target_identifier).ok_or(BodyTableDeletionError::InvalidSource)?;
    if !removed.contains(&target) || source_identifier == target {
        return Ok(());
    }
    if source_removed {
        return Ok(());
    }
    // The body message is rewritten by the archive phase to remove exactly
    // the selected attachment edge. Formula owners are rewritten by the
    // formula phase; their selected table-root edges are therefore expected
    // to disappear before final metadata/archive publication.
    if body_edge_allowed
        && source_identifier == selected_body_identifier
        && message_index == selected_body_message_index
        && target == selected_attachment_identifier
    {
        return Ok(());
    }
    if formula_source && selected_roots.contains(&target) {
        return Ok(());
    }
    Err(BodyTableDeletionError::UnsupportedDependency)
}

fn check_data_inbound(
    _source_identifier: NonZeroU64,
    target_identifier: u64,
    _source_removed: bool,
    _removed: &HashSet<NonZeroU64>,
) -> Result<(), BodyTableDeletionError> {
    let Some(target) = NonZeroU64::new(target_identifier) else {
        return Err(BodyTableDeletionError::InvalidSource);
    };
    // Data-reference identifiers live in the package data namespace, while
    // `removed` contains archive-object identifiers.  Metadata preparation
    // owns the data-owner registry rewrite; graph must only reject malformed
    // zero references here and never infer a cross-namespace dependency from
    // a numeric collision.
    let _ = target;
    Ok(())
}

fn object_at_index<'a>(
    source: &'a Package,
    index: &DeletionObjectIndex,
    identifier: NonZeroU64,
) -> Result<&'a ArchiveObject, BodyTableDeletionError> {
    let location = index
        .get(identifier)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    object_at_route(
        source,
        location.component_index,
        location.object_index,
        identifier,
    )
}

fn object_at_route(
    source: &Package,
    component_index: usize,
    object_index: usize,
    identifier: NonZeroU64,
) -> Result<&ArchiveObject, BodyTableDeletionError> {
    let component = source
        .state
        .source
        .components()
        .get_index(component_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(object_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    if object.archive_info.identifier != Some(identifier.get()) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(object)
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&litchi_iwa_core::RawMessage, BodyTableDeletionError> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.is_some() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        selected = Some(message);
    }
    selected.ok_or(BodyTableDeletionError::InvalidSource)
}

fn storage_options(
    budget: &table_lock::WireBudget,
    source: &[u8],
) -> Result<storage::DecodeOptions, BodyTableDeletionError> {
    let limits = budget.wire_limits();
    let input = source.len().max(1);
    let fields = budget.remaining_wire_fields();
    let work = budget.remaining_wire_work();
    if input > limits.max_input_bytes() || fields == 0 || work == 0 {
        return Err(BodyTableDeletionError::LimitExceeded {
            kind: if input > limits.max_input_bytes() {
                super::BodyTableDeletionLimitKind::WireBytes
            } else if fields == 0 {
                super::BodyTableDeletionLimitKind::WireFields
            } else {
                super::BodyTableDeletionLimitKind::WireWork
            },
            observed: if input > limits.max_input_bytes() {
                usize_to_u64(input)
            } else {
                1
            },
            maximum: if input > limits.max_input_bytes() {
                usize_to_u64(limits.max_input_bytes())
            } else if fields == 0 {
                0
            } else {
                0
            },
        });
    }
    Ok(storage::DecodeOptions::new(
        input,
        fields,
        work,
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        budget.remaining_payload_references(),
        input,
    ))
}

fn charge_storage_report(
    budget: &mut table_lock::WireBudget,
    report: storage::DecodeReport,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(report.source_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            )
        })
        .and_then(|_| budget.charge_payload_work(report.text_bytes()))
        .map_err(super::map_lock_error)
}

fn comment_options(
    budget: &table_lock::WireBudget,
    source: &[u8],
) -> Result<comment_storage_codec::DecodeOptions, BodyTableDeletionError> {
    let limits = budget.wire_limits();
    let input = source.len().max(1);
    let fields = budget.remaining_wire_fields();
    let work = budget.remaining_wire_work();
    if input > limits.max_input_bytes() || fields == 0 || work == 0 {
        return Err(BodyTableDeletionError::LimitExceeded {
            kind: if input > limits.max_input_bytes() {
                super::BodyTableDeletionLimitKind::WireBytes
            } else if fields == 0 {
                super::BodyTableDeletionLimitKind::WireFields
            } else {
                super::BodyTableDeletionLimitKind::WireWork
            },
            observed: usize_to_u64(input.max(1)),
            maximum: usize_to_u64(limits.max_input_bytes()),
        });
    }
    Ok(comment_storage_codec::DecodeOptions::new(
        input,
        fields,
        work,
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        budget.remaining_payload_references(),
        input,
    ))
}

fn charge_comment_report(
    budget: &mut table_lock::WireBudget,
    report: comment_storage_codec::DecodeReport,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(report.source_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            )
        })
        .and_then(|_| budget.charge_payload_work(report.reference_bytes()))
        .and_then(|_| budget.charge_payload_work(report.text_bytes()))
        .and_then(|_| budget.charge_payload_items(report.replies()))
        .map_err(super::map_lock_error)
}

fn local_comment_identifier(
    reference: comment_storage_codec::ReferenceSnapshot,
) -> Result<NonZeroU64, BodyTableDeletionError> {
    if reference.deprecated_is_external() == Some(true) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    NonZeroU64::new(reference.identifier()).ok_or(BodyTableDeletionError::InvalidSource)
}

fn map_storage_error(error: storage::DecodeError) -> BodyTableDeletionError {
    let Some(limit) = error.resource_limit() else {
        return BodyTableDeletionError::InvalidSource;
    };
    match limit {
        storage::DecodeLimit::Bytes { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        storage::DecodeLimit::References { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::PayloadReferences,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        storage::DecodeLimit::Text { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        storage::DecodeLimit::Fields { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireFields,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        storage::DecodeLimit::Work { observed, maximum }
        | storage::DecodeLimit::Retained { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        storage::DecodeLimit::Nesting { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        storage::DecodeLimit::Allocation { requested } => {
            BodyTableDeletionError::Allocation { amount: requested }
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn map_comment_error(error: comment_storage_codec::DecodeError) -> BodyTableDeletionError {
    let Some(limit) = error.resource_limit() else {
        return BodyTableDeletionError::InvalidSource;
    };
    match limit {
        comment_storage_codec::DecodeLimit::Bytes { observed, maximum }
        | comment_storage_codec::DecodeLimit::OutputBytes { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: if matches!(
                    limit,
                    comment_storage_codec::DecodeLimit::OutputBytes { .. }
                ) {
                    super::BodyTableDeletionLimitKind::WireOutputBytes
                } else {
                    super::BodyTableDeletionLimitKind::WireBytes
                },
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        comment_storage_codec::DecodeLimit::References { observed, maximum }
        | comment_storage_codec::DecodeLimit::Replies { observed, maximum }
        | comment_storage_codec::DecodeLimit::ReferenceBytes { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::PayloadReferences,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        comment_storage_codec::DecodeLimit::Text { observed, maximum }
        | comment_storage_codec::DecodeLimit::Work { observed, maximum }
        | comment_storage_codec::DecodeLimit::Scratch { observed, maximum }
        | comment_storage_codec::DecodeLimit::Retained { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        comment_storage_codec::DecodeLimit::Fields { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireFields,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        comment_storage_codec::DecodeLimit::Nesting { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        comment_storage_codec::DecodeLimit::Allocations { observed, maximum } => {
            BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
