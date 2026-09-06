//! Private source-bound planning for removing a selected comment closure.
//!
//! Comment storage is normally private to the selected drawable, but the
//! archive graph can share a storage root (and its replies) with another drawable.
//! This planner therefore computes a fixed point: references from objects that
//! survive the media removal retain the shared storage, and retained storage
//! retains every selected reply reachable from it.  Annotation authors are
//! never removal candidates.  An unused author releases only its current strong
//! component dependency; the native writer retains the author object and its
//! author-storage member.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The private planner keeps source validation, census, and closure facts together."
)]

use std::mem::size_of;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, Limits as ArchiveLimits, MessageInfo};

use super::budget::LifecycleBudget;
use super::comment_graph::CommentGraphPlan;
use super::{Package, SlideMediaLifecycleError};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;

#[cfg(test)]
mod tests;

/// Source-bound removal facts for one selected media transaction.
///
/// `removed_object_ids` is the effective physical-removal set.  Selected
/// comment storage IDs that are reachable from a surviving object are omitted
/// from it, together with any selected replies reachable from those retained
/// roots.  The complete retained closure is in
/// `retained_comment_storage_ids`.
///
/// `unused_external_author_ids` authorizes removal of exact current strong
/// component dependencies only. Native Keynote retains the author objects and
/// their storage even
/// when no surviving object references them after the selected comment closure
/// is removed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct CommentRemovalPlan {
    pub(super) removed_object_ids: Vec<u64>,
    pub(super) retained_comment_storage_ids: Vec<u64>,
    pub(super) unused_external_author_ids: Vec<u64>,
}

/// Plan the effective object-removal set for a selected comment graph.
///
/// The exact immutable [`Package`] is scanned across every parsed component,
/// not only the selected slide component, so an incoming reference from an
/// unrelated drawable or component preserves the shared comment storage.
/// Every object header goes through the strict core ArchiveInfo census before
/// its object and field reference lists are used.  The same lifecycle budget
/// accounts for that census and all planner-owned vectors.
pub(super) fn plan_comment_removal(
    package: &Package,
    component_name: &str,
    selected_object_ids: &[u64],
    comment_graph: &CommentGraphPlan,
    archive_limits: ArchiveLimits,
    budget: &mut LifecycleBudget,
) -> Result<CommentRemovalPlan, SlideMediaLifecycleError> {
    if component_name.is_empty() || comment_graph.component_name.as_ref() != component_name {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    validate_sorted_unique(selected_object_ids)?;
    validate_sorted_unique(&comment_graph.storage_ids)?;
    validate_sorted_unique(&comment_graph.author_ids)?;
    if comment_graph.storage_ids.is_empty() || comment_graph.root_storage_identifier == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    for &identifier in selected_object_ids {
        let Some((actual_component, _)) = package.object_with_component(identifier) else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if actual_component != component_name {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    for &identifier in &comment_graph.storage_ids {
        if selected_object_ids.binary_search(&identifier).is_err() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let Some((actual_component, _)) = package.object_with_component(identifier) else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if actual_component != component_name {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    for &identifier in &comment_graph.author_ids {
        if identifier == 0 || selected_object_ids.binary_search(&identifier).is_ok() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        if package.object_with_component(identifier).is_none() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    validate_author_dependencies(package, comment_graph, budget)?;
    let wire_limits = package
        .semantic_wire_limits()
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;

    let mut retained_comment_storage_ids = Vec::new();
    reserve_vec(
        &mut retained_comment_storage_ids,
        comment_graph.storage_ids.len(),
        budget,
    )?;

    let mut comment_edges = Vec::new();
    let mut selected_comment_author_edges = Vec::new();
    let mut used_author_ids = Vec::new();
    reserve_vec(&mut used_author_ids, comment_graph.author_ids.len(), budget)?;

    // Scan every source object once.  `strict_reference_census` proves that
    // the retained ArchiveInfo projection contains no unknown metadata before
    // this planner reads its object-reference lists.
    for component in package.state.source.components().iter() {
        budget.charge_entries(component.archive().objects.len())?;
        let source_component = component.name() == component_name;
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            if object.archive_info.message_infos.len() != object.messages.len() {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            validate_known_payload_relationships(
                package,
                component.name(),
                object,
                &comment_graph.storage_ids,
                wire_limits,
                budget,
            )?;
            let header_limits =
                super::graph::reserve_core_header_inspection_work(object, archive_limits, budget)?;
            super::graph::strict_reference_census(object, header_limits, budget)?;

            let selected = selected_object_ids.binary_search(&identifier).is_ok();
            let selected_comment = comment_graph.storage_ids.binary_search(&identifier).is_ok();
            let mut field_count = 0usize;
            let mut metadata_items = 0usize;
            for info in &object.archive_info.message_infos {
                field_count = field_count
                    .checked_add(info.field_infos.len())
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?;
                metadata_items = metadata_items
                    .checked_add(info.object_references.len())
                    .and_then(|count| count.checked_add(info.data_references.len()))
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?;
                for field in &info.field_infos {
                    metadata_items = metadata_items
                        .checked_add(field.object_references.len())
                        .and_then(|count| count.checked_add(field.data_references.len()))
                        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
                }
            }
            budget.charge_wire_fields(
                object
                    .archive_info
                    .message_infos
                    .len()
                    .checked_add(field_count)
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?,
            )?;
            budget.charge_wire_work(metadata_items.saturating_add(field_count).max(1))?;

            for info in &object.archive_info.message_infos {
                for &reference in &info.object_references {
                    visit_reference(
                        identifier,
                        reference,
                        selected,
                        selected_comment,
                        source_component,
                        &comment_graph.storage_ids,
                        &comment_graph.author_ids,
                        &mut retained_comment_storage_ids,
                        &mut comment_edges,
                        &mut selected_comment_author_edges,
                        &mut used_author_ids,
                        budget,
                    )?;
                }
                for field in &info.field_infos {
                    for &reference in &field.object_references {
                        visit_reference(
                            identifier,
                            reference,
                            selected,
                            selected_comment,
                            source_component,
                            &comment_graph.storage_ids,
                            &comment_graph.author_ids,
                            &mut retained_comment_storage_ids,
                            &mut comment_edges,
                            &mut selected_comment_author_edges,
                            &mut used_author_ids,
                            budget,
                        )?;
                    }
                }
            }
        }
    }

    close_retained_comment_storage_ids(
        &mut retained_comment_storage_ids,
        &mut comment_edges,
        budget,
    )?;

    for (storage_identifier, author_identifier) in selected_comment_author_edges {
        if retained_comment_storage_ids
            .binary_search(&storage_identifier)
            .is_ok()
        {
            insert_sorted_unique(&mut used_author_ids, author_identifier, budget)?;
        }
    }

    let mut removed_object_ids = Vec::new();
    reserve_vec(&mut removed_object_ids, selected_object_ids.len(), budget)?;
    for &identifier in selected_object_ids {
        if comment_graph.storage_ids.binary_search(&identifier).is_ok()
            && retained_comment_storage_ids
                .binary_search(&identifier)
                .is_ok()
        {
            continue;
        }
        push_vec(&mut removed_object_ids, identifier, budget)?;
    }

    let mut external_author_ids = Vec::new();
    for dependency in &comment_graph.author_dependencies {
        if dependency.component_name.as_ref() != component_name {
            insert_sorted_unique(
                &mut external_author_ids,
                dependency.author_identifier,
                budget,
            )?;
        }
    }
    let mut unused_external_author_ids = Vec::new();
    reserve_vec(
        &mut unused_external_author_ids,
        external_author_ids.len(),
        budget,
    )?;
    for &identifier in &external_author_ids {
        if used_author_ids.binary_search(&identifier).is_err() {
            push_vec(&mut unused_external_author_ids, identifier, budget)?;
        }
    }

    Ok(CommentRemovalPlan {
        removed_object_ids,
        retained_comment_storage_ids,
        unused_external_author_ids,
    })
}

fn validate_author_dependencies(
    package: &Package,
    comment_graph: &CommentGraphPlan,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_wire_work(comment_graph.author_dependencies.len().max(1))?;
    for dependency in &comment_graph.author_dependencies {
        if comment_graph
            .storage_ids
            .binary_search(&dependency.storage_identifier)
            .is_err()
            || comment_graph
                .author_ids
                .binary_search(&dependency.author_identifier)
                .is_err()
        {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        let Some((actual_component, _)) =
            package.object_with_component(dependency.author_identifier)
        else {
            return Err(SlideMediaLifecycleError::InvalidSource);
        };
        if actual_component != dependency.component_name.as_ref()
            || dependency.component_name.as_ref().is_empty()
        {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    Ok(())
}

/// Validate the payload/header relationship for known ownership edges before
/// the generic ArchiveInfo census is used for retention.  Opaque root fields
/// remain untouched and are deliberately ignored here; the known Movie and
/// CommentStorage projections must still account for every corresponding
/// header edge so a stale payload cannot silently remove a live root.
fn validate_known_payload_relationships(
    package: &Package,
    component_name: &str,
    object: &ArchiveObject,
    comment_storage_ids: &[u64],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    for (message, info) in object
        .messages
        .iter()
        .zip(object.archive_info.message_infos.iter())
    {
        match message.type_ {
            MOVIE_MESSAGE_TYPE => {
                if info.type_ != MOVIE_MESSAGE_TYPE {
                    return Err(SlideMediaLifecycleError::InvalidSource);
                }
                validate_movie_payload_relationship(
                    package,
                    component_name,
                    message.data.as_slice(),
                    info,
                    comment_storage_ids,
                    limits,
                    budget,
                )?
            },
            COMMENT_STORAGE_MESSAGE_TYPE => validate_comment_storage_relationship(
                package,
                component_name,
                object
                    .archive_info
                    .identifier
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?,
                info,
                message.data.as_slice(),
                limits,
                budget,
            )?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_movie_payload_relationship(
    package: &Package,
    component_name: &str,
    payload: &[u8],
    info: &MessageInfo,
    comment_storage_ids: &[u64],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let direct_comment = super::graph::direct_drawable_comment(payload, limits, budget)?;
    if let Some(identifier) = direct_comment {
        validate_comment_storage_target(package, component_name, identifier)?;
        let count = info
            .object_references
            .iter()
            .filter(|reference| **reference == identifier)
            .count();
        if count != 1 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    for &reference in &info.object_references {
        if comment_storage_ids.binary_search(&reference).is_ok()
            && direct_comment != Some(reference)
        {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    validate_field_reference_subset(info, budget)
}

fn validate_comment_storage_relationship(
    package: &Package,
    component_name: &str,
    storage_identifier: u64,
    info: &MessageInfo,
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let facts = super::comment_graph::validate_comment_storage_payload_relationship(
        storage_identifier,
        info,
        payload,
        limits,
        budget,
    )?;
    if let Some(author_identifier) = facts.author_identifier {
        validate_annotation_author_target(package, author_identifier)?;
    }
    for reply_identifier in facts.reply_identifiers {
        validate_comment_storage_target(package, component_name, reply_identifier)?;
    }
    Ok(())
}

fn validate_comment_storage_target(
    package: &Package,
    component_name: &str,
    identifier: u64,
) -> Result<(), SlideMediaLifecycleError> {
    let Some((actual_component, object)) = package.object_with_component(identifier) else {
        return Err(SlideMediaLifecycleError::InvalidSource);
    };
    if actual_component != component_name {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    validate_single_message_type(object, identifier, COMMENT_STORAGE_MESSAGE_TYPE)
}

fn validate_annotation_author_target(
    package: &Package,
    identifier: u64,
) -> Result<(), SlideMediaLifecycleError> {
    let Some((_, object)) = package.object_with_component(identifier) else {
        return Err(SlideMediaLifecycleError::InvalidSource);
    };
    validate_single_message_type(object, identifier, 212)
}

fn validate_single_message_type(
    object: &ArchiveObject,
    identifier: u64,
    expected_type: u32,
) -> Result<(), SlideMediaLifecycleError> {
    if object.archive_info.identifier != Some(identifier)
        || object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let message = object
        .messages
        .first()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .first()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if message.type_ != expected_type
        || info.type_ != expected_type
        || usize::try_from(info.length).ok() != Some(message.data.len())
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    Ok(())
}

fn validate_field_reference_subset(
    info: &MessageInfo,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let field_reference_count = info.field_infos.iter().try_fold(0usize, |count, field| {
        count
            .checked_add(field.object_references.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)
    })?;
    let mut aggregate = Vec::new();
    reserve_vec(&mut aggregate, info.object_references.len(), budget)?;
    for &reference in &info.object_references {
        if reference == 0 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        push_vec(&mut aggregate, reference, budget)?;
    }
    let mut fields = Vec::new();
    reserve_vec(&mut fields, field_reference_count, budget)?;
    for field in &info.field_infos {
        for &reference in &field.object_references {
            if reference == 0 {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            push_vec(&mut fields, reference, budget)?;
        }
    }
    let sort_work = aggregate
        .len()
        .saturating_add(fields.len())
        .max(1)
        .checked_mul(usize::BITS as usize)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_wire_work(sort_work)?;
    aggregate.sort_unstable();
    fields.sort_unstable();

    let mut aggregate_index = 0usize;
    for reference in fields {
        while aggregate_index < aggregate.len() && aggregate[aggregate_index] < reference {
            aggregate_index += 1;
        }
        if aggregate_index == aggregate.len() || aggregate[aggregate_index] != reference {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        aggregate_index += 1;
    }
    Ok(())
}

fn visit_reference(
    object_identifier: u64,
    reference: u64,
    selected: bool,
    selected_comment: bool,
    source_component: bool,
    comment_storage_ids: &[u64],
    author_ids: &[u64],
    retained_comment_storage_ids: &mut Vec<u64>,
    comment_edges: &mut Vec<(u64, u64)>,
    selected_comment_author_edges: &mut Vec<(u64, u64)>,
    used_author_ids: &mut Vec<u64>,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if reference == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let is_comment = comment_storage_ids.binary_search(&reference).is_ok();
    let is_author = author_ids.binary_search(&reference).is_ok();
    if !selected && is_comment {
        insert_sorted_unique(retained_comment_storage_ids, reference, budget)?;
    } else if selected_comment && is_comment {
        push_pair(comment_edges, (object_identifier, reference), budget)?;
    }
    if is_author && source_component {
        if !selected {
            insert_sorted_unique(used_author_ids, reference, budget)?;
        } else if selected_comment {
            push_pair(
                selected_comment_author_edges,
                (object_identifier, reference),
                budget,
            )?;
        }
    }
    Ok(())
}

/// Sort and close the shared-storage roots independently of storage ID order.
///
/// A sorted cursor is insufficient here: a reply can have a lower ID than its
/// parent, so appending it to a sorted retained list may place it before the
/// cursor that discovered the parent.  The pending stack and independent
/// processed set visit every newly retained storage exactly once.
fn close_retained_comment_storage_ids(
    retained_comment_storage_ids: &mut Vec<u64>,
    comment_edges: &mut Vec<(u64, u64)>,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let sort_work = comment_edges
        .len()
        .max(1)
        .checked_mul(usize::BITS as usize)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_wire_work(sort_work)?;
    comment_edges.sort_unstable();
    comment_edges.dedup();

    let mut pending = Vec::new();
    reserve_vec(&mut pending, retained_comment_storage_ids.len(), budget)?;
    for &identifier in retained_comment_storage_ids.iter() {
        push_vec(&mut pending, identifier, budget)?;
    }
    let mut processed = Vec::new();
    reserve_vec(&mut processed, retained_comment_storage_ids.len(), budget)?;
    while let Some(source) = pending.pop() {
        if processed.binary_search(&source).is_ok() {
            continue;
        }
        insert_sorted_unique(&mut processed, source, budget)?;
        budget.charge_wire_work(comment_edges.len().max(1))?;
        let mut edge_index = lower_bound_source(comment_edges, source);
        while edge_index < comment_edges.len() && comment_edges[edge_index].0 == source {
            if insert_sorted_unique(
                retained_comment_storage_ids,
                comment_edges[edge_index].1,
                budget,
            )? {
                push_vec(&mut pending, comment_edges[edge_index].1, budget)?;
            }
            edge_index += 1;
        }
    }
    Ok(())
}

fn lower_bound_source(edges: &[(u64, u64)], source: u64) -> usize {
    let mut low = 0usize;
    let mut high = edges.len();
    while low < high {
        let middle = low + (high - low) / 2;
        if edges[middle].0 < source {
            low = middle + 1;
        } else {
            high = middle;
        }
    }
    low
}

fn validate_sorted_unique(values: &[u64]) -> Result<(), SlideMediaLifecycleError> {
    if values.contains(&0) || values.windows(2).any(|window| window[0] >= window[1]) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    Ok(())
}

fn reserve_vec<T>(
    output: &mut Vec<T>,
    additional: usize,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if additional == 0 {
        return Ok(());
    }
    let amount = additional
        .checked_mul(size_of::<T>())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocations(amount)?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| SlideMediaLifecycleError::Allocation { amount })
}

fn push_vec<T>(
    output: &mut Vec<T>,
    value: T,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if output.len() == output.capacity() {
        reserve_vec(output, output.capacity().max(1), budget)?;
    }
    output.push(value);
    Ok(())
}

fn push_pair(
    output: &mut Vec<(u64, u64)>,
    value: (u64, u64),
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    push_vec(output, value, budget)
}

fn insert_sorted_unique(
    output: &mut Vec<u64>,
    value: u64,
    budget: &mut LifecycleBudget,
) -> Result<bool, SlideMediaLifecycleError> {
    match output.binary_search(&value) {
        Ok(_) => Ok(false),
        Err(index) => {
            budget.charge_wire_work(output.len().max(1))?;
            if output.len() == output.capacity() {
                reserve_vec(output, output.capacity().max(1), budget)?;
            }
            output.insert(index, value);
            Ok(true)
        },
    }
}
