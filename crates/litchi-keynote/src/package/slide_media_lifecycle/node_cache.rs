//! Transaction-local maintenance of the native slide-node build cache.
//!
//! A slide's build list lives in its `KN.SlideArchive`, while the four scalar
//! cache values live in the corresponding `KN.SlideNodeArchive`.  Those
//! archives are normally separate components.  This adapter keeps that
//! component split explicit: the slide component is edited in place and a
//! source node component is cloned only when it is different from the edited
//! component.  The caller owns publication of the returned component.

use std::mem::size_of;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{Archive, ArchiveLimits, ArchiveObject, RawMessage};
use litchi_iwa_protos::keynote_media_lifecycle_codec as lifecycle_codec;

use super::budget::LifecycleBudget;
use super::graph::MediaGraphSelection;
use super::{Package, SlideMediaLifecycleError};

const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;

/// An independently publishable component produced when the source slide node
/// does not share the edited slide component.
#[derive(Debug)]
pub(super) struct NodeCacheComponentEdit {
    pub(super) name: Box<str>,
    pub(super) archive: Archive,
}

/// Prepare the slide-node cache edit for one lifecycle candidate.
///
/// When the node belongs to `edited_component_name`, `edited_archive` is
/// mutated in place and `None` is returned.  Otherwise the source node
/// component is cloned under the operation budget and returned as an owned
/// component edit.
pub(super) fn prepare_slide_node_build_cache(
    source: &Package,
    selection: &MediaGraphSelection,
    edited_component_name: &str,
    edited_archive: &mut Archive,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Option<NodeCacheComponentEdit>, SlideMediaLifecycleError> {
    if selection.build_ids.is_empty() && selection.chunk_ids.is_empty() {
        return Ok(None);
    }

    let slide_record = source
        .slide_record_at(selection.slide_position.get())
        .map_err(|_| SlideMediaLifecycleError::Read)?
        .ok_or(SlideMediaLifecycleError::SlidePositionNotFound {
            position: selection.slide_position,
        })?;
    if slide_record.slide_identifier != selection.slide_identifier {
        return Err(SlideMediaLifecycleError::PatchConflict);
    }
    let node_identifier = slide_record.node_identifier;
    if node_identifier == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let (node_component_name, source_node) = source
        .object_with_component(node_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if node_component_name == super::METADATA_COMPONENT {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let source_node_index = unique_slide_node_message(source_node)?;
    let source_node_payload = source_node
        .messages
        .get(source_node_index)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?
        .data
        .as_slice();
    if source_node_cache_is_invalidated(source_node_payload, wire_limits, budget)? {
        return Ok(None);
    }

    if node_component_name == edited_component_name {
        rewrite_node_cache(
            edited_archive,
            node_identifier,
            archive_limits,
            wire_limits,
            budget,
        )?;
        validate_candidate_node_cache(edited_archive, node_identifier, wire_limits, budget)?;
        return Ok(None);
    }

    let source_component = source
        .state
        .source
        .components()
        .get(node_component_name)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let node_archive =
        clone_component_with_budget(source_component.archive(), archive_limits, budget)?;
    let mut node_archive = node_archive;
    rewrite_node_cache(
        &mut node_archive,
        node_identifier,
        archive_limits,
        wire_limits,
        budget,
    )?;
    validate_candidate_node_cache(&node_archive, node_identifier, wire_limits, budget)?;
    let name = copy_component_name(node_component_name, budget)?;
    Ok(Some(NodeCacheComponentEdit {
        name,
        archive: node_archive,
    }))
}

/// Validate a candidate node cache after the candidate component has been
/// assembled.  The helper is intentionally public only inside the lifecycle
/// module so final readback can reuse the same witness without exposing native
/// object identifiers in the format API.
pub(super) fn validate_candidate_node_cache(
    candidate_node_archive: &Archive,
    node_identifier: u64,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let object = find_archive_object(candidate_node_archive, node_identifier, budget)?;
    let index = unique_slide_node_message(object)?;
    let payload = object
        .messages
        .get(index)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?
        .data
        .as_slice();
    let options = super::lifecycle_options(wire_limits, payload.len())?;
    let (snapshot, report) =
        lifecycle_codec::node_cache::decode_slide_node_build_cache_with_report(payload, options)
            .map_err(map_lifecycle_decode_error)?;
    charge_node_cache_decode_report(report, budget)?;

    if !is_invalidated(snapshot) {
        return Err(SlideMediaLifecycleError::Verification);
    }
    Ok(())
}

/// Validate the node cache in a fully reopened candidate package.
pub(super) fn validate_candidate_package_node_cache(
    candidate: &Package,
    selection: &MediaGraphSelection,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if selection.build_ids.is_empty() && selection.chunk_ids.is_empty() {
        return Ok(());
    }
    let slide_record = candidate
        .slide_record_at(selection.slide_position.get())
        .map_err(|_| SlideMediaLifecycleError::Read)?
        .ok_or(SlideMediaLifecycleError::SlidePositionNotFound {
            position: selection.slide_position,
        })?;
    if slide_record.slide_identifier != selection.slide_identifier
        || slide_record.node_identifier == 0
    {
        return Err(SlideMediaLifecycleError::Verification);
    }
    let (_, node) = candidate
        .object_with_component(slide_record.node_identifier)
        .ok_or(SlideMediaLifecycleError::Verification)?;
    let index = unique_slide_node_message(node)?;
    let payload = node
        .messages
        .get(index)
        .ok_or(SlideMediaLifecycleError::Verification)?
        .data
        .as_slice();
    if !source_node_cache_is_invalidated(payload, wire_limits, budget)? {
        return Err(SlideMediaLifecycleError::Verification);
    }
    Ok(())
}

fn source_node_cache_is_invalidated(
    payload: &[u8],
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<bool, SlideMediaLifecycleError> {
    let options = super::lifecycle_options(wire_limits, payload.len())?;
    let (snapshot, report) =
        lifecycle_codec::node_cache::decode_slide_node_build_cache_with_report(payload, options)
            .map_err(map_lifecycle_decode_error)?;
    charge_node_cache_decode_report(report, budget)?;
    Ok(is_invalidated(snapshot))
}

fn is_invalidated(snapshot: lifecycle_codec::node_cache::SlideNodeBuildCacheSnapshot<'_>) -> bool {
    snapshot.build_event_count().is_none()
        && snapshot.build_event_count_cache_version() == Some(u32::MAX)
        && snapshot.has_explicit_builds().is_none()
        && snapshot.has_explicit_builds_cache_version() == Some(u32::MAX)
}

fn rewrite_node_cache(
    archive: &mut Archive,
    node_identifier: u64,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    let index = {
        let object = find_archive_object(archive, node_identifier, budget)?;
        unique_slide_node_message(object)?
    };
    let source_payload = archive
        .object(node_identifier)
        .and_then(|object| object.messages.get(index))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?
        .data
        .as_slice();
    let options = super::lifecycle_options(wire_limits, source_payload.len())?;
    // The neutral codec owns the native invalidation representation: it
    // removes stale count/presence scalars and marks both cache versions
    // invalid while preserving all unknown SlideNodeArchive fields.
    let edit = lifecycle_codec::node_cache::SlideNodeBuildCacheEdit::invalidate();
    let prepared = lifecycle_codec::node_cache::prepare_slide_node_build_cache_rewrite(
        source_payload,
        edit,
        options,
    )
    .map_err(map_lifecycle_decode_error)?;
    budget.charge_allocations(prepared.output_bytes())?;
    let (data, report) = prepared.commit().map_err(map_lifecycle_decode_error)?;
    charge_node_cache_rewrite_report(report, budget)?;

    let source_object = archive
        .object(node_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let narrowed_limits =
        super::graph::reserve_core_header_work(source_object, 0, archive_limits, budget)?;
    let object = archive
        .object_mut(node_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data,
            },
            narrowed_limits,
        )
        .map(|_| ())
        .map_err(map_archive_error)
}

fn unique_slide_node_message(object: &ArchiveObject) -> Result<usize, SlideMediaLifecycleError> {
    if object.archive_info.message_infos.len() != object.messages.len() {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let mut index = None;
    for (position, message) in object.messages.iter().enumerate() {
        if message.type_ != SLIDE_NODE_MESSAGE_TYPE {
            continue;
        }
        if index.replace(position).is_some() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    index.ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn find_archive_object<'archive>(
    archive: &'archive Archive,
    identifier: u64,
    budget: &mut LifecycleBudget,
) -> Result<&'archive ArchiveObject, SlideMediaLifecycleError> {
    budget.charge_wire_work(archive.objects.len())?;
    archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn clone_component_with_budget(
    source: &Archive,
    archive_limits: ArchiveLimits,
    budget: &mut LifecycleBudget,
) -> Result<Archive, SlideMediaLifecycleError> {
    let object_count = source.objects.len();
    budget.charge_entries(object_count)?;
    // Admit the complete source-relative traversal before asking the core to
    // calculate its encoded bound.  This prevents an archive clone from
    // hiding linear metadata work behind an infallible `Clone` implementation.
    budget.charge_wire_work(
        object_count
            .checked_mul(128)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?,
    )?;
    let encoded_bytes = source
        .encoded_len_with_limits(archive_limits)
        .map_err(map_archive_error)?;
    let mut messages = 0usize;
    let mut message_infos = 0usize;
    let mut fields = 0usize;
    let mut references = 0usize;
    for object in &source.objects {
        messages = messages
            .checked_add(object.messages.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        message_infos = message_infos
            .checked_add(object.archive_info.message_infos.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        for info in &object.archive_info.message_infos {
            fields = fields
                .checked_add(info.field_infos.len())
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            references = references
                .checked_add(info.object_references.len())
                .and_then(|total| total.checked_add(info.data_references.len()))
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            for field in &info.field_infos {
                references = references
                    .checked_add(field.object_references.len())
                    .and_then(|total| total.checked_add(field.data_references.len()))
                    .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            }
        }
    }
    let structural_bytes = object_count
        .checked_mul(size_of::<ArchiveObject>())
        .and_then(|bytes| bytes.checked_add(messages.checked_mul(size_of::<RawMessage>())?))
        .and_then(|bytes| {
            bytes.checked_add(message_infos.checked_mul(size_of::<litchi_iwa_core::MessageInfo>())?)
        })
        .and_then(|bytes| {
            bytes.checked_add(fields.checked_mul(size_of::<litchi_iwa_core::FieldInfo>())?)
        })
        .and_then(|bytes| bytes.checked_add(references.checked_mul(size_of::<u64>())?))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let allocation_bound = encoded_bytes
        .checked_add(structural_bytes)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocations(allocation_bound)?;
    Ok(source.clone())
}

fn copy_component_name(
    name: &str,
    budget: &mut LifecycleBudget,
) -> Result<Box<str>, SlideMediaLifecycleError> {
    budget.charge_allocations(name.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(name.len())
        .map_err(|_| SlideMediaLifecycleError::Allocation { amount: name.len() })?;
    output.push_str(name);
    Ok(output.into_boxed_str())
}

fn charge_node_cache_decode_report(
    report: lifecycle_codec::DecodeReport,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    super::charge_lifecycle_decode_report(report, budget)
}

fn charge_node_cache_rewrite_report(
    report: lifecycle_codec::RewriteReport,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(
        usize::try_from(report.max_depth()).map_err(|_| SlideMediaLifecycleError::InvalidSource)?,
    )?;
    budget.charge_allocations(report.scratch_bytes())?;
    for _ in 0..report.allocations() {
        budget.charge_allocations(0)?;
    }
    Ok(())
}

fn map_lifecycle_decode_error(error: lifecycle_codec::DecodeError) -> SlideMediaLifecycleError {
    use lifecycle_codec::DecodeLimit;
    let Some(limit) = error.resource_limit() else {
        return SlideMediaLifecycleError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        DecodeLimit::Bytes { observed, maximum } => (
            super::SlideMediaLifecycleLimitKind::InputBytes,
            observed,
            maximum,
        ),
        DecodeLimit::Fields { observed, maximum } => (
            super::SlideMediaLifecycleLimitKind::WireFields,
            observed,
            maximum,
        ),
        DecodeLimit::Work { observed, maximum } => (
            super::SlideMediaLifecycleLimitKind::WireWork,
            observed,
            maximum,
        ),
        DecodeLimit::OutputBytes { observed, maximum } => (
            super::SlideMediaLifecycleLimitKind::OutputBytes,
            observed,
            maximum,
        ),
        DecodeLimit::References { observed, maximum } => (
            super::SlideMediaLifecycleLimitKind::References,
            observed,
            maximum,
        ),
        DecodeLimit::Nesting { observed, maximum } => {
            return SlideMediaLifecycleError::LimitExceeded {
                kind: super::SlideMediaLifecycleLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        _ => return SlideMediaLifecycleError::InvalidSource,
    };
    let observed = match u64::try_from(observed) {
        Ok(value) => value,
        Err(_) => return SlideMediaLifecycleError::InvalidSource,
    };
    let maximum = match u64::try_from(maximum) {
        Ok(value) => value,
        Err(_) => return SlideMediaLifecycleError::InvalidSource,
    };
    SlideMediaLifecycleError::LimitExceeded {
        kind,
        observed,
        maximum,
    }
}

fn map_archive_error(error: litchi_iwa_core::Error) -> SlideMediaLifecycleError {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::Entries,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideMediaLifecycleError::Allocation { amount: requested }
        },
        _ => SlideMediaLifecycleError::InvalidSource,
    }
}
