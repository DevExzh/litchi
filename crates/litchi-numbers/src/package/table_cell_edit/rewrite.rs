//! Grouped physical publication for a table-cell transaction.
//!
//! This module deliberately knows nothing about tiles, strings, formulae, or
//! table semantics.  Its caller gives it final raw message payloads.  In
//! return it changes each selected IWA component once, removes the canonical
//! root previews once, and reopens one complete candidate package.

use std::{cmp::Ordering, collections::HashSet, fmt};

use litchi_iwa_archive::{
    SourceCatalog,
    package::{Entry, EntryEdit},
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};

use super::super::Package;

/// The three root previews invalidated by a native Numbers cell mutation.
pub(super) const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// One final raw payload for an existing IWA message.
///
/// The payload is borrowed so the semantic transaction remains the sole owner
/// of its bounded mutation buffer.  Replacements are grouped by
/// `component_index`; component groups must be monotonic and message
/// coordinates within each group must be strictly sorted and duplicate-free,
/// so publication can traverse the plan without a second order buffer.
#[derive(Clone, Copy, Debug)]
pub(super) struct MessageReplacement<'a> {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) expected_type: u32,
    pub(super) payload: &'a [u8],
    /// Exact aggregate-only `MessageInfo` object-reference transition.
    ///
    /// This narrow seam covers weak metadata with no selected `FieldInfo`;
    /// richer field-local changes use [`MessageEdit`] in the staged writer.
    pub(super) references: Option<AggregateReferenceDelta<'a>>,
}

/// Borrowed exact aggregate-only archive-header reference transition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct AggregateReferenceDelta<'a> {
    pub(super) before: &'a [u64],
    pub(super) after: &'a [u64],
}

/// The physical portion of a prepared table-cell transaction.
#[derive(Clone, Copy, Debug)]
pub(super) struct RewritePlan<'a> {
    pub(super) replacements: &'a [MessageReplacement<'a>],
    /// Must be the exact subset of [`ROOT_PREVIEWS`] present in the source.
    pub(super) preview_deletions: &'a [&'static str],
}

/// One exact nested `FieldInfo` object-reference transition.
///
/// The writer converts this to the core archive helper's borrowed selector
/// only while the owning component is open. Keeping the path and identifiers
/// owned here lets leaf planners release their temporary buffers before the
/// package is reassembled.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct FieldReferenceDelta {
    pub(super) field_info_index: usize,
    pub(super) expected_path: Vec<u32>,
    pub(super) before: Vec<u64>,
    pub(super) after: Vec<u64>,
}

/// Complete exact object-reference metadata transition for one message.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct ReferenceDelta {
    pub(super) aggregate_before: Vec<u64>,
    pub(super) aggregate_after: Vec<u64>,
    pub(super) fields: Vec<FieldReferenceDelta>,
}

/// One owned payload and exact archive-header transition.
///
/// `references`, when present, is the exact aggregate and selected-field
/// before/after state. The core transition applies simultaneous pruning,
/// retained-reference reordering, and appending in one atomic raw-header edit.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct MessageEdit {
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) expected_type: u32,
    pub(super) payload: Vec<u8>,
    pub(super) references: Option<ReferenceDelta>,
}

/// One exact source object which becomes unreachable in the final overlay.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct ObjectDeletion {
    pub(super) object_index: usize,
    pub(super) expected_identifier: u64,
}

/// A final staged object moved atomically into its selected component.
pub(super) type NewObjectEdit = ArchiveObject;

/// All staged mutations for one exact source component.
///
/// Component edits, messages, and new objects must be strictly sorted. The
/// grouped writer rejects a non-deterministic or duplicate plan instead of
/// silently normalizing ownership decisions made by an upstream planner.
#[derive(Debug, PartialEq)]
pub(super) struct ComponentEdit {
    pub(super) component_index: usize,
    pub(super) messages: Vec<MessageEdit>,
    pub(super) object_deletions: Vec<ObjectDeletion>,
    pub(super) new_objects: Vec<NewObjectEdit>,
}

/// Owned sparse-capable grouped publication plan.
#[derive(Debug, PartialEq)]
pub(super) struct StagedRewritePlan<'a> {
    pub(super) component_edits: Vec<ComponentEdit>,
    /// Must be the exact subset of [`ROOT_PREVIEWS`] present in the source.
    pub(super) preview_deletions: &'a [&'static str],
}

/// Whether a published payload replaced source state or belongs to a newly
/// appended object.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum PublishedMessageKind {
    Existing,
    Deleted,
    Appended,
}

/// Directional object topology change retained for locality verification.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum PublishedObjectKind {
    Deleted,
    Appended,
}

/// Deterministic exact payload evidence retained after the staged buffers are
/// consumed by archive publication.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct PublishedMessage {
    pub(super) component_index: usize,
    pub(super) object_identifier: u64,
    pub(super) source_object_index: Option<usize>,
    pub(super) target_object_index: Option<usize>,
    pub(super) message_index: usize,
    pub(super) expected_type: u32,
    pub(super) kind: PublishedMessageKind,
    pub(super) payload: Vec<u8>,
}

/// Deterministic coordinate and identifier of one appended object.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct PublishedObject {
    pub(super) component_index: usize,
    pub(super) source_object_index: Option<usize>,
    pub(super) target_object_index: Option<usize>,
    pub(super) identifier: u64,
    pub(super) kind: PublishedObjectKind,
}

/// Direction-specific work needed to reopen one exact package artifact.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct ReopenCost {
    pub(super) work: u64,
    pub(super) references: u64,
}

/// Publication counters reserved before the focused package is reassembled.
///
/// Every counter is exact: component serialization and Snappy compression
/// finish before this reservation is formed, and stored-IWA ZIP sizing is
/// arithmetic. A failed reservation therefore allocates no output artifact
/// and performs no candidate reopen.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct PublicationReservation {
    pub(super) components_reassembled: u64,
    pub(super) reassembly_bytes: u64,
    pub(super) preview_bytes_deleted: u64,
    pub(super) locality_bytes: u64,
    /// Conservative work envelope for the post-publication locality proof.
    /// The sparse orchestrator replaces this with `locality::Report::work`
    /// before it records the exact successful publication.
    pub(super) locality_work: u64,
    pub(super) output_artifact_allocations: u64,
    pub(super) output_bytes: u64,
    pub(super) candidate_reopens: u64,
    pub(super) source_reopen: ReopenCost,
    pub(super) target_reopen: ReopenCost,
}

/// Exact publication counters observed for a successful candidate.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct PublicationCost {
    pub(super) components_reassembled: u64,
    pub(super) reassembly_bytes: u64,
    pub(super) preview_bytes_deleted: u64,
    pub(super) locality_bytes: u64,
    /// Initially the checked locality-proof envelope; replaced with the exact
    /// locality report work before publication is recorded.
    pub(super) locality_work: u64,
    pub(super) output_artifact_allocations: u64,
    pub(super) output_bytes: u64,
    pub(super) candidate_reopens: u64,
    pub(super) source_reopen: ReopenCost,
    pub(super) target_reopen: ReopenCost,
}

/// Exact byte and topology counters from the touched-component pipeline.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct ComponentCost {
    pub(super) components: u64,
    pub(super) compressed_input_bytes: u64,
    pub(super) decoded_input_bytes: u64,
    pub(super) serialized_output_bytes: u64,
    pub(super) compressed_output_bytes: u64,
    pub(super) retained_evidence_bytes: u64,
    pub(super) retained_elements: u64,
    pub(super) peak_scratch_bytes: u64,
    pub(super) allocation_events: u64,
    pub(super) reference_edits: u64,
    pub(super) reference_items: u64,
    pub(super) appended_objects: u64,
    pub(super) deleted_objects: u64,
    pub(super) work: u64,
}

/// Whether the grouped writer should retain payload-bearing publication
/// evidence after candidate reopen.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) enum EvidenceRetention {
    /// Retain exact message/object evidence for callers which do not already
    /// own a normalized patch-evidence allocation.
    #[default]
    Retain,
    /// Omit writer evidence because the caller already owns exact coordinates.
    Omit,
}

/// Conservative component-pipeline envelope authorized before any touched
/// component is decoded or any writer-owned component buffer is allocated.
///
/// Source-side counters and staged topology counters are exact. Output,
/// scratch, allocation, and work counters are hard upper bounds derived from
/// the package's validated archive limits. The later [`ComponentCost`] remains
/// the exact successful observation used for settlement.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct ComponentReservation {
    pub(super) components: u64,
    pub(super) compressed_input_bytes: u64,
    pub(super) decoded_input_bytes: u64,
    pub(super) staged_payload_bytes: u64,
    pub(super) maximum_serialized_output_bytes: u64,
    pub(super) maximum_compressed_output_bytes: u64,
    pub(super) maximum_retained_evidence_bytes: u64,
    pub(super) maximum_retained_elements: u64,
    pub(super) maximum_peak_bytes: u64,
    pub(super) maximum_allocation_events: u64,
    pub(super) reference_items: u64,
    pub(super) appended_objects: u64,
    pub(super) deleted_objects: u64,
    pub(super) work: u64,
}

/// Final result of the one grouped physical publication.
pub(super) struct RewriteOutcome {
    pub(super) package: Package,
    pub(super) touched_components: usize,
    pub(super) publication: PublicationCost,
    /// Source-present evidence in source order, followed by source-absent
    /// appended evidence in target order. Production callers own normalized
    /// patch evidence and request [`EvidenceRetention::Omit`].
    #[cfg(test)]
    pub(super) published_messages: Vec<PublishedMessage>,
}

/// Content-free failure from the raw grouped writer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum RewriteError {
    UnsupportedSource,
    InvalidSource,
    Limit,
    Allocation { amount: usize },
    Verification,
    Precharge,
    Candidate,
}

impl fmt::Display for RewriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::UnsupportedSource => "the Numbers source cannot be physically rewritten",
            Self::InvalidSource => "the Numbers source is malformed for a focused rewrite",
            Self::Limit => "the focused Numbers rewrite exceeded a physical limit",
            Self::Allocation { .. } => "the focused Numbers rewrite could not allocate memory",
            Self::Verification => "the focused Numbers rewrite failed verification",
            Self::Precharge => "the focused Numbers rewrite publication was not authorized",
            Self::Candidate => "the focused Numbers rewrite could not reopen its candidate",
        })
    }
}

impl std::error::Error for RewriteError {}

/// Reassemble and reopen one changed candidate after an exact precharge.
///
/// `precharge` is called exactly once after component serialization but before
/// the output package allocation. It receives exact output bytes, events, and
/// directional-reopen counts. The callback is intentionally
/// small: it lets the table-cell budget reserve publication atomically without
/// making this raw writer depend on that budget's still-private representation.
#[cfg(test)]
pub(super) fn rewrite_with_precharge(
    source: &Package,
    plan: RewritePlan<'_>,
    mut precharge: impl FnMut(PublicationReservation) -> Result<(), RewriteError>,
) -> Result<RewriteOutcome, RewriteError> {
    rewrite_with_authorization(
        source,
        plan,
        |_reservation| Ok(()),
        |reservation, _cost| precharge(reservation),
    )
}

/// Publish a classic scalar/text plan with admission before writer buffers.
#[cfg(test)]
pub(super) fn rewrite_with_authorization(
    source: &Package,
    plan: RewritePlan<'_>,
    authorize_components: impl FnMut(ComponentReservation) -> Result<(), RewriteError>,
    precharge: impl FnMut(PublicationReservation, ComponentCost) -> Result<(), RewriteError>,
) -> Result<RewriteOutcome, RewriteError> {
    rewrite_with_evidence_authorization(
        source,
        plan,
        EvidenceRetention::Retain,
        authorize_components,
        precharge,
    )
}

/// Authorized classic publication with caller-selected evidence retention.
pub(super) fn rewrite_with_evidence_authorization(
    source: &Package,
    plan: RewritePlan<'_>,
    evidence_retention: EvidenceRetention,
    mut authorize_components: impl FnMut(ComponentReservation) -> Result<(), RewriteError>,
    mut precharge: impl FnMut(PublicationReservation, ComponentCost) -> Result<(), RewriteError>,
) -> Result<RewriteOutcome, RewriteError> {
    if plan.replacements.is_empty() {
        return Err(RewriteError::InvalidSource);
    }
    let source_catalog = physical_source(source)?;
    validate_preview_deletions(source_catalog, plan.preview_deletions)?;
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let (component_count, replacement_order) =
        validate_classic_replacements(source_catalog, plan.replacements)?;
    let classic_reservation = classic_component_reservation(
        source_catalog,
        plan.replacements,
        component_count,
        replacement_order,
        archive_limits,
    )?;
    authorize_components(classic_reservation.reservation)?;

    let mut physical_entries = Vec::new();
    physical_entries
        .try_reserve_exact(plan.replacements.len())
        .map_err(|_error| RewriteError::Allocation {
            amount: plan.replacements.len(),
        })?;
    physical_entries.resize(plan.replacements.len(), None);
    index_classic_entries(
        source_catalog,
        plan.replacements,
        replacement_order,
        &mut physical_entries,
    )?;
    let source_reopen = reopen_cost(source_catalog, source.source_bytes().len(), &[], &[])?;

    let mut components = Vec::new();
    components
        .try_reserve_exact(component_count)
        .map_err(|_error| RewriteError::Allocation {
            amount: component_count,
        })?;
    let mut range_cursor = match replacement_order {
        ClassicOrder::Ascending => 0,
        ClassicOrder::Descending => plan.replacements.len(),
    };
    while let Some(range) =
        next_classic_range(plan.replacements, replacement_order, &mut range_cursor)
    {
        let component_index = plan.replacements[range.start].component_index;
        let mut matching_entries = physical_entries[range.clone()]
            .iter()
            .filter_map(|entry| *entry);
        let entry = matching_entries.next().ok_or(RewriteError::InvalidSource)?;
        if matching_entries.next().is_some() {
            return Err(RewriteError::InvalidSource);
        }
        components.push(rewrite_component(
            source,
            component_index,
            entry,
            &plan.replacements[range],
            archive_limits,
            snappy_limits,
            evidence_retention,
        )?);
    }
    drop(physical_entries);
    let component_cost = component_cost(
        &components,
        1,
        classic_reservation.admission_reference_items,
        classic_reservation.traversal_work,
    )?;
    validate_component_cost(component_cost, classic_reservation.reservation)?;
    #[cfg(test)]
    let published_messages = take_publication_messages(&mut components)?;

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(components.len())
        .map_err(|_error| RewriteError::Allocation {
            amount: components.len(),
        })?;
    for component in &components {
        edits.push(EntryEdit::new(component.name, &component.compressed));
    }
    let output_bytes = reassembled_output_len(source_catalog, &components, plan.preview_deletions)?;
    let target_reopen = target_reopen_cost(
        source_catalog,
        &components,
        plan.preview_deletions,
        output_bytes,
    )?;
    let reservation = publication_reservation(
        source_catalog,
        &components,
        plan.preview_deletions,
        output_bytes,
        source_reopen,
        target_reopen,
    )?;
    let output =
        with_publication_authorization(reservation, component_cost, &mut precharge, || {
            source_catalog
                .package()
                .reassemble_with_deletions_to_bytes(&edits, plan.preview_deletions, physical_limits)
                .map_err(map_archive_error)
        })?;
    drop(edits);
    if output.len() != output_bytes {
        return Err(RewriteError::Verification);
    }
    let publication = publication_cost(
        source_catalog,
        &components,
        plan.preview_deletions,
        output.len(),
        source_reopen,
        target_reopen,
    )?;
    let package = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(|_error| RewriteError::Candidate)?;
    Ok(RewriteOutcome {
        package,
        touched_components: components.len(),
        publication,
        #[cfg(test)]
        published_messages,
    })
}

/// Authorized staged publication with caller-selected evidence retention.
pub(super) fn rewrite_staged_with_evidence_authorization(
    source: &Package,
    plan: StagedRewritePlan<'_>,
    evidence_retention: EvidenceRetention,
    mut authorize_components: impl FnMut(ComponentReservation) -> Result<(), RewriteError>,
    mut precharge: impl FnMut(PublicationReservation, ComponentCost) -> Result<(), RewriteError>,
) -> Result<RewriteOutcome, RewriteError> {
    validate_staged_plan(&plan)?;
    let source_catalog = physical_source(source)?;
    validate_preview_deletions(source_catalog, plan.preview_deletions)?;
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let component_reservation = component_reservation(
        source,
        source_catalog,
        &plan.component_edits,
        archive_limits,
    )?;
    let preview_deletions = plan.preview_deletions;
    let component_count = plan.component_edits.len();
    let (source_reopen, components) = with_component_authorization(
        component_reservation.reservation,
        &mut authorize_components,
        || {
            validate_staged_source(source, &plan)?;
            let source_reopen = reopen_cost(source_catalog, source.source_bytes().len(), &[], &[])?;
            let mut components = Vec::new();
            components
                .try_reserve_exact(component_count)
                .map_err(|_error| RewriteError::Allocation {
                    amount: component_count,
                })?;
            for edit in plan.component_edits {
                components.push(rewrite_staged_component(
                    source,
                    source_catalog,
                    edit,
                    archive_limits,
                    snappy_limits,
                    evidence_retention,
                )?);
            }
            Ok((source_reopen, components))
        },
    )?;
    #[cfg(test)]
    let mut components = components;
    let identifier_set_allocation =
        usize::from(component_reservation.reservation.appended_objects != 0);
    let component_cost = component_cost(
        &components,
        identifier_set_allocation,
        component_reservation.admission_reference_items,
        0,
    )?;
    validate_component_cost(component_cost, component_reservation.reservation)?;
    #[cfg(test)]
    let published_messages = take_publication_messages(&mut components)?;

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(components.len())
        .map_err(|_error| RewriteError::Allocation {
            amount: components.len(),
        })?;
    for component in &components {
        edits.push(EntryEdit::new(component.name, &component.compressed));
    }
    let output_bytes = reassembled_output_len(source_catalog, &components, preview_deletions)?;
    let target_reopen =
        target_reopen_cost(source_catalog, &components, preview_deletions, output_bytes)?;
    let reservation = publication_reservation(
        source_catalog,
        &components,
        preview_deletions,
        output_bytes,
        source_reopen,
        target_reopen,
    )?;
    let output =
        with_publication_authorization(reservation, component_cost, &mut precharge, || {
            source_catalog
                .package()
                .reassemble_with_deletions_to_bytes(&edits, preview_deletions, physical_limits)
                .map_err(map_archive_error)
        })?;
    drop(edits);
    if output.len() != output_bytes {
        return Err(RewriteError::Verification);
    }
    let publication = publication_cost(
        source_catalog,
        &components,
        preview_deletions,
        output_bytes,
        source_reopen,
        target_reopen,
    )?;
    let package = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(|_error| RewriteError::Candidate)?;
    let touched_components = components.len();
    Ok(RewriteOutcome {
        package,
        touched_components,
        publication,
        #[cfg(test)]
        published_messages,
    })
}

fn with_component_authorization<T>(
    reservation: ComponentReservation,
    authorize: &mut impl FnMut(ComponentReservation) -> Result<(), RewriteError>,
    work: impl FnOnce() -> Result<T, RewriteError>,
) -> Result<T, RewriteError> {
    authorize(reservation)?;
    work()
}

fn with_publication_authorization<T>(
    reservation: PublicationReservation,
    component_cost: ComponentCost,
    precharge: &mut impl FnMut(PublicationReservation, ComponentCost) -> Result<(), RewriteError>,
    publish: impl FnOnce() -> Result<T, RewriteError>,
) -> Result<T, RewriteError> {
    precharge(reservation, component_cost)?;
    publish()
}

fn validate_component_cost(
    actual: ComponentCost,
    reservation: ComponentReservation,
) -> Result<(), RewriteError> {
    if actual.components > reservation.components
        || actual.compressed_input_bytes > reservation.compressed_input_bytes
        || actual.decoded_input_bytes > reservation.decoded_input_bytes
        || actual.serialized_output_bytes > reservation.maximum_serialized_output_bytes
        || actual.compressed_output_bytes > reservation.maximum_compressed_output_bytes
        || actual.retained_evidence_bytes > reservation.maximum_retained_evidence_bytes
        || actual.retained_elements > reservation.maximum_retained_elements
        || actual.peak_scratch_bytes > reservation.maximum_peak_bytes
        || actual.allocation_events > reservation.maximum_allocation_events
        || actual.reference_items > reservation.reference_items
        || actual.appended_objects != reservation.appended_objects
        || actual.deleted_objects != reservation.deleted_objects
        || actual.work > reservation.work
    {
        return Err(RewriteError::Verification);
    }
    Ok(())
}

struct ClassicReservation {
    reservation: ComponentReservation,
    admission_reference_items: usize,
    traversal_work: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ClassicOrder {
    Ascending,
    Descending,
}

fn validate_classic_replacements(
    source: &SourceCatalog,
    replacements: &[MessageReplacement<'_>],
) -> Result<(usize, ClassicOrder), RewriteError> {
    let mut component_count = 0usize;
    let mut previous: Option<&MessageReplacement<'_>> = None;
    let mut order = None;
    for replacement in replacements {
        source
            .components()
            .get_index(replacement.component_index)
            .ok_or(RewriteError::InvalidSource)?;
        if let Some(previous) = previous {
            if previous.component_index == replacement.component_index {
                if (previous.object_index, previous.message_index)
                    >= (replacement.object_index, replacement.message_index)
                {
                    return Err(RewriteError::InvalidSource);
                }
            } else {
                let next_order = if previous.component_index < replacement.component_index {
                    ClassicOrder::Ascending
                } else {
                    ClassicOrder::Descending
                };
                if order.is_some_and(|order| order != next_order) {
                    return Err(RewriteError::InvalidSource);
                }
                order = Some(next_order);
            }
        }
        if previous.is_none_or(|previous| previous.component_index != replacement.component_index) {
            component_count = checked_add(component_count, 1)?;
        }
        previous = Some(replacement);
    }
    Ok((component_count, order.unwrap_or(ClassicOrder::Ascending)))
}

fn next_classic_range(
    replacements: &[MessageReplacement<'_>],
    order: ClassicOrder,
    cursor: &mut usize,
) -> Option<std::ops::Range<usize>> {
    match order {
        ClassicOrder::Ascending => {
            let start = *cursor;
            let component = replacements.get(start)?.component_index;
            let end = replacements[start..]
                .partition_point(|replacement| replacement.component_index == component)
                + start;
            *cursor = end;
            Some(start..end)
        },
        ClassicOrder::Descending => {
            let end = *cursor;
            let component = replacements.get(end.checked_sub(1)?)?.component_index;
            let start = replacements[..end]
                .partition_point(|replacement| replacement.component_index > component);
            *cursor = start;
            Some(start..end)
        },
    }
}

fn selected_component_position(
    source: &SourceCatalog,
    selection_len: usize,
    mut component_index_at: impl FnMut(usize) -> usize,
    name: &str,
) -> (Option<usize>, usize) {
    counted_binary_search_by(selection_len, |position| {
        source
            .components()
            .get_index(component_index_at(position))
            .map_or(Ordering::Greater, |component| component.name().cmp(name))
    })
}

fn counted_binary_search_by(
    len: usize,
    mut compare: impl FnMut(usize) -> Ordering,
) -> (Option<usize>, usize) {
    let mut base = 0usize;
    let mut size = len;
    let mut comparisons = 0usize;
    while size > 0 {
        let half = size / 2;
        let middle = base + half;
        comparisons += 1;
        match compare(middle) {
            Ordering::Less => {
                base = middle + 1;
                size -= half + 1;
            },
            Ordering::Greater => size = half,
            Ordering::Equal => return (Some(middle), comparisons),
        }
    }
    (None, comparisons)
}

fn index_classic_entries<'a>(
    source: &'a SourceCatalog,
    replacements: &[MessageReplacement<'_>],
    order: ClassicOrder,
    entries: &mut [Option<&'a Entry>],
) -> Result<(), RewriteError> {
    for entry in source.package().iter() {
        let (position, _comparisons) = selected_component_position(
            source,
            replacements.len(),
            |position| {
                replacements[classic_ordered_position(order, replacements.len(), position)]
                    .component_index
            },
            entry.name(),
        );
        if let Some(position) = position {
            let ordered_position = classic_ordered_position(order, replacements.len(), position);
            if entries[ordered_position].replace(entry).is_some() {
                return Err(RewriteError::InvalidSource);
            }
        }
    }
    Ok(())
}

fn classic_ordered_position(order: ClassicOrder, len: usize, position: usize) -> usize {
    match order {
        ClassicOrder::Ascending => position,
        ClassicOrder::Descending => len - position - 1,
    }
}

/// Allocation-free conservative admission for sorted classic replacements.
fn classic_component_reservation(
    source: &SourceCatalog,
    replacements: &[MessageReplacement<'_>],
    expected_components: usize,
    order: ClassicOrder,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<ClassicReservation, RewriteError> {
    let mut components = 0usize;
    let mut compressed_input_bytes = 0usize;
    let mut decoded_input_bytes = 0usize;
    let mut source_allocation_bytes = 0usize;
    let mut reference_items = 0usize;
    let mut entry_visits = 0usize;
    let mut comparisons = 0usize;
    for entry in source.package().iter() {
        entry_visits = checked_add(entry_visits, 1)?;
        let (position, entry_comparisons) = selected_component_position(
            source,
            replacements.len(),
            |position| {
                replacements[classic_ordered_position(order, replacements.len(), position)]
                    .component_index
            },
            entry.name(),
        );
        comparisons = checked_add(comparisons, entry_comparisons)?;
        let Some(position) = position else {
            continue;
        };
        let component = source
            .components()
            .get_index(
                replacements[classic_ordered_position(order, replacements.len(), position)]
                    .component_index,
            )
            .ok_or(RewriteError::InvalidSource)?;
        if entry.is_opaque() || entry.metadata().local().compression_method() != 0 {
            return Err(RewriteError::UnsupportedSource);
        }
        components = checked_add(components, 1)?;
        compressed_input_bytes = checked_add(compressed_input_bytes, entry.data().len())?;
        decoded_input_bytes =
            checked_add(decoded_input_bytes, archive_extent(component.archive())?)?;
        source_allocation_bytes = checked_add(
            source_allocation_bytes,
            archive_allocation_cost(component.archive())?,
        )?;
        reference_items = checked_add(
            reference_items,
            archive_reference_count(component.archive())?,
        )?;
    }
    if components == 0 || components != expected_components {
        return Err(RewriteError::InvalidSource);
    }
    let admission_reference_items = reference_items;
    reference_items = checked_mul(reference_items, 2)?;
    let mut transition_reference_items = 0usize;
    let mut transition_target_items = 0usize;
    for replacement in replacements {
        if let Some(delta) = replacement.references {
            transition_reference_items = checked_add(
                transition_reference_items,
                checked_add(delta.before.len(), delta.after.len())?,
            )?;
            // Every final identifier can contribute to the target archive
            // scan; charging the full after-list remains sound for appends as
            // well as the aggregate-only pruning used by rich storage.
            transition_target_items = checked_add(transition_target_items, delta.after.len())?;
        }
    }
    reference_items = checked_add(reference_items, transition_reference_items)?;
    reference_items = checked_add(reference_items, transition_target_items)?;
    let traversal_work = classic_traversal_work(
        entry_visits,
        comparisons,
        replacements.len(),
        expected_components,
    )?;
    let staged_payload_bytes = replacements.iter().try_fold(0usize, |total, replacement| {
        checked_add(total, replacement.payload.len())
    })?;
    let maximum_serialized_output_bytes =
        checked_mul(components, archive_limits.max_archive_bytes())?;
    let maximum_compressed_output_bytes = checked_mul(
        components,
        maximum_snappy_output(archive_limits.max_archive_bytes())?,
    )?;
    let maximum_retained_evidence_bytes = staged_payload_bytes;
    let maximum_peak_bytes = [
        decoded_input_bytes,
        source_allocation_bytes,
        staged_payload_bytes,
        maximum_serialized_output_bytes,
        maximum_compressed_output_bytes,
    ]
    .into_iter()
    .try_fold(0usize, checked_add)?;
    let maximum_allocation_events = checked_add(
        checked_add(
            checked_mul(components, 20)?,
            checked_mul(replacements.len(), 6)?,
        )?,
        3,
    )?;
    let maximum_retained_elements = checked_add(components, replacements.len())?;
    let work = [
        compressed_input_bytes,
        checked_mul(maximum_peak_bytes, 2)?,
        maximum_allocation_events,
        maximum_retained_elements,
        reference_items,
        components,
        traversal_work,
    ]
    .into_iter()
    .try_fold(0usize, checked_add)?;
    let reservation = ComponentReservation {
        components: as_u64(components)?,
        compressed_input_bytes: as_u64(compressed_input_bytes)?,
        decoded_input_bytes: as_u64(decoded_input_bytes)?,
        staged_payload_bytes: as_u64(staged_payload_bytes)?,
        maximum_serialized_output_bytes: as_u64(maximum_serialized_output_bytes)?,
        maximum_compressed_output_bytes: as_u64(maximum_compressed_output_bytes)?,
        maximum_retained_evidence_bytes: as_u64(maximum_retained_evidence_bytes)?,
        maximum_retained_elements: as_u64(maximum_retained_elements)?,
        maximum_peak_bytes: as_u64(maximum_peak_bytes)?,
        maximum_allocation_events: as_u64(maximum_allocation_events)?,
        reference_items: as_u64(reference_items)?,
        appended_objects: 0,
        deleted_objects: 0,
        work: as_u64(work)?,
    };
    Ok(ClassicReservation {
        reservation,
        admission_reference_items,
        traversal_work,
    })
}

fn classic_traversal_work(
    entry_visits: usize,
    comparisons: usize,
    replacements: usize,
    components: usize,
) -> Result<usize, RewriteError> {
    let one_package_pass = checked_add(entry_visits, comparisons)?;
    let grouped_component_visits = replacements
        .checked_sub(1)
        .and_then(|value| value.checked_add(components))
        .ok_or(RewriteError::InvalidSource)?;
    checked_add(
        checked_mul(one_package_pass, 2)?,
        checked_add(checked_mul(replacements, 3)?, grouped_component_visits)?,
    )
}

struct StagedComponentReservation {
    reservation: ComponentReservation,
    admission_reference_items: usize,
}

fn component_reservation(
    source: &Package,
    source_catalog: &SourceCatalog,
    edits: &[ComponentEdit],
    archive_limits: litchi_iwa_core::Limits,
) -> Result<StagedComponentReservation, RewriteError> {
    let mut compressed_input_bytes = 0usize;
    let mut decoded_input_bytes = 0usize;
    let mut staged_payload_bytes = 0usize;
    let mut retained_evidence_bytes = 0usize;
    let mut source_allocation_bytes = 0usize;
    let mut allocation_event_bound = 2usize;
    let mut reference_items = 0usize;
    let mut admission_reference_items = 0usize;
    let mut evidence_messages = 0usize;
    let mut appended_objects = 0usize;
    let mut deleted_objects = 0usize;

    for edit in edits {
        let component = source
            .state
            .components
            .catalog()
            .get_index(edit.component_index)
            .ok_or(RewriteError::InvalidSource)?;
        let physical_component = source_catalog
            .components()
            .get_index(edit.component_index)
            .filter(|candidate| candidate.name() == component.name())
            .ok_or(RewriteError::InvalidSource)?;
        let entry = source_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == component.name())
            .ok_or(RewriteError::InvalidSource)?;
        if entry.is_opaque() || entry.metadata().local().compression_method() != 0 {
            return Err(RewriteError::UnsupportedSource);
        }
        compressed_input_bytes = checked_add(compressed_input_bytes, entry.data().len())?;
        decoded_input_bytes = checked_add(
            decoded_input_bytes,
            archive_extent(physical_component.archive())?,
        )?;
        source_allocation_bytes = checked_add(
            source_allocation_bytes,
            archive_allocation_cost(physical_component.archive())?,
        )?;
        allocation_event_bound = checked_add(
            allocation_event_bound,
            archive_allocation_event_bound(physical_component.archive())?,
        )?;
        let source_references = archive_reference_count(physical_component.archive())?;
        admission_reference_items = checked_add(admission_reference_items, source_references)?;
        reference_items = checked_add(reference_items, source_references)?;

        for message in &edit.messages {
            physical_component
                .archive()
                .objects
                .get(message.object_index)
                .and_then(|object| object.messages.get(message.message_index))
                .filter(|source_message| source_message.type_ == message.expected_type)
                .ok_or(RewriteError::InvalidSource)?;
            staged_payload_bytes = checked_add(staged_payload_bytes, message.payload.len())?;
            retained_evidence_bytes = checked_add(retained_evidence_bytes, message.payload.len())?;
            evidence_messages = checked_add(evidence_messages, 1)?;
            if let Some(delta) = &message.references {
                reference_items = checked_add(
                    reference_items,
                    checked_add(delta.aggregate_before.len(), delta.aggregate_after.len())?,
                )?;
                for field in &delta.fields {
                    reference_items = checked_add(
                        reference_items,
                        checked_add(
                            field.expected_path.len(),
                            checked_add(field.before.len(), field.after.len())?,
                        )?,
                    )?;
                }
            }
        }
        for deletion in &edit.object_deletions {
            let object = physical_component
                .archive()
                .objects
                .get(deletion.object_index)
                .filter(|object| {
                    object.archive_info.identifier == Some(deletion.expected_identifier)
                })
                .ok_or(RewriteError::InvalidSource)?;
            deleted_objects = checked_add(deleted_objects, 1)?;
            evidence_messages = checked_add(evidence_messages, object.messages.len())?;
            for message in &object.messages {
                retained_evidence_bytes = checked_add(retained_evidence_bytes, message.data.len())?;
            }
        }
        appended_objects = checked_add(appended_objects, edit.new_objects.len())?;
        for object in &edit.new_objects {
            allocation_event_bound = checked_add(
                allocation_event_bound,
                object_allocation_event_bound(object)?,
            )?;
            reference_items = checked_add(reference_items, object_reference_count(object)?)?;
            evidence_messages = checked_add(evidence_messages, object.messages.len())?;
            for message in &object.messages {
                staged_payload_bytes = checked_add(staged_payload_bytes, message.data.len())?;
                retained_evidence_bytes = checked_add(retained_evidence_bytes, message.data.len())?;
            }
        }
    }

    // The admission pass scans every source reference. The successful writer
    // scans the final archive as well; source references plus the explicit
    // before/after and new-object items are a conservative bound for it. One
    // structural item per appended object covers linkage metadata introduced
    // outside the staged object's own reference lists.
    reference_items = checked_add(reference_items, admission_reference_items)?;
    reference_items = checked_add(reference_items, appended_objects)?;

    let maximum_serialized_output_bytes =
        checked_mul(edits.len(), archive_limits.max_archive_bytes())?;
    let maximum_compressed_per_component =
        maximum_snappy_output(archive_limits.max_archive_bytes())?;
    let maximum_compressed_output_bytes =
        checked_mul(edits.len(), maximum_compressed_per_component)?;
    let maximum_peak_bytes = [
        decoded_input_bytes,
        source_allocation_bytes,
        staged_payload_bytes,
        maximum_serialized_output_bytes,
        maximum_compressed_output_bytes,
        retained_evidence_bytes,
    ]
    .into_iter()
    .try_fold(0usize, checked_add)?;
    // Every writer allocation owns at least one byte or one non-zero-sized
    // structural item. Byte totals therefore dominate variable allocation
    // counts; the fixed allowance covers the result/evidence vectors and
    // empty-but-capacity-bearing component containers.
    let maximum_output_frames = checked_mul(
        edits.len(),
        archive_limits
            .max_archive_bytes()
            .div_ceil(SnappyStream::WRITE_CHUNK_SIZE),
    )?;
    let maximum_allocation_events = [
        allocation_event_bound,
        checked_mul(edits.len(), 16)?,
        checked_mul(evidence_messages, 2)?,
        appended_objects,
        reference_items,
        checked_mul(maximum_output_frames, 2)?,
    ]
    .into_iter()
    .try_fold(0usize, checked_add)?;
    let maximum_retained_elements = checked_add(
        edits.len(),
        checked_add(
            evidence_messages,
            checked_add(appended_objects, deleted_objects)?,
        )?,
    )?;
    let work = [
        compressed_input_bytes,
        checked_mul(maximum_peak_bytes, 2)?,
        maximum_allocation_events,
        maximum_retained_elements,
        reference_items,
        appended_objects,
        deleted_objects,
        edits.len(),
    ]
    .into_iter()
    .try_fold(0usize, checked_add)?;
    let reservation = ComponentReservation {
        components: as_u64(edits.len())?,
        compressed_input_bytes: as_u64(compressed_input_bytes)?,
        decoded_input_bytes: as_u64(decoded_input_bytes)?,
        staged_payload_bytes: as_u64(staged_payload_bytes)?,
        maximum_serialized_output_bytes: as_u64(maximum_serialized_output_bytes)?,
        maximum_compressed_output_bytes: as_u64(maximum_compressed_output_bytes)?,
        maximum_retained_evidence_bytes: as_u64(retained_evidence_bytes)?,
        maximum_retained_elements: as_u64(maximum_retained_elements)?,
        maximum_peak_bytes: as_u64(maximum_peak_bytes)?,
        maximum_allocation_events: as_u64(maximum_allocation_events)?,
        reference_items: as_u64(reference_items)?,
        appended_objects: as_u64(appended_objects)?,
        deleted_objects: as_u64(deleted_objects)?,
        work: as_u64(work)?,
    };
    Ok(StagedComponentReservation {
        reservation,
        admission_reference_items,
    })
}

fn object_reference_count(object: &ArchiveObject) -> Result<usize, RewriteError> {
    let mut total = 0usize;
    for info in &object.archive_info.message_infos {
        total = checked_add(total, info.object_references.len())?;
        total = checked_add(total, info.data_references.len())?;
        for field in &info.field_infos {
            total = checked_add(total, field.object_references.len())?;
            total = checked_add(total, field.data_references.len())?;
        }
    }
    Ok(total)
}

fn archive_allocation_event_bound(archive: &Archive) -> Result<usize, RewriteError> {
    archive.objects.iter().try_fold(1usize, |total, object| {
        checked_add(total, object_allocation_event_bound(object)?)
    })
}

fn object_allocation_event_bound(object: &ArchiveObject) -> Result<usize, RewriteError> {
    let header_bytes =
        usize::try_from(object.header_length).map_err(|_error| RewriteError::InvalidSource)?;
    // One event per raw-header or staged-metadata allocation byte dominates
    // every nested protobuf-owned container. The fixed object/message
    // allowances cover canonical header encoding, message vectors, payload
    // copies, and raw-header provenance.
    checked_add(
        header_bytes.max(object_allocation_cost_checked(object)?),
        checked_add(4, checked_mul(object.messages.len(), 3)?)?,
    )
}

fn maximum_snappy_output(input_bytes: usize) -> Result<usize, RewriteError> {
    const RAW_OVERHEAD: usize = 32;
    const FRAME_HEADER: usize = 4;

    let mut total = 0usize;
    let mut remaining = input_bytes;
    while remaining != 0 {
        let chunk = remaining.min(SnappyStream::WRITE_CHUNK_SIZE);
        let compressed = checked_add(checked_add(RAW_OVERHEAD, chunk)?, chunk / 6)?;
        total = checked_add(total, checked_add(FRAME_HEADER, compressed)?)?;
        remaining = remaining
            .checked_sub(chunk)
            .ok_or(RewriteError::InvalidSource)?;
    }
    Ok(total)
}

fn validate_staged_plan(plan: &StagedRewritePlan<'_>) -> Result<(), RewriteError> {
    if plan.component_edits.is_empty() {
        return Err(RewriteError::InvalidSource);
    }
    for (component_position, component) in plan.component_edits.iter().enumerate() {
        if component.messages.is_empty()
            && component.object_deletions.is_empty()
            && component.new_objects.is_empty()
        {
            return Err(RewriteError::InvalidSource);
        }
        if component_position != 0
            && plan.component_edits[component_position - 1].component_index
                >= component.component_index
        {
            return Err(RewriteError::InvalidSource);
        }
        for (message_position, message) in component.messages.iter().enumerate() {
            if message_position != 0 {
                let previous = &component.messages[message_position - 1];
                if (previous.object_index, previous.message_index)
                    >= (message.object_index, message.message_index)
                {
                    return Err(RewriteError::InvalidSource);
                }
            }
            if let Some(references) = &message.references {
                for (field_position, field) in references.fields.iter().enumerate() {
                    if field_position != 0
                        && references.fields[field_position - 1].field_info_index
                            >= field.field_info_index
                    {
                        return Err(RewriteError::InvalidSource);
                    }
                }
            }
        }
        for (deletion_position, deletion) in component.object_deletions.iter().enumerate() {
            if deletion.expected_identifier == 0
                || (deletion_position != 0
                    && component.object_deletions[deletion_position - 1].object_index
                        >= deletion.object_index)
            {
                return Err(RewriteError::InvalidSource);
            }
            if component
                .messages
                .binary_search_by_key(&deletion.object_index, |message| message.object_index)
                .is_ok()
            {
                return Err(RewriteError::InvalidSource);
            }
        }
        for (object_position, object) in component.new_objects.iter().enumerate() {
            let identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or(RewriteError::InvalidSource)?;
            if object_position != 0
                && component.new_objects[object_position - 1]
                    .archive_info
                    .identifier
                    .is_none_or(|previous| previous >= identifier)
            {
                return Err(RewriteError::InvalidSource);
            }
        }
    }
    Ok(())
}

fn validate_staged_source(
    source: &Package,
    plan: &StagedRewritePlan<'_>,
) -> Result<(), RewriteError> {
    let new_object_count = plan
        .component_edits
        .iter()
        .try_fold(0usize, |total, component| {
            checked_add(total, component.new_objects.len())
        })?;
    let mut new_identifiers = HashSet::new();
    new_identifiers
        .try_reserve(new_object_count)
        .map_err(|_error| RewriteError::Allocation {
            amount: new_object_count,
        })?;
    for component in &plan.component_edits {
        for deletion in &component.object_deletions {
            let resolved = source
                .state
                .index
                .resolve_ref_id(&source.state.components, deletion.expected_identifier)
                .map_err(|_error| RewriteError::InvalidSource)?
                .ok_or(RewriteError::InvalidSource)?;
            if resolved.component_index != component.component_index
                || resolved.object_index != deletion.object_index
            {
                return Err(RewriteError::InvalidSource);
            }
        }
        for object in &component.new_objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or(RewriteError::InvalidSource)?;
            if !new_identifiers.insert(identifier) {
                return Err(RewriteError::InvalidSource);
            }
            if source
                .state
                .index
                .resolve_ref_id(&source.state.components, identifier)
                .map_err(|_error| RewriteError::InvalidSource)?
                .is_some()
            {
                return Err(RewriteError::InvalidSource);
            }
        }
    }
    drop(new_identifiers);
    Ok(())
}

struct RewrittenComponent<'a> {
    component_index: usize,
    name: &'a str,
    old_compressed_len: usize,
    source_decoded_bytes: usize,
    decoded_bytes: usize,
    allocation_bytes: usize,
    references: usize,
    compressed: Vec<u8>,
    published_messages: Vec<PublishedMessage>,
    published_objects: Vec<PublishedObject>,
    peak_scratch_bytes: usize,
    allocation_events: usize,
    reference_edits: usize,
    reference_items: usize,
    appended_objects: usize,
    deleted_objects: usize,
}

fn rewrite_component<'a>(
    source: &'a Package,
    component_index: usize,
    entry: &'a Entry,
    replacements: &[MessageReplacement<'_>],
    archive_limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    evidence_retention: EvidenceRetention,
) -> Result<RewrittenComponent<'a>, RewriteError> {
    let component = source
        .state
        .components
        .catalog()
        .get_index(component_index)
        .ok_or(RewriteError::InvalidSource)?;
    let name = component.name();
    if entry.name() != name {
        return Err(RewriteError::InvalidSource);
    }
    // Native IWA members are ZIP-stored because their payload is already a
    // Snappy stream. A Deflate-wrapped component would require a second ZIP
    // compression pass merely to learn its physical output size, defeating
    // exact pre-output authorization.
    if entry.is_opaque() || entry.metadata().local().compression_method() != 0 {
        return Err(RewriteError::UnsupportedSource);
    }
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let source_decoded_bytes = stream.as_bytes().len();
    let source_allocation_bytes = archive_allocation_cost(&archive)?;
    drop(stream);

    let mut published_messages = Vec::new();
    published_messages
        .try_reserve_exact(
            usize::from(evidence_retention == EvidenceRetention::Retain) * replacements.len(),
        )
        .map_err(|_error| RewriteError::Allocation {
            amount: replacements.len(),
        })?;
    for replacement in replacements {
        let object = archive
            .objects
            .get_mut(replacement.object_index)
            .ok_or(RewriteError::InvalidSource)?;
        let message = object
            .messages
            .get(replacement.message_index)
            .ok_or(RewriteError::InvalidSource)?;
        let metadata = object
            .archive_info
            .message_infos
            .get(replacement.message_index)
            .ok_or(RewriteError::InvalidSource)?;
        let object_identifier = object
            .archive_info
            .identifier
            .ok_or(RewriteError::InvalidSource)?;
        if message.type_ != replacement.expected_type || metadata.type_ != replacement.expected_type
        {
            return Err(RewriteError::InvalidSource);
        }
        replace_classic_message(object, replacement, archive_limits)?;
        if evidence_retention == EvidenceRetention::Retain {
            published_messages.push(PublishedMessage {
                component_index,
                object_identifier,
                source_object_index: Some(replacement.object_index),
                target_object_index: Some(replacement.object_index),
                message_index: replacement.message_index,
                expected_type: replacement.expected_type,
                kind: PublishedMessageKind::Existing,
                payload: copy_payload(replacement.payload)?,
            });
        }
    }
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let allocation_bytes = archive_allocation_cost(&archive)?;
    let references_count = archive_reference_count(&archive)?;
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    let decoded_bytes = rewritten.len();
    let peak_scratch_bytes = checked_add(source_decoded_bytes, source_allocation_bytes)?
        .max(checked_add(allocation_bytes, decoded_bytes)?)
        .max(checked_add(decoded_bytes, compressed.len())?);
    let nonempty_payloads = replacements
        .iter()
        .filter(|replacement| !replacement.payload.is_empty())
        .count();
    let reference_edit_count = replacements
        .iter()
        .filter(|replacement| replacement.references.is_some())
        .count();
    let reference_item_count = replacements.iter().try_fold(0usize, |total, replacement| {
        let Some(delta) = replacement.references else {
            return Ok(total);
        };
        checked_add(total, checked_add(delta.before.len(), delta.after.len())?)
    })?;
    let allocation_events = checked_add(
        5,
        checked_add(
            checked_add(nonempty_payloads, checked_mul(nonempty_payloads, 2)?)?,
            checked_mul(reference_edit_count, 3)?,
        )?,
    )?;
    drop(rewritten);
    Ok(RewrittenComponent {
        component_index,
        name,
        old_compressed_len: entry.data().len(),
        source_decoded_bytes,
        decoded_bytes,
        allocation_bytes,
        references: references_count,
        compressed,
        published_messages,
        published_objects: Vec::new(),
        peak_scratch_bytes,
        allocation_events,
        reference_edits: reference_edit_count,
        reference_items: checked_add(references_count, reference_item_count)?,
        appended_objects: 0,
        deleted_objects: 0,
    })
}

fn replace_classic_message(
    object: &mut ArchiveObject,
    replacement: &MessageReplacement<'_>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), RewriteError> {
    let message = RawMessage {
        type_: replacement.expected_type,
        data: copy_payload(replacement.payload)?,
    };
    if let Some(delta) = replacement.references {
        object
            .replace_message_transitioning_object_references_preserving_header_with_limits(
                replacement.message_index,
                message,
                ObjectReferenceTransition {
                    aggregate_before: delta.before,
                    aggregate_after: delta.after,
                    fields: &[],
                },
                limits,
            )
            .map_err(map_core_error)?;
    } else {
        object
            .replace_message_preserving_header_with_limits(
                replacement.message_index,
                message,
                limits,
            )
            .map_err(map_core_error)?;
    }
    Ok(())
}

fn rewrite_staged_component<'a>(
    source: &'a Package,
    source_catalog: &'a SourceCatalog,
    edit: ComponentEdit,
    archive_limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    evidence_retention: EvidenceRetention,
) -> Result<RewrittenComponent<'a>, RewriteError> {
    let component = source
        .state
        .components
        .catalog()
        .get_index(edit.component_index)
        .ok_or(RewriteError::InvalidSource)?;
    let name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(RewriteError::InvalidSource)?;
    // See `rewrite_component`: exact publication sizing requires the native
    // stored-IWA seam rather than a second opaque compression layer.
    if entry.is_opaque() || entry.metadata().local().compression_method() != 0 {
        return Err(RewriteError::UnsupportedSource);
    }
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let source_decoded_bytes = stream.as_bytes().len();
    let source_allocation_bytes = archive_allocation_cost(&archive)?;
    drop(stream);

    let reference_edit_count = edit
        .messages
        .iter()
        .filter(|message| message.references.is_some())
        .count();
    let reference_item_count = edit.messages.iter().try_fold(0usize, |total, message| {
        let Some(delta) = &message.references else {
            return Ok(total);
        };
        let mut items = checked_add(delta.aggregate_before.len(), delta.aggregate_after.len())?;
        for field in &delta.fields {
            items = checked_add(items, checked_add(field.before.len(), field.after.len())?)?;
        }
        checked_add(total, items)
    })?;
    let edited_payload_allocations = edit
        .messages
        .iter()
        .filter(|message| !message.payload.is_empty())
        .count();
    let field_transition_allocations = edit
        .messages
        .iter()
        .filter(|message| {
            message
                .references
                .as_ref()
                .is_some_and(|delta| !delta.fields.is_empty())
        })
        .count();
    let appended_object_count = edit.new_objects.len();
    let deleted_object_count = edit.object_deletions.len();

    let appended_message_count = edit.new_objects.iter().try_fold(0usize, |total, object| {
        checked_add(total, object.messages.len())
    })?;
    let appended_evidence_allocations =
        edit.new_objects.iter().try_fold(0usize, |total, object| {
            checked_add(
                total,
                object
                    .messages
                    .iter()
                    .filter(|message| !message.data.is_empty())
                    .count(),
            )
        })?;
    let deleted_message_count =
        edit.object_deletions
            .iter()
            .try_fold(0usize, |total, deletion| {
                let object = archive
                    .objects
                    .get(deletion.object_index)
                    .filter(|object| {
                        object.archive_info.identifier == Some(deletion.expected_identifier)
                    })
                    .ok_or(RewriteError::InvalidSource)?;
                checked_add(total, object.messages.len())
            })?;
    let deleted_evidence_allocations =
        edit.object_deletions
            .iter()
            .try_fold(0usize, |total, deletion| {
                let object = archive
                    .objects
                    .get(deletion.object_index)
                    .filter(|object| {
                        object.archive_info.identifier == Some(deletion.expected_identifier)
                    })
                    .ok_or(RewriteError::InvalidSource)?;
                checked_add(
                    total,
                    object
                        .messages
                        .iter()
                        .filter(|message| !message.data.is_empty())
                        .count(),
                )
            })?;
    let retain_evidence = evidence_retention == EvidenceRetention::Retain;
    let evidence_count = if retain_evidence {
        checked_add(
            checked_add(edit.messages.len(), appended_message_count)?,
            deleted_message_count,
        )?
    } else {
        0
    };
    let mut published_messages = Vec::new();
    published_messages
        .try_reserve_exact(evidence_count)
        .map_err(|_error| RewriteError::Allocation {
            amount: evidence_count,
        })?;
    let mut published_objects = Vec::new();
    let object_evidence_count = if retain_evidence {
        checked_add(edit.object_deletions.len(), edit.new_objects.len())?
    } else {
        0
    };
    published_objects
        .try_reserve_exact(object_evidence_count)
        .map_err(|_error| RewriteError::Allocation {
            amount: edit
                .object_deletions
                .len()
                .saturating_add(edit.new_objects.len()),
        })?;

    for message_edit in edit.messages {
        let object = archive
            .objects
            .get_mut(message_edit.object_index)
            .ok_or(RewriteError::InvalidSource)?;
        let message = object
            .messages
            .get(message_edit.message_index)
            .ok_or(RewriteError::InvalidSource)?;
        let metadata = object
            .archive_info
            .message_infos
            .get(message_edit.message_index)
            .ok_or(RewriteError::InvalidSource)?;
        let object_identifier = object
            .archive_info
            .identifier
            .ok_or(RewriteError::InvalidSource)?;
        if message.type_ != message_edit.expected_type
            || metadata.type_ != message_edit.expected_type
        {
            return Err(RewriteError::InvalidSource);
        }
        apply_message_edit(object, &message_edit, archive_limits)?;
        let target_object_index =
            shifted_object_index(message_edit.object_index, &edit.object_deletions)?;
        if retain_evidence {
            published_messages.push(PublishedMessage {
                component_index: edit.component_index,
                object_identifier,
                source_object_index: Some(message_edit.object_index),
                target_object_index: Some(target_object_index),
                message_index: message_edit.message_index,
                expected_type: message_edit.expected_type,
                kind: PublishedMessageKind::Existing,
                payload: message_edit.payload,
            });
        }
    }

    for deletion in &edit.object_deletions {
        let object = archive
            .objects
            .get(deletion.object_index)
            .filter(|object| object.archive_info.identifier == Some(deletion.expected_identifier))
            .ok_or(RewriteError::InvalidSource)?;
        if retain_evidence {
            published_objects.push(PublishedObject {
                component_index: edit.component_index,
                source_object_index: Some(deletion.object_index),
                target_object_index: None,
                identifier: deletion.expected_identifier,
                kind: PublishedObjectKind::Deleted,
            });
            for (message_index, message) in object.messages.iter().enumerate() {
                published_messages.push(PublishedMessage {
                    component_index: edit.component_index,
                    object_identifier: deletion.expected_identifier,
                    source_object_index: Some(deletion.object_index),
                    target_object_index: None,
                    message_index,
                    expected_type: message.type_,
                    kind: PublishedMessageKind::Deleted,
                    payload: copy_payload(&message.data)?,
                });
            }
        }
    }
    for deletion in edit.object_deletions.iter().rev() {
        let removed = archive
            .remove_object(deletion.expected_identifier)
            .ok_or(RewriteError::InvalidSource)?;
        if removed.archive_info.identifier != Some(deletion.expected_identifier) {
            return Err(RewriteError::InvalidSource);
        }
    }

    let first_new_object = archive.objects.len();
    for (offset, object) in edit.new_objects.iter().enumerate() {
        let object_index = first_new_object
            .checked_add(offset)
            .ok_or(RewriteError::InvalidSource)?;
        let identifier = object
            .archive_info
            .identifier
            .ok_or(RewriteError::InvalidSource)?;
        if retain_evidence {
            published_objects.push(PublishedObject {
                component_index: edit.component_index,
                source_object_index: None,
                target_object_index: Some(object_index),
                identifier,
                kind: PublishedObjectKind::Appended,
            });
            for (message_index, message) in object.messages.iter().enumerate() {
                published_messages.push(PublishedMessage {
                    component_index: edit.component_index,
                    object_identifier: identifier,
                    source_object_index: None,
                    target_object_index: Some(object_index),
                    message_index,
                    expected_type: message.type_,
                    kind: PublishedMessageKind::Appended,
                    payload: copy_payload(&message.data)?,
                });
            }
        }
    }
    archive
        .append_objects_with_limits(edit.new_objects, archive_limits)
        .map_err(map_core_error)?;
    published_messages.sort_unstable_by_key(message_evidence_key);
    published_objects.sort_unstable_by_key(object_evidence_key);
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let allocation_bytes = archive_allocation_cost(&archive)?;
    let references = archive_reference_count(&archive)?;
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    let decoded_bytes = rewritten.len();
    let peak_scratch_bytes = checked_add(source_decoded_bytes, source_allocation_bytes)?
        .max(checked_add(allocation_bytes, decoded_bytes)?)
        .max(checked_add(decoded_bytes, compressed.len())?);
    let evidence_payload_allocations = if retain_evidence {
        checked_add(deleted_evidence_allocations, appended_evidence_allocations)?
    } else {
        0
    };
    let allocation_events = [
        4usize,
        usize::from(evidence_count != 0),
        usize::from(object_evidence_count != 0),
        edited_payload_allocations,
        field_transition_allocations,
        evidence_payload_allocations,
        usize::from(appended_object_count != 0),
    ]
    .into_iter()
    .try_fold(0usize, checked_add)?;
    drop(rewritten);
    Ok(RewrittenComponent {
        component_index: edit.component_index,
        name,
        old_compressed_len: entry.data().len(),
        source_decoded_bytes,
        decoded_bytes,
        allocation_bytes,
        references,
        compressed,
        published_messages,
        published_objects,
        peak_scratch_bytes,
        allocation_events,
        reference_edits: reference_edit_count,
        reference_items: checked_add(reference_item_count, references)?,
        appended_objects: appended_object_count,
        deleted_objects: deleted_object_count,
    })
}

fn shifted_object_index(
    source_index: usize,
    deletions: &[ObjectDeletion],
) -> Result<usize, RewriteError> {
    let removed_before = deletions.partition_point(|deletion| deletion.object_index < source_index);
    source_index
        .checked_sub(removed_before)
        .ok_or(RewriteError::InvalidSource)
}

fn apply_message_edit(
    object: &mut ArchiveObject,
    edit: &MessageEdit,
    limits: litchi_iwa_core::Limits,
) -> Result<(), RewriteError> {
    let replacement = RawMessage {
        type_: edit.expected_type,
        data: copy_payload(&edit.payload)?,
    };
    let Some(delta) = &edit.references else {
        object
            .replace_message_preserving_header_with_limits(edit.message_index, replacement, limits)
            .map_err(map_core_error)?;
        return Ok(());
    };
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(delta.fields.len())
        .map_err(|_error| RewriteError::Allocation {
            amount: delta.fields.len(),
        })?;
    for field in &delta.fields {
        fields.push(FieldObjectReferenceTransition {
            field_info_index: field.field_info_index,
            expected_path: &field.expected_path,
            before: &field.before,
            after: &field.after,
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            edit.message_index,
            replacement,
            ObjectReferenceTransition {
                aggregate_before: &delta.aggregate_before,
                aggregate_after: &delta.aggregate_after,
                fields: &fields,
            },
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn physical_source(package: &Package) -> Result<&SourceCatalog, RewriteError> {
    let source = package
        .state
        .components
        .physical()
        .ok_or(RewriteError::UnsupportedSource)?;
    if !source.source_is_exact() {
        return Err(RewriteError::UnsupportedSource);
    }
    Ok(source)
}

/// Compute the canonical preview deletion list, rejecting duplicate root names.
pub(super) fn root_preview_deletions(source: &Package) -> Result<Vec<&'static str>, RewriteError> {
    let catalog = physical_source(source)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(ROOT_PREVIEWS.len())
        .map_err(|_error| RewriteError::Allocation {
            amount: ROOT_PREVIEWS.len(),
        })?;
    for name in ROOT_PREVIEWS {
        let count = catalog
            .package()
            .iter()
            .filter(|entry| entry.name() == name)
            .count();
        match count {
            0 => {},
            1 => result.push(name),
            _ => return Err(RewriteError::InvalidSource),
        }
    }
    Ok(result)
}

fn validate_preview_deletions(
    source: &SourceCatalog,
    supplied: &[&'static str],
) -> Result<(), RewriteError> {
    let mut supplied_index = 0usize;
    for name in ROOT_PREVIEWS {
        let count = source
            .package()
            .iter()
            .filter(|entry| entry.name() == name)
            .count();
        match count {
            0 => {},
            1 => {
                if supplied.get(supplied_index) != Some(&name) {
                    return Err(RewriteError::InvalidSource);
                }
                supplied_index = supplied_index
                    .checked_add(1)
                    .ok_or(RewriteError::InvalidSource)?;
            },
            _ => return Err(RewriteError::InvalidSource),
        }
    }
    if supplied_index == supplied.len() {
        Ok(())
    } else {
        Err(RewriteError::InvalidSource)
    }
}

fn publication_reservation(
    source: &SourceCatalog,
    components: &[RewrittenComponent<'_>],
    previews: &[&str],
    output_bytes: usize,
    source_reopen: ReopenCost,
    target_reopen: ReopenCost,
) -> Result<PublicationReservation, RewriteError> {
    let preview_bytes_deleted = source
        .package()
        .iter()
        .filter(|entry| previews.contains(&entry.name()))
        .try_fold(0usize, |total, entry| {
            checked_add(total, entry.data().len())
        })?;
    Ok(PublicationReservation {
        components_reassembled: as_u64(components.len())?,
        reassembly_bytes: as_u64(output_bytes)?,
        preview_bytes_deleted: as_u64(preview_bytes_deleted)?,
        locality_bytes: locality_byte_envelope(source, components, output_bytes)?,
        locality_work: locality_work_envelope(source, output_bytes)?,
        // The archive API produces one Vec and Package retains a distinct
        // Arc<[u8]>.  Count both until ingress accepts a zero-copy Vec owner.
        output_artifact_allocations: 2,
        output_bytes: as_u64(output_bytes)?,
        candidate_reopens: 1,
        source_reopen,
        target_reopen,
    })
}

fn publication_cost(
    source: &SourceCatalog,
    components: &[RewrittenComponent<'_>],
    previews: &[&str],
    output_bytes: usize,
    source_reopen: ReopenCost,
    target_reopen: ReopenCost,
) -> Result<PublicationCost, RewriteError> {
    let reservation = publication_reservation(
        source,
        components,
        previews,
        output_bytes,
        source_reopen,
        target_reopen,
    )?;
    Ok(PublicationCost {
        components_reassembled: reservation.components_reassembled,
        reassembly_bytes: reservation.reassembly_bytes,
        preview_bytes_deleted: reservation.preview_bytes_deleted,
        locality_bytes: reservation.locality_bytes,
        locality_work: reservation.locality_work,
        output_artifact_allocations: reservation.output_artifact_allocations,
        output_bytes: reservation.output_bytes,
        candidate_reopens: reservation.candidate_reopens,
        source_reopen: reservation.source_reopen,
        target_reopen: reservation.target_reopen,
    })
}

/// Bound every byte compared by the directional locality proof.
///
/// ZIP/member comparisons are bounded by both physical artifacts. Selected
/// components are also compared as decoded messages, which is not represented
/// by the compressed package lengths. Charging both source-decoded and final
/// serialized archive bytes remains sound for arbitrary mixed growth and
/// shrinkage and is exact from the already prepared component overlay before
/// ZIP reassembly starts.
fn locality_byte_envelope(
    source: &SourceCatalog,
    components: &[RewrittenComponent<'_>],
    output_bytes: usize,
) -> Result<u64, RewriteError> {
    locality_byte_envelope_from_lengths(source.source_bytes().len(), output_bytes, components)
}

fn locality_byte_envelope_from_lengths(
    source_bytes: usize,
    output_bytes: usize,
    components: &[RewrittenComponent<'_>],
) -> Result<u64, RewriteError> {
    let mut bytes = checked_add(source_bytes, output_bytes)?;
    for component in components {
        bytes = checked_add(bytes, component.source_decoded_bytes)?;
        bytes = checked_add(bytes, component.decoded_bytes)?;
    }
    as_u64(bytes)
}

/// Bound an allocation-free directional locality proof before candidate reopen.
///
/// The verifier compares both ZIP envelopes and walks selected archive
/// topology.  Raw bytes can be compared through multiple ZIP views, while
/// topology inspections are scalar work.  The 16x byte term covers the
/// bounded repeated ZIP spans and the separate topology term covers
/// zero-length structural records.  Both source and output are checked
/// physical artifacts, and every addition/multiplication is checked.
fn locality_work_envelope(
    source: &SourceCatalog,
    output_bytes: usize,
) -> Result<u64, RewriteError> {
    let artifact_bytes = checked_add(source.source_bytes().len(), output_bytes)?;
    let mut topology = 0usize;
    for _entry in source.package().iter() {
        topology = checked_add(topology, 1)?;
    }
    for component in source.components().iter() {
        topology = checked_add(topology, 1)?;
        for object in &component.archive().objects {
            topology = checked_add(topology, 1)?;
            topology = checked_add(topology, object.messages.len())?;
            topology = checked_add(topology, object.archive_info.message_infos.len())?;
        }
    }
    let bytes_work = checked_mul(artifact_bytes, 16)?;
    let topology_work = checked_mul(topology, 8)?;
    as_u64(checked_add(bytes_work, topology_work)?)
}

fn reopen_cost(
    catalog: &SourceCatalog,
    raw_package_bytes: usize,
    replacements: &[RewrittenComponent<'_>],
    deletions: &[&str],
) -> Result<ReopenCost, RewriteError> {
    let mut logical_bytes = 0usize;
    for entry in catalog.package().iter() {
        if deletions.contains(&entry.name()) {
            continue;
        }
        logical_bytes = checked_add(logical_bytes, entry.data().len())?;
    }
    for replacement in replacements {
        logical_bytes = logical_bytes
            .checked_sub(replacement.old_compressed_len)
            .and_then(|value| value.checked_add(replacement.compressed.len()))
            .ok_or(RewriteError::InvalidSource)?;
    }
    let mut decoded_bytes = 0usize;
    let mut structure = 0usize;
    let mut references = 0usize;
    let mut replacement_index = 0usize;
    for (component_index, component) in catalog.components().iter().enumerate() {
        if replacement_index < replacements.len()
            && replacements[replacement_index].component_index < component_index
        {
            return Err(RewriteError::InvalidSource);
        }
        let (decoded, component_structure, component_references) = if replacements
            .get(replacement_index)
            .is_some_and(|replacement| replacement.component_index == component_index)
        {
            let replacement = &replacements[replacement_index];
            replacement_index = replacement_index
                .checked_add(1)
                .ok_or(RewriteError::InvalidSource)?;
            (
                replacement.decoded_bytes,
                replacement.allocation_bytes,
                replacement.references,
            )
        } else {
            (
                archive_extent(component.archive())?,
                archive_allocation_cost(component.archive())?,
                archive_reference_count(component.archive())?,
            )
        };
        decoded_bytes = checked_add(decoded_bytes, decoded)?;
        structure = checked_add(structure, component_structure)?;
        references = checked_add(references, component_references)?;
    }
    if replacement_index != replacements.len() {
        return Err(RewriteError::InvalidSource);
    }
    let work = checked_add(
        checked_add(raw_package_bytes, checked_mul(logical_bytes, 2)?)?,
        checked_add(checked_mul(decoded_bytes, 2)?, structure)?,
    )?;
    Ok(ReopenCost {
        work: as_u64(work)?,
        references: as_u64(references)?,
    })
}

fn target_reopen_cost(
    catalog: &SourceCatalog,
    components: &[RewrittenComponent<'_>],
    deletions: &[&str],
    output_bytes: usize,
) -> Result<ReopenCost, RewriteError> {
    reopen_cost(catalog, output_bytes, components, deletions)
}

fn reassembled_output_len(
    catalog: &SourceCatalog,
    components: &[RewrittenComponent<'_>],
    deletions: &[&str],
) -> Result<usize, RewriteError> {
    let mut output = catalog.source_bytes().len();
    for component in components {
        output = output
            .checked_sub(component.old_compressed_len)
            .and_then(|value| value.checked_add(component.compressed.len()))
            .ok_or(RewriteError::InvalidSource)?;
    }
    for entry in catalog.package().iter() {
        if deletions.contains(&entry.name()) {
            output = output
                .checked_sub(entry.raw_record().local_record().len())
                .and_then(|value| {
                    value.checked_sub(entry.raw_record().central_directory_record().len())
                })
                .ok_or(RewriteError::InvalidSource)?;
        }
    }
    Ok(output)
}

fn copy_payload(source: &[u8]) -> Result<Vec<u8>, RewriteError> {
    let mut target = Vec::new();
    target
        .try_reserve_exact(source.len())
        .map_err(|_error| RewriteError::Allocation {
            amount: source.len(),
        })?;
    target.extend_from_slice(source);
    Ok(target)
}

fn component_cost(
    components: &[RewrittenComponent<'_>],
    pre_component_allocation_events: usize,
    admission_reference_items: usize,
    traversal_work: usize,
) -> Result<ComponentCost, RewriteError> {
    let mut result = ComponentCost {
        components: as_u64(components.len())?,
        ..ComponentCost::default()
    };
    let mut evidence_messages = 0usize;
    let mut evidence_objects = 0usize;
    let mut retained_component_bytes = 0usize;
    let mut aggregate_peak_bytes = 0usize;
    result.reference_items = as_u64(admission_reference_items)?;
    for component in components {
        result.compressed_input_bytes = checked_add_u64(
            result.compressed_input_bytes,
            as_u64(component.old_compressed_len)?,
        )?;
        result.decoded_input_bytes = checked_add_u64(
            result.decoded_input_bytes,
            as_u64(component.source_decoded_bytes)?,
        )?;
        result.serialized_output_bytes = checked_add_u64(
            result.serialized_output_bytes,
            as_u64(component.decoded_bytes)?,
        )?;
        result.compressed_output_bytes = checked_add_u64(
            result.compressed_output_bytes,
            as_u64(component.compressed.len())?,
        )?;
        let mut component_evidence_bytes = 0usize;
        for message in &component.published_messages {
            result.retained_evidence_bytes = checked_add_u64(
                result.retained_evidence_bytes,
                as_u64(message.payload.len())?,
            )?;
            component_evidence_bytes =
                checked_add(component_evidence_bytes, message.payload.len())?;
        }
        evidence_messages = checked_add(evidence_messages, component.published_messages.len())?;
        evidence_objects = checked_add(evidence_objects, component.published_objects.len())?;
        aggregate_peak_bytes = aggregate_peak_bytes.max(checked_add(
            retained_component_bytes,
            checked_add(component.peak_scratch_bytes, component_evidence_bytes)?,
        )?);
        retained_component_bytes = checked_add(
            retained_component_bytes,
            checked_add(component.compressed.len(), component_evidence_bytes)?,
        )?;
        result.allocation_events = checked_add_u64(
            result.allocation_events,
            as_u64(component.allocation_events)?,
        )?;
        result.reference_edits =
            checked_add_u64(result.reference_edits, as_u64(component.reference_edits)?)?;
        result.reference_items =
            checked_add_u64(result.reference_items, as_u64(component.reference_items)?)?;
        result.appended_objects =
            checked_add_u64(result.appended_objects, as_u64(component.appended_objects)?)?;
        result.deleted_objects =
            checked_add_u64(result.deleted_objects, as_u64(component.deleted_objects)?)?;
    }
    result.retained_elements = as_u64(checked_add(
        components.len(),
        checked_add(evidence_messages, evidence_objects)?,
    )?)?;
    result.peak_scratch_bytes = as_u64(aggregate_peak_bytes)?;
    // The component result Vec, flattened message/object evidence Vecs, and
    // EntryEdit Vec are each one successful logical writer allocation when
    // nonempty. Empty evidence vectors allocate nothing.
    result.allocation_events = checked_add_u64(
        result.allocation_events,
        as_u64(checked_add(
            checked_add(2, pre_component_allocation_events)?,
            checked_add(
                usize::from(evidence_messages != 0),
                usize::from(evidence_objects != 0),
            )?,
        )?)?,
    )?;
    result.work = [
        result.components,
        result.compressed_input_bytes,
        result.decoded_input_bytes,
        result.serialized_output_bytes,
        result.compressed_output_bytes,
        result.retained_evidence_bytes,
        result.retained_elements,
        result.peak_scratch_bytes,
        result.allocation_events,
        result.reference_items,
        result.appended_objects,
        result.deleted_objects,
        as_u64(traversal_work)?,
    ]
    .into_iter()
    .try_fold(0u64, checked_add_u64)?;
    Ok(result)
}

#[cfg(test)]
fn take_publication_messages(
    components: &mut [RewrittenComponent<'_>],
) -> Result<Vec<PublishedMessage>, RewriteError> {
    let message_count = components.iter().try_fold(0usize, |total, component| {
        checked_add(total, component.published_messages.len())
    })?;
    let object_count = components.iter().try_fold(0usize, |total, component| {
        checked_add(total, component.published_objects.len())
    })?;
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(message_count)
        .map_err(|_error| RewriteError::Allocation {
            amount: message_count,
        })?;
    let mut objects = Vec::new();
    objects
        .try_reserve_exact(object_count)
        .map_err(|_error| RewriteError::Allocation {
            amount: object_count,
        })?;
    for component in components {
        messages.append(&mut component.published_messages);
        objects.append(&mut component.published_objects);
    }
    debug_assert!(messages.windows(2).all(|pair| {
        let left = &pair[0];
        let right = &pair[1];
        (left.component_index, message_evidence_key(left))
            < (right.component_index, message_evidence_key(right))
    }));
    debug_assert!(objects.windows(2).all(|pair| {
        let left = pair[0];
        let right = pair[1];
        (left.component_index, object_evidence_key(&left))
            < (right.component_index, object_evidence_key(&right))
    }));
    drop(objects);
    Ok(messages)
}

fn message_evidence_key(message: &PublishedMessage) -> (bool, usize, usize) {
    // Existing and deleted evidence is consumed by source-locality in source
    // order. Appended evidence has no source coordinate and follows it in
    // target order. This single order therefore remains strict after either
    // directional projection, even when a deleted source index numerically
    // equals an appended target index.
    match message.source_object_index {
        Some(source_object_index) => (false, source_object_index, message.message_index),
        None => (
            true,
            message.target_object_index.unwrap_or(usize::MAX),
            message.message_index,
        ),
    }
}

fn object_evidence_key(object: &PublishedObject) -> (bool, usize) {
    match object.source_object_index {
        Some(source_object_index) => (false, source_object_index),
        None => (true, object.target_object_index.unwrap_or(usize::MAX)),
    }
}

fn archive_extent(archive: &Archive) -> Result<usize, RewriteError> {
    archive.objects.iter().try_fold(0usize, |maximum, object| {
        let end = object
            .header_offset
            .checked_add(object.header_length)
            .and_then(|offset| offset.checked_add(object.data_length))
            .and_then(|value| usize::try_from(value).ok())
            .ok_or(RewriteError::InvalidSource)?;
        Ok(maximum.max(end))
    })
}

fn archive_allocation_cost(archive: &Archive) -> Result<usize, RewriteError> {
    use std::mem::size_of;

    archive.objects.iter().try_fold(
        checked_mul(archive.objects.len(), size_of::<ArchiveObject>())?,
        |total, object| checked_add(total, object_allocation_cost_checked(object)?),
    )
}

fn object_allocation_cost_checked(object: &ArchiveObject) -> Result<usize, RewriteError> {
    use std::mem::size_of;

    fn add(total: &mut usize, amount: usize) -> Result<(), RewriteError> {
        *total = checked_add(*total, amount)?;
        Ok(())
    }

    let mut total = checked_mul(object.messages.len(), size_of::<RawMessage>())?;
    total = checked_add(
        total,
        checked_mul(
            object.archive_info.message_infos.len(),
            size_of::<litchi_iwa_core::MessageInfo>(),
        )?,
    )?;
    total = checked_add(
        total,
        checked_mul(
            usize::try_from(object.header_length).map_err(|_error| RewriteError::InvalidSource)?,
            2,
        )?,
    )?;
    for info in &object.archive_info.message_infos {
        for length in [
            info.versions.len(),
            info.diff_merge_version.len(),
            info.diff_read_version.len(),
        ] {
            add(&mut total, checked_mul(length, size_of::<u32>())?)?;
        }
        if let Some(path) = &info.diff_field_path {
            add(&mut total, size_of::<litchi_iwa_core::FieldPath>())?;
            add(&mut total, checked_mul(path.path.len(), size_of::<u32>())?)?;
        }
        add(
            &mut total,
            checked_mul(
                info.fields_to_remove.len(),
                size_of::<litchi_iwa_core::FieldPath>(),
            )?,
        )?;
        add(
            &mut total,
            checked_mul(info.object_references.len(), size_of::<u64>())?,
        )?;
        add(
            &mut total,
            checked_mul(info.data_references.len(), size_of::<u64>())?,
        )?;
        add(
            &mut total,
            checked_mul(
                info.field_infos.len(),
                size_of::<litchi_iwa_core::FieldInfo>(),
            )?,
        )?;
        for path in &info.fields_to_remove {
            add(&mut total, checked_mul(path.path.len(), size_of::<u32>())?)?;
        }
        for field in &info.field_infos {
            let field_items = checked_add(
                checked_add(field.path.path.len(), field.object_references.len())?,
                field.data_references.len(),
            )?;
            add(&mut total, checked_mul(field_items, size_of::<u64>())?)?;
            add(
                &mut total,
                checked_mul(field.known_field_version.len(), size_of::<u32>())?,
            )?;
            add(
                &mut total,
                field
                    .known_field_feature_identifier
                    .as_ref()
                    .map_or(0, String::len),
            )?;
        }
    }
    Ok(total)
}

fn archive_reference_count(archive: &Archive) -> Result<usize, RewriteError> {
    let mut total = 0usize;
    for object in &archive.objects {
        for info in &object.archive_info.message_infos {
            total = checked_add(total, info.object_references.len())?;
            total = checked_add(total, info.data_references.len())?;
            for field in &info.field_infos {
                total = checked_add(total, field.object_references.len())?;
                total = checked_add(total, field.data_references.len())?;
            }
        }
    }
    Ok(total)
}

fn checked_add(left: usize, right: usize) -> Result<usize, RewriteError> {
    left.checked_add(right).ok_or(RewriteError::InvalidSource)
}

fn checked_mul(left: usize, right: usize) -> Result<usize, RewriteError> {
    left.checked_mul(right).ok_or(RewriteError::InvalidSource)
}

fn checked_add_u64(left: u64, right: u64) -> Result<u64, RewriteError> {
    left.checked_add(right).ok_or(RewriteError::InvalidSource)
}

fn as_u64(value: usize) -> Result<u64, RewriteError> {
    u64::try_from(value).map_err(|_error| RewriteError::InvalidSource)
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> RewriteError {
    match error {
        litchi_iwa_archive::Error::Limit { .. } => RewriteError::Limit,
        litchi_iwa_archive::Error::Allocation { amount, .. } => RewriteError::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => RewriteError::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => RewriteError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> RewriteError {
    match error {
        litchi_iwa_core::Error::Limit { .. } => RewriteError::Limit,
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            RewriteError::Allocation { amount: requested }
        },
        _ => RewriteError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{ArchiveObject, FieldInfo, RawMessage};

    use super::{
        AggregateReferenceDelta, ComponentReservation, FieldReferenceDelta, MessageEdit,
        MessageReplacement, PublicationReservation, PublishedMessage, PublishedMessageKind,
        ReferenceDelta, RewriteError, RewrittenComponent, apply_message_edit,
        classic_traversal_work, component_cost, counted_binary_search_by,
        locality_byte_envelope_from_lengths, message_evidence_key, replace_classic_message,
        shifted_object_index, validate_component_cost, with_component_authorization,
        with_publication_authorization,
    };

    fn referenced_object() -> ArchiveObject {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 7,
                data: vec![0xaa],
            }],
        )
        .expect("test object is valid");
        let info = &mut object.archive_info.message_infos[0];
        info.object_references.extend([10, 20]);
        let mut field = FieldInfo::new(vec![4, 1]);
        field.object_references.push(10);
        info.field_infos.push(field);
        object
    }

    #[test]
    fn exact_reference_delta_prunes_reorders_and_appends_atomically() {
        let mut object = referenced_object();
        let edit = MessageEdit {
            object_index: 0,
            message_index: 0,
            expected_type: 7,
            payload: vec![0xbb, 0xcc],
            references: Some(ReferenceDelta {
                aggregate_before: vec![10, 20],
                aggregate_after: vec![20, 30],
                fields: vec![FieldReferenceDelta {
                    field_info_index: 0,
                    expected_path: vec![4, 1],
                    before: vec![10],
                    after: vec![30],
                }],
            }),
        };
        apply_message_edit(&mut object, &edit, litchi_iwa_core::Limits::default())
            .expect("exact combined transition is valid");
        let info = &object.archive_info.message_infos[0];
        assert_eq!(info.object_references, [20, 30]);
        assert_eq!(info.field_infos[0].object_references, [30]);
        assert_eq!(object.messages[0].data, [0xbb, 0xcc]);
    }

    #[test]
    fn exact_reference_delta_prunes_all_proven_occurrences() {
        let mut object = referenced_object();
        let edit = MessageEdit {
            object_index: 0,
            message_index: 0,
            expected_type: 7,
            payload: vec![0xdd],
            references: Some(ReferenceDelta {
                aggregate_before: vec![10, 20],
                aggregate_after: vec![20],
                fields: vec![FieldReferenceDelta {
                    field_info_index: 0,
                    expected_path: vec![4, 1],
                    before: vec![10],
                    after: Vec::new(),
                }],
            }),
        };
        apply_message_edit(&mut object, &edit, litchi_iwa_core::Limits::default())
            .expect("exact prune is valid");
        let info = &object.archive_info.message_infos[0];
        assert_eq!(info.object_references, [20]);
        assert!(info.field_infos[0].object_references.is_empty());
        assert_eq!(object.messages[0].data, [0xdd]);
    }

    #[test]
    fn classic_weak_metadata_transition_rewrites_payload_and_exact_aggregate_only() {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2001,
                data: vec![0xaa],
            }],
        )
        .expect("test object is valid");
        let info = &mut object.archive_info.message_infos[0];
        info.object_references.extend([903_835, 903_815, 905_312]);
        info.data_references.push(77);
        assert!(info.field_infos.is_empty());

        replace_classic_message(
            &mut object,
            &MessageReplacement {
                component_index: 0,
                object_index: 0,
                message_index: 0,
                expected_type: 2001,
                payload: &[0xbb, 0xcc],
                references: Some(AggregateReferenceDelta {
                    before: &[903_835, 903_815, 905_312],
                    after: &[903_835, 903_815],
                }),
            },
            litchi_iwa_core::Limits::default(),
        )
        .expect("aggregate-only transition succeeds");

        assert_eq!(object.messages[0].data, [0xbb, 0xcc]);
        let info = &object.archive_info.message_infos[0];
        assert_eq!(info.object_references, [903_835, 903_815]);
        assert_eq!(info.data_references, [77]);
        assert!(info.field_infos.is_empty());
    }

    #[test]
    fn deletion_shift_is_deterministic() {
        let deletions = [
            super::ObjectDeletion {
                object_index: 1,
                expected_identifier: 11,
            },
            super::ObjectDeletion {
                object_index: 4,
                expected_identifier: 14,
            },
        ];
        assert_eq!(shifted_object_index(0, &deletions), Ok(0));
        assert_eq!(shifted_object_index(3, &deletions), Ok(2));
        assert_eq!(shifted_object_index(5, &deletions), Ok(3));
    }

    #[test]
    fn evidence_order_is_strict_in_both_directions() {
        let mut evidence = [
            PublishedMessage {
                component_index: 0,
                object_identifier: 30,
                source_object_index: None,
                target_object_index: Some(3),
                message_index: 0,
                expected_type: 7,
                kind: PublishedMessageKind::Appended,
                payload: vec![3],
            },
            PublishedMessage {
                component_index: 0,
                object_identifier: 20,
                source_object_index: Some(3),
                target_object_index: None,
                message_index: 0,
                expected_type: 7,
                kind: PublishedMessageKind::Deleted,
                payload: vec![2],
            },
            PublishedMessage {
                component_index: 0,
                object_identifier: 40,
                source_object_index: Some(2),
                target_object_index: Some(2),
                message_index: 0,
                expected_type: 7,
                kind: PublishedMessageKind::Existing,
                payload: vec![4],
            },
        ];
        evidence.sort_unstable_by_key(message_evidence_key);

        assert_eq!(
            evidence
                .iter()
                .filter_map(|message| message.source_object_index)
                .collect::<Vec<_>>(),
            [2, 3]
        );
        assert_eq!(
            evidence
                .iter()
                .filter_map(|message| message.target_object_index)
                .collect::<Vec<_>>(),
            [2, 3]
        );
        assert_eq!(
            evidence
                .iter()
                .map(|message| message.kind)
                .collect::<Vec<_>>(),
            [
                PublishedMessageKind::Existing,
                PublishedMessageKind::Deleted,
                PublishedMessageKind::Appended,
            ]
        );
    }

    #[test]
    fn max_minus_one_component_authorization_runs_no_component_work() {
        let reservation = ComponentReservation {
            work: 10,
            ..ComponentReservation::default()
        };
        let maximum = reservation.work - 1;
        let mut component_work = 0usize;
        let result = with_component_authorization(
            reservation,
            &mut |requested| {
                if requested.work > maximum {
                    Err(RewriteError::Precharge)
                } else {
                    Ok(())
                }
            },
            || {
                component_work += 1;
                Ok(())
            },
        );
        assert_eq!(result, Err(RewriteError::Precharge));
        assert_eq!(component_work, 0);
    }

    #[test]
    fn exact_component_cost_maps_within_reservation_and_rejects_max_minus_one() {
        let components = [RewrittenComponent {
            component_index: 0,
            name: "Index/Tables/Test.iwa",
            old_compressed_len: 3,
            source_decoded_bytes: 5,
            decoded_bytes: 7,
            allocation_bytes: 9,
            references: 6,
            compressed: vec![1, 2],
            published_messages: vec![PublishedMessage {
                component_index: 0,
                object_identifier: 7,
                source_object_index: Some(0),
                target_object_index: Some(0),
                message_index: 0,
                expected_type: 6_002,
                kind: PublishedMessageKind::Existing,
                payload: vec![3, 4, 5],
            }],
            published_objects: Vec::new(),
            peak_scratch_bytes: 11,
            allocation_events: 4,
            reference_edits: 1,
            reference_items: 6,
            appended_objects: 0,
            deleted_objects: 0,
        }];
        let exact = component_cost(&components, 0, 4, 0).expect("component accounting");
        assert_eq!(exact.retained_elements, 2);
        assert_eq!(exact.peak_scratch_bytes, 14);
        assert_eq!(exact.allocation_events, 7);
        assert_eq!(exact.reference_items, 10);
        let reservation = ComponentReservation {
            components: exact.components,
            compressed_input_bytes: exact.compressed_input_bytes,
            decoded_input_bytes: exact.decoded_input_bytes,
            maximum_serialized_output_bytes: exact.serialized_output_bytes,
            maximum_compressed_output_bytes: exact.compressed_output_bytes,
            maximum_retained_evidence_bytes: exact.retained_evidence_bytes,
            maximum_retained_elements: exact.retained_elements,
            maximum_peak_bytes: exact.peak_scratch_bytes,
            maximum_allocation_events: exact.allocation_events,
            reference_items: exact.reference_items,
            appended_objects: exact.appended_objects,
            deleted_objects: exact.deleted_objects,
            work: exact.work,
            ..ComponentReservation::default()
        };
        assert_eq!(validate_component_cost(exact, reservation), Ok(()));

        let too_small = ComponentReservation {
            maximum_allocation_events: reservation.maximum_allocation_events - 1,
            ..reservation
        };
        assert_eq!(
            validate_component_cost(exact, too_small),
            Err(RewriteError::Verification)
        );
    }

    #[test]
    fn four_to_eight_k_multi_component_traversal_is_linear_and_max_minus_one_is_early() {
        fn counter(components: usize) -> usize {
            let comparisons = (0..components)
                .map(|target| {
                    let (found, comparisons) =
                        counted_binary_search_by(components, |position| position.cmp(&target));
                    assert_eq!(found, Some(target));
                    comparisons
                })
                .sum();
            classic_traversal_work(components, comparisons, components, components)
                .expect("bounded traversal counter")
        }

        let four_k = counter(4_096);
        let eight_k = counter(8_192);
        assert!(eight_k * 100 <= four_k * 220, "4K={four_k}, 8K={eight_k}");

        let reservation = ComponentReservation {
            work: u64::try_from(eight_k).expect("counter fits u64"),
            ..ComponentReservation::default()
        };
        let maximum = reservation.work - 1;
        let mut component_work = 0usize;
        let result = with_component_authorization(
            reservation,
            &mut |requested| {
                if requested.work > maximum {
                    Err(RewriteError::Precharge)
                } else {
                    Ok(())
                }
            },
            || {
                component_work += 1;
                Ok(())
            },
        );
        assert_eq!(result, Err(RewriteError::Precharge));
        assert_eq!(component_work, 0);
    }

    #[test]
    fn locality_envelope_includes_decoded_components_and_max_minus_one_prevents_output() {
        let components = [RewrittenComponent {
            component_index: 0,
            name: "Index/Tables/Test.iwa",
            old_compressed_len: 3,
            source_decoded_bytes: 11,
            decoded_bytes: 13,
            allocation_bytes: 0,
            references: 0,
            compressed: vec![1, 2],
            published_messages: Vec::new(),
            published_objects: Vec::new(),
            peak_scratch_bytes: 0,
            allocation_events: 0,
            reference_edits: 0,
            reference_items: 0,
            appended_objects: 0,
            deleted_objects: 0,
        }];
        let locality_bytes = locality_byte_envelope_from_lengths(17, 19, &components)
            .expect("locality byte envelope");
        assert_eq!(locality_bytes, 60);

        let reservation = PublicationReservation {
            locality_bytes,
            ..PublicationReservation::default()
        };
        let maximum = locality_bytes - 1;
        let mut output_work = 0usize;
        let result = with_publication_authorization(
            reservation,
            super::ComponentCost::default(),
            &mut |requested, _cost| {
                if requested.locality_bytes > maximum {
                    Err(RewriteError::Precharge)
                } else {
                    Ok(())
                }
            },
            || {
                output_work += 1;
                Ok(())
            },
        );
        assert_eq!(result, Err(RewriteError::Precharge));
        assert_eq!(output_work, 0);
    }
}
