//! Staged rich-text ownership for the bounded unique in-place path.
//!
//! This leaf never opens or reassembles a package. The transaction supplies
//! already-resolved raw messages and complete ownership counts; this module
//! validates those proofs, rewrites one `TSWP.StorageArchive`, and returns the
//! exact message and aggregate-reference delta for staged publication.

use core::{fmt, mem::size_of, ops::Range};

use litchi_iwa_common::{
    WireLimits,
    varint::{decode_varint_from_bytes, encoded_len},
    wire::{WireField, parse_wire_fields_with_limits},
};
use litchi_iwa_text_wire::{
    PreparedStorageRewrite, RewriteBehavior, RewriteError, RewriteLimits, StorageRewrite,
    StorageRewriteExecutionLimits, StorageValidation,
    prepare_storage_text_rewrite_with_behavior_and_limits, validate_storage_with_limits,
};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const PAYLOAD_MESSAGE_TYPE: u32 = 6_218;

/// Exact location used by staged component publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct MessageLocation {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
}

/// One archive-header field declaration for object references.
#[derive(Debug, Clone, Copy)]
pub(super) struct FieldReferences<'a> {
    pub(super) root_field: u32,
    pub(super) references: &'a [u64],
}

/// Raw one-message object admitted by the resolver.
#[derive(Clone, Copy)]
pub(super) struct ObjectSource<'source, 'fields, 'metadata> {
    pub(super) location: MessageLocation,
    pub(super) identifier: u64,
    pub(super) message_type: u32,
    pub(super) payload: &'source [u8],
    /// Exact `MessageInfo.object_references`, in metadata order.
    pub(super) object_references: &'metadata [u64],
    pub(super) field_references: &'fields [FieldReferences<'metadata>],
}

impl fmt::Debug for ObjectSource<'_, '_, '_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ObjectSource")
            .field("location", &self.location)
            .field("identifier", &self.identifier)
            .field("message_type", &self.message_type)
            .field("payload_bytes", &self.payload.len())
            .field("object_references", &self.object_references.len())
            .field("field_references", &self.field_references.len())
            .finish_non_exhaustive()
    }
}

/// The message that directly owns the selected list entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum EntryOwner {
    Root,
    Segment {
        object_id: u64,
        entries: u32,
        root_references: u32,
    },
}

/// Root and selected-entry routing for one rich-text list.
#[derive(Debug, Clone, Copy)]
pub(super) struct ListRoute {
    pub(super) root_object_id: u64,
    pub(super) owner: EntryOwner,
}

/// Complete, read-only input for one rich-text scalar transition.
#[derive(Debug, Clone, Copy)]
pub(super) struct Request<'source, 'fields, 'metadata> {
    pub(super) route: ListRoute,
    pub(super) key: u32,
    pub(super) list_ref_count: u32,
    pub(super) payload: ObjectSource<'source, 'fields, 'metadata>,
    pub(super) storage: ObjectSource<'source, 'fields, 'metadata>,
    /// Exact inbound archive-header occurrences for the payload and storage.
    pub(super) payload_inbound_references: u32,
    pub(super) storage_inbound_references: u32,
    /// Strictly ascending, nonzero object identifiers in the package catalog.
    pub(super) local_object_ids: &'metadata [u64],
}

/// Finite raw-wire and text-rewrite policy.
#[derive(Debug, Clone, Copy)]
pub(super) struct Limits {
    pub(super) wire: WireLimits,
    pub(super) text: RewriteLimits,
    pub(super) max_deltas: usize,
    pub(super) max_work: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            wire: WireLimits::default(),
            text: RewriteLimits::default(),
            max_deltas: 1_000_000,
            max_work: RewriteLimits::MAX_REWRITE_WORK,
        }
    }
}

/// Why a rich-text plan was refused. No source text, path, or identifier is
/// retained in the error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Error {
    InvalidSource,
    Limit,
    Allocation { amount: usize },
    Overflow,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid or unbounded Numbers rich-text mutation")
    }
}

impl std::error::Error for Error {}

/// Existing message replacement staged for publication.
#[derive(Clone, PartialEq, Eq)]
pub(super) struct MessageReplacement {
    pub(super) location: MessageLocation,
    pub(super) expected_type: u32,
    /// Explicit because an empty `StorageArchive` payload is canonical for an
    /// empty, default-kind cell and remains a publishable replacement.
    pub(super) kind: ReplacementKind,
    pub(super) payload: Vec<u8>,
    pub(super) references: ReferenceDelta,
}

impl fmt::Debug for MessageReplacement {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MessageReplacement")
            .field("location", &self.location)
            .field("expected_type", &self.expected_type)
            .field("kind", &self.kind)
            .field("payload_bytes", &self.payload.len())
            .field("references", &self.references)
            .finish_non_exhaustive()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ReplacementKind {
    StorageArchive,
}

/// Exact aggregate and field-specific archive-header reference transition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ReferenceDelta {
    pub(super) before: Vec<u64>,
    pub(super) after: Vec<u64>,
    /// Sorted, deduplicated identifiers safe to prune from aggregate and all
    /// field metadata after the leaf proved they no longer occur in payload.
    pub(super) removed: Vec<u64>,
    pub(super) removed_by_field: Vec<(u32, u64)>,
}

/// Exact bounded work retained by the leaf.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct Report {
    /// Aggregate source bytes admitted across payload validation and the
    /// storage validation/rewrite passes (storage is charged twice).
    pub(super) input_bytes: usize,
    /// Planner-retained replacement storage bytes.
    pub(super) output_bytes: usize,
    pub(super) wire_fields: usize,
    /// Conservative governed work bound for payload validation plus the text
    /// storage validation/rewrite passes. It is never an elapsed-time value.
    pub(super) work_bound: usize,
    pub(super) reference_occurrences: usize,
}

/// Exact retained governed units visible after planning. Temporary parser and
/// wire-rewrite scratch remains governed by [`Limits`] and is not retained.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct RetainedAccounting {
    pub(super) elements: usize,
    pub(super) bytes: usize,
    pub(super) allocation_events: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Disposition {
    Unchanged,
    InPlace,
}

/// Complete logical plan for the supported unique in-place path.
pub(super) struct Plan {
    disposition: Disposition,
    result_key: u32,
    replacements: Vec<MessageReplacement>,
    report: Report,
    execution: ExecutionReport,
}

/// Output-free rich-text plan. The source rewrite plan owns no candidate or
/// reference Vec; execution alone creates the staged replacement.
pub(super) struct PreparedPlan<'source, 'replacement> {
    metadata: PreparedMetadata,
    replacement: &'replacement str,
    disposition: Disposition,
    result_key: u32,
    storage: Option<PreparedStorageRewrite<'source, 'replacement>>,
    limits: Limits,
    report: Report,
    requirements: ExecutionRequirements,
}

struct PreparedMetadata {
    location: MessageLocation,
    aggregate_before: Vec<u64>,
    local_object_ids: Vec<u64>,
    field_references: Vec<(u32, u64)>,
    aggregate_only: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct ExecutionRequirements {
    pub(super) output_bytes: usize,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocation_events: usize,
    pub(super) work_bound: usize,
    pub(super) reference_occurrences: usize,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct ExecutionReport {
    pub(super) output_bytes: usize,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocation_events: usize,
    pub(super) work_bound: usize,
    pub(super) reference_occurrences: usize,
}

#[derive(Debug)]
pub(super) struct PlanParts {
    pub(super) disposition: Disposition,
    pub(super) replacements: Vec<MessageReplacement>,
}

impl fmt::Debug for Plan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Plan")
            .field("disposition", &self.disposition)
            .field("result_key", &self.result_key)
            .field("replacements", &self.replacements.len())
            .field("report", &self.report)
            .finish_non_exhaustive()
    }
}

impl Plan {
    #[cfg(test)]
    pub(super) const fn disposition(&self) -> Disposition {
        self.disposition
    }
    #[cfg(test)]
    pub(super) const fn result_key(&self) -> u32 {
        self.result_key
    }
    #[cfg(test)]
    pub(super) fn replacements(&self) -> &[MessageReplacement] {
        &self.replacements
    }
    pub(super) const fn execution_report(&self) -> ExecutionReport {
        self.execution
    }
    pub(super) fn into_parts(self) -> PlanParts {
        PlanParts {
            disposition: self.disposition,
            replacements: self.replacements,
        }
    }
}

impl PreparedPlan<'_, '_> {
    pub(super) const fn disposition(&self) -> Disposition {
        self.disposition
    }

    pub(super) const fn result_key(&self) -> u32 {
        self.result_key
    }

    pub(super) const fn prepare_report(&self) -> Report {
        Report {
            output_bytes: 0,
            ..self.report
        }
    }

    pub(super) const fn execution_requirements(&self) -> ExecutionRequirements {
        self.requirements
    }

    pub(super) fn retained_accounting(&self) -> Result<RetainedAccounting, Error> {
        let mut accounting = RetainedAccounting::default();
        account_vec(
            &mut accounting,
            self.metadata.aggregate_before.capacity(),
            size_of::<u64>(),
        )?;
        account_vec(
            &mut accounting,
            self.metadata.local_object_ids.capacity(),
            size_of::<u64>(),
        )?;
        account_vec(
            &mut accounting,
            self.metadata.field_references.capacity(),
            size_of::<(u32, u64)>(),
        )?;
        Ok(accounting)
    }

    pub(super) fn execute(self, limits: ExecutionRequirements) -> Result<Plan, Error> {
        preflight_execution(self.requirements, limits)?;
        let PreparedPlan {
            metadata,
            replacement,
            disposition,
            result_key,
            storage,
            limits: rich_limits,
            report,
            requirements,
        } = self;
        let Some(storage) = storage else {
            return empty_plan(disposition, result_key, report);
        };
        let text_requirements = storage.execution_requirements();
        let rewrite = storage
            .execute(StorageRewriteExecutionLimits {
                max_output_bytes: text_requirements.output_bytes(),
                max_retained_elements: text_requirements.retained_elements(),
                max_retained_bytes: text_requirements.retained_bytes(),
                max_peak_scratch_bytes: text_requirements.peak_scratch_bytes(),
                max_allocations: text_requirements.allocations(),
                max_work: text_requirements.work(),
            })
            .map_err(map_text)?;
        execute_rich_plan(
            metadata,
            result_key,
            replacement,
            rewrite,
            report,
            requirements,
            rich_limits,
        )
    }
}

/// Plan replacement of the complete rich storage text.
#[cfg(test)]
pub(super) fn plan_text(
    request: Request<'_, '_, '_>,
    replacement: &str,
    limits: Limits,
) -> Result<Plan, Error> {
    let prepared = prepare_text(request, replacement, limits)?;
    let requirements = prepared.execution_requirements();
    prepared.execute(requirements)
}

/// Plan replacement of complete rich storage text without candidate output.
pub(super) fn prepare_text<'source, 'replacement>(
    request: Request<'source, '_, '_>,
    replacement: &'replacement str,
    limits: Limits,
) -> Result<PreparedPlan<'source, 'replacement>, Error> {
    let payload_fields = validate_request(request, limits)?;
    let validation =
        validate_storage_with_limits(request.storage.payload, limits.text).map_err(map_text)?;
    let storage = prepare_storage_text_rewrite_with_behavior_and_limits(
        request.storage.payload,
        Range {
            start: 0,
            end: validation.utf16_len(),
        },
        replacement,
        RewriteBehavior::PreserveOnEqualText,
        limits.text,
    )
    .map_err(map_text)?;
    let storage_requirements = storage.execution_requirements();
    let report = planned_report(
        request,
        validation,
        payload_fields,
        replacement.len(),
        storage_requirements.output_bytes(),
        storage_requirements.reference_occurrences(),
    )?;
    if report.work_bound > limits.max_work {
        return Err(Error::Limit);
    }
    let disposition = if storage.changed() {
        Disposition::InPlace
    } else {
        Disposition::Unchanged
    };
    let metadata = prepare_metadata(request)?;
    let requirements = if disposition == Disposition::InPlace {
        rich_execution_requirements(request, &metadata, storage_requirements, report)?
    } else {
        ExecutionRequirements::default()
    };
    Ok(PreparedPlan {
        metadata,
        replacement,
        disposition,
        result_key: request.key,
        storage: (disposition == Disposition::InPlace).then_some(storage),
        limits,
        report,
        requirements,
    })
}

fn execute_rich_plan(
    metadata: PreparedMetadata,
    result_key: u32,
    _replacement: &str,
    rewrite: StorageRewrite,
    report: Report,
    requirements: ExecutionRequirements,
    limits: Limits,
) -> Result<Plan, Error> {
    validate_prepared_storage_metadata(&metadata, &rewrite, limits)?;
    let references = reference_delta(
        &metadata.aggregate_before,
        &rewrite,
        limits,
        metadata.aggregate_only,
    )?;
    let output_bytes = rewrite.bytes().len();
    let storage_execution = rewrite.execution_report();
    let replacement = MessageReplacement {
        location: metadata.location,
        expected_type: STORAGE_MESSAGE_TYPE,
        kind: ReplacementKind::StorageArchive,
        payload: rewrite.into_bytes(),
        references,
    };
    let mut replacements = Vec::new();
    reserve(&mut replacements, 1)?;
    replacements.push(replacement);
    let retained = retained_replacement_artifact(&replacements)?;
    let peak_scratch_bytes = storage_execution
        .peak_scratch_bytes
        .checked_add(retained.bytes)
        .and_then(|bytes| bytes.checked_add(size_of::<MessageReplacement>()))
        .ok_or(Error::Overflow)?;
    let allocation_events = storage_execution
        .allocations
        .checked_add(retained.allocation_events)
        .and_then(|events| events.checked_add(1))
        .ok_or(Error::Overflow)?;
    Ok(Plan {
        disposition: Disposition::InPlace,
        result_key,
        replacements,
        report: Report {
            output_bytes,
            ..report
        },
        execution: ExecutionReport {
            output_bytes,
            retained_elements: retained.elements,
            retained_bytes: retained.bytes,
            peak_scratch_bytes,
            allocation_events,
            work_bound: requirements.work_bound,
            reference_occurrences: requirements.reference_occurrences,
        },
    })
}

fn validate_request(request: Request<'_, '_, '_>, limits: Limits) -> Result<usize, Error> {
    if limits.max_deltas == 0
        || limits.max_work == 0
        || request.key == 0
        || request.list_ref_count != 1
        || request.payload.identifier == 0
        || request.storage.identifier == 0
        || request.route.root_object_id == 0
        || request.payload.message_type != PAYLOAD_MESSAGE_TYPE
        || request.storage.message_type != STORAGE_MESSAGE_TYPE
        || request.payload_inbound_references != 1
        || request.storage_inbound_references != 1
    {
        return Err(Error::InvalidSource);
    }
    validate_local_ids(request.local_object_ids)?;
    for id in [
        request.route.root_object_id,
        request.payload.identifier,
        request.storage.identifier,
    ] {
        require_local(id, request.local_object_ids)?;
    }
    match request.route.owner {
        EntryOwner::Root => {},
        EntryOwner::Segment {
            object_id,
            entries,
            root_references,
            ..
        } => {
            if entries == 0 || root_references != 1 {
                return Err(Error::InvalidSource);
            }
            require_local(object_id, request.local_object_ids)?;
        },
    }
    validate_payload(
        request.payload,
        request.storage.identifier,
        request.local_object_ids,
        limits.wire,
    )
}

fn validate_payload(
    source: ObjectSource<'_, '_, '_>,
    storage_id: u64,
    local: &[u64],
    limits: WireLimits,
) -> Result<usize, Error> {
    let fields = parse_wire_fields_with_limits(source.payload, limits).map_err(map_wire)?;
    let storage = singular(&fields, 1, 2, source.payload)?;
    let _cell = singular(&fields, 3, 2, source.payload)?;
    if fields.iter().filter(|field| field.number() == 2).count() > 1 {
        return Err(Error::InvalidSource);
    }
    let reference = storage.payload(source.payload).map_err(map_wire)?;
    let nested = parse_wire_fields_with_limits(reference, limits).map_err(map_wire)?;
    let identifier = singular(&nested, 1, 0, reference)?;
    let value = canonical_varint(identifier.payload(reference).map_err(map_wire)?)?;
    if value != storage_id {
        return Err(Error::InvalidSource);
    }
    if let Some(external) = optional_singular(&nested, 3, 0, reference)? {
        if canonical_varint(external.payload(reference).map_err(map_wire)?)? != 0 {
            return Err(Error::InvalidSource);
        }
    }
    if source.object_references != [storage_id] {
        return Err(Error::InvalidSource);
    }
    if !source.field_references.is_empty()
        && (source.field_references.len() != 1
            || source.field_references[0].root_field != 1
            || source.field_references[0].references != [storage_id])
    {
        return Err(Error::InvalidSource);
    }
    require_local(storage_id, local)?;
    fields
        .len()
        .checked_add(nested.len())
        .ok_or(Error::Overflow)
}

fn prepare_metadata(request: Request<'_, '_, '_>) -> Result<PreparedMetadata, Error> {
    let mut aggregate_before = Vec::new();
    reserve(
        &mut aggregate_before,
        request.storage.object_references.len(),
    )?;
    aggregate_before.extend_from_slice(request.storage.object_references);
    let mut local_object_ids = Vec::new();
    reserve(&mut local_object_ids, request.local_object_ids.len())?;
    local_object_ids.extend_from_slice(request.local_object_ids);
    let field_count =
        request
            .storage
            .field_references
            .iter()
            .try_fold(0usize, |count, field| {
                count
                    .checked_add(field.references.len())
                    .ok_or(Error::Overflow)
            })?;
    let mut field_references = Vec::new();
    reserve(&mut field_references, field_count)?;
    for field in request.storage.field_references {
        field_references.extend(
            field
                .references
                .iter()
                .map(|identifier| (field.root_field, *identifier)),
        );
    }
    field_references.sort_unstable();
    Ok(PreparedMetadata {
        location: request.storage.location,
        aggregate_before,
        local_object_ids,
        field_references,
        aggregate_only: request.storage.field_references.is_empty(),
    })
}

fn validate_prepared_storage_metadata(
    metadata: &PreparedMetadata,
    rewrite: &StorageRewrite,
    limits: Limits,
) -> Result<(), Error> {
    let references_match = if metadata.aggregate_only {
        reference_multiset_contains(
            rewrite.object_reference_occurrences_before(),
            &metadata.aggregate_before,
        )?
    } else {
        same_reference_multiset(
            &metadata.aggregate_before,
            rewrite.object_reference_occurrences_before(),
        )?
    };
    if !references_match {
        return Err(Error::InvalidSource);
    }
    if metadata.aggregate_before.len() > limits.max_deltas {
        return Err(Error::Limit);
    }
    for &identifier in &metadata.aggregate_before {
        require_local(identifier, &metadata.local_object_ids)?;
    }
    for &identifier in rewrite.object_reference_occurrences_before() {
        require_local(identifier, &metadata.local_object_ids)?;
    }
    if metadata.aggregate_only {
        for removed in rewrite.removed_object_references_by_field() {
            let before = rewrite
                .object_reference_occurrences_before()
                .iter()
                .filter(|identifier| **identifier == removed.identifier())
                .count();
            if before == 0 {
                return Err(Error::InvalidSource);
            }
        }
        for &removed in rewrite.removed_object_references() {
            if !rewrite
                .object_reference_occurrences_before()
                .contains(&removed)
                || rewrite
                    .object_reference_occurrences_after()
                    .contains(&removed)
            {
                return Err(Error::InvalidSource);
            }
        }
        return Ok(());
    }
    for removed in rewrite.removed_object_references_by_field() {
        if metadata
            .field_references
            .binary_search(&(removed.storage_field_number(), removed.identifier()))
            .is_err()
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn reference_delta(
    aggregate_before: &[u64],
    rewrite: &StorageRewrite,
    limits: Limits,
    aggregate_only: bool,
) -> Result<ReferenceDelta, Error> {
    let count = rewrite
        .object_reference_occurrences_before()
        .len()
        .checked_add(rewrite.object_reference_occurrences_after().len())
        .ok_or(Error::Overflow)?;
    if count > limits.max_deltas {
        return Err(Error::Limit);
    }
    let before = copy_slice(aggregate_before, "before references")?;
    let (after, removed) = if aggregate_only {
        aggregate_only_after(aggregate_before, rewrite)?
    } else {
        (
            copy_slice(
                rewrite.object_reference_occurrences_after(),
                "after references",
            )?,
            copy_slice(rewrite.removed_object_references(), "removed references")?,
        )
    };
    let mut removed_by_field = Vec::new();
    if !aggregate_only {
        reserve(
            &mut removed_by_field,
            rewrite.removed_object_references_by_field().len(),
        )?;
        for item in rewrite.removed_object_references_by_field() {
            removed_by_field.push((item.storage_field_number(), item.identifier()));
        }
    }
    Ok(ReferenceDelta {
        before,
        after,
        removed,
        removed_by_field,
    })
}

fn same_reference_multiset(left: &[u64], right: &[u64]) -> Result<bool, Error> {
    if left.len() != right.len() {
        return Ok(false);
    }
    let mut left_sorted = copy_slice(left, "left reference multiset")?;
    let mut right_sorted = copy_slice(right, "right reference multiset")?;
    left_sorted.sort_unstable();
    right_sorted.sort_unstable();
    Ok(left_sorted == right_sorted)
}

fn reference_multiset_contains(superset: &[u64], subset: &[u64]) -> Result<bool, Error> {
    let mut superset = copy_slice(superset, "reference superset")?;
    let mut subset = copy_slice(subset, "reference subset")?;
    superset.sort_unstable();
    subset.sort_unstable();
    let mut available = superset.into_iter().peekable();
    for requested in subset {
        while available
            .peek()
            .is_some_and(|candidate| *candidate < requested)
        {
            let _ = available.next();
        }
        if available.next() != Some(requested) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn aggregate_only_after(
    aggregate_before: &[u64],
    rewrite: &StorageRewrite,
) -> Result<(Vec<u64>, Vec<u64>), Error> {
    let mut sorted_after = copy_slice(
        rewrite.object_reference_occurrences_after(),
        "raw references after",
    )?;
    sorted_after.sort_unstable();
    let mut counts: Vec<(u64, usize)> = Vec::new();
    reserve(&mut counts, sorted_after.len())?;
    for identifier in sorted_after {
        if let Some((_current, count)) = counts
            .last_mut()
            .filter(|(current, _count)| *current == identifier)
        {
            *count = count.checked_add(1).ok_or(Error::Overflow)?;
        } else {
            counts.push((identifier, 1));
        }
    }
    let mut after = Vec::new();
    reserve(&mut after, aggregate_before.len())?;
    for &identifier in aggregate_before {
        let Some((_identifier, remaining)) = counts
            .binary_search_by_key(&identifier, |(candidate, _count)| *candidate)
            .ok()
            .and_then(|index| counts.get_mut(index))
        else {
            continue;
        };
        if *remaining != 0 {
            after.push(identifier);
            *remaining -= 1;
        }
    }
    let mut sorted_before = copy_slice(aggregate_before, "aggregate references before")?;
    sorted_before.sort_unstable();
    let mut removed = Vec::new();
    reserve(
        &mut removed,
        rewrite
            .removed_object_references()
            .len()
            .min(aggregate_before.len()),
    )?;
    for &identifier in rewrite.removed_object_references() {
        if sorted_before.binary_search(&identifier).is_ok() {
            removed.push(identifier);
        }
    }
    Ok((after, removed))
}

fn planned_report(
    request: Request<'_, '_, '_>,
    validation: StorageValidation,
    payload_fields: usize,
    replacement_bytes: usize,
    output_bytes: usize,
    reference_occurrences: usize,
) -> Result<Report, Error> {
    Ok(Report {
        input_bytes: aggregate_input_bytes(request)?,
        output_bytes,
        wire_fields: validation
            .fields()
            .checked_add(payload_fields)
            .ok_or(Error::Overflow)?,
        work_bound: aggregate_work_bound(
            request,
            validation,
            output_bytes,
            replacement_bytes,
            payload_fields,
        )?,
        reference_occurrences,
    })
}

fn rich_execution_requirements(
    request: Request<'_, '_, '_>,
    metadata: &PreparedMetadata,
    storage: litchi_iwa_text_wire::StorageRewriteExecutionRequirements,
    report: Report,
) -> Result<ExecutionRequirements, Error> {
    let occurrences = storage.reference_occurrences();
    let aggregate = request.storage.object_references.len();
    let after = aggregate.max(occurrences);
    let reference_elements = aggregate
        .checked_add(after)
        .and_then(|value| value.checked_add(occurrences.checked_mul(2)?))
        .ok_or(Error::Overflow)?;
    let retained_elements = storage
        .output_bytes()
        .checked_add(reference_elements)
        .ok_or(Error::Overflow)?;
    let scalar_reference_elements = aggregate
        .checked_add(after)
        .and_then(|value| value.checked_add(occurrences))
        .ok_or(Error::Overflow)?;
    let retained_bytes = storage
        .output_bytes()
        .checked_add(
            scalar_reference_elements
                .checked_mul(size_of::<u64>())
                .ok_or(Error::Overflow)?,
        )
        .and_then(|bytes| {
            occurrences
                .checked_mul(size_of::<(u32, u64)>())
                .and_then(|field_bytes| bytes.checked_add(field_bytes))
        })
        .ok_or(Error::Overflow)?;
    let reference_allocations = usize::from(aggregate != 0)
        .checked_add(usize::from(after != 0))
        .and_then(|value| value.checked_add(usize::from(occurrences != 0) * 2))
        .ok_or(Error::Overflow)?;
    let allocation_events = storage
        .allocations()
        .checked_add(reference_allocations)
        .and_then(|events| events.checked_add(1))
        .ok_or(Error::Overflow)?;
    let plan_bytes = metadata
        .aggregate_before
        .capacity()
        .checked_mul(size_of::<u64>())
        .and_then(|bytes| {
            metadata
                .local_object_ids
                .capacity()
                .checked_mul(size_of::<u64>())
                .and_then(|local| bytes.checked_add(local))
        })
        .and_then(|bytes| {
            metadata
                .field_references
                .capacity()
                .checked_mul(size_of::<(u32, u64)>())
                .and_then(|fields| bytes.checked_add(fields))
        })
        .ok_or(Error::Overflow)?;
    let peak_scratch_bytes = storage
        .peak_scratch_bytes()
        .checked_add(retained_bytes)
        .and_then(|bytes| bytes.checked_add(size_of::<MessageReplacement>()))
        .and_then(|bytes| bytes.checked_add(plan_bytes))
        .ok_or(Error::Overflow)?;
    let reference_work = aggregate
        .checked_add(occurrences.checked_mul(8).ok_or(Error::Overflow)?)
        .ok_or(Error::Overflow)?;
    Ok(ExecutionRequirements {
        output_bytes: storage.output_bytes(),
        retained_elements,
        retained_bytes,
        peak_scratch_bytes,
        allocation_events,
        work_bound: storage
            .work()
            .checked_add(reference_work)
            .ok_or(Error::Overflow)?,
        reference_occurrences: report.reference_occurrences,
    })
}

fn preflight_execution(
    requirements: ExecutionRequirements,
    limits: ExecutionRequirements,
) -> Result<(), Error> {
    if requirements.output_bytes > limits.output_bytes
        || requirements.retained_elements > limits.retained_elements
        || requirements.retained_bytes > limits.retained_bytes
        || requirements.peak_scratch_bytes > limits.peak_scratch_bytes
        || requirements.allocation_events > limits.allocation_events
        || requirements.work_bound > limits.work_bound
        || requirements.reference_occurrences > limits.reference_occurrences
    {
        return Err(Error::Limit);
    }
    Ok(())
}

fn retained_replacement_artifact(
    replacements: &[MessageReplacement],
) -> Result<RetainedAccounting, Error> {
    let mut accounting = RetainedAccounting::default();
    for replacement in replacements {
        account_bytes(&mut accounting, replacement.payload.capacity())?;
        account_reference_delta(&mut accounting, &replacement.references)?;
    }
    Ok(accounting)
}

fn aggregate_work_bound(
    request: Request<'_, '_, '_>,
    validation: StorageValidation,
    output_bytes: usize,
    replacement_bytes: usize,
    payload_fields: usize,
) -> Result<usize, Error> {
    // Mirrors the public wire edge's documented validation/rewrite envelopes
    // while conservatively bounding private tree bytes by the source and
    // result-text bytes by the complete output payload.
    let storage = request.storage.payload.len();
    let payload = request.payload.payload.len();
    let validation_work = storage
        .checked_mul(3)
        .and_then(|work| work.checked_add(validation.utf8_len().checked_mul(6)?))
        .ok_or(Error::Overflow)?;
    let rewrite_work = storage
        .checked_mul(5)
        .and_then(|work| work.checked_add(output_bytes.checked_mul(6)?))
        .and_then(|work| work.checked_add(validation.utf8_len().checked_mul(6)?))
        .and_then(|work| work.checked_add(replacement_bytes))
        .and_then(|work| work.checked_add(validation.fields()))
        .ok_or(Error::Overflow)?;
    payload
        .checked_mul(2)
        .and_then(|work| work.checked_add(payload_fields))
        .and_then(|work| work.checked_add(validation_work))
        .and_then(|work| work.checked_add(rewrite_work))
        .ok_or(Error::Overflow)
}

fn aggregate_input_bytes(request: Request<'_, '_, '_>) -> Result<usize, Error> {
    request
        .storage
        .payload
        .len()
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(request.payload.payload.len()))
        .ok_or(Error::Overflow)
}

fn empty_plan(disposition: Disposition, result_key: u32, report: Report) -> Result<Plan, Error> {
    Ok(Plan {
        disposition,
        result_key,
        replacements: Vec::new(),
        report,
        execution: ExecutionReport::default(),
    })
}

fn singular(
    fields: &[WireField],
    number: u32,
    wire: u8,
    source: &[u8],
) -> Result<WireField, Error> {
    let mut found = fields
        .iter()
        .copied()
        .filter(|field| field.number() == number);
    let field = found.next().ok_or(Error::InvalidSource)?;
    if found.next().is_some() || field.wire_type() != wire {
        return Err(Error::InvalidSource);
    }
    field.validate_canonical_framing(source).map_err(map_wire)?;
    Ok(field)
}

fn optional_singular(
    fields: &[WireField],
    number: u32,
    wire: u8,
    source: &[u8],
) -> Result<Option<WireField>, Error> {
    let mut found = fields
        .iter()
        .copied()
        .filter(|field| field.number() == number);
    let value = found.next();
    if found.next().is_some() {
        return Err(Error::InvalidSource);
    }
    if let Some(field) = value {
        if field.wire_type() != wire {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing(source).map_err(map_wire)?;
    }
    Ok(value)
}

fn canonical_varint(bytes: &[u8]) -> Result<u64, Error> {
    let (value, length) = decode_varint_from_bytes(bytes).map_err(|_| Error::InvalidSource)?;
    if length != bytes.len() || encoded_len(value) != length {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn validate_local_ids(ids: &[u64]) -> Result<(), Error> {
    if ids.first().copied() == Some(0) || ids.windows(2).any(|pair| pair[0] >= pair[1]) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}
fn require_local(id: u64, ids: &[u64]) -> Result<(), Error> {
    if id == 0 || ids.binary_search(&id).is_err() {
        Err(Error::InvalidSource)
    } else {
        Ok(())
    }
}
fn reserve<T>(output: &mut Vec<T>, amount: usize) -> Result<(), Error> {
    let expected = output.len().checked_add(amount).ok_or(Error::Overflow)?;
    output
        .try_reserve_exact(amount)
        .map_err(|_| Error::Allocation { amount })?;
    if size_of::<T>() != 0 && output.capacity() != expected {
        return Err(Error::Allocation { amount });
    }
    Ok(())
}

fn account_vec(
    accounting: &mut RetainedAccounting,
    elements: usize,
    element_bytes: usize,
) -> Result<(), Error> {
    accounting.elements = accounting
        .elements
        .checked_add(elements)
        .ok_or(Error::Overflow)?;
    if elements != 0 {
        accounting.allocation_events = accounting
            .allocation_events
            .checked_add(1)
            .ok_or(Error::Overflow)?;
    }
    accounting.bytes = accounting
        .bytes
        .checked_add(elements.checked_mul(element_bytes).ok_or(Error::Overflow)?)
        .ok_or(Error::Overflow)?;
    Ok(())
}

fn account_bytes(accounting: &mut RetainedAccounting, bytes: usize) -> Result<(), Error> {
    accounting.bytes = accounting.bytes.checked_add(bytes).ok_or(Error::Overflow)?;
    if bytes != 0 {
        accounting.allocation_events = accounting
            .allocation_events
            .checked_add(1)
            .ok_or(Error::Overflow)?;
    }
    Ok(())
}

fn account_reference_delta(
    accounting: &mut RetainedAccounting,
    delta: &ReferenceDelta,
) -> Result<(), Error> {
    for length in [
        delta.before.capacity(),
        delta.after.capacity(),
        delta.removed.capacity(),
    ] {
        account_vec(accounting, length, size_of::<u64>())?;
    }
    account_vec(
        accounting,
        delta.removed_by_field.capacity(),
        size_of::<(u32, u64)>(),
    )
}

fn copy_slice(source: &[u64], _resource: &'static str) -> Result<Vec<u64>, Error> {
    let mut result = Vec::new();
    reserve(&mut result, source.len())?;
    result.extend_from_slice(source);
    Ok(result)
}
fn map_wire(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_common::Error::LimitExceeded { .. }
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::Limit,
        _ => Error::InvalidSource,
    }
}
fn map_text(error: RewriteError) -> Error {
    match error {
        RewriteError::Allocation { amount, .. } => Error::Allocation { amount },
        RewriteError::LimitExceeded { .. } | RewriteError::InvalidLimit { .. } => Error::Limit,
        RewriteError::ArithmeticOverflow { .. } => Error::Overflow,
        _ => Error::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use litchi_iwa_protos::{tsp, tst, tswp};
    use prost::Message;

    use super::*;

    fn fixture_bytes(text: &str) -> (Vec<u8>, Vec<u8>) {
        let storage = tswp::StorageArchive {
            kind: Some(tswp::storage_archive::KindType::Cell as i32),
            text: vec![text.to_owned()],
            ..Default::default()
        }
        .encode_to_vec();
        let payload = tst::RichTextPayloadArchive {
            storage: tsp::Reference {
                identifier: 30,
                ..Default::default()
            },
            range: None,
            cellid: tst::CellId {
                packed_data: 1 << 16,
                expanded_coord: None,
            },
        }
        .encode_to_vec();
        (storage, payload)
    }

    fn request<'a>(
        storage: &'a [u8],
        payload: &'a [u8],
        local: &'a [u64],
        payload_refs: &'a [u64],
        payload_fields: &'a [FieldReferences<'a>],
        storage_refs: &'a [u64],
        storage_fields: &'a [FieldReferences<'a>],
        owner: EntryOwner,
        list_ref_count: u32,
    ) -> Request<'a, 'a, 'a> {
        Request {
            route: ListRoute {
                root_object_id: 10,
                owner,
            },
            key: 2,
            list_ref_count,
            payload: ObjectSource {
                location: MessageLocation {
                    component_index: 0,
                    object_index: 1,
                    message_index: 0,
                },
                identifier: 20,
                message_type: PAYLOAD_MESSAGE_TYPE,
                payload,
                object_references: payload_refs,
                field_references: payload_fields,
            },
            storage: ObjectSource {
                location: MessageLocation {
                    component_index: 1,
                    object_index: 0,
                    message_index: 0,
                },
                identifier: 30,
                message_type: STORAGE_MESSAGE_TYPE,
                payload: storage,
                object_references: storage_refs,
                field_references: storage_fields,
            },
            payload_inbound_references: 1,
            storage_inbound_references: 1,
            local_object_ids: local,
        }
    }

    #[test]
    fn unique_storage_replaces_in_place_and_equal_text_is_exact_noop() {
        let (storage, payload) = fixture_bytes("Original");
        let local = [10, 20, 30];
        let payload_refs = [30];
        let field_refs = [FieldReferences {
            root_field: 1,
            references: &payload_refs,
        }];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &field_refs,
            &[],
            &[],
            EntryOwner::Root,
            1,
        );

        let changed = plan_text(source, "Changed", Limits::default()).unwrap();
        assert_eq!(changed.disposition(), Disposition::InPlace);
        assert_eq!(changed.result_key(), 2);
        assert_eq!(changed.replacements().len(), 1);
        assert_eq!(
            changed.replacements()[0].kind,
            ReplacementKind::StorageArchive
        );
        let rewritten =
            tswp::StorageArchive::decode(changed.replacements()[0].payload.as_slice()).unwrap();
        assert_eq!(rewritten.text, ["Changed"]);

        let equal = plan_text(source, "Original", Limits::default()).unwrap();
        assert_eq!(equal.disposition(), Disposition::Unchanged);
        assert!(equal.replacements().is_empty());
    }

    #[test]
    fn shared_storage_is_refused_without_a_publishable_replacement() {
        let (storage, payload) = fixture_bytes("Original");
        let local = [10, 20, 30];
        let payload_refs = [30];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &[],
            &[],
            &[],
            EntryOwner::Root,
            2,
        );

        assert_eq!(
            plan_text(source, "Changed", Limits::default()).err(),
            Some(Error::InvalidSource)
        );
    }

    #[test]
    fn unique_segment_owner_routes_the_storage_replacement_in_place() {
        let (storage, payload) = fixture_bytes("Original");
        let local = [10, 11, 20, 30];
        let payload_refs = [30];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &[],
            &[],
            &[],
            EntryOwner::Segment {
                object_id: 11,
                entries: 1,
                root_references: 1,
            },
            1,
        );

        let plan = plan_text(source, "Changed", Limits::default()).unwrap();
        assert_eq!(plan.disposition(), Disposition::InPlace);
        assert_eq!(plan.result_key(), 2);
        assert_eq!(plan.replacements().len(), 1);
    }

    #[test]
    fn unique_storage_proves_and_preserves_local_style_ownership() {
        let (storage, payload) = fixture_bytes("Styled");
        let mut decoded = tswp::StorageArchive::decode(storage.as_slice()).unwrap();
        decoded.style_sheet = Some(tsp::Reference {
            identifier: 40,
            ..Default::default()
        });
        let storage = decoded.encode_to_vec();
        let local = [10, 20, 30, 40];
        let payload_refs = [30];
        let payload_fields = [FieldReferences {
            root_field: 1,
            references: &payload_refs,
        }];
        let style_refs = [40];
        let storage_fields = [FieldReferences {
            root_field: 2,
            references: &style_refs,
        }];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &payload_fields,
            &style_refs,
            &storage_fields,
            EntryOwner::Root,
            1,
        );
        let plan = plan_text(source, "Restyled text", Limits::default()).unwrap();
        let delta = &plan.replacements()[0].references;
        assert_eq!(delta.before, [40]);
        assert_eq!(delta.after, [40]);
        assert!(delta.removed_by_field.is_empty());
    }

    #[test]
    fn weak_native_metadata_uses_only_exact_aggregate_storage_ownership() {
        let (storage, payload) = fixture_bytes("Styled");
        let mut decoded = tswp::StorageArchive::decode(storage.as_slice()).unwrap();
        decoded.style_sheet = Some(tsp::Reference {
            identifier: 40,
            ..Default::default()
        });
        let storage = decoded.encode_to_vec();
        let local = [10, 20, 30, 40];
        let payload_refs = [30];
        let style_refs = [40];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &[],
            &style_refs,
            &[],
            EntryOwner::Root,
            1,
        );
        let plan = plan_text(source, "Changed", Limits::default()).unwrap();
        let delta = &plan.replacements()[0].references;
        assert_eq!(delta.before, [40]);
        assert_eq!(delta.after, [40]);
        assert!(delta.removed.is_empty());
        assert!(delta.removed_by_field.is_empty());
    }

    #[test]
    fn weak_native_metadata_rejects_unowned_aggregate_reference() {
        let (storage, payload) = fixture_bytes("Styled");
        let local = [10, 20, 30, 40];
        let payload_refs = [30];
        let hostile_refs = [40];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &[],
            &hostile_refs,
            &[],
            EntryOwner::Root,
            1,
        );
        assert_eq!(
            plan_text(source, "Changed", Limits::default()).err(),
            Some(Error::InvalidSource)
        );
    }

    #[test]
    fn weak_native_payload_rejects_aggregate_storage_disagreement() {
        let (storage, payload) = fixture_bytes("Styled");
        let local = [10, 20, 30, 31];
        let hostile_payload_refs = [31];
        let source = request(
            &storage,
            &payload,
            &local,
            &hostile_payload_refs,
            &[],
            &[],
            &[],
            EntryOwner::Root,
            1,
        );
        assert_eq!(
            plan_text(source, "Changed", Limits::default()).err(),
            Some(Error::InvalidSource)
        );
    }

    #[test]
    fn payload_rejects_additional_empty_field_metadata() {
        let (storage, payload) = fixture_bytes("Styled");
        let local = [10, 20, 30];
        let payload_refs = [30];
        let fields = [
            FieldReferences {
                root_field: 1,
                references: &payload_refs,
            },
            FieldReferences {
                root_field: 99,
                references: &[],
            },
        ];
        let source = request(
            &storage,
            &payload,
            &local,
            &payload_refs,
            &fields,
            &[],
            &[],
            EntryOwner::Root,
            1,
        );
        assert_eq!(
            plan_text(source, "Changed", Limits::default()).err(),
            Some(Error::InvalidSource)
        );
    }
}
