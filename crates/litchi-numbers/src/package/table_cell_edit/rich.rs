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
    RewriteBehavior, RewriteError, RewriteLimits, StorageRewrite, StorageValidation,
    rewrite_storage_text_with_behavior_and_limits, validate_storage_with_limits,
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
pub(super) struct ObjectSource<'a> {
    pub(super) location: MessageLocation,
    pub(super) identifier: u64,
    pub(super) message_type: u32,
    pub(super) payload: &'a [u8],
    /// Exact `MessageInfo.object_references`, in metadata order.
    pub(super) object_references: &'a [u64],
    pub(super) field_references: &'a [FieldReferences<'a>],
}

impl fmt::Debug for ObjectSource<'_> {
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
pub(super) struct Request<'a> {
    pub(super) route: ListRoute,
    pub(super) key: u32,
    pub(super) list_ref_count: u32,
    pub(super) payload: ObjectSource<'a>,
    pub(super) storage: ObjectSource<'a>,
    /// Exact inbound archive-header occurrences for the payload and storage.
    pub(super) payload_inbound_references: u32,
    pub(super) storage_inbound_references: u32,
    /// Strictly ascending, nonzero object identifiers in the package catalog.
    pub(super) local_object_ids: &'a [u64],
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
}

#[derive(Debug)]
pub(super) struct PlanParts {
    pub(super) disposition: Disposition,
    pub(super) result_key: u32,
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
    pub(super) const fn report(&self) -> Report {
        self.report
    }
    pub(super) fn retained_accounting(&self) -> Result<RetainedAccounting, Error> {
        let mut accounting = RetainedAccounting::default();
        account_vec(
            &mut accounting,
            self.replacements.len(),
            size_of::<MessageReplacement>(),
        )?;
        for replacement in &self.replacements {
            account_bytes(&mut accounting, replacement.payload.len())?;
            account_reference_delta(&mut accounting, &replacement.references)?;
        }
        Ok(accounting)
    }
    pub(super) fn into_parts(self) -> PlanParts {
        PlanParts {
            disposition: self.disposition,
            result_key: self.result_key,
            replacements: self.replacements,
        }
    }
}

/// Plan replacement of the complete rich storage text.
pub(super) fn plan_text(
    request: Request<'_>,
    replacement: &str,
    limits: Limits,
) -> Result<Plan, Error> {
    let payload_fields = validate_request(request, limits)?;
    let validation =
        validate_storage_with_limits(request.storage.payload, limits.text).map_err(map_text)?;
    let rewrite = rewrite_storage_text_with_behavior_and_limits(
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
    validate_storage_metadata(request, &rewrite, limits)?;
    let report = base_report(
        request,
        &rewrite,
        validation,
        payload_fields,
        replacement.len(),
    )?;
    if report.work_bound > limits.max_work {
        return Err(Error::Limit);
    }
    if !rewrite.changed() {
        return empty_plan(Disposition::Unchanged, request.key, report);
    }
    let references = reference_delta(
        request.storage.object_references,
        &rewrite,
        limits,
        request.storage.field_references.is_empty(),
    )?;
    let output_bytes = rewrite.bytes().len();
    let replacement = MessageReplacement {
        location: request.storage.location,
        expected_type: STORAGE_MESSAGE_TYPE,
        kind: ReplacementKind::StorageArchive,
        payload: rewrite.into_bytes(),
        references,
    };
    let mut replacements = Vec::new();
    reserve(&mut replacements, 1)?;
    replacements.push(replacement);
    Ok(Plan {
        disposition: Disposition::InPlace,
        result_key: request.key,
        replacements,
        report: Report {
            output_bytes,
            ..report
        },
    })
}

fn validate_request(request: Request<'_>, limits: Limits) -> Result<usize, Error> {
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
    source: ObjectSource<'_>,
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

fn validate_storage_metadata(
    request: Request<'_>,
    rewrite: &StorageRewrite,
    limits: Limits,
) -> Result<(), Error> {
    let aggregate_only = request.storage.field_references.is_empty();
    let references_match = if aggregate_only {
        reference_multiset_contains(
            rewrite.object_reference_occurrences_before(),
            request.storage.object_references,
        )?
    } else {
        same_reference_multiset(
            request.storage.object_references,
            rewrite.object_reference_occurrences_before(),
        )?
    };
    if !references_match {
        return Err(Error::InvalidSource);
    }
    if request.storage.object_references.len() > limits.max_deltas {
        return Err(Error::Limit);
    }
    for &identifier in request.storage.object_references {
        require_local(identifier, request.local_object_ids)?;
    }
    for &identifier in rewrite.object_reference_occurrences_before() {
        require_local(identifier, request.local_object_ids)?;
    }
    if aggregate_only {
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
        let found = request.storage.field_references.iter().any(|field| {
            field.root_field == removed.storage_field_number()
                && field.references.contains(&removed.identifier())
        });
        if !found {
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

fn base_report(
    request: Request<'_>,
    rewrite: &StorageRewrite,
    validation: StorageValidation,
    payload_fields: usize,
    replacement_bytes: usize,
) -> Result<Report, Error> {
    Ok(Report {
        input_bytes: aggregate_input_bytes(request)?,
        output_bytes: rewrite.bytes().len(),
        wire_fields: validation
            .fields()
            .checked_add(payload_fields)
            .ok_or(Error::Overflow)?,
        work_bound: aggregate_work_bound(
            request,
            validation,
            rewrite.bytes().len(),
            replacement_bytes,
            payload_fields,
        )?,
        reference_occurrences: rewrite.object_reference_occurrences_before().len(),
    })
}

fn aggregate_work_bound(
    request: Request<'_>,
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

fn aggregate_input_bytes(request: Request<'_>) -> Result<usize, Error> {
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
    for length in [delta.before.len(), delta.after.len(), delta.removed.len()] {
        account_vec(accounting, length, size_of::<u64>())?;
    }
    account_vec(
        accounting,
        delta.removed_by_field.len(),
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
    ) -> Request<'a> {
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
