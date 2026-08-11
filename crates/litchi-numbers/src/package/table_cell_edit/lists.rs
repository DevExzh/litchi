//! Bounded, source-preserving `TST.TableDataList` string rewrites.
//!
//! This is deliberately a raw-payload layer.  The table-cell transaction owns
//! resource resolution and cell rewrites; this module owns exactly one string
//! list payload after those callers have grouped all changes for that list.
//! In particular, it never decodes a generated `TableDataList` or re-encodes
//! unrelated entries, segments, or unknown fields.

use core::{cmp::Ordering, fmt, mem::size_of};

use litchi_iwa_common::{
    WireLimits,
    varint::{encode_varint_into, encoded_len},
    wire::{WireField, parse_wire_fields_with_limits},
};
use litchi_iwa_protos::{
    numbers_table_cell_storage_codec::{
        DecodeError, DecodeOptions, DecodeReport, StorageVisitor, TableDataListEntrySnapshot,
        decode_table_data_list_entry_with_report, decode_table_data_list_with_visitor,
    },
    tst,
};

const ENTRY_FIELD: u32 = 3;
const NEXT_LIST_ID_FIELD: u32 = 2;
const ENTRY_KEY_FIELD: u32 = 1;
const ENTRY_REF_COUNT_FIELD: u32 = 2;
const ENTRY_STRING_FIELD: u32 = 3;

/// One requested string reference, in caller order.
///
/// Repeated text is intentional: all repeated requests receive the same key
/// and increment that entry's reference count once per request.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct StringRequest<'source> {
    text: &'source str,
}

impl fmt::Debug for StringRequest<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("StringRequest")
            .field("bytes", &self.text.len())
            .finish_non_exhaustive()
    }
}

impl<'source> StringRequest<'source> {
    /// Construct one borrowed text request.
    #[must_use]
    pub(crate) const fn new(text: &'source str) -> Self {
        Self { text }
    }
}

/// The source key selected for one [`StringRequest`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct StringAssignment {
    request: usize,
    key: u32,
}

impl StringAssignment {
    #[must_use]
    pub(crate) const fn request(self) -> usize {
        self.request
    }

    #[must_use]
    pub(crate) const fn key(self) -> u32 {
        self.key
    }
}

/// Exact observations from the read-only string-key assignment pass.
///
/// The formulas are intentionally exposed as scalar getters so the parent
/// transaction can translate them into its own resource envelope without
/// charging a list write or an output payload.  In particular:
///
/// - retained bytes are exactly `assignments * size_of::<StringAssignment>()`;
/// - peak scratch is the raw-field and entry-plan arrays plus the three
///   compact sort indexes;
/// - allocation events cover only this module's explicit fallible vectors;
/// - transaction work covers raw parsing, both strict decode passes, compact
///   indexing, and every request, but no encoding or output bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct AssignmentReport {
    decode: DecodeReport,
    entry_decode: DecodeTotals,
    entries_scanned: usize,
    requests: usize,
    unique_requests: usize,
    retained_bytes: usize,
    peak_scratch_bytes: usize,
    allocations: usize,
    transaction_work: usize,
}

impl AssignmentReport {
    #[must_use]
    pub(crate) const fn decode(self) -> DecodeReport {
        self.decode
    }

    #[must_use]
    pub(crate) const fn entry_decode(self) -> DecodeTotals {
        self.entry_decode
    }

    #[must_use]
    pub(crate) const fn entries_scanned(self) -> usize {
        self.entries_scanned
    }

    #[must_use]
    pub(crate) const fn requests(self) -> usize {
        self.requests
    }

    #[must_use]
    pub(crate) const fn unique_requests(self) -> usize {
        self.unique_requests
    }

    #[must_use]
    pub(crate) const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub(crate) const fn peak_scratch_bytes(self) -> usize {
        self.peak_scratch_bytes
    }

    #[must_use]
    pub(crate) const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub(crate) const fn transaction_work(self) -> usize {
        self.transaction_work
    }
}

/// Stable caller-ordered keys discovered without constructing a list payload.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct StringAssignments {
    assignments: Vec<StringAssignment>,
    report: AssignmentReport,
}

impl fmt::Debug for StringAssignments {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("StringAssignments")
            .field("assignments", &self.assignments.len())
            .field("report", &self.report)
            .finish_non_exhaustive()
    }
}

impl StringAssignments {
    #[must_use]
    pub(crate) fn assignments(&self) -> &[StringAssignment] {
        &self.assignments
    }

    #[must_use]
    pub(crate) const fn report(&self) -> AssignmentReport {
        self.report
    }
}

/// Finite policy for one grouped list plan.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ListLimits {
    decode: DecodeOptions,
    max_entries: usize,
    max_requests: usize,
    max_output_bytes: usize,
    max_retained_bytes: usize,
    max_peak_scratch_bytes: usize,
    max_allocations: usize,
    max_transaction_work: usize,
}

impl ListLimits {
    /// Construct explicit limits.  A transaction normally derives these from
    /// its global budget before allocating any list-local state.
    #[must_use]
    pub(crate) const fn new(
        decode: DecodeOptions,
        max_entries: usize,
        max_requests: usize,
        max_output_bytes: usize,
    ) -> Self {
        Self {
            decode,
            max_entries,
            max_requests,
            max_output_bytes,
            max_retained_bytes: usize::MAX,
            max_peak_scratch_bytes: usize::MAX,
            max_allocations: usize::MAX,
            max_transaction_work: usize::MAX,
        }
    }

    /// Bind list-local retained, scratch, allocation-event, and aggregate work
    /// ceilings to the parent transaction's remaining budget.
    #[must_use]
    pub(crate) const fn with_accounting(
        mut self,
        max_retained_bytes: usize,
        max_peak_scratch_bytes: usize,
        max_allocations: usize,
        max_transaction_work: usize,
    ) -> Self {
        self.max_retained_bytes = max_retained_bytes;
        self.max_peak_scratch_bytes = max_peak_scratch_bytes;
        self.max_allocations = max_allocations;
        self.max_transaction_work = max_transaction_work;
        self
    }
}

/// Exact list-local accounting returned only after the complete source has
/// passed strict decoding and the final payload has been constructed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ListReport {
    decode: DecodeReport,
    entry_decode: DecodeTotals,
    entries_scanned: usize,
    strings_reused: usize,
    strings_added: usize,
    input_bytes: usize,
    output_bytes: usize,
    retained_bytes: usize,
    peak_scratch_bytes: usize,
    allocations: usize,
    transaction_work: usize,
}

impl ListReport {
    #[must_use]
    pub(crate) const fn decode(self) -> DecodeReport {
        self.decode
    }

    /// Aggregate strict per-entry decode work performed after the complete
    /// root decoder accepted the list.  This is separate from `decode()` so a
    /// caller cannot accidentally treat the second, source-preserving pass as
    /// free.
    #[must_use]
    pub(crate) const fn entry_decode(self) -> DecodeTotals {
        self.entry_decode
    }

    #[must_use]
    pub(crate) const fn entries_scanned(self) -> usize {
        self.entries_scanned
    }

    #[must_use]
    pub(crate) const fn strings_reused(self) -> usize {
        self.strings_reused
    }

    #[must_use]
    pub(crate) const fn strings_added(self) -> usize {
        self.strings_added
    }

    #[must_use]
    pub(crate) const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub(crate) const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub(crate) const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub(crate) const fn peak_scratch_bytes(self) -> usize {
        self.peak_scratch_bytes
    }

    #[must_use]
    pub(crate) const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub(crate) const fn transaction_work(self) -> usize {
        self.transaction_work
    }
}

/// Aggregate decoder observations for all strict direct-entry passes.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct DecodeTotals {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    references: usize,
}

impl DecodeTotals {
    #[must_use]
    pub(crate) const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    #[must_use]
    pub(crate) const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub(crate) const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub(crate) const fn references(self) -> usize {
        self.references
    }

    fn add(&mut self, report: DecodeReport) -> Result<(), Failure> {
        self.source_bytes = self
            .source_bytes
            .checked_add(report.source_bytes())
            .ok_or(Failure::Overflow("entry decode source bytes"))?;
        self.fields = self
            .fields
            .checked_add(report.fields())
            .ok_or(Failure::Overflow("entry decode fields"))?;
        self.work_bytes = self
            .work_bytes
            .checked_add(report.work_bytes())
            .ok_or(Failure::Overflow("entry decode work bytes"))?;
        self.references = self
            .references
            .checked_add(report.references())
            .ok_or(Failure::Overflow("entry decode references"))?;
        Ok(())
    }
}

/// One fully prepared, single-payload rewrite.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct ListRewrite {
    payload: Vec<u8>,
    assignments: Vec<StringAssignment>,
    report: ListReport,
}

impl fmt::Debug for ListRewrite {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ListRewrite")
            .field("payload_bytes", &self.payload.len())
            .field("assignments", &self.assignments.len())
            .field("report", &self.report)
            .finish_non_exhaustive()
    }
}

impl ListRewrite {
    #[must_use]
    #[cfg(test)]
    fn payload(&self) -> &[u8] {
        &self.payload
    }

    #[must_use]
    pub(crate) fn assignments(&self) -> &[StringAssignment] {
        &self.assignments
    }

    #[must_use]
    pub(crate) const fn report(&self) -> ListReport {
        self.report
    }

    /// Consume the one final payload after the caller has completed its
    /// transaction-wide verification/precharge.
    #[must_use]
    pub(crate) fn into_payload(self) -> Vec<u8> {
        self.payload
    }
}

/// A failure before a list rewrite is published.
#[derive(Debug)]
pub(crate) enum Failure {
    Decode(DecodeError),
    Wire(litchi_iwa_common::Error),
    Allocation {
        resource: &'static str,
        amount: usize,
    },
    LimitExceeded {
        resource: &'static str,
        observed: usize,
        maximum: usize,
    },
    InvalidSource(&'static str),
    Overflow(&'static str),
}

impl fmt::Display for Failure {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Decode(error) => {
                write!(formatter, "strict table-data-list decode failed: {error}")
            },
            Self::Wire(error) => write!(formatter, "table-data-list wire scan failed: {error}"),
            Self::Allocation { resource, amount } => {
                write!(
                    formatter,
                    "could not allocate {amount} units for {resource}"
                )
            },
            Self::LimitExceeded {
                resource,
                observed,
                maximum,
            } => write!(
                formatter,
                "table-data-list {resource} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::InvalidSource(reason) => {
                write!(formatter, "invalid table-data-list source: {reason}")
            },
            Self::Overflow(reason) => {
                write!(formatter, "table-data-list arithmetic overflow: {reason}")
            },
        }
    }
}

impl std::error::Error for Failure {}

impl From<DecodeError> for Failure {
    fn from(error: DecodeError) -> Self {
        Self::Decode(error)
    }
}

impl From<litchi_iwa_common::Error> for Failure {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self::Wire(error)
    }
}

/// Determine the final stable key of every requested string without encoding.
///
/// This is the read-only first half of a grouped cell rewrite.  It exists for
/// callers that must place string keys into tile records before those tile
/// rewrites reveal the complete set of old keys to release.  The caller then
/// invokes [`plan_string_list`] exactly once with the final net releases and
/// the same additions.  Existing strings always retain their source key,
/// including a string that the final net plan releases to zero and revives in
/// the same batch; unique new strings receive consecutive `next_list_id` keys
/// in lexicographic request-text order, exactly as in [`plan_string_list`].
///
/// The source is strictly decoded and all allocations and arithmetic are
/// fallible.  No payload bytes are copied or encoded.
pub(crate) fn preflight_string_assignments(
    source_payload: &[u8],
    additions: &[StringRequest<'_>],
    limits: ListLimits,
) -> Result<StringAssignments, Failure> {
    if additions.len() > limits.max_requests {
        return Err(Failure::LimitExceeded {
            resource: "string requests",
            observed: additions.len(),
            maximum: limits.max_requests,
        });
    }

    let wire_limits = wire_limits(source_payload.len(), limits)?;
    let root_fields = parse_wire_fields_with_limits(source_payload, wire_limits)?;
    let entry_count = root_fields
        .iter()
        .filter(|field| field.number() == ENTRY_FIELD)
        .count();
    if entry_count > limits.max_entries {
        return Err(Failure::LimitExceeded {
            resource: "list entries",
            observed: entry_count,
            maximum: limits.max_entries,
        });
    }

    let mut staged = StagedEntries::new();
    let (list, decode) =
        decode_table_data_list_with_visitor(source_payload, limits.decode, &mut staged)?;
    if list.list_type() != tst::table_data_list::ListType::String as i32 {
        return Err(Failure::InvalidSource("TableDataList is not a string list"));
    }
    if staged.seen != entry_count {
        return Err(Failure::InvalidSource(
            "strict list entry stream differs from raw entry fields",
        ));
    }

    let (entries, entry_decode) =
        match_entries(source_payload, &root_fields, limits.decode, entry_count)?;
    let mut key_order = indices(entries.len(), "assignment key index")?;
    key_order.sort_unstable_by_key(|index| entries[*index].key);
    ensure_unique_keys(&entries, &key_order)?;

    let mut text_order = indices(entries.len(), "assignment string index")?;
    text_order.retain(|index| entries[*index].string.is_some());
    text_order.sort_unstable_by(|left, right| {
        entries[*left]
            .string
            .cmp(&entries[*right].string)
            .then_with(|| entries[*left].key.cmp(&entries[*right].key))
    });
    ensure_unique_texts(&entries, &text_order)?;

    let mut request_order = request_indices(additions.len())?;
    request_order.sort_unstable_by(|left, right| additions[*left].text.cmp(additions[*right].text));
    let mut assignments = blank_assignments(additions.len())?;
    let mut unique_requests = 0usize;
    let mut strings_new = 0usize;
    let mut request_position = 0usize;
    let mut next_key = list.next_list_id();
    while request_position < request_order.len() {
        let begin = request_position;
        let text = additions[request_order[begin]].text;
        while request_position < request_order.len()
            && additions[request_order[request_position]].text == text
        {
            request_position = request_position
                .checked_add(1)
                .ok_or(Failure::Overflow("assignment request position"))?;
        }
        unique_requests = unique_requests
            .checked_add(1)
            .ok_or(Failure::Overflow("unique string requests"))?;
        let key = if let Some(index) = find_text(&entries, &text_order, text) {
            entries
                .get(index)
                .ok_or(Failure::InvalidSource("assignment string index is invalid"))?
                .key
        } else {
            if next_key == 0 || find_key(&entries, &key_order, next_key).is_some() {
                return Err(Failure::InvalidSource("next string key is unavailable"));
            }
            strings_new = strings_new
                .checked_add(1)
                .ok_or(Failure::Overflow("preflight new strings"))?;
            let assigned = next_key;
            next_key = next_key
                .checked_add(1)
                .ok_or(Failure::Overflow("next string key"))?;
            assigned
        };
        for &request in &request_order[begin..request_position] {
            *assignments
                .get_mut(request)
                .ok_or(Failure::InvalidSource("request assignment is invalid"))? =
                StringAssignment { request, key };
        }
    }

    let final_entries = entry_count
        .checked_add(strings_new)
        .ok_or(Failure::Overflow("preflight final entries"))?;
    check_limit("planned list entries", final_entries, limits.max_entries)?;
    let retained_bytes = assignment_retained_bytes(assignments.len())?;
    let peak_scratch_bytes = scratch_bytes(
        &root_fields,
        &entries,
        &key_order,
        &text_order,
        &request_order,
    )?;
    let allocations = assignment_allocation_count(entry_count, additions.len())?;
    let transaction_work = assignment_transaction_work(
        source_payload.len(),
        root_fields.len(),
        decode,
        entry_decode,
        entry_count,
        additions.len(),
    )?;
    let report = AssignmentReport {
        decode,
        entry_decode,
        entries_scanned: entry_count,
        requests: additions.len(),
        unique_requests,
        retained_bytes,
        peak_scratch_bytes,
        allocations,
        transaction_work,
    };
    ensure_assignment_accounting(report, limits)?;
    Ok(StringAssignments {
        assignments,
        report,
    })
}

/// Plan and encode one string-list rewrite.
///
/// `releases` is `(key, references_to_remove)`.  All additions are grouped in
/// this call, so an existing string is incremented once by the count of all
/// requests that select it and a new string receives one new list key.  The
/// strict list decoder completes successfully *before* any planner result is
/// merged or returned.
///
/// Entries owned by segments are deliberately not materialized here.  Segment
/// resources remain raw references in this root payload; their individual
/// payloads must be passed to this same function as their own grouped resource.
pub(crate) fn plan_string_list(
    source_payload: &[u8],
    releases: &[(u32, u32)],
    additions: &[StringRequest<'_>],
    limits: ListLimits,
) -> Result<ListRewrite, Failure> {
    if additions.len() > limits.max_requests {
        return Err(Failure::LimitExceeded {
            resource: "string requests",
            observed: additions.len(),
            maximum: limits.max_requests,
        });
    }

    let wire_limits = wire_limits(source_payload.len(), limits)?;
    // This compact locator is deliberately built before the streaming codec
    // plan: it supplies exact raw entry spans and lets the visitor reserve once
    // rather than allocate while callbacks are in flight.
    let root_fields = parse_wire_fields_with_limits(source_payload, wire_limits)?;
    let entry_count = root_fields
        .iter()
        .filter(|field| field.number() == ENTRY_FIELD)
        .count();
    if entry_count > limits.max_entries {
        return Err(Failure::LimitExceeded {
            resource: "list entries",
            observed: entry_count,
            maximum: limits.max_entries,
        });
    }

    let mut staged = StagedEntries::new();
    let (list, decode) =
        decode_table_data_list_with_visitor(source_payload, limits.decode, &mut staged)?;
    if list.list_type() != tst::table_data_list::ListType::String as i32 {
        return Err(Failure::InvalidSource("TableDataList is not a string list"));
    }
    if staged.seen != entry_count {
        return Err(Failure::InvalidSource(
            "strict list entry stream differs from raw entry fields",
        ));
    }

    let (mut entries, entry_decode) =
        match_entries(source_payload, &root_fields, limits.decode, entry_count)?;
    let mut key_order = indices(entries.len(), "list key index")?;
    key_order.sort_unstable_by_key(|index| entries[*index].key);
    ensure_unique_keys(&entries, &key_order)?;

    for &(key, count) in releases {
        if count == 0 {
            continue;
        }
        let index = find_key(&entries, &key_order, key)
            .ok_or(Failure::InvalidSource("released string key is absent"))?;
        let entry = entries
            .get_mut(index)
            .ok_or(Failure::InvalidSource("string key index is invalid"))?;
        if entry.string.is_none() {
            return Err(Failure::InvalidSource("released key is not a string entry"));
        }
        entry.decrement(count)?;
    }

    let mut text_order = indices(entries.len(), "list string index")?;
    // Deduplicate against every source string, including one whose release
    // count is currently zero.  A later addition in this same batch revives
    // that key, which is the correct final-state COW result and avoids a
    // needless remove/new-key churn.
    text_order.retain(|index| entries[*index].string.is_some());
    text_order.sort_unstable_by(|left, right| {
        entries[*left]
            .string
            .cmp(&entries[*right].string)
            .then_with(|| entries[*left].key.cmp(&entries[*right].key))
    });
    ensure_unique_texts(&entries, &text_order)?;

    let mut request_order = request_indices(additions.len())?;
    request_order.sort_unstable_by(|left, right| additions[*left].text.cmp(additions[*right].text));
    let strings_to_add = count_new_strings(&entries, &text_order, additions, &request_order)?;
    let final_capacity = entry_count
        .checked_add(strings_to_add)
        .ok_or(Failure::Overflow("planned list entries"))?;
    check_limit("planned list entries", final_capacity, limits.max_entries)?;
    if strings_to_add != 0 {
        entries
            .try_reserve_exact(strings_to_add)
            .map_err(|_allocation| Failure::Allocation {
                resource: "new list entry plans",
                amount: strings_to_add,
            })?;
    }
    let mut assignments = blank_assignments(additions.len())?;
    let mut strings_reused = 0usize;
    let mut strings_added = 0usize;
    let mut request_position = 0usize;
    let mut next_key = list.next_list_id();
    while request_position < request_order.len() {
        let begin = request_position;
        let text = additions[request_order[begin]].text;
        while request_position < request_order.len()
            && additions[request_order[request_position]].text == text
        {
            request_position = request_position
                .checked_add(1)
                .ok_or(Failure::Overflow("request position"))?;
        }
        let occurrences = request_position
            .checked_sub(begin)
            .ok_or(Failure::Overflow("request group length"))?;
        let selected = find_text(&entries, &text_order, text);
        let key = match selected {
            Some(index) => {
                let entry = entries
                    .get_mut(index)
                    .ok_or(Failure::InvalidSource("string index is invalid"))?;
                entry.increment(occurrences)?;
                strings_reused = strings_reused
                    .checked_add(1)
                    .ok_or(Failure::Overflow("reused strings"))?;
                entry.key
            },
            None => {
                if next_key == 0 || find_key(&entries, &key_order, next_key).is_some() {
                    return Err(Failure::InvalidSource("next string key is unavailable"));
                }
                let raw =
                    encode_new_string_entry(next_key, occurrences, text, limits.max_output_bytes)?;
                entries.push(EntryPlan::new_string(next_key, occurrences, raw)?);
                strings_added = strings_added
                    .checked_add(1)
                    .ok_or(Failure::Overflow("added strings"))?;
                let assigned = next_key;
                next_key = next_key
                    .checked_add(1)
                    .ok_or(Failure::Overflow("next string key"))?;
                assigned
            },
        };
        for &request in &request_order[begin..request_position] {
            *assignments
                .get_mut(request)
                .ok_or(Failure::InvalidSource("request assignment is invalid"))? =
                StringAssignment { request, key };
        }
    }
    if strings_added != strings_to_add || entries.len() != final_capacity {
        return Err(Failure::InvalidSource(
            "string assignment changed during final planning",
        ));
    }

    let payload = encode_root(
        source_payload,
        &root_fields,
        &entries,
        list.next_list_id(),
        next_key,
        limits.max_output_bytes,
    )?;
    let entries_rewritten = entries.iter().filter(|entry| entry.changed()).count();
    let report = ListReport {
        decode,
        entry_decode,
        entries_scanned: entry_count,
        strings_reused,
        strings_added,
        input_bytes: source_payload.len(),
        output_bytes: payload.len(),
        retained_bytes: retained_bytes(&payload, &assignments)?,
        peak_scratch_bytes: scratch_bytes(
            &root_fields,
            &entries,
            &key_order,
            &text_order,
            &request_order,
        )?,
        allocations: allocation_count(&entries, additions.len(), strings_added, payload.len())?,
        transaction_work: transaction_work(
            source_payload.len(),
            root_fields.len(),
            decode,
            entry_decode,
            entry_count,
            entries_rewritten,
            payload.len(),
        )?,
    };
    ensure_accounting(report, limits)?;
    Ok(ListRewrite {
        payload,
        assignments,
        report,
    })
}

struct StagedEntries {
    seen: usize,
}

impl StagedEntries {
    const fn new() -> Self {
        Self { seen: 0 }
    }
}

impl StorageVisitor for StagedEntries {
    fn visit_list_entry(
        &mut self,
        entry: TableDataListEntrySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        let _ = entry;
        // The raw locator bounded `expected` before strict traversal.  The
        // post-decode equality check turns any codec/raw disagreement into our
        // typed source failure without manufacturing a private codec error.
        self.seen = self.seen.saturating_add(1);
        Ok(())
    }
}

struct EntryPlan<'source> {
    key: u32,
    initial_ref_count: u32,
    additions: u32,
    removals: u32,
    string: Option<&'source str>,
    raw: EntryRaw<'source>,
    original: bool,
}

impl<'source> EntryPlan<'source> {
    fn from_source(snapshot: TableDataListEntrySnapshot<'source>, raw: &'source [u8]) -> Self {
        Self {
            key: snapshot.key(),
            initial_ref_count: snapshot.ref_count(),
            additions: 0,
            removals: 0,
            string: snapshot.string_value(),
            raw: EntryRaw::Borrowed(raw),
            original: true,
        }
    }

    fn new_string(key: u32, ref_count: usize, raw: Vec<u8>) -> Result<Self, Failure> {
        let initial_ref_count = u32::try_from(ref_count)
            .map_err(|_conversion| Failure::Overflow("new string reference count"))?;
        Ok(Self {
            key,
            initial_ref_count,
            additions: 0,
            removals: 0,
            // Requests are sorted, so a newly allocated text cannot need to
            // participate in a later source-entry lookup in this call.
            string: None,
            raw: EntryRaw::Owned(raw),
            original: false,
        })
    }

    fn decrement(&mut self, count: u32) -> Result<(), Failure> {
        let final_count = self.final_ref_count()?;
        if count > final_count {
            return Err(Failure::InvalidSource("string reference count underflow"));
        }
        self.removals = self
            .removals
            .checked_add(count)
            .ok_or(Failure::Overflow("string reference removals"))?;
        Ok(())
    }

    fn increment(&mut self, count: usize) -> Result<(), Failure> {
        let count =
            u32::try_from(count).map_err(|_conversion| Failure::Overflow("string additions"))?;
        self.additions = self
            .additions
            .checked_add(count)
            .ok_or(Failure::Overflow("string reference additions"))?;
        self.final_ref_count()?;
        Ok(())
    }

    fn final_ref_count(&self) -> Result<u32, Failure> {
        self.initial_ref_count
            .checked_add(self.additions)
            .ok_or(Failure::Overflow("string reference count"))?
            .checked_sub(self.removals)
            .ok_or(Failure::InvalidSource("string reference count underflow"))
    }

    fn changed(&self) -> bool {
        // The grouped writer publishes the final net state.  Equal releases
        // and additions deliberately preserve the complete raw source entry,
        // including unknown fields and non-selected field order.
        !self.original || self.additions != self.removals
    }
}

enum EntryRaw<'source> {
    Borrowed(&'source [u8]),
    Owned(Vec<u8>),
}

impl EntryRaw<'_> {
    fn as_slice(&self) -> &[u8] {
        match self {
            Self::Borrowed(raw) => raw,
            Self::Owned(raw) => raw,
        }
    }
}

fn wire_limits(input_bytes: usize, limits: ListLimits) -> Result<WireLimits, Failure> {
    let input = input_bytes.max(1);
    let output = limits.max_output_bytes.max(1);
    WireLimits::default()
        .with_input_bytes(input)
        .map_err(Failure::Wire)?
        .with_output_bytes(output)
        .map_err(Failure::Wire)
}

fn match_entries<'source>(
    source: &'source [u8],
    fields: &[WireField],
    options: DecodeOptions,
    capacity: usize,
) -> Result<(Vec<EntryPlan<'source>>, DecodeTotals), Failure> {
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(capacity)
        .map_err(|_allocation| Failure::Allocation {
            resource: "list entry plans",
            amount: capacity,
        })?;
    let mut decoded = DecodeTotals::default();
    for field in fields {
        if field.number() != ENTRY_FIELD {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Failure::InvalidSource(
                "list entry has non-message wire type",
            ));
        }
        field.validate_canonical_framing(source)?;
        let raw = field.payload(source)?;
        let (snapshot, report) = decode_table_data_list_entry_with_report(raw, options)?;
        decoded.add(report)?;
        entries.push(EntryPlan::from_source(snapshot, raw));
    }
    Ok((entries, decoded))
}

fn retained_bytes(payload: &[u8], assignments: &[StringAssignment]) -> Result<usize, Failure> {
    payload
        .len()
        .checked_add(
            assignments
                .len()
                .checked_mul(size_of::<StringAssignment>())
                .ok_or(Failure::Overflow("retained assignments bytes"))?,
        )
        .ok_or(Failure::Overflow("retained list bytes"))
}

fn assignment_retained_bytes(assignments: usize) -> Result<usize, Failure> {
    assignments
        .checked_mul(size_of::<StringAssignment>())
        .ok_or(Failure::Overflow("retained assignment bytes"))
}

fn scratch_bytes(
    root_fields: &[WireField],
    entries: &[EntryPlan<'_>],
    key_order: &[usize],
    text_order: &[usize],
    request_order: &[usize],
) -> Result<usize, Failure> {
    let indexes = key_order
        .len()
        .checked_add(text_order.len())
        .and_then(|length| length.checked_add(request_order.len()))
        .and_then(|length| length.checked_mul(size_of::<usize>()))
        .ok_or(Failure::Overflow("scratch index bytes"))?;
    let raw_owned = entries.iter().try_fold(0usize, |total, entry| {
        let bytes = match &entry.raw {
            EntryRaw::Borrowed(_) => 0,
            EntryRaw::Owned(raw) => raw.len(),
        };
        total
            .checked_add(bytes)
            .ok_or(Failure::Overflow("scratch new-entry bytes"))
    })?;
    let root_bytes = root_fields
        .len()
        .checked_mul(size_of::<WireField>())
        .ok_or(Failure::Overflow("scratch root fields bytes"))?;
    let entry_bytes = entries
        .len()
        .checked_mul(size_of::<EntryPlan<'_>>())
        .ok_or(Failure::Overflow("scratch entry-plan bytes"))?;
    root_bytes
        .checked_add(entry_bytes)
        .and_then(|bytes| bytes.checked_add(indexes))
        .and_then(|bytes| bytes.checked_add(raw_owned))
        .ok_or(Failure::Overflow("peak list scratch bytes"))
}

fn allocation_count(
    entries: &[EntryPlan<'_>],
    request_count: usize,
    strings_added: usize,
    output_bytes: usize,
) -> Result<usize, Failure> {
    // Source-field parsing has its own fallible allocation path inside
    // `litchi-iwa-common`; this count is the exact set of explicit list-plan
    // allocations.  Zero-length reservations do not allocate.
    let entry_count = entries.iter().filter(|entry| entry.original).count();
    let changed_outputs = entries.iter().try_fold(0usize, |count, entry| {
        if entry.original && entry.changed() && entry.final_ref_count()? != 0 {
            count
                .checked_add(1)
                .ok_or(Failure::Overflow("changed entry output allocations"))
        } else {
            Ok(count)
        }
    })?;
    let fixed = usize::from(entry_count != 0) // entry plan
        + usize::from(strings_added != 0) // fallible exact entry-plan growth
        + usize::from(entry_count != 0) // key index
        + usize::from(entry_count != 0) // text index
        + usize::from(request_count != 0) // request order
        + usize::from(request_count != 0); // assignments
    fixed
        .checked_add(strings_added) // owned new-entry payloads
        .and_then(|count| count.checked_add(changed_outputs)) // changed entry output buffers
        .and_then(|count| count.checked_add(usize::from(output_bytes != 0))) // final payload
        .ok_or(Failure::Overflow("list allocation events"))
}

fn assignment_allocation_count(entry_count: usize, request_count: usize) -> Result<usize, Failure> {
    usize::from(entry_count != 0) // entry plan
        .checked_add(usize::from(entry_count != 0)) // key index
        .and_then(|count| count.checked_add(usize::from(entry_count != 0))) // text index
        .and_then(|count| count.checked_add(usize::from(request_count != 0))) // request order
        .and_then(|count| count.checked_add(usize::from(request_count != 0))) // assignments
        .ok_or(Failure::Overflow("assignment allocation events"))
}

fn transaction_work(
    source_bytes: usize,
    root_fields: usize,
    root_decode: DecodeReport,
    entry_decode: DecodeTotals,
    entries: usize,
    entries_rewritten: usize,
    output_bytes: usize,
) -> Result<usize, Failure> {
    // Every parser/decoder pass, sort comparison domain, field emission, and
    // retained output byte is charged.  The sort itself is bounded by the
    // compact indexes; its comparison upper bound is charged conservatively.
    let strict_root = root_decode
        .work_bytes()
        .checked_add(root_decode.fields())
        .and_then(|work| work.checked_add(root_decode.references()))
        .ok_or(Failure::Overflow("root strict work"))?;
    let strict_entries = entry_decode
        .work_bytes
        .checked_add(entry_decode.fields)
        .and_then(|work| work.checked_add(entry_decode.references))
        .ok_or(Failure::Overflow("entry strict work"))?;
    let parse = source_bytes
        .checked_add(root_fields)
        .ok_or(Failure::Overflow("root raw parse work"))?;
    let indexing = entries
        .checked_mul(6)
        .and_then(|work| work.checked_add(entries_rewritten.checked_mul(2)?))
        .ok_or(Failure::Overflow("list planning work"))?;
    parse
        .checked_add(strict_root)
        .and_then(|work| work.checked_add(strict_entries))
        .and_then(|work| work.checked_add(indexing))
        .and_then(|work| work.checked_add(output_bytes))
        .ok_or(Failure::Overflow("list transaction work"))
}

fn assignment_transaction_work(
    source_bytes: usize,
    root_fields: usize,
    root_decode: DecodeReport,
    entry_decode: DecodeTotals,
    entries: usize,
    requests: usize,
) -> Result<usize, Failure> {
    let strict_root = root_decode
        .work_bytes()
        .checked_add(root_decode.fields())
        .and_then(|work| work.checked_add(root_decode.references()))
        .ok_or(Failure::Overflow("assignment root strict work"))?;
    let strict_entries = entry_decode
        .work_bytes
        .checked_add(entry_decode.fields)
        .and_then(|work| work.checked_add(entry_decode.references))
        .ok_or(Failure::Overflow("assignment entry strict work"))?;
    let parse = source_bytes
        .checked_add(root_fields)
        .ok_or(Failure::Overflow("assignment raw parse work"))?;
    // Four charged compact operations per source entry (populate and sort the
    // key/text indexes) and per request (populate, sort/group, assign).  This
    // deterministic formula is deliberately independent of allocator and
    // comparison implementation details, so it scales linearly and provides
    // an exact transaction envelope for this leaf.
    let indexing = entries
        .checked_mul(4)
        .and_then(|work| work.checked_add(requests.checked_mul(4)?))
        .ok_or(Failure::Overflow("assignment indexing work"))?;
    parse
        .checked_add(strict_root)
        .and_then(|work| work.checked_add(strict_entries))
        .and_then(|work| work.checked_add(indexing))
        .ok_or(Failure::Overflow("assignment transaction work"))
}

fn ensure_accounting(report: ListReport, limits: ListLimits) -> Result<(), Failure> {
    check_limit(
        "retained list bytes",
        report.retained_bytes,
        limits.max_retained_bytes,
    )?;
    check_limit(
        "list peak scratch bytes",
        report.peak_scratch_bytes,
        limits.max_peak_scratch_bytes,
    )?;
    check_limit(
        "list allocation events",
        report.allocations,
        limits.max_allocations,
    )?;
    check_limit(
        "list transaction work",
        report.transaction_work,
        limits.max_transaction_work,
    )
}

fn ensure_assignment_accounting(
    report: AssignmentReport,
    limits: ListLimits,
) -> Result<(), Failure> {
    check_limit(
        "retained assignment bytes",
        report.retained_bytes,
        limits.max_retained_bytes,
    )?;
    check_limit(
        "assignment peak scratch bytes",
        report.peak_scratch_bytes,
        limits.max_peak_scratch_bytes,
    )?;
    check_limit(
        "assignment allocation events",
        report.allocations,
        limits.max_allocations,
    )?;
    check_limit(
        "assignment transaction work",
        report.transaction_work,
        limits.max_transaction_work,
    )
}

fn check_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<(), Failure> {
    if observed > maximum {
        return Err(Failure::LimitExceeded {
            resource,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn indices(length: usize, resource: &'static str) -> Result<Vec<usize>, Failure> {
    let mut indices = Vec::new();
    indices
        .try_reserve_exact(length)
        .map_err(|_allocation| Failure::Allocation {
            resource,
            amount: length,
        })?;
    for index in 0..length {
        indices.push(index);
    }
    Ok(indices)
}

fn request_indices(length: usize) -> Result<Vec<usize>, Failure> {
    indices(length, "string request index")
}

fn blank_assignments(length: usize) -> Result<Vec<StringAssignment>, Failure> {
    let mut assignments = Vec::new();
    assignments
        .try_reserve_exact(length)
        .map_err(|_allocation| Failure::Allocation {
            resource: "string assignments",
            amount: length,
        })?;
    for request in 0..length {
        assignments.push(StringAssignment { request, key: 0 });
    }
    Ok(assignments)
}

fn ensure_unique_keys(entries: &[EntryPlan<'_>], keys: &[usize]) -> Result<(), Failure> {
    for pair in keys.windows(2) {
        let left = entries
            .get(pair[0])
            .ok_or(Failure::InvalidSource("invalid key index"))?;
        let right = entries
            .get(pair[1])
            .ok_or(Failure::InvalidSource("invalid key index"))?;
        if left.key == right.key {
            return Err(Failure::InvalidSource("duplicate list key"));
        }
    }
    Ok(())
}

fn ensure_unique_texts(entries: &[EntryPlan<'_>], texts: &[usize]) -> Result<(), Failure> {
    for pair in texts.windows(2) {
        let left = entries
            .get(pair[0])
            .and_then(|entry| entry.string)
            .ok_or(Failure::InvalidSource("invalid string index"))?;
        let right = entries
            .get(pair[1])
            .and_then(|entry| entry.string)
            .ok_or(Failure::InvalidSource("invalid string index"))?;
        if left == right {
            return Err(Failure::InvalidSource("duplicate string alias"));
        }
    }
    Ok(())
}

fn count_new_strings(
    entries: &[EntryPlan<'_>],
    texts: &[usize],
    additions: &[StringRequest<'_>],
    requests: &[usize],
) -> Result<usize, Failure> {
    let mut count = 0usize;
    let mut position = 0usize;
    while position < requests.len() {
        let text = additions
            .get(requests[position])
            .ok_or(Failure::InvalidSource("request index is invalid"))?
            .text;
        position = position
            .checked_add(1)
            .ok_or(Failure::Overflow("request group position"))?;
        while position < requests.len()
            && additions
                .get(requests[position])
                .ok_or(Failure::InvalidSource("request index is invalid"))?
                .text
                == text
        {
            position = position
                .checked_add(1)
                .ok_or(Failure::Overflow("request group position"))?;
        }
        if find_text(entries, texts, text).is_none() {
            count = count
                .checked_add(1)
                .ok_or(Failure::Overflow("new string groups"))?;
        }
    }
    Ok(count)
}

fn find_key(entries: &[EntryPlan<'_>], keys: &[usize], key: u32) -> Option<usize> {
    keys.binary_search_by_key(&key, |index| entries[*index].key)
        .ok()
        .and_then(|slot| keys.get(slot).copied())
}

fn find_text(entries: &[EntryPlan<'_>], texts: &[usize], text: &str) -> Option<usize> {
    texts
        .binary_search_by(|index| match entries[*index].string {
            Some(value) => value.cmp(text),
            None => Ordering::Less,
        })
        .ok()
        .and_then(|slot| texts.get(slot).copied())
}

fn encode_new_string_entry(
    key: u32,
    count: usize,
    text: &str,
    maximum: usize,
) -> Result<Vec<u8>, Failure> {
    let count =
        u32::try_from(count).map_err(|_conversion| Failure::Overflow("new string count"))?;
    let text_length = text.len();
    let text_field_length = length_field_length(ENTRY_STRING_FIELD, text_length)?;
    let exact = varint_field_length(ENTRY_KEY_FIELD, u64::from(key))?
        .checked_add(varint_field_length(
            ENTRY_REF_COUNT_FIELD,
            u64::from(count),
        )?)
        .and_then(|length| length.checked_add(text_field_length))
        .ok_or(Failure::Overflow("new string entry bytes"))?;
    if exact > maximum {
        return Err(Failure::LimitExceeded {
            resource: "list output bytes",
            observed: exact,
            maximum,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(exact)
        .map_err(|_allocation| Failure::Allocation {
            resource: "new string entry",
            amount: exact,
        })?;
    append_varint(&mut output, ENTRY_KEY_FIELD, u64::from(key));
    append_varint(&mut output, ENTRY_REF_COUNT_FIELD, u64::from(count));
    append_length(&mut output, ENTRY_STRING_FIELD, text.as_bytes())?;
    if output.len() != exact {
        return Err(Failure::InvalidSource(
            "new string encoding length mismatch",
        ));
    }
    Ok(output)
}

fn encode_root(
    source: &[u8],
    fields: &[WireField],
    entries: &[EntryPlan<'_>],
    original_next_key: u32,
    next_key: u32,
    maximum: usize,
) -> Result<Vec<u8>, Failure> {
    let mut exact = source.len();
    let next_changed = next_key != original_next_key;
    let mut next_count = 0usize;
    for field in fields {
        if field.number() != NEXT_LIST_ID_FIELD {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(Failure::InvalidSource("invalid next string key field"));
        }
        next_count = next_count
            .checked_add(1)
            .ok_or(Failure::Overflow("next string key fields"))?;
        if next_changed {
            let replacement = varint_field_length(NEXT_LIST_ID_FIELD, u64::from(next_key))?;
            exact = exact
                .checked_sub(field.raw(source)?.len())
                .and_then(|length| length.checked_add(replacement))
                .ok_or(Failure::Overflow("list output bytes"))?;
        }
    }
    if next_count != 1 {
        return Err(Failure::InvalidSource(
            "missing or repeated next string key",
        ));
    }
    for entry in entries {
        if !entry.original {
            exact = exact
                .checked_add(length_field_length(
                    ENTRY_FIELD,
                    entry.raw.as_slice().len(),
                )?)
                .ok_or(Failure::Overflow("list output bytes"))?;
        } else if entry.changed() && entry.final_ref_count()? == 0 {
            let original = length_field_length(ENTRY_FIELD, entry.raw.as_slice().len())?;
            exact = exact
                .checked_sub(original)
                .ok_or(Failure::Overflow("list output bytes"))?;
        } else if entry.changed() {
            let original = length_field_length(ENTRY_FIELD, entry.raw.as_slice().len())?;
            let rewritten =
                length_field_length(ENTRY_FIELD, encoded_entry_length(entry, maximum)?)?;
            exact = exact
                .checked_sub(original)
                .and_then(|value| value.checked_add(rewritten))
                .ok_or(Failure::Overflow("list output bytes"))?;
        }
    }
    if exact > maximum {
        return Err(Failure::LimitExceeded {
            resource: "list output bytes",
            observed: exact,
            maximum,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(exact)
        .map_err(|_allocation| Failure::Allocation {
            resource: "rewritten string list",
            amount: exact,
        })?;
    let mut source_entry = 0usize;
    for field in fields {
        if field.number() == NEXT_LIST_ID_FIELD && next_changed {
            append_varint(&mut output, NEXT_LIST_ID_FIELD, u64::from(next_key));
            continue;
        }
        if field.number() != ENTRY_FIELD {
            output.extend_from_slice(field.raw(source)?);
            continue;
        }
        let entry = entries
            .get(source_entry)
            .ok_or(Failure::InvalidSource("raw list entry index is invalid"))?;
        source_entry = source_entry
            .checked_add(1)
            .ok_or(Failure::Overflow("raw list entry index"))?;
        if entry.changed() && entry.final_ref_count()? == 0 {
            continue;
        }
        if entry.changed() {
            let replacement = encoded_entry(entry, maximum)?;
            append_length(&mut output, ENTRY_FIELD, &replacement)?;
        } else {
            output.extend_from_slice(field.raw(source)?);
        }
    }
    if source_entry != entries.iter().filter(|entry| entry.original).count() {
        return Err(Failure::InvalidSource("raw list entry count changed"));
    }
    for entry in entries.iter().filter(|entry| !entry.original) {
        append_length(&mut output, ENTRY_FIELD, entry.raw.as_slice())?;
    }
    if output.len() != exact {
        return Err(Failure::InvalidSource("rewritten list length mismatch"));
    }
    Ok(output)
}

fn encoded_entry(entry: &EntryPlan<'_>, maximum: usize) -> Result<Vec<u8>, Failure> {
    if !entry.original {
        return Err(Failure::InvalidSource(
            "new entry does not need re-encoding",
        ));
    }
    let raw = entry.raw.as_slice();
    let fields = parse_wire_fields_with_limits(raw, WireLimits::default())?;
    let exact = encoded_entry_length_from_fields(entry, raw, &fields, maximum)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(exact)
        .map_err(|_allocation| Failure::Allocation {
            resource: "rewritten list entry",
            amount: exact,
        })?;
    for field in fields {
        if field.number() == ENTRY_REF_COUNT_FIELD {
            append_varint(
                &mut output,
                ENTRY_REF_COUNT_FIELD,
                u64::from(entry.final_ref_count()?),
            );
        } else {
            output.extend_from_slice(field.raw(raw)?);
        }
    }
    if output.len() != exact {
        return Err(Failure::InvalidSource(
            "rewritten list entry length mismatch",
        ));
    }
    Ok(output)
}

fn encoded_entry_length(entry: &EntryPlan<'_>, maximum: usize) -> Result<usize, Failure> {
    if !entry.original {
        return Err(Failure::InvalidSource(
            "new entry does not need re-encoding",
        ));
    }
    let raw = entry.raw.as_slice();
    let fields = parse_wire_fields_with_limits(raw, WireLimits::default())?;
    encoded_entry_length_from_fields(entry, raw, &fields, maximum)
}

fn encoded_entry_length_from_fields(
    entry: &EntryPlan<'_>,
    raw: &[u8],
    fields: &[WireField],
    maximum: usize,
) -> Result<usize, Failure> {
    let replacement_length =
        varint_field_length(ENTRY_REF_COUNT_FIELD, u64::from(entry.final_ref_count()?))?;
    let mut exact = raw.len();
    let mut seen = false;
    for field in fields {
        if field.number() == ENTRY_REF_COUNT_FIELD {
            if field.wire_type() != 0 || seen {
                return Err(Failure::InvalidSource(
                    "invalid string reference-count field",
                ));
            }
            seen = true;
            exact = exact
                .checked_sub(field.raw(raw)?.len())
                .and_then(|length| length.checked_add(replacement_length))
                .ok_or(Failure::Overflow("rewritten list entry bytes"))?;
        }
    }
    if !seen {
        return Err(Failure::InvalidSource(
            "missing string reference-count field",
        ));
    }
    if exact > maximum {
        return Err(Failure::LimitExceeded {
            resource: "list output bytes",
            observed: exact,
            maximum,
        });
    }
    Ok(exact)
}

fn varint_field_length(number: u32, value: u64) -> Result<usize, Failure> {
    encoded_len(u64::from(number) << 3)
        .checked_add(encoded_len(value))
        .ok_or(Failure::Overflow("varint field bytes"))
}

fn length_field_length(number: u32, length: usize) -> Result<usize, Failure> {
    let length =
        u64::try_from(length).map_err(|_conversion| Failure::Overflow("length field bytes"))?;
    encoded_len((u64::from(number) << 3) | 2)
        .checked_add(encoded_len(length))
        .and_then(|total| total.checked_add(usize::try_from(length).ok()?))
        .ok_or(Failure::Overflow("length field bytes"))
}

fn append_varint(output: &mut Vec<u8>, number: u32, value: u64) {
    encode_varint_into(output, u64::from(number) << 3);
    encode_varint_into(output, value);
}

fn append_length(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), Failure> {
    let length =
        u64::try_from(payload.len()).map_err(|_conversion| Failure::Overflow("length payload"))?;
    encode_varint_into(output, (u64::from(number) << 3) | 2);
    encode_varint_into(output, length);
    output.extend_from_slice(payload);
    Ok(())
}

#[cfg(test)]
mod tests {
    use core::mem::size_of;

    use prost::Message;

    use super::{
        Failure, ListLimits, StringAssignment, StringRequest, plan_string_list,
        preflight_string_assignments,
    };
    use litchi_iwa_protos::{numbers_table_cell_storage_codec::DecodeOptions, tst};

    fn string_list(next: u32, entries: &[(u32, u32, &str)]) -> Vec<u8> {
        tst::TableDataList {
            list_type: tst::table_data_list::ListType::String as i32,
            next_list_id: next,
            entries: entries
                .iter()
                .map(|&(key, refcount, text)| tst::table_data_list::ListEntry {
                    key,
                    refcount,
                    string: Some(text.to_owned()),
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn limits(payload: &[u8], max_entries: usize, max_requests: usize) -> ListLimits {
        let input = payload.len().max(1);
        ListLimits::new(
            DecodeOptions::new(
                input,
                input.saturating_mul(8),
                input.saturating_mul(128),
                64,
                usize::MAX,
                usize::MAX,
            ),
            max_entries,
            max_requests,
            input.saturating_add(16_384),
        )
    }

    fn requests<'a>(texts: &'a [&'a str]) -> Vec<StringRequest<'a>> {
        texts.iter().map(|text| StringRequest::new(text)).collect()
    }

    fn keys(assignments: &[StringAssignment]) -> Vec<u32> {
        assignments
            .iter()
            .map(|assignment| assignment.key())
            .collect()
    }

    #[test]
    fn preflight_assigns_unique_new_texts_once_in_stable_order() {
        let payload = string_list(10, &[(7, 1, "alpha")]);
        let additions = requests(&["zeta", "beta", "zeta"]);
        let result =
            preflight_string_assignments(&payload, &additions, limits(&payload, 3, 3)).unwrap();
        let rewrite = plan_string_list(&payload, &[], &additions, limits(&payload, 3, 3)).unwrap();

        assert_eq!(keys(result.assignments()), [11, 10, 11]);
        assert_eq!(keys(rewrite.assignments()), [11, 10, 11]);
        assert_eq!(
            (
                result.report().requests(),
                result.report().unique_requests()
            ),
            (3, 2)
        );
        assert_eq!(
            result.report().retained_bytes(),
            3 * size_of::<StringAssignment>()
        );
    }

    #[test]
    fn preflight_reuses_one_source_key_for_repeated_requests() {
        let payload = string_list(8, &[(7, 2, "alpha")]);
        let additions = requests(&["alpha", "alpha"]);
        let result =
            preflight_string_assignments(&payload, &additions, limits(&payload, 1, 2)).unwrap();

        assert_eq!(keys(result.assignments()), [7, 7]);
    }

    #[test]
    fn release_to_zero_and_add_same_text_nets_to_exact_source_entry() {
        let mut payload = string_list(8, &[(7, 1, "alpha")]);
        // Unknown root data proves that the net-zero path remains raw-source
        // authoritative rather than generated-message authoritative.
        super::append_varint(&mut payload, 100, 77);
        let additions = requests(&["alpha"]);
        let preflight =
            preflight_string_assignments(&payload, &additions, limits(&payload, 1, 1)).unwrap();
        let rewrite =
            plan_string_list(&payload, &[(7, 1)], &additions, limits(&payload, 1, 1)).unwrap();

        assert_eq!(keys(preflight.assignments()), [7]);
        assert_eq!(keys(rewrite.assignments()), [7]);
        assert_eq!(rewrite.payload(), payload);
    }

    #[test]
    fn assignment_envelope_scales_by_the_exposed_linear_formulas() {
        let payload = string_list(1, &[]);
        let small_text: Vec<String> = (0..16).map(|index| format!("s{index:02}")).collect();
        let large_text: Vec<String> = (0..32).map(|index| format!("s{index:02}")).collect();
        let small_refs: Vec<&str> = small_text.iter().map(String::as_str).collect();
        let large_refs: Vec<&str> = large_text.iter().map(String::as_str).collect();
        let small_requests = requests(&small_refs);
        let large_requests = requests(&large_refs);
        let small =
            preflight_string_assignments(&payload, &small_requests, limits(&payload, 16, 16))
                .unwrap()
                .report();
        let large =
            preflight_string_assignments(&payload, &large_requests, limits(&payload, 32, 32))
                .unwrap()
                .report();

        assert_eq!(
            large.retained_bytes() - small.retained_bytes(),
            16 * size_of::<StringAssignment>()
        );
        assert_eq!(
            large.peak_scratch_bytes() - small.peak_scratch_bytes(),
            16 * size_of::<usize>()
        );
        assert_eq!(large.transaction_work() - small.transaction_work(), 16 * 4);
        assert_eq!(large.allocations(), small.allocations());
    }

    #[test]
    fn assignment_exact_accounting_is_inclusive_and_max_minus_one_refuses() {
        let payload = string_list(9, &[(7, 1, "alpha")]);
        let additions = requests(&["alpha", "beta"]);
        let base = limits(&payload, 2, 2);
        let exact = preflight_string_assignments(&payload, &additions, base)
            .unwrap()
            .report();
        let exact_limits = base.with_accounting(
            exact.retained_bytes(),
            exact.peak_scratch_bytes(),
            exact.allocations(),
            exact.transaction_work(),
        );
        assert_eq!(
            preflight_string_assignments(&payload, &additions, exact_limits)
                .unwrap()
                .report(),
            exact
        );

        for constrained in [
            base.with_accounting(
                exact.retained_bytes() - 1,
                exact.peak_scratch_bytes(),
                exact.allocations(),
                exact.transaction_work(),
            ),
            base.with_accounting(
                exact.retained_bytes(),
                exact.peak_scratch_bytes() - 1,
                exact.allocations(),
                exact.transaction_work(),
            ),
            base.with_accounting(
                exact.retained_bytes(),
                exact.peak_scratch_bytes(),
                exact.allocations() - 1,
                exact.transaction_work(),
            ),
            base.with_accounting(
                exact.retained_bytes(),
                exact.peak_scratch_bytes(),
                exact.allocations(),
                exact.transaction_work() - 1,
            ),
        ] {
            assert!(matches!(
                preflight_string_assignments(&payload, &additions, constrained),
                Err(Failure::LimitExceeded { .. })
            ));
        }
    }
}
