//! Source-bound batch rewrites for comment-storage graph edges.
//!
//! A comment graph is represented by one `TSD.CommentStorageArchive` payload
//! per node.  The repeated `replies` references are the graph edges, while
//! `storage_uuid` is the node identity used by iWork readers.  This module
//! provides the small mutation needed by graph owners when they clone a node:
//! every selected reply identifier and, optionally, the storage UUID are
//! rewritten in one source pass.  The parent codec remains the wire and Buffa
//! authority; this module only adds planning and source-preserving emission.
//!
//! No generated message is retained or encoded here.  Unknown fields, nested
//! author/date payloads, and all untouched reply-reference spans are copied
//! byte-for-byte from the source.  Preparation borrows the source and the
//! caller's sorted remap table, and candidate bytes are allocated only after
//! the caller's execution limits have been checked.

use core::mem::size_of;

use super::*;

/// Source-bound mutation of all selected reply references and an optional
/// comment-storage UUID.
///
/// At least one source expectation is required: a complete payload fingerprint
/// and/or the source `storage_uuid`.  The expectation makes a prepared value
/// an optimistic-concurrency witness instead of an unguarded read-modify-write
/// operation.  `reply_remaps` is borrowed and must remain unchanged until
/// preparation completes; the prepared value retains that borrow through
/// execution. The table must be sorted strictly by source identifier, contain
/// only non-zero identifiers, and have unique replacement identifiers. It may
/// contain unrelated package remaps; only entries matching a reply in this
/// payload are applied.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CommentStorageLifecycleRewrite<'remaps> {
    reply_remaps: &'remaps [(u64, u64)],
    expected_fingerprint: Option<u64>,
    expected_storage_uuid: Option<UuidSnapshot>,
    replacement_storage_uuid: Option<UuidSnapshot>,
}

impl<'remaps> CommentStorageLifecycleRewrite<'remaps> {
    /// Build a rewrite using a sorted source-to-replacement table.
    #[must_use]
    pub const fn new(reply_remaps: &'remaps [(u64, u64)]) -> Self {
        Self {
            reply_remaps,
            expected_fingerprint: None,
            expected_storage_uuid: None,
            replacement_storage_uuid: None,
        }
    }

    /// Require the complete source payload to have `fingerprint`.
    #[must_use]
    pub const fn expecting_fingerprint(mut self, fingerprint: u64) -> Self {
        self.expected_fingerprint = Some(fingerprint);
        self
    }

    /// Require the source payload's `storage_uuid` to equal `uuid`.
    #[must_use]
    pub const fn expecting_storage_uuid(mut self, uuid: UuidSnapshot) -> Self {
        self.expected_storage_uuid = Some(uuid);
        self
    }

    /// Replace the source payload's `storage_uuid` with `uuid`.
    #[must_use]
    pub const fn replacing_storage_uuid(mut self, uuid: UuidSnapshot) -> Self {
        self.replacement_storage_uuid = Some(uuid);
        self
    }

    /// Return the borrowed reply remap table.
    #[must_use]
    pub const fn reply_remaps(self) -> &'remaps [(u64, u64)] {
        self.reply_remaps
    }

    /// Return the optional source fingerprint witness.
    #[must_use]
    pub const fn expected_fingerprint(self) -> Option<u64> {
        self.expected_fingerprint
    }

    /// Return the optional source UUID witness.
    #[must_use]
    pub const fn expected_storage_uuid(self) -> Option<UuidSnapshot> {
        self.expected_storage_uuid
    }

    /// Return the optional replacement UUID.
    #[must_use]
    pub const fn replacement_storage_uuid(self) -> Option<UuidSnapshot> {
        self.replacement_storage_uuid
    }
}

/// A prepared, source-witnessed batch comment graph rewrite.
///
/// Preparation performs strict source validation, remap validation, exact
/// output sizing, and finite-resource accounting.  `execute` performs one
/// source-to-output emission pass and one strict candidate readback pass; it
/// never performs a sequence of independent reply rewrites.
#[derive(Debug)]
pub struct PreparedCommentStorageLifecycleRewrite<'source, 'remaps> {
    source: &'source [u8],
    source_reply_ids: Vec<u64>,
    reply_remaps: &'remaps [(u64, u64)],
    replacement_storage_uuid: Option<UuidSnapshot>,
    source_uuid: Option<UuidSnapshot>,
    source_report: DecodeReport,
    candidate_report: DecodeReport,
    candidate_parity_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
}

impl<'source, 'remaps> PreparedCommentStorageLifecycleRewrite<'source, 'remaps> {
    /// Return strict source accounting from preparation.
    #[must_use]
    pub const fn prepare_report(&self) -> DecodeReport {
        self.source_report
    }

    /// Return all physical, wire, and scratch requirements.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute after the enclosing transaction has charged every requirement.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_rewrite_limits(self.requirements, limits)?;

        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_error| {
                DecodeError::resource(DecodeLimit::Allocations {
                    observed: self.requirements.allocations,
                    maximum: limits.allocations,
                })
            })?;
        emit_lifecycle_rewrite(
            &mut output,
            self.source,
            self.reply_remaps,
            self.replacement_storage_uuid,
            self.requirements.output_bytes,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }

        let candidate_options = DecodeOptions {
            max_message_bytes: output.len().max(1),
            ..DecodeOptions::new(
                output.len().max(1),
                self.requirements.fields.max(1),
                self.requirements.work_bytes.max(1),
                self.requirements.max_depth.max(1),
                self.requirements.references.max(1),
                self.source_report.text_bytes.max(1),
            )
        };
        let candidate = scan_comment_storage_raw(&output, candidate_options, None)?;
        let (candidate_snapshot, candidate_parity_report) =
            decode_comment_storage_archive_with_report(&output, candidate_options)?;
        if candidate.report != self.candidate_report
            || candidate_parity_report != self.candidate_parity_report
            || !reply_ids_match_batch(
                &candidate.reply_ids,
                &self.source_reply_ids,
                self.reply_remaps,
            )
            || candidate.storage_uuid_value != self.replacement_storage_uuid.or(self.source_uuid)
            || candidate_snapshot.storage_uuid() != candidate.storage_uuid_value
        {
            return Err(DecodeError::invalid());
        }
        let report = RewriteReport {
            source: self.source_report,
            result: candidate.report,
            input_bytes: self.source.len(),
            output_bytes: output.len(),
            fields: self.requirements.fields,
            work_bytes: self.requirements.work_bytes,
            max_depth: self.requirements.max_depth,
            references: self.requirements.references,
            replies: self.requirements.replies,
            reference_bytes: self.requirements.reference_bytes,
            allocations: self.requirements.allocations,
            scratch_bytes: self.requirements.scratch_bytes,
            retained_bytes: self.requirements.retained_bytes,
            changed: output.as_slice() != self.source,
        };
        Ok(RewriteOutput {
            bytes: output,
            report,
        })
    }
}

/// Prepare one atomic source-preserving comment graph rewrite.
pub fn prepare_comment_storage_lifecycle_rewrite<'source, 'remaps>(
    source: &'source [u8],
    rewrite: CommentStorageLifecycleRewrite<'remaps>,
    options: DecodeOptions,
) -> Result<PreparedCommentStorageLifecycleRewrite<'source, 'remaps>, DecodeError> {
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if rewrite.expected_fingerprint.is_none() && rewrite.expected_storage_uuid.is_none() {
        return Err(DecodeError::invalid());
    }
    let source_scan = scan_comment_storage_raw(source, options, None)?;
    let source_uuid = source_scan.storage_uuid_value;
    let fingerprint_work = usize::from(rewrite.expected_fingerprint.is_some())
        .checked_mul(source.len())
        .ok_or_else(DecodeError::invalid)?;
    let mut preflight_work = source_scan.report.work_bytes;
    preflight_work = add_work_with_limit(preflight_work, fingerprint_work, options)?;
    if rewrite
        .expected_fingerprint
        .is_some_and(|expected| expected != comment_storage_source_fingerprint(source))
        || rewrite
            .expected_storage_uuid
            .is_some_and(|expected| source_uuid != Some(expected))
    {
        return Err(DecodeError::invalid());
    }
    let parity_options =
        options_with_work_budget(options, remaining_work(options, preflight_work)?);
    let (_, source_parity_report) =
        decode_comment_storage_archive_with_report(source, parity_options)?;
    preflight_work = add_work_with_limit(preflight_work, source_parity_report.work_bytes, options)?;
    let validation_options =
        options_with_work_budget(options, remaining_work(options, preflight_work)?);
    let validation = validate_remap_tables(
        &source_scan.reply_ids,
        rewrite.reply_remaps,
        validation_options,
    )?;
    preflight_work = add_work_with_limit(preflight_work, validation.work_bytes, options)?;
    validate_storage_uuid_transition(
        source_uuid,
        source_scan.storage_uuid,
        rewrite.replacement_storage_uuid,
    )?;

    let measure_options =
        options_with_work_budget(options, remaining_work(options, preflight_work)?);
    let (output_bytes, candidate_report, delta_scratch) = measure_lifecycle_rewrite(
        source,
        &source_scan,
        rewrite.reply_remaps,
        rewrite.replacement_storage_uuid,
        measure_options,
    )?;
    let candidate_parity_report =
        project_candidate_report(source_parity_report, output_bytes, &delta_scratch, true)?;
    let planning_work = lifecycle_planning_work(
        source.len(),
        output_bytes,
        source_scan.reply_ids.len(),
        rewrite.reply_remaps.len(),
        rewrite.expected_fingerprint.is_some(),
    )?;
    let source_reply_scratch = source_scan
        .reply_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or_else(DecodeError::invalid)?;
    let remap_scratch = rewrite
        .reply_remaps
        .len()
        .checked_mul(size_of::<(u64, u64)>())
        .ok_or_else(DecodeError::invalid)?;
    let group_depth =
        usize::try_from(delta_scratch.max_group_depth).map_err(|_error| DecodeError::invalid())?;
    let group_scratch = group_depth
        .checked_mul(size_of::<u32>())
        .and_then(|bytes| bytes.checked_mul(2))
        .ok_or_else(DecodeError::invalid)?;
    let scratch_bytes = source_reply_scratch
        .checked_add(remap_scratch)
        .and_then(|bytes| bytes.checked_add(validation.scratch_bytes))
        .and_then(|bytes| bytes.checked_add(group_scratch))
        .ok_or_else(DecodeError::invalid)?;
    let allocations = usize::from(!source_scan.reply_ids.is_empty())
        .checked_add(validation.allocations)
        .and_then(|count| count.checked_add(usize::from(delta_scratch.max_group_depth != 0)))
        .and_then(|count| count.checked_add(1))
        .and_then(|count| count.checked_add(usize::from(candidate_report.replies() != 0)))
        .and_then(|count| count.checked_add(usize::from(delta_scratch.max_group_depth != 0)))
        .ok_or_else(DecodeError::invalid)?;
    let fields = source_scan
        .report
        .fields()
        .checked_add(source_parity_report.fields())
        .and_then(|fields| fields.checked_add(candidate_report.fields()))
        .and_then(|fields| fields.checked_add(candidate_parity_report.fields()))
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = source_scan
        .report
        .work_bytes()
        .checked_add(source_parity_report.work_bytes())
        .and_then(|work| work.checked_add(candidate_report.work_bytes()))
        .and_then(|work| work.checked_add(candidate_parity_report.work_bytes()))
        .and_then(|work| work.checked_add(validation.work_bytes))
        .and_then(|work| work.checked_add(planning_work))
        .ok_or_else(DecodeError::invalid)?;
    let references = source_scan
        .report
        .references()
        .checked_add(source_parity_report.references())
        .and_then(|references| references.checked_add(candidate_report.references()))
        .and_then(|references| references.checked_add(candidate_parity_report.references()))
        .ok_or_else(DecodeError::invalid)?;
    let replies = source_scan
        .report
        .replies()
        .checked_add(source_parity_report.replies())
        .and_then(|replies| replies.checked_add(candidate_report.replies()))
        .and_then(|replies| replies.checked_add(candidate_parity_report.replies()))
        .ok_or_else(DecodeError::invalid)?;
    let reference_bytes = source_scan
        .report
        .reference_bytes()
        .checked_add(source_parity_report.reference_bytes())
        .and_then(|bytes| bytes.checked_add(candidate_report.reference_bytes()))
        .and_then(|bytes| bytes.checked_add(candidate_parity_report.reference_bytes()))
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        max_depth: source_scan
            .report
            .max_depth()
            .max(source_parity_report.max_depth())
            .max(candidate_report.max_depth())
            .max(candidate_parity_report.max_depth()),
        references,
        replies,
        reference_bytes,
        allocations,
        scratch_bytes,
        retained_bytes: output_bytes,
    };
    ensure_rewrite_requirements_limits(requirements, options)?;
    Ok(PreparedCommentStorageLifecycleRewrite {
        source,
        source_reply_ids: source_scan.reply_ids,
        reply_remaps: rewrite.reply_remaps,
        replacement_storage_uuid: rewrite.replacement_storage_uuid,
        source_uuid,
        source_report: source_scan.report,
        candidate_report,
        candidate_parity_report,
        requirements,
    })
}

/// One-shot source-preserving comment graph rewrite.
pub fn rewrite_comment_storage_lifecycle(
    source: &[u8],
    rewrite: CommentStorageLifecycleRewrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_comment_storage_lifecycle_rewrite(source, rewrite, options)?;
    let limits = prepared.execution_requirements().exact();
    prepared.execute(limits)
}

fn options_with_work_budget(options: DecodeOptions, max_work_bytes: usize) -> DecodeOptions {
    DecodeOptions::new(
        options.max_message_bytes,
        options.max_fields,
        max_work_bytes,
        options.recursion_limit,
        options.max_references,
        options.max_text_bytes,
    )
}

fn remaining_work(options: DecodeOptions, consumed: usize) -> Result<usize, DecodeError> {
    options.max_work_bytes.checked_sub(consumed).ok_or_else(|| {
        DecodeError::resource(DecodeLimit::Work {
            observed: consumed,
            maximum: options.max_work_bytes,
        })
    })
}

fn add_work_with_limit(
    consumed: usize,
    additional: usize,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    let observed = consumed
        .checked_add(additional)
        .ok_or_else(DecodeError::invalid)?;
    if observed > options.max_work_bytes {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed,
            maximum: options.max_work_bytes,
        }));
    }
    Ok(observed)
}

fn project_candidate_report(
    source: DecodeReport,
    output_bytes: usize,
    delta: &LifecycleDelta,
    parity: bool,
) -> Result<DecodeReport, DecodeError> {
    let (work_added, work_removed) = if parity {
        let root_added = delta
            .output_added
            .checked_mul(2)
            .ok_or_else(DecodeError::invalid)?;
        let root_removed = delta
            .output_removed
            .checked_mul(2)
            .ok_or_else(DecodeError::invalid)?;
        let uuid_added = delta
            .uuid_work_added
            .checked_mul(2)
            .ok_or_else(DecodeError::invalid)?;
        let uuid_removed = delta
            .uuid_work_removed
            .checked_mul(2)
            .ok_or_else(DecodeError::invalid)?;
        let reply_added = delta
            .reply_work_added
            .checked_mul(2)
            .ok_or_else(DecodeError::invalid)?;
        let reply_removed = delta
            .reply_work_removed
            .checked_mul(2)
            .ok_or_else(DecodeError::invalid)?;
        (
            root_added
                .checked_add(reply_added)
                .and_then(|work| work.checked_add(uuid_added))
                .ok_or_else(DecodeError::invalid)?,
            root_removed
                .checked_add(reply_removed)
                .and_then(|work| work.checked_add(uuid_removed))
                .ok_or_else(DecodeError::invalid)?,
        )
    } else {
        (
            delta
                .output_added
                .checked_add(delta.work_added)
                .ok_or_else(DecodeError::invalid)?,
            delta
                .output_removed
                .checked_add(delta.work_removed)
                .ok_or_else(DecodeError::invalid)?,
        )
    };
    Ok(DecodeReport {
        source_bytes: output_bytes,
        fields: source.fields,
        work_bytes: apply_delta(source.work_bytes, work_added, work_removed)?,
        max_depth: source.max_depth,
        references: source.references,
        replies: source.replies,
        reference_bytes: apply_delta(
            source.reference_bytes,
            delta.reference_added,
            delta.reference_removed,
        )?,
        text_bytes: source.text_bytes,
    })
}

fn lifecycle_planning_work(
    source_bytes: usize,
    output_bytes: usize,
    reply_count: usize,
    remap_count: usize,
    fingerprints: bool,
) -> Result<usize, DecodeError> {
    let fingerprint = if fingerprints { source_bytes } else { 0 };
    // Emission scans the outer payload and may parse every selected nested
    // reference more than once (identifier lookup, replacement, and UUID
    // verification). Six source-length passes are a conservative byte-work
    // ceiling for those bounded loops.
    let emission = source_bytes
        .checked_mul(6)
        .ok_or_else(DecodeError::invalid)?;
    let remap_lookups = remap_lookup_work(reply_count, remap_count)?
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    fingerprint
        .checked_add(source_bytes)
        .and_then(|work| work.checked_add(emission))
        .and_then(|work| work.checked_add(output_bytes))
        .and_then(|work| work.checked_add(remap_lookups))
        .ok_or_else(DecodeError::invalid)
}

#[derive(Debug, Clone, Copy)]
struct RemapValidation {
    scratch_bytes: usize,
    allocations: usize,
    work_bytes: usize,
}

fn validate_remap_tables(
    source_ids: &[u64],
    remaps: &[(u64, u64)],
    options: DecodeOptions,
) -> Result<RemapValidation, DecodeError> {
    if remaps.len() > options.max_references {
        return Err(DecodeError::resource(DecodeLimit::References {
            observed: remaps.len(),
            maximum: options.max_references,
        }));
    }
    // Charge every validation loop, sort, and binary-search lookup before
    // allocating either sorted table. A low work ceiling must not become an
    // allocation probe for a large package-wide remap.
    let source_validation_work = source_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or_else(DecodeError::invalid)?;
    let remap_validation_work = remaps
        .len()
        .checked_mul(size_of::<(u64, u64)>())
        .ok_or_else(DecodeError::invalid)?;
    let source_sort_work = sort_work(source_ids.len())?;
    let replacement_sort_work = sort_work(remaps.len())?;
    let lookup_work = remap_lookup_work(source_ids.len(), remaps.len())?;
    let validation_work = source_validation_work
        .checked_add(remap_validation_work)
        .and_then(|work| work.checked_add(source_sort_work))
        .and_then(|work| work.checked_add(replacement_sort_work))
        .and_then(|work| work.checked_add(lookup_work))
        .ok_or_else(DecodeError::invalid)?;
    if validation_work > options.max_work_bytes {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: validation_work,
            maximum: options.max_work_bytes,
        }));
    }

    for (index, remap) in remaps.iter().enumerate() {
        let (source, replacement) = *remap;
        if source == 0 || replacement == 0 {
            return Err(DecodeError::invalid());
        }
        if index != 0
            && remaps
                .get(index.checked_sub(1).ok_or_else(DecodeError::invalid)?)
                .is_some_and(|previous| previous.0 >= source)
        {
            return Err(DecodeError::invalid());
        }
    }

    let mut sorted_source_ids = Vec::new();
    sorted_source_ids
        .try_reserve_exact(source_ids.len())
        .map_err(|_error| {
            DecodeError::resource(DecodeLimit::Allocations {
                observed: 1,
                maximum: 0,
            })
        })?;
    sorted_source_ids.extend_from_slice(source_ids);
    sorted_source_ids.sort_unstable();
    if sorted_source_ids
        .windows(2)
        .any(|window| window[0] == 0 || window[0] == window[1])
        || sorted_source_ids
            .last()
            .is_some_and(|identifier| *identifier == 0)
    {
        return Err(DecodeError::invalid());
    }

    let mut sorted_replacements = Vec::new();
    sorted_replacements
        .try_reserve_exact(remaps.len())
        .map_err(|_error| {
            DecodeError::resource(DecodeLimit::Allocations {
                observed: 1,
                maximum: 0,
            })
        })?;
    sorted_replacements.extend(remaps.iter().map(|(_, replacement)| *replacement));
    sorted_replacements.sort_unstable();
    if sorted_replacements
        .windows(2)
        .any(|window| window[0] == 0 || window[0] == window[1])
        || sorted_replacements
            .last()
            .is_some_and(|identifier| *identifier == 0)
    {
        return Err(DecodeError::invalid());
    }

    Ok(RemapValidation {
        scratch_bytes: source_ids
            .len()
            .checked_mul(size_of::<u64>())
            .and_then(|bytes| {
                remaps
                    .len()
                    .checked_mul(size_of::<u64>())
                    .and_then(|other| bytes.checked_add(other))
            })
            .ok_or_else(DecodeError::invalid)?,
        allocations: usize::from(!source_ids.is_empty())
            .checked_add(usize::from(!remaps.is_empty()))
            .ok_or_else(DecodeError::invalid)?,
        work_bytes: validation_work,
    })
}

fn sort_work(length: usize) -> Result<usize, DecodeError> {
    if length <= 1 {
        return Ok(0);
    }
    let leading_zeroes =
        usize::try_from((length - 1).leading_zeros()).map_err(|_error| DecodeError::invalid())?;
    let bits = usize::try_from(usize::BITS)
        .map_err(|_error| DecodeError::invalid())?
        .checked_sub(leading_zeroes)
        .ok_or_else(DecodeError::invalid)?;
    length
        .checked_mul(bits)
        .and_then(|comparisons| comparisons.checked_mul(size_of::<u64>()))
        .ok_or_else(DecodeError::invalid)
}

fn remap_lookup_work(reply_count: usize, remap_count: usize) -> Result<usize, DecodeError> {
    if reply_count == 0 || remap_count == 0 {
        return Ok(0);
    }
    let search_length = remap_count
        .checked_add(1)
        .ok_or_else(DecodeError::invalid)?;
    let leading_zeroes =
        usize::try_from(search_length.leading_zeros()).map_err(|_error| DecodeError::invalid())?;
    let bits = usize::try_from(usize::BITS)
        .map_err(|_error| DecodeError::invalid())?
        .checked_sub(leading_zeroes)
        .ok_or_else(DecodeError::invalid)?;
    reply_count
        .checked_mul(bits)
        .and_then(|searches| searches.checked_mul(3))
        .and_then(|searches| searches.checked_mul(size_of::<(u64, u64)>()))
        .ok_or_else(DecodeError::invalid)
}

fn validate_storage_uuid_transition(
    source_uuid: Option<UuidSnapshot>,
    source_field: Option<RawField>,
    replacement: Option<UuidSnapshot>,
) -> Result<(), DecodeError> {
    let Some(replacement) = replacement else {
        return Ok(());
    };
    if replacement.lower() == 0 && replacement.upper() == 0 {
        return Err(DecodeError::invalid());
    }
    if source_uuid.is_none() || source_field.is_none() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct LifecycleDelta {
    output_added: usize,
    output_removed: usize,
    work_added: usize,
    work_removed: usize,
    reply_work_added: usize,
    reply_work_removed: usize,
    uuid_work_added: usize,
    uuid_work_removed: usize,
    reference_added: usize,
    reference_removed: usize,
    max_group_depth: u32,
}

impl LifecycleDelta {
    const fn new() -> Self {
        Self {
            output_added: 0,
            output_removed: 0,
            work_added: 0,
            work_removed: 0,
            reply_work_added: 0,
            reply_work_removed: 0,
            uuid_work_added: 0,
            uuid_work_removed: 0,
            reference_added: 0,
            reference_removed: 0,
            max_group_depth: 0,
        }
    }

    fn add_change(&mut self, old: usize, new: usize, kind: DeltaKind) -> Result<(), DecodeError> {
        let (added, removed) = if new >= old {
            (new - old, 0)
        } else {
            (0, old - new)
        };
        let (add_slot, remove_slot) = match kind {
            DeltaKind::Output => (&mut self.output_added, &mut self.output_removed),
            DeltaKind::Work => (&mut self.work_added, &mut self.work_removed),
            DeltaKind::ReplyWork => (&mut self.reply_work_added, &mut self.reply_work_removed),
            DeltaKind::UuidWork => (&mut self.uuid_work_added, &mut self.uuid_work_removed),
            DeltaKind::Reference => (&mut self.reference_added, &mut self.reference_removed),
        };
        *add_slot = add_slot
            .checked_add(added)
            .ok_or_else(DecodeError::invalid)?;
        *remove_slot = remove_slot
            .checked_add(removed)
            .ok_or_else(DecodeError::invalid)?;
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum DeltaKind {
    Output,
    Work,
    ReplyWork,
    UuidWork,
    Reference,
}

fn apply_delta(base: usize, added: usize, removed: usize) -> Result<usize, DecodeError> {
    base.checked_add(added)
        .and_then(|value| value.checked_sub(removed))
        .ok_or_else(DecodeError::invalid)
}

fn measure_lifecycle_rewrite(
    source: &[u8],
    source_scan: &RawScanSummary,
    remaps: &[(u64, u64)],
    replacement_uuid: Option<UuidSnapshot>,
    options: DecodeOptions,
) -> Result<(usize, DecodeReport, LifecycleDelta), DecodeError> {
    let mut delta = LifecycleDelta::new();
    let mut parser_budget = RewriteScanBudget::new(source, options, options.max_message_bytes)?;
    let mut offset = 0;
    let mut reply_ordinal = 0usize;
    while let Some((field, group_depth)) =
        next_raw_field(source, &mut offset, source.len(), &mut parser_budget, 1)?
    {
        delta.max_group_depth = delta.max_group_depth.max(group_depth);
        if field.number == REPLIES_FIELD {
            let source_identifier = *source_scan
                .reply_ids
                .get(reply_ordinal)
                .ok_or_else(DecodeError::invalid)?;
            reply_ordinal = reply_ordinal
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
            if let Some(remap) = find_remap(remaps, source_identifier) {
                if remap.1 != source_identifier {
                    let old_payload_len = field.payload_len()?;
                    let old_identifier_len =
                        canonical_varint_field_len(REFERENCE_IDENTIFIER_FIELD, source_identifier)?;
                    let new_identifier_len =
                        canonical_varint_field_len(REFERENCE_IDENTIFIER_FIELD, remap.1)?;
                    let new_payload_len = old_payload_len
                        .checked_sub(old_identifier_len)
                        .and_then(|length| length.checked_add(new_identifier_len))
                        .ok_or_else(DecodeError::invalid)?;
                    let old_field_len = field.raw_len()?;
                    let new_field_len = canonical_bytes_field_len(REPLIES_FIELD, new_payload_len)?;
                    delta.add_change(old_field_len, new_field_len, DeltaKind::Output)?;
                    delta.add_change(old_payload_len, new_payload_len, DeltaKind::Work)?;
                    delta.add_change(old_payload_len, new_payload_len, DeltaKind::ReplyWork)?;
                    delta.add_change(old_payload_len, new_payload_len, DeltaKind::Reference)?;
                }
            }
        }
        if field.number == STORAGE_UUID_FIELD {
            if let Some(replacement_uuid) = replacement_uuid {
                let source_uuid = source_scan
                    .storage_uuid_value
                    .ok_or_else(DecodeError::invalid)?;
                if replacement_uuid != source_uuid {
                    let old_uuid_payload_len = field.payload_len()?;
                    let old_lower_len =
                        canonical_varint_field_len(UUID_LOWER_FIELD, source_uuid.lower())?;
                    let old_upper_len =
                        canonical_varint_field_len(UUID_UPPER_FIELD, source_uuid.upper())?;
                    let new_lower_len =
                        canonical_varint_field_len(UUID_LOWER_FIELD, replacement_uuid.lower())?;
                    let new_upper_len =
                        canonical_varint_field_len(UUID_UPPER_FIELD, replacement_uuid.upper())?;
                    let new_uuid_payload_len = old_uuid_payload_len
                        .checked_sub(old_lower_len)
                        .and_then(|length| length.checked_sub(old_upper_len))
                        .and_then(|length| length.checked_add(new_lower_len))
                        .and_then(|length| length.checked_add(new_upper_len))
                        .ok_or_else(DecodeError::invalid)?;
                    let old_field_len = field.raw_len()?;
                    let new_field_len =
                        canonical_bytes_field_len(STORAGE_UUID_FIELD, new_uuid_payload_len)?;
                    delta.add_change(old_field_len, new_field_len, DeltaKind::Output)?;
                    delta.add_change(
                        old_uuid_payload_len,
                        new_uuid_payload_len,
                        DeltaKind::Work,
                    )?;
                    delta.add_change(
                        old_uuid_payload_len,
                        new_uuid_payload_len,
                        DeltaKind::UuidWork,
                    )?;
                }
            }
        }
    }
    if reply_ordinal != source_scan.reply_ids.len() {
        return Err(DecodeError::invalid());
    }
    let output_bytes = apply_delta(source.len(), delta.output_added, delta.output_removed)?;
    let candidate_report = DecodeReport {
        source_bytes: output_bytes,
        fields: source_scan.report.fields,
        work_bytes: apply_delta(
            source_scan.report.work_bytes,
            delta
                .output_added
                .checked_add(delta.work_added)
                .ok_or_else(DecodeError::invalid)?,
            delta
                .output_removed
                .checked_add(delta.work_removed)
                .ok_or_else(DecodeError::invalid)?,
        )?,
        max_depth: source_scan.report.max_depth,
        references: source_scan.report.references,
        replies: source_scan.report.replies,
        reference_bytes: apply_delta(
            source_scan.report.reference_bytes,
            delta.reference_added,
            delta.reference_removed,
        )?,
        text_bytes: source_scan.report.text_bytes,
    };
    Ok((output_bytes, candidate_report, delta))
}

fn find_remap(remaps: &[(u64, u64)], source_identifier: u64) -> Option<(u64, u64)> {
    remaps
        .binary_search_by_key(&source_identifier, |remap| remap.0)
        .ok()
        .and_then(|index| remaps.get(index).copied())
}

fn emit_lifecycle_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    remaps: &[(u64, u64)],
    replacement_uuid: Option<UuidSnapshot>,
    expected_length: usize,
) -> Result<(), DecodeError> {
    emit_lifecycle_rewrite_fields(output, source, remaps, replacement_uuid)?;
    if output.len() == expected_length {
        Ok(())
    } else {
        Err(DecodeError::invalid())
    }
}

fn emit_lifecycle_rewrite_fields(
    output: &mut Vec<u8>,
    source: &[u8],
    remaps: &[(u64, u64)],
    replacement_uuid: Option<UuidSnapshot>,
) -> Result<(), DecodeError> {
    let options = DecodeOptions::new(
        source.len().max(1),
        usize::MAX,
        usize::MAX,
        MAX_RECURSION_LIMIT,
        usize::MAX,
        usize::MAX,
    );
    let mut parser_budget = RewriteScanBudget::new(source, options, source.len())?;
    let mut offset = 0;
    while let Some((field, _group_depth)) =
        next_raw_field(source, &mut offset, source.len(), &mut parser_budget, 1)?
    {
        if field.number == REPLIES_FIELD {
            let source_identifier = source_reply_identifier(source, field)?;
            let remap = find_remap(remaps, source_identifier);
            if let Some(remap) = remap {
                if remap.1 != source_identifier {
                    emit_replaced_reply_field(output, source, field, remap.1)?;
                } else {
                    output.extend_from_slice(&source[field.start..field.end]);
                }
            } else {
                output.extend_from_slice(&source[field.start..field.end]);
            }
        } else if field.number == STORAGE_UUID_FIELD {
            if let Some(replacement_uuid) = replacement_uuid {
                let raw = &source[field.payload_start..field.payload_end];
                let old_uuid = find_uuid_from_source(raw, options)?;
                if old_uuid == replacement_uuid {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    let new_payload_len =
                        uuid_replacement_payload_len(raw, old_uuid, replacement_uuid)?;
                    output.extend_from_slice(&source[field.start..field.value_start]);
                    append_varint(
                        output,
                        u64::try_from(new_payload_len).map_err(|_error| DecodeError::invalid())?,
                    );
                    emit_uuid_payload_replacement(output, raw, replacement_uuid)?;
                }
            } else {
                output.extend_from_slice(&source[field.start..field.end]);
            }
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    Ok(())
}

fn source_reply_identifier(source: &[u8], field: RawField) -> Result<u64, DecodeError> {
    let raw = &source[field.payload_start..field.payload_end];
    let options = DecodeOptions::new(
        raw.len().max(1),
        usize::MAX,
        usize::MAX,
        MAX_RECURSION_LIMIT,
        usize::MAX,
        usize::MAX,
    );
    let mut budget = RewriteScanBudget::new(raw, options, raw.len())?;
    let facts = validate_raw_reference(raw, &mut budget, 2)?;
    Ok(facts.identifier)
}

fn emit_replaced_reply_field(
    output: &mut Vec<u8>,
    source: &[u8],
    field: RawField,
    replacement_identifier: u64,
) -> Result<(), DecodeError> {
    let raw = &source[field.payload_start..field.payload_end];
    let replacement_len =
        canonical_varint_field_len(REFERENCE_IDENTIFIER_FIELD, replacement_identifier)?;
    let new_payload_len = raw
        .len()
        .checked_sub(canonical_varint_field_len(
            REFERENCE_IDENTIFIER_FIELD,
            source_reply_identifier(source, field)?,
        )?)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(DecodeError::invalid)?;
    output.extend_from_slice(&source[field.start..field.value_start]);
    append_varint(
        output,
        u64::try_from(new_payload_len).map_err(|_error| DecodeError::invalid())?,
    );
    emit_reference_identifier_replacement(output, raw, replacement_identifier)?;
    Ok(())
}

fn emit_reference_identifier_replacement(
    output: &mut Vec<u8>,
    source: &[u8],
    replacement_identifier: u64,
) -> Result<(), DecodeError> {
    let options = DecodeOptions::new(
        source.len().max(1),
        usize::MAX,
        usize::MAX,
        MAX_RECURSION_LIMIT,
        usize::MAX,
        usize::MAX,
    );
    let mut budget = RewriteScanBudget::new(source, options, source.len())?;
    let mut offset = 0;
    let mut replaced = false;
    while let Some((field, _group_depth)) =
        next_raw_field(source, &mut offset, source.len(), &mut budget, 2)?
    {
        if field.number == REFERENCE_IDENTIFIER_FIELD {
            if replaced {
                return Err(DecodeError::invalid());
            }
            output.extend_from_slice(&source[field.start..field.value_start]);
            append_varint(output, replacement_identifier);
            replaced = true;
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    if replaced {
        Ok(())
    } else {
        Err(DecodeError::invalid())
    }
}

fn find_uuid_from_source(
    source: &[u8],
    options: DecodeOptions,
) -> Result<UuidSnapshot, DecodeError> {
    let mut budget = RewriteScanBudget::new(source, options, source.len())?;
    validate_raw_uuid(source, &mut budget, 2)
}

fn uuid_replacement_payload_len(
    source: &[u8],
    old_uuid: UuidSnapshot,
    replacement_uuid: UuidSnapshot,
) -> Result<usize, DecodeError> {
    let old_lower = canonical_varint_field_len(UUID_LOWER_FIELD, old_uuid.lower())?;
    let old_upper = canonical_varint_field_len(UUID_UPPER_FIELD, old_uuid.upper())?;
    let new_lower = canonical_varint_field_len(UUID_LOWER_FIELD, replacement_uuid.lower())?;
    let new_upper = canonical_varint_field_len(UUID_UPPER_FIELD, replacement_uuid.upper())?;
    source
        .len()
        .checked_sub(old_lower)
        .and_then(|length| length.checked_sub(old_upper))
        .and_then(|length| length.checked_add(new_lower))
        .and_then(|length| length.checked_add(new_upper))
        .ok_or_else(DecodeError::invalid)
}

fn emit_uuid_payload_replacement(
    output: &mut Vec<u8>,
    source: &[u8],
    replacement_uuid: UuidSnapshot,
) -> Result<(), DecodeError> {
    let options = DecodeOptions::new(
        source.len().max(1),
        usize::MAX,
        usize::MAX,
        MAX_RECURSION_LIMIT,
        usize::MAX,
        usize::MAX,
    );
    let mut budget = RewriteScanBudget::new(source, options, source.len())?;
    let mut offset = 0;
    let mut lower = false;
    let mut upper = false;
    while let Some((field, _group_depth)) =
        next_raw_field(source, &mut offset, source.len(), &mut budget, 2)?
    {
        match field.number {
            UUID_LOWER_FIELD if !lower => {
                output.extend_from_slice(&source[field.start..field.value_start]);
                append_varint(output, replacement_uuid.lower());
                lower = true;
            },
            UUID_UPPER_FIELD if !upper => {
                output.extend_from_slice(&source[field.start..field.value_start]);
                append_varint(output, replacement_uuid.upper());
                upper = true;
            },
            UUID_LOWER_FIELD | UUID_UPPER_FIELD => {
                return Err(DecodeError::invalid());
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
    }
    if lower && upper {
        Ok(())
    } else {
        Err(DecodeError::invalid())
    }
}

fn reply_ids_match_batch(candidate_ids: &[u64], source_ids: &[u64], remaps: &[(u64, u64)]) -> bool {
    if candidate_ids.len() != source_ids.len() {
        return false;
    }
    candidate_ids.iter().enumerate().all(|(index, candidate)| {
        let source_identifier = source_ids[index];
        let expected =
            find_remap(remaps, source_identifier).map_or(source_identifier, |remap| remap.1);
        *candidate == expected
    })
}

#[cfg(test)]
#[path = "lifecycle_tests.rs"]
mod lifecycle_tests;
