//! Source-preserving cloning of one Keynote comment-storage payload.
//!
//! Comment storage is a graph node, rather than an opaque blob: its repeated
//! reply references point at the other nodes cloned with the selected
//! drawable.  The neutral Buffa lifecycle codec owns the wire validation and
//! source-preserving rewrite.  This adapter supplies the Keynote transaction
//! budget and deliberately leaves the author, text, date, UUID, and unknown
//! fields untouched.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The adapter keeps admission, accounting, and execution together."
)]

use std::mem::size_of;

use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::comment_storage_codec::{
    self, RewriteExecutionRequirements,
    lifecycle::{CommentStorageLifecycleRewrite, prepare_comment_storage_lifecycle_rewrite},
};

use super::budget::LifecycleBudget;
use super::{SlideMediaLifecycleError, SlideMediaLifecycleLimitKind};

// The neutral lifecycle path performs four bounded source-width passes: the
// source scan and rewrite measurement during preparation, followed by the
// candidate scan and Buffa parity decode during execution.  A 64-byte pass
// allowance covers the nested scalar/reference walks; keeping the multiplier
// explicit makes the aggregate bound auditable instead of hiding it in a
// large per-byte constant.
const NEUTRAL_PASS_COUNT: usize = 4;
const SOURCE_PASS_WORK_PER_BYTE: usize = 64;
const MAX_REPLY_REWRITE_GROWTH: usize = 18;
const PREPARATION_GROUP_SCRATCH: usize = 64 * size_of::<u32>() * 2;
const PREPARATION_ALLOCATION_EVENTS: usize = 8;

/// Rewrite every reply edge in a cloned comment-storage node.
///
/// The source fingerprint makes the prepared neutral rewrite an optimistic
/// concurrency witness.  Because no replacement UUID is supplied, the
/// neutral codec copies the source UUID and all non-reply spans byte-for-byte,
/// including author/date/text and unknown extensions.
pub(super) fn rewrite_comment_payload(
    payload: &[u8],
    object_remap: &[(u64, u64)],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u8>, SlideMediaLifecycleError> {
    let (options, preparation) = bounded_comment_options(payload.len(), object_remap.len())?;

    // Preparation owns several temporary reply/remap vectors and performs a
    // source and candidate sizing pass.  Charge their conservative maxima
    // before entering the neutral codec, whose allocations are otherwise
    // invisible to the package ledger.
    budget.charge_allocation_plan(preparation.scratch_bytes, preparation.allocations)?;
    budget.charge_wire_work(preparation.work_bytes)?;

    let rewrite = CommentStorageLifecycleRewrite::new(object_remap).expecting_fingerprint(
        comment_storage_codec::comment_storage_source_fingerprint(payload),
    );
    let prepared = prepare_comment_storage_lifecycle_rewrite(payload, rewrite, options)
        .map_err(map_comment_decode_error)?;
    let requirements = prepared.execution_requirements();

    // Execute only after every requirement has been admitted to the same
    // operation-wide ledger.  Reference bytes are charged as wire work since
    // the lifecycle budget has no separate byte-valued reference counter.
    charge_execution_requirements(requirements, budget)?;
    prepared
        .execute(requirements.exact())
        .map(|output| output.into_bytes())
        .map_err(map_comment_decode_error)
}

#[derive(Debug, Clone, Copy)]
struct PreparationBound {
    scratch_bytes: usize,
    work_bytes: usize,
    allocations: usize,
}

fn bounded_comment_options(
    payload_length: usize,
    remap_length: usize,
) -> Result<(comment_storage_codec::DecodeOptions, PreparationBound), SlideMediaLifecycleError> {
    let wire_limits = WireLimits::default();
    if payload_length > wire_limits.max_input_bytes() {
        return Err(limit(
            SlideMediaLifecycleLimitKind::InputBytes,
            payload_length,
            wire_limits.max_input_bytes(),
        ));
    }
    if remap_length > wire_limits.max_fields() {
        return Err(limit(
            SlideMediaLifecycleLimitKind::References,
            remap_length,
            wire_limits.max_fields(),
        ));
    }

    let source_length = payload_length.max(1);
    // A length-delimited field needs at least a one-byte key and one-byte
    // value/length. This is a conservative source-width upper bound for
    // dense varints, unlike using the whole source length as a field count.
    let field_upper_bound = (payload_length / 2).max(1);
    let reply_upper_bound = field_upper_bound;
    let aggregate_field_limit = wire_limits
        .max_fields()
        .checked_mul(NEUTRAL_PASS_COUNT)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let fields = field_upper_bound
        .checked_mul(NEUTRAL_PASS_COUNT)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if fields > aggregate_field_limit {
        return Err(limit(
            SlideMediaLifecycleLimitKind::WireFields,
            fields,
            aggregate_field_limit,
        ));
    }
    let references = reply_upper_bound
        .checked_mul(NEUTRAL_PASS_COUNT)
        .and_then(|count| count.max(remap_length).checked_add(1))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let aggregate_reference_limit = aggregate_field_limit;
    if references > aggregate_reference_limit {
        return Err(limit(
            SlideMediaLifecycleLimitKind::References,
            references,
            aggregate_reference_limit,
        ));
    }

    // Each changed reply can grow its nested identifier varint and the outer
    // length varint. The bound is deliberately expressed in terms of the
    // source-width reply count so it remains finite for dense hostile input.
    let message_growth = reply_upper_bound
        .checked_mul(MAX_REPLY_REWRITE_GROWTH)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let message_bytes = source_length
        .checked_add(message_growth)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if message_bytes > wire_limits.max_output_bytes() {
        return Err(limit(
            SlideMediaLifecycleLimitKind::OutputBytes,
            message_bytes,
            wire_limits.max_output_bytes(),
        ));
    }

    let recursion_limit = u32::try_from(wire_limits.max_nesting().min(64))
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    let source_pass_work = source_length
        .checked_mul(NEUTRAL_PASS_COUNT)
        .and_then(|passes| passes.checked_mul(SOURCE_PASS_WORK_PER_BYTE))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    // The neutral validator sorts source reply IDs and replacement IDs, then
    // binary-searches each possible reply against the replacement table. The
    // terms below mirror those operations and charge the comparison-log work
    // using the same word widths as the neutral implementation.
    let reply_sort_work = sort_work(reply_upper_bound)?;
    let remap_sort_work = sort_work(remap_length)?;
    let remap_lookup_work = lookup_work(reply_upper_bound, remap_length)?;
    let sorting_work = reply_sort_work
        .checked_add(remap_sort_work)
        .and_then(|work| work.checked_add(remap_lookup_work))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let work_bytes = source_pass_work
        .checked_add(sorting_work)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if work_bytes > wire_limits.max_rewrite_work() {
        return Err(limit(
            SlideMediaLifecycleLimitKind::WireWork,
            work_bytes,
            wire_limits.max_rewrite_work(),
        ));
    }

    let source_scratch = reply_upper_bound
        .checked_mul(size_of::<u64>())
        .and_then(|bytes| bytes.checked_mul(2))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let remap_scratch = remap_length
        .checked_mul(size_of::<(u64, u64)>())
        .and_then(|bytes| {
            remap_length
                .checked_mul(size_of::<u64>())
                .and_then(|sorted| bytes.checked_add(sorted))
        })
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let scratch_bytes = source_scratch
        .checked_add(remap_scratch)
        .and_then(|bytes| bytes.checked_add(PREPARATION_GROUP_SCRATCH))
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;

    Ok((
        comment_storage_codec::DecodeOptions::new(
            message_bytes,
            fields,
            work_bytes,
            recursion_limit,
            references,
            source_length,
        ),
        PreparationBound {
            scratch_bytes,
            work_bytes,
            allocations: PREPARATION_ALLOCATION_EVENTS,
        },
    ))
}

fn sort_work(length: usize) -> Result<usize, SlideMediaLifecycleError> {
    if length <= 1 {
        return Ok(0);
    }
    let leading_zeroes = usize::try_from((length - 1).leading_zeros())
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    let bits = usize::try_from(usize::BITS)
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?
        .checked_sub(leading_zeroes)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    length
        .checked_mul(bits)
        .and_then(|comparisons| comparisons.checked_mul(size_of::<u64>()))
        .ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn lookup_work(reply_count: usize, remap_count: usize) -> Result<usize, SlideMediaLifecycleError> {
    if reply_count == 0 || remap_count == 0 {
        return Ok(0);
    }
    let search_length = remap_count
        .checked_add(1)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let leading_zeroes = usize::try_from(search_length.leading_zeros())
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    let bits = usize::try_from(usize::BITS)
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?
        .checked_sub(leading_zeroes)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    reply_count
        .checked_mul(bits)
        .and_then(|searches| searches.checked_mul(3))
        .and_then(|searches| searches.checked_mul(size_of::<(u64, u64)>()))
        .ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn charge_execution_requirements(
    requirements: RewriteExecutionRequirements,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    budget.charge_output(requirements.output_bytes)?;
    budget.charge_wire_fields(requirements.fields)?;
    budget.charge_wire_work(requirements.work_bytes)?;
    budget.charge_wire_work(requirements.reference_bytes)?;
    budget.charge_nesting(
        usize::try_from(requirements.max_depth)
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?,
    )?;
    budget.charge_references(requirements.references)?;
    let retained_and_scratch = requirements
        .retained_bytes
        .checked_add(requirements.scratch_bytes)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocation_plan(retained_and_scratch, requirements.allocations)
}

fn map_comment_decode_error(error: comment_storage_codec::DecodeError) -> SlideMediaLifecycleError {
    let Some(limit_kind) = error.resource_limit() else {
        return SlideMediaLifecycleError::InvalidSource;
    };
    use comment_storage_codec::DecodeLimit;
    match limit_kind {
        DecodeLimit::Bytes { observed, maximum } => {
            mapped_limit(SlideMediaLifecycleLimitKind::InputBytes, observed, maximum)
        },
        DecodeLimit::OutputBytes { observed, maximum }
        | DecodeLimit::Retained { observed, maximum } => {
            mapped_limit(SlideMediaLifecycleLimitKind::OutputBytes, observed, maximum)
        },
        DecodeLimit::References { observed, maximum }
        | DecodeLimit::Replies { observed, maximum }
        | DecodeLimit::ReferenceBytes { observed, maximum } => {
            mapped_limit(SlideMediaLifecycleLimitKind::References, observed, maximum)
        },
        DecodeLimit::Fields { observed, maximum } => {
            mapped_limit(SlideMediaLifecycleLimitKind::WireFields, observed, maximum)
        },
        DecodeLimit::Nesting { observed, maximum } => limit_u64(
            SlideMediaLifecycleLimitKind::WireNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        DecodeLimit::Work { observed, maximum } | DecodeLimit::Text { observed, maximum } => {
            mapped_limit(SlideMediaLifecycleLimitKind::WireWork, observed, maximum)
        },
        DecodeLimit::Allocations { observed, maximum }
        | DecodeLimit::Scratch { observed, maximum } => {
            mapped_limit(SlideMediaLifecycleLimitKind::Allocations, observed, maximum)
        },
        _ => SlideMediaLifecycleError::InvalidSource,
    }
}

fn mapped_limit(
    kind: SlideMediaLifecycleLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideMediaLifecycleError {
    let (Ok(observed), Ok(maximum)) = (u64::try_from(observed), u64::try_from(maximum)) else {
        return SlideMediaLifecycleError::InvalidSource;
    };
    limit_u64(kind, observed, maximum)
}

fn limit(
    kind: SlideMediaLifecycleLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideMediaLifecycleError {
    mapped_limit(kind, observed, maximum)
}

fn limit_u64(
    kind: SlideMediaLifecycleLimitKind,
    observed: u64,
    maximum: u64,
) -> SlideMediaLifecycleError {
    SlideMediaLifecycleError::LimitExceeded {
        kind,
        observed,
        maximum,
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "Focused adapter fixtures use explicit test assertions."
)]
mod tests {
    use std::path::PathBuf;

    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};

    use super::*;

    const TEXT_FIELD: u32 = 1;
    const CREATION_DATE_FIELD: u32 = 2;
    const AUTHOR_FIELD: u32 = 3;
    const REPLIES_FIELD: u32 = 4;
    const STORAGE_UUID_FIELD: u32 = 5;
    const DATE_SECONDS_FIELD: u32 = 1;
    const UUID_LOWER_FIELD: u32 = 1;
    const UUID_UPPER_FIELD: u32 = 2;
    const UNKNOWN_ROOT_FIELD: u32 = 10_000;
    const UNKNOWN_REFERENCE_FIELD: u32 = 90;

    fn native_budget() -> LifecycleBudget {
        let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/media-comments-duplicate-native.key");
        let bytes = std::fs::read(path).expect("native comment fixture");
        let package = super::super::Package::from_bytes(&bytes).expect("native comment package");
        LifecycleBudget::for_package(&package).expect("lifecycle budget")
    }

    fn reference(identifier: u64, marker: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, 1, identifier).expect("reference identifier");
        append_length_delimited_field(&mut output, UNKNOWN_REFERENCE_FIELD, marker)
            .expect("reference marker");
        output
    }

    fn date(seconds_bits: u64) -> Vec<u8> {
        let mut output = Vec::new();
        let key = (u64::from(DATE_SECONDS_FIELD) << 3) | 1;
        output.push(u8::try_from(key).expect("date fixed64 key"));
        output.extend_from_slice(&seconds_bits.to_le_bytes());
        output
    }

    fn uuid(lower: u64, upper: u64) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint_field(&mut output, UUID_LOWER_FIELD, lower).expect("uuid lower");
        append_length_delimited_field(&mut output, UNKNOWN_REFERENCE_FIELD, b"uuid-unknown")
            .expect("uuid marker");
        append_varint_field(&mut output, UUID_UPPER_FIELD, upper).expect("uuid upper");
        output
    }

    fn comment_payload(author: u64, replies: &[u64]) -> Vec<u8> {
        let mut output = Vec::new();
        append_length_delimited_field(&mut output, TEXT_FIELD, b"adapter comment")
            .expect("comment text");
        append_length_delimited_field(
            &mut output,
            CREATION_DATE_FIELD,
            &date(0x0102_0304_0506_0708),
        )
        .expect("comment date");
        append_length_delimited_field(&mut output, AUTHOR_FIELD, &reference(author, b"author"))
            .expect("comment author");
        for &reply in replies {
            append_length_delimited_field(&mut output, REPLIES_FIELD, &reference(reply, b"reply"))
                .expect("comment reply");
        }
        append_length_delimited_field(
            &mut output,
            STORAGE_UUID_FIELD,
            &uuid(0x1122_3344_5566_7788, 0x99aa_bbcc_ddee_ff00),
        )
        .expect("comment uuid");
        append_length_delimited_field(&mut output, UNKNOWN_ROOT_FIELD, b"root-unknown")
            .expect("unknown root field");
        output
    }

    fn reply_ids(payload: &[u8]) -> Vec<u64> {
        let options = bounded_comment_options(payload.len(), 0)
            .expect("bounded comment options")
            .0;
        let mut identifiers = Vec::new();
        let result =
            comment_storage_codec::visit_comment_storage_replies(payload, options, &mut |reply| {
                identifiers.push(reply.identifier());
                Ok(())
            });
        result.expect("comment reply decode");
        identifiers
    }

    #[test]
    fn dense_reply_varint_growth_is_rewritten_with_bounded_options() {
        let source = comment_payload(41, &[1, 127, 128, 16_384]);
        let remaps = [
            (1, 0x0102_0304_0506_0708),
            (127, 0x1112_1314_1516_1718),
            (128, 0x2122_2324_2526_2728),
            (16_384, 0x3132_3334_3536_3738),
        ];
        let mut budget = native_budget();
        let output =
            rewrite_comment_payload(&source, &remaps, &mut budget).expect("dense reply rewrite");

        assert!(output.len() > source.len());
        assert_eq!(
            reply_ids(&output),
            remaps
                .iter()
                .map(|(_, replacement)| *replacement)
                .collect::<Vec<_>>()
        );
        assert!(
            output
                .windows(b"root-unknown".len())
                .any(|window| window == b"root-unknown")
        );
    }

    #[test]
    fn unrelated_global_mappings_do_not_change_author_or_reply_order() {
        let source = comment_payload(5, &[11, 22]);
        let remaps = [(5, 500), (11, 1_111), (22, 2_222), (99, 9_999)];
        let mut budget = native_budget();
        let output = rewrite_comment_payload(&source, &remaps, &mut budget)
            .expect("unrelated mapping rewrite");
        let options = bounded_comment_options(output.len(), remaps.len())
            .expect("bounded output options")
            .0;
        let (snapshot, _) =
            comment_storage_codec::decode_comment_storage_archive_with_report(&output, options)
                .expect("rewritten comment decode");

        assert_eq!(
            snapshot.author().map(|reference| reference.identifier()),
            Some(5)
        );
        assert_eq!(reply_ids(&output), [1_111, 2_222]);
        assert!(
            output
                .windows(b"adapter comment".len())
                .any(|window| window == b"adapter comment")
        );
        assert!(
            output
                .windows(b"root-unknown".len())
                .any(|window| window == b"root-unknown")
        );
    }

    #[test]
    fn source_author_uuid_date_and_unknown_bytes_are_preserved() {
        let source = comment_payload(41, &[7]);
        let source_before = source.clone();
        let remaps = [(7, 700)];
        let mut budget = native_budget();
        let output = rewrite_comment_payload(&source, &remaps, &mut budget)
            .expect("source-preserving rewrite");
        let options = bounded_comment_options(output.len(), remaps.len())
            .expect("bounded output options")
            .0;
        let (snapshot, _) =
            comment_storage_codec::decode_comment_storage_archive_with_report(&output, options)
                .expect("source-preserving decode");

        assert_eq!(snapshot.text(), Some("adapter comment"));
        assert_eq!(
            snapshot.creation_date().map(|date| date.seconds_bits()),
            Some(0x0102_0304_0506_0708)
        );
        assert_eq!(
            snapshot.author().map(|reference| reference.identifier()),
            Some(41)
        );
        assert_eq!(
            snapshot
                .storage_uuid()
                .map(|uuid| (uuid.lower(), uuid.upper())),
            Some((0x1122_3344_5566_7788, 0x99aa_bbcc_ddee_ff00))
        );
        assert!(
            output
                .windows(b"author".len())
                .any(|window| window == b"author")
        );
        assert!(
            output
                .windows(b"reply".len())
                .any(|window| window == b"reply")
        );
        assert!(
            output
                .windows(b"uuid-unknown".len())
                .any(|window| window == b"uuid-unknown")
        );
        assert_eq!(reply_ids(&output), [700]);
        assert_eq!(source, source_before);
    }

    #[test]
    fn oversized_and_unbounded_inputs_are_rejected_before_preparation() {
        let wire_limits = WireLimits::default();
        let input_error = bounded_comment_options(wire_limits.max_input_bytes() + 1, 0)
            .expect_err("oversized input");
        assert!(matches!(
            input_error,
            SlideMediaLifecycleError::LimitExceeded {
                kind: SlideMediaLifecycleLimitKind::InputBytes,
                ..
            }
        ));

        let remap_error = bounded_comment_options(32, wire_limits.max_fields() + 1)
            .expect_err("oversized remap table");
        assert!(matches!(
            remap_error,
            SlideMediaLifecycleError::LimitExceeded {
                kind: SlideMediaLifecycleLimitKind::References,
                ..
            }
        ));

        let pass_factor = NEUTRAL_PASS_COUNT
            .checked_mul(SOURCE_PASS_WORK_PER_BYTE)
            .expect("pass factor");
        let work_length = wire_limits.max_rewrite_work() / pass_factor + 1;
        let work_error = bounded_comment_options(work_length, 0).expect_err("work bound");
        assert!(matches!(
            work_error,
            SlideMediaLifecycleError::LimitExceeded {
                kind: SlideMediaLifecycleLimitKind::WireWork,
                ..
            }
        ));

        let huge_error = bounded_comment_options(usize::MAX, 0).expect_err("huge source");
        assert!(matches!(
            huge_error,
            SlideMediaLifecycleError::LimitExceeded {
                kind: SlideMediaLifecycleLimitKind::InputBytes,
                ..
            }
        ));
    }

    #[test]
    fn non_resource_codec_failures_map_to_invalid_source() {
        let options = comment_storage_codec::DecodeOptions::new(1, 1, 1, 1, 1, 1);
        let rewrite = CommentStorageLifecycleRewrite::new(&[]).expecting_fingerprint(0);
        let error = prepare_comment_storage_lifecycle_rewrite(&[], rewrite, options)
            .expect_err("empty comment payload");
        assert!(matches!(
            map_comment_decode_error(error),
            SlideMediaLifecycleError::InvalidSource
        ));
    }
}
