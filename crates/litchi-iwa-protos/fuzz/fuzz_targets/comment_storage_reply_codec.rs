#![no_main]

//! Bounded fuzzing for ordered `TSD.CommentStorageArchive.replies` rewrites.
//!
//! The target keeps the source borrowed through preparation, executes only
//! after replaying exact requirements, and reads every successful candidate
//! back through the strict visitor.  It deliberately exercises the three
//! ordered operations independently because a reply identifier is not a
//! stable selector after an earlier insertion or removal.

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::comment_storage_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_REFERENCES: usize = 1_024;
const MAX_TEXT_BYTES: usize = 64 * 1024;

// These are deliberately small and source-only.  Larger/deeper recipes live
// in the checked-in corpus so the fuzz target does not allocate them per call.
const FIXED_CASES: &[&[u8]] = &[
    &[],
    // One canonical reply reference.
    &[0x22, 0x02, 0x08, 0x01],
    // A balanced unknown group followed by a canonical reply.
    &[0x9b, 0x05, 0x08, 0x07, 0x9c, 0x05, 0x22, 0x02, 0x08, 0x01],
    // Truncated length-delimited reply and unterminated group.
    &[0x22, 0x03, 0x08],
    &[0x9b, 0x05, 0x08, 0x07],
    // Wrong wire payload for the nested reference.
    &[0x22, 0x01, 0x01],
];

#[derive(Default)]
struct ReplyIds {
    ids: Vec<u64>,
}

impl codec::CommentStorageVisitor for ReplyIds {
    fn visit_reply(&mut self, reply: codec::ReferenceRecord<'_>) -> Result<(), codec::DecodeError> {
        self.ids.push(reply.identifier());
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source);

    // Keep malformed, group, and empty-root paths hot independently of corpus
    // scheduling.  All are bounded and are passed by shared reference.
    for fixed in FIXED_CASES {
        exercise_source(fixed);
    }
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn decode_options(source_len: usize) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        MAX_INPUT_BYTES.max(source_len),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_TEXT_BYTES,
    )
}

fn collect_reply_ids(source: &[u8]) -> Result<(Vec<u64>, codec::DecodeReport), codec::DecodeError> {
    let mut replies = ReplyIds::default();
    let (_, report) = codec::decode_comment_storage_archive_with_visitor(
        source,
        decode_options(source.len()),
        &mut replies,
    )?;
    Ok((replies.ids, report))
}

fn exercise_source(source: &[u8]) {
    if source.len() > MAX_INPUT_BYTES {
        return;
    }
    let before = source.to_vec();
    let Ok((ids, source_report)) = collect_reply_ids(source) else {
        return;
    };

    let new_identifier = fresh_identifier(&ids, source.first().copied().unwrap_or(0));
    exercise_operation(
        source,
        &ids,
        codec::CommentStorageReplyRewrite::append(new_identifier),
        codec::CommentStorageReplyRewrite::remove(ids.len(), new_identifier),
    );

    if let Some(&expected_identifier) = ids.first() {
        let replacement_identifier = fresh_identifier(&ids, 0xa5);
        exercise_operation(
            source,
            &ids,
            codec::CommentStorageReplyRewrite::replace(
                0,
                expected_identifier,
                replacement_identifier,
            ),
            codec::CommentStorageReplyRewrite::replace(
                0,
                replacement_identifier,
                expected_identifier,
            ),
        );

        let wrong_expected = if expected_identifier == u64::MAX {
            1
        } else {
            expected_identifier.saturating_add(1)
        };
        let wrong =
            codec::CommentStorageReplyRewrite::replace(0, wrong_expected, replacement_identifier);
        assert!(
            codec::prepare_comment_storage_reply_rewrite(
                source,
                wrong,
                decode_options(source.len())
            )
            .is_err(),
            "stale expected reply identifier was accepted"
        );
    }

    if let Some(&expected_identifier) = ids.last() {
        exercise_operation(
            source,
            &ids,
            codec::CommentStorageReplyRewrite::remove(ids.len() - 1, expected_identifier),
            codec::CommentStorageReplyRewrite::append(expected_identifier),
        );
    } else {
        let wrong = codec::CommentStorageReplyRewrite::remove(0, 1);
        assert!(
            codec::prepare_comment_storage_reply_rewrite(
                source,
                wrong,
                decode_options(source.len())
            )
            .is_err(),
            "remove with no reply was accepted"
        );
    }

    black_box(source_report);
    assert_eq!(source, before.as_slice(), "reply codec modified its source");
}

fn fresh_identifier(ids: &[u64], seed: u8) -> u64 {
    let mut candidate = 0x1000_0000_0000_0000 | u64::from(seed);
    while candidate == 0 || ids.contains(&candidate) {
        candidate = candidate.wrapping_add(1);
        if candidate == 0 {
            candidate = 1;
        }
    }
    candidate
}

fn exercise_operation(
    source: &[u8],
    source_ids: &[u64],
    operation: codec::CommentStorageReplyRewrite,
    inverse: codec::CommentStorageReplyRewrite,
) {
    let options = decode_options(source.len());
    let Ok(plan) = codec::prepare_comment_storage_reply_rewrite(source, operation, options) else {
        return;
    };
    let requirements = plan.execution_requirements();
    let prepare_report = plan.prepare_report();
    assert_eq!(prepare_report.source_bytes(), source.len());
    assert_eq!(prepare_report.replies(), source_ids.len());
    assert_eq!(requirements.input_bytes, source.len());

    let output = plan
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact reply rewrite execution failed: {error}"));
    assert_output_report(output.report(), requirements);
    let candidate = output.bytes().to_vec();
    let candidate_ids = collect_reply_ids(&candidate)
        .unwrap_or_else(|error| panic!("reply rewrite readback failed: {error}"))
        .0;
    let expected_ids = expected_ids(source_ids, operation);
    assert_eq!(
        candidate_ids, expected_ids,
        "reply rewrite changed order unexpectedly"
    );
    assert_unknown_raw_fields_preserved(source, &candidate);

    let inverse_plan = codec::prepare_comment_storage_reply_rewrite(
        &candidate,
        inverse,
        decode_options(candidate.len()),
    )
    .unwrap_or_else(|error| panic!("inverse-like reply preparation failed: {error}"));
    let inverse_requirements = inverse_plan.execution_requirements();
    let restored = inverse_plan
        .execute(codec::RewriteExecutionLimits::exact(inverse_requirements))
        .unwrap_or_else(|error| panic!("inverse-like reply execution failed: {error}"));
    let restored_ids = collect_reply_ids(restored.bytes())
        .unwrap_or_else(|error| panic!("inverse-like readback failed: {error}"))
        .0;
    assert_eq!(
        restored_ids, source_ids,
        "inverse-like rewrite lost reply order"
    );

    exercise_max_minus_one(source, operation, requirements);
    black_box((candidate, restored));
}

fn expected_ids(source_ids: &[u64], operation: codec::CommentStorageReplyRewrite) -> Vec<u64> {
    let mut expected = source_ids.to_vec();
    match operation {
        codec::CommentStorageReplyRewrite::Append { identifier } => expected.push(identifier),
        codec::CommentStorageReplyRewrite::Replace {
            ordinal,
            replacement_identifier,
            ..
        } => {
            if let Some(identifier) = expected.get_mut(ordinal) {
                *identifier = replacement_identifier;
            }
        },
        codec::CommentStorageReplyRewrite::Remove { ordinal, .. } => {
            if ordinal < expected.len() {
                expected.remove(ordinal);
            }
        },
    }
    expected
}

fn assert_output_report(
    report: codec::RewriteReport,
    requirements: codec::RewriteExecutionRequirements,
) {
    assert_eq!(report.input_bytes(), requirements.input_bytes);
    assert_eq!(report.output_bytes(), requirements.output_bytes);
    assert_eq!(report.fields(), requirements.fields);
    assert_eq!(report.work_bytes(), requirements.work_bytes);
    assert_eq!(report.max_depth(), requirements.max_depth);
    assert_eq!(report.references(), requirements.references);
    assert_eq!(report.replies(), requirements.replies);
    assert_eq!(report.reference_bytes(), requirements.reference_bytes);
    assert_eq!(report.allocations(), requirements.allocations);
    assert_eq!(report.scratch_bytes(), requirements.scratch_bytes);
    assert_eq!(report.retained_bytes(), requirements.retained_bytes);
}

fn execute_with_limits(
    source: &[u8],
    operation: codec::CommentStorageReplyRewrite,
    limits: codec::RewriteExecutionLimits,
) -> Result<codec::RewriteOutput, codec::DecodeError> {
    let plan = codec::prepare_comment_storage_reply_rewrite(
        source,
        operation,
        decode_options(source.len()),
    )?;
    plan.execute(limits)
}

fn exercise_max_minus_one(
    source: &[u8],
    operation: codec::CommentStorageReplyRewrite,
    requirements: codec::RewriteExecutionRequirements,
) {
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let mut probes = Vec::new();
    if requirements.input_bytes > 0 {
        probes.push(exact.with_input_bytes(requirements.input_bytes - 1));
    }
    if requirements.output_bytes > 0 {
        probes.push(exact.with_output_bytes(requirements.output_bytes - 1));
    }
    if requirements.fields > 0 {
        probes.push(exact.with_fields(requirements.fields - 1));
    }
    if requirements.work_bytes > 0 {
        probes.push(exact.with_work_bytes(requirements.work_bytes - 1));
    }
    if requirements.max_depth > 0 {
        probes.push(exact.with_max_depth(requirements.max_depth - 1));
    }
    if requirements.references > 0 {
        probes.push(exact.with_references(requirements.references - 1));
    }
    if requirements.replies > 0 {
        probes.push(exact.with_replies(requirements.replies - 1));
    }
    if requirements.reference_bytes > 0 {
        probes.push(exact.with_reference_bytes(requirements.reference_bytes - 1));
    }
    if requirements.allocations > 0 {
        probes.push(exact.with_allocations(requirements.allocations - 1));
    }
    if requirements.scratch_bytes > 0 {
        probes.push(exact.with_scratch_bytes(requirements.scratch_bytes - 1));
    }
    if requirements.retained_bytes > 0 {
        probes.push(exact.with_retained_bytes(requirements.retained_bytes - 1));
    }
    for limits in probes {
        assert!(
            execute_with_limits(source, operation, limits).is_err(),
            "reply rewrite accepted a max-minus-one execution limit"
        );
    }
}

fn assert_unknown_raw_fields_preserved(source: &[u8], candidate: &[u8]) {
    // Field 91 group and an overlong scalar are the stable raw-preservation
    // markers used by the checked-in recipes.  Other unknown bytes are still
    // covered by the codec's source-preserving emitter.
    for marker in [
        &[0x9b, 0x05, 0x08, 0x07, 0x9c, 0x05][..],
        &[0xd0, 0x05, 0x81, 0x00][..],
    ] {
        if source.windows(marker.len()).any(|window| window == marker) {
            assert!(
                candidate
                    .windows(marker.len())
                    .any(|window| window == marker),
                "reply rewrite dropped an unknown raw field"
            );
        }
    }
}
