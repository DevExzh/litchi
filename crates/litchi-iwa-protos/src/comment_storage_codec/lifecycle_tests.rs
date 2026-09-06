use super::*;

fn varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn field_varint(output: &mut Vec<u8>, number: u32, value: u64) {
    key(output, number, 0);
    varint(output, value);
}

fn field_bytes(output: &mut Vec<u8>, number: u32, value: &[u8]) {
    key(output, number, 2);
    varint(output, value.len() as u64);
    output.extend_from_slice(value);
}

fn date(seconds_bits: u64) -> Vec<u8> {
    let mut output = Vec::new();
    key(&mut output, DATE_SECONDS_FIELD, 1);
    output.extend_from_slice(&seconds_bits.to_le_bytes());
    output
}

fn reference(identifier: u64, marker: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, REFERENCE_IDENTIFIER_FIELD, identifier);
    field_bytes(&mut output, 90, marker);
    output
}

fn uuid(lower: u64, upper: u64) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, UUID_LOWER_FIELD, lower);
    field_bytes(&mut output, 90, b"uuid-unknown");
    field_varint(&mut output, UUID_UPPER_FIELD, upper);
    output
}

fn source() -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 99, 123);
    field_bytes(&mut output, TEXT_FIELD, b"root text");
    field_bytes(
        &mut output,
        CREATION_DATE_FIELD,
        &date(0x8000_0000_0000_0000),
    );
    field_bytes(&mut output, AUTHOR_FIELD, &reference(41, b"author-unknown"));
    field_bytes(&mut output, REPLIES_FIELD, &reference(7, b"reply-seven"));
    field_bytes(&mut output, REPLIES_FIELD, &reference(128, b"reply-128"));
    field_bytes(
        &mut output,
        STORAGE_UUID_FIELD,
        &uuid(0x0102_0304, 0x0506_0708),
    );
    field_bytes(&mut output, 100, b"root-unknown");
    output
}

fn options(bytes: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        bytes.len().saturating_mul(2),
        256,
        bytes.len().saturating_mul(64).max(1),
        16,
        128,
        4096,
    )
}

#[derive(Default)]
struct Replies {
    identifiers: Vec<u64>,
}

impl CommentStorageVisitor for Replies {
    fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.identifiers.push(reply.identifier());
        Ok(())
    }
}

fn reply_ids(bytes: &[u8]) -> Vec<u64> {
    let mut replies = Replies::default();
    decode_comment_storage_archive_with_visitor(bytes, options(bytes), &mut replies)
        .expect("valid comment storage")
        .0;
    replies.identifiers
}

#[test]
fn lifecycle_rewrite_remaps_every_reply_and_preserves_untouched_spans() {
    let source = source();
    let old_uuid = UuidSnapshot::from_parts(0x0102_0304, 0x0506_0708);
    let new_uuid = UuidSnapshot::from_parts(0x0102_0304_0506_0708, 0x1112_1314_1516_1718);
    let remaps = [(7, 100), (128, 1_000)];
    let rewrite = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_fingerprint(comment_storage_source_fingerprint(&source))
        .expecting_storage_uuid(old_uuid)
        .replacing_storage_uuid(new_uuid);
    let prepared = prepare_comment_storage_lifecycle_rewrite(&source, rewrite, options(&source))
        .expect("preparation");
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(requirements.exact())
        .expect("execution")
        .into_bytes();

    assert_eq!(reply_ids(&output), [100, 1_000]);
    let snapshot = decode_comment_storage_archive(&output, options(&output)).expect("readback");
    assert_eq!(snapshot.text(), Some("root text"));
    assert_eq!(
        snapshot.author().map(ReferenceSnapshot::identifier),
        Some(41)
    );
    assert_eq!(snapshot.storage_uuid(), Some(new_uuid));
    for marker in [
        b"author-unknown".as_slice(),
        b"reply-seven".as_slice(),
        b"reply-128".as_slice(),
        b"uuid-unknown".as_slice(),
        b"root-unknown".as_slice(),
    ] {
        assert!(output.windows(marker.len()).any(|window| window == marker));
    }
    assert!(output.len() == requirements.output_bytes);
}

#[test]
fn lifecycle_rewrite_handles_mixed_reply_growth_and_uuid_shrink() {
    let source = source();
    let old_uuid = UuidSnapshot::from_parts(0x0102_0304, 0x0506_0708);
    let compact_uuid = UuidSnapshot::from_parts(1, 2);
    let remaps = [(7, 1), (128, 1_000_000)];
    let rewrite = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_storage_uuid(old_uuid)
        .replacing_storage_uuid(compact_uuid);
    let prepared = prepare_comment_storage_lifecycle_rewrite(&source, rewrite, options(&source))
        .expect("preparation");
    let requirements = prepared.execution_requirements();
    assert!(requirements.output_bytes < source.len());

    let mut emitted = Vec::new();
    emit_lifecycle_rewrite_fields(
        &mut emitted,
        &source,
        prepared.reply_remaps,
        prepared.replacement_storage_uuid,
    )
    .expect("raw emission");
    assert_eq!(
        emitted.len(),
        requirements.output_bytes,
        "emitted length differs from measured length"
    );
    let candidate_options = DecodeOptions {
        max_message_bytes: emitted.len().max(1),
        ..DecodeOptions::new(
            emitted.len().max(1),
            requirements.fields.max(1),
            requirements.work_bytes.max(1),
            requirements.max_depth.max(1),
            requirements.references.max(1),
            prepared.source_report.text_bytes.max(1),
        )
    };
    let candidate = scan_comment_storage_raw(&emitted, candidate_options, None)
        .expect("raw candidate verification");
    let (_, candidate_parity) =
        decode_comment_storage_archive_with_report(&emitted, candidate_options)
            .expect("Buffa candidate verification");
    assert_eq!(
        candidate.report, prepared.candidate_report,
        "raw candidate report differs from prepared report"
    );
    assert_eq!(
        candidate_parity, prepared.candidate_parity_report,
        "Buffa candidate report differs from prepared report"
    );
    let output = prepared
        .execute(requirements.exact())
        .expect("mixed growth/shrink execution")
        .into_bytes();

    assert_eq!(reply_ids(&output), [1, 1_000_000]);
    let snapshot = decode_comment_storage_archive(&output, options(&output)).expect("readback");
    assert_eq!(snapshot.storage_uuid(), Some(compact_uuid));
    assert!(
        output
            .windows(b"uuid-unknown".len())
            .any(|window| { window == b"uuid-unknown" })
    );
}

#[test]
fn lifecycle_rewrite_identity_is_an_exact_noop() {
    let source = source();
    let old_uuid = UuidSnapshot::from_parts(0x0102_0304, 0x0506_0708);
    let remaps = [(7, 7), (128, 128)];
    let rewrite = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_fingerprint(comment_storage_source_fingerprint(&source))
        .expecting_storage_uuid(old_uuid);
    let prepared = prepare_comment_storage_lifecycle_rewrite(&source, rewrite, options(&source))
        .expect("preparation");
    let limits = prepared.execution_requirements().exact();
    let output = prepared.execute(limits).expect("execution");
    assert_eq!(output.bytes(), source.as_slice());
    assert!(!output.report().changed());
}

#[test]
fn lifecycle_rewrite_requires_source_witnesses_and_sorted_unique_remaps() {
    let source = source();
    fn base(remaps: &[(u64, u64)]) -> CommentStorageLifecycleRewrite<'_> {
        CommentStorageLifecycleRewrite::new(remaps)
            .expecting_storage_uuid(UuidSnapshot::from_parts(0x0102_0304, 0x0506_0708))
    }

    assert!(
        prepare_comment_storage_lifecycle_rewrite(
            &source,
            CommentStorageLifecycleRewrite::new(&[]),
            options(&source),
        )
        .is_err()
    );
    assert!(
        prepare_comment_storage_lifecycle_rewrite(
            &source,
            base(&[(128, 100), (7, 101),]),
            options(&source),
        )
        .is_err()
    );
    assert!(
        prepare_comment_storage_lifecycle_rewrite(
            &source,
            base(&[(7, 100), (7, 101),]),
            options(&source),
        )
        .is_err()
    );
    assert!(
        prepare_comment_storage_lifecycle_rewrite(
            &source,
            base(&[(7, 100), (128, 100),]),
            options(&source),
        )
        .is_err()
    );
    assert!(
        prepare_comment_storage_lifecycle_rewrite(&source, base(&[(7, 0)]), options(&source),)
            .is_err()
    );
    let unrelated_remaps = [(999, 100)];
    let unrelated_rewrite = base(&unrelated_remaps);
    let unrelated =
        prepare_comment_storage_lifecycle_rewrite(&source, unrelated_rewrite, options(&source))
            .expect("unrelated package remaps are accepted");
    let limits = unrelated.execution_requirements().exact();
    let unrelated_output = unrelated
        .execute(limits)
        .expect("unrelated package remap execution");
    assert_eq!(unrelated_output.bytes(), source.as_slice());
}

#[test]
fn lifecycle_rewrite_rejects_stale_expectations_and_malformed_references() {
    let source = source();
    let old_uuid = UuidSnapshot::from_parts(0x0102_0304, 0x0506_0708);
    let remaps = [(7, 100)];
    let wrong_fingerprint = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_fingerprint(1)
        .expecting_storage_uuid(old_uuid);
    assert!(
        prepare_comment_storage_lifecycle_rewrite(&source, wrong_fingerprint, options(&source))
            .is_err()
    );
    let wrong_uuid = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_fingerprint(comment_storage_source_fingerprint(&source))
        .expecting_storage_uuid(UuidSnapshot::from_parts(1, 2));
    assert!(
        prepare_comment_storage_lifecycle_rewrite(&source, wrong_uuid, options(&source)).is_err()
    );

    let mut malformed = source;
    let malformed_reply = [0x08, 0x07, 0x08, 0x08];
    field_bytes(&mut malformed, REPLIES_FIELD, &malformed_reply);
    let rewrite = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_fingerprint(comment_storage_source_fingerprint(&malformed));
    assert!(
        prepare_comment_storage_lifecycle_rewrite(&malformed, rewrite, options(&malformed))
            .is_err()
    );
}

#[test]
fn lifecycle_rewrite_replays_exact_limits_before_candidate_allocation() {
    let source = source();
    let old_uuid = UuidSnapshot::from_parts(0x0102_0304, 0x0506_0708);
    let remaps = [(7, 100)];
    let rewrite = CommentStorageLifecycleRewrite::new(&remaps)
        .expecting_fingerprint(comment_storage_source_fingerprint(&source))
        .expecting_storage_uuid(old_uuid);
    let prepared = prepare_comment_storage_lifecycle_rewrite(&source, rewrite, options(&source))
        .expect("preparation");
    let requirements = prepared.execution_requirements();
    prepared
        .execute(requirements.exact())
        .expect("exact limits");

    let cases = [
        requirements
            .exact()
            .with_output_bytes(requirements.output_bytes - 1),
        requirements
            .exact()
            .with_work_bytes(requirements.work_bytes - 1),
        requirements
            .exact()
            .with_scratch_bytes(requirements.scratch_bytes - 1),
        requirements
            .exact()
            .with_allocations(requirements.allocations - 1),
    ];
    for limits in cases {
        let remaps = [(7, 100)];
        let rewrite = CommentStorageLifecycleRewrite::new(&remaps)
            .expecting_fingerprint(comment_storage_source_fingerprint(&source))
            .expecting_storage_uuid(old_uuid);
        let prepared =
            prepare_comment_storage_lifecycle_rewrite(&source, rewrite, options(&source))
                .expect("preparation");
        assert!(prepared.execute(limits).is_err());
    }
}

#[test]
fn lifecycle_rewrite_remap_work_scales_logarithmically_and_preempts_work() {
    let source = source();
    let make_remaps = |count: u64| {
        (1..=count)
            .map(|identifier| (identifier, 1_000_000 + identifier))
            .collect::<Vec<_>>()
    };
    let lifecycle_options = |max_references: usize, max_work_bytes: usize| {
        DecodeOptions::new(
            source.len().saturating_mul(4),
            1_000_000,
            max_work_bytes,
            16,
            max_references,
            4096,
        )
    };

    let remaps_4096 = make_remaps(4096);
    let prepared_4096 = prepare_comment_storage_lifecycle_rewrite(
        &source,
        CommentStorageLifecycleRewrite::new(&remaps_4096)
            .expecting_fingerprint(comment_storage_source_fingerprint(&source)),
        lifecycle_options(8192, usize::MAX),
    )
    .expect("4096 remap preparation");
    let requirements_4096 = prepared_4096.execution_requirements();

    let remaps_8192 = make_remaps(8192);
    let prepared_8192 = prepare_comment_storage_lifecycle_rewrite(
        &source,
        CommentStorageLifecycleRewrite::new(&remaps_8192)
            .expecting_fingerprint(comment_storage_source_fingerprint(&source)),
        lifecycle_options(8192, usize::MAX),
    )
    .expect("8192 remap preparation");
    let requirements_8192 = prepared_8192.execution_requirements();
    assert!(requirements_8192.work_bytes > requirements_4096.work_bytes);

    let source_raw = scan_comment_storage_raw(&source, lifecycle_options(8192, usize::MAX), None)
        .expect("source raw preflight");
    let (_, source_parity) =
        decode_comment_storage_archive_with_report(&source, lifecycle_options(8192, usize::MAX))
            .expect("source Buffa parity preflight");
    // Leave enough budget for the complete source pass and fingerprint, but
    // less than the remap validator's sort/lookup charge. The validator must
    // reject before either sorted Vec is allocated.
    let early_sort_limit = source_raw
        .report
        .work_bytes()
        .checked_add(source.len())
        .and_then(|work| work.checked_add(source_parity.work_bytes()))
        .and_then(|work| work.checked_add(1))
        .expect("bounded work arithmetic");
    let constrained = lifecycle_options(8192, early_sort_limit);
    let error = prepare_comment_storage_lifecycle_rewrite(
        &source,
        CommentStorageLifecycleRewrite::new(&remaps_8192)
            .expecting_fingerprint(comment_storage_source_fingerprint(&source)),
        constrained,
    )
    .expect_err("work ceiling");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::Work { .. })
    ));
}
