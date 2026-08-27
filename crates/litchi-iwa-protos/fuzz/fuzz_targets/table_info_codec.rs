#![no_main]

//! Bounded fuzzing for the strict `TST.TableInfoArchive` projection.
//!
//! The target keeps the caller-owned protobuf bytes authoritative while it
//! exercises both the complete lock-aware snapshot and the model-reference
//! convenience path.  It deliberately does not construct a package, generated
//! archive, or native object graph.  The package owners consume this projection
//! for table ownership and lock admission.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::table_info_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;

// Small complete and malformed wire recipes keep the strict root, nested
// reference, lock presence, unknown framing, and failure boundaries reachable
// even when arbitrary input is not a valid protobuf payload.
const FIXED_CASES: &[&[u8]] = &[
    // TableInfoArchive { super: {}, table_model: { identifier: 42 } }.
    &[0x0a, 0x00, 0x12, 0x02, 0x08, 0x2a],
    // The same source with an explicitly absent-equivalent lock value.
    &[0x0a, 0x02, 0x28, 0x00, 0x12, 0x02, 0x08, 0x2a],
    // The same source with locked=true.
    &[0x0a, 0x02, 0x28, 0x01, 0x12, 0x02, 0x08, 0x2a],
    // Canonical unknown scalar/fixed64/bytes/fixed32 and balanced groups at
    // both selected envelopes, with locked=true retained.
    &[
        0x0a, 0x16, // DrawableArchive.super
        0x08, 0x07, // unknown varint
        0x11, 0, 1, 2, 3, 4, 5, 6, 7, // unknown fixed64
        0x1a, 0x02, 0xaa, 0xbb, // unknown bytes
        0x25, 9, 8, 7, 6, // unknown fixed32
        0x28, 0x01, // locked=true
        0x12, 0x08, // selected table-model reference
        0x08, 0x2a, // identifier=42
        0x10, 0x09, // unknown nested scalar
        0x1b, 0x08, 0x0b, 0x1c, // balanced unknown nested group
        0x98, 0x06, 0x01, // root unknown scalar
        0xa1, 0x06, 0, 1, 2, 3, 4, 5, 6, 7, // root fixed64
        0xaa, 0x06, 0x01, 0xff, // root bytes
        0xb5, 0x06, 0, 1, 2, 3, // root fixed32
        0xbb, 0x06, 0x08, 0x01, 0xbc, 0x06, // root group
    ],
    // Missing required super envelope.
    &[0x12, 0x02, 0x08, 0x2a],
    // Missing required table-model reference.
    &[0x0a, 0x00],
    // Missing nested identifier.
    &[0x0a, 0x00, 0x12, 0x00],
    // Zero nested identifier.
    &[0x0a, 0x00, 0x12, 0x02, 0x08, 0x00],
    // Duplicate required root envelopes.
    &[0x0a, 0x00, 0x0a, 0x00, 0x12, 0x02, 0x08, 0x01],
    // Duplicate selected model reference.
    &[0x0a, 0x00, 0x12, 0x02, 0x08, 0x01, 0x12, 0x02, 0x08, 0x02],
    // Duplicate nested identifier.
    &[0x0a, 0x00, 0x12, 0x04, 0x08, 0x01, 0x08, 0x02],
    // Duplicate lock field.
    &[0x0a, 0x04, 0x28, 0x00, 0x28, 0x01, 0x12, 0x02, 0x08, 0x2a],
    // Known fields with wrong wire types.
    &[0x0a, 0x00, 0x10, 0x2a],
    &[0x08, 0x00, 0x12, 0x02, 0x08, 0x2a],
    &[0x0a, 0x02, 0x2a, 0x00, 0x12, 0x02, 0x08, 0x2a],
    // Non-canonical key, length, identifier, and lock encodings.
    &[0x0a, 0x00, 0x92, 0x00, 0x02, 0x08, 0x2a],
    &[0x0a, 0x80, 0x00, 0x12, 0x02, 0x08, 0x2a],
    &[0x0a, 0x00, 0x12, 0x03, 0x08, 0xaa, 0x00],
    &[0x0a, 0x03, 0x28, 0x81, 0x00, 0x12, 0x02, 0x08, 0x2a],
    // Non-boolean known lock scalar and truncated nested framing.
    &[0x0a, 0x02, 0x28, 0x02, 0x12, 0x02, 0x08, 0x2a],
    &[0x0a, 0x00, 0x12, 0x01, 0x08],
    // Unterminated, mismatched, and nested unknown groups.
    &[0x0a, 0x01, 0x0b, 0x12, 0x02, 0x08, 0x01],
    &[0x0a, 0x03, 0x0b, 0x08, 0x01, 0x12, 0x02, 0x08, 0x01],
    &[0x0a, 0x03, 0x0b, 0x0c, 0x0c, 0x12, 0x02, 0x08, 0x01],
];

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data) {
        exercise_source(&source);
    }

    // Keep deterministic semantic and malformed cases in every campaign.
    static FIXED: OnceLock<()> = OnceLock::new();
    FIXED.get_or_init(|| {
        for source in FIXED_CASES {
            exercise_source(source);
        }
        exercise_depth_boundary();
        exercise_limit_boundaries();
    });
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

fn options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    let options = options(source);
    let snapshot = codec::decode_table_info(source, options);
    let reference = codec::decode_table_model_reference(source, options);
    assert_eq!(source, before.as_slice(), "decode modified caller bytes");

    match (snapshot, reference) {
        (Ok(snapshot), Ok(reference)) => {
            assert_eq!(snapshot.table_model(), reference);
            let locked = snapshot.locked();
            assert!(locked.is_none() || locked == Some(false) || locked == Some(true));
            black_box((reference.identifier(), locked));

            exercise_lock_rewrites(source, snapshot);
        },
        (Err(snapshot_error), Err(reference_error)) => {
            // Both APIs share the same strict preflight. Keep the error values
            // opaque to the fuzz target while ensuring they remain printable
            // and cannot contain native identifiers or source bytes.
            black_box((snapshot_error.to_string(), reference_error.to_string()));
        },
        (snapshot, reference) => panic!(
            "table-info projection disagreement: snapshot={snapshot:?}, reference={reference:?}"
        ),
    }
}

fn rewrite_options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
    .with_max_output_bytes(MAX_INPUT_BYTES)
    .with_max_allocations(MAX_INPUT_BYTES)
    .with_max_retained_bytes(MAX_INPUT_BYTES.saturating_mul(2))
    .with_max_scratch_bytes(MAX_INPUT_BYTES.saturating_mul(3))
}

fn exercise_lock_rewrites(source: &[u8], snapshot: codec::TableInfoSnapshot) {
    let before = source.to_vec();
    let options = rewrite_options(source);
    let fingerprint = codec::table_info_source_fingerprint(source);

    for requested in [None, Some(false), Some(true)] {
        let write = codec::TableInfoLockWrite::explicit(requested);
        let prepared = match codec::prepare_table_info_lock_rewrite(source, write, options) {
            Ok(prepared) => prepared,
            Err(error) => {
                // Inputs accepted by the strict read profile can still exceed a
                // rewrite profile's aggregate ceilings. Keep that path opaque,
                // but never let the failed preparation mutate caller bytes.
                black_box(error.to_string());
                assert_eq!(source, before.as_slice());
                continue;
            },
        };
        let report = prepared.prepare_report();
        let requirements = prepared.execution_requirements();
        assert_eq!(report.input_bytes(), requirements.input_bytes());
        assert_eq!(report.output_bytes(), requirements.output_bytes());
        assert_eq!(report.fields(), requirements.fields());
        assert_eq!(report.work_bytes(), requirements.work_bytes());
        assert_eq!(report.max_depth(), requirements.max_depth());
        assert_eq!(report.allocations(), requirements.allocations());
        assert_eq!(report.retained_bytes(), requirements.retained_bytes());
        assert_eq!(report.scratch_bytes(), requirements.scratch_bytes());
        assert_eq!(report.source_fingerprint(), fingerprint);
        assert_eq!(prepared.source_fingerprint(), fingerprint);
        assert_eq!(report.changed(), requested != snapshot.locked());

        exercise_prepare_limit_boundaries(source, write, options, requirements);
        exercise_execute_limit_boundaries(&prepared, requirements);

        let output = prepared
            .clone()
            .execute(requirements.exact())
            .expect("exact prepared lock limits must execute");
        let candidate = output.as_bytes().to_vec();
        let output_report = output.report();
        assert_eq!(output_report, report);
        assert_eq!(source, before.as_slice(), "execute modified source bytes");

        let candidate_snapshot =
            codec::decode_table_info(&candidate, codec::DecodeOptions::for_source(&candidate))
                .expect("prepared candidate must remain a strict TableInfo");
        assert_eq!(candidate_snapshot.table_model(), snapshot.table_model());
        assert_eq!(candidate_snapshot.locked(), requested);
        if requested == snapshot.locked() {
            assert_eq!(candidate, before, "no-op rewrite changed source bytes");
        }

        let inverse_options = rewrite_options(&candidate);
        let inverse = codec::rewrite_table_info_lock(
            &candidate,
            codec::TableInfoLockWrite::explicit(snapshot.locked()),
            inverse_options,
        )
        .expect("inverse lock rewrite must execute");
        assert_eq!(inverse, before, "lock rewrite inverse was not byte exact");

        let (one_shot, one_shot_report) =
            codec::rewrite_table_info_lock_with_report(source, write, options)
                .expect("one-shot lock rewrite must execute");
        assert_eq!(one_shot, candidate);
        assert_eq!(one_shot_report, report);
        assert_eq!(source, before.as_slice(), "one-shot modified source bytes");

        let mismatch = codec::prepare_table_info_lock_rewrite_with_fingerprint(
            source,
            fingerprint ^ 1,
            write,
            options,
        )
        .expect_err("stale lock rewrite fingerprint must fail");
        assert!(mismatch.is_fingerprint_mismatch());
        assert_eq!(
            source,
            before.as_slice(),
            "stale rewrite changed source bytes"
        );
    }
}

fn exercise_prepare_limit_boundaries(
    source: &[u8],
    write: codec::TableInfoLockWrite,
    options: codec::DecodeOptions,
    requirements: codec::RewriteExecutionRequirements,
) {
    if requirements.input_bytes() > 0 {
        let limited = codec::DecodeOptions::new(
            requirements.input_bytes() - 1,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
        )
        .with_max_output_bytes(MAX_INPUT_BYTES)
        .with_max_allocations(MAX_INPUT_BYTES)
        .with_max_retained_bytes(MAX_INPUT_BYTES.saturating_mul(2))
        .with_max_scratch_bytes(MAX_INPUT_BYTES.saturating_mul(3));
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare input limit must fail");
        black_box(error.wire_resource_limit());
    }
    if requirements.fields() > 0 {
        let limited = codec::DecodeOptions::new(
            source.len().max(1),
            requirements.fields() - 1,
            MAX_WORK_BYTES,
            MAX_RECURSION,
        )
        .with_max_output_bytes(MAX_INPUT_BYTES)
        .with_max_allocations(MAX_INPUT_BYTES)
        .with_max_retained_bytes(MAX_INPUT_BYTES.saturating_mul(2))
        .with_max_scratch_bytes(MAX_INPUT_BYTES.saturating_mul(3));
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare field limit must fail");
        black_box(error.field_limit_values());
    }
    if requirements.work_bytes() > 0 {
        let limited = codec::DecodeOptions::new(
            source.len().max(1),
            MAX_FIELDS,
            requirements.work_bytes() - 1,
            MAX_RECURSION,
        )
        .with_max_output_bytes(MAX_INPUT_BYTES)
        .with_max_allocations(MAX_INPUT_BYTES)
        .with_max_retained_bytes(MAX_INPUT_BYTES.saturating_mul(2))
        .with_max_scratch_bytes(MAX_INPUT_BYTES.saturating_mul(3));
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare work limit must fail");
        black_box(error.work_limit_values());
    }
    if requirements.max_depth() > 0 {
        let limited = codec::DecodeOptions::new(
            source.len().max(1),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            requirements.max_depth() - 1,
        )
        .with_max_output_bytes(MAX_INPUT_BYTES)
        .with_max_allocations(MAX_INPUT_BYTES)
        .with_max_retained_bytes(MAX_INPUT_BYTES.saturating_mul(2))
        .with_max_scratch_bytes(MAX_INPUT_BYTES.saturating_mul(3));
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare nesting limit must fail");
        black_box(error.wire_resource_limit());
    }
    if requirements.output_bytes() > 0 {
        let limited = options.with_max_output_bytes(requirements.output_bytes() - 1);
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare output limit must fail");
        black_box(error.output_limit_values());
    }
    if requirements.allocations() > 0 {
        let limited = options.with_max_allocations(requirements.allocations() - 1);
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare allocation limit must fail");
        black_box(error.allocation_limit_values());
    }
    if requirements.retained_bytes() > 0 {
        let limited = options.with_max_retained_bytes(requirements.retained_bytes() - 1);
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare retained limit must fail");
        black_box(error.retained_limit_values());
    }
    if requirements.scratch_bytes() > 0 {
        let limited = options.with_max_scratch_bytes(requirements.scratch_bytes() - 1);
        let error = codec::prepare_table_info_lock_rewrite(source, write, limited)
            .expect_err("one-less prepare scratch limit must fail");
        black_box(error.scratch_limit_values());
    }
}

fn exercise_execute_limit_boundaries(
    prepared: &codec::PreparedTableInfoLockRewrite<'_>,
    requirements: codec::RewriteExecutionRequirements,
) {
    if requirements.input_bytes() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_input_bytes(requirements.input_bytes() - 1),
            )
            .expect_err("one-less execute input limit must fail");
        black_box(error.wire_resource_limit());
    }
    if requirements.output_bytes() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_output_bytes(requirements.output_bytes() - 1),
            )
            .expect_err("one-less execute output limit must fail");
        black_box(error.output_limit_values());
    }
    if requirements.fields() > 0 {
        let error = prepared
            .clone()
            .execute(requirements.exact().with_fields(requirements.fields() - 1))
            .expect_err("one-less execute field limit must fail");
        black_box(error.field_limit_values());
    }
    if requirements.work_bytes() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_work_bytes(requirements.work_bytes() - 1),
            )
            .expect_err("one-less execute work limit must fail");
        black_box(error.work_limit_values());
    }
    if requirements.max_depth() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_max_depth(requirements.max_depth() - 1),
            )
            .expect_err("one-less execute nesting limit must fail");
        black_box(error.wire_resource_limit());
    }
    if requirements.allocations() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_allocations(requirements.allocations() - 1),
            )
            .expect_err("one-less execute allocation limit must fail");
        black_box(error.allocation_limit_values());
    }
    if requirements.retained_bytes() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_retained_bytes(requirements.retained_bytes() - 1),
            )
            .expect_err("one-less execute retained limit must fail");
        black_box(error.retained_limit_values());
    }
    if requirements.scratch_bytes() > 0 {
        let error = prepared
            .clone()
            .execute(
                requirements
                    .exact()
                    .with_scratch_bytes(requirements.scratch_bytes() - 1),
            )
            .expect_err("one-less execute scratch limit must fail");
        black_box(error.scratch_limit_values());
    }
}

fn exercise_depth_boundary() {
    // A balanced unknown group nested inside the required super envelope is
    // valid at the normal depth profile and fails closed at depth one.
    let source = [0x0a, 0x04, 0x0b, 0x0c, 0x28, 0x00, 0x12, 0x02, 0x08, 0x01];
    let before = source;
    let exact = codec::DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, 64);
    let _ = codec::decode_table_info(&source, exact);
    assert_eq!(source, before);

    let shallow = codec::DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, 1);
    if let Err(error) = codec::decode_table_info(&source, shallow) {
        black_box(error.wire_resource_limit());
    }
}

fn exercise_limit_boundaries() {
    let source = [0x0a, 0x00, 0x12, 0x02, 0x08, 0x01];
    let before = source;

    // The codec's documented fixed profile is exact for this minimal source:
    // three strict fields and 16 bytes of strict-plus-projection work.
    let exact = codec::DecodeOptions::new(source.len(), 3, 16, 2);
    assert!(codec::decode_table_info(&source, exact).is_ok());
    assert_eq!(source, before);

    let one_less_input = codec::DecodeOptions::new(source.len() - 1, 3, 16, 2);
    if let Err(error) = codec::decode_table_info(&source, one_less_input) {
        black_box(error.wire_resource_limit());
    }
    let one_less_fields = codec::DecodeOptions::new(source.len(), 2, 16, 2);
    if let Err(error) = codec::decode_table_info(&source, one_less_fields) {
        black_box(error.field_limit_values());
    }
    let one_less_work = codec::DecodeOptions::new(source.len(), 3, 15, 2);
    if let Err(error) = codec::decode_table_info(&source, one_less_work) {
        black_box(error.work_limit_values());
    }
    let zero_nesting = codec::DecodeOptions::new(source.len(), 3, 16, 0);
    if let Err(error) = codec::decode_table_info(&source, zero_nesting) {
        black_box(error.wire_resource_limit());
    }
    assert_eq!(source, before, "limit probes modified caller bytes");
}
