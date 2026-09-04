#![no_main]

//! Strict, source-preserving fuzzing for the native Numbers Text
//! `FormatStructArchive` projection.
//!
//! Text is intentionally a small format: the selected payload contains only
//! the native discriminator (`260`).  The target still treats the complete
//! source as authoritative, so unknown records/groups are retained and every
//! successful decode/rewrite remains bounded, borrowed, and exact.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_text_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const CANONICAL_TEXT: &[u8] = &[0x08, 0x84, 0x02];
const UNKNOWN_SCALAR: &[u8] = &[0xa0, 0x06, 0x81, 0x00];
const UNKNOWN_GROUP: &[u8] = &[0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06];
const UNKNOWN_LENGTH: &[u8] = &[0xb2, 0x06, 0x03, 0x01, 0x7f, 0x00];

const FIXED_CASES: &[&[u8]] = &[
    // Canonical native Text (field 1 = 260).
    CANONICAL_TEXT,
    // Unknown scalar, fixed-width, and length-delimited records remain
    // opaque while the selected discriminator stays strict.
    &[
        0xa0, 0x06, 0x81, 0x00, 0x08, 0x84, 0x02, 0xe9, 0x06, 0, 1, 2, 3, 4, 5, 6, 7, 0xed, 0x06,
        8, 9, 10, 11,
    ],
    // A balanced unknown group is retained across a source-preserving write.
    &[
        0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06, 0x08, 0x84, 0x02, 0xb2, 0x06, 0x03, 1, 2, 3,
    ],
    // Unknown records may be interleaved on either side of the known field.
    &[
        0xb2, 0x06, 0x03, b't', b'x', b't', 0x08, 0x84, 0x02, 0xa0, 0x06, 0x81, 0x00,
    ],
    // Missing, duplicate, and wrong-wire/type known fields.
    &[],
    &[0x08, 0x84, 0x02, 0x08, 0x84, 0x02],
    &[0x0a, 0x01, 0x00],
    &[0x08, 0x83, 0x02],
    // Non-canonical key/value encodings are rejected before projection.
    &[0x88, 0x00, 0x84, 0x02],
    &[0x08, 0x84, 0x82, 0x00],
    // Other known FormatStructArchive fields cannot be hidden as unknown.
    &[0x08, 0x84, 0x02, 0x10, 0x00],
    &[0x08, 0x84, 0x02, 0x20, 0x00],
    &[0x08, 0x84, 0x02, 0xa0, 0x01, 0x00],
    &[0x08, 0x84, 0x02, 0xe8, 0x02, 0x00],
    // Truncated, mismatched, and unterminated groups remain malformed.
    &[0x08, 0x84],
    &[0x08, 0x84, 0x02, 0xa3, 0x06, 0xa8, 0x06, 0x01],
    &[0x08, 0x84, 0x02, 0xa3, 0x06, 0xa8, 0x06, 0x01, 0xac, 0x06],
    &[0x08, 0x84, 0x02, 0xa4, 0x06],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    // Keep canonical, unknown-field, malformed, and limit paths reachable
    // even when a campaign starts with only arbitrary invalid bytes.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-text-format");
        }
        exercise_canonical_writes();
        exercise_encoding_markers();
        exercise_limit_guards();
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
        MAX_INPUT_BYTES.max(source.len()),
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    )
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_text_format(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "scalar decode modified its source"
    );
    let reported = codec::decode_text_format_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "reported decode modified its source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.format_type(), codec::NATIVE_TEXT_FORMAT_TYPE);
            assert_eq!(snapshot.raw(), source);
            assert_borrowed(source, snapshot.raw());
            assert_report(report, source.len());
            exercise_limit_profiles(source, report);
            black_box((snapshot, report));
            exercise_rewrite(source, snapshot, data, &before);
        },
        (Err(scalar_error), Err(reported_error)) => {
            assert_eq!(
                scalar_error.resource_limit(),
                reported_error.resource_limit(),
                "Text scalar/report error classifications disagree"
            );
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "Text-format scalar/report disagreement: scalar={:?}, report={:?}",
            scalar_result.map(|snapshot| snapshot.format_type()),
            reported_result.map(|(snapshot, _)| snapshot.format_type())
        ),
    }

    assert_eq!(
        source,
        before.as_slice(),
        "Text-format probes modified their source"
    );
}

fn assert_borrowed(source: &[u8], borrowed: &[u8]) {
    if borrowed.is_empty() {
        return;
    }
    let source_start = source.as_ptr() as usize;
    let source_end = source_start
        .checked_add(source.len())
        .expect("bounded source pointer range");
    let borrowed_start = borrowed.as_ptr() as usize;
    let borrowed_end = borrowed_start
        .checked_add(borrowed.len())
        .expect("bounded borrowed pointer range");
    assert!(
        borrowed_start >= source_start && borrowed_end <= source_end,
        "Text-format snapshot did not borrow from its source"
    );
}

fn assert_report(report: codec::DecodeReport, source_len: usize) {
    assert_eq!(report.input_bytes(), source_len);
    assert_eq!(report.output_bytes(), 0);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert_eq!(report.references(), 0);
    assert_eq!(report.items(), 0);
    assert_eq!(report.text_bytes(), 0);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.retained_bytes(), 0);
    assert_eq!(report.scratch_bytes(), 0);
}

fn exercise_rewrite(
    source: &[u8],
    snapshot: codec::TextFormatSnapshot<'_>,
    data: &[u8],
    before: &[u8],
) {
    let write = if data.first().copied().unwrap_or_default() & 1 == 0 {
        codec::TextFormatWrite::from_snapshot(snapshot)
    } else {
        codec::TextFormatWrite::new()
    };
    let prepared = match codec::prepare_text_format_rewrite(source, write, options(source)) {
        Ok(prepared) => prepared,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let requirements = prepared.execution_requirements();
    assert_prepare_report(prepared.prepare_report(), requirements);
    assert_eq!(source, before, "prepare modified its source");

    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact Text-format replay failed: {error}"));
    assert_eq!(source, before, "execute modified its source");
    assert_eq!(
        output.bytes(),
        source,
        "Text preserving rewrite changed bytes"
    );
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    assert_eq!(output.report(), prepared.prepare_report());

    let readback = codec::decode_text_format(output.bytes(), options(output.bytes()))
        .unwrap_or_else(|error| panic!("Text-format candidate readback failed: {error}"));
    assert_eq!(codec::TextFormatWrite::from_snapshot(readback), write);
    assert_known_unknown_spans_preserved(source, output.bytes());

    let one_shot = codec::rewrite_text_format(source, write, options(source))
        .unwrap_or_else(|error| panic!("one-shot Text-format rewrite failed: {error}"));
    assert_eq!(one_shot.bytes(), output.bytes());
    assert_eq!(one_shot.report(), output.report());
    exercise_limit_failures(source, write, prepared, requirements);
    black_box(output);
}

fn assert_prepare_report(
    report: codec::DecodeReport,
    requirements: codec::RewriteExecutionRequirements,
) {
    assert_eq!(report.output_bytes(), requirements.output_bytes());
    assert_eq!(report.fields(), requirements.fields());
    assert_eq!(report.work_bytes(), requirements.work_bytes());
    assert_eq!(report.max_depth(), requirements.max_depth());
    assert_eq!(report.references(), requirements.references());
    assert_eq!(report.items(), requirements.items());
    assert_eq!(report.text_bytes(), requirements.text_bytes());
    assert_eq!(report.allocations(), requirements.allocations());
    assert_eq!(report.retained_bytes(), requirements.retained_bytes());
    assert_eq!(report.scratch_bytes(), requirements.scratch_bytes());
}

fn assert_known_unknown_spans_preserved(source: &[u8], candidate: &[u8]) {
    for unknown in [UNKNOWN_SCALAR, UNKNOWN_GROUP, UNKNOWN_LENGTH] {
        if source
            .windows(unknown.len())
            .any(|window| window == unknown)
        {
            assert!(
                candidate
                    .windows(unknown.len())
                    .any(|window| window == unknown),
                "source unknown Text-format span was not preserved"
            );
        }
    }
}

fn exercise_limit_failures(
    source: &[u8],
    write: codec::TextFormatWrite,
    prepared: codec::PreparedTextFormatRewrite<'_>,
    requirements: codec::RewriteExecutionRequirements,
) {
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let probes = [
        (requirements.output_bytes() > 0)
            .then(|| exact.with_output_bytes(requirements.output_bytes() - 1)),
        (requirements.fields() > 0).then(|| exact.with_fields(requirements.fields() - 1)),
        (requirements.work_bytes() > 0)
            .then(|| exact.with_work_bytes(requirements.work_bytes() - 1)),
        (requirements.max_depth() > 0).then(|| exact.with_max_depth(requirements.max_depth() - 1)),
        (requirements.allocations() > 0)
            .then(|| exact.with_allocations(requirements.allocations() - 1)),
        (requirements.retained_bytes() > 0)
            .then(|| exact.with_retained_bytes(requirements.retained_bytes() - 1)),
        (requirements.scratch_bytes() > 0)
            .then(|| exact.with_scratch_bytes(requirements.scratch_bytes() - 1)),
    ];
    for limits in probes.into_iter().flatten() {
        let result = prepared.execute(limits);
        assert!(
            result.is_err(),
            "Text-format rewrite accepted a strict limit"
        );
        if let Err(error) = result {
            black_box(error.resource_limit());
            observe_error(error);
        }
    }

    let replay = codec::rewrite_text_format(source, write, options(source));
    if let Err(error) = replay {
        observe_error(error);
    }
}

fn exercise_limit_profiles(source: &[u8], report: codec::DecodeReport) {
    if source.is_empty() {
        return;
    }
    let input_limited = codec::DecodeOptions::new(
        source.len() - 1,
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_resource_limit(codec::decode_text_format(source, input_limited), |limit| {
        matches!(limit, codec::DecodeLimit::InputBytes { .. })
    });

    if report.fields() > 0 {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            report.fields() - 1,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_text_format(source, limited), |limit| {
            matches!(limit, codec::DecodeLimit::Fields { .. })
        });
    }

    if report.work_bytes() > 0 {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_FIELDS,
            report.work_bytes() - 1,
            MAX_RECURSION,
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_text_format(source, limited), |limit| {
            matches!(limit, codec::DecodeLimit::Work { .. })
        });
    }

    if report.max_depth() > 0 {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            report.max_depth() - 1,
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_text_format(source, limited), |limit| {
            matches!(limit, codec::DecodeLimit::Nesting { .. })
        });
    }
}

fn assert_resource_limit<T>(
    result: Result<T, codec::DecodeError>,
    predicate: impl FnOnce(codec::DecodeLimit) -> bool,
) {
    let error = match result {
        Ok(_) => panic!("resource-limited Text-format operation unexpectedly succeeded"),
        Err(error) => error,
    };
    let limit = error
        .resource_limit()
        .expect("resource failure did not expose a typed limit");
    assert!(predicate(limit), "unexpected Text-format resource limit");
    observe_error(error);
}

fn exercise_canonical_writes() {
    let write = codec::TextFormatWrite::new();
    let prepared = codec::prepare_text_format_write(write, options(&[]))
        .unwrap_or_else(|error| panic!("canonical Text-format prepare failed: {error}"));
    let requirements = prepared.execution_requirements();
    assert_prepare_report(prepared.prepare_report(), requirements);
    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("canonical Text-format execute failed: {error}"));
    assert_eq!(output.bytes(), CANONICAL_TEXT);
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    let readback = codec::decode_text_format(output.bytes(), options(output.bytes()))
        .unwrap_or_else(|error| panic!("canonical Text-format readback failed: {error}"));
    assert_eq!(codec::TextFormatWrite::from_snapshot(readback), write);
    assert_eq!(
        codec::canonical_text_format(write, options(&[]))
            .unwrap_or_else(|error| panic!("canonical one-shot Text-format failed: {error}"))
            .bytes(),
        output.bytes()
    );
    black_box(output);
}

fn exercise_encoding_markers() {
    assert_eq!(
        codec::decode_text_format_encoding(
            codec::EXPLICIT_TEXT_FORMAT,
            codec::TEXT_CELL_FORMAT_KIND,
        ),
        Ok(codec::TextFormatEncoding::Plain)
    );
    assert_eq!(
        codec::decode_text_format_encoding(
            codec::EXPLICIT_CONVERTED_TEXT_FORMAT,
            codec::TEXT_CELL_FORMAT_KIND,
        ),
        Ok(codec::TextFormatEncoding::Converted)
    );
    assert!(!codec::TextFormatEncoding::Plain.is_converted());
    assert!(codec::TextFormatEncoding::Converted.is_converted());
    for (marker, kind) in [
        (0, codec::TEXT_CELL_FORMAT_KIND),
        (0x82, codec::TEXT_CELL_FORMAT_KIND),
        (codec::EXPLICIT_TEXT_FORMAT, 4),
    ] {
        assert!(codec::decode_text_format_encoding(marker, kind).is_err());
    }
}

fn exercise_limit_guards() {
    let write = codec::TextFormatWrite::new();
    let output_limited = options(&[]).with_max_output_bytes(0);
    assert_resource_limit(
        codec::prepare_text_format_write(write, output_limited),
        |limit| matches!(limit, codec::DecodeLimit::OutputBytes { .. }),
    );

    let prepared =
        codec::prepare_text_format_rewrite(CANONICAL_TEXT, write, options(CANONICAL_TEXT))
            .unwrap_or_else(|error| panic!("Text-format limit guard prepare failed: {error}"));
    let requirements = prepared.execution_requirements();
    assert_resource_limit(
        prepared.execute(codec::RewriteExecutionLimits::exact(requirements).with_allocations(0)),
        |limit| matches!(limit, codec::DecodeLimit::Allocation { .. }),
    );
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
