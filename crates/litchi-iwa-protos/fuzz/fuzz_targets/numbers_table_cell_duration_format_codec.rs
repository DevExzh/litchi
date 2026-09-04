#![no_main]

//! Bounded fuzzing for the native Numbers Duration `FormatStructArchive`
//! seam. The harness keeps the source-preserving contract hot while allowing
//! arbitrary bytes to exercise strict framing, known-field rejection,
//! unknown-group retention, lazy Buffa parity, and finite rewrite limits.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_duration_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const VALID: &[u8] = &[
    0x08, 0x8c, 0x02, // format_type = 268
    0x38, 0x01, // duration_style = abbreviated
    0x78, 0x04, // duration_unit_largest = hours
    0x80, 0x01, 0x20, // duration_unit_smallest = milliseconds
    0xc0, 0x02, 0x01, // use_automatic_duration_units = true
];

const VALID_WITH_UNKNOWN: &[u8] = &[
    0x08, 0x8c, 0x02, 0xa0, 0x06, 0x81, 0x00, // unknown scalar
    0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01, 0xa3, 0x06, 0x98, 0x03, 0x09, 0xa4,
    0x06, // unknown group
];

const FIXED_CASES: &[&[u8]] = &[
    VALID,
    VALID_WITH_UNKNOWN,
    // Missing selected fields.
    &[0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20],
    &[0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01],
    // Duplicate and wrong-wire selected fields.
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x38, 0x00, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x0a, 0x01, 0x01, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x3a, 0x01, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    // Every other known FormatStructArchive field belongs to another family.
    &[
        0x08, 0x8c, 0x02, 0x10, 0x00, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x1a, 0x03, b'U', b'S', b'D', 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20,
        0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0xa8, 0x01, 0x00, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02,
        0x01,
    ],
    // Invalid style, units, order, and Boolean domain.
    &[
        0x08, 0x8c, 0x02, 0x38, 0x03, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x03, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x10, 0x80, 0x01, 0x04, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x02,
    ],
    // Known selected keys and values encoded noncanonically.
    &[
        0x88, 0x00, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0xb8, 0x00, 0x81, 0x00, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    // Unterminated, mismatched, and truncated groups.
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01, 0xa3, 0x06,
        0x08, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01, 0xa3, 0x06,
        0xac, 0x06,
    ],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source);

    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed);
        }
        exercise_canonical_writes();
        exercise_limit_guard();
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

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_duration_format(source, decode_options);
    let reported = codec::decode_duration_format_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "Duration scalar decode modified source"
    );
    assert_eq!(
        source,
        before.as_slice(),
        "Duration report decode modified source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.format_type(), codec::NATIVE_DURATION_FORMAT_TYPE);
            assert!(codec::DurationStyle::from_native(snapshot.duration_style()).is_some());
            assert!(codec::DurationUnit::from_native(snapshot.duration_unit_largest()).is_some());
            assert!(codec::DurationUnit::from_native(snapshot.duration_unit_smallest()).is_some());
            assert_eq!(snapshot.raw(), source);
            assert_eq!(report.input_bytes(), source.len());
            black_box((snapshot, report));
            exercise_rewrite(source, snapshot);
        },
        (Err(left), Err(right)) => assert_eq!(
            left.resource_limit(),
            right.resource_limit(),
            "Duration scalar/report error classifications disagree"
        ),
        (left, right) => {
            panic!("Duration scalar/report disagreement: scalar={left:?}, report={right:?}")
        },
    }
    assert_eq!(source, before.as_slice(), "Duration probes modified source");
}

fn exercise_rewrite(source: &[u8], snapshot: codec::DurationFormatSnapshot<'_>) {
    let write = codec::DurationFormatWrite::from_snapshot(snapshot)
        .with_style(codec::DurationStyle::FullNames)
        .with_largest_unit(codec::DurationUnit::Minutes)
        .with_smallest_unit(codec::DurationUnit::Seconds)
        .with_automatic(false);
    let prepared = codec::prepare_duration_format_rewrite(source, write, options(source));
    let Ok(prepared) = prepared else {
        return;
    };
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .expect("valid Duration rewrite should execute");
    let rewritten = codec::decode_duration_format(output.bytes(), options(output.bytes()))
        .expect("rewritten Duration should read back");
    assert_eq!(rewritten.style(), Some(codec::DurationStyle::FullNames));
    assert_eq!(rewritten.largest_unit(), Some(codec::DurationUnit::Minutes));
    assert_eq!(
        rewritten.smallest_unit(),
        Some(codec::DurationUnit::Seconds)
    );
    assert!(!rewritten.is_automatic());
    assert_eq!(output.report().fields(), requirements.fields());
    assert_eq!(output.report().work_bytes(), requirements.work_bytes());
    assert!(output.bytes().len() <= MAX_OUTPUT_BYTES.max(source.len()));
}

fn exercise_canonical_writes() {
    for style in [
        codec::DurationStyle::Colon,
        codec::DurationStyle::Abbreviated,
        codec::DurationStyle::FullNames,
    ] {
        let write = codec::DurationFormatWrite::from_parts(
            style,
            codec::DurationUnit::Weeks,
            codec::DurationUnit::Milliseconds,
            true,
        );
        let output = codec::canonical_duration_format(write, options(&[])).expect("canonical");
        let snapshot = codec::decode_duration_format(output.bytes(), options(output.bytes()))
            .expect("canonical readback");
        assert_eq!(codec::DurationFormatWrite::from_snapshot(snapshot), write);
        black_box(output);
    }
}

fn exercise_limit_guard() {
    let mut limited = options(VALID);
    limited = limited.with_max_output_bytes(0);
    let write = codec::DurationFormatWrite::from_parts(
        codec::DurationStyle::Colon,
        codec::DurationUnit::Weeks,
        codec::DurationUnit::Milliseconds,
        true,
    );
    let error = codec::prepare_duration_format_rewrite(VALID, write, limited)
        .expect_err("output ceiling")
        .resource_limit();
    assert!(matches!(
        error,
        Some(codec::DecodeLimit::OutputBytes { .. })
    ));

    let prepared = codec::prepare_duration_format_rewrite(VALID, write, options(VALID))
        .expect("limit guard prepare");
    let requirements = prepared.execution_requirements();
    let error = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements).with_allocations(0))
        .expect_err("allocation ceiling")
        .resource_limit();
    assert!(matches!(error, Some(codec::DecodeLimit::Allocation { .. })));
}
