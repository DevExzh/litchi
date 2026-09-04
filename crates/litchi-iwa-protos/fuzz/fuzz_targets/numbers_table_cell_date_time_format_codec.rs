#![no_main]

//! Bounded fuzzing for the native Numbers DateTime `FormatStructArchive`
//! seam. The harness keeps the source-preserving contract hot while allowing
//! arbitrary bytes to exercise strict framing, UTF-8, known-field rejection,
//! unknown-group retention, Buffa parity, and finite rewrite limits.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_date_time_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const VALID: &[u8] = &[
    0x08, 0x85, 0x02, // format_type = 261
    0x72, 0x0a, b'y', b'y', b'y', b'y', b'-', b'M', b'M', b'-', b'd', b'd',
];

const VALID_WITH_UNKNOWN: &[u8] = &[
    0x08, 0x85, 0x02, 0xa0, 0x06, 0x81, 0x00, // unknown scalar
    0x72, 0x0a, b'y', b'y', b'y', b'y', b'-', b'M', b'M', b'-', b'd', b'd', 0xa3, 0x06, 0x98, 0x03,
    0x09, 0xa4, 0x06, // unknown balanced group
];

const FIXED_CASES: &[&[u8]] = &[
    VALID,
    VALID_WITH_UNKNOWN,
    // Missing selected fields.
    &[0x08, 0x85, 0x02],
    &[
        0x72, 0x0a, b'y', b'y', b'y', b'y', b'-', b'M', b'M', b'-', b'd', b'd',
    ],
    // Duplicate and wrong-wire selected fields.
    &[
        0x08, 0x85, 0x02, 0x72, 0x0a, b'y', b'y', b'y', b'y', b'-', b'M', b'M', b'-', b'd', b'd',
        0x72, 0x02, b'X', b'X',
    ],
    &[0x0a, 0x01, 0x01, 0x72, 0x01, b'x'],
    &[0x08, 0x85, 0x02, 0x70, 0x01, 0x72, 0x01, b'x'],
    // Every other known FormatStructArchive field is another family.
    &[0x08, 0x85, 0x02, 0x10, 0x00, 0x72, 0x01, b'x'],
    &[
        0x08, 0x85, 0x02, 0x1a, 0x03, b'U', b'S', b'D', 0x72, 0x01, b'x',
    ],
    &[0x08, 0x85, 0x02, 0xa8, 0x01, 0x00, 0x72, 0x01, b'x'],
    // Invalid pattern values and malformed groups.
    &[0x08, 0x85, 0x02, 0x72, 0x00],
    &[0x08, 0x85, 0x02, 0x72, 0x01, b'\n'],
    &[0x08, 0x85, 0x02, 0x72, 0x02, 0xff, 0xff],
    &[0x08, 0x85, 0x02, 0xa3, 0x06, 0x08, 0x01],
    &[0x08, 0x85, 0x02, 0xa4, 0x06],
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
        exercise_canonical_write();
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
        codec::MAX_DATE_TIME_PATTERN_BYTES,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_date_time_format(source, decode_options);
    let reported = codec::decode_date_time_format_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "DateTime scalar decode modified source"
    );
    assert_eq!(
        source,
        before.as_slice(),
        "DateTime report decode modified source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.format_type(), codec::NATIVE_DATE_TIME_FORMAT_TYPE);
            assert_eq!(snapshot.raw(), source);
            assert!(snapshot.pattern().len() <= codec::MAX_DATE_TIME_PATTERN_BYTES);
            assert!(!snapshot.pattern().trim().is_empty());
            assert!(!snapshot.pattern().chars().any(char::is_control));
            assert_borrowed(source, snapshot.pattern().as_bytes());
            assert_eq!(report.input_bytes(), source.len());
            black_box((snapshot, report));
            exercise_rewrite(source, snapshot);
        },
        (Err(left), Err(right)) => assert_eq!(
            left.resource_limit(),
            right.resource_limit(),
            "DateTime scalar/report error classifications disagree"
        ),
        (left, right) => {
            panic!("DateTime scalar/report disagreement: scalar={left:?}, report={right:?}")
        },
    }
    assert_eq!(source, before.as_slice(), "DateTime probes modified source");
}

fn exercise_rewrite(source: &[u8], snapshot: codec::DateTimeFormatSnapshot<'_>) {
    let write = codec::DateTimeFormatWrite::from_snapshot(snapshot).with_pattern("MM/dd/yyyy");
    let prepared = codec::prepare_date_time_format_rewrite(source, write, options(source))
        .expect("valid DateTime source should prepare");
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .expect("valid DateTime rewrite should execute");
    let rewritten = codec::decode_date_time_format(output.bytes(), options(output.bytes()))
        .expect("rewritten DateTime should read back");
    assert_eq!(rewritten.pattern(), "MM/dd/yyyy");
    assert_eq!(output.report().fields(), requirements.fields());
    assert_eq!(output.report().work_bytes(), requirements.work_bytes());
    assert!(output.bytes().len() <= MAX_OUTPUT_BYTES);
}

fn exercise_canonical_write() {
    let write = codec::DateTimeFormatWrite::new("yyyy-MM-dd H:mm:ss");
    let output = codec::canonical_date_time_format(write, options(VALID)).expect("canonical");
    let snapshot = codec::decode_date_time_format(output.bytes(), options(output.bytes()))
        .expect("canonical readback");
    assert_eq!(snapshot.pattern(), write.pattern());
    assert_eq!(output.report().fields(), 2);
}

fn exercise_limit_guard() {
    let mut limited = options(VALID);
    limited = limited.with_max_text_bytes(4);
    let error = codec::decode_date_time_format(VALID, limited).expect_err("text ceiling");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::Text { .. })
    ));
}

fn assert_borrowed(source: &[u8], value: &[u8]) {
    let source_start = source.as_ptr() as usize;
    let source_end = source_start.saturating_add(source.len());
    let value_start = value.as_ptr() as usize;
    let value_end = value_start.saturating_add(value.len());
    assert!(value_start >= source_start && value_end <= source_end);
}
