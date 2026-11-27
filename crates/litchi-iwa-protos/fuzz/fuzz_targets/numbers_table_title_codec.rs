#![no_main]

//! Bounded fuzzing for the strict Numbers table-title projection.
//!
//! The target treats each input as caller-owned protobuf wire data.  It keeps
//! the generated Buffa view behind the codec's public snapshot, compares the
//! scalar and reported entry points, and runs deterministic wire recipes so
//! all proto2 presence states and malformed boundaries remain reachable.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_title_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_REFERENCES: usize = 2;

const TABLE_NAME_ENABLED_FIELD: u32 = 22;
const TABLE_NAME_STYLE_FIELD: u32 = 30;
const TABLE_NAME_HEIGHT_FIELD: u32 = 33;
const TABLE_NAME_SHAPE_STYLE_FIELD: u32 = 36;
const TABLE_NAME_BORDER_ENABLED_FIELD: u32 = 37;

// Tiny hand-authored protobuf recipes keep canonical presence, nested
// references, unknown framing, and every selected malformed route hot even
// when arbitrary bytes do not form a valid table-title envelope.
const FIXED_CASES: &[&[u8]] = &[
    // Empty envelope: every proto2 field is absent.
    &[],
    // Explicit false values retain presence independently from absence.
    &[0xb0, 0x01, 0x00, 0xa8, 0x02, 0x00],
    // Explicit true values and an IEEE-754 negative-zero height.
    &[
        0xb0, 0x01, 0x01, 0x89, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80, 0xa8, 0x02,
        0x01,
    ],
    // NaN height plus a visible title style and an outlined shape style.  The
    // references include the deprecated signed and boolean scalar fields.
    &[
        0xb0, 0x01, 0x01, 0xf0, 0x01, 0x0f, 0x08, 0x29, 0x10, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
        0xff, 0xff, 0xff, 0xff, 0x01, 0x18, 0x01, 0x89, 0x02, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0xf8, 0x7f, 0xa2, 0x02, 0x0f, 0x08, 0x2a, 0x10, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
        0xff, 0xff, 0xff, 0xff, 0x01, 0x18, 0x00, 0xa8, 0x02, 0x00,
    ],
    // All unknown wire kinds and balanced nested groups surround selected
    // fields.  Unknown spans remain caller-owned; the report still counts
    // every framing record and reaches depth three.
    &[
        0xa0, 0x06, 0x07, // unknown varint field 100
        0xa9, 0x06, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, // fixed64 101
        0xb2, 0x06, 0x02, 0xff, 0x00, // bytes 102
        0xbd, 0x06, 0x09, 0x08, 0x07, 0x06, // fixed32 103
        0xa3, 0x06, 0x08, 0x09, 0xab, 0x06, 0x10, 0x0a, 0xac, 0x06, 0xa4,
        0x06, // groups 104/105
        0xb0, 0x01, 0x01,
    ],
    // Both style references and all scalar fields in a single valid payload.
    &[
        0xb0, 0x01, 0x00, 0xf0, 0x01, 0x02, 0x08, 0x01, 0x89, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0xa2, 0x02, 0x02, 0x08, 0x02, 0xa8, 0x02, 0x01,
    ],
    // Duplicate selected scalar.
    &[0xb0, 0x01, 0x00, 0xb0, 0x01, 0x01],
    // Duplicate selected reference.
    &[0xf0, 0x01, 0x02, 0x08, 0x01, 0xf0, 0x01, 0x02, 0x08, 0x02],
    // Every selected field with an incompatible wire type.
    &[
        0xb1, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // field 22 fixed64
        0x88, 0x02, 0x00, // field 33 varint
        0xf0, 0x01, 0x01, // field 30 varint
        0xa5, 0x02, 0x00, 0x00, 0x00, 0x00, // field 36 fixed32
        0xaa, 0x02, 0x01, 0x00, // field 37 bytes
    ],
    // Non-canonical selected key/value varints.
    &[0xb0, 0x81, 0x00, 0x01],
    &[0xb0, 0x01, 0x80, 0x00],
    // Non-canonical reference key, identifier, and legacy scalar values.
    &[0xf0, 0x01, 0x03, 0x88, 0x00, 0x01],
    &[0xf0, 0x01, 0x03, 0x08, 0x80, 0x00],
    &[0xf0, 0x01, 0x05, 0x08, 0x01, 0x10, 0x80, 0x80, 0x80, 0x04],
    &[0xf0, 0x01, 0x04, 0x08, 0x01, 0x18, 0x02],
    // Missing/zero required identifiers and truncated selected envelopes.
    &[0xf0, 0x01, 0x00],
    &[0xf0, 0x01, 0x02, 0x08, 0x00],
    &[0xf0, 0x01, 0x05, 0x08, 0x01],
    &[0xf0, 0x01, 0x05, 0x08, 0x01, 0x10],
    &[0x89, 0x02, 0x01, 0x02],
    &[0xb0, 0x01, 0x80],
    // Unterminated, mismatched, and unexpected group ends.
    &[0xa3, 0x06, 0x08, 0x01],
    &[0xa3, 0x06, 0x08, 0x01, 0xab, 0x06],
    &[0xa4, 0x06],
    // Invalid field number and truncated tag.
    &[0x00],
    &[0x80],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source);

    // Keep deterministic valid, malformed, and finite-limit cases in every
    // process without adding their bytes to a mutable libFuzzer corpus.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed);
        }
        exercise_scalar_presence_states();
        exercise_limit_boundaries();
        exercise_deep_groups();
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
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_table_title_settings(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "scalar decode modified its source"
    );
    let reported = codec::decode_table_title_settings_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "reported decode modified its source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            assert!(report.references() <= MAX_REFERENCES);
            assert!(report.reference_bytes() <= source.len());
            assert_snapshot_is_canonical(snapshot);
            black_box((snapshot, report));
        },
        (Err(scalar_error), Err(reported_error)) => {
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "table-title scalar/report decode disagreement: scalar_ok={}, report_ok={}",
            scalar_result.is_ok(),
            reported_result.is_ok()
        ),
    }
    assert_eq!(
        source,
        before.as_slice(),
        "decode probes modified its source"
    );
}

fn assert_snapshot_is_canonical(snapshot: codec::TableTitleSettingsSnapshot) {
    if let Some(reference) = snapshot.table_name_style() {
        assert!(reference.identifier() > 0);
        black_box((
            reference.identifier(),
            reference.deprecated_type(),
            reference.deprecated_is_external(),
        ));
    }
    if let Some(reference) = snapshot.table_name_shape_style() {
        assert!(reference.identifier() > 0);
        black_box((
            reference.identifier(),
            reference.deprecated_type(),
            reference.deprecated_is_external(),
        ));
    }
    // Height is intentionally compared as bits by the codec.  Touch the
    // complete value so NaN, infinities, and negative zero remain observable.
    black_box((
        snapshot.table_name_enabled(),
        snapshot.table_name_height_bits(),
        snapshot.table_name_border_enabled(),
    ));
}

fn exercise_scalar_presence_states() {
    for enabled in [None, Some(false), Some(true)] {
        for outlined in [None, Some(false), Some(true)] {
            for height in [
                None,
                Some(0.0f64.to_bits()),
                Some((-0.0f64).to_bits()),
                Some(f64::NAN.to_bits()),
                Some(f64::INFINITY.to_bits()),
                Some(f64::NEG_INFINITY.to_bits()),
                Some(0x0123_4567_89ab_cdef),
            ] {
                let source = title_payload(enabled, outlined, height, None, None);
                let snapshot = codec::decode_table_title_settings(&source, options(&source))
                    .unwrap_or_else(|error| panic!("canonical scalar state rejected: {error}"));
                assert_eq!(snapshot.table_name_enabled(), enabled);
                assert_eq!(snapshot.table_name_border_enabled(), outlined);
                assert_eq!(snapshot.table_name_height_bits(), height);
                assert_eq!(snapshot.table_name_style(), None);
                assert_eq!(snapshot.table_name_shape_style(), None);
            }
        }
    }
}

fn exercise_limit_boundaries() {
    let paragraph = reference_payload(41, Some(-1), Some(true));
    let shape = reference_payload(42, Some(0), Some(false));
    let source = title_payload(
        Some(true),
        Some(false),
        Some((-0.0f64).to_bits()),
        Some(&paragraph),
        Some(&shape),
    );
    let before = source.clone();
    let (_, report) = codec::decode_table_title_settings_with_report(&source, options(&source))
        .expect("rich canonical title must decode");
    assert!(report.fields() > 0);
    assert!(report.work_bytes() > 0);
    assert!(report.max_depth() > 0);
    assert!(report.references() > 0);

    let exact = codec::DecodeOptions::new(
        source.len(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
        report.references(),
    );
    let (_, exact_report) = codec::decode_table_title_settings_with_report(&source, exact)
        .expect("exact title limits must be inclusive");
    assert_eq!(exact_report, report);

    let probes = [
        (
            "bytes",
            codec::DecodeOptions::new(
                source.len() - 1,
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            ),
            codec::DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            },
        ),
        (
            "fields",
            codec::DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            ),
            codec::DecodeLimit::Fields {
                observed: report.fields(),
                maximum: report.fields() - 1,
            },
        ),
        (
            "work",
            codec::DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work_bytes() - 1,
                report.max_depth(),
                report.references(),
            ),
            codec::DecodeLimit::Work {
                observed: report.work_bytes(),
                maximum: report.work_bytes() - 1,
            },
        ),
        (
            "nesting",
            codec::DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work_bytes(),
                report.max_depth() - 1,
                report.references(),
            ),
            codec::DecodeLimit::Nesting {
                observed: report.max_depth(),
                maximum: report.max_depth() - 1,
            },
        ),
        (
            "references",
            codec::DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references() - 1,
            ),
            codec::DecodeLimit::References {
                observed: report.references(),
                maximum: report.references() - 1,
            },
        ),
    ];
    for (name, limited, expected) in probes {
        let error = match codec::decode_table_title_settings(&source, limited) {
            Ok(_) => panic!("{name} max-minus-one title limit succeeded"),
            Err(error) => error,
        };
        assert_eq!(error.resource_limit(), Some(expected), "{name} limit");
        observe_error(error);
    }
    assert_eq!(source, before.as_slice(), "limit probes changed source");
}

fn exercise_deep_groups() {
    let mut source = Vec::new();
    for _ in 0..(MAX_RECURSION as usize + 1) {
        push_key(&mut source, 100, 3);
    }
    push_varint_field(&mut source, TABLE_NAME_ENABLED_FIELD, 1);
    for _ in 0..(MAX_RECURSION as usize + 1) {
        push_key(&mut source, 100, 4);
    }
    let before = source.clone();
    let result = codec::decode_table_title_settings(&source, options(&source));
    assert!(result.is_err(), "deep title group bypassed nesting limit");
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(source, before, "deep-group probe modified source");
}

fn title_payload(
    enabled: Option<bool>,
    outlined: Option<bool>,
    height_bits: Option<u64>,
    paragraph: Option<&[u8]>,
    shape: Option<&[u8]>,
) -> Vec<u8> {
    let mut source = Vec::new();
    if let Some(value) = enabled {
        push_varint_field(&mut source, TABLE_NAME_ENABLED_FIELD, u64::from(value));
    }
    if let Some(reference) = paragraph {
        push_length_field(&mut source, TABLE_NAME_STYLE_FIELD, reference);
    }
    if let Some(bits) = height_bits {
        push_fixed64_field(&mut source, TABLE_NAME_HEIGHT_FIELD, bits);
    }
    if let Some(reference) = shape {
        push_length_field(&mut source, TABLE_NAME_SHAPE_STYLE_FIELD, reference);
    }
    if let Some(value) = outlined {
        push_varint_field(
            &mut source,
            TABLE_NAME_BORDER_ENABLED_FIELD,
            u64::from(value),
        );
    }
    source
}

fn reference_payload(
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
) -> Vec<u8> {
    let mut source = Vec::new();
    push_varint_field(&mut source, 1, identifier);
    if let Some(value) = deprecated_type {
        let encoded = if value < 0 {
            u64::from_ne_bytes(i64::from(value).to_ne_bytes())
        } else {
            u64::try_from(value).expect("non-negative i32 fits u64")
        };
        push_varint_field(&mut source, 2, encoded);
    }
    if let Some(value) = deprecated_is_external {
        push_varint_field(&mut source, 3, u64::from(value));
    }
    source
}

fn push_key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    push_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_key(output, number, 0);
    push_varint(output, value);
}

fn push_fixed64_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_key(output, number, 1);
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_length_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_key(output, number, 2);
    push_varint(
        output,
        u64::try_from(payload.len()).expect("bounded payload length fits u64"),
    );
    output.extend_from_slice(payload);
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = u8::try_from(value & 0x7f).expect("masked varint byte fits u8");
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.resource_limit());
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
