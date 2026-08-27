#![no_main]

//! Bounded fuzzing for the strict Numbers table-header settings projection.
//!
//! Inputs are complete `TST.TableModelArchive`-style wire payloads containing
//! the required row/column dimensions.  The target keeps the source bytes
//! authoritative, exercises all seven optional settings, and replays every
//! prepared execution ceiling before allowing the candidate allocation.

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_header_settings_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_ALLOCATIONS: usize = 16;
const MAX_RETAINED_BYTES: usize = 128 * 1024;
const MAX_SCRATCH_BYTES: usize = 512 * 1024;

const UNKNOWN_GROUP: &[u8] = &[0xa3, 0x06, 0x08, 0x01, 0xa4, 0x06];

// These small wire recipes are independent of any native package.  They keep
// required dimensions and the strict malformed/unknown cases reachable even
// when arbitrary bytes do not happen to form a valid protobuf envelope.
const FIXED_CASES: &[&[u8]] = &[
    // Required dimensions only.
    &[0x30, 0x01, 0x38, 0x01],
    // Missing one required dimension.
    &[0x30, 0x01],
    // All seven optional fields, including explicit false values.
    &[
        0x30, 0x01, 0x38, 0x02, 0x48, 0x01, 0x50, 0x01, 0x58, 0x01, 0x60, 0x01, 0x68, 0x00, 0xe8,
        0x01, 0x01, 0x80, 0x02, 0x00,
    ],
    // Canonical unknown scalar and a balanced unknown group survive rewrites.
    &[
        0x30, 0x01, 0x38, 0x01, 0xa0, 0x06, 0x81, 0x01, 0xa3, 0x06, 0x08, 0x01, 0xa4, 0x06,
    ],
    // Unknown scalar, fixed-width values, bytes, and a balanced group.
    &[
        0x30, 0x01, 0x38, 0x01, 0xa0, 0x06, 0x81, 0x00, 0xad, 0x06, 0x00, 0x00, 0x00, 0x00, 0xb2,
        0x06, 0x02, 0x09, 0x08, 0xa3, 0x06, 0x08, 0x01, 0xa4, 0x06,
    ],
    // Duplicate required dimensions.
    &[0x30, 0x01, 0x30, 0x02, 0x38, 0x01],
    // Duplicate selected optional field.
    &[0x30, 0x01, 0x38, 0x01, 0x48, 0x01, 0x48, 0x02],
    // Wrong wire types for a required and a selected field.
    &[0x32, 0x01, 0x01, 0x38, 0x01],
    &[0x30, 0x01, 0x38, 0x01, 0x4d, 0x00, 0x00, 0x00, 0x00],
    // Non-canonical known key and value encodings.
    &[0xb0, 0x80, 0x00, 0x01, 0x38, 0x01],
    &[0x30, 0x81, 0x00, 0x38, 0x01],
    // A non-canonical boolean value.
    &[0x30, 0x01, 0x38, 0x01, 0x60, 0x02],
    // Truncated varint and unterminated group.
    &[0x30, 0x01, 0x38, 0x01, 0x48],
    &[0x30, 0x01, 0x38, 0x01, 0xa3, 0x06, 0x08, 0x01],
];

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data) {
        exercise_source(&source, data);
    }

    static FIXED: OnceLock<()> = OnceLock::new();
    FIXED.get_or_init(|| {
        for source in FIXED_CASES {
            exercise_source(source, b"fixed-table-header-settings");
        }
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
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
    .with_max_allocations(MAX_ALLOCATIONS)
    .with_max_retained_bytes(MAX_RETAINED_BYTES.max(source.len()))
    .with_max_scratch_bytes(MAX_SCRATCH_BYTES.max(source.len()))
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let decoded = codec::decode_table_header_settings(source, decode_options);
    assert_eq!(
        source,
        original.as_slice(),
        "header decode modified its source"
    );

    let snapshot = match decoded {
        Ok(snapshot) => snapshot,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    black_box((
        snapshot.rows(),
        snapshot.columns(),
        snapshot.header_rows(),
        snapshot.header_columns(),
        snapshot.footer_rows(),
        snapshot.header_rows_frozen(),
        snapshot.header_columns_frozen(),
        snapshot.repeating_header_rows_enabled(),
        snapshot.repeating_header_columns_enabled(),
    ));

    let desired = desired_write(data);
    let prepared =
        match codec::prepare_table_header_settings_rewrite(source, desired, decode_options) {
            Ok(prepared) => prepared,
            Err(error) => {
                observe_error(error);
                return;
            },
        };
    let requirements = prepared.execution_requirements();
    let preparation = prepared.prepare_report();
    assert_eq!(preparation.input_bytes(), source.len());
    assert_eq!(preparation.output_bytes(), requirements.output_bytes);
    assert_eq!(preparation.fields(), requirements.fields);
    assert_eq!(preparation.work_bytes(), requirements.work_bytes);
    assert_eq!(preparation.max_depth(), requirements.max_depth);
    assert_eq!(preparation.allocations(), requirements.allocations);
    assert_eq!(preparation.retained_bytes(), requirements.retained_bytes);
    assert_eq!(preparation.scratch_bytes(), requirements.scratch_bytes);
    assert_eq!(
        source,
        original.as_slice(),
        "header prepare modified its source"
    );

    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact header prepared replay failed: {error}"));
    assert_eq!(output.bytes().len(), requirements.output_bytes);
    assert_eq!(output.report(), preparation);
    let candidate = output.bytes().to_vec();
    assert_eq!(
        source,
        original.as_slice(),
        "header execute modified its source"
    );

    let readback = codec::decode_table_header_settings(&candidate, options(&candidate))
        .unwrap_or_else(|error| panic!("header candidate readback failed: {error}"));
    assert_eq!(readback.optional_write(), desired);
    let one_shot = codec::rewrite_table_header_settings(source, desired, decode_options)
        .unwrap_or_else(|error| panic!("header one-shot rewrite failed: {error}"));
    assert_eq!(one_shot, candidate);
    assert_eq!(
        source,
        original.as_slice(),
        "header rewrite modified its source"
    );

    exercise_limit_failures(source, desired, requirements, &original);
}

fn desired_write(data: &[u8]) -> codec::TableHeaderSettingsWrite {
    let byte = |offset: usize| data.get(offset).copied().unwrap_or_default();
    codec::TableHeaderSettingsWrite::new(
        Some(u32::from(byte(0) % 5)),
        if byte(1) & 1 == 0 {
            Some(u32::from(byte(1) % 5))
        } else {
            None
        },
        if byte(2) & 1 == 0 {
            Some(u32::from(byte(2) % 5))
        } else {
            None
        },
        Some(byte(3) & 1 != 0),
        if byte(4) & 1 == 0 {
            Some(byte(4) & 2 != 0)
        } else {
            None
        },
        Some(byte(5) & 1 != 0),
        if byte(6) & 1 == 0 {
            Some(byte(6) & 2 != 0)
        } else {
            None
        },
    )
}

fn exercise_limit_failures(
    source: &[u8],
    desired: codec::TableHeaderSettingsWrite,
    requirements: codec::RewriteExecutionRequirements,
    original: &[u8],
) {
    // Input admission is the one axis that belongs to DecodeOptions rather
    // than RewriteExecutionRequirements; replay it independently before the
    // prepared execution ceilings.
    let input_limited = codec::DecodeOptions::new(
        source.len().saturating_sub(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES)
    .with_max_allocations(MAX_ALLOCATIONS)
    .with_max_retained_bytes(MAX_RETAINED_BYTES)
    .with_max_scratch_bytes(MAX_SCRATCH_BYTES);
    assert!(
        codec::decode_table_header_settings(source, input_limited).is_err(),
        "a max-minus-one header input ceiling succeeded"
    );
    assert!(
        codec::prepare_table_header_settings_rewrite(source, desired, input_limited).is_err(),
        "a max-minus-one header input ceiling prepared a rewrite"
    );
    assert_eq!(
        source, original,
        "failed input admission modified its source"
    );

    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let limits = [
        exact.with_output_bytes(requirements.output_bytes.saturating_sub(1)),
        exact.with_fields(requirements.fields.saturating_sub(1)),
        exact.with_work_bytes(requirements.work_bytes.saturating_sub(1)),
        exact.with_max_depth(requirements.max_depth.saturating_sub(1)),
        exact.with_allocations(requirements.allocations.saturating_sub(1)),
        exact.with_retained_bytes(requirements.retained_bytes.saturating_sub(1)),
        exact.with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1)),
    ];
    for limits in limits {
        let prepared =
            codec::prepare_table_header_settings_rewrite(source, desired, options(source));
        let Ok(prepared) = prepared else {
            continue;
        };
        let result = prepared.execute(limits);
        assert!(result.is_err(), "a max-minus-one header ceiling succeeded");
        black_box(result.err());
        assert_eq!(
            source, original,
            "failed header execute modified its source"
        );
    }
}

fn exercise_deep_groups() {
    let mut source = vec![0x30, 0x01, 0x38, 0x01];
    for _ in 0..80 {
        source.extend_from_slice(&UNKNOWN_GROUP[..2]);
    }
    source.extend_from_slice(&[0x08, 0x01]);
    for _ in 0..80 {
        source.extend_from_slice(&[0xa4, 0x06]);
    }
    let result = codec::decode_table_header_settings(&source, options(&source));
    assert!(
        result.is_err(),
        "deep unknown groups bypassed nesting limit"
    );
    black_box(result.err());
}

fn observe_error(error: impl std::fmt::Debug + std::fmt::Display) {
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
