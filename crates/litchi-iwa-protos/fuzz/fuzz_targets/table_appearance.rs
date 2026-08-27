#![no_main]

//! Strict table-style/model/stylesheet wire fuzzing.
//!
//! The target keeps every codec pass on bounded caller-owned payloads.  A
//! successful prepared rewrite is executed with its exact requirements and
//! decoded again, which exercises source preservation, candidate verification,
//! unknown-field retention, and the typed output/field/work ceilings.

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::table_appearance_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_STYLES: usize = 4 * 1024;

const FIXED_CASES: &[&[u8]] = &[
    // TableModelArchive.table_style = Reference { identifier: 7 }.
    &[0x1a, 0x02, 0x08, 0x07],
    // TableStyleArchive.super = parent 1, variation=true, stylesheet 8;
    // table properties set banding and automatic sizing.
    &[
        0x0a, 0x0c, 0x1a, 0x02, 0x08, 0x01, 0x20, 0x01, 0x2a, 0x02, 0x08, 0x08, 0x5a, 0x05, 0x08,
        0x01, 0xb0, 0x01, 0x01,
    ],
    // StylesheetArchive.styles = Reference { identifier: 7 }.
    &[0x0a, 0x02, 0x08, 0x07],
    // Unknown overlong scalar and balanced unknown group around a style edge.
    &[
        0x1a, 0x0c, 0x08, 0x07, 0x98, 0x03, 0x81, 0x80, 0x00, 0x93, 0x03, 0x08, 0x01, 0x9c, 0x03,
    ],
    // Duplicate known style edge.
    &[0x1a, 0x02, 0x08, 0x07, 0x1a, 0x02, 0x08, 0x08],
    // Wrong wire and non-canonical known reference scalar.
    &[0x1a, 0x03, 0x0d, 0x07, 0x00, 0x00, 0x00],
    &[0x1a, 0x03, 0x08, 0x87, 0x00],
    // Unterminated unknown group.
    &[0x1a, 0x06, 0x08, 0x07, 0x93, 0x03, 0x08, 0x01],
    // TableStylePresetArchive: index, image, style-network, and a balanced
    // unknown group.  The unknown group is intentionally outside the known
    // three-field projection and remains source-authoritative.
    &[
        0x08, 0x01, 0x12, 0x02, 0x08, 0x07, 0x1a, 0x02, 0x08, 0x08, 0xd3, 0x05, 0xd8, 0x05, 0x01,
        0xd4, 0x05,
    ],
    // TableStyleNetworkArchive: all nine required references plus a balanced
    // unknown group.
    &[
        0x0a, 0x02, 0x08, 0x01, 0x12, 0x02, 0x08, 0x02, 0x1a, 0x02, 0x08, 0x03, 0x22, 0x02, 0x08,
        0x04, 0x2a, 0x02, 0x08, 0x05, 0x32, 0x02, 0x08, 0x06, 0x3a, 0x02, 0x08, 0x07, 0x42, 0x02,
        0x08, 0x08, 0x4a, 0x02, 0x08, 0x09, 0xd3, 0x05, 0xd8, 0x05, 0x01, 0xd4, 0x05,
    ],
    // Duplicate known preset network reference.
    &[
        0x08, 0x01, 0x12, 0x02, 0x08, 0x07, 0x1a, 0x02, 0x08, 0x08, 0x1a, 0x02, 0x08, 0x09,
    ],
    // Wrong wire for the known preset image reference.
    &[0x08, 0x01, 0x10, 0x07, 0x1a, 0x02, 0x08, 0x08],
    // Non-canonical known preset network reference identifier.
    &[
        0x08, 0x01, 0x12, 0x02, 0x08, 0x07, 0x1a, 0x03, 0x08, 0x87, 0x00,
    ],
    // Duplicate required network reference.
    &[
        0x0a, 0x02, 0x08, 0x01, 0x12, 0x02, 0x08, 0x02, 0x1a, 0x02, 0x08, 0x03, 0x22, 0x02, 0x08,
        0x04, 0x2a, 0x02, 0x08, 0x05, 0x32, 0x02, 0x08, 0x06, 0x3a, 0x02, 0x08, 0x07, 0x42, 0x02,
        0x08, 0x08, 0x4a, 0x02, 0x08, 0x09, 0x0a, 0x02, 0x08, 0x0a,
    ],
    // Wrong wire for a required network reference.
    &[0x08, 0x01],
    // Non-canonical required network reference identifier.
    &[
        0x0a, 0x03, 0x08, 0x81, 0x00, 0x12, 0x02, 0x08, 0x02, 0x1a, 0x02, 0x08, 0x03, 0x22, 0x02,
        0x08, 0x04, 0x2a, 0x02, 0x08, 0x05, 0x32, 0x02, 0x08, 0x06, 0x3a, 0x02, 0x08, 0x07, 0x42,
        0x02, 0x08, 0x08, 0x4a, 0x02, 0x08, 0x09,
    ],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-table-appearance");
        }
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
        MAX_STYLES,
    )
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);

    match codec::decode_table_model_with_report(source, decode_options) {
        Ok((snapshot, report)) => {
            assert_eq!(source, before.as_slice(), "model decode modified source");
            assert_eq!(snapshot.raw(), source);
            assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source.len()));
            exercise_model_rewrite(source, snapshot.style_identifier(), data, &before);
        },
        Err(error) => observe_error(error),
    }

    match codec::decode_table_style_with_report(source, decode_options) {
        Ok((snapshot, report)) => {
            assert_eq!(source, before.as_slice(), "style decode modified source");
            assert_eq!(snapshot.raw(), source);
            black_box((
                snapshot.style_identifier(),
                snapshot.parent_identifier(),
                snapshot.stylesheet_identifier(),
                snapshot.is_variation(),
                snapshot.overrides(),
                report,
            ));
            exercise_style_inheritance(snapshot, data);
            exercise_variation(data);
        },
        Err(error) => observe_error(error),
    }

    exercise_style_preset(source, &before);
    exercise_style_network(source, &before);

    match codec::decode_stylesheet_with_report(source, decode_options) {
        Ok((snapshot, report)) => {
            assert_eq!(
                source,
                before.as_slice(),
                "stylesheet decode modified source"
            );
            assert_eq!(snapshot.raw(), source);
            black_box((snapshot.style_count(), snapshot.parent_identifier(), report));
            exercise_stylesheet_append(source, data, &before);
        },
        Err(error) => observe_error(error),
    }

    exercise_variation(data);
    assert_eq!(
        source,
        before.as_slice(),
        "appearance fuzzing modified source"
    );
}

fn exercise_style_preset(source: &[u8], before: &[u8]) {
    let decode_options = options(source);
    match codec::decode_table_style_preset_with_report(source, decode_options) {
        Ok((snapshot, report)) => {
            assert_eq!(source, before, "preset decode modified source");
            assert_eq!(snapshot.raw(), source);
            assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source.len()));
            let replay = codec::decode_table_style_preset(source, decode_options)
                .unwrap_or_else(|error| panic!("preset replay failed: {error}"));
            assert_eq!(replay, snapshot);
            black_box((snapshot.style_network_identifier(), report));
            exercise_preset_limits(source, report);
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_style_network(source: &[u8], before: &[u8]) {
    let decode_options = options(source);
    match codec::decode_table_style_network_with_report(source, decode_options) {
        Ok((snapshot, report)) => {
            assert_eq!(source, before, "network decode modified source");
            assert_eq!(snapshot.raw(), source);
            assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source.len()));
            let replay = codec::decode_table_style_network(source, decode_options)
                .unwrap_or_else(|error| panic!("network replay failed: {error}"));
            assert_eq!(replay, snapshot);
            black_box((snapshot.table_style_identifier(), report));
            exercise_network_limits(source, report);
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_preset_limits(source: &[u8], report: codec::DecodeReport) {
    let fields = report.fields().saturating_sub(1);
    let error = codec::decode_table_style_preset_with_report(
        source,
        options(source).with_max_fields(fields),
    )
    .expect_err("preset fields max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::Fields { .. })
    ));

    let work = report.work_bytes().saturating_sub(1);
    let error = codec::decode_table_style_preset_with_report(
        source,
        options(source).with_max_work_bytes(work),
    )
    .expect_err("preset work max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::WorkBytes { .. })
    ));

    let depth = report.max_depth().saturating_sub(1);
    let error = codec::decode_table_style_preset_with_report(
        source,
        options(source).with_recursion_limit(depth),
    )
    .expect_err("preset nesting max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::Nesting { .. })
    ));

    let input = source.len().saturating_sub(1);
    let error = codec::decode_table_style_preset_with_report(
        source,
        options(source).with_max_input_bytes(input),
    )
    .expect_err("preset input max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::InputBytes { .. })
    ));
}

fn exercise_network_limits(source: &[u8], report: codec::DecodeReport) {
    let fields = report.fields().saturating_sub(1);
    let error = codec::decode_table_style_network_with_report(
        source,
        options(source).with_max_fields(fields),
    )
    .expect_err("network fields max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::Fields { .. })
    ));

    let work = report.work_bytes().saturating_sub(1);
    let error = codec::decode_table_style_network_with_report(
        source,
        options(source).with_max_work_bytes(work),
    )
    .expect_err("network work max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::WorkBytes { .. })
    ));

    let depth = report.max_depth().saturating_sub(1);
    let error = codec::decode_table_style_network_with_report(
        source,
        options(source).with_recursion_limit(depth),
    )
    .expect_err("network nesting max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::Nesting { .. })
    ));

    let input = source.len().saturating_sub(1);
    let error = codec::decode_table_style_network_with_report(
        source,
        options(source).with_max_input_bytes(input),
    )
    .expect_err("network input max-minus-one was accepted");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::InputBytes { .. })
    ));
}

fn exercise_model_rewrite(source: &[u8], old: u64, data: &[u8], before: &[u8]) {
    let mut new = 0x1_0000_u64
        | u64::from(data.first().copied().unwrap_or_default())
        | (u64::from(data.get(1).copied().unwrap_or_default()) << 8);
    if new == old {
        new = old.wrapping_add(1);
    }
    let write = codec::prepare_table_model_style_rewrite(source, old, new, options(source));
    let Ok(plan) = write else {
        if let Err(error) = write {
            observe_error(error);
        }
        return;
    };
    let requirements = plan.execution_requirements();
    assert!(requirements.output_bytes() <= MAX_OUTPUT_BYTES.max(source.len()));
    let exact = requirements.exact_limits();
    let result = plan.execute(exact);
    let Ok((candidate, report)) = result else {
        if let Err(error) = result {
            observe_error(error);
        }
        return;
    };
    assert_eq!(source, before, "model rewrite modified source");
    assert_eq!(report.output_bytes(), candidate.len());
    assert_eq!(report.fields(), requirements.fields());
    assert_eq!(report.work_bytes(), requirements.work_bytes());
    assert!(candidate.len() <= MAX_OUTPUT_BYTES.max(source.len()));
    let decoded = codec::decode_table_model(&candidate, options(&candidate))
        .unwrap_or_else(|error| panic!("model rewrite readback failed: {error}"));
    assert_eq!(decoded.style_identifier(), new);
    black_box((candidate, report));

    if requirements.output_bytes() > 0 {
        let too_small = requirements.exact_limits();
        let too_small = codec::RewriteExecutionLimits {
            max_output_bytes: too_small.max_output_bytes.saturating_sub(1),
            ..too_small
        };
        let retry = codec::prepare_table_model_style_rewrite(source, old, new, options(source))
            .and_then(|plan| plan.execute(too_small));
        assert!(
            retry.is_err(),
            "model rewrite accepted a short output limit"
        );
    }
}

fn exercise_style_inheritance(snapshot: codec::TableStyleSnapshot<'_>, data: &[u8]) {
    let first = 1u64 + u64::from(data.first().copied().unwrap_or_default() & 3);
    let node = codec::TableStyleNode::new(first, snapshot);
    let result = codec::resolve_table_style_appearance(&[node], first, options(snapshot.raw()));
    if let Ok(effective) = result {
        black_box(effective);
    } else if let Err(error) = result {
        observe_error(error);
    }
}

fn exercise_variation(data: &[u8]) {
    let overrides = codec::AppearanceOverrides {
        row_banding: Some(bit(data, 0, 0)),
        row_sizing: Some(bit(data, 0, 1)),
        body_horizontal: Some(bit(data, 1, 0)),
        body_vertical: Some(bit(data, 1, 1)),
        header_columns_horizontal: Some(bit(data, 1, 2)),
        header_rows_vertical: Some(bit(data, 1, 3)),
        footer_rows_vertical: Some(bit(data, 1, 4)),
    };
    let request = codec::TableStyleVariationWrite {
        parent_identifier: 7,
        stylesheet_identifier: 8,
        overrides,
    };
    match codec::canonical_table_style_variation(request, options(&[])) {
        Ok(payload) => {
            assert!(!payload.bytes().is_empty());
            assert_eq!(payload.report().output_bytes(), payload.bytes().len());
            let decoded = codec::decode_table_style(payload.bytes(), options(payload.bytes()))
                .unwrap_or_else(|error| panic!("variation readback failed: {error}"));
            assert_eq!(decoded.parent_identifier(), Some(7));
            black_box(payload);
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_stylesheet_append(source: &[u8], data: &[u8], before: &[u8]) {
    let append = codec::StylesheetStyleAppend {
        style_identifier: 0x2_0000 | u64::from(data.get(1).copied().unwrap_or_default()),
        parent_identifier: None,
    };
    let prepared = codec::prepare_stylesheet_append(source, append, options(source));
    let Ok(plan) = prepared else {
        if let Err(error) = prepared {
            observe_error(error);
        }
        return;
    };
    let result = plan.execute(plan.execution_requirements().exact_limits());
    let Ok((candidate, report)) = result else {
        if let Err(error) = result {
            observe_error(error);
        }
        return;
    };
    assert_eq!(source, before, "stylesheet rewrite modified source");
    assert_eq!(report.output_bytes(), candidate.len());
    let decoded = codec::decode_stylesheet(&candidate, options(&candidate))
        .unwrap_or_else(|error| panic!("stylesheet rewrite readback failed: {error}"));
    assert!(decoded.style_count() > 0);
    black_box((candidate, report));
}

fn bit(data: &[u8], offset: usize, bit: u8) -> bool {
    data.get(offset).copied().unwrap_or_default() & (1 << bit) != 0
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.to_string());
    black_box(error.resource_limit());
    black_box(error.allocation_requested());
}
