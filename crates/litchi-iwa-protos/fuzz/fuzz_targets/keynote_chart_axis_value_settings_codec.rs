#![no_main]

//! Bounded fuzzing for the Keynote chart value-axis settings projection.
//!
//! The target accepts only the generated non-style extension payload.  It
//! deliberately exercises the strict wire pass, the private lazy Buffa view,
//! and the source-span rewrite plan as one contract: a successful decode must
//! agree with its reported form, a prepared rewrite must agree with its
//! one-shot equivalent, and every failed limit replay must leave the source
//! untouched.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_chart_axis_value_settings_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_ALLOCATIONS: usize = 16;
const MAX_RETAINED_BYTES: usize = 256 * 1024;
const MAX_SCRATCH_BYTES: usize = 128 * 1024;

// A source-built extension with all owned fields and an adjacent decades
// field.  Unknown fields in the fixed cases ensure that a campaign reaches
// source-span preservation even before libFuzzer finds a valid input.
const FULL_SOURCE: &[u8] = &[
    0x20, 0x07, // decades = 7 (adjacent, preserved)
    0x28, 0x05, // major steps = 5
    0x30, 0x02, // minor steps = 2
    0x40, 0x02, // logarithmic scale
    0x8a, 0x01, 0x09, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x59, 0x40, // max 100
    0x92, 0x01, 0x09, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0xc0, // min -10
];

const UNKNOWN_SOURCE: &[u8] = &[
    0x08, 0x96, 0x01, // unknown varint
    0x20, 0x09, // decades
    0x98, 0x06, 0x01, // unknown scalar
    0x28, 0x03, // major
    0x12, 0x03, 0x01, 0x02, 0x03, // unknown length-delimited
    0x30, 0x01, // minor
    0x40, 0x02, // scale
    0xa0, 0x06, 0x2a, // unknown varint
    0x5b, 0x08, 0x01, 0x5c, // balanced unknown group
];

const NESTED_UNKNOWN_SOURCE: &[u8] = &[
    0x8a, 0x01, 0x0d, // max wrapper, 13 bytes
    0x10, 0x07, // unknown nested varint
    0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x59, 0x40, // 100
    0x18, 0x01, // unknown nested varint
    0x92, 0x01, 0x09, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0xc0, // min
];

const FIXED_CASES: &[&[u8]] = &[
    &[],
    FULL_SOURCE,
    UNKNOWN_SOURCE,
    NESTED_UNKNOWN_SOURCE,
    &[0x28, 0x05, 0x28, 0x06],             // duplicate major
    &[0x2d, 0x00, 0x00, 0x80, 0x3f],       // wrong major wire type
    &[0x28, 0x80, 0x00],                   // non-canonical major varint
    &[0x8a, 0x01, 0x09, 0x09, 0x00, 0x00], // truncated bound
    &[0x8a, 0x01, 0x00],                   // missing nested number
    &[
        0x8a, 0x01, 0x12, 0x09, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0xc0,
    ],
    // Nested field 1 with the wrong (varint) wire type.
    &[0x8a, 0x01, 0x02, 0x08, 0x01],
    // Nested field 1 appears twice in one bound wrapper.
    &[
        0x8a, 0x01, 0x12, 0x09, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x24, 0xc0, 0x09, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ],
    &[0x5b, 0x08, 0x01], // unterminated group
];

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data) {
        exercise_source(&source, data);
    }

    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for source in FIXED_CASES {
            exercise_source(source, b"fixed-keynote-axis-value-settings");
        }
        exercise_limit_guards();
        exercise_input_limit();
        exercise_deep_groups();
    });
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
            return None;
        }
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
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
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
    .with_max_allocations(MAX_ALLOCATIONS)
    .with_max_retained_bytes(MAX_RETAINED_BYTES.max(source.len()))
    .with_max_scratch_bytes(MAX_SCRATCH_BYTES.max(source.len()))
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_axis_value_settings(source, decode_options);
    assert_eq!(
        source,
        original.as_slice(),
        "scalar decode modified its source"
    );
    let reported = codec::decode_axis_value_settings_with_report(source, decode_options);
    assert_eq!(
        source,
        original.as_slice(),
        "reported decode modified its source"
    );

    let (snapshot, report) = match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            (snapshot, report)
        },
        (Err(scalar_error), Err(reported_error)) => {
            assert_eq!(
                scalar_error, reported_error,
                "scalar/report decode errors disagree"
            );
            observe_error(scalar_error);
            observe_error(reported_error);
            return;
        },
        (scalar_result, reported_result) => panic!(
            "axis-value scalar/report decode disagreement: scalar={:?}, report={:?}",
            scalar_result.as_ref().map(|snapshot| snapshot.settings()),
            reported_result
                .as_ref()
                .map(|(snapshot, _)| snapshot.settings())
        ),
    };

    assert_eq!(snapshot.raw(), source, "snapshot lost its source slice");
    assert_report(report, source.len());
    assert_eq!(source, original.as_slice(), "decode changed its source");
    black_box(snapshot.settings());
    black_box(snapshot.value_axis_settings());
    black_box(snapshot.bounds());
    black_box(snapshot.steps());
    black_box(snapshot.scale());
    black_box(snapshot.decades());
    black_box(snapshot.scale_present());
    black_box(snapshot.explicit_scale());

    let extension = codec::decode_axis_value_settings_extension(source, decode_options)
        .unwrap_or_else(|error| panic!("axis-value extension decode failed: {error}"));
    assert_eq!(
        extension, snapshot,
        "extension decode disagreed with scalar decode"
    );

    exercise_preserve(source, snapshot, decode_options, &original);
    exercise_rewrites(source, snapshot.settings(), data, decode_options, &original);
}

fn exercise_preserve(
    source: &[u8],
    snapshot: codec::AxisValueSettingsSnapshot<'_>,
    decode_options: codec::DecodeOptions,
    original: &[u8],
) {
    let write = codec::AxisValueSettingsWrite::preserve();
    let result = codec::rewrite_axis_value_settings_with_report(source, write, decode_options);
    assert_eq!(source, original, "preserve rewrite modified its source");
    let Ok((output, report)) = result else {
        if let Err(error) = result {
            observe_error(error);
        }
        return;
    };
    assert_eq!(output, source, "semantic preserve was not an exact no-op");
    assert!(!report.changed(), "preserve rewrite reported a change");
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), source.len());
    assert_unknown_spans_preserved(source, &output);
    let readback = codec::decode_axis_value_settings(&output, options(&output))
        .unwrap_or_else(|error| panic!("preserve candidate readback failed: {error}"));
    assert_eq!(readback.settings(), snapshot.settings());
}

fn exercise_rewrites(
    source: &[u8],
    current: codec::AxisValueSettings,
    data: &[u8],
    decode_options: codec::DecodeOptions,
    original: &[u8],
) {
    let desired = desired_settings(data);
    let partial = codec::AxisValueBounds::new(
        Some(codec::AxisValueBound::new(-2.0).expect("finite bound")),
        None,
    )
    .expect("partial bound is valid");
    let bounded_steps = codec::AxisValueSteps::new(Some(3), Some(0)).expect("valid steps");
    let requests = [
        (codec::AxisValueSettingsWrite::new(desired), desired),
        (
            codec::AxisValueSettingsWrite::preserve().with_bounds(partial),
            codec::AxisValueSettings::new(partial, current.steps(), current.scale()),
        ),
        (
            codec::AxisValueSettingsWrite::preserve()
                .with_bounds(codec::AxisValueBounds::automatic()),
            codec::AxisValueSettings::new(
                codec::AxisValueBounds::automatic(),
                current.steps(),
                current.scale(),
            ),
        ),
        (
            codec::AxisValueSettingsWrite::preserve().with_steps(bounded_steps),
            codec::AxisValueSettings::new(current.bounds(), bounded_steps, current.scale()),
        ),
        (
            codec::AxisValueSettingsWrite::preserve()
                .with_steps(codec::AxisValueSteps::automatic()),
            codec::AxisValueSettings::new(
                current.bounds(),
                codec::AxisValueSteps::automatic(),
                current.scale(),
            ),
        ),
        (
            codec::AxisValueSettingsWrite::preserve().with_scale(codec::Scale::Logarithmic),
            codec::AxisValueSettings::new(
                current.bounds(),
                current.steps(),
                codec::Scale::Logarithmic,
            ),
        ),
        (
            codec::AxisValueSettingsWrite::preserve().with_explicit_scale(codec::Scale::Linear),
            codec::AxisValueSettings::new(current.bounds(), current.steps(), codec::Scale::Linear),
        ),
        (
            codec::AxisValueSettingsWrite::preserve().clear_scale(),
            codec::AxisValueSettings::new(current.bounds(), current.steps(), codec::Scale::Linear),
        ),
    ];

    for (write, expected) in requests {
        exercise_prepared(source, write, expected, decode_options, original);
    }
}

fn desired_settings(data: &[u8]) -> codec::AxisValueSettings {
    let low = -f64::from(data.first().copied().unwrap_or(5) % 32);
    let high = low + f64::from(data.get(1).copied().unwrap_or(17) % 32) + 1.0;
    let bounds = codec::AxisValueBounds::fixed(
        codec::AxisValueBound::new(low).expect("finite lower bound"),
        codec::AxisValueBound::new(high).expect("finite upper bound"),
    )
    .expect("generated bounds are ordered");
    let major = u32::from(data.get(2).copied().unwrap_or(3) % 32) + 1;
    let minor = u32::from(data.get(3).copied().unwrap_or(1) % 8);
    let steps = codec::AxisValueSteps::new(Some(major), Some(minor)).expect("generated steps");
    let scale = match data.get(4).copied().unwrap_or_default() & 3 {
        0 => codec::Scale::Linear,
        1 => codec::Scale::Logarithmic,
        2 => codec::Scale::Unsupported(9_001),
        _ => codec::Scale::Unsupported(-7),
    };
    codec::AxisValueSettings::new(bounds, steps, scale)
}

fn exercise_prepared(
    source: &[u8],
    write: codec::AxisValueSettingsWrite,
    expected: codec::AxisValueSettings,
    decode_options: codec::DecodeOptions,
    original: &[u8],
) {
    let prepared = codec::prepare_axis_value_settings_rewrite(source, write, decode_options);
    assert_eq!(source, original, "prepare modified its source");
    let Ok(prepared) = prepared else {
        if let Err(error) = prepared {
            observe_error(error);
        }
        return;
    };

    let requirements = prepared.execution_requirements();
    let preparation = prepared.prepare_report();
    assert_eq!(preparation.source_bytes(), source.len());
    assert!(preparation.fields() <= requirements.fields);
    assert!(preparation.work_bytes() <= requirements.work_bytes);
    assert!(preparation.max_depth() <= requirements.max_depth);
    assert!(preparation.retained_bytes() <= requirements.retained_bytes);
    assert!(preparation.scratch_bytes() <= requirements.scratch_bytes);
    assert_eq!(source, original, "prepare changed its source");

    let exact = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact axis-value prepared replay failed: {error}"));
    assert_eq!(source, original, "execute changed its source");
    assert_eq!(exact.as_bytes().len(), requirements.output_bytes);
    assert_eq!(exact.report().input_bytes(), source.len());
    assert_eq!(exact.report().output_bytes(), exact.as_bytes().len());
    assert_eq!(exact.report().fields(), requirements.fields);
    assert_eq!(exact.report().work_bytes(), requirements.work_bytes);
    assert_eq!(exact.report().max_depth(), requirements.max_depth);
    assert_eq!(exact.report().allocations(), requirements.allocations);
    assert_eq!(exact.report().retained_bytes(), requirements.retained_bytes);
    assert_eq!(exact.report().scratch_bytes(), requirements.scratch_bytes);
    assert_eq!(exact.report().changed(), exact.as_bytes() != source);

    let candidate = exact.as_bytes().to_vec();
    let readback = codec::decode_axis_value_settings(&candidate, options(&candidate))
        .unwrap_or_else(|error| panic!("axis-value candidate readback failed: {error}"));
    assert_eq!(readback.settings(), expected);
    assert_unknown_spans_preserved(source, &candidate);

    let one_shot = codec::rewrite_axis_value_settings_with_report(source, write, decode_options)
        .unwrap_or_else(|error| panic!("axis-value one-shot replay failed: {error}"));
    assert_eq!(one_shot.0, candidate, "prepared and one-shot bytes differ");
    assert_eq!(
        one_shot.1,
        exact.report(),
        "prepared and one-shot reports differ"
    );
    let alias = codec::rewrite_axis_value_settings(source, write, decode_options)
        .unwrap_or_else(|error| panic!("axis-value rewrite alias failed: {error}"));
    assert_eq!(
        alias, candidate,
        "rewrite alias differed from prepared output"
    );
    let extension = codec::rewrite_axis_value_settings_extension(source, write, decode_options)
        .unwrap_or_else(|error| panic!("axis-value extension rewrite failed: {error}"));
    assert_eq!(
        extension, candidate,
        "extension rewrite differed from prepared output"
    );
    assert_eq!(source, original, "one-shot changed its source");

    exercise_limit_failures(source, write, requirements, decode_options, original);
    black_box(exact);
}

fn exercise_limit_failures(
    source: &[u8],
    write: codec::AxisValueSettingsWrite,
    requirements: codec::RewriteExecutionRequirements,
    decode_options: codec::DecodeOptions,
    original: &[u8],
) {
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let probes = [
        (requirements.output_bytes > 0)
            .then(|| exact.with_output_bytes(requirements.output_bytes - 1)),
        (requirements.fields > 0).then(|| exact.with_fields(requirements.fields - 1)),
        (requirements.work_bytes > 0).then(|| exact.with_work_bytes(requirements.work_bytes - 1)),
        (requirements.max_depth > 0).then(|| exact.with_max_depth(requirements.max_depth - 1)),
        (requirements.allocations > 0)
            .then(|| exact.with_allocations(requirements.allocations - 1)),
        (requirements.retained_bytes > 0)
            .then(|| exact.with_retained_bytes(requirements.retained_bytes - 1)),
        (requirements.scratch_bytes > 0)
            .then(|| exact.with_scratch_bytes(requirements.scratch_bytes - 1)),
    ];
    for limits in probes.into_iter().flatten() {
        let prepared = codec::prepare_axis_value_settings_rewrite(source, write, decode_options);
        let Ok(prepared) = prepared else {
            continue;
        };
        let result = prepared.execute(limits);
        assert!(result.is_err(), "axis-value max-minus-one limit succeeded");
        if let Err(error) = result {
            observe_error(error);
        }
        assert_eq!(source, original, "failed limit replay changed its source");
    }
}

fn exercise_limit_guards() {
    let source = FULL_SOURCE;
    let (_, report) = codec::decode_axis_value_settings_with_report(source, options(source))
        .expect("full axis-value fixed source must decode");
    let cases = [
        codec::DecodeOptions::new(source.len() - 1, MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION),
        codec::DecodeOptions::new(
            source.len(),
            report.fields().saturating_sub(1),
            MAX_WORK_BYTES,
            MAX_RECURSION,
        ),
        codec::DecodeOptions::new(
            source.len(),
            MAX_FIELDS,
            report.work_bytes().saturating_sub(1),
            MAX_RECURSION,
        ),
        codec::DecodeOptions::new(
            source.len(),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            report.max_depth().saturating_sub(1),
        ),
    ];
    for decode_options in cases {
        let result = codec::decode_axis_value_settings_with_report(source, decode_options);
        assert!(
            result.is_err(),
            "axis-value max-minus-one decode limit succeeded"
        );
        if let Err(error) = result {
            observe_error(error);
        }
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let source = OVERSIZED.get_or_init(|| vec![0; MAX_INPUT_BYTES + 1].into_boxed_slice());
    let options =
        codec::DecodeOptions::new(MAX_INPUT_BYTES, MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION)
            .with_max_output_bytes(MAX_OUTPUT_BYTES)
            .with_max_allocations(MAX_ALLOCATIONS)
            .with_max_retained_bytes(MAX_RETAINED_BYTES)
            .with_max_scratch_bytes(MAX_SCRATCH_BYTES);
    let result = codec::decode_axis_value_settings(source, options);
    assert!(result.is_err(), "oversized axis-value source was accepted");
    if let Err(error) = result {
        assert!(error.resource_limit().is_some());
        observe_error(error);
    }
}

fn exercise_deep_groups() {
    let mut source = Vec::new();
    for _ in 0..80 {
        source.extend_from_slice(&[0x5b]);
    }
    source.extend_from_slice(&[0x08, 0x01]);
    for _ in 0..80 {
        source.extend_from_slice(&[0x5c]);
    }
    let result = codec::decode_axis_value_settings(&source, options(&source));
    assert!(
        result.is_err(),
        "deep unknown groups bypassed nesting ceiling"
    );
    if let Err(error) = result {
        observe_error(error);
    }
}

fn assert_report(report: codec::DecodeReport, source_bytes: usize) {
    assert_eq!(report.source_bytes(), source_bytes);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert_eq!(report.allocations(), 0);
    assert!(report.retained_bytes() <= MAX_RETAINED_BYTES.max(source_bytes));
    assert!(report.scratch_bytes() <= MAX_SCRATCH_BYTES.max(source_bytes));
}

fn assert_unknown_spans_preserved(source: &[u8], output: &[u8]) {
    let source_spans = wire_field_spans(source).expect("strict source had valid wire spans");
    let output_spans = wire_field_spans(output).expect("strict output had valid wire spans");
    let source_unknown = source_spans
        .iter()
        .filter(|(number, _)| !is_selected(*number))
        .map(|(_, raw)| (*raw).to_vec())
        .collect::<Vec<_>>();
    let output_unknown = output_spans
        .iter()
        .filter(|(number, _)| !is_selected(*number))
        .map(|(_, raw)| (*raw).to_vec())
        .collect::<Vec<_>>();
    assert_eq!(source_unknown, output_unknown, "unknown root spans changed");

    let source_presence = selected_presence(&source_spans);
    let output_presence = selected_presence(&output_spans);
    if source_presence == output_presence {
        assert_eq!(
            source_spans.len(),
            output_spans.len(),
            "selected replacement changed span count"
        );
        for ((source_number, source_raw), (output_number, output_raw)) in
            source_spans.iter().zip(output_spans.iter())
        {
            assert_eq!(
                is_selected(*source_number),
                is_selected(*output_number),
                "selected/unknown interleaving changed"
            );
            if !is_selected(*source_number) {
                assert_eq!(source_raw, output_raw, "unknown root span moved or changed");
            }
        }
    }

    // Bounds are selected outer fields, but their nested unknown records are
    // independently source-authoritative when the wrapper remains present.
    for number in [17, 18] {
        let Some(source_nested) = length_delimited_payload(&source_spans, number) else {
            continue;
        };
        let Some(output_nested) = length_delimited_payload(&output_spans, number) else {
            continue;
        };
        let source_inner = wire_field_spans(source_nested)
            .expect("strict source nested bound had valid spans")
            .into_iter()
            .filter(|(field, _)| *field != 1)
            .map(|(_, raw)| raw.to_vec())
            .collect::<Vec<_>>();
        let output_inner = wire_field_spans(output_nested)
            .expect("strict output nested bound had valid spans")
            .into_iter()
            .filter(|(field, _)| *field != 1)
            .map(|(_, raw)| raw.to_vec())
            .collect::<Vec<_>>();
        assert_eq!(
            source_inner, output_inner,
            "nested unknown bound spans changed"
        );
    }
}

fn is_selected(number: u32) -> bool {
    matches!(number, 5 | 6 | 8 | 17 | 18)
}

fn selected_presence(spans: &[(u32, &[u8])]) -> Vec<u32> {
    let mut selected = spans
        .iter()
        .filter_map(|(number, _)| is_selected(*number).then_some(*number))
        .collect::<Vec<_>>();
    selected.sort_unstable();
    selected
}

fn length_delimited_payload<'source>(
    spans: &[(u32, &'source [u8])],
    number: u32,
) -> Option<&'source [u8]> {
    spans.iter().find_map(|(field, raw)| {
        if *field != number {
            return None;
        }
        let (_, after_key) = read_varint(raw, 0)?;
        let (length, after_length) = read_varint(raw, after_key)?;
        let length = usize::try_from(length).ok()?;
        let end = after_length.checked_add(length)?;
        raw.get(after_length..end)
    })
}

fn wire_field_spans(source: &[u8]) -> Option<Vec<(u32, &[u8])>> {
    let mut spans = Vec::new();
    let mut offset = 0;
    while offset < source.len() {
        let start = offset;
        let (tag, next) = read_varint(source, offset)?;
        offset = next;
        let number = u32::try_from(tag >> 3).ok()?;
        let wire_type = tag & 7;
        if number == 0 || !skip_wire(source, &mut offset, wire_type, number, 0) {
            return None;
        }
        spans.push((number, source.get(start..offset)?));
    }
    Some(spans)
}

fn skip_wire(source: &[u8], offset: &mut usize, wire_type: u64, group: u32, depth: u32) -> bool {
    if depth > MAX_RECURSION + 1 {
        return false;
    }
    match wire_type {
        0 => read_varint(source, *offset)
            .map(|(_, end)| *offset = end)
            .is_some(),
        1 => advance(source, offset, 8),
        2 => {
            let Some((length, end)) = read_varint(source, *offset) else {
                return false;
            };
            let Ok(length) = usize::try_from(length) else {
                return false;
            };
            *offset = end;
            advance(source, offset, length)
        },
        3 => loop {
            let Some((tag, end)) = read_varint(source, *offset) else {
                return false;
            };
            *offset = end;
            let Ok(number) = u32::try_from(tag >> 3) else {
                return false;
            };
            let child_wire = tag & 7;
            if child_wire == 4 {
                return number == group;
            }
            if !skip_wire(source, offset, child_wire, number, depth + 1) {
                return false;
            }
        },
        4 | 6 | 7 => false,
        5 => advance(source, offset, 4),
        _ => false,
    }
}

fn advance(source: &[u8], offset: &mut usize, amount: usize) -> bool {
    let Some(end) = offset.checked_add(amount) else {
        return false;
    };
    if end > source.len() {
        return false;
    }
    *offset = end;
    true
}

fn read_varint(source: &[u8], mut offset: usize) -> Option<(u64, usize)> {
    let mut value = 0u64;
    for shift in (0..64).step_by(7) {
        let byte = *source.get(offset)?;
        offset += 1;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some((value, offset));
        }
    }
    None
}

fn observe_error(error: impl std::fmt::Debug + std::fmt::Display) {
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
