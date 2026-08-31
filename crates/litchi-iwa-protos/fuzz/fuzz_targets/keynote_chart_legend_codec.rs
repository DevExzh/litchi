#![no_main]

//! Strict, source-preserving fuzzing for the generated Keynote chart-legend
//! visibility field.
//!
//! The target deliberately stays at the codec boundary.  It never constructs
//! a chart graph or retains libFuzzer's input: the codec receives a bounded,
//! caller-owned copy and is checked for borrowed reads, exact no-ops,
//! successful rewrite readback, idempotence, and preservation of every
//! unselected wire span.

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_chart_legend_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const LEGEND_FIELD: u64 = 20;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES + 1;

// Keep a few valid and intentionally invalid layouts reachable before a
// campaign discovers them.  The unknown field is placed on both sides of the
// selected field in separate cases so its relative wire order is observable.
const FIXED_CASES: &[&[u8]] = &[
    &[],                                         // absent
    &[0xa0, 0x01, 0x00],                         // explicit false
    &[0xa0, 0x01, 0x01],                         // explicit true
    &[0xa0, 0x01, 0x01, 0xa0, 0x01, 0x00],       // duplicate
    &[0xa2, 0x01, 0x01, 0x01],                   // selected field has wrong wire type
    &[0xa0, 0x01, 0x80, 0x00],                   // non-canonical zero bool
    &[0x80, 0xfa, 0x01, 0x07, 0xa0, 0x01, 0x01], // unknown before
    &[0xa0, 0x01, 0x01, 0x80, 0xfa, 0x01, 0x07], // unknown after
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);

    // Fixed layouts ensure that strict malformed paths and unknown-span
    // rewrites remain covered even when arbitrary input is mostly invalid.
    static FIXED: OnceLock<()> = OnceLock::new();
    FIXED.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-keynote-chart-legend");
        }
        exercise_limit_guards();
        exercise_input_limit();
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

fn options(source: &[u8], data: &[u8]) -> codec::DecodeOptions {
    // Vary each ceiling while keeping it finite and large enough for the
    // source-derived no-op path in ordinary cases.
    let field_slack = usize::from(data.first().copied().unwrap_or(31)) % 256;
    let work_slack = usize::from(data.get(1).copied().unwrap_or(31)) % 1024;
    let depth = 1 + u32::from(data.get(2).copied().unwrap_or(7) % 32);
    codec::DecodeOptions::new(
        source.len().max(1),
        (source.len().saturating_add(field_slack).max(8)).min(MAX_FIELDS),
        source
            .len()
            .saturating_mul(8)
            .saturating_add(work_slack)
            .clamp(8, MAX_WORK_BYTES),
        depth.min(MAX_RECURSION),
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
}

fn generous_options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source, data);

    let decoded = codec::decode_chart_legend(source, decode_options);
    assert_eq!(source, original.as_slice(), "decode modified its source");
    let Ok(snapshot) = decoded else {
        if let Err(error) = decoded {
            observe_error(error);
        }
        return;
    };

    assert_eq!(snapshot.raw(), source, "snapshot lost its source slice");
    if let Ok((reported, report)) =
        codec::decode_chart_legend_visibility_with_report(source, decode_options)
    {
        assert_eq!(reported, snapshot, "reported decode disagreed with decode");
        assert_report(report, source.len());
    }
    assert_eq!(
        source,
        original.as_slice(),
        "reported decode modified its source"
    );

    let current = snapshot.visible();
    black_box(current);

    // A source-derived write must be an exact no-op, including unknown spans
    // and the original selected-field encoding/order.
    exercise_rewrite(source, current, current, decode_options, &original);

    // Exercise all proto2 presence states.  The arbitrary byte chooses the
    // first replacement, while the fixed matrix always covers removal and
    // both explicit bool values.
    let varied = if data.first().copied().unwrap_or_default() & 1 == 0 {
        Some(true)
    } else {
        Some(false)
    };
    for desired in [None, Some(false), Some(true), varied] {
        exercise_rewrite(
            source,
            current,
            desired,
            generous_options(source),
            &original,
        );
    }
}

fn exercise_rewrite(
    source: &[u8],
    before: Option<bool>,
    desired: Option<bool>,
    options: codec::DecodeOptions,
    original: &[u8],
) {
    let write = codec::ChartLegendWrite::explicit(desired);
    let result = codec::rewrite_chart_legend_visibility_with_report(source, write, options);
    assert_eq!(source, original, "rewrite modified its source");
    let Ok((output, report)) = result else {
        if let Err(error) = result {
            observe_error(error);
        }
        return;
    };

    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), output.len());
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(output.len() <= MAX_OUTPUT_BYTES.max(source.len()));
    assert_eq!(source, original, "successful rewrite changed its source");
    assert_unknown_spans_preserved(source, &output);

    let readback_options = generous_options(&output);
    let Ok(readback) = codec::decode_chart_legend(&output, readback_options) else {
        return;
    };
    assert_eq!(readback.visible(), desired);

    // Repeating the same semantic request must not keep appending fields or
    // otherwise normalize an already-written candidate.
    let second =
        codec::rewrite_chart_legend_visibility_with_report(&output, write, readback_options);
    let Ok((idempotent, second_report)) = second else {
        return;
    };
    assert_eq!(idempotent, output, "legend rewrite was not idempotent");
    assert!(!second_report.changed());
    assert_eq!(second_report.input_bytes(), output.len());
    assert_eq!(second_report.output_bytes(), output.len());
    assert_unknown_spans_preserved(source, &idempotent);

    // An inverse restores the semantic value.  Exact source bytes are only
    // required when the selected field's presence was not changed: removing
    // a span necessarily loses its former insertion position.
    let inverse = codec::rewrite_chart_legend_visibility_with_report(
        &output,
        codec::ChartLegendWrite::explicit(before),
        generous_options(&output),
    );
    if let Ok((restored, _)) = inverse {
        let restored_read = codec::decode_chart_legend(&restored, generous_options(&restored));
        if let Ok(restored_read) = restored_read {
            assert_eq!(restored_read.visible(), before);
            assert_unknown_spans_preserved(source, &restored);
        }
    }
}

fn assert_report(report: codec::DecodeReport, source_bytes: usize) {
    assert_eq!(report.source_bytes(), source_bytes);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
}

fn exercise_limit_guards() {
    let source = [0xa0, 0x01, 0x01];

    for options in [
        codec::DecodeOptions::new(
            source.len().saturating_sub(1),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            8,
        ),
        codec::DecodeOptions::new(source.len(), 1, MAX_WORK_BYTES, 8),
        codec::DecodeOptions::new(source.len(), MAX_FIELDS, 1, 8),
        codec::DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, 0),
    ] {
        let result = codec::decode_chart_legend(&source, options);
        if let Err(error) = result {
            observe_error(error);
        }
    }

    let capped = codec::DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, 8)
        .with_max_output_bytes(source.len().saturating_sub(1));
    let result = codec::rewrite_chart_legend_visibility_with_report(
        &source,
        codec::ChartLegendWrite::new(false),
        capped,
    );
    if let Err(error) = result {
        observe_error(error);
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let source = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    let result = codec::decode_chart_legend(
        source,
        codec::DecodeOptions::new(MAX_INPUT_BYTES, MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION),
    );
    if let Err(error) = result {
        observe_error(error);
    }
}

fn assert_unknown_spans_preserved(source: &[u8], output: &[u8]) {
    let (Some(source_spans), Some(output_spans)) =
        (wire_field_spans(source), wire_field_spans(output))
    else {
        return;
    };
    let source_unknown = source_spans
        .iter()
        .filter_map(|(number, raw)| (*number != LEGEND_FIELD).then_some(*raw))
        .collect::<Vec<_>>();
    let output_unknown = output_spans
        .iter()
        .filter_map(|(number, raw)| (*number != LEGEND_FIELD).then_some(*raw))
        .collect::<Vec<_>>();
    assert_eq!(output_unknown, source_unknown, "unknown wire spans changed");
}

fn wire_field_spans<'source>(source: &'source [u8]) -> Option<Vec<(u64, &'source [u8])>> {
    let mut spans = Vec::new();
    let mut offset = 0;
    while offset < source.len() {
        let start = offset;
        let (tag, next) = read_varint(source, offset)?;
        offset = next;
        let number = tag >> 3;
        let wire_type = tag & 7;
        skip_wire(source, &mut offset, wire_type, number)?;
        spans.push((number, &source[start..offset]));
    }
    Some(spans)
}

fn skip_wire(source: &[u8], offset: &mut usize, wire_type: u64, group: u64) -> Option<()> {
    match wire_type {
        0 => {
            *offset = read_varint(source, *offset)?.1;
            Some(())
        },
        1 => advance(source, offset, 8),
        2 => {
            let (length, next) = read_varint(source, *offset)?;
            *offset = next;
            advance(source, offset, usize::try_from(length).ok()?)
        },
        3 => loop {
            let (tag, next) = read_varint(source, *offset)?;
            *offset = next;
            let number = tag >> 3;
            let child_wire = tag & 7;
            if child_wire == 4 {
                return (number == group).then_some(());
            }
            skip_wire(source, offset, child_wire, number)?;
        },
        5 => advance(source, offset, 4),
        _ => None,
    }
}

fn advance(source: &[u8], offset: &mut usize, amount: usize) -> Option<()> {
    let end = offset.checked_add(amount)?;
    (end <= source.len()).then(|| *offset = end)
}

fn read_varint(source: &[u8], mut offset: usize) -> Option<(u64, usize)> {
    let mut value = 0u64;
    for shift in (0..64).step_by(7) {
        let byte = *source.get(offset)?;
        offset = offset.checked_add(1)?;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some((value, offset));
        }
    }
    None
}

fn observe_error(error: impl std::fmt::Debug + std::fmt::Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
