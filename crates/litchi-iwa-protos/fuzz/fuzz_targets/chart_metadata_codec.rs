#![no_main]

//! Bounded fuzzing for the shared borrowed chart-metadata projection.
//!
//! The target feeds the same caller-owned byte slice through both the modern
//! 5021 extension envelope and the legacy 5000 chart-info reader.  Package
//! object lookup and title resolution live at the format-reader boundary;
//! this target stays at the strict wire-codec boundary.

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::chart_metadata_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_DEPTH: u32 = 64;
const MAX_LABEL_COUNT: usize = 4 * 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RETAINED_BYTES: usize = 256 * 1024;

fuzz_target!(|data: &[u8]| {
    FIXED_CASES.get_or_init(run_fixed_cases);
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_both_formats(&source);
    exercise_mutation(&source, data);
});

static FIXED_CASES: OnceLock<()> = OnceLock::new();

fn run_fixed_cases() {
    for source in [
        modern_fixture(),
        modern_fixture_with_unknowns(),
        modern_missing_extension(),
        modern_invalid_label(),
        modern_duplicate_extension(),
        legacy_fixture(),
        legacy_fixture_without_inline_grid(),
        legacy_missing_model(),
        legacy_missing_chart_type(),
        legacy_invalid_label(),
        truncated_varint(),
        unterminated_group(),
    ] {
        exercise_both_formats(&source);
    }

    // A zero identifier is represented as `None` by the borrowed projection.
    // Duplicate tracking must still remember that the singular wire field was
    // present, otherwise a zero followed by a second value can slip through.
    assert_modern_rejected(modern_identifier_zero_then_repeat());
    assert_legacy_rejected(legacy_identifier_zero_then_repeat());
    assert_modern_rejected(modern_non_style_zero_then_repeat());
    assert_legacy_rejected(legacy_non_style_zero_then_repeat());
}

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
    codec::DecodeOptions::for_source(source)
        .with_max_input_bytes(MAX_INPUT_BYTES)
        .with_max_fields(MAX_FIELDS)
        .with_max_work_bytes(MAX_WORK_BYTES)
        .with_max_depth(MAX_DEPTH)
        .with_max_label_count(MAX_LABEL_COUNT)
        .with_max_text_bytes(MAX_TEXT_BYTES)
}

fn exercise_both_formats(source: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    exercise_modern(source, &decode_options);
    exercise_legacy(source, &decode_options);
    assert_eq!(source, original.as_slice(), "decoder modified its source");
}

fn exercise_modern(source: &[u8], options: &codec::DecodeOptions) {
    let decoded = codec::decode_modern_with_report(source, options);
    match decoded {
        Ok((snapshot, report)) => {
            assert_eq!(snapshot.format(), codec::ChartMetadataFormat::Modern);
            assert_eq!(snapshot.source(), source);
            assert_snapshot_bounds(source, &snapshot, report);
            let alias = codec::decode_modern(source, options)
                .unwrap_or_else(|error| panic!("modern aliases disagreed: {error}"));
            assert_eq!(alias, snapshot);
            exercise_grid_source(source, snapshot.grid_source());
            black_box((snapshot.chart_type(), snapshot.contains_default_data()));
            black_box(snapshot.non_style_ref());
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_legacy(source: &[u8], options: &codec::DecodeOptions) {
    let decoded = codec::decode_legacy_with_report(source, options);
    match decoded {
        Ok((snapshot, report)) => {
            assert_eq!(snapshot.format(), codec::ChartMetadataFormat::Legacy);
            assert_eq!(snapshot.source(), source);
            assert_snapshot_bounds(source, &snapshot, report);
            let alias = codec::decode_legacy(source, options)
                .unwrap_or_else(|error| panic!("legacy aliases disagreed: {error}"));
            assert_eq!(alias, snapshot);
            exercise_grid_source(source, snapshot.grid_source());
            black_box((snapshot.chart_type(), snapshot.contains_default_data()));
            black_box(snapshot.non_style_ref());
        },
        Err(error) => observe_error(error),
    }
}

fn assert_snapshot_bounds(
    source: &[u8],
    snapshot: &codec::ChartMetadataSnapshot<'_>,
    report: codec::DecodeReport,
) {
    assert_eq!(report.source_bytes(), source.len());
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.failure_work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_DEPTH);
    assert!(report.label_count() <= MAX_LABEL_COUNT);
    assert!(report.text_bytes() <= MAX_TEXT_BYTES);
    assert!(report.allocations() <= MAX_FIELDS);
    assert!(report.retained_bytes() <= MAX_RETAINED_BYTES);

    for labels in [snapshot.row_labels(), snapshot.column_labels()] {
        assert!(labels.len() <= MAX_LABEL_COUNT);
        for label in labels.iter() {
            assert!(label.len() <= MAX_TEXT_BYTES);
            assert_borrowed(source, label);
        }
    }
}

fn exercise_grid_source(source: &[u8], grid_source: Option<&[u8]>) {
    if let Some(grid) = grid_source {
        assert!(
            is_subslice(source, grid),
            "grid source escaped its root source"
        );
    }
}

fn assert_borrowed(source: &[u8], value: &str) {
    if value.is_empty() {
        return;
    }
    assert!(
        is_subslice(source, value.as_bytes()),
        "label was not borrowed"
    );
}

fn is_subslice(source: &[u8], value: &[u8]) -> bool {
    let source_start = source.as_ptr() as usize;
    let source_end = source_start.saturating_add(source.len());
    let value_start = value.as_ptr() as usize;
    let value_end = value_start.saturating_add(value.len());
    value_start >= source_start && value_end <= source_end
}

fn exercise_mutation(source: &[u8], data: &[u8]) {
    if source.is_empty() {
        return;
    }
    let mut mutated = source.to_vec();
    let index = usize::from(data.first().copied().unwrap_or_default()) % mutated.len();
    mutated[index] ^= 0xff;
    exercise_both_formats(&mutated);
}

fn modern_fixture() -> Vec<u8> {
    let mut archive = Vec::new();
    varint_field(&mut archive, 1, 2);
    varint_field(&mut archive, 6, 1);
    let mut grid = Vec::new();
    bytes_field(&mut grid, 1, b"North");
    bytes_field(&mut grid, 2, b"April");
    bytes_field(&mut archive, 7, &grid);
    let reference = reference(42);
    bytes_field(&mut archive, 10, &reference);
    let mut outer = Vec::new();
    bytes_field(&mut outer, 10000, &archive);
    outer
}

fn modern_fixture_with_unknowns() -> Vec<u8> {
    let mut source = modern_fixture();
    varint_field(&mut source, 999, 7);
    bytes_field(&mut source, 998, b"unknown");
    source
}

fn modern_missing_extension() -> Vec<u8> {
    let mut source = Vec::new();
    varint_field(&mut source, 1, 2);
    source
}

fn modern_invalid_label() -> Vec<u8> {
    let mut archive = Vec::new();
    let mut grid = Vec::new();
    bytes_field(&mut grid, 1, &[0xff]);
    bytes_field(&mut archive, 7, &grid);
    let mut source = Vec::new();
    bytes_field(&mut source, 10000, &archive);
    source
}

fn modern_duplicate_extension() -> Vec<u8> {
    let mut source = modern_fixture();
    let duplicate = modern_fixture();
    source.extend_from_slice(&duplicate);
    source
}

fn modern_identifier_zero_then_repeat() -> Vec<u8> {
    let mut archive = Vec::new();
    bytes_field(&mut archive, 10, &reference_with_identifiers(&[0, 42]));
    let mut source = Vec::new();
    bytes_field(&mut source, 10000, &archive);
    source
}

fn modern_non_style_zero_then_repeat() -> Vec<u8> {
    let mut archive = Vec::new();
    bytes_field(&mut archive, 10, &reference(0));
    bytes_field(&mut archive, 10, &reference(42));
    let mut source = Vec::new();
    bytes_field(&mut source, 10000, &archive);
    source
}

fn legacy_fixture() -> Vec<u8> {
    let mut model = Vec::new();
    bytes_field(&mut model, 2, &reference(7));
    let mut grid = Vec::new();
    varint_field(&mut grid, 1, 0);
    bytes_field(&mut grid, 2, b"North");
    bytes_field(&mut grid, 3, b"April");
    varint_field(&mut grid, 6, 0);
    bytes_field(&mut model, 5, &grid);
    let mut source = Vec::new();
    bytes_field(&mut source, 2, &model);
    varint_field(&mut source, 4, 2);
    source
}

fn legacy_fixture_without_inline_grid() -> Vec<u8> {
    let mut model = Vec::new();
    bytes_field(&mut model, 2, &reference(7));
    let mut source = Vec::new();
    bytes_field(&mut source, 2, &model);
    varint_field(&mut source, 4, 2);
    source
}

fn legacy_missing_model() -> Vec<u8> {
    let mut source = Vec::new();
    varint_field(&mut source, 4, 2);
    source
}

fn legacy_missing_chart_type() -> Vec<u8> {
    let mut model = Vec::new();
    bytes_field(&mut model, 2, &reference(7));
    let mut source = Vec::new();
    bytes_field(&mut source, 2, &model);
    source
}

fn legacy_invalid_label() -> Vec<u8> {
    let mut model = Vec::new();
    bytes_field(&mut model, 2, &reference(7));
    let mut grid = Vec::new();
    varint_field(&mut grid, 1, 0);
    bytes_field(&mut grid, 2, &[0xff]);
    varint_field(&mut grid, 6, 0);
    bytes_field(&mut model, 5, &grid);
    let mut source = Vec::new();
    bytes_field(&mut source, 2, &model);
    varint_field(&mut source, 4, 2);
    source
}

fn legacy_identifier_zero_then_repeat() -> Vec<u8> {
    let mut source = Vec::new();
    bytes_field(&mut source, 2, &[]);
    varint_field(&mut source, 4, 2);
    bytes_field(&mut source, 14, &reference_with_identifiers(&[0, 42]));
    source
}

fn legacy_non_style_zero_then_repeat() -> Vec<u8> {
    let mut source = legacy_fixture_without_inline_grid();
    bytes_field(&mut source, 14, &reference(0));
    bytes_field(&mut source, 14, &reference(42));
    source
}

fn truncated_varint() -> Vec<u8> {
    vec![0x80]
}

fn unterminated_group() -> Vec<u8> {
    vec![0x9b, 0x03, 0x08, 0x01]
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, 1, identifier);
    output
}

fn reference_with_identifiers(identifiers: &[u64]) -> Vec<u8> {
    let mut output = Vec::new();
    for &identifier in identifiers {
        varint_field(&mut output, 1, identifier);
    }
    output
}

fn assert_modern_rejected(source: Vec<u8>) {
    let decode_options = options(&source);
    let result = codec::decode_modern_with_report(&source, &decode_options);
    let error = result.expect_err("modern duplicate-zero recipe was accepted");
    let report = error.report();
    assert!(report.source_bytes() <= MAX_INPUT_BYTES);
    assert!(report.failure_work_bytes() <= MAX_WORK_BYTES);
    black_box(error);
}

fn assert_legacy_rejected(source: Vec<u8>) {
    let decode_options = options(&source);
    let result = codec::decode_legacy_with_report(&source, &decode_options);
    let error = result.expect_err("legacy duplicate-zero recipe was accepted");
    let report = error.report();
    assert!(report.source_bytes() <= MAX_INPUT_BYTES);
    assert!(report.failure_work_bytes() <= MAX_WORK_BYTES);
    black_box(error);
}

fn bytes_field(output: &mut Vec<u8>, field: u32, value: &[u8]) {
    key(output, field, 2);
    put_varint(output, value.len() as u64);
    output.extend_from_slice(value);
}

fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    key(output, field, 0);
    put_varint(output, value);
}

fn key(output: &mut Vec<u8>, field: u32, wire: u32) {
    put_varint(output, (u64::from(field) << 3) | u64::from(wire));
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 128 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn observe_error<E: std::fmt::Debug>(error: E) {
    black_box(format!("{error:?}"));
}
