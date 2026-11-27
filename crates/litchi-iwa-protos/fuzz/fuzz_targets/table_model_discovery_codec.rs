#![no_main]

//! Fuzz the neutral, borrowed TableModel discovery projection.
//!
//! This target intentionally accepts protobuf payloads only.  It does not
//! construct a package, archive, ZIP member, or generated TableModel value.
//! The projection has no prepared rewrite API, so "replay" here means that
//! the scalar snapshot and the explicit facts entry points agree exactly on
//! the same caller-owned bytes.

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::table_model_discovery_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;

const TABLE_ID_FIELD: u32 = 1;
const TABLE_STYLE_FIELD: u32 = 3;
const BASE_DATA_STORE_FIELD: u32 = 4;
const TABLE_ROWS_FIELD: u32 = 6;
const TABLE_COLUMNS_FIELD: u32 = 7;
const TABLE_NAME_FIELD: u32 = 8;
const DEFAULT_ROW_HEIGHT_FIELD: u32 = 16;
const DEFAULT_COLUMN_WIDTH_FIELD: u32 = 17;
const BODY_CELL_STYLE_FIELD: u32 = 18;
const HEADER_ROW_STYLE_FIELD: u32 = 19;
const HEADER_COLUMN_STYLE_FIELD: u32 = 20;
const FOOTER_ROW_STYLE_FIELD: u32 = 21;
const BODY_TEXT_STYLE_FIELD: u32 = 24;
const HEADER_ROW_TEXT_STYLE_FIELD: u32 = 25;
const HEADER_COLUMN_TEXT_STYLE_FIELD: u32 = 26;
const FOOTER_ROW_TEXT_STYLE_FIELD: u32 = 27;

static FIXED_CASES: OnceLock<()> = OnceLock::new();

#[derive(Clone, Copy)]
enum LimitKind {
    Bytes,
    Fields,
    Work,
    Text,
    Nesting,
}

fuzz_target!(|data: &[u8]| {
    FIXED_CASES.get_or_init(run_fixed_cases);
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source);
});

fn run_fixed_cases() {
    let valid = model(&[3], None);

    for source in [
        valid.clone(),
        model_with_unknowns(),
        model_with_unknown_group_depth(8),
        model(&[0x81, 0x00], None),
        model_without(TABLE_NAME_FIELD),
        model_with_duplicate_rows(),
        model_with_wrong_rows_wire(),
        model_with_noncanonical_known_key(),
        model_with_noncanonical_unknown_value(),
        model_with_unterminated_group(),
        model_with_mismatched_group_end(),
        model_with_stray_group_end(),
        model_with_truncated_fixed64(),
        model_with_truncated_length(),
        model_with_invalid_wire(),
        model_with_truncated_varint(),
    ] {
        exercise_source(&source);
    }
}

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if data.starts_with(b"hex:") {
        let encoded = &data[5..];
        if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(1) {
            return None;
        }
        let mut output = Vec::with_capacity(encoded.len() / 2);
        let mut high = None;
        for byte in encoded.iter().copied() {
            if byte.is_ascii_whitespace() {
                continue;
            }
            let nibble = hex_nibble(byte)?;
            if let Some(previous) = high.take() {
                output.push((previous << 4) | nibble);
                if output.len() > MAX_INPUT_BYTES {
                    return None;
                }
            } else {
                high = Some(nibble);
            }
        }
        if high.is_some() {
            return None;
        }
        Some(output)
    } else if data.len() <= MAX_INPUT_BYTES {
        Some(data.to_vec())
    } else {
        None
    }
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn exercise_source(source: &[u8]) {
    let original = source.to_vec();
    let options = decode_options(source);
    let reported_snapshot = codec::decode_table_model_with_report(source, options);
    let reported_facts = codec::decode_table_model_facts_with_report(source, options);

    match (reported_snapshot, reported_facts) {
        (Ok((snapshot, report)), Ok((facts, facts_report))) => {
            assert_eq!(snapshot, facts);
            assert_eq!(report, facts_report);
            assert_snapshot_borrowed(source, snapshot);
            assert_eq!(report.input_bytes(), source.len());
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.text_bytes() <= MAX_TEXT_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            // These are codec-owned report categories, not allocator
            // telemetry: this borrowed projection retains no output bytes or
            // caller-visible scratch container.
            assert_eq!(report.allocations(), 0);
            assert_eq!(report.retained_bytes(), 0);
            assert_eq!(report.scratch_bytes(), 0);

            let scalar = codec::decode_table_model(source, options);
            let facts_alias = codec::decode_table_model_facts(source, options);
            let (Ok(scalar), Ok(facts_alias)) = (scalar, facts_alias) else {
                panic!("reported and scalar discovery entry points disagree");
            };
            assert_eq!(scalar, snapshot);
            assert_eq!(facts_alias, facts);
            exercise_limits(source, report);
        },
        (Err(snapshot_error), Err(facts_error)) => {
            assert_error_shape_matches(&snapshot_error, &facts_error);
            black_box(snapshot_error);
            black_box(facts_error);
        },
        _ => panic!("snapshot and facts discovery entry points disagree"),
    }
    assert_eq!(source, original.as_slice());
}

fn decode_options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::for_source(source)
        .with_max_input_bytes(MAX_INPUT_BYTES)
        .with_max_fields(MAX_FIELDS)
        .with_max_work_bytes(MAX_WORK_BYTES)
        .with_max_text_bytes(MAX_TEXT_BYTES)
        .with_recursion_limit(MAX_RECURSION)
}

fn assert_snapshot_borrowed(source: &[u8], snapshot: codec::TableModelDiscoverySnapshot<'_>) {
    assert_borrowed(source, snapshot.table_id());
    assert_borrowed(source, snapshot.table_name());
}

fn assert_borrowed(source: &[u8], value: &str) {
    if value.is_empty() {
        return;
    }
    let source_start = source.as_ptr() as usize;
    let source_end = source_start
        .checked_add(source.len())
        .expect("source pointer range should fit");
    let value_start = value.as_ptr() as usize;
    assert!(value_start >= source_start && value_start < source_end);
}

fn assert_error_shape_matches(first: &codec::DecodeError, second: &codec::DecodeError) {
    assert_eq!(first.resource_limit(), second.resource_limit());
    assert_eq!(
        first.missing_required_field(),
        second.missing_required_field()
    );
    assert_eq!(
        first.duplicate_singular_field(),
        second.duplicate_singular_field()
    );
    assert_eq!(first.noncanonical_reason(), second.noncanonical_reason());
}

fn exercise_limits(source: &[u8], report: codec::DecodeReport) {
    let options = decode_options(source);
    expect_limit(
        source,
        options.with_max_input_bytes(source.len().saturating_sub(1)),
        LimitKind::Bytes,
    );
    expect_limit(
        source,
        options.with_max_fields(report.fields().saturating_sub(1)),
        LimitKind::Fields,
    );
    expect_limit(
        source,
        options.with_max_work_bytes(report.work_bytes().saturating_sub(1)),
        LimitKind::Work,
    );
    if report.text_bytes() > 0 {
        expect_limit(
            source,
            options.with_max_text_bytes(report.text_bytes() - 1),
            LimitKind::Text,
        );
    }
    let nesting_ceiling = report.max_depth().saturating_sub(1);
    expect_limit(
        source,
        options.with_recursion_limit(nesting_ceiling),
        LimitKind::Nesting,
    );
}

fn expect_limit(source: &[u8], options: codec::DecodeOptions, expected: LimitKind) {
    let first = codec::decode_table_model_with_report(source, options)
        .expect_err("max-minus-one limit should reject");
    let second = codec::decode_table_model_facts_with_report(source, options)
        .expect_err("max-minus-one limit should reject");
    assert_limit_kind(&first, expected);
    assert_limit_kind(&second, expected);
    assert_error_shape_matches(&first, &second);
}

fn assert_limit_kind(error: &codec::DecodeError, expected: LimitKind) {
    let Some(limit) = error.resource_limit() else {
        panic!("expected a typed discovery resource limit");
    };
    match (expected, limit) {
        (LimitKind::Bytes, codec::DecodeLimit::Bytes { .. })
        | (LimitKind::Fields, codec::DecodeLimit::Fields { .. })
        | (LimitKind::Work, codec::DecodeLimit::Work { .. })
        | (LimitKind::Text, codec::DecodeLimit::Text { .. })
        | (LimitKind::Nesting, codec::DecodeLimit::Nesting { .. }) => {},
        _ => panic!("wrong typed discovery resource limit"),
    }
}

fn model(rows: &[u8], omit: Option<u32>) -> Vec<u8> {
    let mut output = Vec::new();
    if omit != Some(TABLE_ID_FIELD) {
        bytes_field(TABLE_ID_FIELD, b"id", &mut output);
    }
    if omit != Some(TABLE_STYLE_FIELD) {
        bytes_field(TABLE_STYLE_FIELD, &[0x0a, 0x01, 0x08, 0x01], &mut output);
    }
    if omit != Some(BASE_DATA_STORE_FIELD) {
        bytes_field(BASE_DATA_STORE_FIELD, &[0x0a, 0x00], &mut output);
    }
    if omit != Some(TABLE_ROWS_FIELD) {
        varint_field_raw(TABLE_ROWS_FIELD, rows, &mut output);
    }
    if omit != Some(TABLE_COLUMNS_FIELD) {
        varint_field(TABLE_COLUMNS_FIELD, 4, &mut output);
    }
    if omit != Some(TABLE_NAME_FIELD) {
        bytes_field(TABLE_NAME_FIELD, b"Table", &mut output);
    }
    if omit != Some(DEFAULT_ROW_HEIGHT_FIELD) {
        fixed64_field(DEFAULT_ROW_HEIGHT_FIELD, &mut output);
    }
    if omit != Some(DEFAULT_COLUMN_WIDTH_FIELD) {
        fixed64_field(DEFAULT_COLUMN_WIDTH_FIELD, &mut output);
    }
    for (field, value) in [
        (BODY_CELL_STYLE_FIELD, 2),
        (HEADER_ROW_STYLE_FIELD, 3),
        (HEADER_COLUMN_STYLE_FIELD, 4),
        (FOOTER_ROW_STYLE_FIELD, 5),
        (BODY_TEXT_STYLE_FIELD, 6),
        (HEADER_ROW_TEXT_STYLE_FIELD, 7),
        (HEADER_COLUMN_TEXT_STYLE_FIELD, 8),
        (FOOTER_ROW_TEXT_STYLE_FIELD, 9),
    ] {
        if omit != Some(field) {
            bytes_field(field, &[0x0a, 0x01, 0x08, value], &mut output);
        }
    }
    output
}

fn model_without(field: u32) -> Vec<u8> {
    model(&[3], Some(field))
}

fn model_with_unknowns() -> Vec<u8> {
    let mut output = model(&[3], None);
    varint_field(100, 7, &mut output);
    fixed64_field(101, &mut output);
    fixed32_field(102, &mut output);
    bytes_field(103, b"unknown", &mut output);
    unknown_group(&mut output, 1);
    output
}

fn model_with_unknown_group_depth(depth: u32) -> Vec<u8> {
    let mut output = model(&[3], None);
    unknown_group(&mut output, depth);
    output
}

fn unknown_group(output: &mut Vec<u8>, depth: u32) {
    for field in 100..100 + depth {
        key(output, field, 3);
    }
    varint_field(120, 1, output);
    bytes_field(121, &[0x7f], output);
    for field in (100..100 + depth).rev() {
        key(output, field, 4);
    }
}

fn model_with_duplicate_rows() -> Vec<u8> {
    let mut output = model(&[3], None);
    varint_field(TABLE_ROWS_FIELD, 9, &mut output);
    output
}

fn model_with_wrong_rows_wire() -> Vec<u8> {
    let mut output = model(&[3], None);
    bytes_field(TABLE_ROWS_FIELD, &[3], &mut output);
    output
}

fn model_with_noncanonical_known_key() -> Vec<u8> {
    let mut output = model(&[3], None);
    noncanonical_key(&mut output, TABLE_ROWS_FIELD, 0);
    output.push(1);
    output
}

fn model_with_noncanonical_unknown_value() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 104, 0);
    output.extend_from_slice(&[0x81, 0x00]);
    output
}

fn model_with_unterminated_group() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 105, 3);
    varint_field(106, 1, &mut output);
    output
}

fn model_with_mismatched_group_end() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 105, 3);
    key(&mut output, 106, 4);
    output
}

fn model_with_stray_group_end() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 105, 4);
    output
}

fn model_with_truncated_fixed64() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 107, 1);
    output.extend_from_slice(&[0; 3]);
    output
}

fn model_with_truncated_length() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 108, 2);
    output.push(4);
    output.push(1);
    output
}

fn model_with_invalid_wire() -> Vec<u8> {
    let mut output = model(&[3], None);
    key(&mut output, 109, 6);
    output
}

fn model_with_truncated_varint() -> Vec<u8> {
    let mut output = model(&[3], None);
    output.push(0x80);
    output
}

fn key(output: &mut Vec<u8>, field: u32, wire: u32) {
    put_varint(output, (u64::from(field) << 3) | u64::from(wire));
}

fn noncanonical_key(output: &mut Vec<u8>, field: u32, wire: u32) {
    let mut encoded = Vec::new();
    put_varint(&mut encoded, (u64::from(field) << 3) | u64::from(wire));
    let last = encoded
        .last_mut()
        .expect("a protobuf key always has one byte");
    *last |= 0x80;
    output.extend_from_slice(&encoded);
    output.push(0);
}

fn varint_field(output_field: u32, value: u64, output: &mut Vec<u8>) {
    key(output, output_field, 0);
    put_varint(output, value);
}

fn varint_field_raw(output_field: u32, value: &[u8], output: &mut Vec<u8>) {
    key(output, output_field, 0);
    output.extend_from_slice(value);
}

fn bytes_field(output_field: u32, value: &[u8], output: &mut Vec<u8>) {
    key(output, output_field, 2);
    put_varint(
        output,
        u64::try_from(value.len()).expect("fuzz fixture length fits u64"),
    );
    output.extend_from_slice(value);
}

fn fixed64_field(output_field: u32, output: &mut Vec<u8>) {
    key(output, output_field, 1);
    output.extend_from_slice(&[0; 8]);
}

fn fixed32_field(output_field: u32, output: &mut Vec<u8>) {
    key(output, output_field, 5);
    output.extend_from_slice(&[0; 4]);
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 128 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}
