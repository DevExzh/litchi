#![no_main]

//! Bounded fuzzing for the fresh chart-grid authoring seam.
//!
//! The creation codec receives semantic labels and numeric cells rather than
//! arbitrary protobuf bytes.  This target therefore uses the input as a
//! bounded generator, then sends the resulting wire payload through the
//! borrowed chart-grid reader.  Every successful prepared operation is run
//! with its exact preflight limits and repeated with the same seed to make
//! deterministic identity generation observable.

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::chart_data_codec::{DecodeOptions, decode_grid};
use litchi_iwa_protos::chart_grid_creation_codec::{
    ChartGridCreationRequest, EncodeOptions, encode_chart_grid, prepare_chart_grid_creation,
};

// Keep generator-owned storage and every codec route finite.  The dimensions
// are deliberately much smaller than the input ceiling: a malformed or
// adversarial byte stream can never turn into a large matrix allocation.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_ALLOCATIONS: usize = 64;
const MAX_RETAINED_BYTES: usize = 512 * 1024;
const MAX_SCRATCH_BYTES: usize = 128 * 1024;
const MAX_DEPTH: u32 = 3;
const MAX_ROWS: usize = 16;
const MAX_COLUMNS: usize = 16;
const MAX_LABEL_BYTES: usize = 24;

fuzz_target!(|data: &[u8]| {
    let Some(input) = normalize_input(data) else {
        return;
    };
    exercise(&input);
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

fn exercise(input: &[u8]) {
    let (row_labels, column_labels, values, seed) = generated_grid(input);
    let request = ChartGridCreationRequest::new(&row_labels, &column_labels, &values, seed);
    let encode_options = options(&request);

    // Preparation is allowed to refuse a finite request, but it must never
    // panic or mutate the caller-owned semantic values.  Valid generated
    // requests are kept small enough that the broad profile normally admits
    // them; the explicit refusal probes below exercise zero and one-below
    // budgets independently.
    let prepared = match prepare_chart_grid_creation(request, encode_options) {
        Ok(prepared) => prepared,
        Err(error) => {
            black_box(format_args!("{error:?}"));
            exercise_refusal(&row_labels, &column_labels, &values, seed);
            return;
        },
    };
    let requirements = prepared.execution_requirements();
    assert!(requirements.output_bytes <= MAX_OUTPUT_BYTES);
    assert!(requirements.fields <= MAX_FIELDS);
    assert!(requirements.work_bytes <= MAX_WORK_BYTES);
    assert!(requirements.allocations <= MAX_ALLOCATIONS);

    let output = prepared
        .execute(requirements.exact())
        .expect("exact chart-grid creation requirements must execute");
    let encoded = output.into_bytes();
    assert_eq!(encoded.len(), requirements.output_bytes);
    assert!(!encoded.is_empty(), "a chart grid must have a wire root");

    let direct_request = ChartGridCreationRequest::new(&row_labels, &column_labels, &values, seed);
    let direct_options = options(&direct_request);
    let direct = encode_chart_grid(direct_request, direct_options)
        .expect("the one-shot chart-grid encoder must match preparation")
        .into_bytes();
    assert_eq!(encoded, direct, "prepared and one-shot output differ");

    // A second preparation with the same semantic input must produce exactly
    // the same labels, values, framing, and seed-derived identity map.
    let repeated_request =
        ChartGridCreationRequest::new(&row_labels, &column_labels, &values, seed);
    let repeated_options = options(&repeated_request);
    let repeated = prepare_chart_grid_creation(repeated_request, repeated_options)
        .expect("the same admitted chart-grid request must prepare twice")
        .execute(requirements.exact())
        .expect("the repeated exact chart-grid execution must succeed")
        .into_bytes();
    assert_eq!(
        encoded, repeated,
        "chart-grid creation is not deterministic"
    );

    let decoded = decode_grid(&encoded, &decode_options(&encoded))
        .expect("a successfully authored chart grid must decode");
    assert_eq!(decoded.row_count(), row_labels.len());
    assert_eq!(decoded.column_count(), column_labels.len());
    for (decoded, expected) in decoded.row_labels().iter().zip(&row_labels) {
        assert_eq!(decoded, expected.as_str());
    }
    for (decoded, expected) in decoded.column_labels().iter().zip(&column_labels) {
        assert_eq!(decoded, expected.as_str());
    }
    for (decoded_row, expected_row) in decoded.rows().iter().zip(&values) {
        assert_eq!(decoded_row.len(), expected_row.len());
        for (decoded_value, expected_value) in decoded_row.values().zip(expected_row) {
            assert_eq!(
                decoded_value.map(f64::to_bits),
                expected_value.map(f64::to_bits)
            );
        }
    }

    exercise_refusal(&row_labels, &column_labels, &values, seed);
    black_box((&decoded, &encoded));
}

fn options(request: &ChartGridCreationRequest<'_>) -> EncodeOptions {
    EncodeOptions::for_request(request)
        .with_max_output_bytes(MAX_OUTPUT_BYTES)
        .with_max_fields(MAX_FIELDS)
        .with_max_work_bytes(MAX_WORK_BYTES)
        .with_max_allocations(MAX_ALLOCATIONS)
        .with_max_retained_bytes(MAX_RETAINED_BYTES)
        .with_max_scratch_bytes(MAX_SCRATCH_BYTES)
        .with_max_cells(MAX_ROWS.saturating_mul(MAX_COLUMNS))
        .with_max_labels(MAX_ROWS.saturating_add(MAX_COLUMNS))
        .with_max_text_bytes(
            MAX_ROWS
                .saturating_add(MAX_COLUMNS)
                .saturating_mul(MAX_LABEL_BYTES),
        )
        .with_max_depth(MAX_DEPTH)
}

fn decode_options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len(),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        64,
        MAX_ROWS.saturating_mul(MAX_COLUMNS),
        MAX_ROWS.saturating_add(MAX_COLUMNS),
        MAX_ROWS
            .saturating_add(MAX_COLUMNS)
            .saturating_mul(MAX_LABEL_BYTES),
    )
}

fn exercise_refusal(
    row_labels: &[String],
    column_labels: &[String],
    values: &[Vec<Option<f64>>],
    seed: u64,
) {
    // The request remains finite, so these calls exercise typed resource
    // refusals before any output vector can be allocated.  A codec may report
    // a different first axis, but both calls must return an error and never
    // panic.
    for options in [
        options(&ChartGridCreationRequest::new(
            row_labels,
            column_labels,
            values,
            seed,
        ))
        .with_max_output_bytes(0),
        options(&ChartGridCreationRequest::new(
            row_labels,
            column_labels,
            values,
            seed,
        ))
        .with_max_fields(0),
        options(&ChartGridCreationRequest::new(
            row_labels,
            column_labels,
            values,
            seed,
        ))
        .with_max_work_bytes(0),
        options(&ChartGridCreationRequest::new(
            row_labels,
            column_labels,
            values,
            seed,
        ))
        .with_max_allocations(0),
    ] {
        let request = ChartGridCreationRequest::new(row_labels, column_labels, values, seed);
        let result = prepare_chart_grid_creation(request, options)
            .and_then(|prepared| prepared.execute(prepared.execution_requirements().exact()));
        assert!(result.is_err(), "an empty resource budget was accepted");
        let _ = black_box(result);
    }
}

fn generated_grid(input: &[u8]) -> (Vec<String>, Vec<String>, Vec<Vec<Option<f64>>>, u64) {
    let byte = |offset: usize| input.get(offset).copied().unwrap_or_default();
    let rows = usize::from(byte(0) % (MAX_ROWS as u8)).saturating_add(1);
    let columns = usize::from(byte(1) % (MAX_COLUMNS as u8)).saturating_add(1);
    let row_labels = labels(input, b"row", rows);
    let column_labels = labels(input, b"column", columns);
    let mut values = Vec::with_capacity(rows);
    for row in 0..rows {
        let mut cells = Vec::with_capacity(columns);
        for column in 0..columns {
            let ordinal = row
                .checked_mul(columns)
                .and_then(|value| value.checked_add(column))
                .expect("bounded grid ordinal");
            cells.push(generated_value(input, ordinal));
        }
        values.push(cells);
    }
    (row_labels, column_labels, values, seed(input))
}

fn labels(input: &[u8], prefix: &[u8], count: usize) -> Vec<String> {
    let mut labels = Vec::with_capacity(count);
    for index in 0..count {
        let mut label = String::with_capacity(prefix.len() + 1 + MAX_LABEL_BYTES);
        label.push_str(std::str::from_utf8(prefix).expect("static ASCII prefix"));
        label.push('-');
        let length = usize::from(
            input
                .get(index.saturating_add(2))
                .copied()
                .unwrap_or_default()
                % (MAX_LABEL_BYTES as u8),
        )
        .saturating_add(1);
        for offset in 0..length {
            let value = input
                .get(index.saturating_add(offset).saturating_add(3))
                .copied()
                .unwrap_or_default();
            label.push(char::from(b'a' + value % 26));
        }
        labels.push(label);
    }
    labels
}

fn generated_value(input: &[u8], ordinal: usize) -> Option<f64> {
    let marker = input
        .get(ordinal.saturating_mul(3).saturating_add(2))
        .copied()
        .unwrap_or_default();
    if marker % 7 == 0 {
        return None;
    }
    if marker == 0x5a {
        return Some(0.0);
    }
    if marker == 0xa5 {
        return Some(-0.0);
    }
    let magnitude = f64::from(marker)
        + f64::from(
            input
                .get(ordinal.saturating_mul(5).saturating_add(3))
                .copied()
                .unwrap_or_default(),
        ) / 256.0
        + 0.125;
    Some(if marker & 1 == 0 {
        magnitude
    } else {
        -magnitude
    })
}

fn seed(input: &[u8]) -> u64 {
    let mut bytes = [0_u8; 8];
    let length = input.len().min(bytes.len());
    bytes[..length].copy_from_slice(&input[..length]);
    u64::from_le_bytes(bytes)
}
