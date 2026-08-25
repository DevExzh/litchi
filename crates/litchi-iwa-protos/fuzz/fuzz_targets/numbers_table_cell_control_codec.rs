#![no_main]

//! Strict-wire fuzzing for the unified Numbers interactive CellSpec seam.
//!
//! The neutral codec accepts interaction types 4--8 (stepper, slider,
//! star-rating, pop-up, and checkbox at the native layer) while keeping raw
//! unknown fields caller-owned.  This target probes every control shape,
//! malformed fixed64/Reference framing, deep and balanced unknown groups,
//! prepared exact replay, and each finite limit without publishing partial
//! output.

use std::{fmt::Debug, hint::black_box};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_control_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_REFERENCES: usize = 1_024;
const MAX_ITEMS: usize = 1_024;
const MAX_TEXT_BYTES: usize = 64 * 1024;

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data) {
        exercise_source(&source);
    }
    for source in fixed_cases() {
        exercise_source(&source);
    }
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        MAX_INPUT_BYTES.max(source.len()),
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    exercise_cell_spec(source, &before);
    exercise_format(source, &before);
    exercise_limit_profiles(source);
    exercise_prepared_writes();
    assert_eq!(
        source,
        before.as_slice(),
        "control codec mutated caller bytes"
    );
}

fn exercise_cell_spec(source: &[u8], before: &[u8]) {
    let scalar = codec::decode_control_cell_spec(source, options(source));
    let reported = codec::decode_control_cell_spec_with_report(source, options(source));
    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert!((4..=8).contains(&snapshot.interaction_type()));
            assert_eq!(snapshot.raw(), source);
            assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source.len()));
            assert!(report.output_bytes() <= MAX_OUTPUT_BYTES.max(source.len()));
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            assert!(report.references() <= MAX_REFERENCES);
            black_box((snapshot, report));
        },
        (Err(scalar_error), Err(reported_error)) => {
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "control CellSpec scalar/report disagreement: scalar={:?}, report={:?}",
            scalar_result.map(|snapshot| snapshot.interaction_type()),
            reported_result.map(|(snapshot, _)| snapshot.interaction_type())
        ),
    }
    assert_eq!(source, before, "control CellSpec decode changed source");
}

fn exercise_format(source: &[u8], before: &[u8]) {
    let scalar = codec::decode_control_format(source, options(source));
    let reported = codec::decode_control_format_with_report(source, options(source));
    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert!(matches!(snapshot.format_type(), 256..=263 | 267..=269));
            assert_eq!(snapshot.raw(), source);
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            black_box((snapshot, report));
        },
        (Err(scalar_error), Err(reported_error)) => {
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "control Format scalar/report disagreement: scalar={:?}, report={:?}",
            scalar_result.map(|snapshot| snapshot.format_type()),
            reported_result.map(|(snapshot, _)| snapshot.format_type())
        ),
    }
    assert_eq!(source, before, "control Format decode changed source");
}

fn exercise_prepared_writes() {
    let write_options = codec::DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    );
    for (interaction, minimum, maximum, increment) in [
        (codec::CHECKBOX_INTERACTION_TYPE, None, None, None),
        (
            codec::STAR_RATING_INTERACTION_TYPE,
            Some(0.0),
            Some(5.0),
            Some(1.0),
        ),
        (
            codec::SLIDER_INTERACTION_TYPE,
            Some(0.0),
            Some(100.0),
            Some(10.0),
        ),
        (
            codec::STEPPER_INTERACTION_TYPE,
            Some(1.0),
            Some(10.0),
            Some(1.0),
        ),
    ] {
        let Ok(plan) =
            codec::prepare_cell_spec_write(interaction, minimum, maximum, increment, write_options)
        else {
            continue;
        };
        let requirements = plan.execution_requirements();
        assert_requirements(plan.prepare_report(), requirements);
        let output = plan
            .execute(codec::RewriteExecutionLimits::exact(requirements))
            .unwrap_or_else(|error| panic!("exact control CellSpec execution failed: {error}"));
        assert_eq!(output.bytes().len(), requirements.output_bytes());
        let readback = codec::decode_control_cell_spec(output.bytes(), options(output.bytes()))
            .unwrap_or_else(|error| panic!("control CellSpec prepared readback failed: {error}"));
        assert_eq!(readback.interaction_type(), interaction);
        black_box(output.report());

        if requirements.output_bytes() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .is_err()
            );
        }
        if requirements.fields() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1)
                )
                .is_err()
            );
        }
        if requirements.work_bytes() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1)
                )
                .is_err()
            );
        }
        if requirements.max_depth() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_max_depth(requirements.max_depth() - 1)
                )
                .is_err()
            );
        }
        if requirements.references() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_references(requirements.references() - 1)
                )
                .is_err()
            );
        }
        if requirements.allocations() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_allocations(requirements.allocations() - 1)
                )
                .is_err()
            );
        }
        if requirements.retained_bytes() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_retained_bytes(requirements.retained_bytes() - 1)
                )
                .is_err()
            );
        }
        if requirements.scratch_bytes() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_scratch_bytes(requirements.scratch_bytes() - 1)
                )
                .is_err()
            );
        }

        let canonical =
            codec::canonical_cell_spec(interaction, minimum, maximum, increment, write_options)
                .unwrap_or_else(|error| panic!("canonical control CellSpec failed: {error}"));
        assert_eq!(canonical.bytes(), output.bytes());
        let rewritten =
            codec::rewrite_cell_spec(interaction, minimum, maximum, increment, write_options)
                .unwrap_or_else(|error| panic!("rewrite control CellSpec failed: {error}"));
        assert_eq!(rewritten.bytes(), output.bytes());
    }

    for format_type in [256_u32, 263, 267] {
        let Ok(plan) = codec::prepare_format_write(format_type, write_options) else {
            continue;
        };
        let requirements = plan.execution_requirements();
        assert_requirements(plan.prepare_report(), requirements);
        let output = plan
            .execute(codec::RewriteExecutionLimits::exact(requirements))
            .unwrap_or_else(|error| panic!("exact control Format execution failed: {error}"));
        let readback = codec::decode_control_format(output.bytes(), options(output.bytes()))
            .unwrap_or_else(|error| panic!("control Format prepared readback failed: {error}"));
        assert_eq!(readback.format_type(), format_type);
        let canonical = codec::canonical_format(format_type, write_options)
            .unwrap_or_else(|error| panic!("canonical control Format failed: {error}"));
        assert_eq!(canonical.bytes(), output.bytes());
        let rewritten = codec::rewrite_format(format_type, write_options)
            .unwrap_or_else(|error| panic!("rewrite control Format failed: {error}"));
        assert_eq!(rewritten.bytes(), output.bytes());
        if requirements.output_bytes() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .is_err()
            );
        }
        if requirements.fields() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1)
                )
                .is_err()
            );
        }
        if requirements.work_bytes() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1)
                )
                .is_err()
            );
        }
        if requirements.allocations() > 0 {
            assert!(
                plan.execute(
                    codec::RewriteExecutionLimits::exact(requirements)
                        .with_allocations(requirements.allocations() - 1)
                )
                .is_err()
            );
        }
    }
}

fn assert_requirements(
    report: codec::DecodeReport,
    requirements: codec::RewriteExecutionRequirements,
) {
    assert_eq!(report.output_bytes(), requirements.output_bytes());
    assert_eq!(report.fields(), requirements.fields());
    assert_eq!(report.work_bytes(), requirements.work_bytes());
    assert_eq!(report.max_depth(), requirements.max_depth());
    assert_eq!(report.references(), requirements.references());
    assert_eq!(report.allocations(), requirements.allocations());
    assert_eq!(report.retained_bytes(), requirements.retained_bytes());
    assert_eq!(report.scratch_bytes(), requirements.scratch_bytes());
}

fn exercise_limit_profiles(source: &[u8]) {
    if source.is_empty() {
        return;
    }
    let input_limit = codec::DecodeOptions::new(
        source.len() - 1,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    );
    observe_result(codec::decode_control_cell_spec(source, input_limit));
    for (fields, work, depth, references) in [
        (0, MAX_WORK_BYTES, MAX_RECURSION, MAX_REFERENCES),
        (MAX_FIELDS, 0, MAX_RECURSION, MAX_REFERENCES),
        (MAX_FIELDS, MAX_WORK_BYTES, 0, MAX_REFERENCES),
        (MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION, 0),
    ] {
        let constrained = codec::DecodeOptions::new(
            MAX_INPUT_BYTES,
            MAX_OUTPUT_BYTES,
            fields,
            work,
            depth,
            references,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        );
        observe_result(codec::decode_control_cell_spec(source, constrained));
        observe_result(codec::decode_control_format(source, constrained));
    }
}

fn fixed_cases() -> Vec<Vec<u8>> {
    let mut cases = vec![
        vec![0x08, 0x08],
        fixed_range(6, 0.0, 5.0, 1.0),
        fixed_range(5, 0.0, 100.0, 10.0),
        fixed_range(4, 1.0, 10.0, 1.0),
        vec![0x08, 0x07, 0x32, 0x02, 0x08, 0x2a, 0x38, 0x01],
        vec![0x08, 0x08, 0x08, 0x08],
        vec![0x0a, 0x01, 0x08],
        vec![0x08, 0x88, 0x00],
        vec![0x08, 0x05, 0x19, 0, 0, 0, 0, 0, 0, 0xf8, 0x7f],
        vec![0x08, 0x87, 0x02],
        vec![0x08, 0x08, 0x12, 0x01, 0x00],
        vec![
            0x08, 0x08, 0x93, 0x05, 0x98, 0x06, 0x81, 0x80, 0x00, 0x94, 0x05,
        ],
        vec![0x08],
    ];
    cases.push(deep_unknown_group(72));
    cases.push(deep_unknown_group(8));
    cases
}

fn fixed_range(interaction: u8, minimum: f64, maximum: f64, increment: f64) -> Vec<u8> {
    let mut output = vec![0x08, interaction];
    append_fixed64(&mut output, 3, minimum.to_bits());
    append_fixed64(&mut output, 4, maximum.to_bits());
    append_fixed64(&mut output, 5, increment.to_bits());
    output
}

fn deep_unknown_group(depth: usize) -> Vec<u8> {
    let mut output = Vec::new();
    for _ in 0..depth {
        append_key(&mut output, 90, 3);
    }
    append_key(&mut output, 99, 0);
    append_varint(&mut output, 7);
    for _ in 0..depth {
        append_key(&mut output, 90, 4);
    }
    output.extend_from_slice(&[0x08, 0x08]);
    output
}

fn append_key(output: &mut Vec<u8>, field: u32, wire: u8) {
    append_varint(output, (u64::from(field) << 3) | u64::from(wire));
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn append_fixed64(output: &mut Vec<u8>, field: u32, value: u64) {
    append_key(output, field, 1);
    output.extend_from_slice(&value.to_le_bytes());
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

fn observe_result<T>(result: Result<T, codec::DecodeError>)
where
    T: Debug,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => {
            observe_error(error);
        },
    }
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
    black_box(error.resource_limit());
    black_box(error.allocation_requested());
}
