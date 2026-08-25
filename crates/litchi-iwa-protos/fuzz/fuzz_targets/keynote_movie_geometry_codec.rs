#![no_main]

//! Bounded fuzzing for the strict `TSD.MovieArchive` geometry projection.
//!
//! The target deliberately exercises the prepared path as well as the
//! convenience rewrite.  Every successful candidate is decoded again and
//! every failed limit replay is required to leave the borrowed source alone.

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_movie_geometry_codec::{
    DecodeOptions, MovieGeometryWrite, RewriteExecutionLimits, decode_movie_geometry_with_report,
    prepare_movie_geometry_rewrite, rewrite_movie_geometry,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const CANONICAL_SOURCE: &[u8] = &[
    0x0a, 0x1c, 0x0a, 0x1a, 0x0a, 0x18, 0x0a, 0x0a, 0x0d, 0x00, 0x00, 0xc8, 0x42, 0x15, 0x00, 0x00,
    0xf0, 0x42, 0x12, 0x0a, 0x0d, 0x00, 0x00, 0x20, 0x44, 0x15, 0x00, 0x00, 0xb4, 0x43,
];

// These inputs are intentionally small and independent of any package seed.
// They keep unknown groups, unknown overlong scalars, truncation, and deep
// nesting in every campaign even when arbitrary bytes do not reach a valid
// MovieArchive envelope.
const FIXED_CASES: &[&[u8]] = &[
    CANONICAL_SOURCE,
    &[],
    // Unknown balanced group followed by an unknown overlong scalar.
    &[0x53, 0x08, 0x01, 0x54, 0x80, 0x01, 0x80, 0x00],
    // A truncated length-delimited envelope.
    &[0x0a, 0x08, 0x01, 0x02],
    // Wrong wire type and truncated fixed-width value.
    &[0x0b, 0x25, 0x00, 0x00],
    &[0x25, 0x00, 0x00],
    // Unterminated unknown group.
    &[0x53, 0x08, 0x01],
];

fuzz_target!(|data: &[u8]| {
    if data.len() <= MAX_INPUT_BYTES {
        exercise_source(data);
    }

    static FIXED: OnceLock<()> = OnceLock::new();
    FIXED.get_or_init(|| {
        for source in FIXED_CASES {
            exercise_source(source);
        }
        exercise_deep_groups();
        exercise_limit_guards();
    });
});

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
}

fn write_for(data: &[u8]) -> MovieGeometryWrite {
    let byte = |offset: usize| f32::from(data.get(offset).copied().unwrap_or_default());
    // Keep all generated values finite and strictly positive for dimensions.
    MovieGeometryWrite::from_values(byte(0) + 1.0, byte(1) + 1.0, byte(2) + 1.0, byte(3) + 1.0)
}

fn exercise_source(source: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let decoded = decode_movie_geometry_with_report(source, decode_options);
    assert_eq!(source, original.as_slice(), "decode modified its source");
    let Ok((snapshot, report)) = decoded else {
        return;
    };
    assert_eq!(report.input_bytes(), source.len());
    black_box(snapshot);

    let write = write_for(source);
    let prepared = match prepare_movie_geometry_rewrite(source, write, decode_options) {
        Ok(prepared) => prepared,
        Err(_) => return,
    };
    let requirements = prepared.execution_requirements();
    let prepare_report = prepared.prepare_report();
    assert_eq!(prepare_report.input_bytes(), source.len());
    assert!(prepare_report.fields() <= requirements.fields);
    assert!(prepare_report.work_bytes() <= requirements.work_bytes);
    assert_eq!(source, original.as_slice(), "prepare modified its source");

    let output = prepared
        .execute(RewriteExecutionLimits::exact(requirements))
        .expect("exact geometry requirements must execute");
    assert_eq!(output.as_bytes().len(), requirements.output_bytes);
    assert_eq!(output.report().fields(), requirements.fields);
    assert_eq!(output.report().work_bytes(), requirements.work_bytes);
    let candidate = output.as_bytes().to_vec();
    let one_shot = rewrite_movie_geometry(source, write, decode_options)
        .expect("one-shot geometry rewrite must match prepared execution");
    assert_eq!(one_shot, candidate);

    let (readback, readback_report) =
        decode_movie_geometry_with_report(&candidate, options(&candidate))
            .expect("prepared geometry candidate must be readable");
    black_box(readback);
    assert_eq!(readback_report.input_bytes(), candidate.len());
    assert_eq!(source, original.as_slice(), "rewrite modified its source");

    for limits in [
        RewriteExecutionLimits::exact(requirements)
            .with_output_bytes(requirements.output_bytes.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_fields(requirements.fields.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_work_bytes(requirements.work_bytes.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_max_depth(requirements.max_depth.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_allocations(requirements.allocations.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_retained_bytes(requirements.retained_bytes.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1)),
    ] {
        let result = prepared.execute(limits);
        assert!(
            result.is_err(),
            "a max-minus-one geometry ceiling succeeded"
        );
        let _ = black_box(result);
    }
}

fn exercise_deep_groups() {
    let mut source = Vec::new();
    for _ in 0..80 {
        source.extend_from_slice(&[0x53]);
    }
    source.extend_from_slice(&[0x08, 0x01]);
    for _ in 0..80 {
        source.extend_from_slice(&[0x54]);
    }
    let result = decode_movie_geometry_with_report(&source, options(&source));
    assert!(
        result.is_err(),
        "deep unknown groups bypassed the nesting ceiling"
    );
    let _ = black_box(result);
}

fn exercise_limit_guards() {
    let source = CANONICAL_SOURCE;
    let decode = decode_movie_geometry_with_report(source, options(source))
        .expect("canonical geometry seed must decode");
    let report = decode.1;
    let limits = [
        DecodeOptions::new(
            source.len().saturating_sub(1),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
        ),
        DecodeOptions::new(
            source.len(),
            report.fields().saturating_sub(1),
            MAX_WORK_BYTES,
            MAX_RECURSION,
        ),
        DecodeOptions::new(
            source.len(),
            MAX_FIELDS,
            report.work_bytes().saturating_sub(1),
            MAX_RECURSION,
        ),
        DecodeOptions::new(
            source.len(),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            report.max_depth().saturating_sub(1),
        ),
        DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION)
            .with_max_scratch_bytes(report.scratch_bytes().saturating_sub(1)),
    ];
    for decode_options in limits {
        let result = decode_movie_geometry_with_report(source, decode_options);
        assert!(result.is_err(), "a max-minus-one decode ceiling succeeded");
        let _ = black_box(result);
    }
}
