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
    DecodeOptions, MovieGeometryWrite, MovieTransformWrite, RewriteExecutionLimits,
    decode_movie_geometry_with_report, decode_movie_transform_with_report,
    prepare_movie_geometry_rewrite, prepare_movie_transform_rewrite, rewrite_movie_geometry,
    rewrite_movie_transform,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const REFLECTION_FLAG: u32 = 1 << 2;

const CANONICAL_SOURCE: &[u8] = &[
    0x0a, 0x1a, 0x0a, 0x18, 0x0a, 0x0a, 0x0d, 0x00, 0x00, 0xc8, 0x42, 0x15, 0x00, 0x00, 0xf0, 0x42,
    0x12, 0x0a, 0x0d, 0x00, 0x00, 0x20, 0x44, 0x15, 0x00, 0x00, 0xb4, 0x43,
];

// The optional native transform fields are intentionally included in this
// fixed source so that every campaign reaches their strict projection even
// when arbitrary input does not.  The geometry codec owns only the known
// scalar values; unknown fields remain source spans.
const TRANSFORM_SOURCE: &[u8] = &[
    0x0a, 0x21, 0x0a, 0x1f, 0x0a, 0x0a, 0x0d, 0x00, 0x00, 0xc8, 0x42, 0x15, 0x00, 0x00, 0xf0, 0x42,
    0x12, 0x0a, 0x0d, 0x00, 0x00, 0x20, 0x44, 0x15, 0x00, 0x00, 0xb4, 0x43, 0x18, 0x07, 0x25, 0x00,
    0x00, 0x80, 0x3f,
];

// These inputs are intentionally small and independent of any package seed.
// They keep unknown groups, unknown overlong scalars, truncation, and deep
// nesting in every campaign even when arbitrary bytes do not reach a valid
// MovieArchive envelope.
const FIXED_CASES: &[&[u8]] = &[
    CANONICAL_SOURCE,
    TRANSFORM_SOURCE,
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
    // A duplicate optional transform field must be rejected as malformed.
    &[
        0x0a, 0x23, 0x0a, 0x21, 0x0a, 0x0a, 0x0d, 0x00, 0x00, 0xc8, 0x42, 0x15, 0x00, 0x00, 0xf0,
        0x42, 0x12, 0x0a, 0x0d, 0x00, 0x00, 0x20, 0x44, 0x15, 0x00, 0x00, 0xb4, 0x43, 0x18, 0x07,
        0x18, 0x08, 0x25, 0x00, 0x00, 0x80, 0x3f,
    ],
    // A canonical envelope with a non-finite optional angle.
    &[
        0x0a, 0x21, 0x0a, 0x1f, 0x0a, 0x0a, 0x0d, 0x00, 0x00, 0xc8, 0x42, 0x15, 0x00, 0x00, 0xf0,
        0x42, 0x12, 0x0a, 0x0d, 0x00, 0x00, 0x20, 0x44, 0x15, 0x00, 0x00, 0xb4, 0x43, 0x18, 0x07,
        0x25, 0x00, 0x00, 0xc0, 0x7f,
    ],
];

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data)
        && source.len() <= MAX_INPUT_BYTES
    {
        exercise_source(&source);
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

    exercise_transform(source, source);

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

fn transform_write_for(
    snapshot: litchi_iwa_protos::keynote_movie_geometry_codec::MovieTransformSnapshot,
    data: &[u8],
) -> MovieTransformWrite {
    let current_flags = snapshot.flags().unwrap_or_default();
    // Toggle only the reflection bit while carrying every other native flag
    // bit through unchanged.  The angle remains finite and deterministic.
    let flags = current_flags ^ REFLECTION_FLAG;
    let angle = f32::from(data.first().copied().unwrap_or_default()) - 64.0;
    MovieTransformWrite::from_values(Some(flags), Some(angle))
}

fn exercise_transform(source: &[u8], data: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let decoded = decode_movie_transform_with_report(source, decode_options);
    assert_eq!(
        source,
        original.as_slice(),
        "transform decode modified its source"
    );
    let Ok((snapshot, report)) = decoded else {
        return;
    };
    assert_eq!(report.input_bytes(), source.len());

    // Preserve is the typed no-op path and must copy all optional-field
    // presence/framing, including unknown native flag bits.
    let preserved = match prepare_movie_transform_rewrite(
        source,
        MovieTransformWrite::preserve(),
        decode_options,
    ) {
        Ok(prepared) => prepared,
        Err(_) => return,
    };
    let preserved_requirements = preserved.execution_requirements();
    let preserved_output = preserved
        .execute(RewriteExecutionLimits::exact(preserved_requirements))
        .expect("preserve transform rewrite must execute");
    assert_eq!(preserved_output.as_bytes(), original.as_slice());
    assert!(!preserved_output.report().changed());
    let cleared = MovieTransformWrite::new(None, None);
    let cleared_output = rewrite_movie_transform(source, cleared, decode_options)
        .expect("clearing transform fields must execute");
    let (cleared_snapshot, _) =
        decode_movie_transform_with_report(&cleared_output, options(&cleared_output))
            .expect("cleared transform candidate must decode");
    assert_eq!(cleared_snapshot.flags(), None);
    assert_eq!(cleared_snapshot.angle_degrees(), None);

    let write = transform_write_for(snapshot, data);
    let prepared = match prepare_movie_transform_rewrite(source, write, decode_options) {
        Ok(prepared) => prepared,
        Err(_) => return,
    };
    let requirements = prepared.execution_requirements();
    let prepare_report = prepared.prepare_report();
    assert_eq!(prepare_report.input_bytes(), source.len());
    assert!(prepare_report.fields() <= requirements.fields);
    assert!(prepare_report.work_bytes() <= requirements.work_bytes);
    assert_eq!(
        source,
        original.as_slice(),
        "transform prepare modified its source"
    );

    let output = prepared
        .execute(RewriteExecutionLimits::exact(requirements))
        .expect("exact transform requirements must execute");
    assert_eq!(output.as_bytes().len(), requirements.output_bytes);
    assert_eq!(output.report().fields(), requirements.fields);
    assert_eq!(output.report().work_bytes(), requirements.work_bytes);
    let candidate = output.as_bytes().to_vec();
    let one_shot = rewrite_movie_transform(source, write, decode_options)
        .expect("one-shot transform rewrite must match prepared execution");
    assert_eq!(one_shot, candidate);

    let (readback, readback_report) =
        decode_movie_transform_with_report(&candidate, options(&candidate))
            .expect("prepared transform candidate must be readable");
    assert_eq!(readback.flags(), Some(write_flags(write)));
    assert_eq!(readback.angle_degrees(), Some(write_angle(write)));
    assert_eq!(
        readback.flags().unwrap_or_default() & !REFLECTION_FLAG,
        snapshot.flags().unwrap_or_default() & !REFLECTION_FLAG,
        "transform rewrite changed opaque native flag bits"
    );
    assert_eq!(readback_report.input_bytes(), candidate.len());
    assert_eq!(
        source,
        original.as_slice(),
        "transform rewrite modified its source"
    );

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
            "a max-minus-one transform ceiling succeeded"
        );
        let _ = black_box(result);
    }
}

fn write_flags(write: MovieTransformWrite) -> u32 {
    match write.flags_update() {
        litchi_iwa_protos::keynote_movie_geometry_codec::TransformField::Set(value) => value,
        _ => unreachable!("transform fuzz write always sets flags"),
    }
}

fn write_angle(write: MovieTransformWrite) -> f32 {
    match write.angle_update() {
        litchi_iwa_protos::keynote_movie_geometry_codec::TransformField::Set(value) => value,
        _ => unreachable!("transform fuzz write always sets angle"),
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
