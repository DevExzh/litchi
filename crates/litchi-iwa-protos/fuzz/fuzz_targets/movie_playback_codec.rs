#![no_main]

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::movie_playback_codec::{
    DecodeOptions, MoviePlaybackWrite, RewriteExecutionLimits, decode_movie_playback_with_report,
    prepare_movie_playback_rewrite, rewrite_movie_playback,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

// Small fixed cases keep strict malformed and unknown-wire paths exercised
// even when CRC-protected package fuzzing does not discover a valid payload.
const FIXED_CASES: &[&[u8]] = &[
    &[],
    // Required super envelope with an empty payload and a finite end time.
    &[0x0a, 0x00, 0x25, 0x00, 0x00, 0x80, 0x3f],
    // Duplicate end time.
    &[
        0x0a, 0x00, 0x25, 0x00, 0x00, 0x80, 0x3f, 0x25, 0x00, 0x00, 0x80, 0x3f,
    ],
    // Wrong wire type for a fixed32 field.
    &[0x0a, 0x00, 0x20, 0x01],
    // Truncated fixed32 field.
    &[0x0a, 0x00, 0x25, 0x00],
    // Unknown balanced group followed by a known field.
    &[0x53, 0x08, 0x01, 0x54, 0x25, 0x00, 0x00, 0x80, 0x3f],
    // Unknown overlong scalar followed by a known field.
    &[0x80, 0x01, 0x80, 0x00, 0x25, 0x00, 0x00, 0x80, 0x3f],
    // Unterminated group.
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

fn exercise_source(source: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let decoded = decode_movie_playback_with_report(source, decode_options);
    assert_eq!(source, original.as_slice(), "decode modified its source");
    let Ok((snapshot, report)) = decoded else {
        return;
    };
    assert_eq!(report.input_bytes(), source.len());
    black_box((snapshot.end_time, snapshot.start_time, snapshot.poster_time));

    let write = MoviePlaybackWrite::from_values(
        snapshot.start_time,
        snapshot.end_time,
        snapshot.poster_time,
        snapshot.loop_mode,
        snapshot.volume,
    );
    let prepared = match prepare_movie_playback_rewrite(source, write, decode_options) {
        Ok(prepared) => prepared,
        Err(_) => return,
    };
    let requirements = prepared.execution_requirements();
    let prepare_report = prepared.prepare_report();
    assert_eq!(prepare_report.input_bytes(), source.len());
    assert_eq!(prepare_report.fields(), requirements.fields);
    assert!(prepare_report.work_bytes() <= requirements.work_bytes);
    assert_eq!(source, original.as_slice(), "prepare modified its source");

    let output = prepared
        .execute(RewriteExecutionLimits::exact(requirements))
        .expect("exact playback requirements must execute");
    assert_eq!(output.as_bytes().len(), requirements.output_bytes);
    assert_eq!(output.report().fields(), requirements.fields);
    assert_eq!(output.report().work_bytes(), requirements.work_bytes);
    let candidate = output.as_bytes().to_vec();
    let one_shot = rewrite_movie_playback(source, write, decode_options)
        .expect("one-shot playback rewrite must match prepared execution");
    assert_eq!(one_shot, candidate);
    let changed_end = snapshot.end_time + 0.5;
    if changed_end.is_finite() && changed_end > snapshot.start_time.unwrap_or(0.0) {
        let changed_write = MoviePlaybackWrite::from_values(
            snapshot.start_time,
            changed_end,
            snapshot.poster_time,
            snapshot.loop_mode,
            snapshot.volume,
        );
        if let Ok(changed) = rewrite_movie_playback(source, changed_write, decode_options) {
            assert_eq!(
                source,
                original.as_slice(),
                "changed rewrite modified its source"
            );
            let changed_readback = decode_movie_playback_with_report(&changed, options(&changed))
                .expect("changed playback candidate must be readable")
                .0;
            assert_eq!(changed_readback.end_time, changed_end);
            black_box(changed);
        }
    }
    let (readback, readback_report) =
        decode_movie_playback_with_report(&candidate, options(&candidate))
            .expect("prepared playback candidate must be readable");
    assert_eq!(readback, snapshot);
    assert_eq!(readback_report.input_bytes(), candidate.len());
    assert_eq!(source, original.as_slice(), "rewrite modified its source");

    let one_below = |value: usize| value.saturating_sub(1);
    for limits in [
        RewriteExecutionLimits::exact(requirements)
            .with_output_bytes(one_below(requirements.output_bytes)),
        RewriteExecutionLimits::exact(requirements).with_fields(one_below(requirements.fields)),
        RewriteExecutionLimits::exact(requirements)
            .with_work_bytes(one_below(requirements.work_bytes)),
        RewriteExecutionLimits {
            max_depth: requirements.max_depth.saturating_sub(1),
            ..RewriteExecutionLimits::exact(requirements)
        },
        RewriteExecutionLimits::exact(requirements)
            .with_allocations(requirements.allocations.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_output_bytes(requirements.retained_bytes.saturating_sub(1)),
        RewriteExecutionLimits::exact(requirements)
            .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1)),
    ] {
        let result = prepared.execute(limits);
        if requirements.output_bytes > 0
            || requirements.fields > 0
            || requirements.work_bytes > 0
            || requirements.allocations > 0
            || requirements.retained_bytes > 0
            || requirements.scratch_bytes > 0
        {
            assert!(
                result.is_err(),
                "a max-minus-one playback ceiling succeeded"
            );
        }
        let _ = black_box(result);
    }
}

fn exercise_limit_guards() {
    let source = FIXED_CASES[1];
    let too_small = DecodeOptions::new(1, MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION)
        .with_max_output_bytes(MAX_OUTPUT_BYTES);
    let result = decode_movie_playback_with_report(source, too_small);
    assert!(result.is_err());
    if let Err(error) = result {
        black_box((error.limit_kind(), error.limit_values()));
    }
}
