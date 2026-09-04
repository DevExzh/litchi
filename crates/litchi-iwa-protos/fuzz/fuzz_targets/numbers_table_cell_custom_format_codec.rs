#![no_main]

//! Bounded fuzzing for the Numbers document-scoped Custom format registry.
//!
//! The harness drives both archive and parallel-list entry points.  Successful
//! values are walked through their borrowed iterators, source-preserving
//! rewrites, canonical writes, and one-below finite resource ceilings.  Fixed
//! cases keep malformed groups, duplicate/zero UUID keys, family mismatches,
//! and parallel-array failures reachable independently of arbitrary input.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_custom_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_REFERENCES: usize = 128;
const MAX_ITEMS: usize = 128;
const MAX_TEXT_BYTES: usize = 16 * 1024;

const FIXED_CASES: &[&[u8]] = &[
    // Empty and malformed roots must fail closed before any lazy view.
    &[],
    &[0x0a, 0x01, 0xff],
    &[0xa3, 0x06, 0x08, 0x01],
    &[0xa3, 0x06, 0x08, 0x01, 0xac, 0x06],
    // A list with a zero UUID and a list with only one side of the pair.
    &[0x0a, 0x04, 0x08, 0x00, 0x10, 0x00],
    &[0x0a, 0x04, 0x08, 0x01, 0x10, 0x02],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source);

    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed);
        }
        exercise_canonical_writes();
        exercise_limit_guards();
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
        MAX_REFERENCES,
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    exercise_archive(source);
    assert_eq!(source, before.as_slice(), "Custom probes modified source");
    exercise_list(source);
    assert_eq!(source, before.as_slice(), "Custom list probe modified source");
}

fn exercise_archive(source: &[u8]) {
    let decode_options = options(source);
    let scalar = codec::decode_custom_format(source, decode_options);
    let reported = codec::decode_custom_format_with_report(source, decode_options);
    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.raw(), source);
            assert!(snapshot.name().len() <= codec::MAX_CUSTOM_NAME_BYTES);
            assert!(snapshot.default_format().pattern().len() <= codec::MAX_CUSTOM_PATTERN_BYTES);
            assert_borrowed(source, snapshot.name().as_bytes());
            assert_borrowed(source, snapshot.default_format().pattern().as_bytes());
            let conditions = snapshot
                .conditions()
                .collect::<Result<Vec<_>, _>>()
                .expect("validated condition iterator");
            assert_eq!(conditions.len(), snapshot.condition_count());
            assert_eq!(report.input_bytes(), source.len());
            black_box((snapshot, report));
            exercise_archive_rewrite(source, snapshot, &conditions);
        },
        (Err(left), Err(right)) => assert_eq!(
            left.resource_limit(),
            right.resource_limit(),
            "Custom archive scalar/report errors disagree"
        ),
        (left, right) => panic!(
            "Custom archive scalar/report disagreement: scalar={left:?}, report={right:?}"
        ),
    }
}

fn exercise_archive_rewrite(
    source: &[u8],
    snapshot: codec::CustomFormatSnapshot<'_>,
    conditions: &[codec::CustomConditionSnapshot<'_>],
) {
    let writes = conditions
        .iter()
        .map(|condition| {
            codec::CustomConditionWrite::new(
                condition.condition_type(),
                condition.condition_value(),
                codec::FormatStructWrite::from_snapshot(condition.condition_format()),
                condition.condition_value_dbl(),
            )
        })
        .collect::<Vec<_>>();
    let write = codec::CustomFormatWrite::from_snapshot(snapshot).with_conditions(&writes);
    let Ok(prepared) = codec::prepare_custom_format_rewrite(source, write, options(source)) else {
        return;
    };
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .expect("validated Custom archive rewrite");
    let rewritten = codec::decode_custom_format(output.bytes(), options(output.bytes()))
        .expect("rewritten Custom archive");
    assert_eq!(rewritten.name(), snapshot.name());
    assert_eq!(output.report().fields(), requirements.fields());
    assert!(output.bytes().len() <= MAX_OUTPUT_BYTES.max(source.len()));
}

fn exercise_list(source: &[u8]) {
    let decode_options = options(source);
    let scalar = codec::decode_custom_format_list(source, decode_options);
    let reported = codec::decode_custom_format_list_with_report(source, decode_options);
    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.uuid_count(), snapshot.format_count());
            let uuids = snapshot
                .uuids()
                .collect::<Result<Vec<_>, _>>()
                .expect("validated UUID iterator");
            let archives = snapshot
                .custom_formats()
                .collect::<Result<Vec<_>, _>>()
                .expect("validated archive iterator");
            assert_eq!(uuids.len(), snapshot.uuid_count());
            assert_eq!(archives.len(), snapshot.format_count());
            for (left, right) in uuids.iter().enumerate() {
                assert_ne!(right.lower(), 0);
                assert_ne!(right.upper(), 0);
                assert!(!uuids[..left].contains(right));
            }
            assert_eq!(report.input_bytes(), source.len());
            black_box((snapshot, report));
            exercise_list_rewrite(source, &uuids, &archives);
        },
        (Err(left), Err(right)) => assert_eq!(
            left.resource_limit(),
            right.resource_limit(),
            "Custom list scalar/report errors disagree"
        ),
        (left, right) => panic!(
            "Custom list scalar/report disagreement: scalar={left:?}, report={right:?}"
        ),
    }
}

fn exercise_list_rewrite(
    source: &[u8],
    uuids: &[codec::Uuid],
    archives: &[codec::CustomFormatSnapshot<'_>],
) {
    let mut condition_sets = Vec::with_capacity(archives.len());
    for archive in archives.iter().copied() {
        let conditions = archive
            .conditions()
            .map(|condition| {
                let condition = condition?;
                Ok(codec::CustomConditionWrite::new(
                    condition.condition_type(),
                    condition.condition_value(),
                    codec::FormatStructWrite::from_snapshot(condition.condition_format()),
                    condition.condition_value_dbl(),
                ))
            })
            .collect::<Result<Vec<_>, codec::DecodeError>>();
        let Ok(conditions) = conditions else {
            return;
        };
        condition_sets.push(conditions);
    }
    let writes = archives
        .iter()
        .copied()
        .zip(condition_sets.iter())
        .map(|(archive, conditions)| {
            codec::CustomFormatWrite::from_snapshot(archive).with_conditions(conditions)
        })
        .collect::<Vec<_>>();
    let write = codec::CustomFormatListWrite::new(uuids, &writes);
    let Ok(prepared) = codec::prepare_custom_format_list_rewrite(source, write, options(source))
    else {
        return;
    };
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .expect("validated Custom list rewrite");
    let rewritten = codec::decode_custom_format_list(output.bytes(), options(output.bytes()))
        .expect("rewritten Custom list");
    assert_eq!(rewritten.uuid_count(), uuids.len());
    assert_eq!(output.report().fields(), requirements.fields());
}

fn exercise_canonical_writes() {
    let default = codec::FormatStructWrite::new(
        codec::NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
        "#,##0.00",
    );
    let condition = codec::CustomConditionWrite::with_double(1, 5.0, default);
    let conditions = [condition];
    let write = codec::CustomFormatWrite::new(
        "Custom",
        codec::NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
        default,
        &conditions,
    );
    let output = codec::canonical_custom_format(write, options(&[])).expect("canonical archive");
    let snapshot = codec::decode_custom_format(output.bytes(), options(output.bytes()))
        .expect("canonical archive readback");
    assert_eq!(snapshot.name(), "Custom");
    assert_eq!(snapshot.condition_count(), 1);

    let uuids = [codec::Uuid::new(1, 2)];
    let formats = [write];
    let list = codec::CustomFormatListWrite::new(&uuids, &formats);
    let output = codec::canonical_custom_format_list(list, options(&[])).expect("canonical list");
    let snapshot = codec::decode_custom_format_list(output.bytes(), options(output.bytes()))
        .expect("canonical list readback");
    assert_eq!(snapshot.uuid_count(), 1);
    assert_eq!(snapshot.format_count(), 1);
    black_box(output);
}

fn exercise_limit_guards() {
    let default = codec::FormatStructWrite::new(
        codec::NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
        "#,##0.00",
    );
    let write = codec::CustomFormatWrite::new(
        "Custom",
        codec::NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
        default,
        &[],
    );
    let uuids = [codec::Uuid::new(1, 2)];
    let formats = [write];
    let source = codec::canonical_custom_format_list(
        codec::CustomFormatListWrite::new(&uuids, &formats),
        options(&[]),
    )
    .expect("limit fixture")
    .into_bytes();
    let error = codec::decode_custom_format_list(&source, options(&source).with_max_references(0))
        .expect_err("reference ceiling");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::References { .. })
    ));
}

fn assert_borrowed(source: &[u8], value: &[u8]) {
    let source_start = source.as_ptr() as usize;
    let source_end = source_start.saturating_add(source.len());
    let value_start = value.as_ptr() as usize;
    let value_end = value_start.saturating_add(value.len());
    assert!(value_start >= source_start && value_end <= source_end);
}
