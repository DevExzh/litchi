#![no_main]

//! Strict, source-preserving fuzzing for the native Numbers Percentage
//! `FormatStructArchive` projection.
//!
//! The Number and Percentage wire shapes are intentionally similar, but their
//! native format types are not interchangeable.  This target keeps the
//! Percentage wrapper honest: successful snapshots must borrow the caller's
//! source, malformed and cross-family messages must be rejected, and both
//! source-preserving and canonical writes must replay under exact limits.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_percentage_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const FIXED_CASES: &[&[u8]] = &[
    // Native Percentage: fixed two places, minus-sign negatives, no
    // separator.
    &[0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00],
    // Native Percentage: automatic places and a visible separator.
    &[0x08, 0x82, 0x02, 0x10, 0xfd, 0x01, 0x20, 0x02, 0x28, 0x01],
    // Unknown scalar and balanced group must stay source-authoritative.
    &[
        0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00, 0xa0, 0x06, 0x81, 0x00, 0xa3, 0x06,
        0xa8, 0x06, 0x01, 0xa4, 0x06,
    ],
    // An unknown length-delimited extension after the selected fields.
    &[
        0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00, 0xa2, 0x06, 0x01, 0x7f,
    ],
    // Missing a required selected field is rejected by the strict route.
    &[0x08, 0x82, 0x02, 0x20, 0x00, 0x28, 0x00],
    // Duplicate and wrong-wire selected fields are rejected.
    &[
        0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00, 0x08, 0x82, 0x02,
    ],
    &[0x0a, 0x01, 0x00, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00],
    // Currency and other known non-Percentage fields cannot be hidden as
    // opaque extensions behind the focused API.
    &[
        0x08, 0x82, 0x02, 0x10, 0x02, 0x1a, 0x00, 0x20, 0x00, 0x28, 0x00,
    ],
    // Invalid native scalar domains.
    &[0x08, 0x82, 0x02, 0x10, 0x1f, 0x20, 0x00, 0x28, 0x00],
    &[0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x04, 0x28, 0x00],
    &[0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x02],
    // A known selected varint encoded noncanonically is rejected.
    &[0x08, 0x82, 0x02, 0x10, 0x82, 0x00, 0x20, 0x00, 0x28, 0x00],
    // A valid Number message must not be accepted by the Percentage wrapper.
    &[0x08, 0x80, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00],
    // Unterminated and unmatched groups remain malformed.
    &[
        0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00, 0xa3, 0x06, 0x08, 0x01,
    ],
    &[
        0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00, 0xa4, 0x06,
    ],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    // Keep semantic and resource-boundary guards reachable even when a
    // campaign's initial inputs are all malformed.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-percentage-format");
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
        0,
        0,
        0,
    )
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_percentage_format(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "scalar decode modified its source"
    );
    let reported = codec::decode_percentage_format_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "reported decode modified its source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.raw(), source);
            assert_borrowed(source, snapshot.raw());
            assert_report(report, source.len());
            black_box((snapshot, report));
            exercise_rewrites(source, snapshot, data, &before);
            exercise_limit_profiles(source, report);
        },
        (Err(scalar_error), Err(reported_error)) => {
            assert_eq!(
                scalar_error.resource_limit(),
                reported_error.resource_limit(),
                "Percentage-format scalar/report error classifications disagree"
            );
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "Percentage-format scalar/report disagreement: scalar={:?}, report={:?}",
            scalar_result.map(|snapshot| snapshot.decimal_places()),
            reported_result.map(|(snapshot, _)| snapshot.decimal_places())
        ),
    }

    assert_eq!(
        source,
        before.as_slice(),
        "Percentage-format probes modified source"
    );
}

fn assert_borrowed(source: &[u8], borrowed: &[u8]) {
    if borrowed.is_empty() {
        return;
    }
    let source_start = source.as_ptr() as usize;
    let source_end = source_start
        .checked_add(source.len())
        .expect("bounded source pointer range");
    let borrowed_start = borrowed.as_ptr() as usize;
    let borrowed_end = borrowed_start
        .checked_add(borrowed.len())
        .expect("bounded borrowed pointer range");
    assert!(
        borrowed_start >= source_start && borrowed_end <= source_end,
        "Percentage-format snapshot did not borrow from its source"
    );
}

fn assert_report(report: codec::DecodeReport, source_len: usize) {
    assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source_len));
    assert!(report.output_bytes() <= MAX_OUTPUT_BYTES.max(source_len));
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert_eq!(report.references(), 0);
    assert_eq!(report.items(), 0);
    assert_eq!(report.text_bytes(), 0);
    assert_eq!(report.allocations(), 0);
}

fn exercise_rewrites(
    source: &[u8],
    snapshot: codec::PercentageFormatSnapshot<'_>,
    data: &[u8],
    before: &[u8],
) {
    let preserved = codec::PercentageFormatWrite::from_snapshot(snapshot);
    let requested = requested_write(data, snapshot);
    for write in [preserved, requested] {
        let prepared =
            match codec::prepare_percentage_format_rewrite(source, write, options(source)) {
                Ok(prepared) => prepared,
                Err(error) => {
                    observe_error(error);
                    continue;
                },
            };
        let requirements = prepared.execution_requirements();
        let preparation = prepared.prepare_report();
        assert_eq!(preparation.output_bytes(), requirements.output_bytes());
        assert_eq!(preparation.fields(), requirements.fields());
        assert_eq!(preparation.work_bytes(), requirements.work_bytes());
        assert_eq!(preparation.max_depth(), requirements.max_depth());
        assert_eq!(preparation.allocations(), requirements.allocations());
        assert_eq!(preparation.retained_bytes(), requirements.retained_bytes());
        assert_eq!(preparation.scratch_bytes(), requirements.scratch_bytes());
        assert_eq!(source, before, "prepare modified its source");

        let output = prepared
            .execute(codec::RewriteExecutionLimits::exact(requirements))
            .unwrap_or_else(|error| panic!("exact Percentage-format replay failed: {error}"));
        assert_eq!(source, before, "execute modified its source");
        assert_eq!(output.bytes().len(), requirements.output_bytes());
        assert_eq!(output.report().output_bytes(), requirements.output_bytes());
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
        assert_eq!(output.report().max_depth(), requirements.max_depth());
        assert_eq!(output.report().allocations(), requirements.allocations());
        assert_eq!(
            output.report().retained_bytes(),
            requirements.retained_bytes()
        );
        assert_eq!(
            output.report().scratch_bytes(),
            requirements.scratch_bytes()
        );

        let readback = codec::decode_percentage_format(output.bytes(), options(output.bytes()))
            .unwrap_or_else(|error| panic!("Percentage-format candidate readback failed: {error}"));
        assert_eq!(codec::PercentageFormatWrite::from_snapshot(readback), write);
        if write == preserved {
            assert_eq!(output.bytes(), source, "preserving rewrite was not a no-op");
        }
        assert_known_unknown_spans_preserved(source, output.bytes());

        let one_shot = codec::rewrite_percentage_format(source, write, options(source))
            .unwrap_or_else(|error| panic!("one-shot Percentage-format rewrite failed: {error}"));
        assert_eq!(one_shot.bytes(), output.bytes());
        assert_eq!(one_shot.report(), output.report());
        exercise_limit_failures(source, write, prepared, requirements);
        black_box(output);
    }
}

fn requested_write(
    data: &[u8],
    snapshot: codec::PercentageFormatSnapshot<'_>,
) -> codec::PercentageFormatWrite {
    let decimal_places = match data.first().copied().unwrap_or_default() % 4 {
        0 => snapshot.decimal_places(),
        1 => 0,
        2 => codec::MAX_PERCENTAGE_DECIMAL_PLACES,
        _ => codec::NATIVE_AUTOMATIC_DECIMAL_PLACES,
    };
    let negative_style = u32::from(data.get(1).copied().unwrap_or_default() % 4);
    let show_thousands_separator = data.get(2).copied().unwrap_or_default() & 1 != 0;
    codec::PercentageFormatWrite::new(decimal_places, negative_style, show_thousands_separator)
}

fn assert_known_unknown_spans_preserved(source: &[u8], candidate: &[u8]) {
    let source_shape = parse_root_shape(source)
        .expect("successful Percentage-format source must have valid root wire records");
    let candidate_shape = parse_root_shape(candidate)
        .expect("rewritten Percentage-format candidate must have valid root wire records");

    // Compare the parsed record shape, not arbitrary byte windows.  The
    // selected fields may change their values, but every unknown root span
    // must retain its exact bytes, relative order, and multiplicity.  Keeping
    // selected field numbers in the shape also catches an implementation that
    // accidentally moves an extension across a known field while rewriting.
    assert_eq!(
        source_shape, candidate_shape,
        "Percentage-format rewrite changed parsed unknown-span order or multiplicity"
    );
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RootField<'source> {
    Selected(u32),
    Unknown(&'source [u8]),
}

fn parse_root_shape(source: &[u8]) -> Option<Vec<RootField<'_>>> {
    let mut fields = Vec::new();
    let mut offset = 0usize;
    while offset < source.len() {
        let start = offset;
        let (number, wire, key_len) = read_wire_key(source, offset)?;
        offset = offset.checked_add(key_len)?;
        offset = skip_wire_value(source, offset, wire, number, 0)?;
        let raw = source.get(start..offset)?;
        let field = match number {
            1 | 2 | 4 | 5 => RootField::Selected(number),
            _ => RootField::Unknown(raw),
        };
        fields.push(field);
    }
    Some(fields)
}

fn read_wire_key(source: &[u8], offset: usize) -> Option<(u32, u8, usize)> {
    let (key, consumed, canonical) = read_wire_varint(source, offset)?;
    if !canonical {
        return None;
    }
    let number = u32::try_from(key >> 3).ok()?;
    let wire = u8::try_from(key & 7).ok()?;
    (number != 0 && number <= 0x1fff_ffff).then_some((number, wire, consumed))
}

fn skip_wire_value(
    source: &[u8],
    offset: usize,
    wire: u8,
    field_number: u32,
    depth: usize,
) -> Option<usize> {
    match wire {
        // Unknown scalar values are intentionally allowed to use a relaxed
        // varint encoding, matching the production preservation path.
        0 => {
            let (_, consumed, _) = read_wire_varint(source, offset)?;
            offset.checked_add(consumed)
        },
        1 => offset.checked_add(8).filter(|end| *end <= source.len()),
        2 => {
            let (length, consumed, canonical) = read_wire_varint(source, offset)?;
            if !canonical {
                return None;
            }
            let payload_start = offset.checked_add(consumed)?;
            let payload_len = usize::try_from(length).ok()?;
            payload_start
                .checked_add(payload_len)
                .filter(|end| *end <= source.len())
        },
        3 => skip_wire_group(source, offset, field_number, depth),
        // End-group records are consumed by skip_wire_group and are invalid
        // at the root or as a value of another non-group field.
        4 => None,
        5 => offset.checked_add(4).filter(|end| *end <= source.len()),
        _ => None,
    }
}

fn skip_wire_group(
    source: &[u8],
    mut offset: usize,
    root_number: u32,
    depth: usize,
) -> Option<usize> {
    if depth >= 64 {
        return None;
    }
    let mut stack = vec![root_number];
    while offset < source.len() {
        let (number, wire, key_len) = read_wire_key(source, offset)?;
        offset = offset.checked_add(key_len)?;
        match wire {
            3 => {
                if stack.len() >= 64 {
                    return None;
                }
                stack.push(number);
            },
            4 => {
                if stack.pop()? != number {
                    return None;
                }
                if stack.is_empty() {
                    return Some(offset);
                }
            },
            _ => {
                offset = skip_wire_value(source, offset, wire, number, depth + stack.len())?;
            },
        }
    }
    None
}

fn read_wire_varint(source: &[u8], offset: usize) -> Option<(u64, usize, bool)> {
    let mut value = 0u64;
    let mut shift = 0u32;
    let mut cursor = offset;
    while cursor < source.len() && cursor.checked_sub(offset)? < 10 {
        let byte = source[cursor];
        let part = u64::from(byte & 0x7f);
        if shift == 63 && part > 1 {
            return None;
        }
        value |= part.checked_shl(shift)?;
        cursor += 1;
        if byte & 0x80 == 0 {
            let consumed = cursor.checked_sub(offset)?;
            return Some((value, consumed, encoded_wire_varint_len(value) == consumed));
        }
        shift += 7;
    }
    None
}

const fn encoded_wire_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn exercise_limit_failures(
    source: &[u8],
    write: codec::PercentageFormatWrite,
    prepared: codec::PreparedPercentageFormatRewrite<'_>,
    requirements: codec::RewriteExecutionRequirements,
) {
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let probes = [
        (requirements.output_bytes() > 0)
            .then(|| exact.with_output_bytes(requirements.output_bytes().saturating_sub(1))),
        (requirements.fields() > 0)
            .then(|| exact.with_fields(requirements.fields().saturating_sub(1))),
        (requirements.work_bytes() > 0)
            .then(|| exact.with_work_bytes(requirements.work_bytes().saturating_sub(1))),
        (requirements.max_depth() > 0)
            .then(|| exact.with_max_depth(requirements.max_depth().saturating_sub(1))),
        (requirements.allocations() > 0)
            .then(|| exact.with_allocations(requirements.allocations().saturating_sub(1))),
        (requirements.retained_bytes() > 0)
            .then(|| exact.with_retained_bytes(requirements.retained_bytes().saturating_sub(1))),
    ];
    for limits in probes.into_iter().flatten() {
        let result = prepared.execute(limits);
        assert!(
            result.is_err(),
            "Percentage-format rewrite accepted a strict limit"
        );
        if let Err(error) = result {
            black_box(error.resource_limit());
            observe_error(error);
        }
    }

    // Re-preparation is intentionally part of this probe: it demonstrates
    // that the request and source remain reusable after every failed execute.
    let replay = codec::rewrite_percentage_format(source, write, options(source));
    if let Err(error) = replay {
        observe_error(error);
    }
}

fn exercise_limit_profiles(source: &[u8], report: codec::DecodeReport) {
    if source.is_empty() {
        return;
    }
    let input_limited = codec::DecodeOptions::new(
        source.len() - 1,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_resource_limit(
        codec::decode_percentage_format(source, input_limited),
        |limit| matches!(limit, codec::DecodeLimit::InputBytes { .. }),
    );

    if report.fields() > 0 {
        let options = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            report.fields() - 1,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_percentage_format(source, options), |limit| {
            matches!(limit, codec::DecodeLimit::Fields { .. })
        });
    }

    if report.work_bytes() > 0 {
        let options = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_FIELDS,
            report.work_bytes() - 1,
            MAX_RECURSION,
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_percentage_format(source, options), |limit| {
            matches!(limit, codec::DecodeLimit::Work { .. })
        });
    }
}

fn assert_resource_limit<T>(
    result: Result<T, codec::DecodeError>,
    predicate: impl FnOnce(codec::DecodeLimit) -> bool,
) {
    let error = match result {
        Ok(_) => panic!("resource-limited Percentage-format operation unexpectedly succeeded"),
        Err(error) => error,
    };
    let limit = error
        .resource_limit()
        .expect("resource failure did not expose a typed limit");
    assert!(
        predicate(limit),
        "unexpected Percentage-format resource limit"
    );
    observe_error(error);
}

fn exercise_canonical_writes() {
    for write in [
        codec::PercentageFormatWrite::new(0, 0, false),
        codec::PercentageFormatWrite::new(30, 3, true),
        codec::PercentageFormatWrite::new(codec::NATIVE_AUTOMATIC_DECIMAL_PLACES, 1, true),
    ] {
        let prepared = codec::prepare_percentage_format_write(write, options(&[]))
            .unwrap_or_else(|error| panic!("canonical Percentage-format prepare failed: {error}"));
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(codec::RewriteExecutionLimits::exact(requirements))
            .unwrap_or_else(|error| panic!("canonical Percentage-format execute failed: {error}"));
        assert_eq!(output.bytes().len(), requirements.output_bytes());
        assert_eq!(output.report().fields(), requirements.fields());
        let readback = codec::decode_percentage_format(output.bytes(), options(output.bytes()))
            .unwrap_or_else(|error| panic!("canonical Percentage-format readback failed: {error}"));
        assert_eq!(codec::PercentageFormatWrite::from_snapshot(readback), write);
        assert_eq!(
            codec::canonical_percentage_format(write, options(&[]))
                .expect("canonical one-shot")
                .bytes(),
            output.bytes()
        );
        black_box(output);
    }
}

fn exercise_limit_guards() {
    let source = FIXED_CASES[0];
    let write = codec::PercentageFormatWrite::new(30, 3, true);
    let mut output_limited = options(source);
    output_limited = output_limited.with_max_output_bytes(0);
    assert_resource_limit(
        codec::prepare_percentage_format_rewrite(source, write, output_limited),
        |limit| matches!(limit, codec::DecodeLimit::OutputBytes { .. }),
    );

    let prepared = codec::prepare_percentage_format_rewrite(source, write, options(source))
        .expect("limit guard prepare");
    let requirements = prepared.execution_requirements();
    assert_resource_limit(
        prepared.execute(codec::RewriteExecutionLimits::exact(requirements).with_allocations(0)),
        |limit| matches!(limit, codec::DecodeLimit::Allocation { .. }),
    );
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
