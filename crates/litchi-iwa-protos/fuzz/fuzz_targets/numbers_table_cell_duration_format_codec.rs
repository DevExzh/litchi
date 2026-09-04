#![no_main]

//! Strict, source-preserving fuzzing for the native Numbers Duration
//! `FormatStructArchive` projection.
//!
//! Duration has a deliberately small semantic surface, but it shares its
//! envelope with every other Numbers display format. This target therefore
//! keeps the family boundary, canonical selected fields, opaque unknown wire
//! records, Buffa parity, and every finite preflight/execution budget hot.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_duration_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const SELECTED_FIELDS: &[u32] = &[1, 7, 15, 16, 40];
const STYLES: [codec::DurationStyle; 3] = [
    codec::DurationStyle::Colon,
    codec::DurationStyle::Abbreviated,
    codec::DurationStyle::FullNames,
];
const UNITS: [codec::DurationUnit; 6] = [
    codec::DurationUnit::Weeks,
    codec::DurationUnit::Days,
    codec::DurationUnit::Hours,
    codec::DurationUnit::Minutes,
    codec::DurationUnit::Seconds,
    codec::DurationUnit::Milliseconds,
];

const VALID: &[u8] = &[
    0x08, 0x8c, 0x02, // format_type = 268
    0x38, 0x01, // duration_style = abbreviated
    0x78, 0x04, // duration_unit_largest = hours
    0x80, 0x01, 0x20, // duration_unit_smallest = milliseconds
    0xc0, 0x02, 0x01, // use_automatic_duration_units = true
];

const VALID_WITH_UNKNOWN: &[u8] = &[
    0x08, 0x8c, 0x02, 0xa0, 0x06, 0x81, 0x00, // unknown scalar
    0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01, 0xa3, 0x06, 0x98, 0x03, 0x09, 0xa4,
    0x06, // unknown group
];

const FIXED_CASES: &[&[u8]] = &[
    VALID,
    VALID_WITH_UNKNOWN,
    // Missing selected fields.
    &[0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20],
    &[0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01],
    // Duplicate selected fields.
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x38, 0x00, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    // Wrong wire types for field 1 and field 7.
    &[
        0x0a, 0x01, 0x01, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x3a, 0x01, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    // A known sibling field belongs to another format family.
    &[
        0x08, 0x8c, 0x02, 0x10, 0x00, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    // Invalid style, units, order, and Boolean domain.
    &[
        0x08, 0x8c, 0x02, 0x38, 0x03, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x03, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x10, 0x80, 0x01, 0x04, 0xc0, 0x02, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x02,
    ],
    // Noncanonical unknown scalar and a malformed group remain source data
    // only after the selected fields have already been parsed.
    &[
        0x08, 0x8c, 0x02, 0xa0, 0x06, 0x81, 0x00, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0,
        0x02, 0x01, 0xa3, 0x06, 0x08, 0x01,
    ],
    &[
        0x08, 0x8c, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01, 0xa3, 0x06,
        0xac, 0x06,
    ],
    // Wrong-family format discriminator.
    &[
        0x08, 0x80, 0x02, 0x38, 0x01, 0x78, 0x04, 0x80, 0x01, 0x20, 0xc0, 0x02, 0x01,
    ],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    // Keep semantic and resource-boundary assertions reachable even when a
    // campaign starts with only malformed or random bytes.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-duration-format");
        }
        exercise_all_valid_shapes();
        exercise_malformed_shapes();
        exercise_invalid_writes();
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
    let scalar = codec::decode_duration_format(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "scalar decode modified its source"
    );
    let reported = codec::decode_duration_format_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "reported decode modified its source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_snapshot(snapshot);
            assert_eq!(snapshot.raw(), source);
            assert_borrowed(source, snapshot.raw());
            assert_decode_report(report, source.len());
            black_box((snapshot, report));
            exercise_rewrites(source, snapshot, data, &before);
            exercise_limit_profiles(source, report);
        },
        (Err(scalar_error), Err(reported_error)) => {
            assert_eq!(
                scalar_error, reported_error,
                "Duration scalar/report failures are not deterministic"
            );
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "Duration scalar/report disagreement: scalar={scalar_result:?}, report={reported_result:?}"
        ),
    }

    assert_eq!(
        source,
        before.as_slice(),
        "Duration probes modified their source"
    );
}

fn assert_snapshot(snapshot: codec::DurationFormatSnapshot<'_>) {
    assert_eq!(snapshot.format_type(), codec::NATIVE_DURATION_FORMAT_TYPE);
    assert!(STYLES.contains(&snapshot.style().expect("valid Duration style")));
    assert!(UNITS.contains(&snapshot.largest_unit().expect("valid largest unit")));
    assert!(UNITS.contains(&snapshot.smallest_unit().expect("valid smallest unit")));
    assert!(snapshot.largest_unit().unwrap() <= snapshot.smallest_unit().unwrap());
    black_box(snapshot.is_automatic());
}

fn assert_decode_report(report: codec::DecodeReport, source_len: usize) {
    assert_eq!(report.input_bytes(), source_len);
    assert_eq!(report.output_bytes(), 0);
    assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source_len));
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert_eq!(report.references(), 0);
    assert_eq!(report.items(), 0);
    assert_eq!(report.text_bytes(), 0);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.retained_bytes(), 0);
    assert_eq!(report.scratch_bytes(), 0);
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
        "Duration snapshot did not borrow from its source"
    );
}

fn exercise_rewrites(
    source: &[u8],
    snapshot: codec::DurationFormatSnapshot<'_>,
    data: &[u8],
    before: &[u8],
) {
    let preserved = codec::DurationFormatWrite::from_snapshot(snapshot);
    let requested = requested_write(data, snapshot);
    for write in [preserved, requested] {
        exercise_one_rewrite(source, snapshot, write, before);
    }
}

fn exercise_one_rewrite(
    source: &[u8],
    snapshot: codec::DurationFormatSnapshot<'_>,
    write: codec::DurationFormatWrite,
    before: &[u8],
) {
    let (_, source_report) = codec::decode_duration_format_with_report(source, options(source))
        .expect("successful Duration source should report before rewrite");
    let prepared = match codec::prepare_duration_format_rewrite(source, write, options(source)) {
        Ok(prepared) => prepared,
        Err(preparation_error) => {
            let rewrite_error = codec::rewrite_duration_format(source, write, options(source))
                .expect_err("rewrite succeeded after prepare failed");
            assert_eq!(
                preparation_error, rewrite_error,
                "Duration prepare/rewrite failures are not deterministic"
            );
            observe_error(preparation_error);
            return;
        },
    };
    assert_eq!(source, before, "prepare modified its source");

    let requirements = prepared.execution_requirements();
    assert_rewrite_requirements(requirements, source.len());
    assert_prepared_report(prepared.prepare_report(), requirements);

    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact Duration rewrite failed: {error}"));
    assert_eq!(source, before, "execute modified its source");
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    assert_output_report(output.report(), requirements);

    let (readback, candidate_report) =
        codec::decode_duration_format_with_report(output.bytes(), options(output.bytes()))
            .unwrap_or_else(|error| panic!("Duration candidate readback failed: {error}"));
    assert_eq!(codec::DurationFormatWrite::from_snapshot(readback), write);
    assert_candidate_report(candidate_report, requirements, source_report.work_bytes());
    assert_known_unknown_spans_preserved(source, output.bytes());
    if write == codec::DurationFormatWrite::from_snapshot(snapshot) {
        assert_eq!(output.bytes(), source, "preserving rewrite was not a no-op");
    }

    let one_shot = codec::rewrite_duration_format(source, write, options(source))
        .unwrap_or_else(|error| panic!("one-shot Duration rewrite failed: {error}"));
    assert_eq!(one_shot.bytes(), output.bytes());
    assert_eq!(one_shot.report(), output.report());
    exercise_rewrite_limit_failures(source, write, prepared, requirements);
    black_box(output);
}

fn requested_write(
    data: &[u8],
    snapshot: codec::DurationFormatSnapshot<'_>,
) -> codec::DurationFormatWrite {
    let style_index = usize::from(data.first().copied().unwrap_or_default()) % STYLES.len();
    let largest_index = usize::from(data.get(1).copied().unwrap_or_default()) % UNITS.len();
    let remaining = UNITS.len() - largest_index;
    let smallest_index =
        largest_index + usize::from(data.get(2).copied().unwrap_or_default()) % remaining;
    let automatic = data.get(3).copied().unwrap_or_default() & 1 != 0;
    // Include source values in the selection so a fuzzer can cheaply reach
    // no-op rewrites as well as every typed writer combination.
    let style = if data.first().is_none() {
        snapshot.style().expect("valid source style")
    } else {
        STYLES[style_index]
    };
    codec::DurationFormatWrite::from_parts(
        style,
        UNITS[largest_index],
        UNITS[smallest_index],
        automatic,
    )
}

fn assert_rewrite_requirements(
    requirements: codec::RewriteExecutionRequirements,
    source_len: usize,
) {
    assert!(requirements.output_bytes() > 0);
    assert!(requirements.fields() >= SELECTED_FIELDS.len());
    assert!(requirements.work_bytes() >= requirements.output_bytes());
    assert!(requirements.max_depth() <= MAX_RECURSION);
    assert_eq!(requirements.references(), 0);
    assert_eq!(requirements.items(), 0);
    assert_eq!(requirements.text_bytes(), 0);
    assert_eq!(requirements.allocations(), 1);
    assert_eq!(requirements.scratch_bytes(), 0);
    assert_eq!(
        requirements.retained_bytes(),
        source_len
            .checked_add(requirements.output_bytes())
            .expect("bounded retained bytes")
    );
}

fn assert_prepared_report(
    report: codec::DecodeReport,
    requirements: codec::RewriteExecutionRequirements,
) {
    assert_eq!(report.input_bytes(), 0);
    assert_eq!(report.output_bytes(), requirements.output_bytes());
    assert_eq!(report.fields(), requirements.fields());
    assert_eq!(report.work_bytes(), requirements.work_bytes());
    assert_eq!(report.max_depth(), requirements.max_depth());
    assert_eq!(report.references(), requirements.references());
    assert_eq!(report.items(), requirements.items());
    assert_eq!(report.text_bytes(), requirements.text_bytes());
    assert_eq!(report.allocations(), requirements.allocations());
    assert_eq!(report.retained_bytes(), requirements.retained_bytes());
    assert_eq!(report.scratch_bytes(), requirements.scratch_bytes());
}

fn assert_output_report(
    report: codec::DecodeReport,
    requirements: codec::RewriteExecutionRequirements,
) {
    assert_eq!(report.input_bytes(), 0);
    assert_eq!(report.output_bytes(), requirements.output_bytes());
    assert_eq!(report.fields(), requirements.fields());
    assert_eq!(report.work_bytes(), requirements.work_bytes());
    assert_eq!(report.max_depth(), requirements.max_depth());
    assert_eq!(report.references(), requirements.references());
    assert_eq!(report.items(), requirements.items());
    assert_eq!(report.text_bytes(), requirements.text_bytes());
    assert_eq!(report.allocations(), requirements.allocations());
    assert_eq!(report.retained_bytes(), requirements.retained_bytes());
    assert_eq!(report.scratch_bytes(), requirements.scratch_bytes());
}

fn assert_candidate_report(
    report: codec::DecodeReport,
    requirements: codec::RewriteExecutionRequirements,
    source_work_bytes: usize,
) {
    assert_eq!(report.input_bytes(), requirements.output_bytes());
    assert_eq!(report.output_bytes(), 0);
    assert_eq!(report.fields(), requirements.fields());
    assert_eq!(
        report.work_bytes(),
        requirements
            .work_bytes()
            .checked_sub(source_work_bytes)
            .and_then(|work| work.checked_sub(requirements.output_bytes()))
            .expect("candidate work is smaller than total rewrite work")
    );
    assert_eq!(report.max_depth(), requirements.max_depth());
    assert_eq!(report.references(), 0);
    assert_eq!(report.items(), 0);
    assert_eq!(report.text_bytes(), 0);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.retained_bytes(), 0);
    assert_eq!(report.scratch_bytes(), 0);
}

fn assert_known_unknown_spans_preserved(source: &[u8], candidate: &[u8]) {
    let source_shape = parse_root_shape(source)
        .expect("successful Duration source must have valid root wire records");
    let candidate_shape = parse_root_shape(candidate)
        .expect("rewritten Duration candidate must have valid root wire records");
    assert_eq!(
        source_shape, candidate_shape,
        "Duration rewrite changed selected order or unknown-span bytes"
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
        let field = if SELECTED_FIELDS.contains(&number) {
            RootField::Selected(number)
        } else {
            RootField::Unknown(raw)
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
        // Unknown scalar values intentionally allow relaxed varints, matching
        // the production source-preservation path.
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
        // End-group records are consumed by skip_wire_group and invalid at the
        // root or as a value of another non-group field.
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
    if depth >= MAX_RECURSION as usize {
        return None;
    }
    let mut stack = [0u32; MAX_RECURSION as usize];
    stack[0] = root_number;
    let mut stack_len = 1usize;
    while offset < source.len() {
        let (number, wire, key_len) = read_wire_key(source, offset)?;
        offset = offset.checked_add(key_len)?;
        match wire {
            3 => {
                if stack_len >= MAX_RECURSION as usize {
                    return None;
                }
                stack[stack_len] = number;
                stack_len += 1;
            },
            4 => {
                if stack_len == 0 || stack[stack_len - 1] != number {
                    return None;
                }
                stack_len -= 1;
                if stack_len == 0 {
                    return Some(offset);
                }
            },
            _ => {
                offset = skip_wire_value(source, offset, wire, number, depth + stack_len)?;
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

fn exercise_rewrite_limit_failures(
    source: &[u8],
    write: codec::DurationFormatWrite,
    prepared: codec::PreparedDurationFormatRewrite<'_>,
    requirements: codec::RewriteExecutionRequirements,
) {
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let probes = [
        (requirements.output_bytes() > 0).then(|| {
            (
                exact.with_output_bytes(requirements.output_bytes().saturating_sub(1)),
                "output",
                0u8,
            )
        }),
        (requirements.fields() > 0).then(|| {
            (
                exact.with_fields(requirements.fields().saturating_sub(1)),
                "fields",
                1u8,
            )
        }),
        (requirements.work_bytes() > 0).then(|| {
            (
                exact.with_work_bytes(requirements.work_bytes().saturating_sub(1)),
                "work",
                2u8,
            )
        }),
        (requirements.max_depth() > 0).then(|| {
            (
                exact.with_max_depth(requirements.max_depth().saturating_sub(1)),
                "nesting",
                3u8,
            )
        }),
        (requirements.allocations() > 0).then(|| {
            (
                exact.with_allocations(requirements.allocations().saturating_sub(1)),
                "allocation",
                4u8,
            )
        }),
        (requirements.retained_bytes() > 0).then(|| {
            (
                exact.with_retained_bytes(requirements.retained_bytes().saturating_sub(1)),
                "retained",
                5u8,
            )
        }),
    ];
    for (limits, name, kind) in probes.into_iter().flatten() {
        let error = prepared
            .execute(limits)
            .expect_err("Duration rewrite accepted a strict one-below limit");
        let limit = error
            .resource_limit()
            .unwrap_or_else(|| panic!("Duration {name} failure had no typed limit"));
        assert!(
            matches_rewrite_limit(kind, limit),
            "unexpected Duration {name} limit: {limit:?}"
        );
        observe_error(error);
    }

    // Scratch is intentionally zero for this codec. Exercise that exact edge
    // explicitly while the other dimensions are at their requirements.
    let scratch_zero = prepared
        .execute(exact.with_scratch_bytes(0))
        .expect("zero scratch budget should execute");
    assert_eq!(scratch_zero.bytes().len(), requirements.output_bytes());
    black_box(scratch_zero);

    // A failed execution must not consume the prepared source or request.
    let replay = codec::rewrite_duration_format(source, write, options(source));
    if let Err(error) = replay {
        observe_error(error);
    }
}

fn matches_rewrite_limit(kind: u8, limit: codec::DecodeLimit) -> bool {
    match kind {
        0 => matches!(limit, codec::DecodeLimit::OutputBytes { .. }),
        1 => matches!(limit, codec::DecodeLimit::Fields { .. }),
        2 => matches!(limit, codec::DecodeLimit::Work { .. }),
        3 => matches!(limit, codec::DecodeLimit::Nesting { .. }),
        4 => matches!(limit, codec::DecodeLimit::Allocation { .. }),
        5 => matches!(limit, codec::DecodeLimit::Retained { .. }),
        _ => false,
    }
}

fn exercise_limit_profiles(source: &[u8], report: codec::DecodeReport) {
    if source.is_empty() {
        return;
    }
    let input_limited = codec::DecodeOptions::new(
        source.len() - 1,
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_decode_limit(source, input_limited, |limit| {
        matches!(limit, codec::DecodeLimit::InputBytes { .. })
    });

    if report.fields() > 0 {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            report.fields() - 1,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            0,
            0,
            0,
        );
        assert_decode_limit(source, limited, |limit| {
            matches!(limit, codec::DecodeLimit::Fields { .. })
        });
    }

    if report.work_bytes() > 0 {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_FIELDS,
            report.work_bytes() - 1,
            MAX_RECURSION,
            0,
            0,
            0,
        );
        assert_decode_limit(source, limited, |limit| {
            matches!(limit, codec::DecodeLimit::Work { .. })
        });
    }

    if report.max_depth() > 0 {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES.max(source.len()),
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            report.max_depth() - 1,
            0,
            0,
            0,
        );
        assert_decode_limit(source, limited, |limit| {
            matches!(limit, codec::DecodeLimit::Nesting { .. })
        });
    }
}

fn assert_decode_limit(
    source: &[u8],
    options: codec::DecodeOptions,
    predicate: impl FnOnce(codec::DecodeLimit) -> bool,
) {
    let scalar = codec::decode_duration_format(source, options);
    let reported = codec::decode_duration_format_with_report(source, options);
    let scalar_error = scalar.expect_err("resource-limited Duration scalar decode succeeded");
    let reported_error = reported.expect_err("resource-limited Duration report decode succeeded");
    assert_eq!(
        scalar_error, reported_error,
        "Duration scalar/report resource failures are not deterministic"
    );
    let limit = scalar_error
        .resource_limit()
        .expect("Duration resource failure did not expose a typed limit");
    assert!(
        predicate(limit),
        "unexpected Duration resource limit: {limit:?}"
    );
    observe_error(scalar_error);
    observe_error(reported_error);
}

fn exercise_all_valid_shapes() {
    for style in STYLES {
        for (largest_index, &largest) in UNITS.iter().enumerate() {
            for &smallest in UNITS.iter().skip(largest_index) {
                for automatic in [false, true] {
                    let source = duration_source(style, largest, smallest, automatic);
                    let snapshot = codec::decode_duration_format(&source, options(&source))
                        .unwrap_or_else(|error| panic!("valid Duration shape rejected: {error}"));
                    assert_snapshot(snapshot);
                    exercise_rewrites(&source, snapshot, &source, &source);
                    black_box(source);
                }
            }
        }
    }

    // Exercise every protobuf wire shape accepted as an opaque extension,
    // including a nested group and the largest legal field number.
    let source = unknown_wire_shapes_source();
    let snapshot = codec::decode_duration_format(&source, options(&source))
        .expect("unknown Duration wire shapes should decode");
    exercise_rewrites(&source, snapshot, b"unknown-wire-shapes", &source);
    let max_field = unknown_max_field_source();
    let snapshot = codec::decode_duration_format(&max_field, options(&max_field))
        .expect("maximum unknown Duration field should decode");
    exercise_rewrites(&max_field, snapshot, b"maximum-unknown-field", &max_field);
}

fn exercise_canonical_writes() {
    for style in STYLES {
        for (largest_index, &largest) in UNITS.iter().enumerate() {
            for &smallest in UNITS.iter().skip(largest_index) {
                for automatic in [false, true] {
                    let write =
                        codec::DurationFormatWrite::from_parts(style, largest, smallest, automatic);
                    let prepared = codec::prepare_duration_format_write(write, options(&[]))
                        .unwrap_or_else(|error| {
                            panic!("canonical Duration prepare failed: {error}")
                        });
                    let requirements = prepared.execution_requirements();
                    assert_rewrite_requirements(requirements, 0);
                    assert_prepared_report(prepared.prepare_report(), requirements);
                    let output = prepared
                        .execute(codec::RewriteExecutionLimits::exact(requirements))
                        .unwrap_or_else(|error| {
                            panic!("canonical Duration execute failed: {error}")
                        });
                    assert_eq!(output.bytes().len(), requirements.output_bytes());
                    assert_output_report(output.report(), requirements);
                    let (snapshot, report) = codec::decode_duration_format_with_report(
                        output.bytes(),
                        options(output.bytes()),
                    )
                    .unwrap_or_else(|error| panic!("canonical Duration readback failed: {error}"));
                    assert_eq!(codec::DurationFormatWrite::from_snapshot(snapshot), write);
                    assert_candidate_report(report, requirements, 0);
                    assert_eq!(
                        parse_root_shape(output.bytes()),
                        Some(
                            SELECTED_FIELDS
                                .iter()
                                .copied()
                                .map(RootField::Selected)
                                .collect()
                        )
                    );
                    let one_shot = codec::canonical_duration_format(write, options(&[]))
                        .expect("canonical Duration one-shot");
                    assert_eq!(one_shot.bytes(), output.bytes());
                    assert_eq!(one_shot.report(), output.report());
                    black_box(output);
                }
            }
        }
    }
}

fn exercise_malformed_shapes() {
    // Every selected field is required exactly once.
    for &omitted in SELECTED_FIELDS {
        assert_invalid(
            &duration_source_without(omitted),
            "missing selected Duration field",
        );
    }

    // Every selected field rejects duplicates and noncanonical selected
    // keys/values independently.
    for &field in SELECTED_FIELDS {
        let mut duplicate = duration_source(
            codec::DurationStyle::Abbreviated,
            codec::DurationUnit::Hours,
            codec::DurationUnit::Milliseconds,
            true,
        );
        append_varint_field(&mut duplicate, field, selected_value(field));
        assert_invalid(&duplicate, "duplicate selected Duration field");
        assert_invalid(
            &raw_duration_source(Some(field), None),
            "noncanonical selected Duration key",
        );
        assert_invalid(
            &raw_duration_source(None, Some(field)),
            "noncanonical selected Duration value",
        );
    }

    // Wrong-wire records for each selected slot.
    for &(field, wire) in &[(1, 2), (7, 1), (15, 2), (16, 5), (40, 1)] {
        let mut source = duration_source(
            codec::DurationStyle::Abbreviated,
            codec::DurationUnit::Hours,
            codec::DurationUnit::Milliseconds,
            true,
        );
        append_wire_field(
            &mut source,
            field,
            wire,
            &[0xde, 0xad, 0xbe, 0xef, 0x01, 0x02, 0x03, 0x04],
        );
        assert_invalid(&source, "wrong-wire selected Duration field");
    }

    // All other known FormatStructArchive fields belong to another family.
    for field in 2..=45 {
        if SELECTED_FIELDS.contains(&field) {
            continue;
        }
        let mut source = duration_source(
            codec::DurationStyle::Abbreviated,
            codec::DurationUnit::Hours,
            codec::DurationUnit::Milliseconds,
            true,
        );
        append_varint_field(&mut source, field, 0);
        assert_invalid(&source, "known sibling Duration field");
    }

    // The type discriminator is a strict family boundary.
    for value in [
        0,
        255,
        256,
        257,
        258,
        259,
        260,
        261,
        262,
        263,
        264,
        265,
        266,
        267,
        269,
        u64::from(u32::MAX) + 1,
    ] {
        assert_invalid(
            &duration_source_values(
                value,
                u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED),
                u64::from(codec::NATIVE_DURATION_UNIT_HOURS),
                u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS),
                1,
            ),
            "wrong-family Duration discriminator",
        );
    }

    for style in [3, 7, u64::from(u32::MAX) + 1] {
        assert_invalid(
            &duration_source_values(
                u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
                style,
                u64::from(codec::NATIVE_DURATION_UNIT_HOURS),
                u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS),
                1,
            ),
            "invalid Duration style",
        );
    }
    for unit in [0, 3, 5, 7, 33, u64::from(u32::MAX) + 1] {
        assert_invalid(
            &duration_source_values(
                u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
                u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED),
                unit,
                u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS),
                1,
            ),
            "invalid largest Duration unit",
        );
        assert_invalid(
            &duration_source_values(
                u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
                u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED),
                u64::from(codec::NATIVE_DURATION_UNIT_HOURS),
                unit,
                1,
            ),
            "invalid smallest Duration unit",
        );
    }
    for (largest_index, &largest) in UNITS.iter().enumerate() {
        for &smallest in UNITS.iter().take(largest_index) {
            assert_invalid(
                &duration_source(codec::DurationStyle::Abbreviated, largest, smallest, true),
                "inverted Duration unit order",
            );
        }
    }
    for automatic in [2, 3, u64::from(u32::MAX) + 1] {
        assert_invalid(
            &duration_source_values(
                u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
                u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED),
                u64::from(codec::NATIVE_DURATION_UNIT_HOURS),
                u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS),
                automatic,
            ),
            "invalid Duration Boolean domain",
        );
    }

    // Root end-groups, reserved wire kinds, truncation, and group mismatch.
    let valid = duration_source(
        codec::DurationStyle::Abbreviated,
        codec::DurationUnit::Hours,
        codec::DurationUnit::Milliseconds,
        true,
    );
    let mut root_end = valid.clone();
    append_raw_key(&mut root_end, 100, 4);
    assert_invalid(&root_end, "root end-group");
    for wire in [6, 7] {
        let mut reserved = valid.clone();
        append_raw_key(&mut reserved, 100, wire);
        assert_invalid(&reserved, "reserved root wire");
    }
    let mut truncated_value = valid.clone();
    append_raw_key(&mut truncated_value, 100, 0);
    truncated_value.push(0x80);
    assert_invalid(&truncated_value, "truncated unknown scalar");
    let mut truncated_length = valid.clone();
    append_raw_key(&mut truncated_length, 100, 2);
    truncated_length.push(0x80);
    assert_invalid(&truncated_length, "truncated unknown length");
    let mut mismatched_group = valid.clone();
    append_raw_key(&mut mismatched_group, 100, 3);
    append_varint_field(&mut mismatched_group, 101, 9);
    append_raw_key(&mut mismatched_group, 102, 4);
    assert_invalid(&mismatched_group, "mismatched unknown group");

    let deep = deep_group_source((MAX_RECURSION + 1) as usize);
    assert_decode_limit(&deep, options(&deep), |limit| {
        matches!(limit, codec::DecodeLimit::Nesting { .. })
    });
}

fn assert_invalid(source: &[u8], label: &str) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_duration_format(source, decode_options);
    let reported = codec::decode_duration_format_with_report(source, decode_options);
    let scalar_error = scalar.expect_err("malformed Duration scalar unexpectedly decoded");
    let reported_error = reported.expect_err("malformed Duration report unexpectedly decoded");
    assert_eq!(source, before.as_slice(), "{label} modified its source");
    assert_eq!(
        scalar_error, reported_error,
        "{label} scalar/report failures are not deterministic"
    );
    assert!(
        scalar_error.resource_limit().is_none(),
        "{label} unexpectedly hit a resource limit: {:?}",
        scalar_error.resource_limit()
    );
    let write = codec::DurationFormatWrite::from_parts(
        codec::DurationStyle::FullNames,
        codec::DurationUnit::Minutes,
        codec::DurationUnit::Seconds,
        false,
    );
    let prepared = codec::prepare_duration_format_rewrite(source, write, decode_options);
    let rewritten = codec::rewrite_duration_format(source, write, decode_options);
    let preparation_error = prepared.expect_err("malformed Duration prepared unexpectedly");
    let rewrite_error = rewritten.expect_err("malformed Duration rewrite unexpectedly");
    assert_eq!(
        preparation_error, rewrite_error,
        "{label} prepare/rewrite failures are not deterministic"
    );
    assert!(
        preparation_error.resource_limit().is_none(),
        "{label} rewrite unexpectedly hit a resource limit"
    );
    observe_error(scalar_error);
    observe_error(reported_error);
    observe_error(preparation_error);
}

fn exercise_invalid_writes() {
    let invalid_writes = [
        codec::DurationFormatWrite::new(3, 1, 1, false),
        codec::DurationFormatWrite::new(1, 3, 32, false),
        codec::DurationFormatWrite::new(1, 16, 4, false),
        codec::DurationFormatWrite::new(1, u32::MAX, 32, false),
        codec::DurationFormatWrite::new(1, 4, u32::MAX, true),
    ];
    for write in invalid_writes {
        let prepared = codec::prepare_duration_format_write(write, options(&[]));
        let canonical = codec::canonical_duration_format(write, options(&[]));
        let preparation_error = prepared.expect_err("invalid Duration write prepared");
        let canonical_error = canonical.expect_err("invalid Duration write encoded");
        assert_eq!(
            preparation_error, canonical_error,
            "invalid Duration canonical prepare/encode failures differ"
        );
        assert!(preparation_error.resource_limit().is_none());

        let prepared = codec::prepare_duration_format_rewrite(VALID, write, options(VALID));
        let rewritten = codec::rewrite_duration_format(VALID, write, options(VALID));
        let preparation_error = prepared.expect_err("invalid Duration rewrite prepared");
        let rewrite_error = rewritten.expect_err("invalid Duration rewrite encoded");
        assert_eq!(
            preparation_error, rewrite_error,
            "invalid Duration rewrite prepare/encode failures differ"
        );
        assert!(preparation_error.resource_limit().is_none());
        observe_error(preparation_error);
    }
}

fn exercise_limit_guards() {
    let source = duration_source(
        codec::DurationStyle::Colon,
        codec::DurationUnit::Weeks,
        codec::DurationUnit::Weeks,
        false,
    );
    let write = codec::DurationFormatWrite::from_parts(
        codec::DurationStyle::FullNames,
        codec::DurationUnit::Weeks,
        codec::DurationUnit::Milliseconds,
        true,
    );

    let output_limited = options(&source).with_max_output_bytes(0);
    assert_prepare_limit(
        codec::prepare_duration_format_rewrite(&source, write, output_limited),
        |limit| matches!(limit, codec::DecodeLimit::OutputBytes { .. }),
        "rewrite output ceiling",
    );

    // Canonical preparation has no source to scan, so these probes exercise
    // each check_options branch directly.
    let canonical_write = codec::DurationFormatWrite::from_parts(
        codec::DurationStyle::Colon,
        codec::DurationUnit::Weeks,
        codec::DurationUnit::Milliseconds,
        true,
    );
    let message_limited = codec::DecodeOptions::new(
        12,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_prepare_limit(
        codec::prepare_duration_format_write(canonical_write, message_limited),
        |limit| matches!(limit, codec::DecodeLimit::InputBytes { .. }),
        "canonical message ceiling",
    );
    let canonical_output_limited = codec::DecodeOptions::new(
        MAX_INPUT_BYTES,
        12,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_prepare_limit(
        codec::prepare_duration_format_write(canonical_write, canonical_output_limited),
        |limit| matches!(limit, codec::DecodeLimit::OutputBytes { .. }),
        "canonical output ceiling",
    );
    let fields_limited = codec::DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_OUTPUT_BYTES,
        4,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_prepare_limit(
        codec::prepare_duration_format_write(canonical_write, fields_limited),
        |limit| matches!(limit, codec::DecodeLimit::Fields { .. }),
        "canonical fields ceiling",
    );
    let work_limited = codec::DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        1,
        MAX_RECURSION,
        0,
        0,
        0,
    );
    assert_prepare_limit(
        codec::prepare_duration_format_write(canonical_write, work_limited),
        |limit| matches!(limit, codec::DecodeLimit::Work { .. }),
        "canonical work ceiling",
    );

    let prepared = codec::prepare_duration_format_rewrite(&source, write, options(&source))
        .expect("Duration limit guard prepare");
    let requirements = prepared.execution_requirements();
    let allocation_error = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements).with_allocations(0))
        .expect_err("Duration allocation ceiling accepted");
    assert!(matches!(
        allocation_error.resource_limit(),
        Some(codec::DecodeLimit::Allocation { .. })
    ));
    let retained_error = prepared
        .execute(
            codec::RewriteExecutionLimits::exact(requirements)
                .with_retained_bytes(requirements.retained_bytes() - 1),
        )
        .expect_err("Duration retained-byte ceiling accepted");
    assert!(matches!(
        retained_error.resource_limit(),
        Some(codec::DecodeLimit::Retained { .. })
    ));
    let scratch_output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements).with_scratch_bytes(0))
        .expect("Duration zero scratch edge failed");
    assert_eq!(scratch_output.bytes().len(), requirements.output_bytes());
    observe_error(allocation_error);
    observe_error(retained_error);
    black_box(scratch_output);

    // Invalid recursion ceilings fail identically through both decode APIs.
    for recursion in [0, MAX_RECURSION + 1] {
        let limited = codec::DecodeOptions::new(
            MAX_INPUT_BYTES,
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            recursion,
            0,
            0,
            0,
        );
        assert_decode_limit(&source, limited, |limit| {
            matches!(limit, codec::DecodeLimit::Nesting { .. })
        });
    }
}

fn assert_prepare_limit<T: std::fmt::Debug>(
    result: Result<T, codec::DecodeError>,
    predicate: impl FnOnce(codec::DecodeLimit) -> bool,
    label: &str,
) {
    let error = result.expect_err("Duration preparation unexpectedly succeeded under a limit");
    let limit = error
        .resource_limit()
        .unwrap_or_else(|| panic!("{label} did not expose a typed limit"));
    assert!(predicate(limit), "unexpected {label}: {limit:?}");
    observe_error(error);
}

fn observe_error(error: codec::DecodeError) {
    black_box(error);
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}

fn duration_source(
    style: codec::DurationStyle,
    largest: codec::DurationUnit,
    smallest: codec::DurationUnit,
    automatic: bool,
) -> Vec<u8> {
    duration_source_values(
        u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
        u64::from(style.native_value()),
        u64::from(largest.native_value()),
        u64::from(smallest.native_value()),
        u64::from(automatic),
    )
}

fn duration_source_values(
    format_type: u64,
    style: u64,
    largest: u64,
    smallest: u64,
    automatic: u64,
) -> Vec<u8> {
    let mut source = Vec::with_capacity(16);
    append_varint_field(&mut source, 1, format_type);
    append_varint_field(&mut source, 7, style);
    append_varint_field(&mut source, 15, largest);
    append_varint_field(&mut source, 16, smallest);
    append_varint_field(&mut source, 40, automatic);
    source
}

fn duration_source_without(omitted: u32) -> Vec<u8> {
    let mut source = Vec::new();
    for (field, value) in [
        (1, u64::from(codec::NATIVE_DURATION_FORMAT_TYPE)),
        (7, u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED)),
        (15, u64::from(codec::NATIVE_DURATION_UNIT_HOURS)),
        (16, u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS)),
        (40, 1),
    ] {
        if field != omitted {
            append_varint_field(&mut source, field, value);
        }
    }
    source
}

fn selected_value(field: u32) -> u64 {
    match field {
        1 => u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
        7 => u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED),
        15 => u64::from(codec::NATIVE_DURATION_UNIT_HOURS),
        16 => u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS),
        40 => 1,
        _ => unreachable!("selected Duration field is bounded"),
    }
}

fn raw_duration_source(noncanonical_key: Option<u32>, noncanonical_value: Option<u32>) -> Vec<u8> {
    let mut source = Vec::new();
    for (field, value) in [
        (1, u64::from(codec::NATIVE_DURATION_FORMAT_TYPE)),
        (7, u64::from(codec::NATIVE_DURATION_STYLE_ABBREVIATED)),
        (15, u64::from(codec::NATIVE_DURATION_UNIT_HOURS)),
        (16, u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS)),
        (40, 1),
    ] {
        let key = if noncanonical_key == Some(field) {
            overlong_varint(u64::from(field) << 3)
        } else {
            encode_varint(u64::from(field) << 3)
        };
        let encoded_value = if noncanonical_value == Some(field) {
            overlong_varint(value)
        } else {
            encode_varint(value)
        };
        source.extend_from_slice(&key);
        source.extend_from_slice(&encoded_value);
    }
    source
}

fn unknown_wire_shapes_source() -> Vec<u8> {
    let mut source = Vec::new();
    append_varint_field(
        &mut source,
        1,
        u64::from(codec::NATIVE_DURATION_FORMAT_TYPE),
    );
    append_raw_varint_field(&mut source, 100, &[0x81, 0x00]);
    append_varint_field(
        &mut source,
        7,
        u64::from(codec::NATIVE_DURATION_STYLE_COLON),
    );
    append_wire_field(
        &mut source,
        101,
        1,
        &[0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08],
    );
    append_varint_field(
        &mut source,
        15,
        u64::from(codec::NATIVE_DURATION_UNIT_WEEKS),
    );
    append_wire_field(&mut source, 102, 2, &[0xde, 0xad, 0xbe, 0xef]);
    append_varint_field(
        &mut source,
        16,
        u64::from(codec::NATIVE_DURATION_UNIT_MILLISECONDS),
    );
    let mut group_body = Vec::new();
    append_raw_varint_field(&mut group_body, 104, &[0x81, 0x00]);
    append_wire_field(&mut group_body, 105, 5, &[0xca, 0xfe, 0xba, 0xbe]);
    append_group(&mut source, 103, &group_body);
    append_varint_field(&mut source, 40, 0);
    source
}

fn unknown_max_field_source() -> Vec<u8> {
    let mut source = duration_source(
        codec::DurationStyle::Abbreviated,
        codec::DurationUnit::Hours,
        codec::DurationUnit::Milliseconds,
        true,
    );
    append_raw_varint_field(&mut source, 0x1fff_ffff, &[0x81, 0x00]);
    source
}

fn deep_group_source(depth: usize) -> Vec<u8> {
    let mut source = duration_source(
        codec::DurationStyle::Abbreviated,
        codec::DurationUnit::Hours,
        codec::DurationUnit::Milliseconds,
        true,
    );
    for index in 0..depth {
        append_raw_key(
            &mut source,
            100 + u32::try_from(index).unwrap_or(u32::MAX),
            3,
        );
    }
    for index in (0..depth).rev() {
        append_raw_key(
            &mut source,
            100 + u32::try_from(index).unwrap_or(u32::MAX),
            4,
        );
    }
    source
}

fn append_group(output: &mut Vec<u8>, number: u32, body: &[u8]) {
    append_raw_key(output, number, 3);
    output.extend_from_slice(body);
    append_raw_key(output, number, 4);
}

fn append_wire_field(output: &mut Vec<u8>, number: u32, wire: u8, payload: &[u8]) {
    append_raw_key(output, number, wire);
    match wire {
        0 | 1 | 5 => output.extend_from_slice(payload),
        2 => {
            push_varint(output, payload.len() as u64);
            output.extend_from_slice(payload);
        },
        _ => output.extend_from_slice(payload),
    }
}

fn append_raw_varint_field(output: &mut Vec<u8>, number: u32, encoded_value: &[u8]) {
    append_raw_key(output, number, 0);
    output.extend_from_slice(encoded_value);
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_raw_key(output, number, 0);
    push_varint(output, value);
}

fn append_raw_key(output: &mut Vec<u8>, number: u32, wire: u8) {
    push_varint(output, (u64::from(number) << 3) | u64::from(wire));
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn encode_varint(value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, value);
    output
}

fn overlong_varint(value: u64) -> Vec<u8> {
    let mut output = encode_varint(value);
    let final_byte = output.pop().expect("canonical varint has a final byte");
    output.push(final_byte | 0x80);
    output.push(0);
    output
}
