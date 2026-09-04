#![no_main]

//! Strict, source-preserving fuzzing for the native Numbers Text
//! `FormatStructArchive` projection.
//!
//! Text is intentionally a small format: the selected payload contains only
//! the native discriminator (`260`).  The target still treats the complete
//! source as authoritative, so unknown records/groups are retained and every
//! successful decode/rewrite remains bounded, borrowed, and exact.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_text_format_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

const CANONICAL_TEXT: &[u8] = &[0x08, 0x84, 0x02];
const UNKNOWN_SCALAR: &[u8] = &[0xa0, 0x06, 0x81, 0x00];
const UNKNOWN_GROUP: &[u8] = &[0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06];
const UNKNOWN_LENGTH: &[u8] = &[0xb2, 0x06, 0x03, 0x01, 0x7f, 0x00];
const UNKNOWN_HIGH_SCALAR: &[u8] = &[0xf8, 0xff, 0xff, 0xff, 0x0f, 0x01];

const FIXED_CASES: &[&[u8]] = &[
    // Canonical native Text (field 1 = 260).
    CANONICAL_TEXT,
    // Unknown scalar, fixed-width, and length-delimited records remain
    // opaque while the selected discriminator stays strict.
    &[
        0xa0, 0x06, 0x81, 0x00, 0x08, 0x84, 0x02, 0xe9, 0x06, 0, 1, 2, 3, 4, 5, 6, 7, 0xed, 0x06,
        8, 9, 10, 11,
    ],
    // A balanced unknown group is retained across a source-preserving write.
    &[
        0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06, 0x08, 0x84, 0x02, 0xb2, 0x06, 0x03, 1, 2, 3,
    ],
    // Unknown records may be interleaved on either side of the known field.
    &[
        0xb2, 0x06, 0x03, b't', b'x', b't', 0x08, 0x84, 0x02, 0xa0, 0x06, 0x81, 0x00,
    ],
    // The maximum legal protobuf field number remains opaque.
    &[0xf8, 0xff, 0xff, 0xff, 0x0f, 0x01, 0x08, 0x84, 0x02],
    // Unknown noncanonical keys and lengths are rejected; only scalar values
    // have the compatibility exception exercised by UNKNOWN_SCALAR above.
    &[0xa0, 0x86, 0x00, 0x01, 0x08, 0x84, 0x02],
    &[0xb2, 0x06, 0x80, 0x00, 0x08, 0x84, 0x02],
    &[
        0xa3, 0x06, 0xb2, 0x06, 0x80, 0x00, 0xa4, 0x06, 0x08, 0x84, 0x02,
    ],
    // Reserved wire types are malformed even for unknown field numbers.
    &[0xd6, 0x02],
    &[0xd7, 0x02],
    // Missing, duplicate, and wrong-wire/type known fields.
    &[],
    &[0x08, 0x84, 0x02, 0x08, 0x84, 0x02],
    &[0x0a, 0x01, 0x00],
    &[0x08, 0x83, 0x02],
    // Non-canonical key/value encodings are rejected before projection.
    &[0x88, 0x00, 0x84, 0x02],
    &[0x08, 0x84, 0x82, 0x00],
    // Other known FormatStructArchive fields cannot be hidden as unknown.
    &[0x08, 0x84, 0x02, 0x10, 0x00],
    &[0x08, 0x84, 0x02, 0x20, 0x00],
    &[0x08, 0x84, 0x02, 0xa0, 0x01, 0x00],
    &[0x08, 0x84, 0x02, 0xe8, 0x02, 0x00],
    // Truncated, mismatched, and unterminated groups remain malformed.
    &[0x08, 0x84],
    &[0x08, 0x84, 0x02, 0xa3, 0x06, 0xa8, 0x06, 0x01],
    &[0x08, 0x84, 0x02, 0xa3, 0x06, 0xa8, 0x06, 0x01, 0xac, 0x06],
    &[0x08, 0x84, 0x02, 0xa4, 0x06],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    // Keep canonical, unknown-field, malformed, and limit paths reachable
    // even when a campaign starts with only arbitrary invalid bytes.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-text-format");
        }
        exercise_unknown_wire_matrix();
        exercise_known_field_rejections();
        exercise_malformed_group_matrix();
        exercise_canonical_writes();
        exercise_encoding_markers();
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
    let scalar = codec::decode_text_format(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "scalar decode modified its source"
    );
    let reported = codec::decode_text_format_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "reported decode modified its source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert_eq!(snapshot.format_type(), codec::NATIVE_TEXT_FORMAT_TYPE);
            assert_eq!(snapshot.raw(), source);
            assert_borrowed(source, snapshot.raw());
            assert_report(report, source.len());
            exercise_limit_profiles(source, report);
            black_box((snapshot, report));
            exercise_rewrite(source, snapshot, data, &before);
        },
        (Err(scalar_error), Err(reported_error)) => {
            assert_eq!(
                scalar_error.resource_limit(),
                reported_error.resource_limit(),
                "Text scalar/report error classifications disagree"
            );
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "Text-format scalar/report disagreement: scalar={:?}, report={:?}",
            scalar_result.map(|snapshot| snapshot.format_type()),
            reported_result.map(|(snapshot, _)| snapshot.format_type())
        ),
    }

    assert_eq!(
        source,
        before.as_slice(),
        "Text-format probes modified their source"
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
        "Text-format snapshot did not borrow from its source"
    );
}

fn assert_report(report: codec::DecodeReport, source_len: usize) {
    assert_eq!(report.input_bytes(), source_len);
    assert_eq!(report.output_bytes(), 0);
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

fn exercise_rewrite(
    source: &[u8],
    snapshot: codec::TextFormatSnapshot<'_>,
    data: &[u8],
    before: &[u8],
) {
    let write = if data.first().copied().unwrap_or_default() & 1 == 0 {
        codec::TextFormatWrite::from_snapshot(snapshot)
    } else {
        codec::TextFormatWrite::new()
    };
    let prepared = match codec::prepare_text_format_rewrite(source, write, options(source)) {
        Ok(prepared) => prepared,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let requirements = prepared.execution_requirements();
    assert_prepare_report(prepared.prepare_report(), requirements);
    assert_eq!(source, before, "prepare modified its source");

    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact Text-format replay failed: {error}"));
    assert_eq!(source, before, "execute modified its source");
    assert_eq!(
        output.bytes(),
        source,
        "Text preserving rewrite changed bytes"
    );
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    assert_eq!(output.report(), prepared.prepare_report());

    let readback = codec::decode_text_format(output.bytes(), options(output.bytes()))
        .unwrap_or_else(|error| panic!("Text-format candidate readback failed: {error}"));
    assert_eq!(codec::TextFormatWrite::from_snapshot(readback), write);
    assert_known_unknown_spans_preserved(source, output.bytes());

    let one_shot = codec::rewrite_text_format(source, write, options(source))
        .unwrap_or_else(|error| panic!("one-shot Text-format rewrite failed: {error}"));
    assert_eq!(one_shot.bytes(), output.bytes());
    assert_eq!(one_shot.report(), output.report());
    exercise_limit_failures(source, write, prepared, requirements);
    black_box(output);
}

fn assert_prepare_report(
    report: codec::DecodeReport,
    requirements: codec::RewriteExecutionRequirements,
) {
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

fn assert_known_unknown_spans_preserved(source: &[u8], candidate: &[u8]) {
    let source_shape = parse_root_shape(source)
        .expect("successful Text-format source must have valid root wire records");
    let candidate_shape = parse_root_shape(candidate)
        .expect("rewritten Text-format candidate must have valid root wire records");
    assert_eq!(
        source_shape, candidate_shape,
        "Text-format rewrite changed unknown-span order or multiplicity"
    );
    for unknown in [
        UNKNOWN_SCALAR,
        UNKNOWN_GROUP,
        UNKNOWN_LENGTH,
        UNKNOWN_HIGH_SCALAR,
    ] {
        if source
            .windows(unknown.len())
            .any(|window| window == unknown)
        {
            assert!(
                candidate
                    .windows(unknown.len())
                    .any(|window| window == unknown),
                "source unknown Text-format span was not preserved"
            );
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RootField<'source> {
    Selected,
    Unknown(&'source [u8]),
}

fn parse_root_shape(source: &[u8]) -> Option<Vec<RootField<'_>>> {
    let mut fields = Vec::new();
    let mut offset = 0usize;
    while offset < source.len() {
        let start = offset;
        let (key, key_len, canonical) = read_wire_varint(source, offset)?;
        if !canonical {
            return None;
        }
        let number = u32::try_from(key >> 3).ok()?;
        let wire = u8::try_from(key & 7).ok()?;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return None;
        }
        offset = offset.checked_add(key_len)?;
        offset = skip_wire_value(source, offset, wire, number, 0)?;
        let raw = source.get(start..offset)?;
        fields.push(if number == 1 {
            RootField::Selected
        } else {
            RootField::Unknown(raw)
        });
    }
    Some(fields)
}

fn skip_wire_value(
    source: &[u8],
    mut offset: usize,
    wire: u8,
    number: u32,
    depth: u32,
) -> Option<usize> {
    match wire {
        0 => {
            let (_, length, _) = read_wire_varint(source, offset)?;
            offset.checked_add(length)
        },
        1 => offset.checked_add(8).filter(|end| *end <= source.len()),
        2 => {
            let (length, length_bytes, canonical) = read_wire_varint(source, offset)?;
            if !canonical {
                return None;
            }
            offset = offset.checked_add(length_bytes)?;
            let length = usize::try_from(length).ok()?;
            offset
                .checked_add(length)
                .filter(|end| *end <= source.len())
        },
        3 => {
            if depth >= MAX_RECURSION {
                return None;
            }
            let mut cursor = offset;
            while cursor < source.len() {
                let (key, key_len, canonical) = read_wire_varint(source, cursor)?;
                if !canonical {
                    return None;
                }
                let nested_number = u32::try_from(key >> 3).ok()?;
                let nested_wire = u8::try_from(key & 7).ok()?;
                if nested_number == 0 || nested_number > MAX_FIELD_NUMBER {
                    return None;
                }
                cursor = cursor.checked_add(key_len)?;
                if nested_wire == 4 {
                    return (nested_number == number).then_some(cursor);
                }
                cursor = skip_wire_value(source, cursor, nested_wire, nested_number, depth + 1)?;
            }
            None
        },
        4 | 6 | 7 => None,
        5 => offset.checked_add(4).filter(|end| *end <= source.len()),
        _ => None,
    }
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
            let canonical = encoded_wire_varint_len(value) == consumed;
            return Some((value, consumed, canonical));
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

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn wire_key(number: u32, wire: u8) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, (u64::from(number) << 3) | u64::from(wire));
    output
}

fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut output = wire_key(number, 0);
    push_varint(&mut output, value);
    output
}

fn fixed32_field(number: u32, value: [u8; 4]) -> Vec<u8> {
    let mut output = wire_key(number, 5);
    output.extend_from_slice(&value);
    output
}

fn fixed64_field(number: u32, value: [u8; 8]) -> Vec<u8> {
    let mut output = wire_key(number, 1);
    output.extend_from_slice(&value);
    output
}

fn length_field(number: u32, value: &[u8]) -> Vec<u8> {
    let mut output = wire_key(number, 2);
    push_varint(&mut output, value.len() as u64);
    output.extend_from_slice(value);
    output
}

fn group_field(number: u32, body: &[u8]) -> Vec<u8> {
    let mut output = wire_key(number, 3);
    output.extend_from_slice(body);
    output.extend_from_slice(&wire_key(number, 4));
    output
}

fn source_with_extra(extra: &[u8]) -> Vec<u8> {
    let mut source = Vec::with_capacity(CANONICAL_TEXT.len() + extra.len());
    source.extend_from_slice(extra);
    source.extend_from_slice(CANONICAL_TEXT);
    source
}

fn exercise_unknown_wire_matrix() {
    let nested = varint_field(50, 9);
    let records = [
        varint_field(46, 7),
        fixed64_field(47, [0x89, 0xab, 0xcd, 0xef, 0x10, 0x32, 0x54, 0x76]),
        length_field(48, &[0x00, 0xff, 0x80]),
        group_field(49, &nested),
        varint_field(MAX_FIELD_NUMBER, 1),
    ];
    let mut source = Vec::new();
    for record in records {
        source.extend_from_slice(&record);
    }
    source.extend_from_slice(CANONICAL_TEXT);
    exercise_source(&source, b"unknown-wire-matrix");

    // A canonical key plus an overlong value is the one noncanonical unknown
    // spelling intentionally admitted by this projection.
    let mut overlong = source_with_extra(UNKNOWN_SCALAR);
    overlong.extend_from_slice(UNKNOWN_GROUP);
    exercise_source(&overlong, b"unknown-overlong-value");
}

fn exercise_known_field_rejections() {
    for number in 2..=45 {
        for extra in [
            varint_field(number, 0),
            fixed64_field(number, [0; 8]),
            length_field(number, &[0]),
            fixed32_field(number, [0; 4]),
            group_field(number, &[]),
        ] {
            exercise_source(&source_with_extra(&extra), b"known-sibling-field");
        }
    }
}

fn exercise_malformed_group_matrix() {
    let malformed = [
        wire_key(46, 4),
        wire_key(46, 6),
        wire_key(46, 7),
        [wire_key(46, 1), vec![0; 7]].concat(),
        [wire_key(46, 5), vec![0; 3]].concat(),
        [wire_key(46, 2), vec![2, 0]].concat(),
        [wire_key(46, 3), varint_field(47, 1)].concat(),
        [wire_key(46, 3), varint_field(47, 1), wire_key(48, 4)].concat(),
        [
            wire_key(46, 3),
            wire_key(47, 3),
            wire_key(46, 4),
            wire_key(47, 4),
        ]
        .concat(),
    ];
    for extra in malformed {
        exercise_source(&source_with_extra(&extra), b"malformed-wire");
    }
}

fn exercise_limit_failures(
    source: &[u8],
    write: codec::TextFormatWrite,
    prepared: codec::PreparedTextFormatRewrite<'_>,
    requirements: codec::RewriteExecutionRequirements,
) {
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    let probes = [
        (requirements.output_bytes() > 0)
            .then(|| exact.with_output_bytes(requirements.output_bytes() - 1)),
        (requirements.fields() > 0).then(|| exact.with_fields(requirements.fields() - 1)),
        (requirements.work_bytes() > 0)
            .then(|| exact.with_work_bytes(requirements.work_bytes() - 1)),
        (requirements.max_depth() > 0).then(|| exact.with_max_depth(requirements.max_depth() - 1)),
        (requirements.allocations() > 0)
            .then(|| exact.with_allocations(requirements.allocations() - 1)),
        (requirements.retained_bytes() > 0)
            .then(|| exact.with_retained_bytes(requirements.retained_bytes() - 1)),
        (requirements.scratch_bytes() > 0)
            .then(|| exact.with_scratch_bytes(requirements.scratch_bytes() - 1)),
    ];
    for limits in probes.into_iter().flatten() {
        let result = prepared.execute(limits);
        assert!(
            result.is_err(),
            "Text-format rewrite accepted a strict limit"
        );
        if let Err(error) = result {
            black_box(error.resource_limit());
            observe_error(error);
        }
    }

    let replay = codec::rewrite_text_format(source, write, options(source));
    if let Err(error) = replay {
        observe_error(error);
    }
}

fn exercise_limit_profiles(source: &[u8], report: codec::DecodeReport) {
    if source.is_empty() {
        return;
    }
    let exact = codec::DecodeOptions::new(
        source.len(),
        source.len(),
        report.fields(),
        report.work_bytes(),
        report.max_depth().max(1),
        0,
        0,
        0,
    );
    codec::decode_text_format(source, exact)
        .unwrap_or_else(|error| panic!("exact Text-format decode limits failed: {error}"));

    let input_limited = codec::DecodeOptions::new(
        source.len() - 1,
        source.len(),
        report.fields(),
        report.work_bytes(),
        report.max_depth().max(1),
        0,
        0,
        0,
    );
    assert_resource_limit(codec::decode_text_format(source, input_limited), |limit| {
        matches!(limit, codec::DecodeLimit::InputBytes { .. })
    });

    if report.fields() > 0 {
        let limited = codec::DecodeOptions::new(
            source.len(),
            source.len(),
            report.fields() - 1,
            report.work_bytes(),
            report.max_depth().max(1),
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_text_format(source, limited), |limit| {
            matches!(limit, codec::DecodeLimit::Fields { .. })
        });
    }

    if report.work_bytes() > 0 {
        let limited = codec::DecodeOptions::new(
            source.len(),
            source.len(),
            report.fields(),
            report.work_bytes() - 1,
            report.max_depth().max(1),
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_text_format(source, limited), |limit| {
            matches!(limit, codec::DecodeLimit::Work { .. })
        });
    }

    if report.max_depth() > 0 {
        let limited = codec::DecodeOptions::new(
            source.len(),
            source.len(),
            report.fields(),
            report.work_bytes(),
            report.max_depth() - 1,
            0,
            0,
            0,
        );
        assert_resource_limit(codec::decode_text_format(source, limited), |limit| {
            matches!(limit, codec::DecodeLimit::Nesting { .. })
        });
    }
}

fn assert_resource_limit<T>(
    result: Result<T, codec::DecodeError>,
    predicate: impl FnOnce(codec::DecodeLimit) -> bool,
) {
    let error = match result {
        Ok(_) => panic!("resource-limited Text-format operation unexpectedly succeeded"),
        Err(error) => error,
    };
    let limit = error
        .resource_limit()
        .expect("resource failure did not expose a typed limit");
    assert!(predicate(limit), "unexpected Text-format resource limit");
    observe_error(error);
}

fn exercise_canonical_writes() {
    let write = codec::TextFormatWrite::new();
    let prepared = codec::prepare_text_format_write(write, options(&[]))
        .unwrap_or_else(|error| panic!("canonical Text-format prepare failed: {error}"));
    let requirements = prepared.execution_requirements();
    assert_prepare_report(prepared.prepare_report(), requirements);
    let output = prepared
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("canonical Text-format execute failed: {error}"));
    assert_eq!(output.bytes(), CANONICAL_TEXT);
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    let readback = codec::decode_text_format(output.bytes(), options(output.bytes()))
        .unwrap_or_else(|error| panic!("canonical Text-format readback failed: {error}"));
    assert_eq!(codec::TextFormatWrite::from_snapshot(readback), write);
    assert_eq!(
        codec::canonical_text_format(write, options(&[]))
            .unwrap_or_else(|error| panic!("canonical one-shot Text-format failed: {error}"))
            .bytes(),
        output.bytes()
    );
    let exact = codec::RewriteExecutionLimits::exact(requirements);
    for limits in [
        exact.with_output_bytes(requirements.output_bytes() - 1),
        exact.with_fields(requirements.fields() - 1),
        exact.with_work_bytes(requirements.work_bytes() - 1),
        exact.with_allocations(0),
        exact.with_retained_bytes(requirements.retained_bytes() - 1),
    ] {
        assert!(
            prepared.execute(limits).is_err(),
            "canonical Text-format write accepted one-below exact limits"
        );
    }
    black_box(output);
}

fn exercise_encoding_markers() {
    assert_eq!(
        codec::decode_text_format_encoding(
            codec::EXPLICIT_TEXT_FORMAT,
            codec::TEXT_CELL_FORMAT_KIND,
        ),
        Ok(codec::TextFormatEncoding::Plain)
    );
    assert_eq!(
        codec::decode_text_format_encoding(
            codec::EXPLICIT_CONVERTED_TEXT_FORMAT,
            codec::TEXT_CELL_FORMAT_KIND,
        ),
        Ok(codec::TextFormatEncoding::Converted)
    );
    assert!(!codec::TextFormatEncoding::Plain.is_converted());
    assert!(codec::TextFormatEncoding::Converted.is_converted());
    for (marker, kind) in [
        (0, codec::TEXT_CELL_FORMAT_KIND),
        (0x82, codec::TEXT_CELL_FORMAT_KIND),
        (codec::EXPLICIT_TEXT_FORMAT, 4),
    ] {
        assert!(codec::decode_text_format_encoding(marker, kind).is_err());
    }
}

fn exercise_limit_guards() {
    let write = codec::TextFormatWrite::new();
    let output_limited = options(&[]).with_max_output_bytes(0);
    assert_resource_limit(
        codec::prepare_text_format_write(write, output_limited),
        |limit| matches!(limit, codec::DecodeLimit::OutputBytes { .. }),
    );

    let prepared =
        codec::prepare_text_format_rewrite(CANONICAL_TEXT, write, options(CANONICAL_TEXT))
            .unwrap_or_else(|error| panic!("Text-format limit guard prepare failed: {error}"));
    let requirements = prepared.execution_requirements();
    assert_resource_limit(
        prepared.execute(codec::RewriteExecutionLimits::exact(requirements).with_allocations(0)),
        |limit| matches!(limit, codec::DecodeLimit::Allocation { .. }),
    );

    // A group chain beyond the finite scanner depth must fail with a typed
    // nesting limit instead of recursing or allocating without bound.
    let deep = deep_group_source((MAX_RECURSION + 1) as usize);
    let error = codec::decode_text_format(&deep, options(&deep))
        .expect_err("deep Text-format group unexpectedly decoded");
    assert!(matches!(
        error.resource_limit(),
        Some(codec::DecodeLimit::Nesting { .. })
    ));
    observe_error(error);
}

fn deep_group_source(depth: usize) -> Vec<u8> {
    let mut source = CANONICAL_TEXT.to_vec();
    for _ in 0..depth {
        source.extend_from_slice(&wire_key(46, 3));
    }
    for _ in 0..depth {
        source.extend_from_slice(&wire_key(46, 4));
    }
    source
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
