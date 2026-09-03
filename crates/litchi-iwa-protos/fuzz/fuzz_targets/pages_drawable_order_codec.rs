#![no_main]

//! Bounded fuzzing for the strict, borrowed Pages drawable-order codec.
//!
//! The harness keeps arbitrary malformed wire reachable while deterministic
//! recipes exercise complete valid orders, source-preserving record moves,
//! optional reference fields, unknown scalar/group spans, duplicate and zero
//! identifiers, and finite resource ceilings.  Successful snapshots are
//! required to borrow every nested reference from the caller-owned source;
//! successful rewrites are decoded again before the source is checked for
//! mutation.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::pages_drawable_order_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_REFERENCES: usize = 4 * 1024;
const MAX_RECURSION: u32 = 64;

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);

    // Arbitrary bytes seldom contain a complete reference permutation.  Keep
    // the valid and malformed semantic branches hot in every worker without
    // placing generated native package bytes in the checked-in corpus.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(exercise_fixtures);
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
    )
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let decode = codec::decode_drawable_order_with_report(source, options(source));
    assert_eq!(source, before.as_slice(), "decode modified its source");

    let (snapshot, report) = match decode {
        Ok(value) => value,
        Err(error) => {
            black_box(error);
            // A malformed source must reject a rewrite before mutating the
            // caller-owned bytes, even when the requested order is empty.
            let empty: [u64; 0] = [];
            let _ = black_box(codec::rewrite_drawable_order(
                source,
                codec::DrawableOrderWrite::new(&empty),
                options(source),
            ));
            exercise_mutations(source);
            assert_eq!(source, before.as_slice(), "malformed probe modified source");
            return;
        },
    };
    assert_eq!(snapshot.raw(), source, "snapshot changed its source");
    assert_eq!(snapshot.len(), report.references());

    let references: Vec<_> = snapshot.references().collect();
    assert_eq!(references.len(), snapshot.len());
    for reference in &references {
        assert_borrowed(source, reference.raw());
        black_box((
            reference.identifier(),
            reference.deprecated_type(),
            reference.deprecated_is_external(),
        ));
    }
    let identifiers: Vec<_> = snapshot.identifiers().collect();
    assert_eq!(identifiers.len(), snapshot.len());
    assert_eq!(
        identifiers,
        references
            .iter()
            .map(|item| item.identifier())
            .collect::<Vec<_>>()
    );

    // A valid decode is a source-preserving no-op candidate.  Resource
    // ceilings are intentionally observed rather than asserted here because
    // a large arbitrary valid source can exhaust the rewrite work profile.
    let no_op = codec::rewrite_drawable_order_with_report(
        source,
        codec::DrawableOrderWrite::new(&identifiers),
        options(source),
    );
    if let Ok((candidate, rewrite_report)) = no_op {
        assert_eq!(candidate, source, "no-op rewrite changed bytes");
        assert!(!rewrite_report.changed());
        assert_eq!(rewrite_report.output_bytes(), candidate.len());
        black_box(rewrite_report);
    }

    if identifiers.len() > 1 {
        let mut requested = identifiers.clone();
        if data.first().copied().unwrap_or_default() & 1 == 0 {
            requested.reverse();
        } else {
            requested.rotate_left(1);
        }
        if requested != identifiers {
            let rewrite = codec::rewrite_drawable_order_with_report(
                source,
                codec::DrawableOrderWrite::new(&requested),
                options(source),
            );
            if let Ok((candidate, rewrite_report)) = rewrite {
                assert_eq!(source, before.as_slice(), "rewrite modified its source");
                let (decoded, _) =
                    codec::decode_drawable_order_with_report(&candidate, options(&candidate))
                        .expect("a successful rewrite must decode");
                assert_eq!(decoded.identifiers().collect::<Vec<_>>(), requested);
                assert_eq!(rewrite_report.output_bytes(), candidate.len());
                black_box(rewrite_report);
            }
        }
    }

    exercise_limit_profiles(source, report);
    exercise_mutations(source);
    assert_eq!(
        source,
        before.as_slice(),
        "codec probes modified their source"
    );
}

fn assert_borrowed(source: &[u8], nested: &[u8]) {
    if nested.is_empty() {
        return;
    }
    let source_start = source.as_ptr() as usize;
    let source_end = source_start.saturating_add(source.len());
    let nested_start = nested.as_ptr() as usize;
    let nested_end = nested_start.saturating_add(nested.len());
    assert!(
        nested_start >= source_start && nested_end <= source_end,
        "borrowed reference escaped its source"
    );
}

fn exercise_limit_profiles(source: &[u8], report: codec::DecodeReport) {
    let decode_profiles = [
        codec::DecodeOptions::new(
            source.len().saturating_sub(1),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            report.fields().saturating_sub(1),
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            report.work_bytes().saturating_sub(1),
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            report.max_depth().saturating_sub(1),
            MAX_REFERENCES,
        ),
        codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            report.references().saturating_sub(1),
        ),
    ];
    for profile in decode_profiles {
        let _ = black_box(codec::decode_drawable_order(source, profile));
    }

    if !source.is_empty() {
        let mut too_small = source.to_vec();
        too_small.pop();
        let _ = black_box(codec::decode_drawable_order(
            &too_small,
            options(&too_small),
        ));
    }

    // Invalid configured profiles must fail closed before any source-owned
    // snapshot can escape, including the Buffa-recursion boundary.
    for profile in [
        codec::DecodeOptions::new(
            0,
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        codec::DecodeOptions::new(MAX_INPUT_BYTES, 0, 0, 0, 0, 0),
        codec::DecodeOptions::new(
            MAX_INPUT_BYTES,
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION + 1,
            MAX_REFERENCES,
        ),
    ] {
        let _ = black_box(codec::decode_drawable_order(source, profile));
    }
}

fn exercise_mutations(source: &[u8]) {
    if source.is_empty() {
        return;
    }
    let mut mutated = source.to_vec();
    mutated[0] ^= 1;
    let _ = black_box(codec::decode_drawable_order(&mutated, options(&mutated)));
    mutated.push(0);
    let _ = black_box(codec::decode_drawable_order(&mutated, options(&mutated)));
}

fn exercise_fixtures() {
    let first = reference(
        11,
        &[
            0x10, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x01,
        ],
    );
    let second = reference(22, &[0x18, 0x01]);
    let third = reference(33, &[0x98, 0x03, 0x81, 0x00]);
    let valid = order(
        &[first, second, third],
        &[0xa3, 0x03, 0x08, 0x01, 0xa4, 0x03, 0x98, 0x03, 0x81, 0x00],
    );
    exercise_source(&valid, &[0]);

    let optional = order(
        &[reference(
            7,
            &[
                0x10, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x01, 0x18, 0x01,
            ],
        )],
        &[],
    );
    exercise_source(&optional, &[1]);

    let malformed = [
        vec![0x08, 0x01],             // root field has wrong wire type
        vec![0x0a, 0x00],             // missing required identifier
        vec![0x0a, 0x02, 0x08, 0x00], // zero identifier
        order(&[reference(1, &[]), reference(1, &[])], &[]), // duplicate identifier
        vec![0x0a, 0x03, 0x08, 0x81, 0x00], // non-canonical identifier
        vec![0x0a, 0x02, 0x10, 0x01], // required field has wrong wire
        vec![0xdb, 0x03, 0x08, 0x01, 0xdc, 0x03], // unknown balanced group
        vec![0xdb, 0x03, 0x08, 0x01], // unterminated unknown group
    ];
    for source in malformed {
        let _ = black_box(codec::decode_drawable_order(&source, options(&source)));
    }

    let mut deep = Vec::new();
    for _ in 0..(MAX_RECURSION as usize + 2) {
        deep.extend_from_slice(&[0xdb, 0x03]);
    }
    deep.extend_from_slice(&[0x08, 0x01]);
    for _ in 0..(MAX_RECURSION as usize + 2) {
        deep.extend_from_slice(&[0xdc, 0x03]);
    }
    let _ = black_box(codec::decode_drawable_order(&deep, options(&deep)));
}

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return output;
        }
    }
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    [varint(u64::from(number) << 3), varint(value)].concat()
}

fn field_bytes(number: u32, value: &[u8]) -> Vec<u8> {
    [
        varint((u64::from(number) << 3) | 2),
        varint(value.len() as u64),
        value.to_vec(),
    ]
    .concat()
}

fn reference(identifier: u64, tail: &[u8]) -> Vec<u8> {
    [field_varint(1, identifier), tail.to_vec()].concat()
}

fn order(references: &[Vec<u8>], tail: &[u8]) -> Vec<u8> {
    let mut source = Vec::new();
    for reference in references {
        source.extend_from_slice(&field_bytes(1, reference));
    }
    source.extend_from_slice(tail);
    source
}
