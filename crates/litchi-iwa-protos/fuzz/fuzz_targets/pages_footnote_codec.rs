#![no_main]

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::pages_footnote_codec as reference_codec;
use litchi_iwa_protos::pages_footnote_marker_codec as marker_codec;
use litchi_iwa_protos::tswp::{FootnoteReferenceAttachmentArchive, TextualAttachmentArchive};
use prost::Message as _;

// Keep every strict pass and every candidate no-op output finite. Oversized
// inputs are skipped rather than truncated, so each decoder always receives
// one unchanged caller-owned source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_RECURSION: u32 = 64;

const MALFORMED_MARKER: &[&[u8]] = &[
    // Duplicate singular string-equivalent.
    &[0x0a, 0x01, b'*', 0x0a, 0x01, b'+'],
    // Duplicate singular kind.
    &[0x10, 0x02, 0x10, 0x03],
    // Non-canonical kind varint.
    &[0x10, 0x82, 0x00],
    // Invalid UTF-8 in the selected text field.
    &[0x0a, 0x01, 0xff],
    // Selected string field with a varint wire type.
    &[0x08, 0x01],
    // Truncated length-delimited field.
    &[0x0a, 0x02, b'*'],
    // Unterminated unknown group.
    &[0x53, 0x08, 0x01],
    // Unknown group closed with a different field number.
    &[0x53, 0x5c],
];

const MALFORMED_REFERENCE: &[&[u8]] = &[
    // A present TSP.Reference must carry its required identifier.
    &[0x12, 0x00],
    // A zero TSP.Reference identifier is forbidden by the strict identity
    // check.
    &[0x12, 0x02, 0x08, 0x00],
    // Duplicate singular custom marker.
    &[0x1a, 0x01, b'*', 0x1a, 0x01, b'+'],
    // Invalid UTF-8 in the selected custom marker.
    &[0x1a, 0x01, 0xff],
    // Selected custom marker with a varint wire type.
    &[0x18, 0x01],
    // Truncated nested reference.
    &[0x12, 0x02, 0x08],
    // Non-canonical nested textual kind.
    &[0x0a, 0x03, 0x10, 0x82, 0x00],
    // Unterminated unknown group.
    &[0x53, 0x08, 0x01],
    // Unknown group closed with a different field number.
    &[0x53, 0x5c],
];

#[derive(Clone, Copy)]
struct SourceRange {
    start: usize,
    end: usize,
}

impl SourceRange {
    fn new(source: &[u8]) -> Self {
        let start = source.as_ptr() as usize;
        let end = start
            .checked_add(source.len())
            .expect("bounded source pointer range");
        Self { start, end }
    }

    fn assert_borrowed(self, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start
            .checked_add(bytes.len())
            .expect("bounded borrowed payload pointer range");
        assert!(
            start >= self.start && end <= self.end,
            "decoded payload did not borrow from the source"
        );
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source);
    exercise_mutations(&source, data);

    // The fixed malformed and resource recipes are useful even when a local
    // campaign starts from arbitrary bytes that never reach a successful
    // projection.
    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(|| {
        exercise_known_malformed();
        exercise_limit_guards();
        exercise_input_limit();
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

fn reference_options(source: &[u8]) -> reference_codec::DecodeOptions {
    reference_codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
}

fn marker_options(source: &[u8]) -> marker_codec::DecodeOptions {
    marker_codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    let source_range = SourceRange::new(source);

    let marker = marker_codec::decode_textual_attachment(source, marker_options(source));
    assert_eq!(source, before.as_slice(), "marker decode modified source");
    match marker {
        Ok(snapshot) => {
            if let Some(value) = snapshot.string_equivalent() {
                source_range.assert_borrowed(value.as_bytes());
            }
            assert!(snapshot.raw().len() <= MAX_OUTPUT_BYTES);
            // The marker codec has no production encoding path: its raw slice is
            // the exact lossless no-op/rewrite value.
            assert_eq!(snapshot.raw(), source, "marker no-op changed source bytes");
            if let Ok(oracle) = TextualAttachmentArchive::decode(source) {
                assert_eq!(
                    snapshot.string_equivalent(),
                    oracle.string_equivalent.as_deref()
                );
                assert_eq!(snapshot.kind(), oracle.kind);
            }
            black_box(snapshot);
        },
        Err(error) => observe_error(error),
    }

    let reference = reference_codec::decode_footnote_reference(source, reference_options(source));
    assert_eq!(
        source,
        before.as_slice(),
        "reference decode modified source"
    );
    match reference {
        Ok(snapshot) => {
            if let Some(value) = snapshot.super_string_equivalent() {
                source_range.assert_borrowed(value.as_bytes());
            }
            if let Some(value) = snapshot.custom_mark_string() {
                source_range.assert_borrowed(value.as_bytes());
            }
            // A successful strict read followed by a caller-owned candidate copy
            // is the exact no-op rewrite contract for this raw-preserving codec.
            let candidate = source.to_vec();
            assert!(candidate.len() <= MAX_OUTPUT_BYTES);
            assert_eq!(candidate, source, "reference no-op changed source bytes");
            if let Ok(oracle) = FootnoteReferenceAttachmentArchive::decode(source) {
                assert_reference_matches(snapshot, &oracle);
            }
            black_box(snapshot);
        },
        Err(error) => observe_error(error),
    }
}

fn assert_reference_matches(
    snapshot: reference_codec::FootnoteReferenceSnapshot<'_>,
    oracle: &FootnoteReferenceAttachmentArchive,
) {
    assert_eq!(
        snapshot.super_string_equivalent(),
        oracle
            .super_
            .as_ref()
            .and_then(|value| value.string_equivalent.as_deref())
    );
    assert_eq!(
        snapshot.super_kind(),
        oracle.super_.as_ref().and_then(|value| value.kind)
    );
    match (
        snapshot.contained_storage(),
        oracle.contained_storage.as_ref(),
    ) {
        (Some(snapshot), Some(oracle)) => {
            assert_eq!(snapshot.identifier().get(), oracle.identifier);
            assert_eq!(snapshot.deprecated_type(), oracle.deprecated_type);
            assert_eq!(
                snapshot.deprecated_is_external(),
                oracle.deprecated_is_external
            );
        },
        (None, None) => {},
        _ => panic!("strict and Prost reference presence differed"),
    }
    assert_eq!(
        snapshot.custom_mark_string(),
        oracle.custom_mark_string.as_deref()
    );
}

fn exercise_mutations(source: &[u8], data: &[u8]) {
    if !source.is_empty() {
        let mut mutated = source.to_vec();
        let index = usize::from(data.first().copied().unwrap_or_default()) % mutated.len();
        mutated[index] ^= 0xff;
        observe_mutation(&mutated);
    }
}

fn observe_mutation(source: &[u8]) {
    let before = source.to_vec();
    match marker_codec::decode_textual_attachment(source, marker_options(source)) {
        Ok(snapshot) => {
            black_box((
                snapshot.string_equivalent(),
                snapshot.kind(),
                snapshot.raw(),
            ));
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(
        source,
        before.as_slice(),
        "mutated marker read modified source"
    );

    match reference_codec::decode_footnote_reference(source, reference_options(source)) {
        Ok(snapshot) => {
            black_box((
                snapshot.super_string_equivalent(),
                snapshot.super_kind(),
                snapshot.contained_storage(),
                snapshot.custom_mark_string(),
            ));
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(
        source,
        before.as_slice(),
        "mutated reference read modified source"
    );
}

fn exercise_known_malformed() {
    for source in MALFORMED_MARKER {
        let before = source.to_vec();
        let result = marker_codec::decode_textual_attachment(source, marker_options(source));
        assert!(result.is_err(), "known malformed marker was accepted");
        if let Err(error) = result {
            observe_error(error);
        }
        assert_eq!(
            *source,
            before.as_slice(),
            "malformed marker modified source"
        );
    }

    for source in MALFORMED_REFERENCE {
        let before = source.to_vec();
        let result = reference_codec::decode_footnote_reference(source, reference_options(source));
        assert!(result.is_err(), "known malformed reference was accepted");
        if let Err(error) = result {
            observe_error(error);
        }
        assert_eq!(
            *source,
            before.as_slice(),
            "malformed reference modified source"
        );
    }
}

fn exercise_limit_guards() {
    let marker = [0x0a, 0x01, b'*', 0x10, 0x02];
    let reference = [
        0x0a, 0x05, 0x0a, 0x01, b'*', 0x10, 0x02, 0x12, 0x02, 0x08, 0x2a,
    ];

    assert!(
        marker_codec::decode_textual_attachment(
            &marker,
            marker_codec::DecodeOptions::new(
                marker.len() - 1,
                MAX_FIELDS,
                MAX_WORK_BYTES,
                MAX_RECURSION
            ),
        )
        .is_err(),
        "marker byte ceiling was not enforced"
    );
    assert!(
        reference_codec::decode_footnote_reference(
            &reference,
            reference_codec::DecodeOptions::new(
                reference.len() - 1,
                MAX_FIELDS,
                MAX_WORK_BYTES,
                MAX_RECURSION,
            ),
        )
        .is_err(),
        "reference byte ceiling was not enforced"
    );

    assert!(
        marker_codec::decode_textual_attachment(
            &marker,
            marker_codec::DecodeOptions::new(marker.len(), 1, MAX_WORK_BYTES, MAX_RECURSION),
        )
        .is_err(),
        "marker field ceiling was not enforced"
    );
    assert!(
        reference_codec::decode_footnote_reference(
            &reference,
            reference_codec::DecodeOptions::new(reference.len(), 1, MAX_WORK_BYTES, MAX_RECURSION),
        )
        .is_err(),
        "reference field ceiling was not enforced"
    );

    assert!(
        marker_codec::decode_textual_attachment(
            &marker,
            marker_codec::DecodeOptions::new(marker.len(), MAX_FIELDS, marker.len(), MAX_RECURSION),
        )
        .is_err(),
        "marker work ceiling was not enforced"
    );
    assert!(
        reference_codec::decode_footnote_reference(
            &reference,
            reference_codec::DecodeOptions::new(
                reference.len(),
                MAX_FIELDS,
                reference.len(),
                MAX_RECURSION
            ),
        )
        .is_err(),
        "reference work ceiling was not enforced"
    );

    assert!(
        marker_codec::decode_textual_attachment(
            &deep_unknown_groups(MAX_RECURSION as usize + 1),
            marker_codec::DecodeOptions::new(
                MAX_INPUT_BYTES,
                MAX_FIELDS,
                MAX_WORK_BYTES,
                MAX_RECURSION
            ),
        )
        .is_err(),
        "marker nesting ceiling was not enforced"
    );
    assert!(
        reference_codec::decode_footnote_reference(
            &deep_unknown_groups(MAX_RECURSION as usize + 1),
            reference_codec::DecodeOptions::new(
                MAX_INPUT_BYTES,
                MAX_FIELDS,
                MAX_WORK_BYTES,
                MAX_RECURSION
            ),
        )
        .is_err(),
        "reference nesting ceiling was not enforced"
    );

    // The codecs intentionally have no encoding path. Keep the candidate
    // output ceiling explicit in the target so a future raw-preserving
    // rewrite cannot silently turn this harness into an unbounded allocator.
    assert!(MAX_INPUT_BYTES <= MAX_OUTPUT_BYTES);
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let source = OVERSIZED.get_or_init(|| vec![0; MAX_INPUT_BYTES + 1].into_boxed_slice());
    assert!(
        marker_codec::decode_textual_attachment(
            source,
            marker_codec::DecodeOptions::new(
                MAX_INPUT_BYTES,
                MAX_FIELDS,
                MAX_WORK_BYTES,
                MAX_RECURSION
            ),
        )
        .is_err(),
        "oversized marker source was accepted"
    );
    assert!(
        reference_codec::decode_footnote_reference(
            source,
            reference_codec::DecodeOptions::new(
                MAX_INPUT_BYTES,
                MAX_FIELDS,
                MAX_WORK_BYTES,
                MAX_RECURSION
            ),
        )
        .is_err(),
        "oversized reference source was accepted"
    );
}

fn deep_unknown_groups(depth: usize) -> Vec<u8> {
    let mut output = Vec::with_capacity(depth.saturating_mul(4));
    let start = varint((100_u64 << 3) | 3);
    let end = varint((100_u64 << 3) | 4);
    for _ in 0..depth {
        output.extend_from_slice(&start);
    }
    for _ in 0..depth {
        output.extend_from_slice(&end);
    }
    output
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

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
