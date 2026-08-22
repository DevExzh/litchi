#![no_main]

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::{pages_body_codec, tp, tsp, tswp};
use prost::Message as _;

// Inputs are skipped rather than truncated.  The two body/footnote entry
// points therefore always inspect one unchanged caller-owned source under the
// same finite profile used by the Pages package.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source);

    // Arbitrary bytes rarely reach the selected fields.  These deterministic
    // recipes keep the required root/singular-reference and body-boundary
    // paths hot in every local campaign without committing native package
    // bytes to the corpus.
    static KNOWN: OnceLock<()> = OnceLock::new();
    KNOWN.get_or_init(exercise_known_recipes);
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

fn options(source: &[u8]) -> pages_body_codec::DecodeOptions {
    pages_body_codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();

    let document = pages_body_codec::decode_document_body(source, options(source));
    assert_eq!(source, before.as_slice(), "document decode modified source");
    if let Ok(snapshot) = document {
        if let Ok(archive) = tp::DocumentArchive::decode(source) {
            assert_document_matches(snapshot, &archive);
        }
    }

    let boundary = pages_body_codec::decode_section_boundary(source, options(source));
    assert_eq!(source, before.as_slice(), "boundary decode modified source");
    if let Ok(snapshot) = boundary {
        if let Ok(archive) = tswp::object_attribute_table::ObjectAttribute::decode(source) {
            assert_boundary_matches(snapshot, &archive);
        }
    }

    exercise_limit_profiles(source);
    exercise_mutations(source);
    assert_eq!(
        source,
        before.as_slice(),
        "body/footnote probes modified source"
    );
}

fn assert_document_matches(
    snapshot: pages_body_codec::DocumentBodySnapshot,
    archive: &tp::DocumentArchive,
) {
    assert_reference_matches(snapshot.body_storage(), archive.body_storage.as_ref());
    assert_reference_matches(snapshot.initial_section(), archive.section.as_ref());
}

fn assert_boundary_matches(
    snapshot: pages_body_codec::SectionBoundarySnapshot,
    archive: &tswp::object_attribute_table::ObjectAttribute,
) {
    assert_eq!(snapshot.character_index(), archive.character_index);
    assert_reference_matches(snapshot.section(), archive.object.as_ref());
}

fn assert_reference_matches(
    snapshot: Option<pages_body_codec::ReferenceSnapshot>,
    archive: Option<&tsp::Reference>,
) {
    match (snapshot, archive) {
        (None, None) => {},
        (Some(snapshot), Some(archive)) => {
            assert_eq!(snapshot.identifier().get(), archive.identifier);
            assert_eq!(snapshot.deprecated_type(), archive.deprecated_type);
            assert_eq!(
                snapshot.deprecated_is_external(),
                archive.deprecated_is_external
            );
        },
        _ => panic!("strict and Prost reference presence differed"),
    }
}

fn exercise_limit_profiles(source: &[u8]) {
    // Invalid configured limits must be rejected before a lazy projection can
    // allocate or retain anything.  Keep these checks independent for the
    // document and boundary shapes because one source can be valid for only
    // one of them.
    let invalid_bytes =
        pages_body_codec::DecodeOptions::new(usize::MAX, MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION);
    observe_document(pages_body_codec::decode_document_body(
        source,
        invalid_bytes,
    ));
    observe_boundary(pages_body_codec::decode_section_boundary(
        source,
        invalid_bytes,
    ));

    let zero_recursion =
        pages_body_codec::DecodeOptions::new(source.len().max(1), MAX_FIELDS, MAX_WORK_BYTES, 0);
    assert!(pages_body_codec::decode_document_body(source, zero_recursion).is_err());
    assert!(pages_body_codec::decode_section_boundary(source, zero_recursion).is_err());

    let over_recursion = pages_body_codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION + 1,
    );
    assert!(pages_body_codec::decode_document_body(source, over_recursion).is_err());
    assert!(pages_body_codec::decode_section_boundary(source, over_recursion).is_err());

    if source.is_empty() {
        return;
    }
    let too_small = pages_body_codec::DecodeOptions::new(
        source.len() - 1,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    );
    observe_document(pages_body_codec::decode_document_body(source, too_small));
    observe_boundary(pages_body_codec::decode_section_boundary(source, too_small));

    // For an actually accepted shape these exact one-under ceilings must be
    // rejected as resource failures.  Invalid arbitrary bytes may reject on
    // wire framing first, so this remains an observation rather than a hard
    // assertion on every input.
    let field_limited =
        pages_body_codec::DecodeOptions::new(source.len(), 0, MAX_WORK_BYTES, MAX_RECURSION);
    observe_document(pages_body_codec::decode_document_body(
        source,
        field_limited,
    ));
    observe_boundary(pages_body_codec::decode_section_boundary(
        source,
        field_limited,
    ));

    let work_limit = source.len().saturating_mul(2).saturating_sub(1);
    let work_limited =
        pages_body_codec::DecodeOptions::new(source.len(), MAX_FIELDS, work_limit, MAX_RECURSION);
    observe_document(pages_body_codec::decode_document_body(source, work_limited));
    observe_boundary(pages_body_codec::decode_section_boundary(
        source,
        work_limited,
    ));
}

fn exercise_mutations(source: &[u8]) {
    if source.is_empty() {
        return;
    }
    let mut mutated = source.to_vec();
    mutated[0] ^= 0x01;
    observe_document(pages_body_codec::decode_document_body(
        &mutated,
        options(&mutated),
    ));
    observe_boundary(pages_body_codec::decode_section_boundary(
        &mutated,
        options(&mutated),
    ));

    mutated.push(0);
    observe_document(pages_body_codec::decode_document_body(
        &mutated,
        options(&mutated),
    ));
    observe_boundary(pages_body_codec::decode_section_boundary(
        &mutated,
        options(&mutated),
    ));
}

fn exercise_known_recipes() {
    let document = canonical_document();
    let boundary = canonical_boundary();
    for source in [document, vec![0x7a, 0x00], vec![0x22, 0x00, 0x7a, 0x00]] {
        exercise_source(&source);
    }
    for source in [boundary, vec![0x08, 0x00], vec![0x08, 0x01, 0x12, 0x00]] {
        exercise_source(&source);
    }

    // Selected malformed cases are exercised independently of arbitrary input
    // and are expected to remain rejected by both shape decoders.
    for source in [
        vec![0x22, 0x02, 0x08, 0x01],             // missing required root super
        vec![0x22, 0x02, 0x08, 0x00, 0x7a, 0x00], // zero body identity
        vec![0x22, 0x02, 0x08, 0x01, 0x22, 0x02, 0x08, 0x02, 0x7a, 0x00],
        vec![0x20, 0x01, 0x7a, 0x00], // root reference wrong wire
        vec![0x08, 0x00, 0x08, 0x01], // duplicate boundary index
        vec![0x08, 0x00, 0x12, 0x02, 0x08, 0x00], // zero boundary identity
        vec![0x08, 0x80, 0x80, 0x80, 0x80, 0x10], // oversized boundary index
        vec![0x80],                   // truncated tag
    ] {
        observe_document(pages_body_codec::decode_document_body(
            &source,
            options(&source),
        ));
        observe_boundary(pages_body_codec::decode_section_boundary(
            &source,
            options(&source),
        ));
    }
}

fn canonical_document() -> Vec<u8> {
    let mut output = Vec::new();
    append_bytes(&mut output, 4, &reference(42, Some(-7), Some(false)));
    append_bytes(&mut output, 5, &reference(99, None, None));
    append_bytes(&mut output, 15, b"opaque-super");
    output
}

fn canonical_boundary() -> Vec<u8> {
    let mut output = Vec::new();
    append_varint(&mut output, 1, 65_535);
    append_bytes(&mut output, 2, &reference(u64::MAX, Some(4), Some(true)));
    output
}

fn reference(identifier: u64, deprecated_type: Option<i32>, external: Option<bool>) -> Vec<u8> {
    let mut output = Vec::new();
    append_varint(&mut output, 1, identifier);
    if let Some(value) = deprecated_type {
        append_varint(&mut output, 2, value as u64);
    }
    if let Some(value) = external {
        append_varint(&mut output, 3, u64::from(value));
    }
    output
}

fn append_bytes(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
    put_varint(output, (u64::from(field) << 3) | 2);
    put_varint(
        output,
        u64::try_from(payload.len()).expect("bounded body fuzz payload"),
    );
    output.extend_from_slice(payload);
}

fn append_varint(output: &mut Vec<u8>, field: u32, value: u64) {
    put_varint(output, u64::from(field) << 3);
    put_varint(output, value);
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn observe_document(
    result: Result<pages_body_codec::DocumentBodySnapshot, pages_body_codec::DecodeError>,
) {
    if let Err(error) = result {
        black_box(error.to_string());
        black_box(format!("{error:?}"));
    }
}

fn observe_boundary(
    result: Result<pages_body_codec::SectionBoundarySnapshot, pages_body_codec::DecodeError>,
) {
    if let Err(error) = result {
        black_box(error.to_string());
        black_box(format!("{error:?}"));
    }
}
