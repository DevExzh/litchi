#![no_main]

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::{pages_header_footer_codec, tp, tsp};
use prost::Message as _;

// Keep every strict pass and candidate rewrite within one finite profile.
// Oversized inputs are skipped rather than truncated so all checks observe one
// unchanged caller-owned source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_REFERENCES: usize = 8 * 1024;

const MALFORMED: &[&[u8]] = &[
    // Missing required TSP.Reference.identifier.
    &[0x0a, 0x00],
    // A root field with a varint wire type instead of a nested reference.
    &[0x08, 0x01],
    // Duplicate required identifier in one nested reference.
    &[0x0a, 0x04, 0x08, 0x01, 0x08, 0x02],
    // Non-canonical identifier varint.
    &[0x0a, 0x03, 0x08, 0x81, 0x00],
    // Truncated nested reference.
    &[0x0a, 0x02, 0x08],
    // Unterminated unknown group.
    &[0x53, 0x08, 0x01],
    // Unknown group closed with a different field number.
    &[0x53, 0x5c],
    // Truncated root field length.
    &[0x0a, 0x03, 0x08, 0x01],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);

    // Fixed malformed and resource recipes keep strict failure paths hot
    // even before arbitrary input discovers a valid reference envelope.
    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(|| {
        for malformed in MALFORMED {
            exercise_source(malformed, malformed);
        }
        exercise_known_valid();
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

fn options(source: &[u8]) -> pages_header_footer_codec::DecodeOptions {
    pages_header_footer_codec::DecodeOptions::new(
        source.len().max(1),
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
    )
}

fn exercise_source(source: &[u8], command: &[u8]) {
    let before = source.to_vec();
    let decoded =
        pages_header_footer_codec::decode_section_template_with_report(source, options(source));
    assert_eq!(source, before.as_slice(), "decode modified its source");

    let Ok((snapshot, report)) = decoded else {
        if let Err(error) = decoded {
            observe_error(error);
        }
        return;
    };
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(
        report.references(),
        snapshot.header_count() + snapshot.footer_count()
    );
    black_box((snapshot, report));

    let headers = snapshot
        .headers()
        .map(|reference| pages_header_footer_codec::ReferenceWrite::new(reference.raw()))
        .collect::<Vec<_>>();
    let footers = snapshot
        .footers()
        .map(|reference| pages_header_footer_codec::ReferenceWrite::new(reference.raw()))
        .collect::<Vec<_>>();
    let write = pages_header_footer_codec::SectionTemplateWrite::new(&headers, &footers);
    let rewrite = pages_header_footer_codec::rewrite_section_template_with_report(
        source,
        write,
        options(source),
    );
    assert_eq!(source, before.as_slice(), "rewrite modified its source");
    let Ok((candidate, rewrite_report)) = rewrite else {
        if let Err(error) = rewrite {
            observe_error(error);
        }
        return;
    };
    assert_eq!(rewrite_report.input_bytes(), source.len());
    assert_eq!(rewrite_report.output_bytes(), candidate.len());
    assert_eq!(candidate, source, "borrowed no-op rewrite changed bytes");
    let readback =
        pages_header_footer_codec::decode_section_template(&candidate, options(&candidate))
            .unwrap_or_else(|error| panic!("header/footer no-op readback failed: {error}"));
    assert_eq!(readback.header_count(), snapshot.header_count());
    assert_eq!(readback.footer_count(), snapshot.footer_count());

    exercise_replacement(source, snapshot, command);
}

fn exercise_replacement(
    source: &[u8],
    snapshot: pages_header_footer_codec::SectionTemplateSnapshot<'_>,
    command: &[u8],
) {
    let mut header_bytes = snapshot
        .headers()
        .map(|reference| reference_bytes(reference.identifier(), command))
        .collect::<Vec<_>>();
    let mut footer_bytes = snapshot
        .footers()
        .map(|reference| reference_bytes(reference.identifier(), command))
        .collect::<Vec<_>>();
    if header_bytes.is_empty() && footer_bytes.is_empty() {
        return;
    }

    let headers = header_bytes
        .iter()
        .map(|raw| pages_header_footer_codec::ReferenceWrite::new(raw))
        .collect::<Vec<_>>();
    let footers = footer_bytes
        .iter()
        .map(|raw| pages_header_footer_codec::ReferenceWrite::new(raw))
        .collect::<Vec<_>>();
    let write = pages_header_footer_codec::SectionTemplateWrite::new(&headers, &footers);
    let before = source.to_vec();
    let result = pages_header_footer_codec::rewrite_section_template_with_report(
        source,
        write,
        options(source),
    );
    assert_eq!(source, before.as_slice(), "replacement modified its source");
    let Ok((candidate, report)) = result else {
        if let Err(error) = result {
            observe_error(error);
        }
        return;
    };
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), candidate.len());
    let readback =
        pages_header_footer_codec::decode_section_template(&candidate, options(&candidate))
            .unwrap_or_else(|error| panic!("replacement readback failed: {error}"));
    assert_eq!(readback.header_count(), snapshot.header_count());
    assert_eq!(readback.footer_count(), snapshot.footer_count());
    black_box((&mut header_bytes, &mut footer_bytes, report));
}

fn reference_bytes(original: u64, command: &[u8]) -> Vec<u8> {
    let low = u64::from(command.first().copied().unwrap_or_default());
    let high = u64::from(command.get(1).copied().unwrap_or_default()) << 8;
    let mut identifier = 1 | low | high;
    if identifier == original {
        identifier = original.saturating_add(1).max(1);
    }
    tsp::Reference {
        identifier,
        deprecated_type: (command.get(2).copied().unwrap_or_default() & 1 != 0).then_some(1),
        deprecated_is_external: (command.get(3).copied().unwrap_or_default() & 1 != 0)
            .then_some(true),
        ..tsp::Reference::default()
    }
    .encode_to_vec()
}

fn exercise_known_valid() {
    let source = tp::SectionTemplateArchive {
        headers: vec![reference(1), reference(2)],
        footers: vec![reference(3)],
        ..tp::SectionTemplateArchive::default()
    }
    .encode_to_vec();
    exercise_source(&source, &source);
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn exercise_limit_guards() {
    let source = tp::SectionTemplateArchive {
        headers: vec![reference(1)],
        footers: vec![reference(2)],
        ..tp::SectionTemplateArchive::default()
    }
    .encode_to_vec();
    let limited = pages_header_footer_codec::DecodeOptions::new(
        source.len().saturating_sub(1),
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
    );
    observe_decode(pages_header_footer_codec::decode_section_template(
        &source, limited,
    ));
    observe_decode(pages_header_footer_codec::decode_section_template(
        &source,
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            0,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
    ));
    observe_decode(pages_header_footer_codec::decode_section_template(
        &source,
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            0,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
    ));
    observe_decode(pages_header_footer_codec::decode_section_template(
        &source,
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            0,
            MAX_REFERENCES,
        ),
    ));
    observe_decode(pages_header_footer_codec::decode_section_template(
        &source,
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            0,
        ),
    ));
    observe_decode(pages_header_footer_codec::decode_section_template(
        &source, limited,
    ));

    let snapshot = pages_header_footer_codec::decode_section_template(&source, options(&source))
        .unwrap_or_else(|error| panic!("known limit seed must decode: {error}"));
    let headers = snapshot
        .headers()
        .map(|reference| pages_header_footer_codec::ReferenceWrite::new(reference.raw()))
        .collect::<Vec<_>>();
    let footers = snapshot
        .footers()
        .map(|reference| pages_header_footer_codec::ReferenceWrite::new(reference.raw()))
        .collect::<Vec<_>>();
    let write = pages_header_footer_codec::SectionTemplateWrite::new(&headers, &footers);
    for limited in [
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            source.len().saturating_sub(1),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            0,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            0,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            0,
            MAX_REFERENCES,
        ),
        pages_header_footer_codec::DecodeOptions::new(
            source.len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            0,
        ),
    ] {
        observe_rewrite(pages_header_footer_codec::rewrite_section_template(
            &source, write, limited,
        ));
    }
}

fn observe_decode<T>(result: Result<T, pages_header_footer_codec::DecodeError>) {
    if let Err(error) = result {
        observe_error(error);
    }
}

fn observe_rewrite(result: Result<Vec<u8>, pages_header_footer_codec::DecodeError>) {
    if let Err(error) = result {
        observe_error(error);
    }
}

fn observe_error(error: impl std::fmt::Display + std::fmt::Debug) {
    black_box((error.to_string(), format!("{error:?}")));
}
