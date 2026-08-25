#![no_main]

use std::{hint::black_box, num::NonZeroU64, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::pages_footnote_graph_codec::{
    self as graph, BodyFootnoteEntryWrite, BodyFootnoteTableWrite, DecodeError, DecodeLimit,
    DecodeOptions, FootnoteGraphWrite,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 256 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_ENTRIES: usize = 1_024;
const MAX_RECURSION: u32 = 64;
const MAX_TEXT_COMMAND_BYTES: usize = 512;

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_graph(&source);

    // Keep the exact-output, unknown-span, and malformed-table paths hot in
    // every local campaign without adding generated libFuzzer artifacts to
    // this focused corpus.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_graph(fixed);
        }
    });
});

const FIXED_CASES: &[&[u8]] = &[
    b"graph-basic",
    b"graph-unicode\xf0\x9f\x98\x80",
    b"graph-custom-mark",
    b"graph-structural-marker\x0e",
    b"graph-rewrite-unknown-group",
];

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

fn options() -> DecodeOptions {
    DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_ENTRIES,
    )
}

fn exercise_graph(source: &[u8]) {
    let text = command_text(source);
    let custom_mark = (control(source, 0) & 1 != 0).then(|| command_mark(source));
    let write = graph_write(source, &text, custom_mark.as_deref());
    let source_before = source.to_vec();

    let encoded = graph::encode_footnote_graph_with_report(write, options());
    assert_eq!(
        source,
        source_before.as_slice(),
        "graph encode modified input"
    );

    let Ok((payloads, report)) = encoded else {
        return;
    };
    assert_eq!(
        report.output_bytes(),
        report.reference_bytes()
            + report.storage_bytes()
            + report.marker_bytes()
            + report.body_entry_bytes()
    );
    assert_eq!(
        report.output_bytes(),
        payloads.reference().len()
            + payloads.storage().len()
            + payloads.marker().len()
            + payloads.body_entry().len()
    );
    assert!(report.output_bytes() <= MAX_OUTPUT_BYTES);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.allocations() <= 4);
    assert_eq!(report.retained_bytes(), report.output_bytes());
    assert_eq!(
        report.scratch_bytes(),
        "\u{fffc} ".len() + write.text().len()
    );

    let repeated = graph::encode_footnote_graph(write, options())
        .unwrap_or_else(|error| panic!("reported graph must encode identically: {error}"));
    assert_eq!(repeated, payloads);

    exercise_encode_limits(write, report);
    exercise_body_table(source, payloads.body_entry(), write);
}

fn exercise_encode_limits(write: FootnoteGraphWrite<'_>, report: graph::GraphEncodeReport) {
    let exact = DecodeOptions::new(
        MAX_INPUT_BYTES,
        report.output_bytes(),
        report.fields(),
        report.work_bytes(),
        MAX_RECURSION,
        MAX_ENTRIES,
    );
    graph::encode_footnote_graph(write, exact)
        .unwrap_or_else(|error| panic!("exact graph limits must accept: {error}"));

    if report.output_bytes() > 0 {
        assert_limit(
            graph::encode_footnote_graph(
                write,
                exact.with_max_output_bytes(report.output_bytes() - 1),
            ),
            DecodeLimit::OutputBytes {
                observed: report.output_bytes(),
                maximum: report.output_bytes() - 1,
            },
        );
    }
    if report.fields() > 0 {
        assert_limit(
            graph::encode_footnote_graph(
                write,
                DecodeOptions::new(
                    MAX_INPUT_BYTES,
                    report.output_bytes(),
                    report.fields() - 1,
                    report.work_bytes(),
                    MAX_RECURSION,
                    MAX_ENTRIES,
                ),
            ),
            DecodeLimit::Fields {
                observed: report.fields(),
                maximum: report.fields() - 1,
            },
        );
    }
    if report.work_bytes() > 0 {
        assert_limit(
            graph::encode_footnote_graph(
                write,
                DecodeOptions::new(
                    MAX_INPUT_BYTES,
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes() - 1,
                    MAX_RECURSION,
                    MAX_ENTRIES,
                ),
            ),
            DecodeLimit::WorkBytes {
                observed: report.work_bytes(),
                maximum: report.work_bytes() - 1,
            },
        );
    }
}

fn assert_limit(result: Result<graph::FootnoteGraphPayloads, DecodeError>, expected: DecodeLimit) {
    let error = result.expect_err("one-under graph resource budget must reject");
    assert_eq!(error.resource_limit(), Some(expected));
    black_box(error);
}

fn exercise_body_table(source: &[u8], body_entry: &[u8], write: FootnoteGraphWrite<'_>) {
    let mut table = Vec::with_capacity(body_entry.len().saturating_add(12));
    // Unknown overlong scalar and a balanced group are retained by rewrites.
    table.extend_from_slice(&[0x80, 0x02, 0x80, 0x00, 0x8b, 0x02, 0x80, 0x00, 0x8c, 0x02]);
    table.extend_from_slice(body_entry);
    let before = table.clone();
    let decoded = graph::decode_body_footnote_table_with_report(&table, options());
    let Ok((snapshot, report)) = decoded else {
        return;
    };
    assert_eq!(table, before);
    assert_eq!(snapshot.len(), 1);
    assert_eq!(report.source_bytes(), table.len());
    assert_eq!(report.entries(), 1);
    let entry = snapshot.entries().next().expect("one generated body entry");
    assert_eq!(entry.character_index(), write.character_index());
    assert_eq!(entry.reference_identifier(), write.reference_identifier());
    assert_eq!(entry.raw(), body_entry.get(2..).unwrap_or_default());
    black_box((source, report));

    let preserved = BodyFootnoteEntryWrite::preserve(entry);
    let operation = control(source, 1) % 3;
    let mut requested = Vec::new();
    match operation {
        0 => requested.push(preserved),
        1 => {
            let new_index = write.character_index().saturating_add(1);
            let new_identifier = distinct_identifier(write.reference_identifier(), 0x7fff);
            requested.push(preserved);
            requested.push(BodyFootnoteEntryWrite::new(new_index, new_identifier));
        },
        _ => {},
    }

    let rewritten = graph::rewrite_body_footnote_table_with_report(
        &table,
        BodyFootnoteTableWrite::new(&requested),
        options(),
    );
    let Ok((output, rewrite_report)) = rewritten else {
        return;
    };
    assert_eq!(output.len(), rewrite_report.output_bytes());
    assert_eq!(output.len(), rewrite_report.retained_bytes());
    assert_eq!(rewrite_report.entries_before(), 1);
    assert_eq!(rewrite_report.entries_after(), requested.len());
    assert_eq!(rewrite_report.allocations(), 1);
    assert!(rewrite_report.rewrite_work_bytes() <= MAX_WORK_BYTES);
    assert!(output.windows(2).any(|window| window == [0x80, 0x00]));
    let (candidate, candidate_report) =
        graph::decode_body_footnote_table_with_report(&output, options())
            .unwrap_or_else(|error| panic!("rewrite candidate must decode: {error}"));
    assert_eq!(candidate.len(), requested.len());
    assert_eq!(candidate_report.entries(), requested.len());
    for (actual, expected) in candidate.entries().zip(requested.iter().copied()) {
        assert_eq!(actual.character_index(), expected.character_index());
        assert_eq!(
            actual.reference_identifier(),
            expected.reference_identifier()
        );
    }
}

fn graph_write<'text>(
    source: &[u8],
    text: &'text str,
    custom_mark: Option<&'text str>,
) -> FootnoteGraphWrite<'text> {
    let reference = distinct_identifier(nonzero(read_u64(source, 0)), 1);
    let storage = distinct_identifier(reference, 2);
    let marker = distinct_identifier(storage, 3);
    let stylesheet = (control(source, 1) & 1 != 0).then(|| distinct_identifier(marker, 4));
    let paragraph =
        (control(source, 2) & 1 != 0).then(|| distinct_identifier(stylesheet.unwrap_or(marker), 5));
    let list = (control(source, 3) & 1 != 0)
        .then(|| distinct_identifier(paragraph.or(stylesheet).unwrap_or(marker), 6));
    FootnoteGraphWrite::new(reference, storage, marker, read_u32(source, 4), text)
        .with_custom_mark(custom_mark)
        .with_storage_template(stylesheet, paragraph, list, Some("en-US"))
}

fn command_text(source: &[u8]) -> String {
    let amount = source.len().min(MAX_TEXT_COMMAND_BYTES);
    let mut text = String::from_utf8_lossy(&source[..amount]).into_owned();
    text.retain(|character| character != '\u{000e}' && character != '\u{fffc}');
    if text.is_empty() {
        text.push_str("footnote");
    }
    text
}

fn command_mark(source: &[u8]) -> String {
    let mut mark = String::from_utf8_lossy(&source[source.len().min(1)..])
        .chars()
        .take(32)
        .filter(|character| *character != '\u{000e}' && *character != '\u{fffc}')
        .collect::<String>();
    if mark.is_empty() {
        mark.push('*');
    }
    mark
}

fn nonzero(value: u64) -> NonZeroU64 {
    NonZeroU64::new(value.max(1)).expect("positive graph identifier")
}

fn distinct_identifier(previous: NonZeroU64, offset: u64) -> NonZeroU64 {
    nonzero(previous.get().wrapping_add(offset).max(1))
}

fn read_u32(source: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        control(source, offset),
        control(source, offset + 1),
        control(source, offset + 2),
        control(source, offset + 3),
    ])
}

fn read_u64(source: &[u8], offset: usize) -> u64 {
    u64::from_le_bytes([
        control(source, offset),
        control(source, offset + 1),
        control(source, offset + 2),
        control(source, offset + 3),
        control(source, offset + 4),
        control(source, offset + 5),
        control(source, offset + 6),
        control(source, offset + 7),
    ])
}

fn control(source: &[u8], index: usize) -> u8 {
    source.get(index).copied().unwrap_or_default()
}
