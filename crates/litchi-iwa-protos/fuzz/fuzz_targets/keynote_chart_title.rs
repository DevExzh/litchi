#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_chart_title_codec::{
    ChartTitleWrite, DecodeLimit, DecodeOptions, DecodeReport, WireResourceLimit,
    decode_chart_title, decode_chart_title_text, decode_visible_chart_title,
    rewrite_chart_title_with_report,
};

// Keep both libFuzzer input and every codec allocation finite. Inputs larger
// than this are skipped rather than truncated, so the decoder always sees one
// unchanged caller-owned source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_TITLE_BYTES: usize = 64 * 1024;
const MAX_TITLE_INPUT_BYTES: usize = 4 * 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES + 1;
const PRIVATE_SENTINEL: &str = "__litchi_private_keynote_chart_title_6a45__";

const MALFORMED_SUFFIXES: &[&[u8]] = &[
    // Field 21 with a length-delimited wire type.
    &[0xaa, 0x01, 0x01],
    // Field 21 with a truncated varint value.
    &[0xa8, 0x01],
    // Field 23 with a non-canonical length varint.
    &[0xba, 0x01, 0x80, 0x00],
    // Field 23 with invalid UTF-8.
    &[0xba, 0x01, 0x01, 0xff],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);
    exercise_malformed_mutations(&source, data);

    // These checks are independent of the fuzzer's package source and need
    // only be constructed once per process. Keeping them in the target makes
    // the finite ceilings and redaction contract part of every campaign.
    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(|| {
        exercise_redacted_malformed();
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

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
    .with_max_title_bytes(MAX_TITLE_BYTES)
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let decoded = decode_chart_title(source, decode_options);
    assert_eq!(source, original.as_slice(), "decode modified its source");

    let Ok(snapshot) = decoded else {
        // Arbitrary mutations are expected to be rejected often. The error
        // still gets formatted so fuzzing covers every content-free failure
        // path without retaining or printing source bytes.
        if let Err(error) = decoded {
            observe_error(error);
        }
        return;
    };

    assert_eq!(
        snapshot.raw(),
        source,
        "snapshot did not retain exact source"
    );
    if let Some(title) = snapshot.title() {
        assert_borrowed(source, title.as_bytes());
    }
    if let Some(title) = snapshot.visible_title() {
        if !title.is_empty() {
            assert_borrowed(source, title.as_bytes());
        }
    }

    let (reported_snapshot, report) =
        decode_chart_title_with_report_checked(source, decode_options);
    assert_eq!(reported_snapshot, snapshot);
    assert_report(report, source.len());
    assert_eq!(
        source,
        original.as_slice(),
        "reported decode modified its source"
    );

    assert_eq!(
        decode_chart_title_text(source, decode_options)
            .unwrap_or_else(|error| panic!("title read disagreed with strict decode: {error}")),
        snapshot.title(),
    );
    assert_eq!(
        decode_visible_chart_title(source, decode_options).unwrap_or_else(|error| panic!(
            "visible title read disagreed with strict decode: {error}"
        )),
        snapshot.visible_title(),
    );
    assert_eq!(source, original.as_slice(), "read helpers modified source");

    exercise_noop(
        source,
        snapshot.title_visible(),
        snapshot.title(),
        decode_options,
    );

    let replacement = replacement_title(data);
    exercise_rewrite(
        source,
        snapshot.title_visible(),
        snapshot.title(),
        Some(true),
        Some(replacement.as_ref()),
        decode_options,
    );
    // Native removal is represented by an explicit hidden switch and an
    // absent title field, retaining proto2 presence semantics.
    exercise_rewrite(
        source,
        snapshot.title_visible(),
        snapshot.title(),
        Some(false),
        None,
        decode_options,
    );
}

fn decode_chart_title_with_report_checked<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> (
    litchi_iwa_protos::keynote_chart_title_codec::ChartTitleSnapshot<'source>,
    DecodeReport,
) {
    litchi_iwa_protos::keynote_chart_title_codec::decode_chart_title_with_report(source, options)
        .unwrap_or_else(|error| panic!("strict read disagreed with prior successful read: {error}"))
}

fn assert_report(report: DecodeReport, source_bytes: usize) {
    assert_eq!(report.source_bytes(), source_bytes);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(report.output_bytes() <= MAX_OUTPUT_BYTES.max(source_bytes));
    assert!(report.title_bytes() <= MAX_TITLE_BYTES);
}

fn assert_borrowed(source: &[u8], payload: &[u8]) {
    if payload.is_empty() {
        return;
    }
    let source_start = source.as_ptr() as usize;
    let source_end = source_start
        .checked_add(source.len())
        .expect("bounded source pointer range");
    let payload_start = payload.as_ptr() as usize;
    let payload_end = payload_start
        .checked_add(payload.len())
        .expect("bounded payload pointer range");
    assert!(
        payload_start >= source_start && payload_end <= source_end,
        "decoded title was not borrowed from the source"
    );
}

fn exercise_noop(
    source: &[u8],
    title_visible: Option<bool>,
    title: Option<&str>,
    options: DecodeOptions,
) {
    let (output, report) = rewrite_chart_title_with_report(
        source,
        ChartTitleWrite::new(title_visible, title),
        options,
    )
    .unwrap_or_else(|error| panic!("matching chart-title rewrite failed: {error}"));
    assert_eq!(output, source, "a matching rewrite was not byte-exact");
    assert!(!report.changed());
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), source.len());
    assert_eq!(
        source,
        &source.to_vec(),
        "no-op rewrite modified its source"
    );
    black_box(report);
}

fn exercise_rewrite(
    source: &[u8],
    before_visible: Option<bool>,
    before_title: Option<&str>,
    after_visible: Option<bool>,
    after_title: Option<&str>,
    decode_options: DecodeOptions,
) {
    let before = source.to_vec();
    let write = ChartTitleWrite::new(after_visible, after_title);
    let rewritten = rewrite_chart_title_with_report(source, write, decode_options);
    assert_eq!(source, before.as_slice(), "rewrite modified its source");
    let Ok((output, report)) = rewritten else {
        if let Err(error) = rewritten {
            // Near the finite source ceilings a changed write may correctly
            // run out of aggregate fields/work/output budget.
            observe_error(error);
        }
        return;
    };

    assert_eq!(
        report.changed(),
        before_visible != after_visible || before_title != after_title
    );
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), output.len());
    assert!(output.len() <= MAX_OUTPUT_BYTES.max(source.len()));

    let readback_options = options(&output);
    let (readback, readback_report) =
        decode_chart_title_with_report_checked(&output, readback_options);
    assert_eq!(readback.title_visible(), after_visible);
    assert_eq!(readback.title(), after_title);
    assert_eq!(
        readback.visible_title(),
        visible_title(after_visible, after_title)
    );
    assert_report(readback_report, output.len());
    assert_unknown_fields_preserved(source, &output);

    // Rewriting with the original presence/value pair is the wire-local
    // inverse. It must restore every unknown field and every original span,
    // not merely the selected semantic values.
    let restored = rewrite_chart_title_with_report(
        &output,
        ChartTitleWrite::new(before_visible, before_title),
        readback_options,
    );
    let Ok((restored, inverse_report)) = restored else {
        if let Err(error) = restored {
            observe_error(error);
        }
        return;
    };
    assert_eq!(
        restored, source,
        "chart-title inverse did not restore source"
    );
    assert_eq!(restored, before.as_slice());
    black_box(inverse_report);
}

fn visible_title<'title>(visible: Option<bool>, title: Option<&'title str>) -> Option<&'title str> {
    (visible == Some(true)).then(|| title.unwrap_or_default())
}

fn replacement_title(data: &[u8]) -> Cow<'_, str> {
    let start = data.len().min(8);
    let end = data.len().min(start.saturating_add(MAX_TITLE_INPUT_BYTES));
    if start == end {
        return Cow::Borrowed("fuzz chart title");
    }
    Cow::Owned(String::from_utf8_lossy(&data[start..end]).into_owned())
}

fn exercise_malformed_mutations(source: &[u8], data: &[u8]) {
    if !source.is_empty() {
        let mut mutated = source.to_vec();
        let index = usize::from(data.first().copied().unwrap_or_default()) % mutated.len();
        mutated[index] ^= 0xff;
        observe_mutation(&mutated, data);
    }

    for suffix in MALFORMED_SUFFIXES {
        // Keep the malformed recipe independent from the arbitrary source.
        // Appending a malformed suffix to arbitrary bytes can accidentally
        // turn the suffix into a valid unknown field (or make its bytes part
        // of a preceding length-delimited span), causing a false harness
        // assertion rather than testing the intended rejection path.
        let malformed = suffix.to_vec();
        let before = malformed.clone();
        let decoded = decode_chart_title(&malformed, options(&malformed));
        assert!(decoded.is_err(), "known malformed mutation was accepted");
        if let Err(error) = decoded {
            observe_error(error);
        }
        assert_eq!(
            malformed,
            before.as_slice(),
            "malformed read modified source"
        );

        let rewritten = rewrite_chart_title_with_report(
            &malformed,
            ChartTitleWrite::new(Some(true), Some("mutation")),
            options(&malformed),
        );
        assert!(rewritten.is_err(), "malformed mutation was rewritten");
        if let Err(error) = rewritten {
            observe_error(error);
        }
        assert_eq!(
            malformed,
            before.as_slice(),
            "malformed rewrite modified source"
        );
    }
}

fn observe_mutation(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    match decode_chart_title(source, options(source)) {
        Ok(snapshot) => {
            black_box((
                snapshot.title_visible(),
                snapshot.title(),
                snapshot.visible_title(),
            ));
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(source, before.as_slice(), "mutated read modified source");

    let rewritten = rewrite_chart_title_with_report(
        source,
        ChartTitleWrite::new(Some(true), Some(replacement_title(data).as_ref())),
        options(source),
    );
    match rewritten {
        Ok((output, report)) => {
            black_box((output, report));
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(source, before.as_slice(), "mutated rewrite modified source");
}

fn exercise_redacted_malformed() {
    let mut malformed = field_text(4000, PRIVATE_SENTINEL.as_bytes());
    malformed.extend_from_slice(&[0xa8, 0x01]);
    let options = options(&malformed);
    let decode_error = decode_chart_title(&malformed, options)
        .expect_err("private malformed chart-title sentinel was accepted");
    observe_redacted(decode_error, PRIVATE_SENTINEL);
    let rewrite_error = rewrite_chart_title_with_report(
        &malformed,
        ChartTitleWrite::new(Some(true), Some("redacted")),
        options,
    )
    .expect_err("private malformed chart-title rewrite was accepted");
    observe_redacted(rewrite_error, PRIVATE_SENTINEL);
}

fn exercise_limit_guards() {
    let source = [field_varint(21, 1), field_text(23, b"abcd")].concat();

    let field_error = decode_chart_title(
        &source,
        DecodeOptions::new(source.len(), 1, source.len() * 8, 8),
    )
    .expect_err("field limit was not enforced");
    assert_eq!(field_error.field_limit_values(), Some((2, 1)));

    let work_error = decode_chart_title(
        &source,
        DecodeOptions::new(source.len(), 8, source.len(), 8),
    )
    .expect_err("work limit was not enforced");
    assert_eq!(
        work_error.work_limit_values(),
        Some((source.len() * 2, source.len()))
    );

    let output_error = decode_chart_title(
        &source,
        DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_output_bytes(3),
    )
    .expect_err("output limit was not enforced");
    assert_eq!(output_error.output_limit_values(), Some((4, 3)));

    let title_error = decode_chart_title(
        &source,
        DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_title_bytes(3),
    )
    .expect_err("title limit was not enforced");
    assert_eq!(title_error.title_limit_values(), Some((4, 3)));

    let nesting = [0x0b, 0x13, 0x14, 0x0c];
    let nesting_error = decode_chart_title(
        &nesting,
        DecodeOptions::new(nesting.len(), 8, nesting.len() * 8, 1),
    )
    .expect_err("nesting limit was not enforced");
    assert_eq!(
        nesting_error.wire_resource_limit(),
        Some(WireResourceLimit::Nesting {
            observed: 2,
            maximum: 1,
        })
    );

    // Keep all typed limit variants reachable from the target binary even if
    // a local campaign happens to start from only malformed wire inputs.
    black_box((
        DecodeLimit::Bytes {
            observed: MAX_INPUT_BYTES + 1,
            maximum: MAX_INPUT_BYTES,
        },
        DecodeLimit::Fields {
            observed: MAX_FIELDS + 1,
            maximum: MAX_FIELDS,
        },
    ));
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let source = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    let options = DecodeOptions::new(MAX_INPUT_BYTES, MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION)
        .with_max_output_bytes(MAX_OUTPUT_BYTES)
        .with_max_title_bytes(MAX_TITLE_BYTES);
    let error =
        decode_chart_title(source, options).expect_err("oversized chart-title source was accepted");
    assert_eq!(
        error.wire_resource_limit(),
        Some(WireResourceLimit::Bytes {
            observed: OVERSIZED_INPUT_BYTES,
            maximum: MAX_INPUT_BYTES,
        })
    );
    observe_error(error);
}

fn assert_unknown_fields_preserved(source: &[u8], output: &[u8]) {
    for raw in unknown_field_spans(source) {
        assert!(
            output.windows(raw.len()).any(|window| window == raw),
            "rewrite dropped an unknown chart-title wire span"
        );
    }
}

fn unknown_field_spans(source: &[u8]) -> Vec<&[u8]> {
    let mut spans = Vec::new();
    let mut offset = 0;
    while offset < source.len() {
        let start = offset;
        let Some((tag, next)) = read_varint(source, offset) else {
            return Vec::new();
        };
        offset = next;
        let field_number = tag >> 3;
        let wire_type = tag & 7;
        if !skip_wire(source, &mut offset, wire_type, field_number) {
            return Vec::new();
        }
        if field_number != 21 && field_number != 23 {
            spans.push(&source[start..offset]);
        }
    }
    spans
}

fn skip_wire(source: &[u8], offset: &mut usize, wire_type: u64, group: u64) -> bool {
    match wire_type {
        0 => read_varint(source, *offset)
            .map(|(_, end)| *offset = end)
            .is_some(),
        1 => advance(source, offset, 8),
        2 => {
            let Some((length, end)) = read_varint(source, *offset) else {
                return false;
            };
            let Ok(length) = usize::try_from(length) else {
                return false;
            };
            *offset = end;
            advance(source, offset, length)
        },
        3 => loop {
            let Some((tag, end)) = read_varint(source, *offset) else {
                return false;
            };
            *offset = end;
            let field_number = tag >> 3;
            let child_wire = tag & 7;
            if child_wire == 4 {
                return field_number == group;
            }
            if !skip_wire(source, offset, child_wire, field_number) {
                return false;
            }
        },
        4 | 6 | 7 => false,
        5 => advance(source, offset, 4),
        _ => false,
    }
}

fn advance(source: &[u8], offset: &mut usize, amount: usize) -> bool {
    let Some(end) = offset.checked_add(amount) else {
        return false;
    };
    if end > source.len() {
        return false;
    }
    *offset = end;
    true
}

fn read_varint(source: &[u8], mut offset: usize) -> Option<(u64, usize)> {
    let start = offset;
    let mut value = 0u64;
    for shift in (0..64).step_by(7) {
        let byte = *source.get(offset)?;
        offset = offset.checked_add(1)?;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some((value, offset));
        }
    }
    let _ = start;
    None
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    [varint(u64::from(number) << 3), varint(value)].concat()
}

fn field_text(number: u32, value: &[u8]) -> Vec<u8> {
    [
        varint((u64::from(number) << 3) | 2),
        varint(value.len() as u64),
        value.to_vec(),
    ]
    .concat()
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

fn observe_redacted(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
