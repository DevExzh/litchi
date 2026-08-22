#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_chart_title_codec::{
    ChartTitleWrite, DecodeLimit, DecodeOptions, DecodeReport, WireResourceLimit,
    decode_chart_title, decode_chart_title_extension, decode_chart_title_text,
    decode_visible_chart_title, rewrite_chart_title, rewrite_chart_title_extension,
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
    exercise_selected_unknown_interleavings(data);
    exercise_malformed_mutations(&source, data);

    // These checks are independent of the fuzzer's package source and need
    // only be constructed once per process. Keeping them in the target makes
    // the finite ceilings and redaction contract part of every campaign.
    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(|| {
        exercise_redacted_malformed();
        exercise_limit_guards();
        exercise_input_limit();
        exercise_atomic_rewrite_limit();
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
    assert_eq!(
        decode_chart_title_extension(source, decode_options)
            .unwrap_or_else(|error| panic!("extension read disagreed with strict decode: {error}")),
        snapshot,
        "extension read disagreed with strict decode"
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
        snapshot.title_visible(),
        snapshot.title(),
        decode_options,
    );

    let replacement = replacement_title(data);
    exercise_semantic_matrix(
        source,
        snapshot.title_visible(),
        snapshot.title(),
        replacement.as_ref(),
        decode_options,
    );
    exercise_atomic_rewrite_failure(
        source,
        snapshot.title_visible(),
        snapshot.title(),
        replacement.as_ref(),
        decode_options,
    );
}

fn exercise_selected_unknown_interleavings(data: &[u8]) {
    // Random mutations rarely retain a complete pair of selected fields and
    // then place unknown spans on both sides of that pair. Keep a few tiny,
    // valid layouts in the target so wire-local replacement is exercised for
    // every selected/unknown ordering. The selected fields are deliberately
    // reversed in one case; protobuf field-number order is not a wire-order
    // requirement, and the codec must preserve that source layout.
    let title = replacement_title(data);
    let unknown_before = field_text(4000, b"before");
    let unknown_between = field_varint(4001, u64::from(data.first().copied().unwrap_or(7)));
    let unknown_after = field_text(4002, b"after");
    let unknown_tail = field_varint(4003, u64::from(data.get(1).copied().unwrap_or(9)));
    let cases = [
        [
            unknown_before.clone(),
            field_varint(21, 1),
            unknown_between.clone(),
            field_text(23, title.as_bytes()),
            unknown_after.clone(),
        ]
        .concat(),
        [
            field_text(23, title.as_bytes()),
            unknown_before.clone(),
            field_varint(21, 0),
            unknown_between.clone(),
            unknown_tail.clone(),
        ]
        .concat(),
        [
            unknown_before,
            unknown_between,
            field_text(23, title.as_bytes()),
            unknown_after,
            field_varint(21, 1),
            unknown_tail,
        ]
        .concat(),
    ];

    for source in cases {
        exercise_source(&source, data);
    }
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
    expected_visible: Option<bool>,
    expected_title: Option<&str>,
    write_visible: Option<bool>,
    write_title: Option<&str>,
    options: DecodeOptions,
) {
    let (output, report) = rewrite_chart_title_with_report(
        source,
        ChartTitleWrite::new(write_visible, write_title),
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
    let readback = decode_chart_title(source, options)
        .unwrap_or_else(|error| panic!("no-op rewrite made source unreadable: {error}"));
    assert_eq!(readback.title_visible(), expected_visible);
    assert_eq!(readback.title(), expected_title);

    let alias_output = rewrite_chart_title(
        source,
        ChartTitleWrite::new(write_visible, write_title),
        options,
    )
    .unwrap_or_else(|error| panic!("chart-title rewrite alias failed: {error}"));
    assert_eq!(alias_output, source, "rewrite alias was not byte-exact");

    let extension_output = rewrite_chart_title_extension(
        source,
        ChartTitleWrite::new(write_visible, write_title),
        options,
    )
    .unwrap_or_else(|error| panic!("chart-title extension rewrite failed: {error}"));
    assert_eq!(
        extension_output, source,
        "extension rewrite alias was not byte-exact"
    );
    assert_eq!(source, &source.to_vec(), "no-op aliases modified source");
    black_box(report);
}

fn exercise_semantic_matrix(
    source: &[u8],
    before_visible: Option<bool>,
    before_title: Option<&str>,
    replacement: &str,
    options: DecodeOptions,
) {
    // Keep the complete proto2 presence/value matrix in the target. In
    // particular, a hidden title is still a present field, and an explicitly
    // visible title without field 23 reads back as the native empty default.
    // The final pair is the editor's set operation; (false, None) is clear.
    let writes = [
        ChartTitleWrite::new(None, None),
        ChartTitleWrite::new(None, Some("")),
        ChartTitleWrite::new(None, Some(replacement)),
        ChartTitleWrite::new(Some(false), None),
        ChartTitleWrite::new(Some(false), Some("")),
        ChartTitleWrite::new(Some(false), Some(replacement)),
        ChartTitleWrite::new(Some(true), None),
        ChartTitleWrite::new(Some(true), Some("")),
        ChartTitleWrite::new(Some(true), Some(replacement)),
    ];
    for write in writes {
        exercise_rewrite(
            source,
            before_visible,
            before_title,
            write.title_visible(),
            write.title(),
            options,
        );
    }
}

fn exercise_atomic_rewrite_failure(
    source: &[u8],
    before_visible: Option<bool>,
    before_title: Option<&str>,
    replacement: &str,
    options: DecodeOptions,
) {
    // Force the candidate title below its configured ceiling. The strict
    // source read may succeed, but publication must fail before any candidate
    // can escape. Re-read the source after the error to make the atomicity
    // assertion observable rather than only checking the borrowed slice.
    let before = source.to_vec();
    let capped = options.with_max_title_bytes(replacement.len().saturating_sub(1));
    let result = rewrite_chart_title_with_report(
        source,
        ChartTitleWrite::new(Some(true), Some(replacement)),
        capped,
    );
    let error = result.expect_err("title-capped rewrite unexpectedly succeeded");
    observe_error(error);
    assert_eq!(
        source,
        before.as_slice(),
        "failed rewrite modified its source"
    );

    let after = decode_chart_title(source, options).unwrap_or_else(|decode_error| {
        panic!("source became unreadable after failed rewrite: {decode_error}")
    });
    assert_eq!(after.title_visible(), before_visible);
    assert_eq!(after.title(), before_title);
}

fn exercise_rewrite(
    source: &[u8],
    before_visible: Option<bool>,
    before_title: Option<&str>,
    after_visible: Option<bool>,
    after_title: Option<&str>,
    decode_options: DecodeOptions,
) {
    // The public semantic editor treats clearing an already-hidden or absent
    // title as an exact no-op. A hidden field-23 value is stale wire state
    // rather than a visible title, so do not ask the wire-level codec to
    // represent a transition that changes only hidden presence state. The
    // visible case still rewrites field 21 to false and removes field 23.
    if clear_is_noop(before_visible, after_visible, after_title) {
        exercise_noop(
            source,
            before_visible,
            before_title,
            after_visible,
            after_title,
            decode_options,
        );
        return;
    }

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
    if selected_field_presence(source) == selected_field_presence(&output) {
        assert_selected_unknown_interleaving_preserved(source, &output);
    }

    // Rewriting with the original presence/value pair is the wire-local
    // inverse. It must restore every unknown field and every original span,
    // not merely the selected semantic values. A clear request against a
    // hidden/absent title is an intentional exact no-op, so that inverse
    // cannot restore a previously present hidden field-23 span.
    // The first rewrite may temporarily remove every selected field, making
    // the candidate shorter than the original. Give the inverse pass a
    // source ceiling that covers both sides of the round trip; reusing the
    // candidate-sized policy would reject a valid restoration before the
    // codec can check it.
    let inverse_options = if source.len() >= output.len() {
        options(source)
    } else {
        options(&output)
    };
    let restored = rewrite_chart_title_with_report(
        &output,
        ChartTitleWrite::new(before_visible, before_title),
        inverse_options,
    );
    let (restored, inverse_report) = restored.unwrap_or_else(|error| {
        panic!("successful chart-title rewrite was not invertible: {error}")
    });
    if clear_is_noop(after_visible, before_visible, before_title) {
        assert_eq!(
            restored, output,
            "clear no-op inverse changed a hidden/absent title source"
        );
        assert!(!inverse_report.changed());
        assert_eq!(inverse_report.input_bytes(), output.len());
        assert_eq!(inverse_report.output_bytes(), output.len());
        assert_unknown_fields_preserved(source, &restored);
        black_box(inverse_report);
        return;
    }

    let selected_positions_lost =
        selected_field_presence(source) != selected_field_presence(&output);
    if restored != source && selected_positions_lost {
        // Removing or replacing a selected span erases its historical wire
        // position. The inverse can only append that span after the surviving
        // source spans; check semantic restoration and exact unknown
        // preservation without pretending that the original selected/unknown
        // interleaving remains recoverable.
        let inverse_readback =
            decode_chart_title_with_report_checked(&restored, options(&restored));
        assert_eq!(inverse_readback.0.title_visible(), before_visible);
        assert_eq!(inverse_readback.0.title(), before_title);
        assert_report(inverse_readback.1, restored.len());
        assert_unknown_fields_preserved(source, &restored);
    } else if restored != source && selected_field_order_is_only_difference(source, &restored) {
        // With no selected span removed by the first write, the codec's
        // deterministic inverse can differ only by swapping fields 21 and
        // 23 within their existing selected slots.
        let inverse_readback =
            decode_chart_title_with_report_checked(&restored, options(&restored));
        assert_eq!(inverse_readback.0.title_visible(), before_visible);
        assert_eq!(inverse_readback.0.title(), before_title);
        assert_report(inverse_readback.1, restored.len());
        assert_unknown_fields_preserved(source, &restored);
    } else {
        assert_eq!(
            restored, source,
            "chart-title inverse did not restore source"
        );
        assert_eq!(restored, before.as_slice());
    }
    assert_eq!(inverse_report.changed(), report.changed());
    assert_eq!(inverse_report.input_bytes(), output.len());
    assert_eq!(inverse_report.output_bytes(), restored.len());
    black_box(inverse_report);
}

fn clear_is_noop(
    current_visible: Option<bool>,
    write_visible: Option<bool>,
    write_title: Option<&str>,
) -> bool {
    write_visible == Some(false) && write_title.is_none() && current_visible != Some(true)
}

fn visible_title<'title>(visible: Option<bool>, title: Option<&'title str>) -> Option<&'title str> {
    (visible == Some(true)).then(|| title.unwrap_or_default())
}

fn selected_field_order_is_only_difference(source: &[u8], output: &[u8]) -> bool {
    let Some(source_spans) = wire_field_spans(source) else {
        return false;
    };
    let Some(output_spans) = wire_field_spans(output) else {
        return false;
    };
    if source_spans.len() != output_spans.len() {
        return false;
    }

    // Unknown spans must occupy exactly the same wire slots. Selected spans
    // may swap with one another, but they may not cross an unknown span.
    let mut source_selected = Vec::new();
    let mut output_selected = Vec::new();
    for ((source_number, source_raw), (output_number, output_raw)) in
        source_spans.iter().zip(output_spans.iter())
    {
        let source_is_selected = is_selected_field(*source_number);
        let output_is_selected = is_selected_field(*output_number);
        if source_is_selected != output_is_selected {
            return false;
        }
        if source_is_selected {
            source_selected.push(*source_raw);
            output_selected.push(*output_raw);
        } else if source_raw != output_raw {
            return false;
        }
    }

    source_selected.len() == output_selected.len()
        && source_selected
            .iter()
            .all(|span| output_selected.iter().any(|candidate| candidate == span))
        && output_selected
            .iter()
            .all(|span| source_selected.iter().any(|candidate| candidate == span))
}

fn assert_selected_unknown_interleaving_preserved(source: &[u8], output: &[u8]) {
    let source_spans =
        wire_field_spans(source).expect("strict source decode produced unparseable wire spans");
    let output_spans =
        wire_field_spans(output).expect("strict output decode produced unparseable wire spans");
    assert_eq!(
        source_spans.len(),
        output_spans.len(),
        "rewrite changed selected/unknown span count without changing presence"
    );
    for ((source_number, source_raw), (output_number, output_raw)) in
        source_spans.iter().zip(output_spans.iter())
    {
        assert_eq!(
            is_selected_field(*source_number),
            is_selected_field(*output_number),
            "rewrite changed selected/unknown field interleaving"
        );
        if !is_selected_field(*source_number) {
            assert!(
                source_raw == output_raw,
                "rewrite moved or changed an unknown wire span"
            );
        }
    }
}

fn selected_field_presence(source: &[u8]) -> Vec<u64> {
    let mut fields = wire_field_spans(source)
        .unwrap_or_default()
        .into_iter()
        .filter_map(|(number, _)| is_selected_field(number).then_some(number))
        .collect::<Vec<_>>();
    fields.sort_unstable();
    fields
}

fn is_selected_field(number: u64) -> bool {
    matches!(number, 21 | 23)
}

fn wire_field_spans<'source>(source: &'source [u8]) -> Option<Vec<(u64, &'source [u8])>> {
    let mut spans = Vec::new();
    let mut offset = 0;
    while offset < source.len() {
        let start = offset;
        let (tag, next) = read_varint(source, offset)?;
        offset = next;
        let field_number = tag >> 3;
        let wire_type = tag & 7;
        if !skip_wire(source, &mut offset, wire_type, field_number) {
            return None;
        }
        spans.push((field_number, &source[start..offset]));
    }
    Some(spans)
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

fn exercise_atomic_rewrite_limit() {
    // This reaches the rewrite's output preflight after a successful source
    // decode. It verifies that an output-budget failure has no partial result
    // and leaves the caller-owned source semantically untouched.
    let source = field_varint(4000, 7);
    let before = source.clone();
    let options = DecodeOptions::new(source.len(), 128, source.len() * 8, 8)
        .with_max_output_bytes(source.len())
        .with_max_title_bytes(128);
    let error = rewrite_chart_title_with_report(
        &source,
        ChartTitleWrite::new(Some(true), Some("atomic")),
        options,
    )
    .expect_err("output-capped rewrite unexpectedly succeeded");
    assert!(
        error.output_limit_values().is_some(),
        "output-capped rewrite returned an unrelated error: {error}"
    );
    observe_error(error);
    assert_eq!(source, before, "failed output rewrite modified its source");
    let snapshot = decode_chart_title(
        &source,
        DecodeOptions::new(source.len(), 128, source.len() * 8, 8)
            .with_max_output_bytes(128)
            .with_max_title_bytes(128),
    )
    .expect("output-capped rewrite made source unreadable");
    assert_eq!(snapshot.title_visible(), None);
    assert_eq!(snapshot.title(), None);
}

fn assert_unknown_fields_preserved(source: &[u8], output: &[u8]) {
    assert_eq!(
        unknown_field_spans(output),
        unknown_field_spans(source),
        "rewrite changed the order or bytes of an unknown chart-title wire span"
    );
}

fn selected_field_spans(source: &[u8]) -> Vec<&[u8]> {
    wire_field_spans(source)
        .unwrap_or_default()
        .into_iter()
        .filter_map(|(number, raw)| is_selected_field(number).then_some(raw))
        .collect()
}

fn unknown_field_spans(source: &[u8]) -> Vec<&[u8]> {
    wire_field_spans(source)
        .unwrap_or_default()
        .into_iter()
        .filter_map(|(number, raw)| (!is_selected_field(number)).then_some(raw))
        .collect()
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
    let mut value = 0u64;
    for shift in (0..64).step_by(7) {
        let byte = *source.get(offset)?;
        offset = offset.checked_add(1)?;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some((value, offset));
        }
    }
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
