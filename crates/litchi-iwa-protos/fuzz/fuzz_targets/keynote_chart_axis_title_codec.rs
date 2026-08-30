#![no_main]

//! Strict, source-preserving fuzzing for Keynote chart-axis title fields.
//!
//! The target owns only the generated extension payload.  Chart graph
//! selection and package transactions belong to the neighboring package
//! target.  Every successful case keeps category (13/15) and value (14/16)
//! fields in one source, with unknown spans interleaved around the selected
//! fields so a rewrite cannot silently normalize unrelated data.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_chart_axis_title_codec::{
    AxisTitleKind, AxisTitleSnapshot, AxisTitleWrite, DecodeLimit, DecodeOptions, DecodeReport,
    RewriteExecutionLimits, WireResourceLimit, decode_axis_title, decode_axis_title_extension,
    decode_axis_title_text, decode_axis_titles, decode_axis_titles_with_report,
    decode_visible_axis_title, prepare_axis_title_rewrite, rewrite_axis_title,
    rewrite_axis_title_with_report,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_TITLE_BYTES: usize = 64 * 1024;
const MAX_TITLE_INPUT_BYTES: usize = 4 * 1024;
const MAX_ALLOCATIONS: usize = 4;
const MAX_RETAINED_BYTES: usize = 256 * 1024;
const MAX_SCRATCH_BYTES: usize = 128 * 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES + 1;
const PRIVATE_SENTINEL: &str = "__litchi_private_keynote_chart_axis_title_112__";

const KINDS: [AxisTitleKind; 2] = [AxisTitleKind::Category, AxisTitleKind::Value];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);
    exercise_matrix(data);
    exercise_malformed(data);

    // These fixed guards make the finite resource and redaction contracts
    // reachable even when a campaign starts with only malformed input.
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
    .with_max_allocations(MAX_ALLOCATIONS)
    .with_max_retained_bytes(MAX_RETAINED_BYTES.max(source.len()))
    .with_max_scratch_bytes(MAX_SCRATCH_BYTES.max(source.len()))
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let decoded = decode_axis_titles(source, decode_options);
    assert_eq!(source, original.as_slice(), "decode modified its source");
    let Ok(snapshot) = decoded else {
        if let Err(error) = decoded {
            observe_error(error);
        }
        return;
    };

    assert_eq!(snapshot.raw(), source, "snapshot lost the source slice");
    assert_eq!(
        decode_axis_title(source, decode_options)
            .unwrap_or_else(|error| panic!("axis-title singular read disagreed: {error}")),
        snapshot
    );
    assert_eq!(
        decode_axis_title_extension(source, decode_options)
            .unwrap_or_else(|error| panic!("axis-title extension read disagreed: {error}")),
        snapshot
    );
    assert_eq!(
        decode_axis_title_text(source, decode_options)
            .unwrap_or_else(|error| panic!("axis-title text read disagreed: {error}")),
        snapshot.title()
    );
    assert_eq!(
        decode_visible_axis_title(source, decode_options)
            .unwrap_or_else(|error| panic!("visible axis-title read disagreed: {error}")),
        snapshot.visible_title(AxisTitleKind::Category)
    );
    if let Some(title) = snapshot.title() {
        assert_borrowed(source, title.as_bytes());
    }
    if let Some(title) = snapshot.value_title() {
        assert_borrowed(source, title.as_bytes());
    }
    for kind in KINDS {
        if let Some(title) = snapshot.visible_title(kind)
            && !title.is_empty()
        {
            assert_borrowed(source, title.as_bytes());
        }
    }
    let (reported, report) = decode_axis_titles_with_report(source, decode_options)
        .unwrap_or_else(|error| panic!("reported axis-title decode disagreed: {error}"));
    assert_eq!(reported, snapshot);
    assert_report(report, source.len());
    assert_eq!(
        source,
        original.as_slice(),
        "reported decode modified its source"
    );

    for kind in KINDS {
        black_box(snapshot.visible_title(kind));
        exercise_kind(source, snapshot, kind, data, decode_options);
    }

    // The source-derived write carries both category and value fields.  A
    // matching write is required to be byte-exact, including unknown spans.
    let preserved = AxisTitleWrite::preserve();
    let (output, report) = rewrite_axis_title_with_report(source, preserved, decode_options)
        .unwrap_or_else(|error| panic!("matching axis-title rewrite failed: {error}"));
    assert_eq!(
        output, source,
        "matching axis-title rewrite was not a no-op"
    );
    assert!(!report.changed());
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), source.len());
    assert_unknown_spans_preserved(source, &output);
    let alias = rewrite_axis_title(source, preserved, decode_options)
        .unwrap_or_else(|error| panic!("axis-title rewrite alias failed: {error}"));
    assert_eq!(alias, source, "axis-title rewrite alias was not byte-exact");
    exercise_prepared_write(
        source,
        preserved,
        snapshot.visible_title(AxisTitleKind::Category),
        snapshot.visible_title(AxisTitleKind::Value),
        decode_options,
    );
}

fn exercise_kind(
    source: &[u8],
    snapshot: AxisTitleSnapshot<'_>,
    kind: AxisTitleKind,
    data: &[u8],
    decode_options: DecodeOptions,
) {
    let before = snapshot.visible_title(kind);
    let replacement = replacement_title(data);
    let writes = [Some(replacement.as_str()), Some(""), None];

    for requested in writes {
        let write = AxisTitleWrite::preserve().with_title(kind, requested);
        let source_before = source.to_vec();
        let rewritten = rewrite_axis_title_with_report(source, write, decode_options);
        assert_eq!(
            source,
            source_before.as_slice(),
            "rewrite modified its source"
        );
        let Ok((output, report)) = rewritten else {
            if let Err(error) = rewritten {
                observe_error(error);
            }
            continue;
        };
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.output_bytes(), output.len());
        assert!(output.len() <= MAX_OUTPUT_BYTES.max(source.len()));
        assert_unknown_spans_preserved(source, &output);

        let readback = decode_axis_titles(&output, options(&output))
            .unwrap_or_else(|error| panic!("axis-title rewrite made output unreadable: {error}"));
        let expected = requested.map(|title| title.to_owned());
        assert_eq!(readback.visible_title(kind), expected.as_deref());
        for opposite in KINDS {
            if opposite != kind {
                assert_eq!(
                    readback.visible_title(opposite),
                    snapshot.visible_title(opposite),
                    "axis-title rewrite changed the opposite axis"
                );
            }
        }
        let expected_category = if kind == AxisTitleKind::Category {
            expected.as_deref()
        } else {
            snapshot.visible_title(AxisTitleKind::Category)
        };
        let expected_value = if kind == AxisTitleKind::Value {
            expected.as_deref()
        } else {
            snapshot.visible_title(AxisTitleKind::Value)
        };
        exercise_prepared_write(
            source,
            write,
            expected_category,
            expected_value,
            decode_options,
        );

        // A second source-derived write is the semantic inverse.  If a
        // selected field was absent, exact historical wire placement cannot
        // be recovered; unknown spans and both semantic axes still must be.
        let inverse = AxisTitleWrite::preserve().with_title(kind, before);
        let restored = rewrite_axis_title(&output, inverse, options(&output));
        match restored {
            Ok(restored) => {
                let restored_snapshot = decode_axis_titles(&restored, options(&restored))
                    .unwrap_or_else(|error| {
                        panic!("axis-title inverse made output unreadable: {error}")
                    });
                assert_eq!(restored_snapshot.visible_title(kind), before);
                assert_eq!(
                    restored_snapshot.visible_title(if kind == AxisTitleKind::Category {
                        AxisTitleKind::Value
                    } else {
                        AxisTitleKind::Category
                    }),
                    snapshot.visible_title(if kind == AxisTitleKind::Category {
                        AxisTitleKind::Value
                    } else {
                        AxisTitleKind::Category
                    })
                );
                // A remove followed by an inverse add has no retained wire
                // position for the missing selected field. The codec appends
                // that field by contract, so only unrelated spans—not the
                // historical selected/unknown interleaving—are invariant.
                assert_unknown_span_bytes_preserved(source, &restored);
            },
            Err(error) => observe_error(error),
        }
    }
}

fn exercise_prepared_write(
    source: &[u8],
    write: AxisTitleWrite<'_>,
    expected_category: Option<&str>,
    expected_value: Option<&str>,
    decode_options: DecodeOptions,
) {
    let original = source.to_vec();
    let prepared = match prepare_axis_title_rewrite(source, write, decode_options) {
        Ok(prepared) => prepared,
        Err(error) => {
            // A valid wire payload can still be outside the aggregate work,
            // output, or retained-byte policy. Those are expected fuzzer
            // outcomes, and the source must remain untouched in all cases.
            observe_error(error);
            assert_eq!(source, original.as_slice(), "prepare modified its source");
            return;
        },
    };
    let requirements = prepared.execution_requirements();
    let preparation = prepared.prepare_report();
    assert_eq!(preparation.source_bytes(), source.len());
    assert!(preparation.fields() <= requirements.fields);
    assert!(preparation.work_bytes() <= requirements.work_bytes);
    assert!(preparation.max_depth() <= requirements.max_depth);
    assert_eq!(source, original.as_slice(), "prepare modified its source");

    let output = prepared
        .execute(RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact axis-title prepared replay failed: {error}"));
    assert_eq!(output.as_bytes().len(), requirements.output_bytes);
    assert_eq!(output.report().fields(), requirements.fields);
    assert_eq!(output.report().work_bytes(), requirements.work_bytes);
    assert_eq!(output.report().max_depth(), requirements.max_depth);
    assert_eq!(output.report().allocations(), requirements.allocations);
    assert_eq!(
        output.report().retained_bytes(),
        requirements.retained_bytes
    );
    assert_eq!(output.report().scratch_bytes(), requirements.scratch_bytes);
    assert_eq!(
        output.report().changed(),
        output.as_bytes() != source,
        "prepared changed flag disagreed with output bytes"
    );
    assert_eq!(source, original.as_slice(), "execute modified its source");

    let candidate = output.as_bytes().to_vec();
    let readback = decode_axis_titles(&candidate, options(&candidate))
        .unwrap_or_else(|error| panic!("axis-title prepared candidate was unreadable: {error}"));
    assert_eq!(
        readback.visible_title(AxisTitleKind::Category),
        expected_category
    );
    assert_eq!(readback.visible_title(AxisTitleKind::Value), expected_value);
    assert_unknown_spans_preserved(source, &candidate);

    let one_shot = rewrite_axis_title(source, write, decode_options)
        .unwrap_or_else(|error| panic!("axis-title one-shot replay failed: {error}"));
    assert_eq!(one_shot, candidate, "prepared and one-shot rewrites differ");
    assert_eq!(source, original.as_slice(), "one-shot modified its source");

    let exact = RewriteExecutionLimits::exact(requirements);
    let mut limited = Vec::new();
    if requirements.output_bytes > 0 {
        limited.push(exact.with_output_bytes(requirements.output_bytes - 1));
    }
    if requirements.fields > 0 {
        limited.push(exact.with_fields(requirements.fields - 1));
    }
    if requirements.work_bytes > 0 {
        limited.push(exact.with_work_bytes(requirements.work_bytes - 1));
    }
    if requirements.max_depth > 0 {
        limited.push(exact.with_max_depth(requirements.max_depth - 1));
    }
    if requirements.allocations > 0 {
        limited.push(exact.with_allocations(requirements.allocations - 1));
    }
    if requirements.retained_bytes > 0 {
        limited.push(exact.with_retained_bytes(requirements.retained_bytes - 1));
    }
    if requirements.scratch_bytes > 0 {
        limited.push(exact.with_scratch_bytes(requirements.scratch_bytes - 1));
    }
    for limits in limited {
        let result = prepared.execute(limits);
        assert!(
            result.is_err(),
            "a max-minus-one axis-title ceiling succeeded"
        );
        if let Err(error) = result {
            observe_error(error);
        }
        assert_eq!(
            source,
            original.as_slice(),
            "failed execute modified its source"
        );
    }
}

fn exercise_matrix(data: &[u8]) {
    let title = replacement_title(data);
    let sources = [
        [
            field_varint(4000, 7),
            field_varint(13, 1),
            field_text(17, b"secondary-axis-marker"),
            field_text(15, title.as_bytes()),
            field_varint(14, 1),
            field_text(16, b"value axis"),
            field_varint(4001, 9),
        ]
        .concat(),
        [
            field_text(16, b"value first"),
            field_varint(18, 3),
            field_varint(14, 0),
            field_text(15, b"category"),
            field_varint(13, 0),
            field_text(4002, b"opaque tail"),
        ]
        .concat(),
        [
            field_varint(4003, u64::from(data.first().copied().unwrap_or(11))),
            field_varint(13, 1),
            field_varint(14, 1),
            field_text(15, b""),
            field_text(16, b""),
            field_text(4004, b"unknown between selected fields"),
        ]
        .concat(),
    ];
    for source in sources {
        exercise_source(&source, data);
    }
}

fn exercise_malformed(data: &[u8]) {
    let malformed = [
        vec![0x6a, 0x01, 0x01],       // field 13 has a length-delimited wire type
        vec![0x68, 0x01, 0x68, 0x00], // duplicate category visibility
        vec![0x70, 0x01, 0x70, 0x00], // duplicate value visibility
        vec![0x7a, 0x01, b'a', 0x7a, 0x01, b'b'], // duplicate category title
        vec![0x82, 0x01, 0x01, b'a', 0x82, 0x01, 0x01, b'b'], // duplicate value title
        vec![0x7a, 0x01, 0xff],       // invalid category UTF-8
        vec![0x82, 0x01, 0x01, 0xff], // invalid value UTF-8
        vec![0x7a, 0x80, 0x00],       // non-canonical category length
        vec![0x68, 0x80],             // truncated bool varint
        vec![0x82, 0x01, 0x80],       // truncated value-title length
        vec![0x6f],                   // unsupported category wire type
        vec![0x00],                   // field number zero
        vec![0x6b, 0x68, 0x01],       // unterminated group
        vec![0x53, 0x68, 0x01],       // unterminated unknown group
    ];
    for source in malformed {
        let before = source.clone();
        let decoded = decode_axis_titles(&source, options(&source));
        assert!(decoded.is_err(), "known malformed axis wire was accepted");
        if let Err(error) = decoded {
            observe_error(error);
        }
        assert_eq!(source, before, "malformed decode modified its source");
        let rewritten = rewrite_axis_title(
            &source,
            AxisTitleWrite::preserve().with_title(AxisTitleKind::Category, Some("mutation")),
            options(&source),
        );
        assert!(rewritten.is_err(), "malformed axis wire was rewritten");
        if let Err(error) = rewritten {
            observe_error(error);
        }
        assert_eq!(source, before, "malformed rewrite modified its source");
    }
    black_box(data);
}

fn exercise_redacted_malformed() {
    let mut source = field_text(4000, PRIVATE_SENTINEL.as_bytes());
    source.extend_from_slice(&[0x68, 0x80]);
    let error = decode_axis_titles(&source, options(&source))
        .expect_err("private malformed axis sentinel was accepted");
    observe_redacted(error, PRIVATE_SENTINEL);
    let error = rewrite_axis_title(
        &source,
        AxisTitleWrite::preserve().with_title(AxisTitleKind::Value, Some("redacted")),
        options(&source),
    )
    .expect_err("private malformed axis rewrite was accepted");
    observe_redacted(error, PRIVATE_SENTINEL);
}

fn exercise_limit_guards() {
    let source = [field_varint(13, 1), field_text(15, b"abcd")].concat();
    let field_error = decode_axis_titles(
        &source,
        DecodeOptions::new(source.len(), 1, source.len() * 8, 8),
    )
    .expect_err("axis field limit was not enforced");
    assert_eq!(field_error.field_limit_values(), Some((2, 1)));

    let work_error = decode_axis_titles(
        &source,
        DecodeOptions::new(source.len(), 8, source.len(), 8),
    )
    .expect_err("axis work limit was not enforced");
    assert_eq!(
        work_error.work_limit_values(),
        Some((source.len() * 2, source.len()))
    );

    let output_error = decode_axis_titles(
        &source,
        DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_output_bytes(3),
    )
    .expect_err("axis output limit was not enforced");
    assert_eq!(output_error.output_limit_values(), Some((4, 3)));

    let title_error = decode_axis_titles(
        &source,
        DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_title_bytes(3),
    )
    .expect_err("axis title limit was not enforced");
    assert_eq!(title_error.title_limit_values(), Some((4, 3)));

    let nesting = [0x0b, 0x13, 0x14, 0x0c];
    let nesting_error = decode_axis_titles(
        &nesting,
        DecodeOptions::new(nesting.len(), 8, nesting.len() * 8, 1),
    )
    .expect_err("axis nesting limit was not enforced");
    assert_eq!(
        nesting_error.wire_resource_limit(),
        Some(WireResourceLimit::Nesting {
            observed: 2,
            maximum: 1,
        })
    );

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
        decode_axis_titles(source, options).expect_err("oversized axis-title source was accepted");
    assert_eq!(
        error.resource_limit(),
        Some(DecodeLimit::Bytes {
            observed: OVERSIZED_INPUT_BYTES,
            maximum: MAX_INPUT_BYTES,
        })
    );
    observe_error(error);
}

fn exercise_atomic_rewrite_limit() {
    let source = [field_varint(4000, 7), field_varint(13, 1)].concat();
    let before = source.clone();
    let rewrite_options = DecodeOptions::new(source.len(), 128, source.len() * 8, 8)
        .with_max_output_bytes(source.len())
        .with_max_title_bytes(128);
    let error = rewrite_axis_title(
        &source,
        AxisTitleWrite::preserve().with_title(AxisTitleKind::Category, Some("atomic")),
        rewrite_options,
    )
    .expect_err("output-capped axis rewrite unexpectedly succeeded");
    assert!(error.output_limit_values().is_some());
    observe_error(error);
    assert_eq!(source, before, "failed axis rewrite modified its source");
    let snapshot = decode_axis_titles(&source, options(&source))
        .expect("output-capped axis rewrite made source unreadable");
    assert_eq!(snapshot.visible_title(AxisTitleKind::Category), Some(""));
}

fn assert_report(report: DecodeReport, source_bytes: usize) {
    assert_eq!(report.source_bytes(), source_bytes);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(report.output_bytes() <= MAX_OUTPUT_BYTES.max(source_bytes));
    assert!(report.title_bytes() <= MAX_TITLE_BYTES);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.retained_bytes(), source_bytes);
    assert_eq!(report.scratch_bytes(), 0);
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
        "decoded axis title was not borrowed from the source"
    );
}

fn assert_unknown_spans_preserved(source: &[u8], output: &[u8]) {
    let source_spans =
        wire_field_spans(source).expect("strict source decode produced unparseable wire spans");
    let output_spans =
        wire_field_spans(output).expect("strict output decode produced unparseable wire spans");
    let source_unknown = source_spans
        .iter()
        .into_iter()
        .filter(|(number, _)| !is_selected(*number))
        .map(|(_, raw)| (*raw).to_vec())
        .collect::<Vec<_>>();
    let output_unknown = output_spans
        .iter()
        .into_iter()
        .filter(|(number, _)| !is_selected(*number))
        .map(|(_, raw)| (*raw).to_vec())
        .collect::<Vec<_>>();
    assert_eq!(source_unknown, output_unknown, "unknown axis spans changed");
    if selected_field_presence(&source_spans) == selected_field_presence(&output_spans) {
        assert_selected_unknown_interleaving_preserved(&source_spans, &output_spans);
    }
}

fn assert_unknown_span_bytes_preserved(source: &[u8], output: &[u8]) {
    let source_unknown = wire_field_spans(source)
        .expect("strict source decode produced unparseable wire spans")
        .into_iter()
        .filter(|(number, _)| !is_selected(*number))
        .map(|(_, raw)| raw.to_vec())
        .collect::<Vec<_>>();
    let output_unknown = wire_field_spans(output)
        .expect("strict output decode produced unparseable wire spans")
        .into_iter()
        .filter(|(number, _)| !is_selected(*number))
        .map(|(_, raw)| raw.to_vec())
        .collect::<Vec<_>>();
    assert_eq!(
        source_unknown, output_unknown,
        "inverse changed unrelated axis spans"
    );
}

fn is_selected(number: u64) -> bool {
    matches!(number, 13..=16)
}

fn wire_field_spans(source: &[u8]) -> Option<Vec<(u64, &[u8])>> {
    let mut spans = Vec::new();
    let mut offset = 0;
    while offset < source.len() {
        let start = offset;
        let (tag, next) = read_varint(source, offset)?;
        offset = next;
        let number = tag >> 3;
        let wire_type = tag & 7;
        if !skip_wire(source, &mut offset, wire_type, number) {
            return None;
        }
        spans.push((number, &source[start..offset]));
    }
    Some(spans)
}

fn selected_field_presence(spans: &[(u64, &[u8])]) -> Vec<u64> {
    let mut fields = spans
        .iter()
        .filter_map(|(number, _)| is_selected(*number).then_some(*number))
        .collect::<Vec<_>>();
    fields.sort_unstable();
    fields
}

fn assert_selected_unknown_interleaving_preserved(
    source: &[(u64, &[u8])],
    output: &[(u64, &[u8])],
) {
    assert_eq!(
        source.len(),
        output.len(),
        "rewrite changed span count without changing selected presence"
    );
    for ((source_number, source_raw), (output_number, output_raw)) in
        source.iter().zip(output.iter())
    {
        assert_eq!(
            is_selected(*source_number),
            is_selected(*output_number),
            "rewrite changed selected/unknown interleaving: source={source:?} output={output:?}"
        );
        if !is_selected(*source_number) {
            assert_eq!(
                source_raw, output_raw,
                "rewrite moved or changed an unknown wire span"
            );
        }
    }
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
            let number = tag >> 3;
            let child_wire = tag & 7;
            if child_wire == 4 {
                return number == group;
            }
            if !skip_wire(source, offset, child_wire, number) {
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

fn replacement_title(data: &[u8]) -> String {
    let start = data.len().min(8);
    let end = data.len().min(start.saturating_add(MAX_TITLE_INPUT_BYTES));
    if start == end {
        return "fuzz axis title".to_owned();
    }
    String::from_utf8_lossy(&data[start..end]).into_owned()
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
