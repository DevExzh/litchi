#![no_main]

//! Bounded strict-wire fuzzing for `TST.PopUpMenuModel` and popup cell specs.
//!
//! The target keeps every decode on caller-owned bytes, compares scalar,
//! reported, and streaming visitor routes, and probes each finite limit
//! independently.  Once the prepared rewrite API is available, the same
//! source harness will execute its exact requirements and candidate readback;
//! the decode/visitor matrix remains useful for malformed and unknown-wire
//! coverage independently of package lifecycle fuzzing.

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_pop_up_menu_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_REFERENCES: usize = 1_024;
const MAX_ITEMS: usize = 1_024;
const MAX_TEXT_BYTES: usize = 64 * 1024;

const FIXED_CASES: &[&[u8]] = &[
    // NIL sentinel followed by the string "A" with an empty format archive.
    &[
        0x12, 0x02, 0x08, 0x01, 0x12, 0x12, 0x08, 0x05, 0x2a, 0x0e, 0x0a, 0x01, 0x41, 0x12, 0x03,
        0x08, 0xac, 0x02, 0x20, 0x00, 0x28, 0x00, 0x30, 0x00,
    ],
    // NIL plus two choices.
    &[
        0x12, 0x02, 0x08, 0x01, 0x12, 0x12, 0x08, 0x05, 0x2a, 0x0e, 0x0a, 0x01, 0x41, 0x12, 0x03,
        0x08, 0xac, 0x02, 0x20, 0x00, 0x28, 0x00, 0x30, 0x00, 0x12, 0x18, 0x08, 0x05, 0x2a, 0x16,
        0x0a, 0x05, 0x4c, 0x6f, 0x77, 0x65, 0x72, 0x12, 0x03, 0x08, 0xac, 0x02, 0x20, 0x00, 0x28,
        0x00, 0x30, 0x00,
    ],
    // Deprecated PopUpMenuModel.item must be rejected by the strict route.
    &[0x0a, 0x02, 0x08, 0x01],
    // Missing the required first NIL sentinel.
    &[
        0x12, 0x12, 0x08, 0x05, 0x2a, 0x0e, 0x0a, 0x01, 0x41, 0x12, 0x03, 0x08, 0xac, 0x02, 0x20,
        0x00, 0x28, 0x00, 0x30, 0x00,
    ],
    // Duplicate NIL sentinel.
    &[0x12, 0x02, 0x08, 0x01, 0x12, 0x02, 0x08, 0x01],
    // A non-string CellValueArchive type with a string payload.
    &[
        0x12, 0x09, 0x08, 0x01, 0x2a, 0x05, 0x0a, 0x01, 0x41, 0x12, 0x00,
    ],
    // Known string options that are forbidden by the popup policy.
    &[
        0x12, 0x11, 0x08, 0x05, 0x2a, 0x0d, 0x0a, 0x01, 0x41, 0x12, 0x00, 0x18, 0x01, 0x20, 0x01,
        0x28, 0x01, 0x30, 0x01,
    ],
    // Duplicate known nested fields.
    &[
        0x12, 0x11, 0x08, 0x05, 0x2a, 0x0d, 0x0a, 0x01, 0x41, 0x12, 0x00, 0x18, 0x00, 0x18, 0x00,
        0x28, 0x00,
    ],
    // Unknown overlong scalar and a balanced unknown group around a valid item.
    &[
        0x98, 0x03, 0x81, 0x80, 0x00, 0x93, 0x03, 0xa0, 0x03, 0x81, 0x00, 0x94, 0x03, 0x12, 0x02,
        0x08, 0x01,
    ],
    // Truncated length-delimited field and unterminated group.
    &[0x12, 0x09, 0x08],
    &[0x93, 0x03, 0x08, 0x01],
];

#[derive(Clone, Copy)]
struct SourceRange {
    start: usize,
    end: usize,
}

impl SourceRange {
    fn new(source: &[u8]) -> Self {
        let start = source.as_ptr() as usize;
        Self {
            start,
            end: start
                .checked_add(source.len())
                .expect("bounded popup source pointer range"),
        }
    }

    fn assert_borrowed(self, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start
            .checked_add(bytes.len())
            .expect("bounded popup borrowed range");
        assert!(
            start >= self.start && end <= self.end,
            "popup visitor value did not borrow from source"
        );
    }
}

#[derive(Default)]
struct ItemVisitor {
    source: Option<SourceRange>,
    count: usize,
    text_bytes: usize,
}

impl codec::PopUpMenuVisitor for ItemVisitor {
    fn visit_item(&mut self, item: codec::PopUpMenuItem<'_>) -> Result<(), codec::DecodeError> {
        if let Some(source) = self.source {
            source.assert_borrowed(item.value().as_bytes());
        }
        self.count = self.count.checked_add(1).expect("bounded popup item count");
        self.text_bytes = self
            .text_bytes
            .checked_add(item.value().len())
            .expect("bounded popup text bytes");
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source);

    // Keep deterministic malformed and boundary cases hot in every campaign.
    for fixed in FIXED_CASES {
        exercise_source(fixed);
    }
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
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    )
}

fn exercise_source(source: &[u8]) {
    let before = source.to_vec();
    let source_range = SourceRange::new(source);
    let decode_options = options(source);

    let scalar = codec::decode_popup_menu_model(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "popup scalar decode modified source"
    );
    let reported = codec::decode_popup_menu_model_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "popup report decode modified source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert!(snapshot.has_nil_sentinel());
            assert!(snapshot.item_count() <= MAX_ITEMS);
            assert!(snapshot.text_bytes() <= MAX_TEXT_BYTES);
            assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source.len()));
            assert!(report.output_bytes() <= MAX_OUTPUT_BYTES.max(source.len()));
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            assert!(report.references() <= MAX_REFERENCES);
            assert!(report.items() <= MAX_ITEMS);
            assert!(report.text_bytes() <= MAX_TEXT_BYTES);
            black_box((snapshot.raw(), snapshot.items(), report));

            let mut visitor = ItemVisitor {
                source: Some(source_range),
                ..ItemVisitor::default()
            };
            let visitor_report =
                codec::decode_popup_menu_model_with_visitor(source, decode_options, &mut visitor);
            assert_eq!(source, before.as_slice(), "popup visitor modified source");
            let visitor_report = visitor_report
                .unwrap_or_else(|error| panic!("popup scalar/visitor routes diverged: {error}"));
            assert_eq!(visitor.count, snapshot.item_count().saturating_sub(1));
            assert_eq!(visitor.text_bytes, snapshot.text_bytes());
            assert_eq!(visitor_report, report);
            black_box(visitor_report);
            exercise_prepared_writes();
        },
        (Err(scalar_error), Err(reported_error)) => {
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "popup scalar/report disagreement: scalar={:?}, report={:?}",
            scalar_result.map(|snapshot| snapshot.item_count()),
            reported_result.map(|(snapshot, _)| snapshot.item_count())
        ),
    }

    exercise_cell_spec(source, &before);
    exercise_limit_profiles(source);
    assert_eq!(source, before.as_slice(), "popup probes modified source");
}

fn exercise_prepared_writes() {
    let items = ["Low", "Medium", "High"];
    let write_options = codec::DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    );
    let Ok(plan) = codec::prepare_popup_menu_model_write(&items, write_options) else {
        return;
    };
    let requirements = plan.execution_requirements();
    let prepare_report = plan.prepare_report();
    assert_eq!(prepare_report.output_bytes(), requirements.output_bytes());
    assert_eq!(prepare_report.fields(), requirements.fields());
    assert_eq!(prepare_report.work_bytes(), requirements.work_bytes());
    assert_eq!(prepare_report.max_depth(), requirements.max_depth());
    assert_eq!(prepare_report.references(), requirements.references());
    assert_eq!(prepare_report.items(), requirements.items());
    assert_eq!(prepare_report.text_bytes(), requirements.text_bytes());

    let output = plan
        .execute(codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact popup write execution failed: {error}"));
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    assert_eq!(output.report().output_bytes(), output.bytes().len());
    let output_bytes = output.bytes().to_vec();
    let readback = codec::decode_popup_menu_model(
        output.bytes(),
        codec::DecodeOptions::new(
            output.bytes().len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        ),
    )
    .unwrap_or_else(|error| panic!("popup write readback failed: {error}"));
    assert_eq!(readback.item_count(), items.len());

    if requirements.output_bytes() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_output_bytes(requirements.output_bytes() - 1)
            )
            .is_err()
        );
    }
    if requirements.fields() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_fields(requirements.fields() - 1)
            )
            .is_err()
        );
    }
    if requirements.work_bytes() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_work_bytes(requirements.work_bytes() - 1)
            )
            .is_err()
        );
    }
    if requirements.max_depth() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_max_depth(requirements.max_depth() - 1)
            )
            .is_err()
        );
    }
    if requirements.items() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_items(requirements.items() - 1)
            )
            .is_err()
        );
    }
    if requirements.text_bytes() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_text_bytes(requirements.text_bytes() - 1)
            )
            .is_err()
        );
    }
    if requirements.allocations() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_allocations(requirements.allocations() - 1)
            )
            .is_err()
        );
    }
    if requirements.retained_bytes() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_retained_bytes(requirements.retained_bytes() - 1)
            )
            .is_err()
        );
    }
    if requirements.scratch_bytes() > 0 {
        assert!(
            plan.execute(
                codec::RewriteExecutionLimits::exact(requirements)
                    .with_scratch_bytes(requirements.scratch_bytes() - 1)
            )
            .is_err()
        );
    }

    let canonical = codec::canonical_popup_menu_model(&items, write_options)
        .unwrap_or_else(|error| panic!("canonical popup write failed: {error}"));
    assert_eq!(canonical.bytes(), output_bytes);
    black_box((canonical, output));

    let Ok(spec_plan) = codec::prepare_cell_spec_write(0x7f, true, write_options) else {
        return;
    };
    let spec_requirements = spec_plan.execution_requirements();
    let spec_output = spec_plan
        .execute(codec::RewriteExecutionLimits::exact(spec_requirements))
        .unwrap_or_else(|error| panic!("exact popup cell-spec write failed: {error}"));
    let spec = codec::decode_cell_spec(
        spec_output.bytes(),
        codec::DecodeOptions::new(
            spec_output.bytes().len(),
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        ),
    )
    .unwrap_or_else(|error| panic!("popup cell-spec readback failed: {error}"));
    assert_eq!(spec.interaction_type(), 7);
    assert_eq!(spec.popup_model().identifier(), 0x7f);
    black_box(spec_output);
}

fn exercise_cell_spec(source: &[u8], before: &[u8]) {
    match codec::decode_cell_spec_with_report(source, options(source)) {
        Ok((snapshot, report)) => {
            assert_eq!(source, before, "cell-spec decode modified source");
            assert_eq!(snapshot.raw(), source);
            assert_eq!(snapshot.interaction_type(), 7);
            assert!(snapshot.popup_model().identifier() != 0);
            black_box((snapshot.starts_with_first(), report));
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_limit_profiles(source: &[u8]) {
    if source.is_empty() {
        return;
    }
    let too_small = codec::DecodeOptions::new(
        source.len() - 1,
        MAX_OUTPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_ITEMS,
        MAX_TEXT_BYTES,
    );
    observe_result(codec::decode_popup_menu_model(source, too_small));
    for (fields, work, depth, references, items, text) in [
        (
            0,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        ),
        (
            MAX_FIELDS,
            0,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        ),
        (
            MAX_FIELDS,
            MAX_WORK_BYTES,
            0,
            MAX_REFERENCES,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        ),
        (
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            0,
            MAX_ITEMS,
            MAX_TEXT_BYTES,
        ),
        (
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            0,
            MAX_TEXT_BYTES,
        ),
        (
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_ITEMS,
            0,
        ),
    ] {
        let constrained = codec::DecodeOptions::new(
            MAX_INPUT_BYTES,
            MAX_OUTPUT_BYTES,
            fields,
            work,
            depth,
            references,
            items,
            text,
        );
        observe_result(codec::decode_popup_menu_model(source, constrained));
        assert!(
            source.len() <= MAX_INPUT_BYTES,
            "source normalization failed"
        );
    }
}

fn observe_result<T>(result: Result<T, codec::DecodeError>)
where
    T: std::fmt::Debug,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => observe_error(error),
    }
}

fn observe_error(error: codec::DecodeError) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
    black_box(error.resource_limit());
    black_box(error.allocation_requested());
}
