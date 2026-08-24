#![no_main]

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::table_dimension_codec::{
    DecodeError, DecodeLimit, DecodeOptions, HeaderRecord, HeaderSizeEdit, StorageVisitor,
    decode_header_storage_bucket, decode_header_storage_bucket_with_report,
    decode_header_storage_bucket_with_visitor, execute_header_storage_bucket_size_plan,
    plan_header_storage_bucket_sizes, rewrite_header_storage_bucket_sizes,
};

// Inputs are skipped rather than truncated so every strict pass receives one
// unchanged caller-owned source. The target is intentionally bounded even
// when it is run without cargo-fuzz's `-max_len` flag.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_REFERENCES: usize = 1_024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;

const FIXED_CASES: &[&[u8]] = &[
    // HeaderStorageBucket { bucket_hash_function: 7 }.
    &[0x08, 0x07],
    // One complete Header { index: 0, size: 32.0, hiding_state: 0,
    // number_of_cells: 0 }.
    &[
        0x08, 0x07, 0x12, 0x0b, 0x08, 0x00, 0x15, 0x00, 0x00, 0x00, 0x42, 0x18, 0x00, 0x20, 0x00,
    ],
    // The same valid bucket with an unknown overlong scalar and balanced
    // unknown group at both root and selected-header levels.
    &[
        0x08, 0x07, 0x98, 0x06, 0x81, 0x00, 0x93, 0x03, 0xa0, 0x06, 0x81, 0x00, 0x94, 0x03, 0x12,
        0x17, 0x08, 0x00, 0x15, 0x00, 0x00, 0x00, 0x42, 0x18, 0x00, 0x20, 0x00, 0x98, 0x06, 0x81,
        0x00, 0x93, 0x03, 0xa0, 0x06, 0x81, 0x00, 0x94, 0x03,
    ],
    // Duplicate required root field.
    &[0x08, 0x07, 0x08, 0x08],
    // Wrong wire type for required root field.
    &[0x0d, 0x07, 0x00, 0x00, 0x00, 0x00],
    // Non-canonical required root value.
    &[0x08, 0x87, 0x00],
    // Header with a duplicate selected index.
    &[
        0x08, 0x07, 0x12, 0x0d, 0x08, 0x00, 0x08, 0x01, 0x15, 0x00, 0x00, 0x00, 0x42, 0x18, 0x00,
        0x20, 0x00,
    ],
    // Header with a wrong wire type for the fixed32 size field.
    &[
        0x08, 0x07, 0x12, 0x08, 0x08, 0x00, 0x10, 0x01, 0x18, 0x00, 0x20, 0x00,
    ],
    // Header with a truncated length-delimited reference.
    &[0x08, 0x07, 0x12, 0x03, 0x08, 0x00, 0x2a],
    // Unterminated unknown group.
    &[0x08, 0x07, 0x93, 0x03, 0x08, 0x01],
    // Unknown group closed with a different field number.
    &[0x08, 0x07, 0x93, 0x03, 0x08, 0x01, 0x9c, 0x03],
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
            .expect("bounded borrowed payload range");
        assert!(
            start >= self.start && end <= self.end,
            "header payload did not borrow from the source"
        );
    }
}

#[derive(Default)]
struct HeaderCollector {
    source: Option<SourceRange>,
    count: usize,
}

impl StorageVisitor for HeaderCollector {
    fn visit_header_record(&mut self, record: HeaderRecord<'_>) -> Result<(), DecodeError> {
        if let Some(source) = self.source {
            source.assert_borrowed(record.raw());
        }
        self.count = self.count.checked_add(1).expect("bounded header count");
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    // Keep deterministic malformed and resource-boundary cases in every
    // campaign without mixing them into any existing target's corpus.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-table-dimension");
        }
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

fn options() -> DecodeOptions {
    DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_TEXT_BYTES,
    )
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let source_range = SourceRange::new(source);

    let scalar = decode_header_storage_bucket(source, options());
    assert_eq!(source, before.as_slice(), "scalar decode modified source");

    let aggregate = decode_header_storage_bucket_with_report(source, options());
    assert_eq!(source, before.as_slice(), "reported decode modified source");

    match (scalar, aggregate) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert!(report.source_bytes() <= MAX_INPUT_BYTES);
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            black_box((snapshot.bucket_hash_function(), report));

            let mut collector = HeaderCollector {
                source: Some(source_range),
                count: 0,
            };
            let visited =
                decode_header_storage_bucket_with_visitor(source, options(), &mut collector);
            assert_eq!(source, before.as_slice(), "visitor decode modified source");
            assert!(visited.is_ok(), "visitor and scalar routes diverged");
            assert!(collector.count <= MAX_FIELDS);

            exercise_rewrites(source, data, &before);
            exercise_limits(source, &before);
        },
        (Err(scalar_error), Err(reported_error)) => {
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => {
            panic!(
                "scalar/report decode disagreement: scalar={:?}, report={:?}",
                scalar_result.map(|snapshot| snapshot.bucket_hash_function()),
                reported_result.map(|(snapshot, _)| snapshot.bucket_hash_function())
            );
        },
    }
}

fn exercise_rewrites(source: &[u8], data: &[u8], before: &[u8]) {
    let index = u32::from(data.first().copied().unwrap_or_default()) % 32;
    let replacement = match data.get(1).copied().unwrap_or_default() % 3 {
        0 => 0.0f32.to_bits(),
        1 => 32.0f32.to_bits(),
        _ => f32::from(data.get(2).copied().unwrap_or(1)).to_bits(),
    };
    let generous = options();

    let no_op = rewrite_header_storage_bucket_sizes(source, u32::MAX, &[], generous);
    assert_eq!(source, before, "no-op rewrite modified source");
    if let Ok((candidate, report)) = no_op {
        assert_eq!(candidate, source, "no-op rewrite changed bytes");
        assert_eq!(report.output_bytes(), source.len());
        black_box(report);
    }

    for edit in [
        HeaderSizeEdit::set(index, replacement),
        HeaderSizeEdit::remove(index),
    ] {
        let result = rewrite_header_storage_bucket_sizes(source, u32::MAX, &[edit], generous);
        assert_eq!(source, before, "rewrite modified source");
        let Ok((candidate, report)) = result else {
            continue;
        };
        assert!(candidate.len() <= MAX_INPUT_BYTES);
        assert_eq!(report.output_bytes(), candidate.len());
        assert!(report.source().fields() <= MAX_FIELDS);
        assert!(report.result().fields() <= MAX_FIELDS);
        assert!(report.source().work_bytes() <= MAX_WORK_BYTES);
        assert!(report.result().work_bytes() <= MAX_WORK_BYTES);
        black_box((candidate, report));
    }

    let edits = [HeaderSizeEdit::set(index, replacement)];
    let Ok(plan) = plan_header_storage_bucket_sizes(source, u32::MAX, &edits, generous) else {
        return;
    };
    let requirements = plan.requirements();
    assert!(requirements.output_bytes() <= MAX_INPUT_BYTES);
    let exact = execute_header_storage_bucket_size_plan(plan, generous);
    assert_eq!(source, before, "plan execution modified source");
    let Ok((candidate, report)) = exact else {
        return;
    };
    assert_eq!(candidate.len(), requirements.output_bytes());
    assert_eq!(report.output_bytes(), requirements.output_bytes());
    black_box((candidate, report));

    // Re-plan because execution consumes the borrowed plan. An output ceiling
    // one byte below the preflight requirement must fail before allocation.
    if requirements.output_bytes() > 0 {
        let Ok(plan) = plan_header_storage_bucket_sizes(source, u32::MAX, &edits, generous) else {
            return;
        };
        let too_small = DecodeOptions::new(
            requirements.output_bytes() - 1,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_TEXT_BYTES,
        );
        let result = execute_header_storage_bucket_size_plan(plan, too_small);
        assert!(result.is_err());
        assert_eq!(source, before, "failed execute modified source");
    }
}

fn exercise_limits(source: &[u8], before: &[u8]) {
    if !source.is_empty() {
        let input_limited = DecodeOptions::new(
            source.len() - 1,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
            MAX_TEXT_BYTES,
        );
        let _ = decode_header_storage_bucket(source, input_limited);
        assert_eq!(source, before, "input-limit decode modified source");
    }

    for (fields, work, recursion) in [
        (0, MAX_WORK_BYTES, MAX_RECURSION),
        (MAX_FIELDS, 0, MAX_RECURSION),
        (MAX_FIELDS, MAX_WORK_BYTES, 0),
    ] {
        let constrained = DecodeOptions::new(
            MAX_INPUT_BYTES,
            fields,
            work,
            recursion,
            MAX_REFERENCES,
            MAX_TEXT_BYTES,
        );
        let _ = decode_header_storage_bucket(source, constrained);
        assert_eq!(source, before, "limited decode modified source");
    }
}

fn observe_error(error: DecodeError) {
    if let Some(limit) = error.resource_limit() {
        match limit {
            DecodeLimit::Bytes { observed, maximum }
            | DecodeLimit::Fields { observed, maximum }
            | DecodeLimit::Work { observed, maximum }
            | DecodeLimit::References { observed, maximum }
            | DecodeLimit::Text { observed, maximum } => {
                assert!(observed > maximum || observed == 0);
                black_box((observed, maximum));
            },
            DecodeLimit::Nesting { observed, maximum } => {
                assert!(observed > maximum || observed == 0);
                black_box((observed, maximum));
            },
            DecodeLimit::Allocation { requested } => {
                black_box(requested);
            },
            _ => {},
        }
    }
    black_box(error);
}
