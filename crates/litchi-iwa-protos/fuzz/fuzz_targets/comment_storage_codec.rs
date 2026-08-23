#![no_main]

use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::comment_storage_codec::{
    CommentStorageVisitor, DateSnapshot, DecodeError, DecodeLimit, DecodeOptions, ReferenceRecord,
    ReferenceSnapshot, UuidSnapshot, decode_comment_storage_archive_with_report,
    decode_comment_storage_archive_with_visitor, decode_reference, visit_comment_storage_replies,
};
use litchi_iwa_protos::tsd::CommentStorageArchive;
use litchi_iwa_protos::tsp::{Reference, Uuid};
use prost::Message as _;

// Keep libFuzzer's source and the strict decoder's aggregate accounting
// finite. Oversized inputs are skipped rather than truncated, so every entry
// point sees one unchanged caller-owned source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_REFERENCES: usize = 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_RETAINED_REPLIES: usize = MAX_FIELDS / 32;
const PROBE_MAX_FIELDS: usize = 128;
const PROBE_MAX_WORK_BYTES: usize = 16 * 1024;
const PROBE_MAX_REFERENCES: usize = 16;
const PROBE_MAX_TEXT_BYTES: usize = 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct SliceSummary {
    pointer: usize,
    length: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ReferenceSummary {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ReplySummary {
    ordinal: usize,
    raw: SliceSummary,
    reference: ReferenceSummary,
}

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

    fn assert_borrowed(self, bytes: &[u8]) -> SliceSummary {
        let summary = SliceSummary {
            pointer: bytes.as_ptr() as usize,
            length: bytes.len(),
        };
        if bytes.is_empty() {
            return summary;
        }
        let payload_end = summary
            .pointer
            .checked_add(summary.length)
            .expect("bounded payload pointer range");
        assert!(
            summary.pointer >= self.start && payload_end <= self.end,
            "decoded payload did not borrow from the source"
        );
        summary
    }
}

struct ReplyCollector<'expected> {
    source: SourceRange,
    expected: Option<&'expected [Reference]>,
    replies: Vec<ReplySummary>,
    total: usize,
}

impl<'expected> ReplyCollector<'expected> {
    fn new(source: &[u8], expected: Option<&'expected [Reference]>) -> Self {
        Self {
            source: SourceRange::new(source),
            expected,
            replies: Vec::with_capacity(MAX_RETAINED_REPLIES),
            total: 0,
        }
    }

    fn observe(&mut self, reply: ReferenceRecord<'_>) {
        let raw = self.source.assert_borrowed(reply.raw());
        let standalone = decode_reference(reply.raw(), reference_options(reply.raw()));
        let standalone = standalone
            .unwrap_or_else(|error| panic!("strict reply was not standalone-decodable: {error}"));
        assert_eq!(standalone, reply.reference());
        if let Some(expected) = self.expected {
            let reference = expected
                .get(self.total)
                .unwrap_or_else(|| panic!("streamed reply count exceeded Prost reply count"));
            assert_reference_matches(reply.reference(), reference);
            let decoded = Reference::decode(reply.raw())
                .unwrap_or_else(|error| panic!("strict reply disagreed with Prost: {error}"));
            assert_eq!(&decoded, reference);
        }
        let summary = ReplySummary {
            ordinal: self.total,
            raw,
            reference: summarize_reference(reply.reference()),
        };
        self.total = self
            .total
            .checked_add(1)
            .expect("bounded reply callback count");
        if self.replies.len() < MAX_RETAINED_REPLIES {
            self.replies.push(summary);
        }
    }
}

impl CommentStorageVisitor for ReplyCollector<'_> {
    fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.observe(reply);
        Ok(())
    }
}

struct ErroringVisitor {
    source: SourceRange,
    fail_on: usize,
    error: DecodeError,
    replies: Vec<ReplySummary>,
    total: usize,
}

impl ErroringVisitor {
    fn new(source: &[u8], fail_on: usize, error: DecodeError) -> Self {
        Self {
            source: SourceRange::new(source),
            fail_on,
            error,
            replies: Vec::with_capacity(MAX_RETAINED_REPLIES),
            total: 0,
        }
    }
}

impl CommentStorageVisitor for ErroringVisitor {
    fn visit_reply(&mut self, reply: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        let raw = self.source.assert_borrowed(reply.raw());
        let ordinal = self.total;
        self.total = self.total.checked_add(1).expect("bounded callback count");
        if self.replies.len() < MAX_RETAINED_REPLIES {
            self.replies.push(ReplySummary {
                ordinal,
                raw,
                reference: summarize_reference(reply.reference()),
            });
        }
        if self.total == self.fail_on {
            return Err(self.error.clone());
        }
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise(&source);
    exercise_standalone_reference(&source);
    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(exercise_deterministic_probes);
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

fn reference_options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        1,
        MAX_TEXT_BYTES,
    )
}

fn probe_options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        PROBE_MAX_FIELDS,
        PROBE_MAX_WORK_BYTES,
        8,
        PROBE_MAX_REFERENCES,
        PROBE_MAX_TEXT_BYTES,
    )
}

fn options_with_limits(
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
    max_text_bytes: usize,
) -> DecodeOptions {
    DecodeOptions::new(
        max_message_bytes,
        max_fields,
        max_work_bytes,
        recursion_limit,
        max_references,
        max_text_bytes,
    )
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = u8::try_from(value & 0x7f).expect("varint chunk fits in a byte");
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

fn append_key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    append_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_key(output, number, 0);
    append_varint(output, value);
}

fn append_bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
    append_key(output, number, 2);
    append_varint(
        output,
        u64::try_from(value.len()).expect("bounded probe payload length fits in u64"),
    );
    output.extend_from_slice(value);
}

fn reference_wire(identifier: u64, deprecated_type: i32, deprecated_is_external: bool) -> Vec<u8> {
    let mut output = Vec::with_capacity(24);
    append_varint_field(&mut output, 1, identifier);
    append_varint_field(
        &mut output,
        2,
        u64::from_ne_bytes(i64::from(deprecated_type).to_ne_bytes()),
    );
    append_varint_field(&mut output, 3, u64::from(deprecated_is_external));
    output
}

fn standalone_reference_probe(seed: &[u8]) -> (Vec<u8>, u64, i32, bool) {
    let identifier = seed.iter().take(8).fold(0_u64, |value, byte| {
        value.wrapping_mul(257).wrapping_add(u64::from(*byte))
    });
    let deprecated_type = i32::from(seed.first().copied().unwrap_or_default()) - 128;
    let deprecated_is_external = seed.get(1).copied().is_some_and(|byte| byte & 1 != 0);
    (
        reference_wire(identifier, deprecated_type, deprecated_is_external),
        identifier,
        deprecated_type,
        deprecated_is_external,
    )
}

fn deterministic_archive_probe() -> Vec<u8> {
    let mut output = Vec::with_capacity(32);
    append_bytes_field(&mut output, 1, b"probe");
    append_bytes_field(&mut output, 4, &reference_wire(7, 0, false));
    append_bytes_field(&mut output, 4, &reference_wire(8, -3, true));
    output
}

const MALFORMED_REFERENCE_CASES: &[&[u8]] = &[
    &[],
    &[0x08],
    &[0x08, 0x80],
    &[0x08, 0x81, 0x00],
    &[0x0a, 0x01, 0x01],
    &[0x08, 0x01, 0x08, 0x02],
    &[0x1b, 0x08, 0x01],
];

const MALFORMED_ARCHIVE_CASES: &[(&[u8], usize)] = &[
    (&[0x0a, 0x01, 0xff], 0),
    (&[0x0a, 0x01, b'a', 0x0a, 0x01, b'b'], 0),
    (&[0x12, 0x00], 0),
    (&[0x1b, 0x08, 0x01], 0),
    (&[0x08, 0x01], 0),
    (&[0x0a, 0x04, b'a'], 0),
    (&[0x08, 0x81, 0x00], 0),
    (
        &[0x22, 0x02, 0x08, 0x01, 0x0a, 0x01, b'a', 0x0a, 0x01, b'b'],
        1,
    ),
];

fn assert_limit_error<T>(result: Result<T, DecodeError>, expected: DecodeLimit) {
    let error = match result {
        Ok(_) => panic!("deterministic limit probe unexpectedly succeeded"),
        Err(error) => error,
    };
    assert_eq!(error.resource_limit(), Some(expected));
}

fn malformed_reference_error() -> DecodeError {
    decode_reference(&[0x08], reference_options(&[0x08]))
        .expect_err("truncated standalone reference probe unexpectedly succeeded")
}

fn exercise_standalone_reference(seed: &[u8]) {
    let (reference, identifier, deprecated_type, deprecated_is_external) =
        standalone_reference_probe(seed);
    let before = reference.clone();
    let decoded = decode_reference(&reference, reference_options(&reference))
        .expect("canonical standalone reference probe was rejected");
    assert_eq!(decoded.identifier(), identifier);
    assert_eq!(decoded.deprecated_type(), Some(deprecated_type));
    assert_eq!(
        decoded.deprecated_is_external(),
        Some(deprecated_is_external)
    );
    assert_eq!(
        reference,
        before.as_slice(),
        "reference decode modified source"
    );

    let bytes_limit = options_with_limits(
        reference.len() - 1,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        1,
        MAX_TEXT_BYTES,
    );
    assert_limit_error(
        decode_reference(&reference, bytes_limit),
        DecodeLimit::Bytes {
            observed: reference.len(),
            maximum: reference.len() - 1,
        },
    );
    assert_eq!(
        reference,
        before.as_slice(),
        "byte-limit probe modified source"
    );

    let fields_limit = options_with_limits(
        reference.len(),
        0,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        1,
        MAX_TEXT_BYTES,
    );
    assert_limit_error(
        decode_reference(&reference, fields_limit),
        DecodeLimit::Fields {
            observed: 1,
            maximum: 0,
        },
    );

    let work_limit = options_with_limits(
        reference.len(),
        MAX_FIELDS,
        0,
        MAX_RECURSION,
        1,
        MAX_TEXT_BYTES,
    );
    assert_limit_error(
        decode_reference(&reference, work_limit),
        DecodeLimit::Work {
            observed: reference.len(),
            maximum: 0,
        },
    );

    let references_limit = options_with_limits(
        reference.len(),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        0,
        MAX_TEXT_BYTES,
    );
    assert_limit_error(
        decode_reference(&reference, references_limit),
        DecodeLimit::References {
            observed: 1,
            maximum: 0,
        },
    );

    let recursion_limit = options_with_limits(
        reference.len(),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        0,
        1,
        MAX_TEXT_BYTES,
    );
    assert_limit_error(
        decode_reference(&reference, recursion_limit),
        DecodeLimit::Nesting {
            observed: 0,
            maximum: MAX_RECURSION,
        },
    );
    assert_eq!(
        reference,
        before.as_slice(),
        "standalone limit probes modified source"
    );

    for malformed in MALFORMED_REFERENCE_CASES {
        let before = (*malformed).to_vec();
        assert!(
            decode_reference(malformed, reference_options(malformed)).is_err(),
            "malformed standalone reference probe unexpectedly succeeded"
        );
        assert_eq!(
            *malformed,
            before.as_slice(),
            "malformed probe modified source"
        );
    }
}

fn exercise_malformed_archives() {
    for &(source, expected_callbacks) in MALFORMED_ARCHIVE_CASES {
        let before = source.to_vec();
        assert!(
            decode_comment_storage_archive_with_report(source, options()).is_err(),
            "malformed archive probe unexpectedly succeeded on report path"
        );
        assert_eq!(source, before.as_slice(), "report probe modified source");

        let mut visitor = ReplyCollector::new(source, None);
        assert!(
            decode_comment_storage_archive_with_visitor(source, options(), &mut visitor).is_err(),
            "malformed archive probe unexpectedly succeeded on visitor path"
        );
        assert_eq!(source, before.as_slice(), "visitor probe modified source");
        assert_eq!(visitor.total, expected_callbacks);
        assert_reply_summaries(&visitor.replies);
    }
}

fn assert_archive_limit(source: &[u8], decode_options: DecodeOptions, expected: DecodeLimit) {
    let before = source.to_vec();
    assert_limit_error(
        decode_comment_storage_archive_with_report(source, decode_options),
        expected,
    );
    assert_eq!(
        source,
        before.as_slice(),
        "report limit probe modified source"
    );

    let mut visitor = ReplyCollector::new(source, None);
    assert_limit_error(
        decode_comment_storage_archive_with_visitor(source, decode_options, &mut visitor),
        expected,
    );
    assert_eq!(
        source,
        before.as_slice(),
        "visitor limit probe modified source"
    );
    assert_reply_summaries(&visitor.replies);
}

fn exercise_archive_limits(
    source: &[u8],
    report: litchi_iwa_protos::comment_storage_codec::DecodeReport,
) {
    assert_archive_limit(
        source,
        options_with_limits(
            source.len() - 1,
            PROBE_MAX_FIELDS,
            PROBE_MAX_WORK_BYTES,
            8,
            PROBE_MAX_REFERENCES,
            PROBE_MAX_TEXT_BYTES,
        ),
        DecodeLimit::Bytes {
            observed: source.len(),
            maximum: source.len() - 1,
        },
    );
    assert_archive_limit(
        source,
        options_with_limits(
            source.len(),
            report.fields() - 1,
            PROBE_MAX_WORK_BYTES,
            8,
            PROBE_MAX_REFERENCES,
            PROBE_MAX_TEXT_BYTES,
        ),
        DecodeLimit::Fields {
            observed: report.fields(),
            maximum: report.fields() - 1,
        },
    );
    assert_archive_limit(
        source,
        options_with_limits(
            source.len(),
            PROBE_MAX_FIELDS,
            report.work_bytes() - 1,
            8,
            PROBE_MAX_REFERENCES,
            PROBE_MAX_TEXT_BYTES,
        ),
        DecodeLimit::Work {
            observed: report.work_bytes(),
            maximum: report.work_bytes() - 1,
        },
    );
    assert_archive_limit(
        source,
        options_with_limits(
            source.len(),
            PROBE_MAX_FIELDS,
            PROBE_MAX_WORK_BYTES,
            8,
            report.references() - 1,
            PROBE_MAX_TEXT_BYTES,
        ),
        DecodeLimit::References {
            observed: report.references(),
            maximum: report.references() - 1,
        },
    );
    assert_archive_limit(
        source,
        options_with_limits(
            source.len(),
            PROBE_MAX_FIELDS,
            PROBE_MAX_WORK_BYTES,
            8,
            PROBE_MAX_REFERENCES,
            report.text_bytes() - 1,
        ),
        DecodeLimit::Text {
            observed: report.text_bytes(),
            maximum: report.text_bytes() - 1,
        },
    );
    assert_archive_limit(
        source,
        options_with_limits(
            source.len(),
            PROBE_MAX_FIELDS,
            PROBE_MAX_WORK_BYTES,
            1,
            PROBE_MAX_REFERENCES,
            PROBE_MAX_TEXT_BYTES,
        ),
        DecodeLimit::Nesting {
            observed: 2,
            maximum: 1,
        },
    );
}

fn exercise_callback_errors(source: &[u8]) {
    let before = source.to_vec();
    let callback_error = malformed_reference_error();
    let mut visitor = ErroringVisitor::new(source, 2, callback_error.clone());
    assert!(
        decode_comment_storage_archive_with_visitor(source, probe_options(source), &mut visitor)
            .is_err(),
        "failing visitor probe unexpectedly succeeded"
    );
    assert_eq!(source, before.as_slice(), "failing visitor modified source");
    assert_eq!(visitor.total, 2);
    assert_reply_summaries(&visitor.replies);
    assert_eq!(
        visitor
            .replies
            .iter()
            .map(|reply| reply.reference.identifier)
            .collect::<Vec<_>>(),
        [7, 8]
    );

    let source_range = SourceRange::new(source);
    let mut replies = Vec::with_capacity(MAX_RETAINED_REPLIES);
    let mut total = 0usize;
    let result = visit_comment_storage_replies(source, probe_options(source), &mut |reply| {
        let raw = source_range.assert_borrowed(reply.raw());
        let ordinal = total;
        total = total.checked_add(1).expect("bounded callback count");
        if replies.len() < MAX_RETAINED_REPLIES {
            replies.push(ReplySummary {
                ordinal,
                raw,
                reference: summarize_reference(reply.reference()),
            });
        }
        if total == 2 {
            Err(callback_error.clone())
        } else {
            Ok(())
        }
    });
    assert!(
        result.is_err(),
        "failing closure probe unexpectedly succeeded"
    );
    assert_eq!(source, before.as_slice(), "failing closure modified source");
    assert_eq!(total, 2);
    assert_reply_summaries(&replies);
}

fn exercise_deterministic_probes() {
    exercise_standalone_reference(b"comment-storage-probe");
    exercise_malformed_archives();

    let source = deterministic_archive_probe();
    let before = source.clone();
    let (snapshot, report) =
        decode_comment_storage_archive_with_report(&source, probe_options(&source))
            .expect("deterministic archive probe was rejected on report path");
    assert_eq!(source, before.as_slice(), "probe report modified source");
    assert_eq!(snapshot.text(), Some("probe"));
    assert_eq!(report.source_bytes(), source.len());
    assert_eq!(report.references(), 2);
    assert_eq!(report.replies(), 2);
    assert_eq!(report.reply_references(), 2);
    assert_eq!(report.text_bytes(), 5);

    let mut visitor = ReplyCollector::new(&source, None);
    let (streamed_snapshot, streamed_report) =
        decode_comment_storage_archive_with_visitor(&source, probe_options(&source), &mut visitor)
            .expect("deterministic archive probe was rejected on visitor path");
    assert_eq!(source, before.as_slice(), "probe visitor modified source");
    assert_eq!(streamed_snapshot, snapshot);
    assert_eq!(streamed_report, report);
    assert_eq!(visitor.total, 2);
    assert_reply_summaries(&visitor.replies);

    let mut closure_visitor = ReplyCollector::new(&source, None);
    let (closure_snapshot, closure_report) =
        visit_comment_storage_replies(&source, probe_options(&source), &mut |reply| {
            closure_visitor.observe(reply);
            Ok(())
        })
        .expect("deterministic archive probe was rejected on closure path");
    assert_eq!(source, before.as_slice(), "probe closure modified source");
    assert_eq!(closure_snapshot, snapshot);
    assert_eq!(closure_report, report);
    assert_eq!(closure_visitor.total, 2);
    assert_reply_summaries(&closure_visitor.replies);

    exercise_archive_limits(&source, report);
    exercise_callback_errors(&source);
}

fn exercise(source: &[u8]) {
    let before = source.to_vec();
    let strict = decode_comment_storage_archive_with_report(source, options());
    assert_eq!(source, before.as_slice(), "scalar decode modified source");

    // Prost is deliberately an oracle only after strict acceptance. Its
    // proto2 decoder does not enforce the strict codec's required-field and
    // canonical-wire policy.
    let prost = strict.as_ref().ok().map(|_| {
        CommentStorageArchive::decode(source)
            .unwrap_or_else(|error| panic!("strict acceptance disagreed with Prost: {error}"))
    });
    assert_eq!(source, before.as_slice(), "Prost decode modified source");

    let expected_replies = prost.as_ref().map(|archive| archive.replies.as_slice());
    let mut collector = ReplyCollector::new(source, expected_replies);
    let streamed = decode_comment_storage_archive_with_visitor(source, options(), &mut collector);
    assert_eq!(source, before.as_slice(), "visitor decode modified source");

    match (strict, streamed) {
        (Ok((strict_snapshot, strict_report)), Ok((streamed_snapshot, streamed_report))) => {
            assert_eq!(streamed_snapshot, strict_snapshot);
            assert_eq!(streamed_report, strict_report);
            let archive = prost
                .as_ref()
                .expect("strict success must have a Prost oracle");
            assert_snapshot_matches(strict_snapshot, archive);
            SourceRange::new(source)
                .assert_borrowed(strict_snapshot.text().unwrap_or_default().as_bytes());
            assert_eq!(collector.total, archive.replies.len());
            assert_reply_summaries(&collector.replies);
        },
        (Err(_), Err(_)) => {
            // A visitor may have observed a valid reply prefix before a later
            // malformed field. No partial scalar result is published; source
            // and callback borrow checks still apply above.
        },
        _ => panic!("scalar and visitor paths disagreed on success"),
    }
}

fn summarize_reference(reference: ReferenceSnapshot) -> ReferenceSummary {
    ReferenceSummary {
        identifier: reference.identifier(),
        deprecated_type: reference.deprecated_type(),
        deprecated_is_external: reference.deprecated_is_external(),
    }
}

fn assert_snapshot_matches(
    strict: litchi_iwa_protos::comment_storage_codec::CommentStorageSnapshot<'_>,
    prost: &CommentStorageArchive,
) {
    assert_eq!(strict.text(), prost.text.as_deref());
    assert_eq!(
        strict.creation_date().map(DateSnapshot::seconds_bits),
        prost
            .creation_date
            .as_ref()
            .map(|date| date.seconds.to_bits())
    );
    assert_reference_matches_optional(strict.author(), prost.author.as_ref());
    assert_uuid_matches_optional(strict.storage_uuid(), prost.storage_uuid.as_ref());
}

fn assert_reference_matches_optional(strict: Option<ReferenceSnapshot>, prost: Option<&Reference>) {
    match (strict, prost) {
        (Some(strict), Some(prost)) => assert_reference_matches(strict, prost),
        (None, None) => {},
        _ => panic!("reference presence differed between strict and Prost paths"),
    }
}

fn assert_reference_matches(strict: ReferenceSnapshot, prost: &Reference) {
    assert_eq!(strict.identifier(), prost.identifier);
    assert_eq!(strict.deprecated_type(), prost.deprecated_type);
    assert_eq!(
        strict.deprecated_is_external(),
        prost.deprecated_is_external
    );
}

fn assert_uuid_matches_optional(strict: Option<UuidSnapshot>, prost: Option<&Uuid>) {
    match (strict, prost) {
        (Some(strict), Some(prost)) => {
            assert_eq!(strict.lower(), prost.lower);
            assert_eq!(strict.upper(), prost.upper);
        },
        (None, None) => {},
        _ => panic!("UUID presence differed between strict and Prost paths"),
    }
}

fn assert_reply_summaries(summaries: &[ReplySummary]) {
    for (index, summary) in summaries.iter().enumerate() {
        assert_eq!(summary.ordinal, index);
        assert_slice_summary(summary.raw);
        std::hint::black_box(summary.reference);
    }
}

fn assert_slice_summary(summary: SliceSummary) {
    assert!(summary.pointer != 0 || summary.length == 0);
}
