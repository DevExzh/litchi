#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::comment_storage_codec::{
    CommentStorageVisitor, DateSnapshot, DecodeError, DecodeOptions, ReferenceRecord,
    ReferenceSnapshot, UuidSnapshot, decode_comment_storage_archive_with_report,
    decode_comment_storage_archive_with_visitor,
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

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise(&source);
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
