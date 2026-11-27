#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    DecodeError, DecodeOptions, ReferenceRecord, ReferenceSnapshot, StorageVisitor,
    TableDataListEntrySnapshot, TableDataListSegmentSnapshot,
    decode_table_data_list_segment_with_report, decode_table_data_list_segment_with_visitor,
    decode_table_data_list_with_report, decode_table_data_list_with_visitor,
};
use litchi_iwa_protos::tsp::Reference;
use litchi_iwa_protos::tst::{TableDataList, TableDataListSegment};
use prost::Message as _;

// Keep both libFuzzer's input and the strict decoder's aggregate accounting
// finite. Inputs are skipped rather than truncated so both entry points always
// receive one unchanged source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_REFERENCES: usize = 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;
// Prefixes are enough to make callback state inspectable without allocating a
// vector proportional to a potentially large repeated field. Borrow checks
// and source-order checks still run for every callback.
const MAX_RETAINED_RECORDS: usize = MAX_FIELDS / 32;

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
struct EntrySummary {
    ordinal: usize,
    key: u32,
    ref_count: u32,
    string_value: Option<SliceSummary>,
    reference: Option<ReferenceSummary>,
    formula: Option<SliceSummary>,
    format: Option<SliceSummary>,
    custom_format: Option<SliceSummary>,
    rich_text_payload: Option<ReferenceSummary>,
    comment_storage: Option<ReferenceSummary>,
    import_warning_set: Option<SliceSummary>,
    cell_spec: Option<SliceSummary>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct SegmentSummary {
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

struct RootCollector<'expected> {
    source: SourceRange,
    expected_entries: Option<&'expected [litchi_iwa_protos::tst::table_data_list::ListEntry]>,
    expected_segments: Option<&'expected [Reference]>,
    entries: Vec<EntrySummary>,
    segments: Vec<SegmentSummary>,
    total_entries: usize,
    total_segments: usize,
}

impl<'expected> RootCollector<'expected> {
    fn new(
        source: &[u8],
        expected_entries: Option<&'expected [litchi_iwa_protos::tst::table_data_list::ListEntry]>,
        expected_segments: Option<&'expected [Reference]>,
    ) -> Self {
        Self {
            source: SourceRange::new(source),
            expected_entries,
            expected_segments,
            entries: Vec::with_capacity(MAX_RETAINED_RECORDS),
            segments: Vec::with_capacity(MAX_RETAINED_RECORDS),
            total_entries: 0,
            total_segments: 0,
        }
    }

    fn observe_entry(&mut self, entry: TableDataListEntrySnapshot<'_>) {
        let summary = summarize_entry(self.source, entry, self.total_entries);
        if let Some(expected_entries) = self.expected_entries {
            let expected = expected_entries
                .get(self.total_entries)
                .unwrap_or_else(|| panic!("streamed root entry count exceeded Prost entry count"));
            assert_entry_matches(entry, expected);
        }
        self.total_entries = self
            .total_entries
            .checked_add(1)
            .expect("bounded entry callback count");
        if self.entries.len() < MAX_RETAINED_RECORDS {
            self.entries.push(summary);
        }
    }

    fn observe_segment(&mut self, record: ReferenceRecord<'_>) {
        let summary = SegmentSummary {
            ordinal: self.total_segments,
            raw: self.source.assert_borrowed(record.raw()),
            reference: summarize_reference(record.reference()),
        };
        if let Some(expected_segments) = self.expected_segments {
            let expected = expected_segments
                .get(self.total_segments)
                .unwrap_or_else(|| {
                    panic!("streamed root segment count exceeded Prost segment count")
                });
            assert_reference_matches(record.reference(), expected);
        }
        self.total_segments = self
            .total_segments
            .checked_add(1)
            .expect("bounded segment callback count");
        if self.segments.len() < MAX_RETAINED_RECORDS {
            self.segments.push(summary);
        }
    }
}

impl StorageVisitor for RootCollector<'_> {
    fn visit_list_entry(
        &mut self,
        entry: TableDataListEntrySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        self.observe_entry(entry);
        Ok(())
    }

    fn visit_list_segment(&mut self, record: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.observe_segment(record);
        Ok(())
    }
}

struct SegmentCollector<'expected> {
    source: SourceRange,
    expected_entries: Option<&'expected [litchi_iwa_protos::tst::table_data_list::ListEntry]>,
    entries: Vec<EntrySummary>,
    total_entries: usize,
}

impl<'expected> SegmentCollector<'expected> {
    fn new(
        source: &[u8],
        expected_entries: Option<&'expected [litchi_iwa_protos::tst::table_data_list::ListEntry]>,
    ) -> Self {
        Self {
            source: SourceRange::new(source),
            expected_entries,
            entries: Vec::with_capacity(MAX_RETAINED_RECORDS),
            total_entries: 0,
        }
    }

    fn observe_entry(&mut self, entry: TableDataListEntrySnapshot<'_>) {
        let summary = summarize_entry(self.source, entry, self.total_entries);
        if let Some(expected_entries) = self.expected_entries {
            let expected = expected_entries.get(self.total_entries).unwrap_or_else(|| {
                panic!("streamed segment entry count exceeded Prost entry count")
            });
            assert_entry_matches(entry, expected);
        }
        self.total_entries = self
            .total_entries
            .checked_add(1)
            .expect("bounded entry callback count");
        if self.entries.len() < MAX_RETAINED_RECORDS {
            self.entries.push(summary);
        }
    }
}

impl StorageVisitor for SegmentCollector<'_> {
    fn visit_list_entry(
        &mut self,
        entry: TableDataListEntrySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        self.observe_entry(entry);
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
    // Root and segment calls deliberately have independent outcomes. A root
    // rejection must not prevent the segment entry point from seeing the same
    // source, and vice versa.
    exercise_root(source);
    exercise_segment(source);
}

fn exercise_root(source: &[u8]) {
    let before = source.to_vec();
    let strict = decode_table_data_list_with_report(source, options());
    assert_eq!(
        source,
        before.as_slice(),
        "root scalar decode modified source"
    );

    // The generated Prost oracle is intentionally consulted only after strict
    // acceptance. Prost's permissive proto2 required-field behavior is not the
    // acceptance policy; it is a parity oracle for strict successes.
    let prost = strict.as_ref().ok().map(|_| {
        TableDataList::decode(source)
            .unwrap_or_else(|error| panic!("strict root acceptance disagreed with Prost: {error}"))
    });
    assert_eq!(
        source,
        before.as_slice(),
        "Prost root decode modified source"
    );
    let expected_entries = prost.as_ref().map(|root| root.entries.as_slice());
    let expected_segments = prost.as_ref().map(|root| root.segments.as_slice());
    let mut collector = RootCollector::new(source, expected_entries, expected_segments);
    let streamed = decode_table_data_list_with_visitor(source, options(), &mut collector);
    assert_eq!(
        source,
        before.as_slice(),
        "root visitor decode modified source"
    );

    match (strict, streamed) {
        (Ok((strict_snapshot, strict_report)), Ok((streamed_snapshot, streamed_report))) => {
            assert_eq!(streamed_snapshot, strict_snapshot);
            assert_eq!(streamed_report, strict_report);
            let prost = prost
                .as_ref()
                .expect("strict root success must have a Prost oracle");
            assert_root_matches(strict_snapshot, prost);
            assert_eq!(collector.total_entries, prost.entries.len());
            assert_eq!(collector.total_segments, prost.segments.len());
            assert_entry_prefix_summaries(&collector.entries);
            assert_segment_prefix_summaries(&collector.segments);
        },
        (Err(_), Err(_)) => {
            // A visitor can have observed a valid prefix before a later
            // malformed field. Since the decode failed, that prefix is never
            // published; borrow checks above still cover every callback.
        },
        _ => panic!("root scalar and visitor paths disagreed on success"),
    }
}

fn exercise_segment(source: &[u8]) {
    let before = source.to_vec();
    let strict = decode_table_data_list_segment_with_report(source, options());
    assert_eq!(
        source,
        before.as_slice(),
        "segment scalar decode modified source"
    );

    let prost = strict.as_ref().ok().map(|_| {
        TableDataListSegment::decode(source).unwrap_or_else(|error| {
            panic!("strict segment acceptance disagreed with Prost: {error}")
        })
    });
    assert_eq!(
        source,
        before.as_slice(),
        "Prost segment decode modified source"
    );
    let expected_entries = prost.as_ref().map(|segment| segment.entries.as_slice());
    let mut collector = SegmentCollector::new(source, expected_entries);
    let streamed = decode_table_data_list_segment_with_visitor(source, options(), &mut collector);
    assert_eq!(
        source,
        before.as_slice(),
        "segment visitor decode modified source"
    );

    match (strict, streamed) {
        (Ok((strict_snapshot, strict_report)), Ok((streamed_snapshot, streamed_report))) => {
            assert_eq!(streamed_snapshot, strict_snapshot);
            assert_eq!(streamed_report, strict_report);
            let prost = prost
                .as_ref()
                .expect("strict segment success must have a Prost oracle");
            assert_segment_matches(strict_snapshot, prost);
            SourceRange::new(source).assert_borrowed(strict_snapshot.key_range());
            assert_eq!(collector.total_entries, prost.entries.len());
            assert_entry_prefix_summaries(&collector.entries);
        },
        (Err(_), Err(_)) => {},
        _ => panic!("segment scalar and visitor paths disagreed on success"),
    }
}

fn summarize_entry(
    source: SourceRange,
    entry: TableDataListEntrySnapshot<'_>,
    ordinal: usize,
) -> EntrySummary {
    EntrySummary {
        ordinal,
        key: entry.key(),
        ref_count: entry.ref_count(),
        string_value: entry
            .string_value()
            .map(|value| source.assert_borrowed(value.as_bytes())),
        reference: entry.reference().map(summarize_reference),
        formula: entry.formula().map(|value| source.assert_borrowed(value)),
        format: entry.format().map(|value| source.assert_borrowed(value)),
        custom_format: entry
            .custom_format()
            .map(|value| source.assert_borrowed(value)),
        rich_text_payload: entry.rich_text_payload().map(summarize_reference),
        comment_storage: entry.comment_storage().map(summarize_reference),
        import_warning_set: entry
            .import_warning_set()
            .map(|value| source.assert_borrowed(value)),
        cell_spec: entry.cell_spec().map(|value| source.assert_borrowed(value)),
    }
}

fn summarize_reference(reference: ReferenceSnapshot) -> ReferenceSummary {
    ReferenceSummary {
        identifier: reference.identifier(),
        deprecated_type: reference.deprecated_type(),
        deprecated_is_external: reference.deprecated_is_external(),
    }
}

fn assert_entry_matches(
    strict: TableDataListEntrySnapshot<'_>,
    prost: &litchi_iwa_protos::tst::table_data_list::ListEntry,
) {
    assert_eq!(strict.key(), prost.key);
    assert_eq!(strict.ref_count(), prost.refcount);
    assert_eq!(strict.string_value(), prost.string.as_deref());
    assert_reference_matches_optional(strict.reference(), prost.reference.as_ref());
    assert_opaque_message_matches(strict.formula(), &prost.formula);
    assert_opaque_message_matches(strict.format(), &prost.format);
    assert_opaque_message_matches(strict.custom_format(), &prost.custom_format);
    assert_reference_matches_optional(strict.rich_text_payload(), prost.rich_text_payload.as_ref());
    assert_reference_matches_optional(strict.comment_storage(), prost.comment_storage.as_ref());
    assert_opaque_message_matches(strict.import_warning_set(), &prost.import_warning_set);
    assert_opaque_message_matches(strict.cell_spec(), &prost.cell_spec);
}

fn assert_opaque_message_matches<T>(strict: Option<&[u8]>, prost: &Option<T>)
where
    T: prost::Message + Default + PartialEq + std::fmt::Debug,
{
    match (strict, prost.as_ref()) {
        (Some(raw), Some(expected)) => {
            let decoded = T::decode(raw).unwrap_or_else(|error| {
                panic!("strict opaque message acceptance disagreed with Prost: {error}")
            });
            assert_eq!(&decoded, expected);
        },
        (None, None) => {},
        _ => panic!("opaque message presence differed between strict and Prost paths"),
    }
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

fn assert_root_matches(
    strict: litchi_iwa_protos::numbers_table_cell_storage_codec::TableDataListSnapshot,
    prost: &TableDataList,
) {
    assert_eq!(strict.list_type(), prost.list_type);
    assert_eq!(strict.next_list_id(), prost.next_list_id);
    assert_eq!(strict.is_new_for_bnc(), prost.is_new_for_bnc);
}

fn assert_segment_matches(strict: TableDataListSegmentSnapshot<'_>, prost: &TableDataListSegment) {
    assert_eq!(strict.list_type(), prost.list_type);
    assert_eq!(strict.key_range_location(), prost.key_range.location);
    assert_eq!(strict.key_range_length(), prost.key_range.length);
}

fn assert_entry_prefix_summaries(summaries: &[EntrySummary]) {
    // The concrete fields are checked as callbacks arrive. Retaining this
    // bounded prefix makes pointer/length and source-order state inspectable
    // after the decode without retaining every repeated item.
    for (index, summary) in summaries.iter().enumerate() {
        assert_eq!(summary.ordinal, index);
        assert_optional_slice_summary(summary.string_value);
        assert_optional_slice_summary(summary.formula);
        assert_optional_slice_summary(summary.format);
        assert_optional_slice_summary(summary.custom_format);
        assert_optional_slice_summary(summary.import_warning_set);
        assert_optional_slice_summary(summary.cell_spec);
        std::hint::black_box((
            summary.key,
            summary.ref_count,
            summary.reference,
            summary.rich_text_payload,
            summary.comment_storage,
        ));
    }
}

fn assert_segment_prefix_summaries(summaries: &[SegmentSummary]) {
    for (index, summary) in summaries.iter().enumerate() {
        assert_eq!(summary.ordinal, index);
        assert_slice_summary(summary.raw);
        std::hint::black_box(summary.reference);
    }
}

fn assert_optional_slice_summary(summary: Option<SliceSummary>) {
    if let Some(summary) = summary {
        assert_slice_summary(summary);
    }
}

fn assert_slice_summary(summary: SliceSummary) {
    assert!(summary.pointer != 0 || summary.length == 0);
}
