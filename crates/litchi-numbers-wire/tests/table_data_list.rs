//! Independent contract tests for the shared Numbers table-data-list
//! coordinator.
//!
//! The envelopes below are handwritten protobuf wire messages.  The adapter
//! used by the tests delegates envelope parsing to the strict generated-free
//! storage codec, while keeping admission, topology, and publication under
//! test at the wire-crate boundary.  In particular, a candidate can be
//! visited before a later malformed field is discovered; no visitor prefix
//! may escape as a successful list in that case.

use std::collections::HashSet;

use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    DecodeError, DecodeLimit, DecodeOptions, ReferenceRecord, StorageVisitor,
    TableDataListEntryRecord, decode_table_data_list_segment_type_with_report,
    decode_table_data_list_segment_with_visitor, decode_table_data_list_type_with_report,
    decode_table_data_list_with_visitor,
};
use litchi_numbers_wire::table_data_list::{
    Candidate, CoordinatorIssue, KeyRange, ListDecoder, ListReadPolicy, Message,
    NATIVE_TABLE_DATA_LIST_MESSAGE_KIND, OverflowPolicy, RootOrSegment,
    TABLE_DATA_LIST_MESSAGE_KIND, TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND, read_list,
};

const STRING_LIST: i32 = 1;
const FORMULA_LIST: i32 = 3;

#[derive(Debug, PartialEq, Eq)]
enum HarnessError {
    Codec(DecodeError),
    Issue(CoordinatorIssue),
    Semantic(&'static str),
}

#[derive(Debug)]
struct CodecDecoder {
    options: DecodeOptions,
    max_visited_entries: Option<usize>,
    inject_semantic_error: bool,
    decode_calls: Vec<(RootOrSegment, bool, u64)>,
    visited_entries: usize,
    borrowed_entry_spans: Vec<(usize, usize)>,
    mapped_issues: Vec<CoordinatorIssue>,
}

impl Default for CodecDecoder {
    fn default() -> Self {
        Self {
            options: DecodeOptions::new(16 * 1024, 4_096, 256 * 1024, 64, 4_096, 64 * 1024),
            max_visited_entries: None,
            inject_semantic_error: false,
            decode_calls: Vec::new(),
            visited_entries: 0,
            borrowed_entry_spans: Vec::new(),
            mapped_issues: Vec::new(),
        }
    }
}

impl CodecDecoder {
    fn new() -> Self {
        Self {
            options: DecodeOptions::new(16 * 1024, 4_096, 256 * 1024, 64, 4_096, 64 * 1024),
            ..Self::default()
        }
    }

    fn with_entry_ledger(max_entries: usize) -> Self {
        Self {
            max_visited_entries: Some(max_entries),
            ..Self::new()
        }
    }
}

struct RecordingVisitor {
    admit: bool,
    max_entries: Option<usize>,
    values: Vec<(u32, String)>,
    keys: HashSet<u32>,
    segment_refs: Vec<u64>,
    minimum: Option<u32>,
    maximum: Option<u32>,
    duplicate_key: Option<u32>,
    raw_spans: Vec<(usize, usize)>,
}

type VisitorParts = (
    Vec<(u32, String)>,
    HashSet<u32>,
    Vec<u64>,
    Option<u32>,
    Option<u32>,
    Option<u32>,
    Vec<(usize, usize)>,
);

impl RecordingVisitor {
    fn new(admit: bool, max_entries: Option<usize>) -> Self {
        Self {
            admit,
            max_entries,
            values: Vec::new(),
            keys: HashSet::new(),
            segment_refs: Vec::new(),
            minimum: None,
            maximum: None,
            duplicate_key: None,
            raw_spans: Vec::new(),
        }
    }

    fn finish(self) -> VisitorParts {
        (
            self.values,
            self.keys,
            self.segment_refs,
            self.minimum,
            self.maximum,
            self.duplicate_key,
            self.raw_spans,
        )
    }
}

impl StorageVisitor for RecordingVisitor {
    fn visit_list_entry_record(
        &mut self,
        record: TableDataListEntryRecord<'_>,
    ) -> Result<(), DecodeError> {
        let entry = record.snapshot();
        if let Some(maximum) = self.max_entries
            && self.keys.len() >= maximum
        {
            return Err(DecodeError::allocation(self.keys.len().saturating_add(1)));
        }
        self.raw_spans
            .push((record.raw().as_ptr() as usize, record.raw().len()));
        let key = entry.key();
        self.minimum = Some(self.minimum.map_or(key, |current| current.min(key)));
        self.maximum = Some(self.maximum.map_or(key, |current| current.max(key)));
        if !self.keys.insert(key) && self.duplicate_key.is_none() {
            self.duplicate_key = Some(key);
        }
        if self.admit {
            self.values
                .push((key, entry.string_value().unwrap_or_default().to_owned()));
        }
        Ok(())
    }

    fn visit_list_segment(&mut self, record: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.segment_refs.push(record.reference().identifier());
        Ok(())
    }
}

impl<'source> ListDecoder<'source> for CodecDecoder {
    type Value = String;
    type Error = HarnessError;

    fn probe(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        _object_id: u64,
    ) -> Result<i32, Self::Error> {
        let result = match kind {
            RootOrSegment::Root => decode_table_data_list_type_with_report(source, self.options),
            RootOrSegment::Segment => {
                decode_table_data_list_segment_type_with_report(source, self.options)
            },
        };
        result
            .map(|(snapshot, _report)| snapshot.list_type())
            .map_err(HarnessError::Codec)
    }

    fn decode(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        object_id: u64,
        admit: bool,
    ) -> Result<Candidate<Self::Value, Self::Error>, Self::Error> {
        self.decode_calls.push((kind, admit, object_id));
        let mut visitor = RecordingVisitor::new(admit, self.max_visited_entries);
        let decoded = match kind {
            RootOrSegment::Root => {
                decode_table_data_list_with_visitor(source, self.options, &mut visitor)
                    .map(|(snapshot, report)| (snapshot.list_type(), None, report))
            },
            RootOrSegment::Segment => {
                decode_table_data_list_segment_with_visitor(source, self.options, &mut visitor).map(
                    |(snapshot, report)| {
                        (
                            snapshot.list_type(),
                            Some(KeyRange {
                                location: snapshot.key_range_location(),
                                length: snapshot.key_range_length(),
                            }),
                            report,
                        )
                    },
                )
            },
        };
        let visited_before_error = visitor.raw_spans.len();
        self.visited_entries = self.visited_entries.saturating_add(visited_before_error);
        let (list_type, key_range, _report) = decoded.map_err(HarnessError::Codec)?;
        let (values, keys, segment_refs, minimum, maximum, duplicate_key, raw_spans) =
            visitor.finish();
        self.borrowed_entry_spans.extend(raw_spans);
        Ok(Candidate {
            list_type,
            values,
            keys,
            segment_refs,
            key_range,
            entry_bounds: minimum.zip(maximum).map(|(minimum, maximum)| {
                litchi_numbers_wire::table_data_list::EntryBounds { minimum, maximum }
            }),
            structural_error: duplicate_key
                .map(|key| HarnessError::Issue(CoordinatorIssue::DuplicateEntryKey { key })),
            semantic_error: self
                .inject_semantic_error
                .then_some(HarnessError::Semantic("deferred conversion failure")),
        })
    }

    fn map_issue(&mut self, issue: CoordinatorIssue) -> Self::Error {
        self.mapped_issues.push(issue);
        HarnessError::Issue(issue)
    }
}

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    loop {
        let byte = u8::try_from(value & 0x7f).expect("varint chunk fits in u8");
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return output;
        }
        output.push(byte | 0x80);
    }
}

fn field_key(number: u32, wire_type: u8) -> Vec<u8> {
    varint((u64::from(number) << 3) | u64::from(wire_type))
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    let mut output = field_key(number, 0);
    output.extend(varint(value));
    output
}

fn field_bytes(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = field_key(number, 2);
    output.extend(varint(
        u64::try_from(payload.len()).expect("fixture fits in u64"),
    ));
    output.extend(payload);
    output
}

fn string_entry(key: u32, value: &str) -> Vec<u8> {
    let mut output = field_varint(1, u64::from(key));
    output.extend(field_varint(2, 1));
    output.extend(field_bytes(3, value.as_bytes()));
    output
}

fn reference(identifier: u64) -> Vec<u8> {
    field_varint(1, identifier)
}

fn root(list_type: i32, entries: &[Vec<u8>], segments: &[u64]) -> Vec<u8> {
    let mut output = field_varint(
        1,
        u64::try_from(list_type).expect("list type is nonnegative"),
    );
    output.extend(field_varint(2, 100));
    for entry in entries {
        output.extend(field_bytes(3, entry));
    }
    for segment in segments {
        output.extend(field_bytes(4, &reference(*segment)));
    }
    output
}

fn segment(list_type: i32, location: u32, length: u32, entries: &[Vec<u8>]) -> Vec<u8> {
    let mut range = field_varint(1, u64::from(location));
    range.extend(field_varint(2, u64::from(length)));
    let mut output = field_varint(
        1,
        u64::try_from(list_type).expect("list type is nonnegative"),
    );
    output.extend(field_bytes(2, &range));
    for entry in entries {
        output.extend(field_bytes(3, entry));
    }
    output
}

fn unknown(mut source: Vec<u8>) -> Vec<u8> {
    source.extend(field_bytes(100, &[0xde, 0xad, 0xbe, 0xef]));
    source
}

fn malformed_later_entry(mut source: Vec<u8>) -> Vec<u8> {
    // The routing probe validates the outer length-delimited framing but does
    // not descend into repeated entries. The full decoder must still reject
    // this later entry after publishing no earlier callback prefix.
    source.extend(field_bytes(3, &field_varint(1, 2)));
    source
}

fn contains_span(source: &[u8], span: (usize, usize)) -> bool {
    let start = source.as_ptr() as usize;
    let end = start.saturating_add(source.len());
    span.0 >= start && span.0.saturating_add(span.1) <= end
}

fn issue<T: std::fmt::Debug>(result: Result<T, HarnessError>) -> CoordinatorIssue {
    match result {
        Err(HarnessError::Issue(issue)) => issue,
        other => panic!("expected coordinator issue, got {other:?}"),
    }
}

#[test]
fn valid_root_and_segment_are_sorted_and_borrow_entry_source() {
    let root_source = root(
        STRING_LIST,
        &[string_entry(8, "root-eight"), string_entry(10, "root-ten")],
        &[41],
    );
    let segment_source = segment(
        STRING_LIST,
        1,
        8,
        &[
            string_entry(4, "segment-four"),
            string_entry(2, "segment-two"),
        ],
    );
    let roots = vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)];
    let mut decoder = CodecDecoder::new();
    let values = read_list(
        90,
        STRING_LIST,
        roots,
        |segment_id| {
            Ok((segment_id == 41).then(|| {
                vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &segment_source,
                )]
            }))
        },
        &mut decoder,
        ListReadPolicy::default(),
    )
    .expect("valid root and segment");

    assert_eq!(
        values,
        vec![
            (2, "segment-two".to_owned()),
            (4, "segment-four".to_owned()),
            (8, "root-eight".to_owned()),
            (10, "root-ten".to_owned()),
        ]
    );
    assert_eq!(decoder.visited_entries, 4);
    assert!(
        decoder
            .borrowed_entry_spans
            .iter()
            .any(|&span| contains_span(&root_source, span))
    );
    assert!(
        decoder
            .borrowed_entry_spans
            .iter()
            .any(|&span| contains_span(&segment_source, span))
    );
}

#[test]
fn wrong_list_candidates_are_strictly_walked_without_admission() {
    let wrong = unknown(root(FORMULA_LIST, &[string_entry(99, "discarded")], &[]));
    let selected = root(STRING_LIST, &[string_entry(7, "kept")], &[]);
    let roots = vec![
        Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &wrong),
        Message::new(NATIVE_TABLE_DATA_LIST_MESSAGE_KIND, &selected),
    ];
    let mut decoder = CodecDecoder::new();
    let values = read_list(
        17,
        STRING_LIST,
        roots,
        |_| Ok(None::<Vec<Message<'_>>>),
        &mut decoder,
        ListReadPolicy::default(),
    )
    .expect("wrong type is ignored and selected root is admitted");

    assert_eq!(values, vec![(7, "kept".to_owned())]);
    assert_eq!(decoder.visited_entries, 2);
    assert_eq!(decoder.decode_calls.len(), 2);
    assert_eq!(decoder.decode_calls[0], (RootOrSegment::Root, false, 17));
    assert_eq!(decoder.decode_calls[1], (RootOrSegment::Root, true, 17));
}

#[test]
fn duplicate_roots_and_segment_payloads_are_walked_then_rejected() {
    let first_root = root(STRING_LIST, &[string_entry(1, "first")], &[]);
    let second_root = unknown(root(STRING_LIST, &[string_entry(2, "second")], &[]));
    let roots = vec![
        Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &first_root),
        Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &second_root),
    ];
    let mut decoder = CodecDecoder::new();
    assert_eq!(
        issue(read_list(
            90,
            STRING_LIST,
            roots,
            |_| Ok(None::<Vec<Message<'_>>>),
            &mut decoder,
            ListReadPolicy::default(),
        )),
        CoordinatorIssue::DuplicateRoot {
            object_id: 90,
            expected_type: STRING_LIST,
        }
    );
    assert_eq!(decoder.visited_entries, 2);
    assert_eq!(decoder.decode_calls[1], (RootOrSegment::Root, false, 90));

    let root_source = root(STRING_LIST, &[], &[41]);
    let first_segment = segment(STRING_LIST, 1, 1, &[string_entry(1, "first")]);
    let second_segment = unknown(segment(STRING_LIST, 2, 1, &[string_entry(2, "second")]));
    let mut decoder = CodecDecoder::new();
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |segment_id| {
            Ok((segment_id == 41).then(|| {
                vec![
                    Message::new(TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND, &first_segment),
                    Message::new(TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND, &second_segment),
                ]
            }))
        },
        &mut decoder,
        ListReadPolicy::default(),
    );
    assert_eq!(
        issue(result),
        CoordinatorIssue::DuplicateSegmentPayload { segment_id: 41 }
    );
    assert_eq!(decoder.visited_entries, 2);
    assert_eq!(decoder.decode_calls[2], (RootOrSegment::Segment, false, 41));
}

#[test]
fn missing_root_and_missing_segment_are_distinct_structural_failures() {
    let wrong_root = root(FORMULA_LIST, &[], &[]);
    let mut decoder = CodecDecoder::new();
    assert_eq!(
        issue(read_list(
            7,
            STRING_LIST,
            vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &wrong_root)],
            |_| Ok(None::<Vec<Message<'_>>>),
            &mut decoder,
            ListReadPolicy::default(),
        )),
        CoordinatorIssue::MissingRoot {
            object_id: 7,
            expected_type: STRING_LIST,
        }
    );

    let selected = root(STRING_LIST, &[], &[404]);
    let mut decoder = CodecDecoder::new();
    assert_eq!(
        issue(read_list(
            7,
            STRING_LIST,
            vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &selected)],
            |_| Ok(None::<Vec<Message<'_>>>),
            &mut decoder,
            ListReadPolicy::default(),
        )),
        CoordinatorIssue::MissingSegment {
            object_id: 7,
            segment_id: 404,
        }
    );

    let selected = root(STRING_LIST, &[], &[405]);
    let unrelated = root(STRING_LIST, &[], &[]);
    let mut decoder = CodecDecoder::new();
    assert_eq!(
        issue(read_list(
            7,
            STRING_LIST,
            vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &selected)],
            |_| Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_MESSAGE_KIND,
                &unrelated,
            )])),
            &mut decoder,
            ListReadPolicy::default(),
        )),
        CoordinatorIssue::MissingSegmentPayload { segment_id: 405 }
    );
}

#[test]
fn wrong_segment_type_is_walked_without_admitting_values() {
    let root_source = root(STRING_LIST, &[], &[41]);
    let wrong_segment = unknown(segment(FORMULA_LIST, 1, 1, &[string_entry(1, "ignored")]));
    let mut decoder = CodecDecoder::new();
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |segment_id| {
            Ok((segment_id == 41).then(|| {
                vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &wrong_segment,
                )]
            }))
        },
        &mut decoder,
        ListReadPolicy::default(),
    );
    assert_eq!(
        issue(result),
        CoordinatorIssue::WrongSegmentType {
            segment_id: 41,
            expected_type: STRING_LIST,
            actual_type: FORMULA_LIST,
        }
    );
    assert_eq!(decoder.visited_entries, 1);
    assert_eq!(decoder.decode_calls[1], (RootOrSegment::Segment, false, 41));
}

#[test]
fn repeated_keys_across_root_and_segments_are_rejected() {
    let root_source = root(STRING_LIST, &[string_entry(5, "root")], &[41, 42]);
    let first_segment = segment(STRING_LIST, 1, 5, &[string_entry(2, "first")]);
    let second_segment = segment(STRING_LIST, 5, 3, &[string_entry(5, "duplicate")]);
    let mut decoder = CodecDecoder::new();
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |segment_id| {
            Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                if segment_id == 41 {
                    &first_segment
                } else {
                    &second_segment
                },
            )]))
        },
        &mut decoder,
        ListReadPolicy::default(),
    );
    assert_eq!(
        issue(result),
        CoordinatorIssue::DuplicateEntryKey { key: 5 }
    );
}

#[test]
fn repeated_keys_between_segments_are_rejected_even_when_root_is_empty() {
    let root_source = root(STRING_LIST, &[], &[41, 42]);
    let first_segment = segment(STRING_LIST, 1, 2, &[string_entry(1, "first")]);
    let second_segment = segment(STRING_LIST, 1, 2, &[string_entry(1, "duplicate")]);
    let mut decoder = CodecDecoder::new();
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |segment_id| {
            Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                if segment_id == 41 {
                    &first_segment
                } else {
                    &second_segment
                },
            )]))
        },
        &mut decoder,
        ListReadPolicy::default(),
    );
    assert_eq!(
        issue(result),
        CoordinatorIssue::DuplicateEntryKey { key: 1 }
    );
}

#[test]
fn segment_range_overflow_and_out_of_range_entries_fail_closed() {
    let overflow_root = root(STRING_LIST, &[], &[41, 42]);
    let overflow_segment = segment(STRING_LIST, u32::MAX, 1, &[]);
    let malformed_later_segment =
        malformed_later_entry(segment(STRING_LIST, 10, 1, &[string_entry(2, "later")]));
    let mut immediate_decoder = CodecDecoder::new();
    let mut immediate_resolved = Vec::new();
    assert_eq!(
        issue(read_list(
            90,
            STRING_LIST,
            vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &overflow_root)],
            |segment_id| {
                immediate_resolved.push(segment_id);
                Ok(Some(vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    if segment_id == 41 {
                        &overflow_segment
                    } else {
                        &malformed_later_segment
                    },
                )]))
            },
            &mut immediate_decoder,
            ListReadPolicy {
                max_entries: usize::MAX,
                overflow: OverflowPolicy::Deferred,
                range_overflow: OverflowPolicy::Immediate,
            },
        )),
        CoordinatorIssue::KeyRangeOverflow { segment_id: 41 }
    );
    assert_eq!(immediate_resolved, vec![41]);

    let mut deferred_decoder = CodecDecoder::new();
    let mut deferred_resolved = Vec::new();
    let deferred = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &overflow_root)],
        |segment_id| {
            deferred_resolved.push(segment_id);
            Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                if segment_id == 41 {
                    &overflow_segment
                } else {
                    &malformed_later_segment
                },
            )]))
        },
        &mut deferred_decoder,
        ListReadPolicy {
            max_entries: usize::MAX,
            overflow: OverflowPolicy::Deferred,
            range_overflow: OverflowPolicy::Deferred,
        },
    );
    assert!(matches!(deferred, Err(HarnessError::Codec(_))));
    assert_eq!(deferred_resolved, vec![41, 42]);

    let outside_root = root(STRING_LIST, &[], &[41]);
    let outside_segment = segment(STRING_LIST, 10, 1, &[string_entry(9, "outside")]);
    let mut decoder = CodecDecoder::new();
    assert_eq!(
        issue(read_list(
            90,
            STRING_LIST,
            vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &outside_root)],
            |_segment_id| {
                Ok(Some(vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &outside_segment,
                )]))
            },
            &mut decoder,
            ListReadPolicy::default(),
        )),
        CoordinatorIssue::EntryOutsideKeyRange { segment_id: 41 }
    );
}

#[test]
fn malformed_later_wire_overrides_visited_semantic_prefix() {
    let malformed = malformed_later_entry(root(STRING_LIST, &[string_entry(1, "prefix")], &[]));
    let mut decoder = CodecDecoder::new();
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &malformed)],
        |_| Ok(None::<Vec<Message<'_>>>),
        &mut decoder,
        ListReadPolicy::default(),
    );
    assert!(matches!(result, Err(HarnessError::Codec(_))));
    assert_eq!(decoder.visited_entries, 1);
    assert!(decoder.mapped_issues.is_empty());
}

#[test]
fn structural_error_precedes_deferred_candidate_semantics() {
    let root_source = root(STRING_LIST, &[], &[41]);
    let overflow_segment = segment(STRING_LIST, u32::MAX, 1, &[]);
    let mut decoder = CodecDecoder::new();
    decoder.inject_semantic_error = true;
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |_| {
            Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                &overflow_segment,
            )]))
        },
        &mut decoder,
        ListReadPolicy::default(),
    );
    assert_eq!(
        issue(result),
        CoordinatorIssue::KeyRangeOverflow { segment_id: 41 }
    );
}

#[test]
fn unknown_fields_are_strictly_scanned_without_changing_projection() {
    let entry = unknown(string_entry(3, "opaque-tail"));
    let selected = unknown(root(STRING_LIST, &[entry], &[]));
    let mut decoder = CodecDecoder::new();
    let values = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &selected)],
        |_| Ok(None::<Vec<Message<'_>>>),
        &mut decoder,
        ListReadPolicy::default(),
    )
    .expect("unknown fields remain opaque");
    assert_eq!(values, vec![(3, "opaque-tail".to_owned())]);
    assert_eq!(decoder.visited_entries, 1);
}

#[test]
fn entry_limit_is_inclusive_and_overflow_policy_controls_traversal() {
    let root_source = root(
        STRING_LIST,
        &[string_entry(1, "one"), string_entry(3, "three")],
        &[41, 42],
    );
    let segment_source = segment(STRING_LIST, 2, 1, &[string_entry(2, "two")]);
    let second_segment_source = segment(STRING_LIST, 4, 1, &[string_entry(4, "four")]);

    let mut resolved_immediate = Vec::new();
    let mut immediate_decoder = CodecDecoder::new();
    let immediate = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |segment_id| {
            resolved_immediate.push(segment_id);
            Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                if segment_id == 41 {
                    &segment_source
                } else {
                    &second_segment_source
                },
            )]))
        },
        &mut immediate_decoder,
        ListReadPolicy {
            max_entries: 2,
            overflow: OverflowPolicy::Immediate,
            range_overflow: OverflowPolicy::Immediate,
        },
    );
    assert_eq!(
        issue(immediate),
        CoordinatorIssue::EntryLimit {
            observed: 3,
            maximum: 2,
        }
    );
    assert_eq!(resolved_immediate, vec![41]);

    let mut resolved_deferred = Vec::new();
    let mut deferred_decoder = CodecDecoder::new();
    let deferred = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root_source)],
        |segment_id| {
            resolved_deferred.push(segment_id);
            Ok(Some(vec![Message::new(
                TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                if segment_id == 41 {
                    &segment_source
                } else {
                    &second_segment_source
                },
            )]))
        },
        &mut deferred_decoder,
        ListReadPolicy {
            max_entries: 2,
            overflow: OverflowPolicy::Deferred,
            range_overflow: OverflowPolicy::Deferred,
        },
    );
    assert_eq!(
        issue(deferred),
        CoordinatorIssue::EntryLimit {
            observed: 3,
            maximum: 2,
        }
    );
    assert_eq!(resolved_deferred, vec![41, 42]);

    let boundary_root = root(
        STRING_LIST,
        &[string_entry(1, "one"), string_entry(3, "three")],
        &[41],
    );
    let mut decoder = CodecDecoder::new();
    let at_boundary = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &boundary_root)],
        |segment_id| {
            Ok((segment_id == 41).then(|| {
                vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &segment_source,
                )]
            }))
        },
        &mut decoder,
        ListReadPolicy {
            max_entries: 3,
            overflow: OverflowPolicy::Deferred,
            range_overflow: OverflowPolicy::Deferred,
        },
    )
    .expect("the inclusive boundary admits exactly three entries");
    assert_eq!(at_boundary.len(), 3);
}

#[test]
fn fallible_entry_ledger_refusal_is_returned_without_publication() {
    let selected = root(
        STRING_LIST,
        &[string_entry(1, "first"), string_entry(2, "second")],
        &[],
    );
    let mut decoder = CodecDecoder::with_entry_ledger(1);
    let result = read_list(
        90,
        STRING_LIST,
        vec![Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &selected)],
        |_| Ok(None::<Vec<Message<'_>>>),
        &mut decoder,
        ListReadPolicy::default(),
    );
    match result {
        Err(HarnessError::Codec(error)) => {
            assert_eq!(
                error.resource_limit(),
                Some(DecodeLimit::Allocation { requested: 2 })
            );
        },
        other => panic!("expected fallible ledger refusal, got {other:?}"),
    }
    assert_eq!(decoder.visited_entries, 1);
    assert!(decoder.mapped_issues.is_empty());
}
