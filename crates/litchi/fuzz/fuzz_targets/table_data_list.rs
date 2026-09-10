#![no_main]

//! Bounded raw and structured fuzzing for the Numbers table-data-list
//! coordinator.
//!
//! The coordinator receives borrowed root and segment messages from the
//! adapter. This target supplies a small generated-free storage decoder hook,
//! backed by the real table-cell storage codec, and keeps every synthetic
//! envelope within 1024 bytes, 16 entries, and four segment references.

use std::collections::HashSet;
use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;
use litchi_numbers_wire::table_data_list::{
    Candidate, CoordinatorIssue, EntryBounds, KeyRange, ListDecoder, ListReadPolicy, Message,
    NATIVE_TABLE_DATA_LIST_MESSAGE_KIND, OverflowPolicy, RootOrSegment,
    TABLE_DATA_LIST_MESSAGE_KIND, TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND, read_list,
};

const MAX_SOURCE_BYTES: usize = 1_024;
const MAX_ENTRIES: usize = 16;
const MAX_SEGMENTS: usize = 4;
const MAX_FIELDS: usize = 256;
const MAX_WORK_BYTES: usize = 16 * 1024;
const MAX_REFERENCES: usize = 64;
const MAX_TEXT_BYTES: usize = 1_024;
const MAX_RECURSION: u32 = 16;

const VALID_CASE: u8 = 0;
const NATIVE_ROOT_CASE: u8 = 1;
const MAX_CASE: u8 = 2;
const WRONG_ROOT_KIND_CASE: u8 = 3;
const DUPLICATE_ROOT_CASE: u8 = 4;
const DUPLICATE_SEGMENT_REFERENCE_CASE: u8 = 5;
const MISSING_SEGMENT_CASE: u8 = 6;
const MISSING_SEGMENT_PAYLOAD_CASE: u8 = 7;
const WRONG_SEGMENT_TYPE_CASE: u8 = 8;
const DUPLICATE_SEGMENT_PAYLOAD_CASE: u8 = 9;
const OUTSIDE_KEY_RANGE_CASE: u8 = 10;
const OVERFLOW_KEY_RANGE_CASE: u8 = 11;
const MALFORMED_KEY_RANGE_CASE: u8 = 12;
const MALFORMED_LATER_CASE: u8 = 13;
const DUPLICATE_ENTRY_KEY_CASE: u8 = 14;
const DUPLICATE_FIELD_CASE: u8 = 15;
const WRONG_WIRE_CASE: u8 = 16;
const BYTES_LIMIT_CASE: u8 = 17;
const FIELDS_LIMIT_CASE: u8 = 18;
const WORK_LIMIT_CASE: u8 = 19;
const REFERENCES_LIMIT_CASE: u8 = 20;
const TEXT_LIMIT_CASE: u8 = 21;
const POLICY_IMMEDIATE_CASE: u8 = 22;
const POLICY_DEFERRED_CASE: u8 = 23;
const CASE_COUNT: u8 = 24;

#[derive(Debug)]
enum HookError {
    Codec(storage::DecodeError),
    Coordinator(CoordinatorIssue),
}

/// A stack-only protobuf byte buffer. No generated message or AST is built
/// for a structured input, and no profile can exceed the source cap.
#[derive(Clone, Copy)]
struct FixedBytes {
    bytes: [u8; MAX_SOURCE_BYTES],
    len: usize,
}

impl FixedBytes {
    fn new() -> Self {
        Self {
            bytes: [0; MAX_SOURCE_BYTES],
            len: 0,
        }
    }

    fn push(&mut self, byte: u8) {
        assert!(
            self.len < self.bytes.len(),
            "table-list fixture exceeded cap"
        );
        self.bytes[self.len] = byte;
        self.len += 1;
    }

    fn extend(&mut self, bytes: &[u8]) {
        let end = self
            .len
            .checked_add(bytes.len())
            .expect("table-list fixture length overflow");
        assert!(end <= self.bytes.len(), "table-list fixture exceeded cap");
        self.bytes[self.len..end].copy_from_slice(bytes);
        self.len = end;
    }

    fn as_slice(&self) -> &[u8] {
        &self.bytes[..self.len]
    }
}

struct TableStorageDecoder {
    case: u8,
}

impl TableStorageDecoder {
    fn new(case: u8) -> Self {
        Self { case }
    }

    fn storage_options(&self, source_len: usize) -> storage::DecodeOptions {
        let mut max_message_bytes = MAX_SOURCE_BYTES;
        let mut max_fields = MAX_FIELDS;
        let mut max_work_bytes = MAX_WORK_BYTES;
        let mut max_references = MAX_REFERENCES;
        let mut max_text_bytes = MAX_TEXT_BYTES;
        match self.case {
            BYTES_LIMIT_CASE => max_message_bytes = source_len.saturating_sub(1),
            FIELDS_LIMIT_CASE => max_fields = 1,
            WORK_LIMIT_CASE => max_work_bytes = source_len.saturating_sub(1),
            REFERENCES_LIMIT_CASE => max_references = 0,
            TEXT_LIMIT_CASE => max_text_bytes = 0,
            _ => {},
        }
        storage::DecodeOptions::new(
            max_message_bytes,
            max_fields,
            max_work_bytes,
            MAX_RECURSION,
            max_references,
            max_text_bytes,
        )
    }
}

impl<'source> ListDecoder<'source> for TableStorageDecoder {
    type Value = u32;
    type Error = HookError;

    fn probe(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        _object_id: u64,
    ) -> Result<i32, Self::Error> {
        let options = self.storage_options(source.len());
        let list_type = match kind {
            RootOrSegment::Root => {
                let (snapshot, report) =
                    storage::decode_table_data_list_type_with_report(source, options)
                        .map_err(HookError::Codec)?;
                observe_report(report);
                snapshot.list_type()
            },
            RootOrSegment::Segment => {
                let (snapshot, report) =
                    storage::decode_table_data_list_segment_type_with_report(source, options)
                        .map_err(HookError::Codec)?;
                observe_report(report);
                snapshot.list_type()
            },
        };
        Ok(list_type)
    }

    fn decode(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        _object_id: u64,
        admit: bool,
    ) -> Result<Candidate<Self::Value, Self::Error>, Self::Error> {
        let mut collector = StorageCollector::new(source, admit);
        let options = self.storage_options(source.len());
        let (list_type, key_range) = match kind {
            RootOrSegment::Root => {
                let (snapshot, report) =
                    storage::decode_table_data_list_with_visitor(source, options, &mut collector)
                        .map_err(HookError::Codec)?;
                observe_report(report);
                (snapshot.list_type(), None)
            },
            RootOrSegment::Segment => {
                let (snapshot, report) = storage::decode_table_data_list_segment_with_visitor(
                    source,
                    options,
                    &mut collector,
                )
                .map_err(HookError::Codec)?;
                observe_report(report);
                (
                    snapshot.list_type(),
                    Some(KeyRange {
                        location: snapshot.key_range_location(),
                        length: snapshot.key_range_length(),
                    }),
                )
            },
        };
        black_box(collector.borrowed_bytes);

        Ok(Candidate {
            list_type,
            values: collector.values,
            keys: collector.keys,
            segment_refs: collector.segment_refs,
            key_range,
            entry_bounds: collector.entry_bounds,
            structural_error: None,
            semantic_error: None,
        })
    }

    fn map_issue(&mut self, issue: CoordinatorIssue) -> Self::Error {
        HookError::Coordinator(issue)
    }
}

struct StorageCollector {
    source_start: usize,
    source_end: usize,
    admit: bool,
    keys: HashSet<u32>,
    values: Vec<(u32, u32)>,
    segment_refs: Vec<u64>,
    entry_bounds: Option<EntryBounds>,
    borrowed_bytes: usize,
}

impl StorageCollector {
    fn new(source: &[u8], admit: bool) -> Self {
        let source_start = source.as_ptr() as usize;
        Self {
            source_start,
            source_end: source_start.saturating_add(source.len()),
            admit,
            keys: HashSet::new(),
            values: Vec::new(),
            segment_refs: Vec::new(),
            entry_bounds: None,
            borrowed_bytes: 0,
        }
    }

    fn observe_borrowed(&mut self, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start.saturating_add(bytes.len());
        assert!(start >= self.source_start);
        assert!(end <= self.source_end);
        self.borrowed_bytes = self.borrowed_bytes.saturating_add(bytes.len());
    }
}

impl storage::StorageVisitor for StorageCollector {
    fn visit_list_entry_record(
        &mut self,
        record: storage::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        self.observe_borrowed(record.raw());
        let entry = record.snapshot();
        if !self.keys.contains(&entry.key()) && self.keys.len() >= MAX_ENTRIES {
            return Err(storage::DecodeError::allocation(
                self.keys.len().saturating_add(1),
            ));
        }
        if self.admit && self.values.len() >= MAX_ENTRIES {
            return Err(storage::DecodeError::allocation(
                self.values.len().saturating_add(1),
            ));
        }
        if let Some(value) = entry.string_value() {
            self.observe_borrowed(value.as_bytes());
        }
        for payload in [
            entry.formula(),
            entry.format(),
            entry.custom_format(),
            entry.import_warning_set(),
            entry.cell_spec(),
        ]
        .into_iter()
        .flatten()
        {
            self.observe_borrowed(payload);
        }
        let key = entry.key();
        self.keys.insert(key);
        self.entry_bounds = Some(match self.entry_bounds {
            Some(bounds) => EntryBounds {
                minimum: bounds.minimum.min(key),
                maximum: bounds.maximum.max(key),
            },
            None => EntryBounds {
                minimum: key,
                maximum: key,
            },
        });
        if self.admit {
            self.values.push((key, key));
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        record: storage::ReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if self.segment_refs.len() >= MAX_SEGMENTS {
            return Err(storage::DecodeError::allocation(
                self.segment_refs.len().saturating_add(1),
            ));
        }
        self.observe_borrowed(record.raw());
        self.segment_refs.push(record.reference().identifier());
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let bounded = &data[..data.len().min(MAX_SOURCE_BYTES)];
    exercise_raw(bounded);

    let selected = control_byte(data, 0) % CASE_COUNT;
    let root = root_fixture(data, selected);
    let expected_type = list_type(data);
    let segment = segment_fixture(data, selected, expected_type);
    exercise_coordinator(data, selected, &root, &segment, expected_type);
});

fn exercise_raw(source: &[u8]) {
    let mut decoder = TableStorageDecoder::new(VALID_CASE);
    let roots = [Message::new(TABLE_DATA_LIST_MESSAGE_KIND, source)];
    let result = read_list(
        0x1001,
        1,
        roots,
        no_segments,
        &mut decoder,
        ListReadPolicy::default(),
    );
    observe_result(result);

    let mut segment_decoder = TableStorageDecoder::new(VALID_CASE);
    observe_result(segment_decoder.probe(source, RootOrSegment::Segment, 0x1002));
    observe_result(segment_decoder.decode(source, RootOrSegment::Segment, 0x1002, false));
}

fn no_segments<'source>(_segment_id: u64) -> Result<Option<[Message<'source>; 1]>, HookError> {
    Ok(None)
}

fn exercise_coordinator(
    data: &[u8],
    case: u8,
    root: &FixedBytes,
    segment: &FixedBytes,
    expected_type: i32,
) {
    let root_kind = if case == WRONG_ROOT_KIND_CASE {
        0
    } else if case == NATIVE_ROOT_CASE {
        NATIVE_TABLE_DATA_LIST_MESSAGE_KIND
    } else {
        TABLE_DATA_LIST_MESSAGE_KIND
    };
    let root_messages = [
        Message::new(root_kind, root.as_slice()),
        if case == DUPLICATE_ROOT_CASE {
            Message::new(root_kind, root.as_slice())
        } else {
            Message::new(0, &[])
        },
    ];
    let segment_kind = if case == MISSING_SEGMENT_PAYLOAD_CASE {
        0
    } else {
        TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND
    };
    let duplicate_payload = case == DUPLICATE_SEGMENT_PAYLOAD_CASE;
    let missing_segment = case == MISSING_SEGMENT_CASE;
    let segment_data = segment.as_slice();
    let mut decoder = TableStorageDecoder::new(case);
    let mut policy = ListReadPolicy::default();
    if matches!(case, POLICY_IMMEDIATE_CASE | POLICY_DEFERRED_CASE) {
        policy.max_entries = 1;
        policy.overflow = if case == POLICY_IMMEDIATE_CASE {
            OverflowPolicy::Immediate
        } else {
            OverflowPolicy::Deferred
        };
    }
    if case == OVERFLOW_KEY_RANGE_CASE {
        policy.range_overflow = if control_byte(data, 79) & 1 == 0 {
            OverflowPolicy::Immediate
        } else {
            OverflowPolicy::Deferred
        };
    }
    let result = read_list(
        u64::from(control_u32(data, 80) | 1),
        expected_type,
        root_messages,
        |segment_id| {
            if missing_segment {
                return Ok(None);
            }
            let first = Message::new(segment_kind, segment_data);
            let second = if duplicate_payload {
                Message::new(segment_kind, segment_data)
            } else {
                Message::new(0, &[])
            };
            black_box(segment_id);
            Ok(Some([first, second]))
        },
        &mut decoder,
        policy,
    );
    observe_result(result);
}

fn observe_result<T>(result: Result<T, HookError>) {
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => match error {
            HookError::Codec(error) => {
                black_box(error);
            },
            HookError::Coordinator(issue) => {
                black_box(issue);
            },
        },
    }
}

fn observe_report(report: storage::DecodeReport) {
    black_box((
        report.source_bytes(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
        report.references(),
        report.reference_bytes(),
        report.text_bytes(),
    ));
}

fn root_fixture(data: &[u8], case: u8) -> FixedBytes {
    let mut output = FixedBytes::new();
    let list_type = list_type(data);
    put_varint_field(&mut output, 1, encode_int32(list_type));
    put_varint_field(&mut output, 2, u64::from(control_u32(data, 4) | 1));
    if case == DUPLICATE_FIELD_CASE {
        put_varint_field(&mut output, 1, encode_int32(list_type));
    }
    if control_byte(data, 8) & 1 != 0 {
        put_varint_field(&mut output, 5, u64::from(control_byte(data, 9) & 1));
    }

    let entries = entry_count(data, case);
    for index in 0..entries {
        if case == WRONG_WIRE_CASE && index == 0 {
            put_varint_field(&mut output, 3, 1);
            continue;
        }
        let entry = if case == MALFORMED_LATER_CASE && index + 1 == entries {
            malformed_entry(index)
        } else {
            entry_fixture(data, index, case, true)
        };
        put_bytes_field(&mut output, 3, entry.as_slice());
    }

    let segments = segment_count(data, case);
    for index in 0..segments {
        let identifier = if case == DUPLICATE_SEGMENT_REFERENCE_CASE {
            701
        } else {
            701_u64.saturating_add(index as u64)
        };
        let reference = reference_fixture(data, identifier, index + 24);
        put_bytes_field(&mut output, 4, reference.as_slice());
    }
    output
}

fn segment_fixture(data: &[u8], case: u8, expected_type: i32) -> FixedBytes {
    let mut output = FixedBytes::new();
    let list_type = if case == WRONG_SEGMENT_TYPE_CASE {
        if expected_type == 12 {
            1
        } else {
            expected_type.saturating_add(1)
        }
    } else {
        expected_type
    };
    put_varint_field(&mut output, 1, encode_int32(list_type));
    let range = range_fixture(data, case, segment_entry_count(data, case));
    if case == WRONG_WIRE_CASE {
        put_varint_field(&mut output, 2, 1);
    } else {
        put_bytes_field(&mut output, 2, range.as_slice());
        if case == DUPLICATE_FIELD_CASE {
            put_bytes_field(&mut output, 2, range.as_slice());
        }
    }

    let entries = segment_entry_count(data, case);
    for index in 0..entries {
        if case == WRONG_WIRE_CASE && index == 0 {
            put_varint_field(&mut output, 3, 1);
            continue;
        }
        let entry = if case == MALFORMED_LATER_CASE && index + 1 == entries {
            malformed_entry(index)
        } else {
            entry_fixture(data, index, case, false)
        };
        put_bytes_field(&mut output, 3, entry.as_slice());
    }
    output
}

fn entry_fixture(data: &[u8], index: usize, case: u8, root: bool) -> FixedBytes {
    let mut output = FixedBytes::new();
    let key = if case == DUPLICATE_ENTRY_KEY_CASE {
        1
    } else if root {
        u32::try_from(index + 1).unwrap_or(u32::MAX)
    } else {
        100_u32.saturating_add(u32::try_from(index).unwrap_or(u32::MAX))
    };
    put_varint_field(&mut output, 1, u64::from(key));
    put_varint_field(
        &mut output,
        2,
        u64::from(control_byte(data, index + 20) % 4),
    );

    let payload_kind = if case == TEXT_LIMIT_CASE {
        0
    } else if case == REFERENCES_LIMIT_CASE {
        1
    } else {
        (usize::from(control_byte(data, index + 24)) + index) % 9
    };
    match payload_kind {
        0 => put_bytes_field(&mut output, 3, b"cell"),
        1 => {
            let reference = reference_fixture(
                data,
                u64::try_from(40 + index).unwrap_or(u64::MAX),
                32 + index,
            );
            put_bytes_field(&mut output, 4, reference.as_slice());
        },
        2 => put_bytes_field(&mut output, 5, &[]),
        3 => put_bytes_field(&mut output, 6, &[]),
        4 => put_bytes_field(&mut output, 8, &[]),
        5 => {
            let reference = reference_fixture(
                data,
                u64::try_from(50 + index).unwrap_or(u64::MAX),
                40 + index,
            );
            put_bytes_field(&mut output, 9, reference.as_slice());
        },
        6 => {
            let reference = reference_fixture(
                data,
                u64::try_from(60 + index).unwrap_or(u64::MAX),
                48 + index,
            );
            put_bytes_field(&mut output, 10, reference.as_slice());
        },
        7 => put_bytes_field(&mut output, 11, &[]),
        _ => put_bytes_field(&mut output, 12, &[]),
    }
    if case == DUPLICATE_FIELD_CASE && index == 0 {
        put_bytes_field(&mut output, 3, b"again");
    }
    output
}

fn malformed_entry(index: usize) -> FixedBytes {
    let mut output = FixedBytes::new();
    put_varint_field(
        &mut output,
        1,
        u64::from(u32::try_from(index + 1).unwrap_or(u32::MAX)),
    );
    output
}

fn reference_fixture(data: &[u8], identifier: u64, offset: usize) -> FixedBytes {
    let mut output = FixedBytes::new();
    let controlled = if control_byte(data, offset) & 1 == 0 {
        identifier
    } else {
        identifier.saturating_add(u64::from(control_byte(data, offset)))
    } | 1;
    put_varint_field(&mut output, 1, controlled);
    if control_byte(data, offset + 1) & 1 != 0 {
        put_varint_field(
            &mut output,
            2,
            u64::from(control_byte(data, offset + 2) % 4),
        );
    }
    output
}

fn range_fixture(data: &[u8], case: u8, segment_entries: usize) -> FixedBytes {
    let mut output = FixedBytes::new();
    let (location, length) = match case {
        OUTSIDE_KEY_RANGE_CASE => (1_000, 1),
        OVERFLOW_KEY_RANGE_CASE => (u32::MAX - 1, 4),
        DUPLICATE_ENTRY_KEY_CASE => (1, 1),
        _ => (
            100,
            u32::try_from(segment_entries.max(1)).unwrap_or(u32::MAX),
        ),
    };
    put_varint_field(&mut output, 1, u64::from(location));
    if case != MALFORMED_KEY_RANGE_CASE {
        put_varint_field(&mut output, 2, u64::from(length));
    }
    if case == DUPLICATE_FIELD_CASE {
        put_varint_field(&mut output, 1, u64::from(control_u32(data, 76)));
    }
    output
}

fn entry_count(data: &[u8], case: u8) -> usize {
    match case {
        MAX_CASE => MAX_ENTRIES,
        MALFORMED_LATER_CASE => 2 + usize::from(control_byte(data, 70) % 4),
        DUPLICATE_ENTRY_KEY_CASE | POLICY_IMMEDIATE_CASE | POLICY_DEFERRED_CASE => 3,
        _ => 2 + usize::from(control_byte(data, 70) % 3),
    }
}

fn segment_entry_count(data: &[u8], case: u8) -> usize {
    match case {
        MAX_CASE => MAX_ENTRIES,
        MALFORMED_LATER_CASE => 2 + usize::from(control_byte(data, 71) % 4),
        DUPLICATE_ENTRY_KEY_CASE | POLICY_IMMEDIATE_CASE | POLICY_DEFERRED_CASE => 2,
        _ => 1 + usize::from(control_byte(data, 71) % 3),
    }
}

fn segment_count(data: &[u8], case: u8) -> usize {
    match case {
        MAX_CASE => MAX_SEGMENTS,
        DUPLICATE_SEGMENT_REFERENCE_CASE => 2,
        _ => 1 + usize::from(control_byte(data, 72) % 2),
    }
}

fn list_type(data: &[u8]) -> i32 {
    1 + i32::from(control_byte(data, 73) % 12)
}

fn encode_int32(value: i32) -> u64 {
    if value < 0 {
        u64::from_ne_bytes(i64::from(value).to_ne_bytes())
    } else {
        u64::try_from(value).unwrap_or(0)
    }
}

fn put_varint_field(output: &mut FixedBytes, number: u32, value: u64) {
    put_key(output, number, 0);
    put_varint(output, value);
}

fn put_bytes_field(output: &mut FixedBytes, number: u32, value: &[u8]) {
    put_key(output, number, 2);
    put_varint(output, u64::try_from(value.len()).unwrap_or(u64::MAX));
    output.extend(value);
}

fn put_key(output: &mut FixedBytes, number: u32, wire_type: u8) {
    put_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn put_varint(output: &mut FixedBytes, mut value: u64) {
    loop {
        let mut byte = u8::try_from(value & 0x7f).unwrap_or(0);
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

fn control_byte(data: &[u8], offset: usize) -> u8 {
    data.get(offset)
        .copied()
        .unwrap_or((offset as u8).wrapping_mul(29).wrapping_add(7))
}

fn control_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        control_byte(data, offset),
        control_byte(data, offset + 1),
        control_byte(data, offset + 2),
        control_byte(data, offset + 3),
    ])
}
