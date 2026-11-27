//! Focused generated-free segment ingress coverage.
//!
//! These tests pass a segment payload directly to the codec.  They deliberately
//! do not construct an archive object or assign a native message identifier:
//! the codec validates the bytes supplied by its caller, while object routing
//! remains the responsibility of the Numbers package adapter.

use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    DecodeError, DecodeOptions, StorageVisitor, TableDataListEntrySnapshot,
    decode_table_data_list_segment_with_visitor,
};

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        usize::MAX,
        usize::MAX,
        64,
        usize::MAX,
        usize::MAX,
    )
}

fn push_varint(bytes: &mut Vec<u8>, field: u32, mut value: u64) {
    push_varint_value(bytes, u64::from(field) << 3);
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn push_varint_value(bytes: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn push_bytes(bytes: &mut Vec<u8>, field: u32, payload: &[u8]) {
    push_varint_value(bytes, (u64::from(field) << 3) | 2);
    push_varint_value(
        bytes,
        u64::try_from(payload.len()).expect("test payload length fits in u64"),
    );
    bytes.extend_from_slice(payload);
}

fn range(location: u32, length: u32) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_varint(&mut bytes, 1, u64::from(location));
    push_varint(&mut bytes, 2, u64::from(length));
    bytes
}

fn entry(key: u32) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_varint(&mut bytes, 1, u64::from(key));
    push_varint(&mut bytes, 2, 1);
    bytes
}

fn segment(entries: &[Vec<u8>]) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_varint(&mut bytes, 1, 7);
    let key_range = range(10, 2);
    push_bytes(&mut bytes, 2, &key_range);
    for value in entries {
        push_bytes(&mut bytes, 3, value);
    }
    bytes
}

#[derive(Default)]
struct EntryCollector {
    keys: Vec<u32>,
}

impl StorageVisitor for EntryCollector {
    fn visit_list_entry(
        &mut self,
        entry: TableDataListEntrySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        self.keys.push(entry.key());
        Ok(())
    }
}

fn decode_transactionally(
    source: &[u8],
    published: &mut Option<Vec<u32>>,
) -> Result<(), DecodeError> {
    let mut staged = EntryCollector::default();
    let decoded = decode_table_data_list_segment_with_visitor(source, options(source), &mut staged);
    if decoded.is_ok() {
        *published = Some(staged.keys);
    }
    decoded.map(|_| ())
}

#[test]
fn segment_projection_preserves_source_and_projects_entries() {
    let source = segment(&[entry(10), entry(11)]);
    let before = source.clone();
    let mut visitor = EntryCollector::default();
    let (snapshot, report) =
        decode_table_data_list_segment_with_visitor(&source, options(&source), &mut visitor)
            .expect("synthetic segment should decode");

    assert_eq!(source, before);
    assert_eq!(snapshot.list_type(), 7);
    assert_eq!(snapshot.key_range_location(), 10);
    assert_eq!(snapshot.key_range_length(), 2);
    assert_eq!(visitor.keys, vec![10, 11]);
    assert!(report.fields() > 0);
    assert!(report.work_bytes() >= source.len().saturating_mul(2));
}

#[test]
fn segment_projection_rejects_duplicate_required_fields_without_mutation() {
    let mut duplicate_type = segment(&[]);
    push_varint(&mut duplicate_type, 1, 8);
    let before_type = duplicate_type.clone();
    assert!(decode_transactionally(&duplicate_type, &mut None).is_err());
    assert_eq!(duplicate_type, before_type);

    let mut duplicate_range = segment(&[]);
    let key_range = range(10, 2);
    push_bytes(&mut duplicate_range, 2, &key_range);
    let before_range = duplicate_range.clone();
    assert!(decode_transactionally(&duplicate_range, &mut None).is_err());
    assert_eq!(duplicate_range, before_range);
}

#[test]
fn segment_projection_stages_callbacks_until_complete_wire_validation() {
    let mut malformed_entry = entry(11);
    push_varint(&mut malformed_entry, 1, 12);
    let source = segment(&[entry(10), malformed_entry]);
    let before = source.clone();
    let mut published = None;

    assert!(decode_transactionally(&source, &mut published).is_err());
    assert!(published.is_none());
    assert_eq!(source, before);
}
