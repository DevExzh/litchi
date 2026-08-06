use super::{
    DeliveryOption, Editor, Metadata, NarrowString, Protection, Recipient, RecipientSelector,
    Snapshot, TransactionError, parse, parse_bytes, to_bytes,
};
use crate::parts::fib::FileInformationBlock;
use litchi_cfb::OleWriter;
use std::io::Cursor;

fn sample() -> Metadata {
    Metadata::try_new(
        true,
        false,
        true,
        Protection::Annotation,
        1,
        DeliveryOption::Parallel,
        NarrowString::new(vec![0x80, 0xff, b's', b'u', b'b']),
        NarrowString::new(vec![0x81, b'm', b's', b'g']),
        NarrowString::new(vec![0x82, b'o', b'k']),
        NarrowString::new(vec![0x83, b't', b'i', b't', b'l', b'e']),
        vec![
            Recipient::try_new(
                vec![0, 0xff, 1],
                NarrowString::new(vec![0x90, b'a', b'l', b'i', b'a', b's']),
            )
            .expect("sample recipient is valid"),
            Recipient::try_new(vec![], NarrowString::new(vec![0xfe, b'b']))
                .expect("sample recipient is valid"),
        ],
    )
    .expect("sample route slip is valid")
}

#[test]
fn round_trip_preserves_narrow_bytes_and_entry_ids() {
    let route_slip = sample();
    let bytes = to_bytes(&route_slip).expect("sample serializes");
    let parsed = parse_bytes(&bytes).expect("sample parses");
    assert_eq!(parsed, route_slip);
    assert_eq!(parsed.to_bytes().expect("parsed value serializes"), bytes);
    assert_eq!(parsed.subject.as_bytes(), &[0x80, 0xff, b's', b'u', b'b']);
    assert_eq!(parsed.recipients[0].entry_id, [0, 0xff, 1]);
}

#[test]
fn parses_from_fib_table_pointer_without_document_integration() {
    let payload = to_bytes(&sample()).expect("sample serializes");
    let offset = 7usize;
    let mut table_stream = vec![0xaa; offset];
    table_stream.extend_from_slice(&payload);

    let fib = fib_with_route_slip_pointer(offset, payload.len());
    assert_eq!(
        parse(&fib, &table_stream).expect("FIB route slip parses"),
        Some(sample())
    );
}

#[test]
fn rejects_invalid_bool_dirty_enum_counts_and_relationships() {
    let valid = to_bytes(&sample()).expect("sample serializes");

    for (offset, value) in [(0, 2u16), (2, 2), (4, 2), (6, 1)] {
        let mut bytes = valid.clone();
        bytes[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
        assert!(parse_bytes(&bytes).is_err(), "field at offset {offset}");
    }

    let mut invalid_stage = valid.clone();
    invalid_stage[10..12].copy_from_slice(&2i16.to_le_bytes());
    assert!(parse_bytes(&invalid_stage).is_err());

    let mut invalid_delivery = valid.clone();
    invalid_delivery[12..14].copy_from_slice(&(-1i16).to_le_bytes());
    assert!(parse_bytes(&invalid_delivery).is_err());

    let mut invalid_count = valid;
    invalid_count[14..16].copy_from_slice(&0i16.to_le_bytes());
    assert!(parse_bytes(&invalid_count).is_err());
}

#[test]
fn rejects_truncation_trailing_bytes_and_invalid_signed_lengths() {
    let valid = to_bytes(&sample()).expect("sample serializes");
    assert!(parse_bytes(&valid[..valid.len() - 1]).is_err());

    let mut trailing = valid.clone();
    trailing.push(0);
    assert!(parse_bytes(&trailing).is_err());

    let mut negative_recipient_count = valid.clone();
    negative_recipient_count[14..16].copy_from_slice(&(-1i16).to_le_bytes());
    assert!(parse_bytes(&negative_recipient_count).is_err());

    let mut empty_name = valid.clone();
    let name_length_offset = first_recipient_name_length_offset(&empty_name);
    empty_name[name_length_offset..name_length_offset + 2].copy_from_slice(&0i16.to_le_bytes());
    assert!(parse_bytes(&empty_name).is_err());

    let mut oversized_subject = valid;
    oversized_subject[16..18].copy_from_slice(&256u16.to_le_bytes());
    assert!(parse_bytes(&oversized_subject).is_err());
}

#[test]
fn validates_models_before_serializing() {
    let mut invalid = sample();
    invalid.stage = 2;
    assert!(invalid.to_bytes().is_err());

    let mut invalid = sample();
    invalid.recipients[0].name = NarrowString::new(Vec::new());
    assert!(invalid.to_bytes().is_err());

    let mut invalid = sample();
    invalid.subject = NarrowString::new(vec![0; 256]);
    assert!(invalid.to_bytes().is_err());

    assert!(Recipient::parse_bytes(&[0, 0, 1, 0, b'a', 0]).is_err());
}

#[test]
fn selectors_reject_ambiguous_names_and_out_of_bounds_indices() {
    let mut metadata = editable_sample();
    metadata.recipients[0].name = NarrowString::new(b"same".to_vec());
    metadata.recipients[1].name = NarrowString::new(b"same".to_vec());
    let mut transaction = Snapshot::new(metadata).unwrap().edit();

    assert!(matches!(
        transaction.set_stage(RecipientSelector::Name(b"same")),
        Err(TransactionError::Selection(
            super::RecipientSelectionError::AmbiguousName { matches, .. }
        )) if matches == vec![0, 1]
    ));
    assert!(matches!(
        transaction.set_stage(RecipientSelector::Index(2)),
        Err(TransactionError::Selection(
            super::RecipientSelectionError::IndexOutOfBounds { index: 2, len: 2 }
        ))
    ));
}

#[test]
fn transaction_rollback_and_commit_produce_reversible_semantic_patch() {
    let source = Snapshot::new(editable_sample()).unwrap();
    let mut transaction = source.edit();
    transaction.set_stage(RecipientSelector::Index(0)).unwrap();
    assert!(transaction.is_changed());
    transaction.rollback();
    assert!(!transaction.is_changed());

    transaction.set_stage(RecipientSelector::Index(1)).unwrap();
    let committed = transaction.commit().unwrap();
    assert_eq!(committed.snapshot().metadata().unwrap().stage, 1);
    assert_eq!(
        committed
            .patch()
            .inverse()
            .apply(committed.snapshot())
            .unwrap(),
        source
    );
}

#[test]
fn protected_route_metadata_rejects_semantic_edits() {
    let mut transaction = Snapshot::new(sample()).unwrap().edit();
    assert!(matches!(
        transaction.set_stage(RecipientSelector::Index(0)),
        Err(TransactionError::Protected(Protection::Annotation))
    ));
}

#[test]
fn package_lifecycle_appends_route_data_and_clears_only_fib_pointer() {
    let original_table = b"unrelated table bytes";
    let original = write_base_doc(original_table);
    let mut editor = Editor::open(original.clone()).unwrap();
    let metadata = editable_sample();
    let committed = editor.set(metadata.clone()).unwrap();
    let routed = committed.snapshot().finish().unwrap();

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(routed.clone())).unwrap();
    let table = ole.open_stream(&["0Table"]).unwrap();
    assert!(table.starts_with(original_table));
    assert_eq!(
        Editor::open(routed.clone()).unwrap().metadata(),
        Some(&metadata)
    );
    assert!(committed.package_patch().inverse().apply(&routed).is_ok());

    let completed = Editor::open(routed).unwrap().complete().unwrap();
    let finished = completed.snapshot().finish().unwrap();
    let reopened = Editor::open(finished.clone()).unwrap();
    assert!(reopened.metadata().is_none());
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(finished)).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    let pointer = 154 + 70 * 8;
    assert_eq!(&word[pointer..pointer + 8], &[0; 8]);
}

fn editable_sample() -> Metadata {
    let mut value = sample();
    value.protection = Protection::Off;
    value
}

fn write_base_doc(table: &[u8]) -> Vec<u8> {
    let pointer = 154 + 70 * 8;
    let mut word = vec![0; pointer + 8];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&71u16.to_le_bytes());

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn first_recipient_name_length_offset(bytes: &[u8]) -> usize {
    let mut offset = 16;
    for _ in 0..4 {
        let length = usize::from(u16::from_le_bytes([bytes[offset], bytes[offset + 1]]));
        offset += 2 + length;
    }
    offset + 2
}

fn fib_with_route_slip_pointer(offset: usize, length: usize) -> FileInformationBlock {
    let pointer_offset = 154 + 70 * 8;
    let mut data = vec![0; pointer_offset + 8];
    data[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    data[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    data[152..154].copy_from_slice(&71u16.to_le_bytes());
    data[pointer_offset..pointer_offset + 4].copy_from_slice(
        &u32::try_from(offset)
            .expect("test offset fits")
            .to_le_bytes(),
    );
    data[pointer_offset + 4..pointer_offset + 8].copy_from_slice(
        &u32::try_from(length)
            .expect("test length fits")
            .to_le_bytes(),
    );
    FileInformationBlock::parse(&data).expect("test FIB is valid")
}
