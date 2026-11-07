use super::package::Editor;
use super::record::{Error as TransactionError, Snapshot};
use super::{
    DeliveryOption, Metadata, NarrowString, Protection, Recipient, RecipientSelector, parse,
    parse_bytes, to_bytes,
};
use crate::SourceLimits;
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
    let bytes = to_bytes(&route_slip, &SourceLimits::default()).expect("sample serializes");
    let parsed = parse_bytes(&bytes, &SourceLimits::default()).expect("sample parses");
    assert_eq!(parsed, route_slip);
    assert_eq!(
        parsed
            .to_bytes(&SourceLimits::default())
            .expect("parsed value serializes"),
        bytes
    );
    assert_eq!(parsed.subject.as_bytes(), &[0x80, 0xff, b's', b'u', b'b']);
    assert_eq!(parsed.recipients[0].entry_id, [0, 0xff, 1]);
}

#[test]
fn parses_from_fib_table_pointer_without_document_integration() {
    let payload = to_bytes(&sample(), &SourceLimits::default()).expect("sample serializes");
    let offset = 7usize;
    let mut table_stream = vec![0xaa; offset];
    table_stream.extend_from_slice(&payload);

    let fib = fib_with_route_slip_pointer(offset, payload.len());
    assert_eq!(
        parse(&fib, &table_stream, &SourceLimits::default()).expect("FIB route slip parses"),
        Some(sample())
    );
}

#[test]
fn rejects_invalid_bool_dirty_enum_counts_and_relationships() {
    let valid = to_bytes(&sample(), &SourceLimits::default()).expect("sample serializes");

    for (offset, value) in [(0, 2u16), (2, 2), (4, 2), (6, 1)] {
        let mut bytes = valid.clone();
        bytes[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
        assert!(
            parse_bytes(&bytes, &SourceLimits::default()).is_err(),
            "field at offset {offset}"
        );
    }

    let mut invalid_stage = valid.clone();
    invalid_stage[10..12].copy_from_slice(&2i16.to_le_bytes());
    assert!(parse_bytes(&invalid_stage, &SourceLimits::default()).is_err());

    let mut invalid_delivery = valid.clone();
    invalid_delivery[12..14].copy_from_slice(&(-1i16).to_le_bytes());
    assert!(parse_bytes(&invalid_delivery, &SourceLimits::default()).is_err());

    let mut invalid_count = valid;
    invalid_count[14..16].copy_from_slice(&0i16.to_le_bytes());
    assert!(parse_bytes(&invalid_count, &SourceLimits::default()).is_err());
}

#[test]
fn rejects_truncation_trailing_bytes_and_invalid_signed_lengths() {
    let valid = to_bytes(&sample(), &SourceLimits::default()).expect("sample serializes");
    assert!(parse_bytes(&valid[..valid.len() - 1], &SourceLimits::default()).is_err());

    let mut trailing = valid.clone();
    trailing.push(0);
    assert!(parse_bytes(&trailing, &SourceLimits::default()).is_err());

    let mut negative_recipient_count = valid.clone();
    negative_recipient_count[14..16].copy_from_slice(&(-1i16).to_le_bytes());
    assert!(parse_bytes(&negative_recipient_count, &SourceLimits::default()).is_err());

    let mut empty_name = valid.clone();
    let name_length_offset = first_recipient_name_length_offset(&empty_name);
    empty_name[name_length_offset..name_length_offset + 2].copy_from_slice(&0i16.to_le_bytes());
    assert!(parse_bytes(&empty_name, &SourceLimits::default()).is_err());

    let mut oversized_subject = valid;
    oversized_subject[16..18].copy_from_slice(&256u16.to_le_bytes());
    assert!(parse_bytes(&oversized_subject, &SourceLimits::default()).is_err());
}

#[test]
fn recipient_count_is_checked_against_minimum_wire_size_before_reservation() {
    let mut bytes = vec![0u8; 16];
    bytes[14..16].copy_from_slice(&100i16.to_le_bytes());
    for _ in 0..4 {
        bytes.extend_from_slice(&0u16.to_le_bytes());
    }
    assert!(parse_bytes(&bytes, &SourceLimits::default()).is_err());
}

#[test]
fn raw_codec_reports_entry_limit_before_decoding() {
    let bytes = to_bytes(&sample(), &SourceLimits::default()).expect("sample serializes");
    let limits = SourceLimits::builder()
        .max_entry_bytes(bytes.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        parse_bytes(&bytes, &limits),
        Err(crate::package::Error::Limit {
            kind: crate::package::SourceLimitKind::EntryBytes,
            ..
        })
    ));
}

#[test]
fn raw_writer_honors_exact_and_one_under_entry_limits() {
    let value = sample();
    let bytes = to_bytes(&value, &SourceLimits::default()).unwrap();
    let exact = SourceLimits::builder()
        .max_entry_bytes(bytes.len())
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(to_bytes(&value, &exact).unwrap(), bytes);
    let one_under = SourceLimits::builder()
        .max_entry_bytes(bytes.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        to_bytes(&value, &one_under),
        Err(crate::package::Error::Limit {
            kind: crate::package::SourceLimitKind::EntryBytes,
            ..
        })
    ));
}

#[test]
fn validates_models_before_serializing() {
    let mut invalid = sample();
    invalid.stage = 2;
    assert!(invalid.to_bytes(&SourceLimits::default()).is_err());

    let mut invalid = sample();
    invalid.recipients[0].name = NarrowString::new(Vec::new());
    assert!(invalid.to_bytes(&SourceLimits::default()).is_err());

    let mut invalid = sample();
    invalid.subject = NarrowString::new(vec![0; 256]);
    assert!(invalid.to_bytes(&SourceLimits::default()).is_err());

    assert!(Recipient::parse_bytes(&[0, 0, 1, 0, b'a', 0], &SourceLimits::default()).is_err());
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
fn dropped_edit_and_commit_produce_reversible_semantic_patch() {
    let source = Snapshot::new(editable_sample()).unwrap();
    let mut transaction = source.edit();
    transaction.set_stage(RecipientSelector::Index(0)).unwrap();
    assert_ne!(transaction.staged(), &source);
    // Dropping an edit discards it; a fresh one starts from the source.
    drop(transaction);
    let mut transaction = source.edit();
    assert_eq!(transaction.staged(), &source);

    transaction.set_stage(RecipientSelector::Index(1)).unwrap();
    let committed = transaction.commit().unwrap();
    assert_eq!(committed.package().metadata().unwrap().stage, 1);
    assert_eq!(
        committed
            .patch()
            .inverse()
            .apply(committed.package())
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
fn protected_route_metadata_rejects_package_replacement_but_keeps_noops() {
    let original = write_base_doc(b"retained table");
    let mut editor = Editor::open(original).unwrap();
    let value = sample();
    let bytes = editor
        .set(value.clone())
        .unwrap()
        .snapshot()
        .finish()
        .unwrap();
    let mut editor = Editor::open(bytes.clone()).unwrap();
    editor.set(value.clone()).unwrap();
    editor.replace(value.clone()).unwrap();
    assert!(!editor.is_changed());
    let mut changed = value.clone();
    changed.protection = Protection::Off;
    for result in [editor.set(changed.clone()), editor.replace(changed)] {
        assert!(
            matches!(result, Err(crate::Error::InvalidFormat(message)) if message == "route-slip transaction failed: route-slip metadata is protected by Annotation")
        );
    }
    assert_eq!(editor.metadata(), Some(&value));
    assert_eq!(editor.finish().unwrap(), bytes);
}

#[test]
fn package_lifecycle_appends_route_data_and_clears_only_fib_pointer() {
    let original_table = b"unrelated table bytes";
    let original = write_base_doc(original_table);
    let mut editor = Editor::open(original.clone()).unwrap();
    let metadata = editable_sample();
    let committed = editor.set(metadata.clone()).unwrap();
    let routed = committed.snapshot().finish().unwrap();

    let mut ole = litchi_cfb::OleFile::open(
        Cursor::new(routed.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
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
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(finished), litchi_cfb::OleFileLimits::default())
            .unwrap();
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
    const DOP_INDEX: usize = 31;
    const POINTER_COUNT: usize = 136;
    let pointer_end = 154 + POINTER_COUNT * 8;
    let mut word = vec![0; pointer_end + 4];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    // FibBase.csw and cslw are fixed MS-DOC counts.
    word[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    word[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&(POINTER_COUNT as u16).to_le_bytes());
    word[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    word[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());

    let mut table = table.to_vec();
    let dop_offset = table.len();
    table.extend_from_slice(
        &crate::parts::document_properties::DocumentProperties::writer_bytes(
            false, false, false, true,
        ),
    );
    let dop_pointer = 154 + DOP_INDEX * 8;
    word[dop_pointer..dop_pointer + 4].copy_from_slice(&(dop_offset as u32).to_le_bytes());
    word[dop_pointer + 4..dop_pointer + 8].copy_from_slice(&594u32.to_le_bytes());

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], &table).unwrap();
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
    const POINTER_COUNT: usize = 136;
    let pointer_offset = 154 + 70 * 8;
    let pointer_end = 154 + POINTER_COUNT * 8;
    let mut data = vec![0; pointer_end + 4];
    data[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    // FibBase.csw and cslw are fixed MS-DOC counts.
    data[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    data[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    data[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    data[152..154].copy_from_slice(&(POINTER_COUNT as u16).to_le_bytes());
    data[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    data[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());
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
