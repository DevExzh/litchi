use super::{
    Editor, FileInformationBlock, Limits, SmartTagBookmarkInfo, SmartTagOrigin,
    SmartTagRecognizerState, Snapshot, TableKind,
};
use litchi_codepage::Ansi;
use litchi_ole_common::smart_tags::{
    Property, PropertyBag, PropertyBagStore, PropertyBagString, PropertyBagStringEncoding, Type,
};

fn property_string(value: &str) -> PropertyBagString {
    PropertyBagString {
        value: value.to_owned(),
        encoding: PropertyBagStringEncoding::Ansi,
    }
}

fn source() -> (FileInformationBlock, Vec<u8>) {
    let ansi = Ansi::WINDOWS_1252;
    let store = PropertyBagStore {
        ansi,
        reserved_factoid_count: 9,
        types: vec![Type {
            id: 7,
            namespace_uri: property_string("urn:test"),
            tag_name: property_string("place"),
            download_url: property_string(""),
        }],
        strings: vec![property_string("key"), property_string("v1")],
    };
    let bags = vec![PropertyBag {
        type_id: 7,
        properties: vec![Property {
            key_index: 0,
            value_index: 1,
        }],
    }];
    let factoid_data = store.to_bytes_with_bags(&bags).unwrap();

    let mut infos = vec![0xff, 0xff, 1, 0, 0, 0];
    infos.extend_from_slice(&6u16.to_le_bytes());
    infos.extend_from_slice(&41u32.to_le_bytes());
    infos.extend_from_slice(&0u16.to_le_bytes());
    infos.extend_from_slice(&2u16.to_le_bytes());
    infos.extend_from_slice(&0xdead_beefu32.to_le_bytes());

    let mut starts = Vec::new();
    starts.extend_from_slice(&2u32.to_le_bytes());
    starts.extend_from_slice(&20u32.to_le_bytes());
    starts.extend_from_slice(&0u16.to_le_bytes());
    starts.extend_from_slice(&0x0033u16.to_le_bytes());
    starts.extend_from_slice(&1u16.to_le_bytes());

    let mut ends = Vec::new();
    ends.extend_from_slice(&10u32.to_le_bytes());
    ends.extend_from_slice(&20u32.to_le_bytes());
    ends.extend_from_slice(&0u16.to_le_bytes());
    ends.extend_from_slice(&0u16.to_le_bytes());

    let mut recognizer = Vec::new();
    recognizer.extend_from_slice(&0u32.to_le_bytes());
    recognizer.extend_from_slice(&20u32.to_le_bytes());
    recognizer.extend_from_slice(&7u16.to_le_bytes());

    let mut table = vec![0xa1; 7];
    let unknown_prefix_len = 3usize;
    let info_offset = table.len();
    table.extend_from_slice(&infos);
    table.extend_from_slice(&[0xb2; 5]);
    let starts_offset = table.len();
    table.extend_from_slice(&starts);
    table.extend_from_slice(&[0xc3; 4]);
    let ends_offset = table.len();
    table.extend_from_slice(&ends);
    table.extend_from_slice(&[0xd4; 6]);
    let factoid_offset = table.len();
    table.extend_from_slice(&factoid_data);
    table.extend_from_slice(&[0xe5; 2]);
    let recognizer_offset = table.len();
    table.extend_from_slice(&recognizer);

    let mut fib_bytes = vec![0u8; 154 + 136 * 8];
    fib_bytes[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    fib_bytes[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    fib_bytes[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
    fib_bytes[10..12].copy_from_slice(&0x0200u16.to_le_bytes());
    fib_bytes[76..80].copy_from_slice(&20u32.to_le_bytes());
    fib_bytes[152..154].copy_from_slice(&136u16.to_le_bytes());

    let pointers = [
        (20, 0, unknown_prefix_len),
        (114, info_offset, infos.len()),
        (115, starts_offset, starts.len()),
        (117, ends_offset, ends.len()),
        (118, factoid_offset, factoid_data.len()),
        (132, recognizer_offset, recognizer.len()),
    ];
    for &(index, offset, length) in &pointers {
        let pointer = 154 + index * 8;
        fib_bytes[pointer..pointer + 4].copy_from_slice(&(offset as u32).to_le_bytes());
        fib_bytes[pointer + 4..pointer + 8].copy_from_slice(&(length as u32).to_le_bytes());
    }
    (FileInformationBlock::parse(&fib_bytes).unwrap(), table)
}

#[test]
fn edits_are_source_checked_and_preserve_opaque_table_bytes() {
    let (fib, table) = source();
    let snapshot = Snapshot::parse_with(&fib, &table, Ansi::WINDOWS_1252).unwrap();
    assert_eq!(snapshot.tags().len(), 1);
    assert_eq!(
        snapshot.recognizer_ranges()[0].state,
        SmartTagRecognizerState::Clean
    );
    assert_eq!(snapshot.store().unwrap().reserved_factoid_count, 9);

    let mut transaction = snapshot.edit();
    transaction
        .set_bookmark_info(
            0,
            SmartTagBookmarkInfo {
                id: 99,
                is_sub_entity: true,
                origin: SmartTagOrigin::GrammarChecker,
            },
        )
        .unwrap();
    transaction
        .set_property_value(0, 0, property_string("v2"))
        .unwrap();
    transaction
        .set_recognizer_state(0, SmartTagRecognizerState::Dirty)
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_noop());

    let edited = commit.snapshot();
    assert_eq!(edited.fib_bytes(), snapshot.fib_bytes());
    assert_eq!(edited.topology(), snapshot.topology());
    assert_eq!(edited.table_stream()[0..7], table[0..7]);
    assert_eq!(
        edited.table_stream()[edited
            .topology()
            .range(TableKind::BookmarkInfo)
            .unwrap()
            .end()
            .unwrap() as usize..][0..5],
        [0xb2; 5]
    );
    let info = snapshot.topology().range(TableKind::BookmarkInfo).unwrap();
    let info_start = info.offset as usize;
    assert_eq!(
        &edited.table_stream()[info_start + 16..info_start + 20],
        &table[info_start + 16..info_start + 20]
    );
    assert_eq!(
        edited.metadata().tags[0].property_bag,
        snapshot.metadata().tags[0].property_bag
    );
    assert_eq!(edited.store().unwrap().string(1), Some("v2"));
    assert_eq!(
        edited.recognizer_ranges()[0].state,
        SmartTagRecognizerState::Dirty
    );

    let restored = commit.patch().inverse().apply(edited).unwrap();
    assert_eq!(restored, snapshot);

    let mut stale_table = table.clone();
    stale_table[0] ^= 1;
    let stale = Snapshot::parse_with(&fib, &stale_table, Ansi::WINDOWS_1252).unwrap();
    assert!(commit.patch().apply(&stale).is_err());
}

#[test]
fn no_op_and_failed_edits_are_atomic() {
    let (fib, table) = source();
    let snapshot = Snapshot::parse_with_limits(&fib, &table, Limits::default()).unwrap();
    let no_op = snapshot.edit().commit().unwrap();
    assert!(!no_op.changed());
    assert_eq!(no_op.snapshot().table_stream(), table.as_slice());

    let mut transaction = snapshot.edit();
    let before = transaction.metadata().clone();
    assert!(
        transaction
            .set_property_value(0, 0, property_string("value-too-long"))
            .is_err()
    );
    assert_eq!(transaction.metadata(), &before);

    let mut invalid = snapshot.edit();
    assert!(
        invalid
            .set_string(1, property_string("bad\0value"))
            .is_err()
    );
    assert_eq!(invalid.metadata(), snapshot.metadata());
}

#[test]
fn package_editor_publishes_only_source_checked_transactions() {
    let (fib, table) = source();
    let mut editor = Editor::open(&fib, &table).unwrap();
    let mut transaction = editor.edit();
    transaction
        .set_recognizer_state(0, SmartTagRecognizerState::Pending)
        .unwrap();
    let commit = editor.apply(transaction).unwrap();
    assert_eq!(editor.snapshot(), commit.snapshot());
    let (fib_bytes, table_bytes) = editor.finish();
    assert_eq!(fib_bytes, fib.raw_data());
    assert_eq!(table_bytes.len(), table.len());
    assert_eq!(table_bytes[0..7], table[0..7]);
    let factoid = editor
        .snapshot()
        .topology()
        .range(TableKind::PropertyBags)
        .unwrap();
    let factoid = factoid.offset as usize..factoid.end().unwrap() as usize;
    assert_eq!(&table_bytes[factoid.clone()], &table[factoid]);
}
