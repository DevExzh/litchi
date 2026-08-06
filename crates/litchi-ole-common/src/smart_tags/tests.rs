use super::*;

use litchi_codepage::Ansi;

fn ansi(value: &[u8]) -> Vec<u8> {
    let mut data = (0x8000 | u16::try_from(value.len()).unwrap())
        .to_le_bytes()
        .to_vec();
    data.extend_from_slice(value);
    data
}

fn store_and_bag() -> Vec<u8> {
    let mut factoid = 7u32.to_le_bytes().to_vec();
    factoid.extend(ansi(b"urn:test"));
    factoid.extend(ansi(b"place"));
    factoid.extend(ansi(b""));
    let mut data = 1u32.to_le_bytes().to_vec();
    data.extend_from_slice(&(factoid.len() as u32).to_le_bytes());
    data.extend(factoid);
    data.extend_from_slice(&0x000cu16.to_le_bytes());
    data.extend_from_slice(&0x0100u16.to_le_bytes());
    data.extend_from_slice(&99u32.to_le_bytes());
    data.extend_from_slice(&2u32.to_le_bytes());
    data.extend(ansi(b"city"));
    data.extend(ansi(b"Paris"));
    data.extend_from_slice(&7u16.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&1u32.to_le_bytes());
    data
}

#[test]
fn parses_store_prefix_and_resolves_indexed_bag() {
    let data = store_and_bag();
    let (store, consumed) =
        PropertyBagStore::parse_prefix(&data, Ansi::WINDOWS_1252, Limits::default()).unwrap();
    let bags = store
        .parse_bags_to_end(&data[consumed..], Limits::default())
        .unwrap();
    assert_eq!(store.reserved_factoid_count, 99);
    assert_eq!(store.types[0].tag_name.value, "place");
    assert_eq!(bags[0].type_id, 7);
    assert_eq!(
        store.resolve_property(bags[0].properties[0]),
        Some(("city", "Paris"))
    );
}

#[test]
fn rejects_bad_counts_references_reserved_fields_and_truncation() {
    let data = store_and_bag();
    let mut huge = data.clone();
    huge[..4].copy_from_slice(&u32::MAX.to_le_bytes());
    assert!(PropertyBagStore::parse_prefix(&huge, Ansi::WINDOWS_1252, Limits::default(),).is_err());

    let (store, consumed) =
        PropertyBagStore::parse_prefix(&data, Ansi::WINDOWS_1252, Limits::default()).unwrap();
    let mut bag = data[consumed..].to_vec();
    bag[4..6].copy_from_slice(&1u16.to_le_bytes());
    assert!(store.parse_bags_to_end(&bag, Limits::default()).is_err());
    let mut bag = data[consumed..].to_vec();
    bag[6..10].copy_from_slice(&2u32.to_le_bytes());
    assert!(store.parse_bags_to_end(&bag, Limits::default()).is_err());
    assert!(
        PropertyBagStore::parse_prefix(&data[..8], Ansi::WINDOWS_1252, Limits::default(),).is_err()
    );

    let mut embedded_nul = data.clone();
    let city = embedded_nul
        .windows(b"city".len())
        .position(|window| window == b"city")
        .expect("fixture contains city");
    embedded_nul[city + 1] = 0;
    assert!(
        PropertyBagStore::parse_prefix(&embedded_nul, Ansi::WINDOWS_1252, Limits::default(),)
            .is_err()
    );
}

#[test]
fn serializes_store_and_bags_exactly_and_refuses_lossy_ansi() {
    let data = store_and_bag();
    let (store, consumed) =
        PropertyBagStore::parse_prefix(&data, Ansi::WINDOWS_1252, Limits::default()).unwrap();
    let bags = store
        .parse_bags_to_end(&data[consumed..], Limits::default())
        .unwrap();
    assert_eq!(store.to_bytes_with_bags(&bags).unwrap(), data);

    let mut unrepresentable = store.clone();
    unrepresentable.strings[0] = PropertyBagString {
        value: "東京".to_string(),
        encoding: PropertyBagStringEncoding::Ansi,
    };
    assert!(unrepresentable.to_bytes_with_bags(&bags).is_err());
    unrepresentable.strings[0].encoding = PropertyBagStringEncoding::Utf16;
    assert!(unrepresentable.to_bytes_with_bags(&bags).is_ok());
    unrepresentable.strings[0].value = "bad\0value".to_string();
    assert!(unrepresentable.to_bytes_with_bags(&bags).is_err());
}

#[test]
fn snapshot_retains_exact_source_and_all_bags_for_a_noop() {
    let mut data = store_and_bag();
    data.extend_from_slice(&7u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());

    let snapshot = Snapshot::parse(&data, Ansi::WINDOWS_1252, Limits::default()).unwrap();
    assert_eq!(snapshot.bytes(), data.as_slice());
    assert_eq!(snapshot.bags().len(), 2);
    assert_eq!(snapshot.resolved_property(0, 0), Some(("city", "Paris")));
    assert_eq!(snapshot.store().strings[0].value, "city");

    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert!(commit.patch().change().is_none());
    assert_eq!(commit.snapshot().bytes(), data.as_slice());
    assert_eq!(commit.patch().apply(&snapshot).unwrap(), snapshot);
}

#[test]
fn typed_property_edits_preserve_unknown_strings_and_other_bags() {
    let mut data = store_and_bag();
    data.extend_from_slice(&7u16.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&1u32.to_le_bytes());

    let snapshot = Snapshot::parse(&data, Ansi::WINDOWS_1252, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .set_property_value(
            0,
            0,
            PropertyBagString {
                value: "Rome".to_string(),
                encoding: PropertyBagStringEncoding::Ansi,
            },
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().resolved_property(0, 0),
        Some(("city", "Rome"))
    );
    assert_eq!(
        commit.snapshot().resolved_property(1, 0),
        Some(("city", "Rome"))
    );
    assert_eq!(commit.snapshot().store().strings[0].value, "city");
    assert_eq!(commit.snapshot().bags().len(), 2);
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        snapshot
    );
}

#[test]
fn stale_sources_and_invalid_or_oversized_edits_are_rejected_atomically() {
    let data = store_and_bag();
    let snapshot = Snapshot::parse(&data, Ansi::WINDOWS_1252, Limits::default()).unwrap();

    let mut changed = snapshot.edit();
    changed
        .set_property_value(
            0,
            0,
            PropertyBagString {
                value: "Rome".to_string(),
                encoding: PropertyBagStringEncoding::Ansi,
            },
        )
        .unwrap();
    let changed = changed.commit().unwrap();
    assert!(changed.patch().apply(&snapshot).is_ok());
    assert!(changed.patch().apply(changed.snapshot()).is_err());

    let before_store = snapshot.store().clone();
    let before_bags = snapshot.bags().to_vec();
    let mut invalid = snapshot.edit();
    assert!(
        invalid
            .set_property_value(
                0,
                0,
                PropertyBagString {
                    value: "東京".to_string(),
                    encoding: PropertyBagStringEncoding::Ansi,
                },
            )
            .is_err()
    );
    assert!(
        invalid
            .set_property(
                0,
                99,
                Property {
                    key_index: 0,
                    value_index: 0
                }
            )
            .is_err()
    );
    assert_eq!(invalid.store(), &before_store);
    assert_eq!(invalid.bags(), before_bags.as_slice());

    let limits = Limits {
        max_bytes: data.len() + 1,
        ..Limits::default()
    };
    let bounded = Snapshot::parse(&data, Ansi::WINDOWS_1252, limits).unwrap();
    let before = (bounded.store().clone(), bounded.bags().to_vec());
    let oversized = "x".repeat(64);
    let mut oversized_edit = bounded.edit();
    assert!(
        oversized_edit
            .set_property_value(
                0,
                0,
                PropertyBagString {
                    value: oversized,
                    encoding: PropertyBagStringEncoding::Utf16,
                },
            )
            .is_err()
    );
    assert_eq!(oversized_edit.store(), &before.0);
    assert_eq!(oversized_edit.bags(), before.1.as_slice());
}

#[test]
fn snapshot_enforces_exact_bag_count_and_source_size() {
    let data = store_and_bag();
    assert!(Snapshot::parse_bags(&data, 0, Ansi::WINDOWS_1252, Limits::default()).is_err());
    let limits = Limits {
        max_bytes: data.len() - 1,
        ..Limits::default()
    };
    assert!(Snapshot::parse(&data, Ansi::WINDOWS_1252, limits).is_err());
}
