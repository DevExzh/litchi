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
