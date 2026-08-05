use std::io::Cursor;
use std::path::PathBuf;

use litchi_cfb::{OleFile, OleWriter};

use super::codec::parse_properties_with_limits;
use super::*;

const ITEM_ID: &str = "{E0FCA697-D525-4175-A08B-DFD1F1FC7C9F}";

fn properties() -> Properties {
    Properties {
        item_id: ITEM_ID.parse().unwrap(),
        schema_references: vec![
            "http://schemas.openxmlformats.org/officeDocument/2006/bibliography".to_string(),
        ],
    }
}

#[test]
fn property_xml_and_compact_storage_name_round_trip() {
    let value = properties();
    let xml = write_properties(&value).unwrap();
    assert_eq!(parse_properties(&xml).unwrap(), value);
    assert_eq!(value.item_id.to_string(), ITEM_ID);
    assert_eq!(value.item_id.storage_name().len(), 26);
    assert!(
        value
            .item_id
            .storage_name()
            .bytes()
            .all(|byte| { byte.is_ascii_uppercase() || byte.is_ascii_digit() })
    );
}

#[test]
fn synthetic_store_round_trips_with_modified_promotion() {
    let item = Item::new(
        properties().item_id.storage_name(),
        br#"<?xml version="1.0"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>"#
            .to_vec(),
        properties(),
    )
    .unwrap();
    let store = Store::new(Promotion::Modified, vec![item]).unwrap();
    let mut writer = OleWriter::new();
    write(&mut writer, &store).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
    assert_eq!(inspect(&mut ole).unwrap().unwrap(), store);
}

#[test]
fn reads_real_word_custom_xml_store() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/inline-endnote-and-footnote.doc");
    let mut ole = OleFile::open(std::fs::File::open(path).unwrap()).unwrap();
    let store = inspect(&mut ole).unwrap().unwrap();
    assert_eq!(store.promotion, Promotion::Unspecified);
    assert_eq!(store.items().len(), 1);
    assert_eq!(store.items()[0].properties().item_id.to_string(), ITEM_ID);
    assert_eq!(store.items()[0].root_name().local_name, "Sources");
}

#[test]
fn rejects_conflicting_markers_bad_shape_and_hostile_xml() {
    let mut writer = OleWriter::new();
    writer.create_storage(&[STORE_STORAGE]).unwrap();
    writer
        .create_storage(&[REDUNDANT_PROMOTION_STORAGE])
        .unwrap();
    writer
        .create_storage(&[MODIFIED_PROMOTION_STORAGE])
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
    assert!(inspect(&mut ole).is_err());

    assert!(Item::new("item", b"<!DOCTYPE x><x/>".to_vec(), properties()).is_err());
    assert!(
        parse_properties(
            br#"<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{E0FCA697-D525-4175-A08B-DFD1F1FC7C9F}"><ds:bad/></ds:datastoreItem>"#
        )
        .is_err()
    );
    assert!(Item::new("item", b"<x>\0</x>".to_vec(), properties()).is_err());
    let mut invalid_properties = properties();
    invalid_properties.schema_references = vec!["urn:\0test".to_string()];
    assert!(write_properties(&invalid_properties).is_err());
}

#[test]
fn zero_count_limits_can_disable_items_and_schema_references() {
    let mut writer = OleWriter::new();
    writer.create_storage(&[STORE_STORAGE]).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
    let limits = Limits {
        max_items: 0,
        max_schema_references: 0,
        ..Limits::default()
    };
    assert!(
        inspect_with_limits(&mut ole, limits)
            .unwrap()
            .unwrap()
            .items()
            .is_empty()
    );

    let xml = write_properties(&properties()).unwrap();
    assert!(parse_properties_with_limits(&xml, &limits).is_err());
}

#[test]
fn accepts_and_preserves_utf16_xml_streams() {
    fn utf16(value: &str, little_endian: bool) -> Vec<u8> {
        let mut output = if little_endian {
            vec![0xFF, 0xFE]
        } else {
            vec![0xFE, 0xFF]
        };
        for unit in value.encode_utf16() {
            let bytes = if little_endian {
                unit.to_le_bytes()
            } else {
                unit.to_be_bytes()
            };
            output.extend_from_slice(&bytes);
        }
        output
    }

    let properties_xml = utf16(
        &String::from_utf8(write_properties(&properties()).unwrap()).unwrap(),
        true,
    );
    assert_eq!(parse_properties(&properties_xml).unwrap(), properties());

    let item_xml = utf16(
        r#"<?xml version="1.0" encoding="UTF-16"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>"#,
        false,
    );
    let item = Item::new("UTF16ITEM", item_xml.clone(), properties()).unwrap();
    assert_eq!(item.xml(), item_xml);
    assert_eq!(item.root_name().local_name, "Sources");
}
