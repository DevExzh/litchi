use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;

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
    assert_eq!(
        parse_properties(
            br#"<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{E0FCA697-D525-4175-A08B-DFD1F1FC7C9F}"><ds:bad/></ds:datastoreItem>"#
        )
        .unwrap(),
        Properties {
            item_id: ITEM_ID.parse().unwrap(),
            schema_references: Vec::new(),
        }
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

fn snapshot_item(byte: u8, storage_name: &str) -> Item {
    let properties = Properties {
        item_id: ItemId::from_bytes([byte; 16]),
        schema_references: vec![format!("urn:source:{byte:02x}")],
    };
    Item::new(
        storage_name,
        format!(r#"<source xmlns="urn:source"><value>{byte}</value></source>"#).into_bytes(),
        properties,
    )
    .unwrap()
}

fn snapshot_store() -> Snapshot {
    Snapshot::from_store(
        Store::new(Promotion::Unspecified, vec![snapshot_item(1, "ITEMONE")]).unwrap(),
    )
    .unwrap()
}

#[test]
fn transaction_edits_store_item_and_properties_without_losing_opaque_bytes() {
    let mut item = snapshot_item(1, "ITEMONE");
    let source_properties = format!(
        r#"<?xml version="1.0"?><ds:datastoreItem xmlns:ds="{CUSTOM_XML_NAMESPACE}" xmlns:x="urn:future" x:marker="keep" ds:itemID="{ITEM_ID}"><!--future property--><x:future>opaque</x:future><ds:schemaRefs><ds:schemaRef ds:uri="urn:source:01"/></ds:schemaRefs></ds:datastoreItem>"#
    );
    item.properties_xml = Arc::from(source_properties.into_bytes());
    item.properties = parse_properties(item.properties_xml()).unwrap();
    let source_item_xml = item.xml().to_vec();
    let source =
        Snapshot::from_store(Store::new(Promotion::Unspecified, vec![item]).unwrap()).unwrap();

    let mut transaction = source.edit();
    transaction.set_promotion(Promotion::Modified).unwrap();
    assert!(transaction.set_storage_name(0, "RENAMED".into()).unwrap());
    assert!(
        transaction
            .set_item_id(0, ItemId::from_bytes([2; 16]))
            .unwrap()
    );
    assert!(
        transaction
            .set_schema_references(0, ["urn:edited"])
            .unwrap()
    );
    assert!(
        transaction
            .set_xml(
                0,
                br#"<source xmlns="urn:source"><future marker="retained"/></source>"#.to_vec(),
            )
            .unwrap()
    );

    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().promotion(), Promotion::Modified);
    assert_eq!(commit.snapshot().items()[0].storage_name(), "RENAMED");
    assert_eq!(
        commit.snapshot().items()[0].xml(),
        b"<source xmlns=\"urn:source\"><future marker=\"retained\"/></source>"
    );
    assert_eq!(source.items()[0].xml(), source_item_xml.as_slice());
    assert!(
        std::str::from_utf8(commit.snapshot().items()[0].properties_xml())
            .unwrap()
            .contains("future property")
    );
    assert!(
        std::str::from_utf8(commit.snapshot().items()[0].properties_xml())
            .unwrap()
            .contains("x:marker=\"keep\"")
    );
    assert!(
        std::str::from_utf8(commit.snapshot().items()[0].properties_xml())
            .unwrap()
            .contains("<x:future>opaque</x:future>")
    );
    assert!(
        std::str::from_utf8(commit.snapshot().items()[0].properties_xml())
            .unwrap()
            .contains("ds:uri=\"urn:edited\"")
    );
    assert!(
        std::str::from_utf8(commit.snapshot().items()[0].properties_xml())
            .unwrap()
            .contains("02020202-0202-0202-0202-020202020202")
    );
}

#[test]
fn no_op_commit_reuses_the_exact_source_and_patch_is_empty() {
    let source = snapshot_store();
    let commit = source.edit().commit().unwrap();

    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert!(commit.patch().change().is_none());
    assert_eq!(commit.snapshot(), &source);
    assert!(Arc::ptr_eq(
        &source.store().items[0].xml,
        &commit.snapshot().store().items[0].xml
    ));
    assert!(Arc::ptr_eq(
        &source.store().items[0].properties_xml,
        &commit.snapshot().store().items[0].properties_xml
    ));
}

#[test]
fn inverse_and_stale_patch_application_are_source_checked() {
    let source = snapshot_store();
    let mut transaction = source.edit();
    transaction
        .set_schema_references(0, ["urn:edited"])
        .unwrap();
    let commit = transaction.commit().unwrap();

    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );

    let mut stale_transaction = source.edit();
    stale_transaction
        .set_xml(
            0,
            b"<source xmlns=\"urn:source\"><stale/></source>".to_vec(),
        )
        .unwrap();
    let stale = stale_transaction.commit().unwrap().snapshot().clone();
    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(
        stale.items()[0].xml(),
        b"<source xmlns=\"urn:source\"><stale/></source>"
    );
}

#[test]
fn failed_edits_and_malformed_or_limited_candidates_are_atomic() {
    let source = snapshot_store();
    let mut transaction = source.edit();
    let before = transaction.items().to_vec();
    assert!(
        transaction
            .set_xml(0, b"<!DOCTYPE x><x/>".to_vec())
            .is_err()
    );
    assert_eq!(transaction.items(), before.as_slice());
    assert!(
        transaction
            .set_properties_xml(0, b"<not-properties/>".to_vec())
            .is_err()
    );
    assert_eq!(transaction.items(), before.as_slice());

    let duplicate = Item::new(
        "ITEMTWO",
        b"<source xmlns=\"urn:source\"><duplicate/></source>".to_vec(),
        source.items()[0].properties().clone(),
    )
    .unwrap();
    assert!(transaction.insert(duplicate).is_err());
    assert_eq!(transaction.items(), before.as_slice());

    let too_small = Limits {
        max_item_bytes: 1,
        ..Limits::default()
    };
    assert!(Snapshot::from_store_with_limits(source.store().clone(), too_small).is_err());
    let no_items = Limits {
        max_items: 0,
        ..Limits::default()
    };
    assert!(Snapshot::from_store_with_limits(source.store().clone(), no_items).is_err());

    let mut oversized_properties = source.items()[0].properties().clone();
    oversized_properties.schema_references = vec!["x".repeat(128)];
    let limited_properties = Limits {
        max_string_bytes: 64,
        ..Limits::default()
    };
    let limited =
        Snapshot::from_store_with_limits(source.store().clone(), limited_properties).unwrap();
    let mut edit = limited.edit();
    assert!(edit.set_properties(0, oversized_properties).is_err());
    assert_eq!(edit.items(), limited.items());
}
