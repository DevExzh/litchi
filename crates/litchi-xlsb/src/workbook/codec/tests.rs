#![allow(
    clippy::cast_possible_truncation,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Focused tests for the private workbook codec owners.

use super::super::super::Workbook;
use crate::package::error::Result;
use crate::package::shared_strings::SharedString;
use crate::raw::{Kind, Records, Writer, kind};

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn record_stream(records: &[(Kind, Vec<u8>)]) -> Result<Vec<u8>> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    for (record_kind, payload) in records {
        writer.write_record(*record_kind, payload)?;
    }
    Ok(data)
}

fn push_independent_record(bytes: &mut Vec<u8>, record_kind: Kind, payload: &[u8]) {
    let mut value = u32::from(record_kind.get());
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }

    let mut value = payload.len() as u32;
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }
    bytes.extend_from_slice(payload);
}

fn nullable_wide_string(value: Option<&str>) -> Vec<u8> {
    value.map_or_else(
        || u32::MAX.to_le_bytes().to_vec(),
        |value| wide_string(value),
    )
}

fn independent_table_header() -> Vec<u8> {
    let mut payload = Vec::new();
    for value in [0u32, 1, 0, 0] {
        payload.extend_from_slice(&value.to_le_bytes());
    }
    payload.extend_from_slice(&2u32.to_le_bytes()); // LTXML
    payload.extend_from_slice(&7u32.to_le_bytes()); // idList
    payload.extend_from_slice(&1u32.to_le_bytes()); // crwHeader
    payload.extend_from_slice(&0u32.to_le_bytes()); // crwTotals
    payload.extend_from_slice(&0u32.to_le_bytes()); // flags
    for _ in 0..6 {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    payload.extend_from_slice(&0u32.to_le_bytes()); // dwConnID
    payload.extend_from_slice(&nullable_wide_string(Some("MappedTable")));
    payload.extend_from_slice(&nullable_wide_string(Some("MappedTable")));
    for _ in 0..4 {
        payload.extend_from_slice(&nullable_wide_string(None));
    }
    payload
}

fn independent_table_column() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes()); // idField
    payload.extend_from_slice(&0u32.to_le_bytes()); // ilta
    for _ in 0..3 {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    payload.extend_from_slice(&0u32.to_le_bytes()); // idqsif
    payload.extend_from_slice(&nullable_wide_string(None));
    payload.extend_from_slice(&nullable_wide_string(Some("MappedColumn")));
    for _ in 0..4 {
        payload.extend_from_slice(&nullable_wide_string(None));
    }
    payload
}

fn independent_xml_properties() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes()); // map ID
    payload.extend_from_slice(&2u32.to_le_bytes()); // can map a single cell
    payload.extend_from_slice(&1u32.to_le_bytes()); // XML data type
    payload.extend_from_slice(&wide_string("/root/item"));
    payload
}

#[test]
fn shared_string_owner_validates_declared_count() {
    let mut item = vec![0];
    item.extend_from_slice(&wide_string("value"));
    let data = record_stream(&[
        (
            kind::BEGIN_SST,
            [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
        ),
        (kind::SST_ITEM, item),
        (kind::END_SST, Vec::new()),
    ])
    .unwrap();
    let mut records = Records::new(&data);
    let mut strings: Vec<SharedString> = Vec::new();
    Workbook::read_shared_strings(&mut records, &mut strings).unwrap();
    assert_eq!(strings.len(), 1);
    assert_eq!(strings[0].text, "value");
}

#[test]
fn pivot_cache_record_owner_preserves_relationship_identity() {
    let mut cache = 12u32.to_le_bytes().to_vec();
    cache.extend_from_slice(&wide_string("rIdCache"));
    let data = record_stream(&[
        (kind::BEGIN_PIVOT_CACHE_IDS, Vec::new()),
        (kind::BEGIN_PIVOT_CACHE_ID, cache),
        (kind::END_PIVOT_CACHE_ID, Vec::new()),
        (kind::END_PIVOT_CACHE_IDS, Vec::new()),
    ])
    .unwrap();
    assert_eq!(
        Workbook::parse_pivot_cache_ids(&data).unwrap(),
        vec![(12, "rIdCache".to_string())]
    );
}

#[test]
fn ordinary_table_owner_ignores_opaque_balanced_xml_map_extensions() {
    let mut data = Vec::new();
    push_independent_record(&mut data, kind::BEGIN_LIST, &independent_table_header());
    push_independent_record(&mut data, kind::BEGIN_LIST_COLS, &1u32.to_le_bytes());
    push_independent_record(&mut data, kind::BEGIN_LIST_COL, &independent_table_column());
    push_independent_record(
        &mut data,
        kind::BEGIN_LIST_XML_CPR,
        &independent_xml_properties(),
    );
    push_independent_record(&mut data, Kind::new(0x0FFE).unwrap(), b"opaque");
    push_independent_record(&mut data, kind::FRT_BEGIN, &[0; 4]);
    push_independent_record(&mut data, kind::AC_BEGIN, &[]);
    // A known table record is opaque while protected by extension wrappers;
    // it must not close the host parser's real column.
    push_independent_record(&mut data, kind::END_LIST_COL, &[]);
    push_independent_record(&mut data, kind::AC_END, &[]);
    push_independent_record(&mut data, kind::FRT_END, &[]);
    push_independent_record(&mut data, kind::END_LIST_XML_CPR, &[]);
    push_independent_record(&mut data, kind::END_LIST_COL, &[]);
    push_independent_record(&mut data, kind::END_LIST_COLS, &[]);
    push_independent_record(&mut data, kind::END_LIST, &[]);

    let table = Workbook::parse_table_definition(&data, 3).unwrap();
    assert_eq!(table.table_id(), 7);
    assert_eq!(table.sheet_index(), 3);
    assert_eq!(table.display_name(), "MappedTable");
    assert_eq!(table.columns().len(), 1);
}

#[test]
fn ordinary_table_owner_rejects_mismatched_xml_map_extensions() {
    let mut data = Vec::new();
    push_independent_record(&mut data, kind::BEGIN_LIST, &independent_table_header());
    push_independent_record(&mut data, kind::BEGIN_LIST_COLS, &1u32.to_le_bytes());
    push_independent_record(&mut data, kind::BEGIN_LIST_COL, &independent_table_column());
    push_independent_record(
        &mut data,
        kind::BEGIN_LIST_XML_CPR,
        &independent_xml_properties(),
    );
    push_independent_record(&mut data, kind::FRT_BEGIN, &[0; 4]);
    push_independent_record(&mut data, kind::AC_END, &[]);

    assert!(Workbook::parse_table_definition(&data, 0).is_err());
}
