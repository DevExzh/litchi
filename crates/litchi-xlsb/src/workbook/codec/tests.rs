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
