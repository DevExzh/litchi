//! Regression tests for the PivotTable-view model and codec.

use super::*;
use crate::raw::{Writer, kind};

fn view_stream(name: &str, cache_id: u32) -> Vec<u8> {
    let mut begin = vec![0u8; 32];
    begin[28..32].copy_from_slice(&cache_id.to_le_bytes());
    begin.extend_from_slice(&(name.encode_utf16().count() as u32).to_le_bytes());
    for unit in name.encode_utf16() {
        begin.extend_from_slice(&unit.to_le_bytes());
    }
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    writer.write_record(kind::BEGIN_SX_VIEW, &begin).unwrap();
    writer
        .write_record(kind::BEGIN_SX_LOCATION, &[0; 36])
        .unwrap();
    writer.write_record(kind::END_SX_LOCATION, &[]).unwrap();
    writer.write_record(kind::END_SX_VIEW, &[]).unwrap();
    bytes
}

#[test]
fn preserves_complete_view_stream_and_extracts_binding() {
    let bytes = view_stream("Revenue Pivot", 17);
    let view = Part::from_bytes(bytes.clone()).unwrap();
    assert_eq!(view.name(), "Revenue Pivot");
    assert_eq!(view.cache_id(), 17);
    assert_eq!(view.version_created(), 0);
    assert_eq!(view.as_bytes(), bytes);
}

#[test]
fn refuses_truncation_and_records_outside_view() {
    let mut truncated = view_stream("P", 1);
    truncated.pop();
    assert!(Part::from_bytes(truncated).is_err());

    let mut trailing = view_stream("P", 1);
    Writer::new(&mut trailing)
        .write_record(kind::END_SX_LOCATION, &[])
        .unwrap();
    assert!(Part::from_bytes(trailing).is_err());
}
