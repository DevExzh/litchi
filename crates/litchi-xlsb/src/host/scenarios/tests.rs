#![allow(
    clippy::expect_used,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

use super::codec;
use super::package::{parse_worksheet, replace_worksheet};
use super::{CellRange, ChangedCell, Manager, Scenario};
use crate::package::error::Error;
use crate::raw::{Header, Kind, Limits, Records, Writer, kind};

fn scenario() -> Scenario {
    let mut scenario = Scenario::new("Forecast", "Analyst").unwrap();
    scenario.set_comment("bounded what-if input").unwrap();
    scenario.set_locked(true);
    scenario
        .set_changed_cells(vec![
            ChangedCell::with_number_format(3, 4, 7, "42").unwrap(),
            ChangedCell::new(8, 2, "North").unwrap(),
        ])
        .unwrap();
    scenario
}

fn manager() -> Manager {
    let mut manager = Manager::new(vec![scenario()]).unwrap();
    manager.set_current(Some(0)).unwrap();
    manager.set_shown(Some(0)).unwrap();
    manager
        .set_result_ranges(vec![CellRange::new(10, 10, 1, 2).unwrap()])
        .unwrap();
    manager
}

fn worksheet(manager: &Manager) -> Vec<u8> {
    let mut data = codec::write_manager(manager).unwrap();
    Writer::new(&mut data)
        .write_record(kind::END_SHEET, &[])
        .unwrap();
    data
}

fn record(kind: Kind, payload: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    Writer::new(&mut data).write_record(kind, payload).unwrap();
    data
}

fn insert_before_kind(data: &[u8], target: Kind, inserted: &[u8]) -> Vec<u8> {
    let offset = Records::new(data)
        .map(|result| result.unwrap())
        .find(|record| record.kind() == target)
        .map(|record| record.offset())
        .expect("target record is present");
    let mut output = Vec::with_capacity(data.len() + inserted.len());
    output.extend_from_slice(&data[..offset]);
    output.extend_from_slice(inserted);
    output.extend_from_slice(&data[offset..]);
    output
}

#[test]
fn typed_manager_round_trips_known_metadata() {
    let original = manager();
    let data = worksheet(&original);
    let parsed = parse_worksheet(&data).unwrap().unwrap();

    assert_eq!(parsed.current(), Some(0));
    assert_eq!(parsed.shown(), Some(0));
    assert_eq!(
        parsed.result_ranges(),
        [CellRange::new(10, 10, 1, 2).unwrap()]
    );
    assert_eq!(parsed.scenarios().len(), 1);
    assert_eq!(parsed.scenarios()[0].name(), "Forecast");
    assert!(parsed.scenarios()[0].locked());
    assert_eq!(parsed.scenarios()[0].changed_cells()[0].value(), "42");
    assert_eq!(
        codec::write_manager(&parsed).unwrap(),
        codec::write_manager(&original).unwrap()
    );
    assert_eq!(replace_worksheet(&data, Some(&parsed)).unwrap(), data);
}

#[test]
fn opaque_records_keep_order_and_unchanged_package_bytes() {
    let manager = manager();
    let mut data = worksheet(&manager);
    let manager_unknown = record(Kind::new(0x0ffe).unwrap(), b"manager extension");
    let scenario_unknown = record(Kind::new(0x0ffd).unwrap(), b"scenario extension");
    data = insert_before_kind(&data, codec::begin_scenario(), &manager_unknown);
    data = insert_before_kind(&data, codec::scenario_cell(), &scenario_unknown);

    let parsed = parse_worksheet(&data).unwrap().unwrap();
    assert_eq!(parsed.unknown_records().len(), 1);
    assert_eq!(parsed.unknown_records()[0].payload(), b"manager extension");
    assert_eq!(parsed.scenarios()[0].unknown_records().len(), 1);
    assert_eq!(
        parsed.scenarios()[0].unknown_records()[0].payload(),
        b"scenario extension"
    );
    assert_eq!(replace_worksheet(&data, Some(&parsed)).unwrap(), data);
}

#[test]
fn opaque_manager_rejects_semantic_edits_and_removal() {
    let manager = manager();
    let mut data = worksheet(&manager);
    let unknown = record(Kind::new(0x0ffe).unwrap(), b"must not be discarded");
    data = insert_before_kind(&data, codec::begin_scenario(), &unknown);
    let parsed = parse_worksheet(&data).unwrap().unwrap();

    let mut candidate = parsed.clone();
    candidate.set_current(None).unwrap();
    assert!(matches!(
        replace_worksheet(&data, Some(&candidate)),
        Err(Error::UnsupportedFeature(_))
    ));
    assert!(matches!(
        replace_worksheet(&data, None),
        Err(Error::UnsupportedFeature(_))
    ));
}

#[test]
fn package_splice_inserts_replaces_and_removes_only_the_known_block() {
    let prefix = record(Kind::new(0x0ffc).unwrap(), b"worksheet metadata");
    let mut absent = prefix.clone();
    Writer::new(&mut absent)
        .write_record(kind::END_SHEET, &[])
        .unwrap();

    let original = manager();
    let inserted = replace_worksheet(&absent, Some(&original)).unwrap();
    assert_eq!(&inserted[..prefix.len()], &prefix);
    assert!(parse_worksheet(&inserted).unwrap().is_some());

    let mut edited = original.clone();
    let mut scenario = edited.scenarios()[0].clone();
    scenario
        .set_changed_cell(0, ChangedCell::new(3, 4, "99").unwrap())
        .unwrap();
    edited.set_scenario(0, scenario).unwrap();
    let replaced = replace_worksheet(&inserted, Some(&edited)).unwrap();
    assert_eq!(&replaced[..prefix.len()], &prefix);
    assert_eq!(
        parse_worksheet(&replaced).unwrap().unwrap().scenarios()[0].changed_cells()[0].value(),
        "99"
    );
    assert_eq!(replace_worksheet(&replaced, None).unwrap(), absent);
}

#[test]
fn reserved_changed_cell_bits_are_rejected() {
    let mut data = worksheet(&manager());
    let (record_offset, header_len) = {
        let record = Records::new(&data)
            .map(|result| result.unwrap())
            .find(|record| record.kind() == codec::scenario_cell())
            .expect("scenario cell is present");
        let (header, header_len) =
            Header::parse(&data[record.offset()..], Limits::DEFAULT).unwrap();
        assert_eq!(header.kind(), codec::scenario_cell());
        (record.offset(), header_len)
    };
    let payload_offset = record_offset + header_len;
    data[payload_offset + 8..payload_offset + 12].copy_from_slice(&1_u32.to_le_bytes());
    assert!(parse_worksheet(&data).is_err());
}

#[test]
fn model_rejects_invalid_selection_and_result_ranges() {
    let mut manager = Manager::new(Vec::new()).unwrap();
    assert!(manager.set_current(Some(0)).is_err());

    let mut manager = Manager::new(vec![scenario()]).unwrap();
    assert!(
        manager
            .set_result_ranges(vec![CellRange::new(0, 0, 0, 32).unwrap()])
            .is_err()
    );
}
