//! Focused worksheet-anchor wire and invariant coverage.

use super::*;
use litchi_odraw::{Record, RecordKind};

fn record(kind: u16, version: u8, instance: u16, body: &[u8]) -> Vec<u8> {
    let ver_inst = u16::from(version) | (instance << 4);
    let mut output = Vec::with_capacity(8 + body.len());
    output.extend_from_slice(&ver_inst.to_le_bytes());
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(body.len() as u32).to_le_bytes());
    output.extend_from_slice(body);
    output
}

fn anchor() -> SheetAnchor {
    SheetAnchor::new(
        AnchorPoint::new(2, 4, -3, 7).unwrap(),
        AnchorPoint::new(9, 18, 1001, -11).unwrap(),
        AnchorBehavior::MoveAndSize,
    )
    .unwrap()
}

#[test]
fn client_anchor_round_trips_exactly() {
    let expected = anchor();
    let bytes = expected.to_record_bytes();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(record.kind(), RecordKind::ClientAnchor);
    assert_eq!(decode_sheet_anchor(&record).unwrap(), expected);
    assert_eq!(expected.to_record_bytes(), bytes);
}

#[test]
fn malformed_anchor_flags_and_extent_are_rejected() {
    let mut bytes = anchor().to_record_bytes();
    bytes[8..10].copy_from_slice(&0x0001u16.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(decode_sheet_anchor(&record).is_err());

    let mut truncated = anchor().to_record_bytes();
    truncated[4..8].copy_from_slice(&17u32.to_le_bytes());
    truncated.truncate(truncated.len() - 1);
    let (record, _) = Record::parse(&truncated, 0).unwrap();
    assert!(decode_sheet_anchor(&record).is_err());
}

#[test]
fn malformed_columns_and_endpoint_order_are_rejected() {
    let mut bytes = anchor().to_record_bytes();
    bytes[10..12].copy_from_slice(&256u16.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(decode_sheet_anchor(&record).is_err());

    let mut body = anchor().to_record_bytes();
    body[18..20].copy_from_slice(&1u16.to_le_bytes());
    let (record, _) = Record::parse(&body, 0).unwrap();
    assert!(decode_sheet_anchor(&record).is_err());
}

#[test]
fn wrong_officeart_record_identity_is_rejected() {
    let encoded = anchor().to_record_bytes();
    let bytes = record(0xF011, 0, 0, &encoded[8..]);
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(decode_sheet_anchor(&record).is_err());
}
