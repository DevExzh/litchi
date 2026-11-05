#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Source limits enforced by the bounded record parser (moved from the
//! embedded-font vectors when `Record::parse` became
//! crate-private, record 0357).

use super::Record;
use crate::{Error, SourceLimitKind, SourceLimits};

const RT_FONT_COLLECTION: u16 = 0x07d5;
const RT_FONT_ENTITY_ATOM: u16 = 0x0fb7;

fn record(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn collection(record_type: u16, payload: &[u8]) -> Vec<u8> {
    record(0x0f, 0, record_type, payload)
}

fn nested_record(depth: usize) -> Vec<u8> {
    let mut value = record(0, 0, RT_FONT_ENTITY_ATOM, &[]);
    for _ in 0..depth {
        value = collection(RT_FONT_COLLECTION, &value);
    }
    value
}

#[test]
fn source_limits_reject_oversize_font_records_before_materialization() {
    let bytes = collection(RT_FONT_COLLECTION, &[0; 128]);
    let limits = SourceLimits::builder()
        .max_entry_bytes(bytes.len())
        .unwrap()
        .max_record_bytes(64)
        .unwrap()
        .max_record_payload_bytes(56)
        .unwrap()
        .max_aggregate_payload_bytes(56)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Record::parse(&bytes, 0, &limits),
        Err(Error::Limit {
            kind: SourceLimitKind::RecordBytes | SourceLimitKind::RecordPayloadBytes,
            ..
        })
    ));
}

#[test]
fn source_limits_bound_record_events_and_depth() {
    let event_bytes = collection(
        RT_FONT_COLLECTION,
        &[
            record(0, 0, RT_FONT_ENTITY_ATOM, &[]),
            record(0, 1, RT_FONT_ENTITY_ATOM, &[]),
        ]
        .concat(),
    );
    let event_limits = SourceLimits::builder()
        .max_events(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Record::parse(&event_bytes, 0, &event_limits),
        Err(Error::Limit {
            kind: SourceLimitKind::Events,
            observed,
            maximum: maximum @ 1,
        }) if observed > maximum
    ));

    let depth_limits = SourceLimits::builder()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Record::parse(&nested_record(3), 0, &depth_limits),
        Err(Error::Limit {
            kind: SourceLimitKind::Depth,
            observed,
            maximum: maximum @ 1,
        }) if observed > maximum
    ));
}
