use super::*;

use crate::revision_log::Revision;
use crate::revision_records::{EOF_RECORD_TYPE, RR_AUTO_FMT_RECORD_TYPE, RRD_INFO_RECORD_TYPE};

fn record(record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut result = Vec::with_capacity(RECORD_HEADER_LEN + payload.len());
    result.extend_from_slice(&record_type.to_le_bytes());
    result.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    result.extend_from_slice(payload);
    result
}

fn rrd(revision_type: u16, revision_id: i32, tab_id: u16) -> [u8; RRD_LEN] {
    let mut result = [0; RRD_LEN];
    result[0..4].copy_from_slice(&26u32.to_le_bytes());
    result[4..8].copy_from_slice(&revision_id.to_le_bytes());
    result[8..10].copy_from_slice(&revision_type.to_le_bytes());
    result[12..14].copy_from_slice(&tab_id.to_le_bytes());
    result
}

fn short_dtr() -> [u8; 8] {
    [0xE7, 0x07, 6, 30, 9, 15, 0, 5]
}

fn fixed_string(field_len: usize, value: &str) -> Vec<u8> {
    let mut result = vec![0; field_len];
    result[1..1 + value.len()].copy_from_slice(value.as_bytes());
    result
}

fn revision_info() -> Vec<u8> {
    let mut payload = vec![0; 50];
    payload[0..2].copy_from_slice(&8u16.to_le_bytes());
    payload[4..6].copy_from_slice(&0x000Bu16.to_le_bytes());
    payload[38..42].copy_from_slice(&99i32.to_le_bytes());
    payload[42..46].copy_from_slice(&4u32.to_le_bytes());
    payload[46..48].copy_from_slice(&45u16.to_le_bytes());
    record(RRD_INFO_RECORD_TYPE, &payload)
}

fn revision_header(user_name: &str) -> Vec<u8> {
    let mut payload = Vec::with_capacity(158);
    let mut header = rrd(0x0020, 0, 0xFFFF);
    header[0..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    payload.extend_from_slice(&header);
    payload.extend_from_slice(&[0x2A; 16]);
    payload.extend_from_slice(&1200u16.to_le_bytes());
    payload.extend_from_slice(&(user_name.len() as u16).to_le_bytes());
    payload.extend_from_slice(&fixed_string(RRD_HEAD_USER_FIELD_LEN, user_name));
    payload.extend_from_slice(&short_dtr());
    payload.extend_from_slice(&4i16.to_le_bytes());
    record(RRD_HEAD_RECORD_TYPE, &payload)
}

fn rename_sheet() -> Vec<u8> {
    let mut payload = Vec::with_capacity(528);
    payload.extend_from_slice(&rrd(0x0009, 21, 1));
    payload.extend_from_slice(&6u16.to_le_bytes());
    payload.extend_from_slice(&fixed_string(255, "Budget"));
    payload.extend_from_slice(&7u16.to_le_bytes());
    payload.extend_from_slice(&fixed_string(255, "Budget2"));
    record(RRD_REN_SHEET_RECORD_TYPE, &payload)
}

fn source_stream() -> Vec<u8> {
    let mut stream = revision_info();
    stream.extend_from_slice(&revision_header("Alice"));
    stream.extend_from_slice(&rename_sheet());
    // An opaque revision remains byte-identical through a flag edit.
    stream.extend_from_slice(&record(RR_AUTO_FMT_RECORD_TYPE, &[0xA1, 0xB2, 0xC3]));
    stream.extend_from_slice(&record(EOF_RECORD_TYPE, &[]));
    stream
}

#[test]
fn snapshot_noop_preserves_exact_stream_and_fingerprint() {
    let bytes = source_stream();
    let snapshot = Snapshot::parse(bytes.clone()).unwrap();

    assert_eq!(snapshot.finish(), bytes);
    assert_eq!(snapshot.bytes(), bytes.as_slice());
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().source_fingerprint(), snapshot.fingerprint());
    assert_eq!(commit.snapshot().finish(), bytes);
}

#[test]
fn revision_flags_edit_is_typed_and_preserves_opaque_bytes() {
    let bytes = source_stream();
    let snapshot = Snapshot::parse(bytes.clone()).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .set_revision_flags(21, RevisionFlags::new(true, true))
        .unwrap();
    let commit = transaction.commit().unwrap();

    assert!(commit.changed());
    assert_eq!(commit.snapshot().revision_log().headers().len(), 1);
    match &commit.snapshot().revision_log().headers()[0].revisions()[0] {
        Revision::RenSheet(sheet) => {
            assert!(sheet.header().is_accepted());
            assert!(sheet.header().is_undo_action());
        },
        other => panic!("expected rename revision, got {other:?}"),
    }
    assert!(
        commit
            .snapshot()
            .bytes()
            .windows(3)
            .any(|window| window == [0xA1, 0xB2, 0xC3])
    );
    assert_eq!(commit.patch().apply(&snapshot).unwrap(), *commit.snapshot());
}

#[test]
fn header_name_edit_supports_unicode_and_exact_noop() {
    let snapshot = Snapshot::parse(source_stream()).unwrap();
    let mut transaction = snapshot.edit();
    transaction.set_header_user_name(0, "Élodie").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().revision_log().headers()[0]
            .head()
            .user_name(),
        "Élodie"
    );

    let mut noop = commit.snapshot().edit();
    noop.set_header_user_name(0, "Élodie").unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.changed());
    assert_eq!(noop_commit.snapshot().bytes(), commit.snapshot().bytes());
}

#[test]
fn invalid_edits_are_failure_atomic() {
    let snapshot = Snapshot::parse(source_stream()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();

    assert!(transaction.set_revision_accepted(999, true).is_err());
    assert!(transaction.set_header_user_name(4, "Nobody").is_err());
    assert!(
        transaction
            .set_header_user_name(0, &"x".repeat(55))
            .is_err()
    );
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn patch_rejects_stale_source_and_supports_inverse() {
    let bytes = source_stream();
    let snapshot = Snapshot::parse(bytes.clone()).unwrap();
    let mut transaction = snapshot.edit();
    transaction.set_revision_undo_action(21, true).unwrap();
    let commit = transaction.commit().unwrap();

    let mut stale_bytes = bytes;
    let opaque = stale_bytes
        .windows(3)
        .position(|window| window == [0xA1, 0xB2, 0xC3])
        .unwrap();
    stale_bytes[opaque] ^= 1;
    let stale = Snapshot::parse(stale_bytes).unwrap();
    let stale_before = stale.finish();
    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(stale.finish(), stale_before);

    let applied = commit.patch().apply(&snapshot).unwrap();
    let reverted = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(reverted.bytes(), snapshot.bytes());
}
