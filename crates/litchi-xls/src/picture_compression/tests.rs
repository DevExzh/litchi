use super::{RECORD_TYPE, Record, Settings, Snapshot, parse, write};

fn frame(record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut value = Vec::with_capacity(4 + payload.len());
    value.extend_from_slice(&record_type.to_le_bytes());
    value.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    value.extend_from_slice(payload);
    value
}

fn payload(recommend: u32, tail: &[u8]) -> Vec<u8> {
    let mut value = Vec::from(RECORD_TYPE.to_le_bytes());
    value.extend_from_slice(&[0; 10]);
    value.extend_from_slice(&recommend.to_le_bytes());
    value.extend_from_slice(tail);
    value
}

#[test]
fn round_trip_preserves_unknown_records_and_extension_tail() {
    let mut input = frame(0x1234, &[0xAA, 0x00]);
    input.extend_from_slice(&frame(RECORD_TYPE, &payload(1, &[7, 8, 9])));
    let snapshot = parse(&input).unwrap();
    assert!(snapshot.settings().unwrap().recommends_compression());
    assert_eq!(snapshot.settings().unwrap().opaque_tail(), &[7, 8, 9]);
    assert_eq!(write(&snapshot).unwrap(), input);
}

#[test]
fn transaction_applies_atomically_and_checks_the_base() {
    let snapshot = Snapshot::try_new(vec![Record::Settings(Settings::new(false))]).unwrap();
    let mut edit = snapshot.edit();
    edit.set_settings(Settings::new(true)).unwrap();
    edit.insert_unknown(0, 0x4321, [1, 2, 3]).unwrap();
    let patch = edit.commit().unwrap();

    let mut target = snapshot.clone();
    patch.apply(&mut target).unwrap();
    assert!(target.settings().unwrap().recommends_compression());
    assert_eq!(
        target.unknown_records().next().unwrap().record_type(),
        0x4321
    );

    let mut changed = snapshot.clone();
    let mut diverged = changed.edit();
    diverged.set_settings(Settings::new(true)).unwrap();
    diverged.commit().unwrap().apply(&mut changed).unwrap();
    let before_failed_apply = changed.clone();
    assert!(patch.apply(&mut changed).is_err());
    assert_eq!(changed, before_failed_apply);
}

#[test]
fn malformed_payloads_and_duplicate_settings_are_rejected() {
    assert!(parse(&[0x9B, 0x08, 0x10]).is_err());
    assert!(parse(&frame(RECORD_TYPE, &payload(2, &[]))).is_err());

    let mut wrong_header = payload(0, &[]);
    wrong_header[0] = 0;
    assert!(parse(&frame(RECORD_TYPE, &wrong_header)).is_err());

    let mut duplicate = frame(RECORD_TYPE, &payload(0, &[]));
    duplicate.extend_from_slice(&frame(RECORD_TYPE, &payload(1, &[])));
    assert!(parse(&duplicate).is_err());
}

#[test]
fn unbounded_unknown_payload_is_rejected() {
    assert!(Record::unknown(0x4321, vec![0; 8_225]).is_err());
}
