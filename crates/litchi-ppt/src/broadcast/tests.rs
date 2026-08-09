#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::{BROADCAST_INFO_RECORD_TYPE, C_STRING_RECORD_TYPE, validate_system_time};
use super::model::{Broadcast, BroadcastProperties};
use super::transaction::Snapshot;
use crate::records::Record;
use crate::slide_sync::SystemTime;

fn time(hour: u16) -> SystemTime {
    SystemTime {
        year: 2026,
        month: 7,
        day_of_week: 0,
        day: 19,
        hour,
        minute: 30,
        second: 15,
        millisecond: 125,
    }
}

fn broadcast() -> Broadcast {
    Broadcast {
        title: Some("Quarterly update".into()),
        description: Some("Roadmap and results".into()),
        speaker: Some("Ada".into()),
        contact: Some("Grace".into()),
        remote_server_name: Some("CAMERA01".into()),
        email_address: Some("feedback@example.test".into()),
        email_name: Some("Feedback".into()),
        chat_url: Some("http://chat.example.test/room?id=7".into()),
        archive_directory: Some("C:\\Archive".into()),
        netshow_files_base_directory: Some("\\\\server\\share".into()),
        netshow_files_directory: Some("\\\\server\\share\\netshow".into()),
        netshow_server_name: Some("NETSHOW01".into()),
        ppt_files_base_directory: "\\\\server\\share".into(),
        ppt_files_directory: "\\\\server\\share\\ppt".into(),
        ppt_files_base_url: "http://slides.example.test/base".into(),
        user_name: "scheduler".into(),
        broadcast_date_time: "2026-07-19T09-30".into(),
        presentation_name: "quarterly.ppt".into(),
        asd_file_name: "\\\\server\\share\\stream.asd".into(),
        entry_id: Some("calendar-item-42".into()),
        properties: BroadcastProperties {
            send_audio: true,
            send_video: true,
            camera_remote: true,
            use_netshow: true,
            use_other_server: false,
            can_email: true,
            can_chat: true,
            archive: true,
            speaker_notes: true,
            quarter_screen: false,
            show_tools: true,
            record_only: false,
            start_time: time(9),
            end_time: time(10),
        },
    }
}

#[test]
fn round_trips_all_broadcast_atoms_and_flags() {
    let expected = broadcast();
    let record = expected.to_record().unwrap();
    let children = Record::parse_sequence_strict(&record.data, "test").unwrap();
    assert_eq!(children.len(), 21);
    assert_eq!(children[0].instance, 1);
    assert_eq!(children[19].instance, 20);
    assert_eq!(children[20].record_type_raw, BROADCAST_INFO_RECORD_TYPE);
    assert_eq!(Broadcast::parse(&record).unwrap(), expected);
}

#[test]
fn rejects_dependency_order_reserved_and_lexical_failures() {
    let mut value = broadcast();
    value.remote_server_name = None;
    assert!(value.validate().is_err());
    value = broadcast();
    value.netshow_server_name = None;
    assert!(value.validate().is_err());
    value = broadcast();
    value.email_name = None;
    assert!(value.validate().is_err());
    value = broadcast();
    value.chat_url = Some("https://wrong-scheme.example".into());
    assert!(value.validate().is_err());

    let valid = broadcast().to_record().unwrap();
    let children = Record::parse_sequence_strict(&valid.data, "test").unwrap();
    let mut data =
        super::codec::record_bytes(0, 2, C_STRING_RECORD_TYPE, &children[1].data).unwrap();
    data.extend_from_slice(
        &super::codec::record_bytes(0, 1, C_STRING_RECORD_TYPE, &children[0].data).unwrap(),
    );
    for child in &children[2..] {
        data.extend_from_slice(
            &super::codec::record_bytes(
                child.version,
                child.instance,
                child.record_type_raw,
                &child.data,
            )
            .unwrap(),
        );
    }
    let mut wrong_order = valid.clone();
    wrong_order.data = data;
    wrong_order.data_length = u32::try_from(wrong_order.data.len()).unwrap();
    assert!(Broadcast::parse(&wrong_order).is_err());

    let mut reserved = valid;
    let atom_start = reserved.data.len() - 34;
    reserved.data[atom_start + 1] |= 0x10;
    assert!(Broadcast::parse(&reserved).is_err());
}

#[test]
fn validates_system_time_and_exact_strict_string_bounds() {
    let mut invalid = time(0);
    invalid.year = 2025;
    invalid.month = 2;
    invalid.day = 29;
    assert!(validate_system_time(invalid).is_err());
    invalid.year = 2024;
    assert!(validate_system_time(invalid).is_ok());
    let mut value = broadcast();
    value.archive_directory = Some("x".repeat(255));
    assert!(value.validate().is_err());
    value = broadcast();
    value.ppt_files_base_directory = "not-unc".into();
    assert!(value.validate().is_err());
    value = broadcast();
    value.user_name = "bad/name".into();
    assert!(value.validate().is_err());
}

fn record_bytes(record: &Record) -> Vec<u8> {
    super::codec::record_bytes(
        record.version,
        record.instance,
        record.record_type_raw,
        &record.data,
    )
    .unwrap()
}

#[test]
fn snapshot_no_op_keeps_exact_source_bytes() {
    let source = record_bytes(&broadcast().to_record().unwrap());
    let snapshot = Snapshot::parse(source.clone()).unwrap();
    let commit = snapshot.edit().commit().unwrap();

    assert!(commit.patch().is_empty());
    assert!(commit.patch().changes().is_empty());
    assert_eq!(commit.patch().before(), source.as_slice());
    assert_eq!(commit.patch().after(), source.as_slice());
    assert_eq!(commit.snapshot().bytes(), source.as_slice());
}

#[test]
fn invalid_inert_targets_are_failure_atomic() {
    let source = record_bytes(&broadcast().to_record().unwrap());
    let snapshot = Snapshot::parse(source).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.broadcast().clone();

    assert!(
        transaction
            .set_chat_url(Some("https://chat.example.test".into()))
            .is_err()
    );
    assert_eq!(transaction.broadcast(), &before);
    assert!(!transaction.is_changed());

    assert!(transaction.set_remote_server_name(None).is_err());
    assert_eq!(transaction.broadcast(), &before);
    assert!(transaction.changes().is_empty());
}

#[test]
fn edits_retain_unknown_children_and_support_inverse_and_stale_rejection() {
    let canonical = broadcast().to_record().unwrap();
    let opaque = super::codec::record_bytes(0, 7, 0x7ffe, &[0xde, 0xad, 0xbe, 0xef]).unwrap();
    let mut data = opaque.clone();
    data.extend_from_slice(&canonical.data);
    let source = super::codec::record_bytes(
        canonical.version,
        canonical.instance,
        canonical.record_type_raw,
        &data,
    )
    .unwrap();
    let snapshot = Snapshot::parse(source.clone()).unwrap();
    assert_eq!(snapshot.unknown_records().len(), 1);
    assert_eq!(
        snapshot.unknown_records()[0].to_record_bytes().unwrap(),
        opaque
    );

    let mut transaction = snapshot.edit();
    transaction.set_title(Some("Updated title".into())).unwrap();
    let commit = transaction.commit().unwrap();
    let target = commit.snapshot();
    assert_eq!(target.unknown_records().len(), 1);
    assert_eq!(
        target.unknown_records()[0].to_record_bytes().unwrap(),
        opaque
    );
    let (target_record, consumed) = Record::parse_strict(target.bytes(), 0).unwrap();
    assert_eq!(consumed, target.bytes().len());
    let children = Record::parse_sequence_strict(&target_record.data, "test").unwrap();
    assert_eq!(children[0].record_type_raw, 0x7ffe);
    assert_eq!(children[0].data, [0xde, 0xad, 0xbe, 0xef]);

    let applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(applied.bytes(), target.bytes());
    assert_eq!(
        commit.patch().undo(target).unwrap().bytes(),
        source.as_slice()
    );
    assert_eq!(
        commit.patch().inverse().apply(target).unwrap().bytes(),
        source.as_slice()
    );

    let mut stale_bytes = source;
    stale_bytes[16] = b'R';
    let stale = Snapshot::parse(stale_bytes).unwrap();
    assert!(commit.patch().apply(&stale).is_err());
}
