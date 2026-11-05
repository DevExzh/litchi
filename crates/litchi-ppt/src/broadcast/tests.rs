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
    wrong_order.data = data.into();
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
    let snapshot = Snapshot::parse(source.clone(), &crate::SourceLimits::default()).unwrap();
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
    let snapshot = Snapshot::parse(source, &crate::SourceLimits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.broadcast().clone();

    assert!(
        transaction
            .set_chat_url(Some("https://chat.example.test".into()))
            .is_err()
    );
    assert_eq!(transaction.broadcast(), &before);
    assert!(transaction.changes().is_empty());

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
    let snapshot = Snapshot::parse(source.clone(), &crate::SourceLimits::default()).unwrap();
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
    let stale = Snapshot::parse(stale_bytes, &crate::SourceLimits::default()).unwrap();
    assert!(commit.patch().apply(&stale).is_err());
}

/// A built one-slide package whose live document carries the test broadcast
/// in its `___PPT9` programmable tag.
fn package_with_broadcast() -> (crate::Package, Vec<u8>) {
    use crate::consts::RecordType;
    use crate::record_wire as wire;
    use crate::slide_order::records::{Owner, Rewrite, publish};

    let mut builder = crate::Builder::new();
    builder.add_slide().unwrap();
    let package = builder.build().unwrap();
    let base = crate::scoped::Base::capture(&package).unwrap();
    let native = base.native().unwrap();
    let document = base.owner_record(&native, Owner::Document).unwrap();
    let children = wire::child_list(&document).unwrap();
    let info = children
        .iter()
        .position(|child| wire::record_type(child).unwrap() == RecordType::DocInfoList as u16);
    let list = match info {
        Some(index) => children[index].to_vec(),
        None => wire::container(RecordType::DocInfoList as u16, 0, &[]).unwrap(),
    };
    let blob = crate::binary_tag_edit::locate(
        &list,
        RecordType::DocInfoList as u16,
        "___PPT9",
        "broadcast",
    )
    .unwrap();
    let mut records = blob
        .records()
        .iter()
        .map(|record| record.to_vec())
        .collect::<Vec<_>>();
    records.push(record_bytes(&broadcast().to_record().unwrap()));
    let list = blob.rebuild(&list, &records, &[]).unwrap();
    let mut data = Vec::new();
    for (index, child) in children.iter().enumerate() {
        if Some(index) == info {
            data.extend_from_slice(&list);
            continue;
        }
        // A new information list goes before the closing EndDocumentAtom.
        if info.is_none() && index + 1 == children.len() {
            data.extend_from_slice(&list);
        }
        data.extend_from_slice(child);
    }
    let after = wire::rebuild(&document, &data).unwrap();
    let published = publish(
        &native,
        &[Rewrite {
            owner: Owner::Document,
            before: document.into(),
            after: after.into(),
        }],
    )
    .unwrap();
    let bytes = published.bytes().to_vec();
    (crate::Package::from_bytes(&bytes).unwrap(), bytes)
}

#[test]
fn package_broadcast_edits_publish_replay_and_invert_exactly() {
    use litchi_core::Position;

    let (source, source_bytes) = package_with_broadcast();
    let read = source.presentation().unwrap().broadcasts().unwrap();
    assert_eq!(read.broadcasts, vec![broadcast()]);

    let no_op = source
        .edit_broadcast(Position::new(0))
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_broadcast(Position::new(0)).unwrap();
    edit.set_title(Some("Scoped broadcast".into())).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let published = commit
        .package()
        .presentation()
        .unwrap()
        .broadcasts()
        .unwrap();
    assert_eq!(
        published.broadcasts[0].title.as_deref(),
        Some("Scoped broadcast")
    );

    let replayed = source.apply_broadcast(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_broadcast(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_broadcast(commit.patch()),
        Err(super::Error::PatchConflict)
    ));
    let Err(super::Error::Refused(refusal)) = source
        .edit_broadcast(Position::new(1))
        .and_then(|edit| edit.broadcast().map(|_| ()))
    else {
        panic!("a position past the last broadcast must be refused");
    };
    assert_eq!(refusal.kind(), litchi_core::edit::RefusalKind::NotFound);
}

#[test]
fn package_broadcast_edits_admit_a_signed_package_and_refuse_a_changed_commit() {
    use litchi_core::Position;

    let (source, _) = package_with_broadcast();
    let signed = crate::scoped::signed(&source);
    let bytes = signed.to_bytes().unwrap();

    let no_op = signed
        .edit_broadcast(Position::new(0))
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), bytes);

    let mut edit = signed.edit_broadcast(Position::new(0)).unwrap();
    edit.set_title(Some("Scoped broadcast".into())).unwrap();
    assert!(matches!(edit.commit(), Err(super::Error::Signed)));
    assert_eq!(signed.to_bytes().unwrap(), bytes);
}
