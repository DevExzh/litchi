#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn record(
    record_type: RecordType,
    version: u16,
    instance: u16,
    data: Vec<u8>,
    children: Vec<Record>,
) -> Record {
    Record {
        record_type,
        record_type_raw: record_type.as_u16(),
        version,
        instance,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.into(),
        children,
    }
}

fn atom(record_type: RecordType, version: u16, instance: u16, data: &[u8]) -> Record {
    record(record_type, version, instance, data.to_vec(), Vec::new())
}

fn wire(record: &Record) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + record.data.len());
    bytes.extend_from_slice(&(record.version | (record.instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&record.data_length.to_le_bytes());
    bytes.extend_from_slice(&record.data);
    bytes
}

fn root(children: Vec<Record>) -> Record {
    let data = children.iter().flat_map(wire).collect::<Vec<_>>();
    record(RecordType::Slide, 0x0f, 0, data, children)
}

fn source() -> Snapshot {
    Snapshot::from_record(
        root(vec![
            atom(RecordType::SlideAtom, 2, 0, &[0; 24]),
            atom(RecordType::Unknown, 0, 7, &[0xaa, 0xbb]),
        ]),
        &Limits::default(),
    )
    .unwrap()
}

fn time(year: u16) -> SystemTime {
    SystemTime::new(year, 1, 1, 1, 0, 0, 0, 0).unwrap()
}

fn synchronization() -> Synchronization {
    Synchronization::new(
        "server-slide-42",
        "http://example.com/slides?id=42",
        time(2024),
        time(2025),
    )
    .unwrap()
}

#[test]
fn typed_values_validate_and_preserve_url_text() {
    let value = synchronization();
    assert_eq!(value.server_slide_id().as_str(), "server-slide-42");
    assert_eq!(
        value.slide_library_url().as_str(),
        "http://example.com/slides?id=42"
    );
    assert_eq!(value.server_modified().year(), 2024);
    assert!(ServerId::new("bad\u{0}").is_err());
    assert!(LibraryUrl::new("https://example.com").is_err());
    // 2023 is not a leap year, so February 29 does not exist.
    assert!(SystemTime::new(2023, 2, 0, 29, 0, 0, 0, 0).is_err());
    // 2024 is a leap year: February 29 is valid and the weekday is retained
    // verbatim (SYSTEMTIME weekdays are informational, not recomputed).
    assert!(SystemTime::new(2024, 2, 4, 29, 0, 0, 0, 0).is_ok());
}

#[test]
fn parses_and_round_trips_slide_sync_without_editing_unknown_records() {
    let value = synchronization();
    let mut editor = source().edit();
    editor.set(value.clone()).unwrap();
    let commit = editor.commit().unwrap();

    let target = commit.snapshot();
    assert_eq!(target.synchronization().unwrap(), Some(value));
    assert_eq!(target.record().children[1].record_type, RecordType::Unknown);
    assert_eq!(target.record().children[1].data, [0xaa, 0xbb]);
    let reparsed = Snapshot::parse(target.bytes(), &Limits::default()).unwrap();
    assert_eq!(reparsed.bytes(), target.bytes());
}

#[test]
fn synchronization_strings_do_not_enter_visible_slide_text() {
    let utf16 = |text: &str| {
        text.encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>()
    };
    let original = Snapshot::from_record(
        root(vec![
            atom(RecordType::TextCharsAtom, 0, 0, &utf16("Visible body")),
            atom(RecordType::CString, 0, 0, &utf16("Visible CString")),
        ]),
        &Limits::default(),
    )
    .unwrap();
    let mut edit = original.edit();
    edit.set(synchronization()).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().record().extract_text().unwrap(),
        "Visible body\nVisible CString"
    );
    assert_eq!(
        commit.snapshot().synchronization().unwrap(),
        Some(synchronization())
    );
    let reopened = Snapshot::parse(commit.snapshot().bytes(), &Limits::default()).unwrap();
    assert_eq!(
        reopened.record().extract_text().unwrap(),
        original.record().extract_text().unwrap()
    );
}

#[test]
fn set_and_clear_are_atomic_and_reversible() {
    let original = source();
    let original_bytes = original.bytes().to_vec();
    let mut editor = original.edit();
    editor.set(synchronization()).unwrap();
    let commit = editor.commit().unwrap();
    assert_eq!(commit.changes().changes().len(), 1);
    assert_ne!(commit.snapshot().bytes(), original_bytes.as_slice());

    let undone = commit.undo(commit.snapshot()).unwrap();
    assert_eq!(undone.bytes(), original_bytes.as_slice());
    assert!(undone.synchronization().unwrap().is_none());
    let redone = commit.redo(&undone).unwrap();
    assert_eq!(redone.bytes(), commit.snapshot().bytes());

    let mut clear = commit.snapshot().edit();
    assert!(clear.clear().unwrap());
    let cleared = clear.commit().unwrap();
    assert!(cleared.synchronization().unwrap().is_none());
}

#[test]
fn rejects_duplicate_or_malformed_sync_records_before_mutation() {
    let value = synchronization();
    let first = codec::encode_sync(&value).unwrap();
    let mut duplicate = source().record().clone();
    duplicate.children.push(first.clone());
    duplicate.children.push(first.clone());
    duplicate.data = duplicate.children.iter().flat_map(wire).collect();
    duplicate.data_length = u32::try_from(duplicate.data.len()).unwrap();
    assert!(Snapshot::from_record(duplicate, &Limits::default()).is_err());

    let malformed = root(vec![record(
        RecordType::RoundTripSlideSyncInfo12,
        0x0f,
        0,
        vec![1, 2, 3],
        Vec::new(),
    )]);
    assert!(Snapshot::from_record(malformed, &Limits::default()).is_err());
}

#[test]
fn rejects_wrong_root_and_stale_revision_operations() {
    let wrong = atom(RecordType::Notes, 0x0f, 0, &[]);
    assert!(Snapshot::from_record(wrong, &Limits::default()).is_err());

    let original = source();
    let mut editor = original.edit();
    editor.set(synchronization()).unwrap();
    let commit = editor.commit().unwrap();
    let mut changed = commit.snapshot().edit();
    changed.clear().unwrap();
    let changed_commit = changed.commit().unwrap();
    assert!(commit.undo(changed_commit.snapshot()).is_err());
}

fn built_package() -> crate::Package {
    let mut builder = crate::Builder::new();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 10, 10, 200, 40, "Synced")
        .unwrap();
    builder.add_slide().unwrap();
    builder.build().unwrap()
}

#[test]
fn package_slide_sync_edits_publish_replay_and_invert_exactly() {
    use litchi_core::Position;
    use litchi_core::edit::RefusalKind;

    let source = built_package();
    let source_bytes = source.to_bytes().unwrap();
    let slide = Position::new(0);
    assert!(
        source
            .edit_slide_sync(slide)
            .unwrap()
            .synchronization()
            .unwrap()
            .is_none()
    );

    let no_op = source.edit_slide_sync(slide).unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    // Clearing absent metadata and setting-then-clearing are exact no-ops.
    let mut edit = source.edit_slide_sync(slide).unwrap();
    assert!(!edit.clear().unwrap());
    edit.set(synchronization()).unwrap();
    assert!(edit.clear().unwrap());
    let net_no_op = edit.commit().unwrap();
    assert!(!net_no_op.diagnostics().changed());
    assert_eq!(net_no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_slide_sync(slide).unwrap();
    edit.set(synchronization()).unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_slide_sync(slide).unwrap();
    edit.set(synchronization()).unwrap();
    assert_eq!(edit.changes().unwrap().len(), 1);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );
    assert_eq!(
        commit
            .package()
            .edit_slide_sync(slide)
            .unwrap()
            .synchronization()
            .unwrap(),
        Some(synchronization())
    );
    assert!(
        commit
            .package()
            .edit_slide_sync(Position::new(1))
            .unwrap()
            .synchronization()
            .unwrap()
            .is_none()
    );

    let replayed = source.apply_slide_sync(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_slide_sync(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_slide_sync(commit.patch()),
        Err(Error::PatchConflict)
    ));

    // Clearing the published metadata republishes the slide without it.
    let mut clear = commit.package().edit_slide_sync(slide).unwrap();
    assert!(clear.clear().unwrap());
    let cleared = clear.commit().unwrap();
    assert!(cleared.diagnostics().changed());
    assert!(
        cleared
            .package()
            .edit_slide_sync(slide)
            .unwrap()
            .synchronization()
            .unwrap()
            .is_none()
    );

    let Err(Error::Refused(refusal)) = source
        .edit_slide_sync(Position::new(2))
        .and_then(|edit| edit.synchronization().map(|_| ()))
    else {
        panic!("a missing slide must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
}

#[test]
fn package_slide_sync_edits_admit_a_signed_package_and_refuse_a_changed_commit() {
    use litchi_core::Position;

    let signed = crate::scoped::signed(&built_package());
    let bytes = signed.to_bytes().unwrap();
    let slide = Position::new(0);

    let no_op = signed.edit_slide_sync(slide).unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), bytes);

    let mut edit = signed.edit_slide_sync(slide).unwrap();
    edit.set(synchronization()).unwrap();
    assert!(matches!(edit.commit(), Err(Error::Signed)));
    assert_eq!(signed.to_bytes().unwrap(), bytes);
}
