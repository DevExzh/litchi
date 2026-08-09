#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn atom(raw: u16, version: u16, instance: u16, data: &[u8]) -> Record {
    Record {
        record_type: RecordType::from(raw),
        record_type_raw: raw,
        version,
        instance,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.to_vec(),
        children: Vec::new(),
    }
}

fn container(raw: u16, children: Vec<Record>) -> Record {
    Record {
        record_type: RecordType::from(raw),
        record_type_raw: raw,
        version: 0x0f,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children,
    }
}

fn slide_atom(context: Context) -> Record {
    let mut data = vec![0; 24];
    match context {
        Context::Title => {
            data[..4].copy_from_slice(&2u32.to_le_bytes());
            data[12..16].copy_from_slice(&0x8000_0000u32.to_le_bytes());
        },
        Context::Main | Context::Notes | Context::Handout => {},
    }
    atom(RecordType::SlideAtom.as_u16(), 2, 0, &data)
}

fn drawing() -> Record {
    atom(RecordType::PPDrawing.as_u16(), 0x0f, 0, &[0xaa, 0xbb, 0xcc])
}

fn unknown() -> Record {
    atom(0x7abc, 0, 3, &[1, 2, 3, 4])
}

fn layout(context: Context) -> Record {
    let children = match context {
        Context::Main | Context::Title => vec![slide_atom(context), drawing(), unknown()],
        Context::Notes => vec![
            atom(RecordType::NotesAtom.as_u16(), 1, 0, &[0; 8]),
            drawing(),
        ],
        Context::Handout => vec![drawing(), unknown()],
    };
    container(context.expected_record_type().as_u16(), children)
}

#[test]
fn inventories_all_contexts_without_prefix_expanded_types() {
    let document = container(
        RecordType::Document.as_u16(),
        vec![
            layout(Context::Main),
            layout(Context::Title),
            layout(Context::Notes),
            layout(Context::Handout),
        ],
    );
    let inventory = inventory(&document).unwrap();
    assert_eq!(inventory.len(), 4);
    assert_eq!(inventory.entries()[0].context(), Context::Main);
    assert_eq!(inventory.entries()[1].context(), Context::Title);
    assert_eq!(inventory.entries()[2].context(), Context::Notes);
    assert_eq!(inventory.entries()[3].context(), Context::Handout);
    assert_eq!(inventory.entries()[1].path().as_slice(), &[1]);
}

#[test]
fn snapshot_round_trip_preserves_unknown_records() {
    let source = Snapshot::from_record(Context::Main, layout(Context::Main)).unwrap();
    let original = source.bytes().to_vec();
    let unknown_before = source
        .record()
        .children
        .iter()
        .find(|record| record.record_type_raw == 0x7abc)
        .unwrap()
        .clone();

    let mut edit = source.edit();
    edit.add(Path::root(), 1, atom(0x7abd, 0, 0, &[9, 8, 7]))
        .unwrap();
    let commit = edit.commit().unwrap();
    let committed = commit.snapshot();

    assert_eq!(source.bytes(), original.as_slice());
    assert_eq!(committed.context(), Context::Main);
    assert_eq!(
        committed
            .record()
            .children
            .iter()
            .find(|record| record.record_type_raw == 0x7abc),
        Some(&unknown_before)
    );
    assert_eq!(commit.changes().changes().len(), 1);
}

#[test]
fn remove_and_replace_are_transactional() {
    let source = Snapshot::from_record(Context::Main, layout(Context::Main)).unwrap();
    let original = source.bytes().to_vec();
    let mut edit = source.edit();
    let removed = edit.remove(Path::from(2usize)).unwrap();
    assert_eq!(removed.record_type_raw, 0x7abc);
    edit.replace(
        Path::from(1usize),
        atom(RecordType::PPDrawing.as_u16(), 0x0f, 0, &[4, 5]),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.bytes(), original.as_slice());
    assert_eq!(commit.changes().changes().len(), 2);
    assert_eq!(commit.snapshot().record().children.len(), 2);
}

#[test]
fn failed_commit_never_changes_source_snapshot() {
    let source = Snapshot::from_record(Context::Main, layout(Context::Main)).unwrap();
    let original = source.bytes().to_vec();
    let mut edit = source.edit();
    edit.remove(Path::from(0usize)).unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(source.bytes(), original.as_slice());
    assert!(
        source
            .record()
            .children
            .iter()
            .any(|record| { record.record_type == RecordType::SlideAtom })
    );
}

#[test]
fn contextual_validation_rejects_wrong_geometry_and_wrong_root() {
    assert!(Snapshot::from_record(Context::Title, layout(Context::Main)).is_err());

    let mut title = layout(Context::Title);
    title.children[0].data[..4].copy_from_slice(&0x10u32.to_le_bytes());
    title.children[0].data_length = 24;
    assert!(Snapshot::from_record(Context::Title, title).is_err());

    let mut missing_reference = layout(Context::Title);
    missing_reference.children[0].data[12..16].fill(0);
    assert!(Snapshot::from_record(Context::Title, missing_reference).is_err());
}

#[test]
fn failed_path_operation_is_non_mutating() {
    let source = Snapshot::from_record(Context::Main, layout(Context::Main)).unwrap();
    let mut edit = source.edit();
    let before = edit.revision();
    assert!(edit.add(Path::root(), usize::MAX, unknown()).is_err());
    assert_eq!(edit.revision(), before);
    assert_eq!(edit.record(), source.record());
}

#[test]
fn committed_changes_can_be_undone_and_redone() {
    let source = Snapshot::from_record(Context::Main, layout(Context::Main)).unwrap();
    let mut edit = source.edit();
    edit.add(Path::root(), 1, atom(0x7abd, 0, 0, &[9, 8, 7]))
        .unwrap();
    let commit = edit.commit().unwrap();

    let undone = commit.changes().undo(commit.snapshot()).unwrap();
    assert_eq!(undone.bytes(), source.bytes());
    let redone = commit.changes().redo(&undone).unwrap();
    assert_eq!(redone.bytes(), commit.snapshot().bytes());
    assert!(commit.changes().undo(&source).is_err());
}
