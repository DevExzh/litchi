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
        data,
        children,
    }
}

fn atom(record_type: RecordType, version: u16, data: &[u8]) -> Record {
    record(record_type, version, 0, data.to_vec(), Vec::new())
}

fn wire(record: &Record) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + record.data.len());
    bytes.extend_from_slice(&(record.version | (record.instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&record.data_length.to_le_bytes());
    bytes.extend_from_slice(&record.data);
    bytes
}

fn children_wire(children: &[Record]) -> Vec<u8> {
    children.iter().flat_map(wire).collect()
}

fn utf16(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn slide_atom(context: Context) -> Record {
    let mut data = vec![0; 24];
    if context == Context::Title {
        data[0..4].copy_from_slice(&2u32.to_le_bytes());
        data[12..16].copy_from_slice(&0x8000_0000u32.to_le_bytes());
    }
    atom(RecordType::SlideAtom, 2, &data)
}

fn master(context: Context, extras: Vec<Record>) -> Record {
    let mut children = Vec::new();
    match context {
        Context::Main | Context::Title => children.push(slide_atom(context)),
        Context::Notes => children.push(atom(RecordType::NotesAtom, 1, &[0; 8])),
        Context::Handout => {},
    }
    children.push(atom(RecordType::PPDrawing, 0x0f, &[0xaa, 0xbb]));
    children.extend(extras);
    let record_type = match context {
        Context::Main => RecordType::MainMaster,
        Context::Title => RecordType::Slide,
        Context::Notes => RecordType::Notes,
        Context::Handout => RecordType::Handout,
    };
    record(record_type, 0x0f, 0, children_wire(&children), children)
}

fn unknown() -> Record {
    Record {
        record_type: RecordType::Unknown,
        record_type_raw: 0x7abc,
        version: 0,
        instance: 7,
        data_length: 3,
        data: vec![1, 2, 3],
        children: Vec::new(),
    }
}

fn name_record(value: &str) -> Record {
    record(RecordType::CString, 0, 3, utf16(value), Vec::new())
}

#[test]
fn reads_and_authors_names_for_all_master_contexts() {
    for context in [
        Context::Main,
        Context::Title,
        Context::Notes,
        Context::Handout,
    ] {
        let snapshot = Snapshot::from_record(context, master(context, Vec::new())).unwrap();
        assert!(snapshot.name().unwrap().is_none());

        let mut edit = snapshot.edit();
        edit.set_name("Quarterly master").unwrap();
        let committed = edit.commit().unwrap();
        assert_eq!(
            committed.snapshot().name().unwrap().unwrap().as_str(),
            "Quarterly master"
        );
        assert_eq!(committed.changes().changes().len(), 1);
    }
}

#[test]
fn replacement_and_clear_are_atomic_and_keep_unknown_records() {
    let opaque_before = unknown();
    let opaque_after = Record {
        record_type: RecordType::Unknown,
        record_type_raw: 0x7abd,
        version: 0,
        instance: 2,
        data_length: 2,
        data: vec![9, 8],
        children: Vec::new(),
    };
    let source = Snapshot::from_record(
        Context::Main,
        master(
            Context::Main,
            vec![
                opaque_before.clone(),
                name_record("Old"),
                opaque_after.clone(),
            ],
        ),
    )
    .unwrap();

    let mut edit = source.edit();
    edit.set_name("New").unwrap();
    let committed = edit.commit().unwrap();
    let updated = committed.snapshot();
    assert_eq!(updated.name().unwrap().unwrap().as_str(), "New");
    assert_eq!(updated.record().children[2].record_type_raw, 0x7abc);
    assert_eq!(updated.record().children[2].data, opaque_before.data);
    assert_eq!(updated.record().children[4].record_type_raw, 0x7abd);
    assert_eq!(updated.record().children[4].data, opaque_after.data);

    let undone = committed.undo(updated).unwrap();
    assert_eq!(undone.name().unwrap().unwrap().as_str(), "Old");
    let redone = committed.redo(&undone).unwrap();
    assert_eq!(redone.name().unwrap().unwrap().as_str(), "New");

    let mut clear = redone.edit();
    assert!(clear.clear_name().unwrap());
    let cleared = clear.commit().unwrap();
    assert!(cleared.snapshot().name().unwrap().is_none());
    assert_eq!(
        cleared.snapshot().record().children[2].record_type_raw,
        0x7abc
    );
    assert_eq!(
        cleared.snapshot().record().children[3].record_type_raw,
        0x7abd
    );
}

#[test]
fn inserts_before_main_template_and_round_trip_tail() {
    let template = record(RecordType::CString, 0, 2, utf16("Template"), Vec::new());
    let round_trip = atom(RecordType::RoundTripTheme12Atom, 0, &[]);
    let source = Snapshot::from_record(
        Context::Main,
        master(Context::Main, vec![unknown(), template.clone(), round_trip]),
    )
    .unwrap();

    let mut edit = source.edit();
    edit.set_name("Named").unwrap();
    let committed = edit.commit().unwrap();
    let children = &committed.snapshot().record().children;
    assert_eq!(children[2].record_type_raw, 0x7abc);
    assert_eq!(children[3].record_type, RecordType::CString);
    assert_eq!(children[3].instance, 3);
    assert_eq!(children[4].record_type, RecordType::CString);
    assert_eq!(children[4].instance, 2);
    assert_eq!(children[5].record_type, RecordType::RoundTripTheme12Atom);
    assert_eq!(wire(&children[4]), wire(&template));
}

#[test]
fn rejects_duplicate_malformed_and_oversized_names() {
    assert!(Name::new("a".repeat(MAX_NAME_BYTES / 2 + 1)).is_err());

    let duplicate = Snapshot::from_record(
        Context::Handout,
        master(
            Context::Handout,
            vec![name_record("one"), name_record("two")],
        ),
    );
    assert!(duplicate.is_err());

    let odd = name_record("ignored");
    let mut odd_root = master(Context::Notes, vec![odd]);
    odd_root.children[2].data = vec![0];
    odd_root.children[2].data_length = 1;
    odd_root.data = children_wire(&odd_root.children);
    odd_root.data_length = u32::try_from(odd_root.data.len()).unwrap();
    assert!(Snapshot::from_record(Context::Notes, odd_root).is_err());

    let invalid_utf16 = record(RecordType::CString, 0, 3, vec![0, 0xd8], Vec::new());
    assert!(
        Snapshot::from_record(
            Context::Handout,
            master(Context::Handout, vec![invalid_utf16])
        )
        .is_err()
    );

    let wrong_header = record(RecordType::CString, 1, 3, utf16("bad"), Vec::new());
    assert!(
        Snapshot::from_record(
            Context::Handout,
            master(Context::Handout, vec![wrong_header])
        )
        .is_err()
    );
}
