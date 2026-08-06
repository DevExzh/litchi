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
        data_length: data.len() as u32,
        data,
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

fn children_wire(children: &[Record]) -> Vec<u8> {
    children.iter().flat_map(wire).collect()
}

fn utf16(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn name_record(instance: u16, value: &str) -> Record {
    atom(RecordType::CString, 0, instance, &utf16(value))
}

fn unknown(raw: u16, data: &[u8]) -> Record {
    Record {
        record_type: RecordType::Unknown,
        record_type_raw: raw,
        version: 0,
        instance: 7,
        data_length: data.len() as u32,
        data: data.to_vec(),
        children: Vec::new(),
    }
}

fn master(extras: Vec<Record>) -> Record {
    let children = vec![
        atom(RecordType::SlideAtom, 2, 0, &[0; 24]),
        atom(RecordType::PPDrawing, 0x0f, 0, &[0xaa, 0xbb]),
    ]
    .into_iter()
    .chain(extras)
    .collect::<Vec<_>>();
    record(
        RecordType::MainMaster,
        0x0f,
        0,
        children_wire(&children),
        children,
    )
}

#[test]
fn authors_exactly_bounded_template_name_and_round_trips() {
    let value = "x".repeat(MAX_NAME_BYTES / 2);
    let name = Name::new(value.clone()).unwrap();
    let source = Snapshot::from_record(master(Vec::new())).unwrap();
    let mut edit = source.edit();
    edit.set_name(name.as_str()).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().name().unwrap().unwrap().as_str(), value);

    let parsed = Snapshot::parse(commit.snapshot().bytes()).unwrap();
    assert_eq!(parsed.bytes(), commit.snapshot().bytes());
    assert_eq!(parsed.name().unwrap(), commit.snapshot().name().unwrap());
}

#[test]
fn inserts_after_round_trip_and_preserves_unknown_and_slide_name_records() {
    let first = unknown(0x7abc, &[1, 2, 3]);
    let round_trip = atom(RecordType::RoundTripTheme12Atom, 0, 0, &[4, 5]);
    let slide_name = name_record(3, "master-display-name");
    let tail = unknown(0x7abd, &[9, 8]);
    let source = Snapshot::from_record(master(vec![
        first.clone(),
        round_trip.clone(),
        slide_name.clone(),
        tail.clone(),
    ]))
    .unwrap();
    let mut edit = source.edit();
    edit.set_name("design").unwrap();
    let committed = edit.commit().unwrap();
    let children = &committed.snapshot().record().children;
    assert_eq!(children[2], first);
    assert_eq!(children[3], round_trip);
    assert_eq!(children[4], slide_name);
    assert_eq!(children[5], tail);
    assert_eq!(children[6].record_type, RecordType::CString);
    assert_eq!(children[6].instance, 2);
    assert_eq!(
        committed.snapshot().name().unwrap().unwrap().as_str(),
        "design"
    );
}

#[test]
fn replacement_clear_rollback_undo_and_redo_are_transactional() {
    let template = name_record(2, "old-design");
    let source = Snapshot::from_record(master(vec![unknown(0x7abc, &[1]), template])).unwrap();
    let original = source.bytes().to_vec();
    let mut edit = source.edit();
    edit.set_name("new-design").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.bytes(), original.as_slice());
    assert_eq!(
        commit.snapshot().name().unwrap().unwrap().as_str(),
        "new-design"
    );

    let undone = commit.undo(commit.snapshot()).unwrap();
    assert_eq!(undone.name().unwrap().unwrap().as_str(), "old-design");
    let redone = commit.redo(&undone).unwrap();
    assert_eq!(redone.bytes(), commit.snapshot().bytes());

    let rollback = source.edit().rollback();
    assert_eq!(rollback.bytes(), source.bytes());

    let mut clear = commit.snapshot().edit();
    assert!(clear.clear_name().unwrap());
    let cleared = clear.commit().unwrap();
    assert!(cleared.snapshot().name().unwrap().is_none());
    assert_eq!(
        cleared.snapshot().record().children[2].record_type_raw,
        0x7abc
    );
}

#[test]
fn rejects_wrong_context_duplicate_headers_utf16_and_bounds() {
    assert!(
        Snapshot::from_record(master(vec![name_record(2, "one"), name_record(2, "two"),])).is_err()
    );

    let mut odd = name_record(2, "bad");
    odd.data = vec![0];
    odd.data_length = 1;
    assert!(Snapshot::from_record(master(vec![odd])).is_err());

    let invalid_utf16 = atom(RecordType::CString, 0, 2, &[0, 0xd8]);
    assert!(Snapshot::from_record(master(vec![invalid_utf16])).is_err());

    let wrong_header = atom(RecordType::CString, 1, 2, &utf16("bad"));
    assert!(Snapshot::from_record(master(vec![wrong_header])).is_err());

    let oversized = atom(RecordType::CString, 0, 2, &vec![b'x'; MAX_NAME_BYTES + 2]);
    assert!(Snapshot::from_record(master(vec![oversized])).is_err());

    let title_children = vec![atom(RecordType::SlideAtom, 2, 0, &[0; 24])];
    let title = record(
        RecordType::Slide,
        0x0f,
        0,
        children_wire(&title_children),
        title_children,
    );
    assert!(Snapshot::from_record(title).is_err());
}

#[test]
fn failed_authoring_does_not_change_source() {
    let source = Snapshot::from_record(master(Vec::new())).unwrap();
    let original = source.bytes().to_vec();
    let mut edit = source.edit();
    assert!(edit.set_name("x".repeat(MAX_NAME_BYTES / 2 + 1)).is_err());
    assert!(!edit.is_changed());
    assert_eq!(edit.snapshot().unwrap().bytes(), original.as_slice());
    assert_eq!(source.bytes(), original.as_slice());
}
