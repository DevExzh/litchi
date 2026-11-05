#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::path::PathBuf;

use litchi_core::Position;
use litchi_ppt::master::{Kind, Objects, parse};
use litchi_ppt::master_layout::{self, Context, Path, Snapshot};
use litchi_ppt::persist::PersistMapping;
use litchi_ppt::records::Record;
use litchi_ppt::{Package, RecordType};

const UNKNOWN_TYPE: u16 = 0x7abc;

fn atom(record_type: RecordType, version: u16, instance: u16, data: &[u8]) -> Record {
    Record {
        record_type,
        record_type_raw: record_type.as_u16(),
        version,
        instance,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.into(),
        children: Vec::new(),
    }
}

fn unknown(data: &[u8]) -> Record {
    Record {
        record_type: RecordType::Unknown,
        record_type_raw: UNKNOWN_TYPE,
        version: 0,
        instance: 7,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.into(),
        children: Vec::new(),
    }
}

fn wire(record: &Record) -> Vec<u8> {
    let packed = record.version | (record.instance << 4);
    let mut bytes = Vec::with_capacity(8 + record.data.len());
    bytes.extend_from_slice(&packed.to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&record.data_length.to_le_bytes());
    bytes.extend_from_slice(&record.data);
    bytes
}

fn container(record_type: RecordType, children: Vec<Record>) -> Record {
    let data = children.iter().flat_map(wire).collect::<Vec<_>>();
    Record {
        record_type,
        record_type_raw: record_type.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.into(),
        children,
    }
}

fn slide_atom(geometry: u32, master_id: u32) -> Record {
    let mut data = [0u8; 24];
    data[0..4].copy_from_slice(&geometry.to_le_bytes());
    data[12..16].copy_from_slice(&master_id.to_le_bytes());
    atom(RecordType::SlideAtom, 2, 0, &data)
}

fn master_persist(persist_id: u32, master_id: u32) -> Record {
    let mut data = [0u8; 20];
    data[0..4].copy_from_slice(&persist_id.to_le_bytes());
    data[12..16].copy_from_slice(&master_id.to_le_bytes());
    atom(RecordType::SlidePersistAtom, 0, 0, &data)
}

fn notes_atom() -> Record {
    atom(RecordType::NotesAtom, 1, 0, &[0; 8])
}

fn drawing() -> Record {
    atom(RecordType::PPDrawing, 0x0f, 0, &[0x10, 0x20, 0x30])
}

fn document(notes_id: u32, handout_id: u32, references: Vec<Record>) -> Record {
    let mut data = [0u8; 40];
    data[24..28].copy_from_slice(&notes_id.to_le_bytes());
    data[28..32].copy_from_slice(&handout_id.to_le_bytes());
    let document_atom = atom(RecordType::DocumentAtom, 1, 0, &data);
    let mut list = container(RecordType::SlideListWithText, references);
    list.instance = 1;
    list.data = list.children.iter().flat_map(wire).collect();
    list.data_length = u32::try_from(list.data.len()).unwrap();
    container(RecordType::Document, vec![document_atom, list])
}

fn mapping(ids: &[u32]) -> PersistMapping {
    let mut mapping = PersistMapping::new();
    for (index, id) in ids.iter().copied().enumerate() {
        mapping.add_mapping(id, u32::try_from(index * 128).unwrap());
    }
    mapping
}

#[test]
fn inventories_all_four_masters_contextually() {
    let main = container(
        RecordType::MainMaster,
        vec![slide_atom(1, 0), drawing(), unknown(&[0xaa, 0xbb, 0xcc])],
    );
    let title = container(RecordType::Slide, vec![slide_atom(2, 0x8000_0000)]);
    let notes = container(RecordType::Notes, vec![notes_atom()]);
    let handout = container(RecordType::Handout, Vec::new());
    let root = document(
        4,
        5,
        vec![
            master_persist(2, 0x8000_0000),
            master_persist(3, 0x8000_0001),
        ],
    );
    let objects =
        Objects::from_records([(2, &main), (3, &title), (4, &notes), (5, &handout)]).unwrap();

    let inventory = parse(&root, &objects, &mapping(&[2, 3, 4, 5])).unwrap();
    assert_eq!(inventory.main().len(), 1);
    assert_eq!(inventory.title().len(), 1);
    assert_eq!(inventory.notes().unwrap().persist().id(), 4);
    assert_eq!(inventory.handout().unwrap().persist().id(), 5);
    assert_eq!(inventory.title()[0].based_on().master_id(), 0x8000_0000);
}

#[test]
fn masters_retain_unknown_records_and_package_edits_are_atomic() {
    let opaque = unknown(&[0xaa, 0xbb, 0xcc]);
    let main = container(
        RecordType::MainMaster,
        vec![slide_atom(1, 0), drawing(), opaque.clone()],
    );
    let root = document(0, 0, vec![master_persist(2, 0x8000_0000)]);
    let objects = Objects::from_records([(2, &main)]).unwrap();
    let inventory = parse(&root, &objects, &mapping(&[2])).unwrap();
    assert_eq!(
        inventory.main()[0].unknown()[0].wire().unwrap(),
        wire(&opaque)
    );

    let snapshot =
        Snapshot::from_record(Context::Main, main, &master_layout::Limits::default()).unwrap();
    assert!(
        snapshot
            .bytes()
            .windows(wire(&opaque).len())
            .any(|window| window == wire(&opaque))
    );

    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/ppt/SampleShow.ppt");
    let source = Package::from_bytes(&std::fs::read(path).unwrap()).unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source
        .edit_master_layout(Kind::Main, Position::new(0))
        .unwrap();
    assert_eq!(edit.context().unwrap(), Context::Main);
    let appended = edit.record().unwrap().children.len();
    edit.append(Path::root(), opaque.clone()).unwrap();
    let committed = edit.commit().unwrap();
    assert!(committed.diagnostics().changed());
    let published = committed
        .package()
        .edit_master_layout(Kind::Main, Position::new(0))
        .unwrap();
    assert_eq!(published.record().unwrap().children[appended], opaque);

    // A candidate that is no longer a valid master is refused at commit and
    // leaves the source untouched.
    let mut invalid = source
        .edit_master_layout(Kind::Main, Position::new(0))
        .unwrap();
    let atom = invalid
        .record()
        .unwrap()
        .children
        .iter()
        .position(|child| child.record_type == RecordType::SlideAtom)
        .unwrap();
    invalid.remove(Path::root().child(atom)).unwrap();
    assert!(matches!(
        invalid.commit(),
        Err(master_layout::Error::Refused(_))
    ));
    assert_eq!(source.to_bytes().unwrap(), before);
}
