use super::*;
use crate::consts::RecordType;
use crate::persist::PersistMapping;
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

fn atom(record_type: RecordType, version: u16, data: &[u8]) -> Record {
    record(record_type, version, 0, data.to_vec(), Vec::new())
}

fn slide_atom(master: u32, notes: u32) -> Record {
    let mut data = vec![0u8; 24];
    data[12..16].copy_from_slice(&master.to_le_bytes());
    data[16..20].copy_from_slice(&notes.to_le_bytes());
    atom(RecordType::SlideAtom, 2, &data)
}

fn master_persist(persist: u32, master: u32) -> Record {
    let mut data = vec![0u8; 20];
    data[0..4].copy_from_slice(&persist.to_le_bytes());
    data[12..16].copy_from_slice(&master.to_le_bytes());
    atom(RecordType::SlidePersistAtom, 0, &data)
}

fn notes_master() -> Record {
    let notes_atom = atom(RecordType::NotesAtom, 1, &[0; 8]);
    record(
        RecordType::Notes,
        0x0f,
        0,
        wire(&notes_atom),
        vec![notes_atom],
    )
}

fn handout_master() -> Record {
    record(RecordType::Handout, 0x0f, 0, Vec::new(), Vec::new())
}

fn main_master() -> Record {
    let unknown = unknown(0x7abc, &[1, 2, 3, 4]);
    let atom = slide_atom(0, 0);
    let children = vec![atom, unknown];
    let data = children_wire(&children);
    record(RecordType::MainMaster, 0x0f, 0, data, children)
}

fn title_master() -> Record {
    let atom = slide_atom(0x8000_0000, 0);
    record(RecordType::Slide, 0x0f, 0, wire(&atom), vec![atom])
}

fn document(notes: u32, handout: u32, masters: Vec<Record>) -> Record {
    let master_list = record(
        RecordType::SlideListWithText,
        0x0f,
        1,
        children_wire(&masters),
        masters,
    );
    let mut atom_data = vec![0u8; 40];
    atom_data[24..28].copy_from_slice(&notes.to_le_bytes());
    atom_data[28..32].copy_from_slice(&handout.to_le_bytes());
    let atom = atom(RecordType::DocumentAtom, 1, &atom_data);
    let children = vec![atom, master_list];
    record(
        RecordType::Document,
        0x0f,
        0,
        children_wire(&children),
        children,
    )
}

fn children_wire(children: &[Record]) -> Vec<u8> {
    children.iter().flat_map(wire).collect()
}

fn wire(record: &Record) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + record.data.len());
    bytes.extend_from_slice(&(record.version | (record.instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&record.data_length.to_le_bytes());
    bytes.extend_from_slice(&record.data);
    bytes
}

fn mapping(ids: &[u32]) -> PersistMapping {
    let mut mapping = PersistMapping::new();
    for (index, id) in ids.iter().copied().enumerate() {
        mapping.add_mapping(id, (index * 128) as u32);
    }
    mapping
}

#[test]
fn inventories_all_four_contextual_master_kinds() {
    let main = main_master();
    let title = title_master();
    let notes = notes_master();
    let handout = handout_master();
    let list_entries = vec![
        master_persist(2, 0x8000_0000),
        master_persist(3, 0x8000_0001),
    ];
    let root = document(4, 5, list_entries);
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
fn unknown_records_retain_reference_and_wire_bytes() {
    let main = main_master();
    let root = document(0, 0, vec![master_persist(2, 0x8000_0000)]);
    let objects = Objects::from_records([(2, &main)]).unwrap();
    let inventory = parse(&root, &objects, &mapping(&[2])).unwrap();

    let unknown = inventory.main()[0].unknown();
    assert_eq!(unknown.len(), 1);
    assert_eq!(unknown[0].raw_type(), 0x7abc);
    assert_eq!(unknown[0].bytes(), &[1, 2, 3, 4]);
    assert_eq!(unknown[0].wire().unwrap(), wire(&main.children[1]));
    assert_eq!(
        inventory.main()[0].reference().data,
        root.children[1].children[0].data
    );
    assert_eq!(inventory.main()[0].record().data, main.data);
}

#[test]
fn missing_or_wrong_persist_targets_fail_before_inventory_publication() {
    let main = main_master();
    let root = document(0, 0, vec![master_persist(2, 0x8000_0000)]);
    let objects = Objects::from_records([(2, &main)]).unwrap();

    assert!(parse(&root, &objects, &mapping(&[9])).is_err());

    let wrong = handout_master();
    let objects = Objects::from_records([(2, &wrong)]).unwrap();
    assert!(parse(&root, &objects, &mapping(&[2])).is_err());
}

#[test]
fn invalid_master_identity_and_title_base_are_rejected() {
    let main = main_master();
    let root = document(0, 0, vec![master_persist(2, 7)]);
    let objects = Objects::from_records([(2, &main)]).unwrap();
    assert!(parse(&root, &objects, &mapping(&[2])).is_err());

    let title = {
        let atom = slide_atom(0x8000_00ff, 0);
        record(RecordType::Slide, 0x0f, 0, wire(&atom), vec![atom])
    };
    let root = document(0, 0, vec![master_persist(2, 0x8000_0000)]);
    let objects = Objects::from_records([(2, &title)]).unwrap();
    assert!(parse(&root, &objects, &mapping(&[2])).is_err());
}

#[test]
fn object_catalog_rejects_duplicate_and_null_ids() {
    let main = main_master();
    assert!(Objects::from_records([(0, &main)]).is_err());
    assert!(Objects::from_records([(2, &main), (2, &main)]).is_err());
}
