use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn atom(record_type: RecordType, version: u16, instance: u16, data: Vec<u8>) -> Record {
    Record {
        record_type,
        record_type_raw: record_type.as_u16(),
        version,
        instance,
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    }
}

fn opaque(raw_type: u16, version: u16, instance: u16, data: Vec<u8>) -> Record {
    Record {
        record_type: RecordType::Unknown,
        record_type_raw: raw_type,
        version,
        instance,
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    }
}

fn container(
    record_type: RecordType,
    version: u16,
    instance: u16,
    children: Vec<Record>,
) -> Record {
    let mut record = atom(record_type, version, instance, Vec::new());
    record.children = children;
    record
}

fn document_atom() -> Record {
    atom(RecordType::DocumentAtom, 1, 0, vec![0; 40])
}

fn end_document() -> Record {
    atom(RecordType::EndDocument, 0, 0, Vec::new())
}

fn custom_table_styles(data: &[u8]) -> Record {
    atom(
        RecordType::RoundTripCustomTableStyles12Atom,
        0,
        0,
        data.to_vec(),
    )
}

fn master_persist(persist_id: u32, master_id: u32) -> Record {
    let mut data = vec![0; 20];
    data[0..4].copy_from_slice(&persist_id.to_le_bytes());
    data[12..16].copy_from_slice(&master_id.to_le_bytes());
    atom(RecordType::SlidePersistAtom, 0, 0, data)
}

fn slide_persist(persist_id: u32, slide_id: u32, text_count: u32) -> Record {
    let mut data = vec![0; 20];
    data[0..4].copy_from_slice(&persist_id.to_le_bytes());
    data[8..12].copy_from_slice(&text_count.to_le_bytes());
    data[12..16].copy_from_slice(&slide_id.to_le_bytes());
    atom(RecordType::SlidePersistAtom, 0, 0, data)
}

fn text_header() -> Record {
    atom(RecordType::TextHeaderAtom, 0, 0, vec![0; 4])
}

fn document(children: Vec<Record>) -> Record {
    container(RecordType::Document, 0x0f, 0, children)
}

fn sample_document(with_styles: bool) -> Record {
    let masters = container(
        RecordType::SlideListWithText,
        0x0f,
        1,
        vec![
            opaque(0x7a01, 3, 9, vec![1, 2, 3]),
            master_persist(10, 0x8000_0010),
            master_persist(11, 0x8000_0020),
        ],
    );
    let slides = container(
        RecordType::SlideListWithText,
        0x0f,
        0,
        vec![
            opaque(0x7a02, 2, 4, vec![9, 8]),
            slide_persist(20, 100, 1),
            text_header(),
            opaque(0x7a03, 1, 2, vec![7, 6, 5]),
            slide_persist(21, 200, 0),
        ],
    );
    let mut children = vec![
        opaque(0x7a00, 2, 17, vec![0xde, 0xad, 0xbe, 0xef]),
        document_atom(),
        masters,
        slides,
    ];
    if with_styles {
        children.push(custom_table_styles(b"opaque-style-package"));
    }
    children.push(end_document());
    document(children)
}

#[test]
fn accepts_both_defined_custom_table_style_placements() {
    let prefix = document_atom();
    let end = end_document();
    let styles = custom_table_styles(&[]);
    assert_eq!(
        DocumentStructure::parse(&document(vec![prefix.clone(), styles.clone(), end.clone()]))
            .unwrap()
            .custom_table_styles,
        Some(CustomTableStylesPlacement::BeforeEndDocument)
    );
    assert_eq!(
        DocumentStructure::parse(&document(vec![prefix, end, styles]))
            .unwrap()
            .custom_table_styles,
        Some(CustomTableStylesPlacement::AfterEndDocument)
    );
}

#[test]
fn rejects_missing_duplicate_nonempty_and_nonterminal_end_records() {
    let end = end_document();
    assert!(DocumentStructure::parse(&document(vec![document_atom()])).is_err());
    assert!(
        DocumentStructure::parse(&document(vec![document_atom(), end.clone(), end.clone()]))
            .is_err()
    );
    assert!(
        DocumentStructure::parse(&document(vec![
            document_atom(),
            atom(RecordType::EndDocument, 0, 0, vec![0]),
        ]))
        .is_err()
    );
    assert!(
        DocumentStructure::parse(&document(vec![
            document_atom(),
            end,
            opaque(0x7711, 0, 0, vec![1]),
        ]))
        .is_err()
    );
}

#[test]
fn source_snapshot_is_exact_on_noop_and_retains_unknown_atoms() {
    let source = Snapshot::from_record(sample_document(true)).unwrap();
    let parsed = Snapshot::parse(source.bytes()).unwrap();

    assert_eq!(parsed.bytes(), source.bytes());
    let unknown = parsed.unknown_atoms().collect::<Vec<_>>();
    assert_eq!(unknown.len(), 1);
    assert_eq!(unknown[0].record_type_raw, 0x7a00);
    assert_eq!(unknown[0].version, 2);
    assert_eq!(unknown[0].instance, 17);
    assert_eq!(unknown[0].data, [0xde, 0xad, 0xbe, 0xef]);

    let commit = parsed.edit().commit().unwrap();
    assert_eq!(commit.snapshot().bytes(), parsed.bytes());
    assert_eq!(commit.snapshot().revision(), parsed.revision());
    assert!(commit.patch().is_empty());
    assert!(commit.patch().changes().is_empty());
}

#[test]
fn typed_inventories_validate_and_reorder_owned_groups() {
    let source = Snapshot::from_record(sample_document(false)).unwrap();
    assert_eq!(
        source
            .masters()
            .iter()
            .map(|master| master.master_id())
            .collect::<Vec<_>>(),
        [0x8000_0010, 0x8000_0020]
    );
    assert_eq!(
        source
            .slides()
            .iter()
            .map(|slide| slide.slide_id())
            .collect::<Vec<_>>(),
        [100, 200]
    );

    let mut edit = source.edit();
    edit.move_slide_id(100, 1).unwrap();
    edit.move_master(0, 1).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .slides()
            .iter()
            .map(|slide| slide.slide_id())
            .collect::<Vec<_>>(),
        [200, 100]
    );
    assert_eq!(
        commit
            .snapshot()
            .masters()
            .iter()
            .map(|master| master.master_id())
            .collect::<Vec<_>>(),
        [0x8000_0020, 0x8000_0010]
    );
    assert_eq!(commit.patch().changes().len(), 2);

    let slide_list = commit.snapshot().record().children[3].children.as_slice();
    assert_eq!(slide_list[0].record_type_raw, 0x7a02);
    assert_eq!(slide_list[1].data[12..16], 200u32.to_le_bytes());
    assert_eq!(slide_list[3].data, [0; 4]);
    assert_eq!(slide_list[4].data, [7, 6, 5]);

    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );
}

#[test]
fn invalid_relationships_and_failed_edits_are_atomic() {
    let duplicate = document(vec![
        document_atom(),
        container(
            RecordType::SlideListWithText,
            0x0f,
            0,
            vec![slide_persist(20, 100, 0), slide_persist(21, 100, 0)],
        ),
        end_document(),
    ]);
    assert!(Snapshot::from_record(duplicate).is_err());

    let out_of_order = document(vec![
        document_atom(),
        container(
            RecordType::SlideListWithText,
            0x0f,
            0,
            vec![slide_persist(20, 100, 0)],
        ),
        container(
            RecordType::SlideListWithText,
            0x0f,
            1,
            vec![master_persist(10, 0x8000_0010)],
        ),
        end_document(),
    ]);
    assert!(Snapshot::from_record(out_of_order).is_err());

    let source = Snapshot::from_record(sample_document(false)).unwrap();
    let mut edit = source.edit();
    let before = edit.record().clone();
    assert!(edit.reorder_slides(&[0, 0]).is_err());
    assert_eq!(edit.record(), &before);
    assert!(!edit.is_changed());
    assert!(edit.changes().is_empty());
}

#[test]
fn stale_patches_are_rejected_and_custom_styles_are_reversible() {
    let source = Snapshot::from_record(sample_document(true)).unwrap();
    let mut edit = source.edit();
    edit.move_custom_table_styles(CustomTableStylesPlacement::AfterEndDocument)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().structure().custom_table_styles,
        Some(CustomTableStylesPlacement::AfterEndDocument)
    );
    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );

    let stale = Snapshot::from_record(sample_document(false)).unwrap();
    assert!(commit.patch().apply(&stale).is_err());
}
