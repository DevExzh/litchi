#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::records::Record;

const XML: &[u8] =
    br#"<p:txStyles xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;

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

fn notes_master(extras: Vec<Record>) -> Record {
    let children = vec![
        atom(RecordType::NotesAtom, 1, &[0; 8]),
        atom(RecordType::PPDrawing, 0x0f, &[0xaa, 0xbb]),
    ]
    .into_iter()
    .chain(extras)
    .collect::<Vec<_>>();
    let data = children.iter().flat_map(wire).collect();
    record(RecordType::Notes, 0x0f, 0, data, children)
}

fn unknown(raw: u16, data: &[u8]) -> Record {
    let mut record = record(RecordType::Unknown, 0, 7, data.to_vec(), Vec::new());
    record.record_type_raw = raw;
    record
}

fn contextual_master(context: Context) -> Record {
    let (record_type, first) = match context {
        Context::Main => (
            RecordType::MainMaster,
            atom(RecordType::SlideAtom, 2, &[0; 24]),
        ),
        Context::Title => {
            let mut data = [0u8; 24];
            data[0..4].copy_from_slice(&2u32.to_le_bytes());
            data[12..16].copy_from_slice(&0x8000_0000u32.to_le_bytes());
            (RecordType::Slide, atom(RecordType::SlideAtom, 2, &data))
        },
        Context::Notes => (RecordType::Notes, atom(RecordType::NotesAtom, 1, &[0; 8])),
        Context::Handout => (
            RecordType::Handout,
            atom(RecordType::PPDrawing, 0x0f, &[0xaa]),
        ),
    };
    let mut children = vec![first];
    if !matches!(context, Context::Handout) {
        children.push(atom(RecordType::PPDrawing, 0x0f, &[0xaa]));
    }
    let data = children.iter().flat_map(wire).collect();
    record(record_type, 0x0f, 0, data, children)
}

#[test]
fn builds_bounded_package_and_preserves_exact_snapshot_bytes() {
    let first = Styles::from_xml(XML).unwrap();
    let second = Styles::from_xml(XML).unwrap();
    assert_eq!(first.bytes(), second.bytes());
    assert_eq!(first.package().part_count(), 1);
    assert_eq!(
        first.package().xml_part_name(),
        "/drs/slideMasters/slideMaster1.xml"
    );

    let parsed = Styles::from_package(first.bytes().to_vec()).unwrap();
    assert_eq!(parsed, first);
}

#[test]
fn authors_notes_master_styles_atomically_and_preserves_unknown_tail() {
    let opaque = unknown(0x7abc, &[1, 2, 3]);
    let source =
        super::super::Snapshot::from_record(Context::Notes, notes_master(vec![opaque.clone()]))
            .unwrap();
    let original = source.bytes().to_vec();
    let styles = Styles::from_xml(XML).unwrap();

    let mut edit = source.edit();
    edit.set_notes_styles(styles.clone()).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.bytes(), original.as_slice());
    assert_eq!(commit.changes().changes().len(), 1);
    assert_eq!(commit.snapshot().notes_styles().unwrap(), Some(styles));
    assert_eq!(commit.snapshot().record().children[2], opaque);

    let undone = commit.undo(commit.snapshot()).unwrap();
    assert!(undone.notes_styles().unwrap().is_none());
    let redone = commit.redo(&undone).unwrap();
    assert_eq!(redone.bytes(), commit.snapshot().bytes());
}

#[test]
fn clears_and_rejects_wrong_context_without_mutating_the_editor() {
    let styles = Styles::from_xml(XML).unwrap();
    let source =
        super::super::Snapshot::from_record(Context::Notes, notes_master(Vec::new())).unwrap();
    let mut edit = source.edit();
    edit.set_notes_styles(styles).unwrap();
    let committed = edit.commit().unwrap();

    let mut clear = committed.snapshot().edit();
    assert!(clear.clear_notes_styles().unwrap());
    let cleared = clear.commit().unwrap();
    assert!(cleared.notes_styles().unwrap().is_none());

    for context in [Context::Main, Context::Title, Context::Handout] {
        let snapshot =
            super::super::Snapshot::from_record(context, contextual_master(context)).unwrap();
        let original = snapshot.bytes().to_vec();
        let mut rejected_edit = snapshot.edit();
        assert!(
            rejected_edit
                .set_notes_styles(Styles::from_xml(XML).unwrap())
                .is_err()
        );
        assert!(!rejected_edit.is_changed());
        assert_eq!(
            rejected_edit.snapshot().unwrap().bytes(),
            original.as_slice()
        );
    }
}

#[test]
fn rejects_malformed_and_oversized_xml_before_authoring() {
    assert!(Styles::from_xml(b"<p:wrong/>\n").is_err());
    assert!(Styles::from_xml(b"<p:txStyles").is_err());
    assert!(Styles::from_xml(vec![b'x'; MAX_XML_BYTES + 1]).is_err());

    let mut malformed = notes_master(Vec::new());
    malformed.children.push(record(
        RecordType::RoundTripNotesMasterTextStyles12Atom,
        1,
        0,
        vec![1, 2, 3],
        Vec::new(),
    ));
    malformed.data = malformed.children.iter().flat_map(wire).collect();
    malformed.data_length = u32::try_from(malformed.data.len()).unwrap();
    assert!(super::super::Snapshot::from_record(Context::Notes, malformed).is_err());
}
