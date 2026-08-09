#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::model::Presentation;
use crate::Error;
use crate::consts::RecordType;
use crate::parsers::RecordParser;
use crate::persist::PersistMapping;
use crate::records::Record;
use crate::slide::SlideDirectory;

fn record(record_type: RecordType, data: Vec<u8>, children: Vec<Record>) -> Record {
    Record {
        record_type,
        record_type_raw: 0,
        version: 0,
        instance: 0,
        data_length: u32::try_from(data.len()).unwrap(),
        data,
        children,
    }
}

fn record_bytes(version: u16, instance: u16, record_type: RecordType, data: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + data.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.as_u16().to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(data);
    bytes
}

#[cfg(feature = "vba-inspection")]
fn presentation_with_vba_storage() -> Presentation {
    let mut atom_data = Vec::new();
    atom_data.extend_from_slice(&41u32.to_le_bytes());
    atom_data.extend_from_slice(&1u32.to_le_bytes());
    atom_data.extend_from_slice(&2u32.to_le_bytes());
    let atom = record_bytes(2, 0, RecordType::VBAInfoAtom, &atom_data);
    let vba_info = record_bytes(0x0f, 1, RecordType::VBAInfo, &atom);

    let mut storage_data = Vec::new();
    storage_data.extend_from_slice(&4096u32.to_le_bytes());
    storage_data.extend_from_slice(&[0x78, 0x9c, 1, 2, 3]);
    let storage = record_bytes(0, 1, RecordType::ExternalOleObjectStg, &storage_data);
    let storage_offset = u32::try_from(vba_info.len()).unwrap();

    let mut powerpoint_document = vba_info;
    powerpoint_document.extend_from_slice(&storage);
    let mut parser = RecordParser::new();
    parser.parse_document(&powerpoint_document).unwrap();
    let mut persist_mapping = PersistMapping::new();
    persist_mapping.add_mapping(41, storage_offset);

    Presentation {
        powerpoint_document,
        parser,
        persist_mapping,
        slide_directory: SlideDirectory::new_for_test(0),
        pictures_data: None,
        record_limits: crate::RecordLimits::default(),
    }
}

#[test]
fn lazy_live_document_reuses_presentation_record_limits() {
    let powerpoint_document = record_bytes(0x0f, 0, RecordType::Document, &[1]);
    let mut persist_mapping = PersistMapping::new();
    persist_mapping.add_mapping(1, 0);
    let presentation = Presentation {
        powerpoint_document,
        parser: RecordParser::new(),
        persist_mapping,
        slide_directory: SlideDirectory::new_for_test(0),
        pictures_data: None,
        record_limits: crate::RecordLimits {
            max_record_payload_bytes: 0,
            ..crate::RecordLimits::default()
        },
    };

    assert!(matches!(
        presentation.live_document_record(),
        Err(Error::ResourceLimit(_))
    ));
}

fn named_shows(children: Vec<Record>) -> Record {
    record(RecordType::NamedShows, Vec::new(), children)
}

fn named_show(name: &str, slide_ids: &[u32]) -> Record {
    let name_bytes: Vec<u8> = name.encode_utf16().flat_map(u16::to_le_bytes).collect();
    let slide_bytes: Vec<u8> = slide_ids.iter().flat_map(|id| id.to_le_bytes()).collect();
    record(
        RecordType::NamedShow,
        Vec::new(),
        vec![
            record(RecordType::CString, name_bytes, Vec::new()),
            record(RecordType::NamedShowSlides, slide_bytes, Vec::new()),
        ],
    )
}

#[test]
fn parses_named_shows_container() {
    let container = named_shows(vec![
        named_show("Demo Show", &[0x101, 0x103]),
        named_show("Short", &[0x100]),
    ]);

    let mut shows = Vec::new();
    Presentation::parse_named_shows(&container, &mut shows);

    assert_eq!(shows.len(), 2);
    assert_eq!(shows[0].name, "Demo Show");
    assert_eq!(shows[0].slide_indices, vec![1, 3]);
    assert_eq!(shows[1].name, "Short");
    assert_eq!(shows[1].slide_indices, vec![0]);
}

#[test]
fn ignores_trailing_partial_slide_id_bytes() {
    let mut show = named_show("Odd", &[0x102]);
    // Append 3 stray bytes to the NamedShowSlides atom.
    show.children[1].data.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
    let container = named_shows(vec![show]);

    let mut shows = Vec::new();
    Presentation::parse_named_shows(&container, &mut shows);

    assert_eq!(shows.len(), 1);
    assert_eq!(shows[0].slide_indices, vec![2]);
}

#[test]
fn skips_named_show_without_name() {
    let show = record(
        RecordType::NamedShow,
        Vec::new(),
        vec![record(
            RecordType::NamedShowSlides,
            0x101u32.to_le_bytes().to_vec(),
            Vec::new(),
        )],
    );
    let container = named_shows(vec![show]);

    let mut shows = Vec::new();
    Presentation::parse_named_shows(&container, &mut shows);
    assert!(shows.is_empty());
}

#[test]
#[cfg(feature = "vba-inspection")]
fn vba_project_storage_returns_only_outer_metadata() {
    let presentation = presentation_with_vba_storage();

    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert_eq!(storage.persist_id_ref(), 41);
    assert!(storage.has_macros());
    assert!(storage.has_persisted_storage());
    assert_eq!(storage.stored_payload_len(), Some(5));
    assert_eq!(storage.declared_uncompressed_len(), Some(4096));
    assert_eq!(
        storage.compression(),
        Some(crate::embedded::storage::Compression::Zlib)
    );
    assert!(storage.may_contain_macro_code());
    assert_eq!(presentation.vba_info().unwrap(), Some(storage.info()));
}
