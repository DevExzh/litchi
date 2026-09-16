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
        data: data.into(),
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
        powerpoint_document: std::sync::Arc::new(powerpoint_document),
        parser,
        persist_mapping,
        slide_directory: SlideDirectory::new_for_test(0),
        pictures_data: None,
        pictures_source: None,
        record_limits: crate::RecordLimits::default(),
    }
}

#[test]
fn lazy_live_document_reuses_presentation_record_limits() {
    let powerpoint_document = record_bytes(0x0f, 0, RecordType::Document, &[1]);
    let mut persist_mapping = PersistMapping::new();
    persist_mapping.add_mapping(1, 0);
    let presentation = Presentation {
        powerpoint_document: std::sync::Arc::new(powerpoint_document),
        parser: RecordParser::new(),
        persist_mapping,
        slide_directory: SlideDirectory::new_for_test(0),
        pictures_data: None,
        pictures_source: None,
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
    show.children[1]
        .data
        .to_mut()
        .extend_from_slice(&[0xAA, 0xBB, 0xCC]);
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

/// Fixtures whose record trees exercise the per-slide re-parse: a large
/// multi-slide deck, a deck with speaker notes, and a small one.
const SHARED_SOURCE_FIXTURES: [&str; 4] = [
    "poi/test-data/slideshow/45543.ppt",
    "poi/test-data/slideshow/headers_footers_2007.ppt",
    "ole/ppt/SampleShow.ppt",
    "poi/test-data/slideshow/41246-1.ppt",
];

fn open_fixture(
    relative: &str,
) -> Option<(crate::Package<std::io::Cursor<Vec<u8>>>, Presentation)> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative);
    let bytes = std::fs::read(path).ok()?;
    let mut package = crate::Package::from_reader(std::io::Cursor::new(bytes)).ok()?;
    let presentation = package.presentation().ok()?;
    Some((package, presentation))
}

fn owned_factory(presentation: &Presentation) -> crate::slide::SlideFactory<'_> {
    crate::slide::SlideFactory::new_with_limits(
        presentation.powerpoint_document.as_slice(),
        &presentation.persist_mapping,
        &presentation.slide_directory,
        presentation.record_limits,
    )
}

fn shared_factory(presentation: &Presentation) -> crate::slide::SlideFactory<'_> {
    crate::slide::SlideFactory::new_shared_with_limits(
        &presentation.powerpoint_document,
        &presentation.persist_mapping,
        &presentation.slide_directory,
        presentation.record_limits,
    )
}

/// Every payload a borrowing per-slide re-parse produces is the value the
/// copying one produced, down to the speaker-notes page and the full slide
/// text, on real fixtures.
#[test]
fn the_shared_and_owned_slide_factories_parse_identical_slides() {
    let mut checked = 0usize;
    let mut opened = 0usize;
    for relative in SHARED_SOURCE_FIXTURES {
        let Some((_package, presentation)) = open_fixture(relative) else {
            continue;
        };
        opened += 1;
        let owned = owned_factory(&presentation);
        let shared = shared_factory(&presentation);
        assert_eq!(owned.slide_ids(), shared.slide_ids(), "{relative}");

        for (index, persist_id) in owned.slide_ids().into_iter().enumerate() {
            let owned_slide = owned.parse_slide(persist_id);
            let shared_slide = shared.parse_slide(persist_id);
            match (owned_slide, shared_slide) {
                (Ok(owned_data), Ok(shared_data)) => {
                    assert_eq!(owned_data.record, shared_data.record, "{relative}");
                    assert_eq!(owned_data.offset, shared_data.offset, "{relative}");
                    assert_eq!(owned_data.slide_id, shared_data.slide_id, "{relative}");
                    assert_eq!(
                        format!("{:?}", owned_data.note_descriptor),
                        format!("{:?}", shared_data.note_descriptor),
                        "{relative}"
                    );
                    assert_eq!(owned_data.doc_data(), shared_data.doc_data(), "{relative}");
                    let owned_slide = crate::Slide::from_slide_data(owned_data, index + 1);
                    let shared_slide = crate::Slide::from_slide_data(shared_data, index + 1);
                    assert_eq!(
                        format!("{:?}", owned_slide.text()),
                        format!("{:?}", shared_slide.text()),
                        "{relative}"
                    );
                    assert_eq!(
                        format!("{:?}", owned_slide.shapes()),
                        format!("{:?}", shared_slide.shapes()),
                        "{relative}"
                    );
                    assert_eq!(
                        format!(
                            "{:?}",
                            owned_slide
                                .speaker_notes()
                                .map(|notes| notes.map(|value| value.text().map(str::to_string)))
                        ),
                        format!(
                            "{:?}",
                            shared_slide
                                .speaker_notes()
                                .map(|notes| notes.map(|value| value.text().map(str::to_string)))
                        ),
                        "{relative}"
                    );
                    checked += 1;
                },
                (owned_error, shared_error) => {
                    assert_eq!(
                        format!("{:?}", owned_error.map(|data| data.offset)),
                        format!("{:?}", shared_error.map(|data| data.offset)),
                        "{relative}"
                    );
                },
            }
        }
    }
    assert_eq!(
        opened,
        SHARED_SOURCE_FIXTURES.len(),
        "every fixture must be readable"
    );
    assert!(checked >= opened, "every fixture must produce a slide");
}

/// `max_copied_payload_bytes` is charged per record whether or not the payload
/// is copied, so a borrowing re-parse refuses exactly the slides a copying one
/// refused.
#[test]
fn the_shared_slide_factory_charges_the_same_copy_budget() {
    fn payload_bytes(record: &Record) -> usize {
        record.data.len() + record.children.iter().map(payload_bytes).sum::<usize>()
    }

    let mut checked = 0usize;
    for relative in SHARED_SOURCE_FIXTURES {
        let Some((_package, presentation)) = open_fixture(relative) else {
            continue;
        };
        let Some(&persist_id) = owned_factory(&presentation).slide_ids().first() else {
            continue;
        };
        let Ok(slide) = owned_factory(&presentation).parse_slide(persist_id) else {
            continue;
        };
        let total = payload_bytes(&slide.record);
        assert!(total > 0, "{relative}");

        for budget in [total, total - 1] {
            let limits = crate::RecordLimits {
                max_copied_payload_bytes: budget,
                ..presentation.record_limits
            };
            let owned = crate::slide::SlideFactory::new_with_limits(
                presentation.powerpoint_document.as_slice(),
                &presentation.persist_mapping,
                &presentation.slide_directory,
                limits,
            )
            .parse_slide(persist_id);
            let shared = crate::slide::SlideFactory::new_shared_with_limits(
                &presentation.powerpoint_document,
                &presentation.persist_mapping,
                &presentation.slide_directory,
                limits,
            )
            .parse_slide(persist_id);
            assert_eq!(
                owned.is_ok(),
                shared.is_ok(),
                "{relative}: copy budget {budget} must refuse identically"
            );
            assert_eq!(
                owned.is_ok(),
                budget == total,
                "{relative}: copy budget {budget}"
            );
            if let (Err(owned_error), Err(shared_error)) = (owned, shared) {
                assert_eq!(
                    format!("{owned_error:?}"),
                    format!("{shared_error:?}"),
                    "{relative}"
                );
            }
        }
        checked += 1;
    }
    assert_eq!(
        checked,
        SHARED_SOURCE_FIXTURES.len(),
        "every fixture must charge a slide budget"
    );
}

/// The borrowing per-slide re-parse never writes through to the stream it
/// spans, and its siblings keep their bytes.
#[test]
fn editing_a_shared_slide_payload_leaves_the_document_stream_intact() {
    for relative in SHARED_SOURCE_FIXTURES {
        let Some((_package, presentation)) = open_fixture(relative) else {
            continue;
        };
        let Some(&persist_id) = shared_factory(&presentation).slide_ids().first() else {
            continue;
        };
        let Ok(mut slide) = shared_factory(&presentation).parse_slide(persist_id) else {
            continue;
        };
        let stream = presentation.powerpoint_document.as_ref().clone();
        let sibling = slide
            .record
            .children
            .first()
            .map(|child| child.data.to_vec());
        slide.record.data.to_mut()[0] ^= 0xFF;
        assert_eq!(
            presentation.powerpoint_document.as_ref(),
            &stream,
            "{relative}: editing a borrowed payload must not reach the stream"
        );
        if let Some(sibling) = sibling {
            assert_eq!(
                slide
                    .record
                    .children
                    .first()
                    .map(|child| child.data.to_vec()),
                Some(sibling),
                "{relative}"
            );
        }
        return;
    }
    panic!("no fixture produced a slide to edit");
}
