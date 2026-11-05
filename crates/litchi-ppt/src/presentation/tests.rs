#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::model::{Presentation, PresentationData};
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
    parser
        .parse_document(&powerpoint_document, &crate::SourceLimits::default())
        .unwrap();
    let mut persist_mapping = PersistMapping::new();
    persist_mapping.add_mapping(41, storage_offset);

    #[cfg(feature = "encryption")]
    let encryption = None;
    Presentation::from_data(PresentationData {
        powerpoint_document: std::sync::Arc::new(powerpoint_document),
        parser,
        persist_mapping,
        slide_directory: SlideDirectory::new_for_test(0),
        pictures_data: None,
        pictures_source: None,
        source_limits: crate::SourceLimits::default(),
        #[cfg(feature = "vba-inspection")]
        vba_project_limits: crate::VbaProjectLimits::default(),
        tolerance: crate::leniency::ToleranceReport::default(),
        source: None,
        #[cfg(feature = "encryption")]
        encryption,
        shape_memo: crate::slide::ShapeMemo::default(),
    })
}

#[test]
fn lazy_live_document_reuses_presentation_source_limits() {
    let powerpoint_document = record_bytes(0x0f, 0, RecordType::Document, &[1]);
    let mut persist_mapping = PersistMapping::new();
    persist_mapping.add_mapping(1, 0);
    #[cfg(feature = "encryption")]
    let encryption = None;
    let presentation = Presentation::from_data(PresentationData {
        powerpoint_document: std::sync::Arc::new(powerpoint_document),
        parser: RecordParser::new(),
        persist_mapping,
        slide_directory: SlideDirectory::new_for_test(0),
        pictures_data: None,
        pictures_source: None,
        source_limits: crate::SourceLimits {
            max_record_payload_bytes: 0,
            ..crate::SourceLimits::default()
        },
        #[cfg(feature = "vba-inspection")]
        vba_project_limits: crate::VbaProjectLimits::default(),
        tolerance: crate::leniency::ToleranceReport::default(),
        source: None,
        #[cfg(feature = "encryption")]
        encryption,
        shape_memo: crate::slide::ShapeMemo::default(),
    });

    assert!(matches!(
        presentation.live_document_record(),
        Err(Error::Limit {
            kind: crate::SourceLimitKind::RecordPayloadBytes,
            ..
        })
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
) -> Option<(crate::raw::Reader<std::io::Cursor<Vec<u8>>>, Presentation)> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative);
    let bytes = std::fs::read(path).ok()?;
    let mut package = crate::raw::Reader::from_reader(std::io::Cursor::new(bytes)).ok()?;
    let presentation = package.presentation().ok()?;
    Some((package, presentation))
}

fn owned_factory(presentation: &Presentation) -> crate::slide::SlideFactory<'_> {
    crate::slide::SlideFactory::new_with_limits(
        presentation.inner.powerpoint_document.as_slice(),
        &presentation.inner.persist_mapping,
        &presentation.inner.slide_directory,
        presentation.inner.source_limits,
    )
}

fn shared_factory(presentation: &Presentation) -> crate::slide::SlideFactory<'_> {
    crate::slide::SlideFactory::new_shared_with_limits(
        &presentation.inner.powerpoint_document,
        &presentation.inner.persist_mapping,
        &presentation.inner.slide_directory,
        presentation.inner.source_limits,
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
            let limits = crate::SourceLimits {
                max_copied_payload_bytes: budget,
                ..presentation.inner.source_limits
            };
            let owned = crate::slide::SlideFactory::new_with_limits(
                presentation.inner.powerpoint_document.as_slice(),
                &presentation.inner.persist_mapping,
                &presentation.inner.slide_directory,
                limits,
            )
            .parse_slide(persist_id);
            let shared = crate::slide::SlideFactory::new_shared_with_limits(
                &presentation.inner.powerpoint_document,
                &presentation.inner.persist_mapping,
                &presentation.inner.slide_directory,
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
        let stream = presentation.inner.powerpoint_document.as_ref().clone();
        let sibling = slide
            .record
            .children
            .first()
            .map(|child| child.data.to_vec());
        slide.record.data.to_mut()[0] ^= 0xFF;
        assert_eq!(
            presentation.inner.powerpoint_document.as_ref(),
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

/// L-PPT-R2: the notes index is built once per presentation, at open, from
/// the `DocumentContainer` the slide directory parsed. Selecting a slide
/// parses that slide only: a factory over a stream whose `DocumentContainer`
/// header no longer reads still finds every slide's notes page, which a
/// factory that re-parsed the document for its own notes index could not.
#[test]
fn slide_selection_reads_notes_from_the_index_built_at_open() {
    use crate::slide::{Slide, SlideFactory};

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow/basic_test_ppt_file.ppt");
    let package = crate::Package::open(path).unwrap();
    let presentation = package.presentation().unwrap();
    let expected = [
        "These are the notes for page 1",
        "These are the notes on page two, again lacking formatting",
    ];

    let mut document = presentation.inner.powerpoint_document.as_ref().clone();
    let offset = presentation.inner.slide_directory.document_offset();
    document[offset + 2..offset + 4].copy_from_slice(&0xffffu16.to_le_bytes());
    let factory = SlideFactory::new(
        &document,
        &presentation.inner.persist_mapping,
        &presentation.inner.slide_directory,
    );
    for (index, slide_data) in factory.slides().enumerate() {
        let slide = Slide::from_slide_data(slide_data.unwrap(), index + 1);
        let notes = slide.speaker_notes().unwrap().expect("notes page");
        assert_eq!(notes.text().unwrap(), expected[index]);
    }

    // Each `slide_at` builds its own factory and shares the same index.
    for (position, text) in expected.iter().enumerate() {
        let slide = presentation.slide_at(position).unwrap().expect("slide");
        let notes = slide.speaker_notes().unwrap().expect("notes page");
        assert_eq!(notes.text().unwrap(), *text);
    }
}

/// Record 0260: the presentation's one-slide memo shares the most recently
/// parsed slide's shape array between handles, and is invisible otherwise.
#[test]
fn slide_handles_share_the_memoized_shape_array() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/ppt/SampleShow.ppt");
    let package = crate::Package::open(path).unwrap();
    let presentation = package.presentation().unwrap();
    let texts = |shapes: &[crate::shapes::ShapeEnum<'static>]| {
        shapes
            .iter()
            .map(|shape| shape.text().unwrap())
            .collect::<Vec<_>>()
    };

    let first = presentation
        .slide_at(0)
        .unwrap()
        .unwrap()
        .shared_shapes()
        .unwrap();
    let again = presentation
        .slide_at(0)
        .unwrap()
        .unwrap()
        .shared_shapes()
        .unwrap();
    assert!(std::sync::Arc::ptr_eq(&first, &again));
    // A clone of the projection shares the memo.
    let cloned = package.presentation().unwrap();
    let shared = cloned
        .slide_at(0)
        .unwrap()
        .unwrap()
        .shared_shapes()
        .unwrap();
    assert!(std::sync::Arc::ptr_eq(&first, &shared));

    // Another slide replaces the entry; the first slide is parsed again and
    // reads exactly as before.
    let second = presentation
        .slide_at(1)
        .unwrap()
        .unwrap()
        .shared_shapes()
        .unwrap();
    assert!(!std::sync::Arc::ptr_eq(&first, &second));
    let reparsed = presentation
        .slide_at(0)
        .unwrap()
        .unwrap()
        .shared_shapes()
        .unwrap();
    assert!(!std::sync::Arc::ptr_eq(&first, &reparsed));
    assert_eq!(texts(&first), texts(&reparsed));

    // `slides()` handles consult the same memo, and the repairs a slide
    // reports do not depend on which handle parsed it.
    let slides = presentation.slides().unwrap();
    assert!(std::sync::Arc::ptr_eq(
        &slides[0].shared_shapes().unwrap(),
        &reparsed
    ));
    let fresh = crate::Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt/SampleShow.ppt"),
    )
    .unwrap()
    .presentation()
    .unwrap();
    for position in 0..presentation.slide_count() {
        let memoized = presentation.slide_at(position).unwrap().unwrap();
        let independent = fresh.slide_at(position).unwrap().unwrap();
        assert_eq!(
            memoized.tolerance_report().unwrap(),
            independent.tolerance_report().unwrap()
        );
        assert_eq!(
            texts(memoized.shapes().unwrap()),
            texts(independent.shapes().unwrap())
        );
    }
}
