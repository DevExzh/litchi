use super::*;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::package::IWorkPackage;
use crate::protobuf::tsp::Reference;
use crate::protobuf::{kn, tp, tswp};
use prost::Message;

fn empty_inputs() -> (Bundle, ObjectIndex) {
    let bytes = crate::package::IWorkPackage::new()
        .to_bytes()
        .expect("an empty package is serializable");
    let bundle = Bundle::from_bytes(&bytes).expect("an empty package is a valid ZIP bundle");
    let object_index = ObjectIndex::from_bundle(&bundle).expect("empty bundle indexes cleanly");
    (bundle, object_index)
}

fn bundle_with_archives(
    archives: impl IntoIterator<Item = (&'static str, Archive)>,
) -> (Bundle, ObjectIndex) {
    let mut package = IWorkPackage::new();
    for (name, archive) in archives {
        package
            .replace_archive(name, &archive)
            .expect("synthetic archive should be accepted");
    }
    let bytes = package
        .to_bytes()
        .expect("synthetic package should serialize");
    let bundle = Bundle::from_bytes(&bytes).expect("synthetic package should parse");
    let object_index = ObjectIndex::from_bundle(&bundle).expect("synthetic package should index");
    (bundle, object_index)
}

fn object<T: Message>(identifier: u64, message_type: u32, message: T) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )
    .expect("synthetic object should be valid")
}

fn raw_object(identifier: u64, message_type: u32, data: Vec<u8>) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )
    .expect("synthetic object should be valid")
}

fn reference(identifier: u64) -> Reference {
    Reference {
        identifier,
        ..Default::default()
    }
}

#[test]
fn each_focused_extractor_returns_leaf_owned_empty_results() {
    let (bundle, object_index) = empty_inputs();

    let tables: Vec<litchi_numbers::Table> = extract_tables(&bundle, &object_index).unwrap();

    assert!(tables.is_empty());
    assert!(matches!(
        extract_slides(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(_))
    ));
    assert!(matches!(
        extract_sections(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(_))
    ));
}

#[test]
fn empty_pages_root_is_a_valid_empty_section() {
    let root = object(1, 10_000, tp::DocumentArchive::default());
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![root],
        },
    )]);

    let sections = extract_sections(&bundle, &object_index).unwrap();

    assert_eq!(sections.len(), 1);
    assert!(sections[0].all_text().is_empty());
    assert!(sections[0].paragraphs().is_empty());
    let structured = extract_all(&bundle, &object_index).unwrap();
    assert_eq!(structured.summary(), "Tables: 0, Slides: 0, Sections: 1");
}

#[test]
fn pages_root_and_referenced_storage_fail_closed() {
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![raw_object(1, 10_000, vec![0x80])],
        },
    )]);
    assert!(matches!(
        extract_sections(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(message)) if message.contains("root payload is invalid")
    ));

    let root = object(
        1,
        10_000,
        tp::DocumentArchive {
            body_storage: Some(reference(42)),
            ..Default::default()
        },
    );
    let body = raw_object(42, 2_001, vec![0x80]);
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![root, body],
        },
    )]);
    assert!(matches!(
        extract_sections(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(message)) if message.contains("text storage payload is invalid")
    ));
}

#[test]
fn empty_keynote_show_is_a_valid_empty_presentation() {
    let root = object(1, 1, kn::DocumentArchive::default());
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![root],
        },
    )]);
    assert!(extract_slides(&bundle, &object_index).unwrap().is_empty());

    let root = object(
        1,
        1,
        kn::DocumentArchive {
            show: reference(2),
            ..Default::default()
        },
    );
    let show = object(
        2,
        2,
        kn::ShowArchive {
            slide_tree: kn::SlideTreeArchive::default(),
            ..Default::default()
        },
    );
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![root, show],
        },
    )]);

    let slides = extract_slides(&bundle, &object_index).unwrap();

    assert!(slides.is_empty());
    assert!(extract_all(&bundle, &object_index).unwrap().is_empty());
}

#[test]
fn keynote_required_chain_does_not_return_partial_slides() {
    let root = object(
        1,
        1,
        kn::DocumentArchive {
            show: reference(2),
            ..Default::default()
        },
    );
    let show = object(
        2,
        2,
        kn::ShowArchive {
            slide_tree: kn::SlideTreeArchive {
                slides: vec![reference(3)],
                ..Default::default()
            },
            ..Default::default()
        },
    );
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![root, show],
        },
    )]);

    assert!(matches!(
        extract_slides(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(message)) if message.contains("slide-tree node object 3 is missing")
    ));
}

#[test]
fn keynote_invalid_root_and_dangling_text_are_typed_errors() {
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![raw_object(1, 1, vec![0x80])],
        },
    )]);
    assert!(matches!(
        extract_slides(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(message)) if message.contains("root payload is invalid")
    ));

    let root = object(
        1,
        1,
        kn::DocumentArchive {
            show: reference(2),
            ..Default::default()
        },
    );
    let show = object(
        2,
        2,
        kn::ShowArchive {
            slide_tree: kn::SlideTreeArchive {
                slides: vec![reference(3)],
                ..Default::default()
            },
            ..Default::default()
        },
    );
    let node = object(
        3,
        4,
        kn::SlideNodeArchive {
            slide: Some(reference(4)),
            ..Default::default()
        },
    );
    let slide = object(
        4,
        5,
        kn::SlideArchive {
            body_placeholder: Some(reference(7)),
            ..Default::default()
        },
    );
    let placeholder = object(
        7,
        7,
        kn::PlaceholderArchive {
            super_: tswp::ShapeInfoArchive {
                owned_storage: Some(reference(8)),
                ..Default::default()
            },
            ..Default::default()
        },
    );
    let (bundle, object_index) = bundle_with_archives([(
        "Index/Document.iwa",
        Archive {
            objects: vec![root, show, node, slide, placeholder],
        },
    )]);

    assert!(matches!(
        extract_slides(&bundle, &object_index),
        Err(crate::Error::InvalidFormat(message)) if message.contains("object 8 is missing")
    ));
}

#[test]
fn extract_all_keeps_empty_application_results_independent() {
    let (bundle, object_index) = empty_inputs();
    let structured = extract_all(&bundle, &object_index).unwrap();

    assert!(structured.is_empty());
    assert_eq!(structured.summary(), "Tables: 0, Slides: 0, Sections: 0");
    assert!(structured.all_text().is_empty());
}

#[test]
fn numbers_table_creation_uses_the_leaf_model() {
    let mut builder =
        litchi_numbers::Table::builder("Test Table", litchi_numbers::Dimensions::new(2, 2));
    assert!(
        builder
            .set(
                litchi_numbers::Position::new(0, 0),
                litchi_numbers::cell::Value::Text("Header 1".to_owned()),
            )
            .is_ok()
    );
    assert!(
        builder
            .set(
                litchi_numbers::Position::new(0, 1),
                litchi_numbers::cell::Value::Text("Header 2".to_owned()),
            )
            .is_ok()
    );
    assert!(
        builder
            .set(
                litchi_numbers::Position::new(1, 0),
                litchi_numbers::cell::Value::number(42.0).expect("finite cell number"),
            )
            .is_ok()
    );
    assert!(
        builder
            .set(
                litchi_numbers::Position::new(1, 1),
                litchi_numbers::cell::Value::Boolean(true),
            )
            .is_ok()
    );

    let table = builder.finish().expect("valid leaf table");
    assert_eq!(table.name(), "Test Table");
    assert_eq!(table.row_count(), 2);
    assert_eq!(table.column_count(), 2);
    assert_eq!(table.cell_count(), 4);
    assert!(table.to_csv().contains("Header 1"));
}

#[test]
fn keynote_slide_creation_preserves_leaf_text_order() {
    let mut builder = litchi_keynote::Slide::builder(0);
    builder.set_title(Some("Introduction".to_owned()));
    builder.push_text("Point 1".to_owned());
    builder.push_text("Point 2".to_owned());
    builder.set_notes(Some("Speaker notes".to_owned()));

    let slide = builder.build();
    assert_eq!(slide.index(), 0);
    assert_eq!(
        slide.all_text(),
        ["Introduction", "Point 1", "Point 2", "Speaker notes"]
    );
}

#[test]
fn pages_section_creation_preserves_leaf_text_order() {
    let mut section_builder = litchi_pages::Section::builder(0, litchi_pages::SectionType::Body);
    section_builder.set_heading(Some("Chapter 1".to_owned()));
    section_builder.push_paragraph("First paragraph.".to_owned());
    section_builder.push_paragraph("Second paragraph.".to_owned());
    let section = section_builder.build();

    assert_eq!(
        section.all_text(),
        ["Chapter 1", "First paragraph.", "Second paragraph."]
    );
}

#[test]
fn structured_text_aggregation_does_not_change_order() {
    let table = litchi_numbers::Table::new("Data", litchi_numbers::Dimensions::new(1, 1));
    let mut slide_builder = litchi_keynote::Slide::builder(0);
    slide_builder.set_title(Some("Title".to_owned()));
    slide_builder.push_text("Body".to_owned());
    let slide = slide_builder.build();
    let mut section_builder = litchi_pages::Section::builder(0, litchi_pages::SectionType::Body);
    section_builder.set_heading(Some("Heading".to_owned()));
    let section = section_builder.build();

    let data = StructuredData::from_parts(vec![table], vec![slide], vec![section])
        .expect("structured semantic values should form a valid snapshot");

    assert_eq!(data.all_text(), ["Table: Data", "Title", "Body", "Heading"]);
    assert_eq!(data.summary(), "Tables: 1, Slides: 1, Sections: 1");
    assert_eq!(data.table(0).map(litchi_numbers::Table::name), Some("Data"));
    assert_eq!(data.slide(0).map(litchi_keynote::Slide::index), Some(0));
    assert_eq!(data.section(0).map(litchi_pages::Section::index), Some(0));
    assert!(data.table(1).is_none());
    assert!(data.slide(1).is_none());
    assert!(data.section(1).is_none());
}
