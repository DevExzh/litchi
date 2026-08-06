use super::*;

fn empty_inputs() -> (Bundle, ObjectIndex) {
    let bytes = crate::package::IWorkPackage::new()
        .to_bytes()
        .expect("an empty package is serializable");
    let bundle = Bundle::from_bytes(&bytes).expect("an empty package is a valid ZIP bundle");
    let object_index = ObjectIndex::from_bundle(&bundle).expect("empty bundle indexes cleanly");
    (bundle, object_index)
}

#[test]
fn each_focused_extractor_returns_leaf_owned_empty_results() {
    let (bundle, object_index) = empty_inputs();

    let tables: Vec<litchi_numbers::Table> = extract_tables(&bundle, &object_index).unwrap();
    let slides: Vec<litchi_keynote::Slide> = extract_slides(&bundle, &object_index).unwrap();
    let sections: Vec<litchi_pages::Section> = extract_sections(&bundle, &object_index).unwrap();

    assert!(tables.is_empty());
    assert!(slides.is_empty());
    assert!(sections.is_empty());
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
                litchi_numbers::cell::Value::Number(42.0),
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
    let mut section = litchi_pages::Section::new(0, litchi_pages::SectionType::Body);
    section.heading = Some("Chapter 1".to_owned());
    section.paragraphs.push("First paragraph.".to_owned());
    section.paragraphs.push("Second paragraph.".to_owned());

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
    let mut section = litchi_pages::Section::new(0, litchi_pages::SectionType::Body);
    section.heading = Some("Heading".to_owned());

    let data = StructuredData::from_parts(vec![table], vec![slide], vec![section])
        .expect("structured semantic values should form a valid snapshot");

    assert_eq!(data.all_text(), ["Table: Data", "Title", "Body", "Heading"]);
    assert_eq!(data.summary(), "Tables: 1, Slides: 1, Sections: 1");
    assert_eq!(data.table(0).map(litchi_numbers::Table::name), Some("Data"));
    assert_eq!(data.slide(0).map(litchi_keynote::Slide::index), Some(0));
    assert_eq!(data.section(0).map(|value| value.index), Some(0));
    assert!(data.table(1).is_none());
    assert!(data.slide(1).is_none());
    assert!(data.section(1).is_none());
}
