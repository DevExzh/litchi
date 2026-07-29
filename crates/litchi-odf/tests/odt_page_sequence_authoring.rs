//! Authoring of explicit ODT page sequences (`text:page-sequence`, ODF 1.3
//! §5.3) through `DocumentBuilder` and `MutableDocument`.
//!
//! The model preserves only the ordered `text:master-page-name` assignments;
//! litchi never paginates or resolves the referenced master pages.

use litchi_odf::{Document, DocumentBuilder, MutableDocument, OdtPageSequence};

fn sequence() -> OdtPageSequence {
    OdtPageSequence::new(vec![
        "First".to_string(),
        "Left".to_string(),
        "Right".to_string(),
    ])
    .unwrap()
}

fn content_xml(bytes: &[u8]) -> String {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut entry = archive.by_name("content.xml").unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut entry, &mut xml).unwrap();
    xml
}

#[test]
fn builder_authored_page_sequence_round_trips_the_package() {
    let mut builder = DocumentBuilder::new();
    builder.set_page_sequence(Some(sequence())).unwrap();
    builder.add_paragraph("Body text").unwrap();
    let bytes = builder.build().unwrap();

    let xml = content_xml(&bytes);
    let sequence_start = xml.find("<text:page-sequence").unwrap();
    let paragraph_start = xml.find("<text:p>").unwrap();
    assert!(
        sequence_start < paragraph_start,
        "page-sequence must be the first office:text child"
    );
    assert!(xml.contains(r#"<text:page text:master-page-name="Left"/>"#));

    let document = Document::from_bytes(bytes).unwrap();
    assert_eq!(document.page_sequence().unwrap(), Some(sequence()));
    // Clearing works through the builder as well.
    let mut builder = DocumentBuilder::new();
    builder.set_page_sequence(Some(sequence())).unwrap();
    builder.set_page_sequence(None).unwrap();
    builder.add_paragraph("Body text").unwrap();
    let document = Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(document.page_sequence().unwrap(), None);
}

#[test]
fn builder_rejects_invalid_page_sequences() {
    let mut builder = DocumentBuilder::new();
    assert!(OdtPageSequence::new(Vec::new()).is_err());
    assert!(OdtPageSequence::new(vec![String::new()]).is_err());
    assert!(
        builder
            .set_page_sequence(Some(OdtPageSequence {
                master_page_names: vec!["Bad\nName".to_string()],
            }))
            .is_err()
    );
    // The rejected value was not stored; the build has no sequence.
    let document = Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(document.page_sequence().unwrap(), None);
}

#[test]
fn mutable_sets_replaces_and_removes_page_sequence() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Existing body").unwrap();
    let document = Document::from_bytes(builder.build().unwrap()).unwrap();
    let mut mutable = MutableDocument::from_document(document).unwrap();
    assert_eq!(mutable.page_sequence().unwrap(), None);

    // Insert ahead of the existing content.
    mutable.set_page_sequence(Some(&sequence())).unwrap();
    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.page_sequence().unwrap(), Some(sequence()));

    // Replace the assignments in place.
    let replacement = OdtPageSequence::new(vec!["Standard".to_string()]).unwrap();
    let mut mutable = MutableDocument::from_document(reopened).unwrap();
    mutable.set_page_sequence(Some(&replacement)).unwrap();
    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.page_sequence().unwrap(), Some(replacement));

    // Remove the sequence; removing again is a no-op.
    let mut mutable = MutableDocument::from_document(reopened).unwrap();
    mutable.set_page_sequence(None).unwrap();
    mutable.set_page_sequence(None).unwrap();
    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.page_sequence().unwrap(), None);
    let xml = content_xml(&mutable.to_bytes().unwrap());
    assert!(!xml.contains("page-sequence"));
    // The body content survives every rewrite.
    assert!(xml.contains("Existing body"));
}
