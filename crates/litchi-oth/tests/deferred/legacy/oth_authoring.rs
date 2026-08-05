//! Authoring for legacy `.oth` Writer/Web templates: create from scratch,
//! save-as-template conversion, and edit round-trips through the reader.

use litchi_odf::{
    Document, DocumentBuilder, MutableWebDocument, Family, Package,
    OwnedPackage, WebDocument, WebDocumentBuilder, constants,
};

const ODT_FIXTURE: &[u8] = include_bytes!("../../../test-data/odfdo/tests/samples/variable.odt");

fn assert_web_package_sanity(bytes: &[u8]) {
    // The mimetype entry is the first, uncompressed ZIP entry.
    assert_eq!(&bytes[..4], b"PK\x03\x04");
    assert_eq!(&bytes[30..38], b"mimetype");
    assert_eq!(
        &bytes[38..38 + constants::ODF_WEB.len()],
        constants::ODF_WEB.as_bytes()
    );

    // Package-level MIME type and manifest root entry agree.
    let package = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(package.mimetype().unwrap(), constants::ODF_WEB);
    let manifest = package.package().unwrap();
    assert_eq!(
        manifest.manifest().get_media_type("/"),
        Some(constants::ODF_WEB)
    );

    // The format-neutral reader classifies it as a Web template.
    let generic = Package::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(generic.family(), Family::Web);
    assert!(generic.is_template());
}

#[test]
fn create_web_template_from_scratch_round_trips() {
    let mut builder = WebDocumentBuilder::new();
    builder
        .builder_mut()
        .add_heading("Portal template", 1)
        .unwrap();
    builder
        .builder_mut()
        .add_paragraph("Reusable web body text.")
        .unwrap();

    let bytes = builder.build().unwrap();
    assert_web_package_sanity(&bytes);

    let document = WebDocument::from_bytes(bytes).unwrap();
    assert_eq!(document.mimetype(), constants::ODF_WEB);
    assert!(document.is_template());
    assert_eq!(
        document.text().unwrap(),
        "Portal template\nReusable web body text."
    );
}

#[test]
fn build_document_returns_validated_web_document() {
    let mut builder = WebDocumentBuilder::new();
    builder
        .builder_mut()
        .add_paragraph("Only paragraph")
        .unwrap();
    let document = builder.build_document().unwrap();
    assert_eq!(document.document().paragraph_count().unwrap(), 1);
    assert_eq!(document.text().unwrap(), "Only paragraph");
}

#[test]
fn convert_existing_text_document_to_web_template() {
    let document = Document::from_bytes(ODT_FIXTURE.to_vec()).unwrap();
    assert!(!document.paragraphs().unwrap().is_empty());

    let mut mutable = MutableWebDocument::from_document(document).unwrap();
    mutable
        .document_mut()
        .update_paragraph(0, "Converted to web template")
        .unwrap();

    let bytes = mutable.to_bytes().unwrap();
    assert_web_package_sanity(&bytes);

    let reopened = WebDocument::from_bytes(bytes).unwrap();
    let text = reopened.text().unwrap();
    assert!(text.starts_with("Converted to web template"));
    // Untouched paragraphs survive the conversion.
    assert!(text.contains("This document"));
}

#[test]
fn open_edit_save_reopen_round_trip() {
    let mut builder = WebDocumentBuilder::new();
    builder
        .builder_mut()
        .add_heading("Draft heading", 1)
        .unwrap();
    builder.builder_mut().add_paragraph("Keep me").unwrap();
    let document = builder.build_document().unwrap();

    let mut mutable = document.into_mutable().unwrap();
    mutable
        .document_mut()
        .update_paragraph(0, "Edited heading")
        .unwrap();

    let path = std::env::temp_dir().join("litchi-oth-authoring-roundtrip.oth");
    mutable.save(&path).unwrap();

    let reopened = WebDocument::open(&path).unwrap();
    let text = reopened.text().unwrap();
    // The heading is untouched; the edited paragraph was replaced.
    assert!(text.contains("Draft heading"));
    assert!(text.contains("Edited heading"));
    assert_eq!(reopened.mimetype(), constants::ODF_WEB);
    std::fs::remove_file(&path).unwrap();
}

#[test]
fn to_mutable_leaves_original_template_untouched() {
    let mut builder = WebDocumentBuilder::new();
    builder
        .builder_mut()
        .add_paragraph("Original text")
        .unwrap();
    let bytes = builder.build().unwrap();
    let document = WebDocument::from_bytes(bytes.clone()).unwrap();

    let mut mutable = document.to_mutable().unwrap();
    mutable
        .document_mut()
        .update_paragraph(0, "Changed copy")
        .unwrap();
    let changed = mutable.to_bytes().unwrap();
    assert!(
        WebDocument::from_bytes(changed)
            .unwrap()
            .text()
            .unwrap()
            .contains("Changed copy")
    );

    // The original template is byte-identical and unchanged.
    assert_eq!(document.as_bytes(), bytes.as_slice());
    assert_eq!(document.text().unwrap(), "Original text");
}

#[test]
fn web_authoring_rejects_non_text_documents() {
    // A non-text package cannot be converted through the text editor.
    let spreadsheet = litchi_odf::FlatSpreadsheet::from_bytes(
        include_bytes!("../../../test-data/libreoffice-core/sc/qa/unit/data/draw-image-link.fods")
            .to_vec(),
    );
    assert!(spreadsheet.is_ok());
    // WebDocument still rejects standard text MIME packages.
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("plain odt").unwrap();
    let odt = builder.build().unwrap();
    assert!(WebDocument::from_bytes(odt).is_err());
}
