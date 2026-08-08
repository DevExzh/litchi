use litchi_odt::elements::parser::OrderElement;
use litchi_odt::flat::Document;

const FLAT_TEXT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"><office:meta><meta:generator>LibreOffice/24.2</meta:generator><dc:title>Flat &amp; Exact</dc:title></office:meta><office:body><office:text><text:h text:outline-level="1">Lorem ipsum dolor sit amet</text:h><text:p>Lorem <text:span>Ipsum</text:span></text:p><table:table table:name="Table1"><table:table-row><table:table-cell><text:p>One</text:p></table:table-cell></table:table-row></table:table><text:p>Cras eu leo sed justo</text:p><table:table table:name="Table2"><table:table-row><table:table-cell><text:p>Two</text:p></table:table-cell></table:table-row></table:table></office:text></office:body></office:document>"#;
const REAL_FLAT_TEXT: &[u8] =
    include_bytes!("../../../test-data/odf/odt/font-face-declarations-flat.fodt");
const REAL_DEFAULT_PAGE_LAYOUT: &[u8] =
    include_bytes!("../../../test-data/odf/odt/libreoffice-page-columns-compact.fodt");
const REAL_LINKED_OBJECT: &[u8] =
    include_bytes!("../../../test-data/odf/odt/libreoffice-linked-object-compact.fodt");
const REAL_USER_FIELD: &[u8] =
    include_bytes!("../../../test-data/odf/odt/libreoffice-user-field-compact.fodt");
const EXACT_MACRO_METADATA: &str = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" o:mimetype="application/vnd.oasis.opendocument.text"><o:scripts><o:script s:language="ooo:Basic"><ooo:libraries><ooo:library name="Standard"/></ooo:libraries></o:script></o:scripts><o:body><o:text><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">Body</text:p></o:text></o:body></o:document>"#;

#[test]
fn flat_text_document_exposes_paragraphs_headings_and_tables() {
    let document = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let text = document.text().unwrap();
    assert!(text.contains("Lorem Ipsum"));
    assert!(text.contains("Cras eu leo sed justo"));
    assert!(text.contains("Lorem ipsum dolor sit amet"));

    let elements = document.elements().unwrap();
    let paragraphs: Vec<_> = elements
        .iter()
        .filter_map(|element| match element {
            OrderElement::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    let headings: Vec<_> = elements
        .iter()
        .filter_map(|element| match element {
            OrderElement::Heading(heading) => Some(heading),
            _ => None,
        })
        .collect();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0].text().unwrap(), "Lorem Ipsum");
    assert_eq!(headings.len(), 1);
    assert_eq!(headings[0].text().unwrap(), "Lorem ipsum dolor sit amet");

    let tables = document.tables().unwrap();
    assert_eq!(tables.len(), 2);
    assert_eq!(tables[0].name(), Some("Table1"));
    assert_eq!(tables[1].name(), Some("Table2"));
}

#[test]
fn flat_text_document_exposes_metadata_and_exact_bytes() {
    let document = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let metadata = document.metadata().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Flat & Exact"));
    assert_eq!(metadata.application.as_deref(), Some("LibreOffice/24.2"));
    assert_eq!(document.as_bytes(), FLAT_TEXT.as_bytes());
    assert_eq!(document.to_bytes(), FLAT_TEXT.as_bytes());
}

#[test]
fn flat_text_document_rejects_garbage_empty_and_packaged_input() {
    assert!(Document::from_bytes(b"not xml at all".to_vec()).is_err());
    assert!(Document::from_bytes(Vec::new()).is_err());
    assert!(Document::from_bytes(b"PK\x03\x04mimetype".to_vec()).is_err());
}

#[test]
fn compact_real_fodt_opens_reads_and_saves_byte_exactly() {
    let document = Document::from_bytes(REAL_FLAT_TEXT.to_vec()).unwrap();
    assert_eq!(document.text().unwrap(), "Body");
    assert_eq!(document.as_bytes(), REAL_FLAT_TEXT);
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("font-face-declarations-flat.fodt");
    document.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), REAL_FLAT_TEXT);
    assert_eq!(Document::open(path).unwrap().as_bytes(), REAL_FLAT_TEXT);
}

#[test]
fn compact_real_fodt_fixtures_expose_layout_objects_and_declarations() {
    for bytes in [
        REAL_DEFAULT_PAGE_LAYOUT,
        REAL_LINKED_OBJECT,
        REAL_USER_FIELD,
    ] {
        litchi_odf_common::compact_xml::validate(bytes).unwrap();
    }

    let layout = Document::from_bytes(REAL_DEFAULT_PAGE_LAYOUT.to_vec())
        .unwrap()
        .default_page_layout()
        .unwrap()
        .unwrap();
    assert!(layout.name.is_empty());
    assert!(layout.properties.is_some());

    let objects = Document::from_bytes(REAL_LINKED_OBJECT.to_vec())
        .unwrap()
        .embedded_objects()
        .unwrap();
    assert_eq!(objects.len(), 1);
    assert!(matches!(
        &objects[0].source,
        litchi_odt::embedded::Source::Linked { href }
            if href == "http://192.0.2.1:12345/probe"
    ));

    let declarations = Document::from_bytes(REAL_USER_FIELD.to_vec())
        .unwrap()
        .variable_declarations()
        .unwrap();
    assert!(
        declarations
            .declarations()
            .any(|declaration| declaration.name() == "user-field-decl-name-example")
    );
}

#[test]
fn compact_exact_fodt_exposes_inert_macro_metadata() {
    litchi_odf_common::compact_xml::validate(EXACT_MACRO_METADATA.as_bytes()).unwrap();
    let scripts = Document::from_bytes(EXACT_MACRO_METADATA.as_bytes().to_vec())
        .unwrap()
        .document_scripts()
        .unwrap()
        .unwrap();
    assert_eq!(scripts.scripts.len(), 1);
    assert_eq!(scripts.scripts[0].language, "ooo:Basic");
    assert!(scripts.scripts[0].content_xml.contains("ooo:library"));
}
