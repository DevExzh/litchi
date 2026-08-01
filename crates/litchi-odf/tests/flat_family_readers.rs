//! Semantic reading of flat OpenDocument XML documents through the packaged
//! family models, exercised against real producer fixtures.

use litchi_odf::elements::parser::DocumentOrderElement;
use litchi_odf::{
    CellValue, FlatChartDocument, FlatDrawingDocument, FlatImageDocument, FlatPresentation,
    FlatSpreadsheet, FlatTextDocument,
};

const FLAT_TEXT: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/pageColumns.fodt"
);
const FLAT_TEXT_TABLES: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/IndexElementsInHiddenSections.fodt"
);
const FLAT_TEXT_IMAGES: &str = include_str!("../../../test-data/odfdo/tests/samples/images.fodt");
const FLAT_SPREADSHEET: &str =
    include_str!("../../../test-data/odfdo/tests/samples/test_flat_lo.fods");
const FLAT_PRESENTATION: &str = include_str!(
    "../../../test-data/libreoffice-core/sd/qa/unit/tiledrendering/data/slide-background-link.fodp"
);
const FLAT_DRAWING: &str = include_str!(
    "../../../test-data/libreoffice-core/xmloff/qa/unit/data/tdf161327_LatheEndAngle.fodg"
);

#[test]
fn flat_text_document_exposes_paragraphs_and_headings() {
    let document = FlatTextDocument::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();

    let text = document.document().text().unwrap();
    assert!(text.contains("Lorem Ipsum"));
    assert!(text.contains("Lorem ipsum dolor sit amet"));

    let elements = document.document().elements().unwrap();
    let paragraphs: Vec<_> = elements
        .iter()
        .filter_map(|element| match element {
            DocumentOrderElement::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    let headings: Vec<_> = elements
        .iter()
        .filter_map(|element| match element {
            DocumentOrderElement::Heading(heading) => Some(heading),
            _ => None,
        })
        .collect();

    assert!(!paragraphs.is_empty());
    assert_eq!(paragraphs[0].text().unwrap(), "Lorem Ipsum");
    assert!(
        paragraphs
            .iter()
            .any(|paragraph| paragraph.text().unwrap().contains("Cras eu leo sed justo"))
    );
    assert!(headings.len() >= 2);
    assert!(headings.iter().any(|heading| {
        heading
            .text()
            .unwrap()
            .contains("Lorem ipsum dolor sit amet")
    }));
    assert!(
        headings
            .iter()
            .any(|heading| heading.text().unwrap().contains("Curabitur at sodales leo"))
    );
}

#[test]
fn flat_text_document_exposes_tables() {
    let document = FlatTextDocument::from_bytes(FLAT_TEXT_TABLES.as_bytes().to_vec()).unwrap();

    let tables = document.document().tables().unwrap();
    assert_eq!(tables.len(), 2);
    assert_eq!(tables[0].name(), Some("Table1"));
    assert_eq!(tables[1].name(), Some("Table2"));
}

#[test]
fn flat_text_document_exposes_metadata_and_round_trips() {
    let bytes = FLAT_TEXT_IMAGES.as_bytes().to_vec();
    let document = FlatTextDocument::from_bytes(bytes.clone()).unwrap();

    let metadata = document.metadata().unwrap();
    assert!(
        metadata
            .application
            .as_deref()
            .is_some_and(|generator| generator.contains("LibreOffice"))
    );

    // Read-only wrappers save the original flat bytes exactly.
    assert_eq!(document.as_bytes(), bytes.as_slice());
    assert_eq!(document.to_bytes(), bytes);
    let path = std::env::temp_dir().join("litchi-flat-semantic-roundtrip.fodt");
    document.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), FLAT_TEXT_IMAGES.as_bytes());
    std::fs::remove_file(&path).unwrap();
}

#[test]
fn flat_spreadsheet_exposes_sheets_cells_and_values() {
    let mut flat = FlatSpreadsheet::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).unwrap();
    let spreadsheet = flat.spreadsheet_mut();

    assert_eq!(spreadsheet.sheet_count().unwrap(), 2);

    let first = spreadsheet.sheet_by_index(0).unwrap().unwrap();
    assert_eq!(first.name().unwrap(), "Sheet1");
    let rows = first.rows().unwrap();
    assert_eq!(
        rows[0].cell(0).unwrap().unwrap().value().unwrap(),
        &CellValue::Text("test".to_string())
    );
    assert_eq!(
        rows[1].cell(1).unwrap().unwrap().value().unwrap(),
        &CellValue::Number(123.0)
    );

    let second = spreadsheet.sheet_by_name("Sheet2").unwrap().unwrap();
    let rows = second.rows().unwrap();
    assert_eq!(
        rows[0].cell(0).unwrap().unwrap().value().unwrap(),
        &CellValue::Text("abc".to_string())
    );

    let metadata = flat.metadata().unwrap();
    assert!(
        metadata
            .application
            .as_deref()
            .is_some_and(|generator| generator.contains("LibreOffice"))
    );
    assert_eq!(flat.as_bytes(), FLAT_SPREADSHEET.as_bytes());
}

#[test]
fn flat_spreadsheet_expands_repeated_cells() {
    let xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3">"#,
        r#"<office:body><office:spreadsheet><table:table table:name="R">"#,
        r#"<table:table-row>"#,
        r#"<table:table-cell office:value-type="float" office:value="7" "#,
        r#"table:number-columns-repeated="3"><text:p>7</text:p></table:table-cell>"#,
        r#"<table:table-cell office:value-type="string"><text:p>end</text:p></table:table-cell>"#,
        r#"</table:table-row>"#,
        r#"</table:table></office:spreadsheet></office:body></office:document>"#,
    );
    let mut flat = FlatSpreadsheet::from_bytes(xml.as_bytes().to_vec()).unwrap();
    let sheet = flat.spreadsheet_mut().sheet_by_name("R").unwrap().unwrap();
    let rows = sheet.rows().unwrap();
    let cells = rows[0].cells().unwrap();
    assert_eq!(cells.len(), 4);
    for cell in &cells[..3] {
        assert_eq!(cell.value().unwrap(), &CellValue::Number(7.0));
    }
    assert_eq!(
        cells[3].value().unwrap(),
        &CellValue::Text("end".to_string())
    );
}

#[test]
fn flat_presentation_exposes_slides() {
    let flat = FlatPresentation::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).unwrap();
    let presentation = flat.presentation();

    assert_eq!(presentation.slide_count().unwrap(), 1);
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 1);
    assert_eq!(slides[0].index(), 0);
    assert!(
        slides[0]
            .text()
            .unwrap()
            .contains("Slide with remote background image.")
    );

    assert_eq!(flat.as_bytes(), FLAT_PRESENTATION.as_bytes());
}

#[test]
fn flat_presentation_exposes_slide_notes() {
    // No corpus fixture ships parseable speaker notes in flat form, so this
    // minimal two-slide document exercises the same flat splitting path.
    let xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" "#,
        r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"office:mimetype="application/vnd.oasis.opendocument.presentation" office:version="1.3">"#,
        r#"<office:body><office:presentation>"#,
        r#"<draw:page draw:name="page1">"#,
        r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box>"#,
        r#"<text:p>First slide</text:p>"#,
        r#"</draw:text-box></draw:frame>"#,
        r#"</draw:page>"#,
        r#"<draw:page draw:name="page2">"#,
        r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box>"#,
        r#"<text:p>Second slide</text:p>"#,
        r#"</draw:text-box></draw:frame>"#,
        r#"<presentation:notes>"#,
        r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box>"#,
        r#"<text:p>Remember to breathe</text:p>"#,
        r#"</draw:text-box></draw:frame>"#,
        r#"</presentation:notes>"#,
        r#"</draw:page>"#,
        r#"</office:presentation></office:body></office:document>"#,
    );
    let flat = FlatPresentation::from_bytes(xml.as_bytes().to_vec()).unwrap();
    let slides = flat.presentation().slides().unwrap();
    assert_eq!(slides.len(), 2);
    assert!(slides[0].text().unwrap().contains("First slide"));
    assert!(slides[1].text().unwrap().contains("Second slide"));
    assert_eq!(slides[0].notes().unwrap(), None);
    assert_eq!(slides[1].notes().unwrap(), Some("Remember to breathe"));
    assert_eq!(flat.as_bytes(), xml.as_bytes());
}

#[test]
fn flat_drawing_exposes_pages() {
    let flat = FlatDrawingDocument::from_bytes(FLAT_DRAWING.as_bytes().to_vec()).unwrap();
    assert_eq!(flat.drawing().page_count(), 1);
    assert_eq!(flat.drawing().pages().len(), 1);
    assert_eq!(flat.as_bytes(), FLAT_DRAWING.as_bytes());
}

#[test]
fn flat_chart_exposes_chart_tree() {
    let xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3">"#,
        r#"<office:body><office:chart>"#,
        r#"<chart:chart chart:class="chart:bar">"#,
        r#"<chart:title><text:p>Flat Chart Title</text:p></chart:title>"#,
        r#"</chart:chart></office:chart></office:body></office:document>"#,
    );
    let flat = FlatChartDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
    assert!(flat.chart().text().contains("Flat Chart Title"));
    assert_eq!(flat.as_bytes(), xml.as_bytes());
}

#[test]
fn flat_image_exposes_frame() {
    let xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
        r#"office:mimetype="application/vnd.oasis.opendocument.image" office:version="1.3">"#,
        r#"<office:body><office:image>"#,
        r#"<draw:frame draw:name="flat-image" svg:width="1cm" svg:height="1cm">"#,
        r#"<draw:image><office:binary-data>iVBORw0KGgo=</office:binary-data></draw:image>"#,
        r#"</draw:frame></office:image></office:body></office:document>"#,
    );
    let flat = FlatImageDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
    assert_eq!(flat.image().images().len(), 1);
    assert_eq!(flat.as_bytes(), xml.as_bytes());
}

#[test]
fn family_wrappers_reject_wrong_flat_families() {
    assert!(FlatTextDocument::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).is_err());
    assert!(FlatSpreadsheet::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).is_err());
    assert!(FlatPresentation::from_bytes(FLAT_TEXT.as_bytes().to_vec()).is_err());
    assert!(FlatDrawingDocument::from_bytes(FLAT_TEXT.as_bytes().to_vec()).is_err());
    assert!(FlatChartDocument::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).is_err());
    assert!(FlatImageDocument::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).is_err());
}

#[test]
fn family_wrappers_reject_packaged_and_garbage_input() {
    assert!(FlatTextDocument::from_bytes(b"not xml at all".to_vec()).is_err());
    assert!(FlatSpreadsheet::from_bytes(Vec::new()).is_err());
    // A packaged ODF (ZIP) must not open through the flat readers.
    let zip_prefix = b"PK\x03\x04mimetype".to_vec();
    assert!(FlatPresentation::from_bytes(zip_prefix).is_err());
}
