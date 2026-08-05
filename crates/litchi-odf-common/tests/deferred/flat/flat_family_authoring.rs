//! Semantic authoring on flat OpenDocument XML documents: edits through the
//! packaged mutable models are written back into flat XML form and survive a
//! reopen.

use litchi_odf::{
    CellValue, DrawingLayer, DrawingPageProperties, DrawingShapeKind, FlatChartDocument,
    FlatDrawingDocument, FlatDocument, FlatPresentation, FlatSpreadsheet, FlatTextDocument,
    Family, Shape,
};

const FLAT_TEXT: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/pageColumns.fodt"
);
const FLAT_SPREADSHEET: &str =
    include_str!("../../../test-data/libreoffice-core/sc/qa/unit/data/draw-image-link.fods");
const FLAT_SPREADSHEET_DDE: &str = include_str!("fixtures/odfpy-sheet-dde-source.fods");
const FLAT_PRESENTATION: &str = include_str!(
    "../../../test-data/libreoffice-core/sd/qa/unit/tiledrendering/data/slide-background-link.fodp"
);
const FLAT_DRAWING: &str = include_str!("../../../test-data/odf/drawing/fill-image-inline.fodg");

const FLAT_DRAWING_WITH_SHAPE: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"office:mimetype="application/vnd.oasis.opendocument.graphics" office:version="1.3">"#,
    r#"<office:body><office:drawing><draw:page draw:name="p1" draw:master-page-name="Default">"#,
    r#"<draw:layer-set><draw:layer draw:name="layout"/></draw:layer-set>"#,
    r#"<draw:rect draw:name="box" draw:layer="layout" svg:width="2cm" svg:height="1cm">"#,
    r#"<text:p>Old label</text:p></draw:rect>"#,
    r#"</draw:page></office:drawing></office:body></office:document>"#,
);

const FLAT_CHART: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3">"#,
    r#"<office:styles><style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="keep-me"/></office:styles>"#,
    r#"<office:body><office:chart>"#,
    r#"<chart:chart chart:class="chart:bar">"#,
    r#"<chart:title><text:p>Revenue</text:p></chart:title>"#,
    r#"<chart:plot-area>"#,
    r#"<chart:axis chart:dimension="x" chart:name="primary-x"/>"#,
    r#"<chart:series chart:values-cell-range-address="Sheet1.A1:A3" chart:class="chart:bar"/>"#,
    r#"</chart:plot-area></chart:chart></office:chart></office:body></office:document>"#,
);

#[test]
fn flat_text_paragraph_edit_round_trips() {
    let document = FlatTextDocument::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut mutable = document.into_mutable().unwrap();
    mutable
        .document_mut()
        .update_paragraph(0, "Rewritten Intro")
        .unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let output = String::from_utf8(bytes.clone()).unwrap();

    // The splice output is a valid flat document of the same family.
    let flat = FlatDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), Family::Text);
    // Sections the mutation did not touch keep their original content.
    assert!(output.contains("<office:settings>"));
    assert!(output.contains("<office:master-styles>"));
    assert!(output.contains("office:mimetype=\"application/vnd.oasis.opendocument.text\""));

    let reopened = FlatTextDocument::from_bytes(bytes).unwrap();
    let text = reopened.document().text().unwrap();
    assert!(text.contains("Rewritten Intro"));
    assert!(text.contains("Cras eu leo sed justo"));

    // The save path writes the same flat bytes to disk.
    let path = std::env::temp_dir().join("litchi-flat-authoring-roundtrip.fodt");
    mutable.save(&path).unwrap();
    let reopened_from_disk = FlatTextDocument::open(&path).unwrap();
    assert!(
        reopened_from_disk
            .document()
            .text()
            .unwrap()
            .contains("Rewritten Intro")
    );
    std::fs::remove_file(&path).unwrap();
}

#[test]
fn flat_text_mutation_preserves_settings_bytes() {
    // The settings section is not carried by the packaged mutation flows; the
    // splice must keep the original section verbatim.
    let settings_start = FLAT_TEXT.find("<office:settings>").unwrap();
    let settings_end = FLAT_TEXT.find("</office:settings>").unwrap() + "</office:settings>".len();
    let original_settings = &FLAT_TEXT[settings_start..settings_end];

    let document = FlatTextDocument::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut mutable = document.into_mutable().unwrap();
    mutable
        .document_mut()
        .update_paragraph(0, "Unrelated edit")
        .unwrap();
    let output = String::from_utf8(mutable.to_bytes().unwrap()).unwrap();
    assert!(output.contains(original_settings));
}

#[test]
fn flat_spreadsheet_cell_edit_round_trips() {
    let document = FlatSpreadsheet::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).unwrap();
    let mut mutable = document.into_mutable().unwrap();
    mutable
        .spreadsheet_mut()
        .set_cell(0, 0, 1, CellValue::Number(42.0))
        .unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let output = String::from_utf8(bytes.clone()).unwrap();
    let flat = FlatDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), Family::Spreadsheet);
    // The linked image stays inert stored metadata; it is never fetched.
    assert!(output.contains("tracking-pixel.png"));

    let mut reopened = FlatSpreadsheet::from_bytes(bytes).unwrap();
    let sheet = reopened
        .spreadsheet_mut()
        .sheet_by_index(0)
        .unwrap()
        .unwrap();
    let rows = sheet.rows().unwrap();
    assert_eq!(
        rows[0].cell(0).unwrap().unwrap().value().unwrap(),
        &CellValue::Text("test".to_string())
    );
    assert_eq!(
        rows[0].cell(1).unwrap().unwrap().value().unwrap(),
        &CellValue::Number(42.0)
    );
}

#[test]
fn flat_spreadsheet_odfpy_fixture_edit_round_trips() {
    // odfpy's flat DDE fixture has no office:meta section; the mutated package
    // always carries one, so the splice inserts it as a new section while the
    // inert DDE declaration survives untouched.
    let document = FlatSpreadsheet::from_bytes(FLAT_SPREADSHEET_DDE.as_bytes().to_vec()).unwrap();
    let mut mutable = document.to_mutable().unwrap();
    mutable
        .spreadsheet_mut()
        .set_cell(0, 0, 0, CellValue::Text("hello".to_string()))
        .unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let output = String::from_utf8(bytes.clone()).unwrap();
    assert!(output.contains("<office:meta>"));
    assert!(output.contains("never/contacted.ods"));

    // `to_mutable` leaves the original wrapper usable.
    assert_eq!(
        document.flat_document().family(),
        Family::Spreadsheet
    );

    let mut reopened = FlatSpreadsheet::from_bytes(bytes).unwrap();
    let sheet = reopened
        .spreadsheet_mut()
        .sheet_by_name("Live Data")
        .unwrap()
        .unwrap();
    assert_eq!(
        sheet.rows().unwrap()[0]
            .cell(0)
            .unwrap()
            .unwrap()
            .value()
            .unwrap(),
        &CellValue::Text("hello".to_string())
    );
}

#[test]
fn flat_presentation_slide_edit_round_trips() {
    let document = FlatPresentation::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).unwrap();
    let mut mutable = document.into_mutable().unwrap();
    mutable
        .presentation_mut()
        .update_slide(0, "New Title", "New Body")
        .unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let flat = FlatDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), Family::Presentation);

    let reopened = FlatPresentation::from_bytes(bytes).unwrap();
    let slides = reopened.presentation().slides().unwrap();
    assert_eq!(slides.len(), 1);
    assert_eq!(slides[0].text().unwrap(), "New Body");
}

#[test]
fn flat_unmodified_read_only_save_stays_byte_exact() {
    // The read-only wrappers keep their lossless guarantee; only the mutable
    // flow re-serializes.
    for bytes in [
        FLAT_TEXT.as_bytes(),
        FLAT_SPREADSHEET.as_bytes(),
        FLAT_PRESENTATION.as_bytes(),
    ] {
        let document = FlatTextDocument::from_bytes(bytes.to_vec());
        if let Ok(document) = document {
            assert_eq!(document.to_bytes(), bytes);
        }
    }
    let spreadsheet = FlatSpreadsheet::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).unwrap();
    assert_eq!(spreadsheet.to_bytes(), FLAT_SPREADSHEET.as_bytes());
    let presentation = FlatPresentation::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).unwrap();
    assert_eq!(presentation.to_bytes(), FLAT_PRESENTATION.as_bytes());
}

#[test]
fn flat_mutable_wrappers_reject_wrong_families() {
    assert!(FlatTextDocument::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).is_err());
    assert!(FlatSpreadsheet::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).is_err());
    assert!(FlatPresentation::from_bytes(FLAT_TEXT.as_bytes().to_vec()).is_err());
    assert!(FlatDrawingDocument::from_bytes(FLAT_CHART.as_bytes().to_vec()).is_err());
    assert!(FlatChartDocument::from_bytes(FLAT_DRAWING.as_bytes().to_vec()).is_err());
}

#[test]
fn flat_drawing_page_and_shape_authoring_round_trips() {
    // The fixture has an empty `<office:drawing/>` body and a fill-image
    // style section; authoring adds a page while the styles stay untouched.
    let document = FlatDrawingDocument::from_bytes(FLAT_DRAWING.as_bytes().to_vec()).unwrap();
    let mut mutable = document.into_mutable().unwrap();
    let mut properties = DrawingPageProperties::new();
    properties.set_name(Some("Page 1"));
    mutable.drawing_mut().add_page(properties).unwrap();
    mutable
        .drawing_mut()
        .add_layer(0, DrawingLayer::new("layout"))
        .unwrap();
    let shape = Shape {
        drawing_kind: Some(DrawingShapeKind::Rectangle),
        name: Some("added-rect".to_string()),
        text: "Added flat shape".to_string(),
        layer: Some("layout".to_string()),
        ..Shape::new()
    };
    mutable.drawing_mut().add_shape(0, shape).unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let output = String::from_utf8(bytes.clone()).unwrap();
    let flat = FlatDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), Family::Drawing);
    assert!(output.contains("libreoffice_5f_0"));

    let reopened = FlatDrawingDocument::from_bytes(bytes).unwrap();
    assert_eq!(reopened.drawing().page_count(), 1);
    let page = reopened.drawing().page(0).unwrap();
    assert_eq!(page.shapes().len(), 1);
    assert_eq!(page.shapes()[0].name.as_deref(), Some("added-rect"));
    assert_eq!(page.shapes()[0].text, "Added flat shape");
}

#[test]
fn flat_drawing_existing_shape_text_edit_round_trips() {
    let document =
        FlatDrawingDocument::from_bytes(FLAT_DRAWING_WITH_SHAPE.as_bytes().to_vec()).unwrap();
    let mut mutable = document.into_mutable().unwrap();
    let mut shape = mutable.drawing().page(0).unwrap().shapes()[0].clone();
    shape.text = "New label".to_string();
    mutable.drawing_mut().set_shape(0, 0, shape).unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let reopened = FlatDrawingDocument::from_bytes(bytes).unwrap();
    let page = reopened.drawing().page(0).unwrap();
    assert_eq!(page.shapes()[0].text, "New label");
    assert_eq!(page.layers().len(), 1);
}

#[test]
fn flat_chart_axis_update_round_trips() {
    let document = FlatChartDocument::from_bytes(FLAT_CHART.as_bytes().to_vec()).unwrap();
    let mut mutable = document.to_mutable().unwrap();
    let update = litchi_odf::ChartAxisUpdate {
        name: Some(Some("renamed-x".to_string())),
        ..Default::default()
    };
    mutable.chart_mut().update_axis(0, &update).unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let output = String::from_utf8(bytes.clone()).unwrap();
    let flat = FlatDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), Family::Chart);
    assert!(output.contains("keep-me"));

    let reopened = FlatChartDocument::from_bytes(bytes).unwrap();
    assert!(reopened.chart().find_axis("renamed-x").is_some());
    assert!(reopened.chart().find_axis("primary-x").is_none());
    assert!(reopened.chart().text().contains("Revenue"));

    // `to_mutable` left the original wrapper unchanged.
    assert!(document.chart().find_axis("primary-x").is_some());
}
