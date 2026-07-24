//! Semantic authoring on flat OpenDocument XML documents: edits through the
//! packaged mutable models are written back into flat XML form and survive a
//! reopen.

use litchi_odf::{
    CellValue, FlatOpenDocument, FlatPresentation, FlatSpreadsheet, FlatTextDocument,
    OpenDocumentFamily,
};

const FLAT_TEXT: &str =
    include_str!("../../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/pageColumns.fodt");
const FLAT_SPREADSHEET: &str =
    include_str!("../../../test-data/libreoffice-core/sc/qa/unit/data/draw-image-link.fods");
const FLAT_SPREADSHEET_DDE: &str = include_str!("fixtures/odfpy-sheet-dde-source.fods");
const FLAT_PRESENTATION: &str =
    include_str!("../../../test-data/libreoffice-core/sd/qa/unit/tiledrendering/data/slide-background-link.fodp");

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
    let flat = FlatOpenDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), OpenDocumentFamily::Text);
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
    let flat = FlatOpenDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), OpenDocumentFamily::Spreadsheet);
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
    assert_eq!(document.flat_document().family(), OpenDocumentFamily::Spreadsheet);

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
    let flat = FlatOpenDocument::from_bytes(bytes.clone()).unwrap();
    assert_eq!(flat.family(), OpenDocumentFamily::Presentation);

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
    let presentation =
        FlatPresentation::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).unwrap();
    assert_eq!(presentation.to_bytes(), FLAT_PRESENTATION.as_bytes());
}

#[test]
fn flat_mutable_wrappers_reject_wrong_families() {
    assert!(FlatTextDocument::from_bytes(FLAT_SPREADSHEET.as_bytes().to_vec()).is_err());
    assert!(FlatSpreadsheet::from_bytes(FLAT_PRESENTATION.as_bytes().to_vec()).is_err());
    assert!(FlatPresentation::from_bytes(FLAT_TEXT.as_bytes().to_vec()).is_err());
}
