use litchi_ods::{
    DdeConversionMode, DdeSource, MutableSpreadsheet, Spreadsheet, SpreadsheetBuilder,
};
use std::io::{Cursor, Write};

#[test]
fn parses_odfpy_sheet_dde_source_reference() {
    let content = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/odfpy-sheet-dde-source.fods"),
    )
    .unwrap();
    let mut spreadsheet = Spreadsheet::from_bytes(package_with_content(&content)).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    let source = sheets[0].dde_source().unwrap();
    assert_eq!(source.application, "soffice");
    assert_eq!(source.topic, "file:///never/contacted.ods");
    assert_eq!(source.item, "Sheet1.A1:B2");
    assert_eq!(source.name.as_deref(), Some("Reference Source"));
    assert_eq!(source.conversion_mode, Some(DdeConversionMode::KeepText));
    assert_eq!(source.automatic_update, Some(true));
}

#[test]
fn builder_and_mutable_round_trip_sheet_dde_source_inertly() {
    let mut source = DdeSource::new("calc&server", "file:///never/opened.ods", "Sheet 1.A1");
    source.name = Some("Prices <live>".to_string());
    source.conversion_mode = Some(DdeConversionMode::IntoEnglishNumber);
    source.automatic_update = Some(true);

    let mut builder = SpreadsheetBuilder::new();
    builder
        .add_sheet("Prices")
        .unwrap()
        .set_sheet_dde_source(Some(source.clone()))
        .unwrap()
        .add_row_with_values(&["cached", "only"])
        .unwrap();
    let mut parsed = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(parsed.sheets().unwrap()[0].dde_source(), Some(&source));

    let mut mutable = MutableSpreadsheet::from_spreadsheet(parsed).unwrap();
    mutable.set_sheet_dde_source(0, None).unwrap();
    mutable
        .set_sheet_dde_source(0, Some(source.clone()))
        .unwrap();
    let mut reparsed = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.sheets().unwrap()[0].dde_source(), Some(&source));
}

#[test]
fn rejects_malformed_spoofed_duplicate_and_oversized_sources() {
    for child in [
        r#"<office:dde-source office:dde-application="a" office:dde-topic="t"/>"#,
        r#"<office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i" office:conversion-mode="convert"/>"#,
        r#"<office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i" bad="x"/>"#,
        r#"<evil:dde-source xmlns:evil="urn:evil" office:dde-application="a" office:dde-topic="t" office:dde-item="i"/>"#,
        r#"<office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i"><table:table-row/></office:dde-source>"#,
        r#"<office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i">text</office:dde-source>"#,
        r#"<office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i"/><office:dde-source office:dde-application="b" office:dde-topic="u" office:dde-item="j"/>"#,
        r#"<table:scenario table:scenario-ranges=".A1:.B2" table:is-active="false"/><office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i"/>"#,
    ] {
        let content = content_with_child(child);
        let result = Spreadsheet::from_bytes(package_with_content(&content))
            .and_then(|mut spreadsheet| spreadsheet.sheets().map(|_| ()));
        assert!(result.is_err(), "accepted {child}");
    }

    let mut builder = SpreadsheetBuilder::new();
    let oversized = DdeSource::new("a".repeat(65_537), "topic", "item");
    assert!(builder.set_sheet_dde_source(Some(oversized)).is_err());
}

fn content_with_child(child: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="Sheet1">{child}<table:table-column/><table:table-row/></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

fn package_with_content(content: &str) -> Vec<u8> {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(
        format!(r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#).as_bytes(),
    )
    .unwrap();
    zip.finish().unwrap();
    output.into_inner()
}
