use litchi_odf::{
    MutableSpreadsheet, OdfImage, OdfImageFrame, OdfImagePart, OdfImageSource, Spreadsheet,
    SpreadsheetBuilder,
};
use std::io::{Cursor, Write};

fn linked_image(href: &str) -> OdfImage {
    OdfImage {
        part: OdfImagePart::Content,
        source: OdfImageSource::Linked {
            href: href.to_string(),
        },
        frame: Some(OdfImageFrame {
            name: Some("Revenue & forecast".to_string()),
            title: Some("Quarterly chart".to_string()),
            description: Some("External preview; never fetched".to_string()),
            x: Some("2cm".to_string()),
            y: Some("-0.5cm".to_string()),
            width: Some("5cm".to_string()),
            height: Some("30mm".to_string()),
            ..OdfImageFrame::default()
        }),
        xml_id: Some("image-1".to_string()),
        filter_name: None,
        declared_media_type: Some("image/png".to_string()),
        link_type: Some("simple".to_string()),
        show: Some("embed".to_string()),
        actuate: Some("onLoad".to_string()),
        alternative_index: 0,
    }
}

#[test]
fn builder_and_mutable_round_trip_inert_sheet_images() {
    let mut builder = SpreadsheetBuilder::new();
    builder
        .add_sheet("Forecast")
        .unwrap()
        .add_row_with_values(&["Q1", "10"])
        .unwrap()
        .add_sheet_image(linked_image("https://example.invalid/chart.png?no-fetch"))
        .unwrap();
    let bytes = builder.build().unwrap();

    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let image = {
        let sheets = spreadsheet.sheets().unwrap();
        let sheet = &sheets[0];
        assert_eq!(sheet.images().len(), 1);
        sheet.images()[0].clone()
    };
    assert!(image.frame.as_ref().unwrap().sheet_shape);
    assert_eq!(
        image.frame.as_ref().unwrap().description.as_deref(),
        Some("External preview; never fetched")
    );
    assert!(matches!(
        &image.source,
        OdfImageSource::Linked { href } if href == "https://example.invalid/chart.png?no-fetch"
    ));
    assert!(spreadsheet.image_bytes(&image).unwrap().is_none());

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let removed = mutable.remove_sheet_image(0, 0).unwrap();
    mutable.add_sheet_image(0, removed).unwrap();
    let mut reparsed = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.sheets().unwrap()[0].images().len(), 1);
}

#[test]
fn parses_libreoffice_table_shapes_reference_and_rejects_invalid_forms() {
    let source = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../3rdparty/libreoffice-core/sc/qa/unit/data/draw-image-link.fods",
        ),
    )
    .unwrap();
    let package = package_with_content(&source);
    let mut parsed = Spreadsheet::from_bytes(package).unwrap();
    let image = parsed.sheets().unwrap()[0].images()[0].clone();
    assert_eq!(image.frame.as_ref().unwrap().name.as_deref(), Some("img1"));
    assert_eq!(image.frame.as_ref().unwrap().width.as_deref(), Some("5cm"));
    assert!(matches!(&image.source, OdfImageSource::Linked { href } if href.contains("192.0.2.1")));

    for body in [
        r#"<table:shapes/><table:shapes/>"#,
        r#"<evil:shapes xmlns:evil="urn:evil"><draw:frame/></evil:shapes>"#,
        r#"<table:shapes><draw:frame svg:width="5cm" svg:height="5cm"><draw:image xlink:href="a"/><draw:image xlink:href="b"/></draw:frame></table:shapes>"#,
    ] {
        let content = content_with_shapes(body);
        let result = Spreadsheet::from_bytes(package_with_content(&content))
            .and_then(|mut spreadsheet| spreadsheet.sheets().map(|_| ()));
        assert!(result.is_err(), "accepted {body}");
    }

    let mut builder = SpreadsheetBuilder::new();
    let mut invalid = linked_image("relative.png");
    invalid.frame.as_mut().unwrap().width = Some("-1cm".to_string());
    assert!(builder.add_sheet_image(invalid).is_err());
}

fn content_with_shapes(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row/>{body}</table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

fn package_with_content(content: &str) -> Vec<u8> {
    let mimetype = "application/vnd.oasis.opendocument.spreadsheet";
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(
        format!(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
        )
        .as_bytes(),
    )
    .unwrap();
    zip.finish().unwrap();
    output.into_inner()
}
