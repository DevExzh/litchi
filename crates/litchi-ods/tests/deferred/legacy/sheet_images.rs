use litchi_ods::{
    CellValue, Image, ImageFrame, ImagePart, ImageSource, MutableSpreadsheet, OdfLength,
    OpenDocumentPackage, OwnedPackage, Spreadsheet, SpreadsheetBuilder,
};
use std::io::{Cursor, Write};

const PNG_PAYLOAD: &[u8] = b"\x89PNG\r\n\x1a\nfake-png-payload";
const GIF_PAYLOAD: &[u8] = b"GIF89afake-gif-payload";

fn linked_image(href: &str) -> Image {
    Image {
        part: ImagePart::Content,
        source: ImageSource::Linked {
            href: href.to_string(),
        },
        frame: Some(ImageFrame {
            name: Some("Revenue & forecast".to_string()),
            title: Some("Quarterly chart".to_string()),
            description: Some("External preview; never fetched".to_string()),
            x: Some("2cm".to_string()),
            y: Some("-0.5cm".to_string()),
            width: Some("5cm".to_string()),
            height: Some("30mm".to_string()),
            ..ImageFrame::default()
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
        ImageSource::Linked { href } if href == "https://example.invalid/chart.png?no-fetch"
    ));
    assert!(spreadsheet.image_bytes(&image).unwrap().is_none());

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let removed = mutable.remove_sheet_image(0, 0).unwrap();
    mutable.add_sheet_image(0, removed).unwrap();
    let mut reparsed = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.sheets().unwrap()[0].images().len(), 1);
}

#[test]
fn parses_libreoffice_table_shapes_references_alternatives_and_invalid_forms() {
    let source = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/sc/qa/unit/data/draw-image-link.fods"),
    )
    .unwrap();
    let package = package_with_content(&source);
    let mut parsed = Spreadsheet::from_bytes(package).unwrap();
    let image = parsed.sheets().unwrap()[0].images()[0].clone();
    assert_eq!(image.frame.as_ref().unwrap().name.as_deref(), Some("img1"));
    assert_eq!(image.frame.as_ref().unwrap().width.as_deref(), Some("5cm"));
    assert!(matches!(&image.source, ImageSource::Linked { href } if href.contains("192.0.2.1")));

    for body in [
        r#"<table:shapes/><table:shapes/>"#,
        r#"<evil:shapes xmlns:evil="urn:evil"><draw:frame/></evil:shapes>"#,
    ] {
        let content = content_with_shapes(body);
        let result = Spreadsheet::from_bytes(package_with_content(&content))
            .and_then(|mut spreadsheet| spreadsheet.sheets().map(|_| ()));
        assert!(result.is_err(), "accepted {body}");
    }

    let alternatives = content_with_shapes(
        r#"<table:shapes><draw:frame draw:name="preview" svg:width="5cm" svg:height="5cm"><draw:image xlink:href="Pictures/vector.svg"/><draw:image xlink:href="Pictures/fallback.png"/></draw:frame></table:shapes>"#,
    );
    let mut spreadsheet = Spreadsheet::from_bytes(package_with_content(&alternatives)).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    let images = sheets[0].images();
    assert_eq!(images.len(), 2);
    assert_eq!(
        (images[0].alternative_index, images[1].alternative_index),
        (0, 1)
    );
    assert_eq!(images[0].frame, images[1].frame);
    assert!(
        matches!(&images[1].source, ImageSource::Linked { href } if href == "Pictures/fallback.png")
    );
    let mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let mut reparsed = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let reparsed_sheets = reparsed.sheets().unwrap();
    let reparsed_images = reparsed_sheets[0].images();
    assert_eq!(reparsed_images.len(), 2);
    assert_eq!(reparsed_images[1].alternative_index, 1);

    let mut builder = SpreadsheetBuilder::new();
    let mut invalid = linked_image("relative.png");
    invalid.frame.as_mut().unwrap().width = Some("-1cm".to_string());
    assert!(builder.add_sheet_image(invalid).is_err());
}

#[test]
fn insert_image_round_trips_discoverable_package_pictures() {
    let mut mutable = MutableSpreadsheet::new();
    mutable.add_sheet("Data").unwrap();
    mutable.add_sheet("Live Data").unwrap();
    mutable
        .set_cell(0, 0, 0, CellValue::Text("keep".to_string()))
        .unwrap();

    let first = mutable
        .insert_image(
            0,
            0,
            0,
            PNG_PAYLOAD,
            &OdfLength::centimeters(5.0),
            &OdfLength::centimeters(3.0),
        )
        .unwrap();
    let second = mutable
        .insert_image(
            0,
            2,
            3,
            GIF_PAYLOAD,
            &OdfLength::millimeters(20.0),
            &OdfLength::millimeters(10.0),
        )
        .unwrap();
    let third = mutable
        .insert_image(
            1,
            1,
            1,
            PNG_PAYLOAD,
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
        )
        .unwrap();
    assert_eq!(first, "Pictures/image1.png");
    assert_eq!(second, "Pictures/image2.gif");
    assert_eq!(third, "Pictures/image3.png");

    let bytes = mutable.to_bytes().unwrap();

    // Payloads are stored verbatim with manifest media types.
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.get_file(&first).unwrap(), PNG_PAYLOAD);
    assert_eq!(package.get_file(&second).unwrap(), GIF_PAYLOAD);
    let manifest = package.package().unwrap();
    assert_eq!(
        manifest.manifest().get_media_type(&first),
        Some("image/png")
    );
    assert_eq!(
        manifest.manifest().get_media_type(&second),
        Some("image/gif")
    );

    // The media-level scan discovers the frames with sheet attribution.
    let generic = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
    let scanned = generic.images().unwrap();
    assert_eq!(scanned.len(), 3);
    let data_frames: Vec<_> = scanned
        .iter()
        .filter(|image| {
            image
                .frame
                .as_ref()
                .and_then(|frame| frame.sheet_name.as_deref())
                == Some("Data")
        })
        .collect();
    assert_eq!(data_frames.len(), 2);
    let frame = data_frames[0].frame.as_ref().unwrap();
    assert!(frame.sheet_shape);
    assert_eq!(frame.anchor_type.as_deref(), Some("cell"));
    assert_eq!(frame.width.as_deref(), Some("5cm"));
    assert_eq!(frame.height.as_deref(), Some("3cm"));
    assert_eq!(frame.end_cell_address.as_deref(), Some("Data.A1"));
    assert_eq!(
        data_frames[1]
            .frame
            .as_ref()
            .unwrap()
            .end_cell_address
            .as_deref(),
        Some("Data.D3")
    );
    let quoted = scanned
        .iter()
        .find(|image| {
            image
                .frame
                .as_ref()
                .and_then(|frame| frame.sheet_name.as_deref())
                == Some("Live Data")
        })
        .unwrap();
    assert_eq!(
        quoted.frame.as_ref().unwrap().end_cell_address.as_deref(),
        Some("'Live Data'.B2")
    );
    assert!(
        matches!(&data_frames[0].source, ImageSource::PackagePart { path, .. } if path == &first)
    );
    // Payloads resolve through the inert package-part accessor.
    assert_eq!(
        generic.image_bytes(data_frames[0]).unwrap().as_deref(),
        Some(PNG_PAYLOAD)
    );

    // The sheet model attributes images to their sheets; content is intact.
    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    assert_eq!(sheets[0].images().len(), 2);
    assert_eq!(sheets[1].images().len(), 1);
    assert_eq!(
        sheets[0].rows().unwrap()[0]
            .cell(0)
            .unwrap()
            .unwrap()
            .value()
            .unwrap(),
        &CellValue::Text("keep".to_string())
    );
}

#[test]
fn insert_image_continues_numbering_across_generations() {
    let mut mutable = MutableSpreadsheet::new();
    mutable.add_sheet("Data").unwrap();
    let first = mutable
        .insert_image(
            0,
            0,
            0,
            PNG_PAYLOAD,
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
        )
        .unwrap();
    let second = mutable
        .insert_image(
            0,
            0,
            1,
            GIF_PAYLOAD,
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
        )
        .unwrap();
    let first_generation = mutable.to_bytes().unwrap();

    let spreadsheet = Spreadsheet::from_bytes(first_generation).unwrap();
    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let third = mutable
        .insert_image(
            0,
            1,
            0,
            PNG_PAYLOAD,
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(2.0),
        )
        .unwrap();
    // Existing parts block both their own stems and sibling extensions.
    assert_eq!(third, "Pictures/image3.png");

    let bytes = mutable.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.get_file(&first).unwrap(), PNG_PAYLOAD);
    assert_eq!(package.get_file(&second).unwrap(), GIF_PAYLOAD);
    assert_eq!(package.get_file(&third).unwrap(), PNG_PAYLOAD);
    let mut reparsed = Spreadsheet::from_bytes(bytes).unwrap();
    assert_eq!(reparsed.sheets().unwrap()[0].images().len(), 3);
}

#[test]
fn insert_image_preserves_existing_fixture_parts() {
    let fixture =
        include_bytes!("../../../test-data/odfdo/tests/samples/simple_table_named_range.ods");
    let spreadsheet = Spreadsheet::from_bytes(fixture.to_vec()).unwrap();
    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let path = mutable
        .insert_image(
            0,
            0,
            1,
            GIF_PAYLOAD,
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(2.0),
        )
        .unwrap();
    assert_eq!(path, "Pictures/image1.gif");

    let bytes = mutable.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    // Auxiliary parts of the source package survive the edit.
    assert!(package.has_file("settings.xml").unwrap());
    assert!(package.has_file("Thumbnails/thumbnail.png").unwrap());
    assert_eq!(package.get_file(&path).unwrap(), GIF_PAYLOAD);

    let mut reparsed = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = reparsed.sheets().unwrap();
    assert_eq!(sheets[0].name().unwrap(), "Example1");
    assert_eq!(sheets[0].images().len(), 1);
    let frame = sheets[0].images()[0].frame.as_ref().unwrap();
    assert!(frame.sheet_shape);
    assert_eq!(frame.end_cell_address.as_deref(), Some("Example1.B1"));
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
        format!(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
        )
        .as_bytes(),
    )
    .unwrap();
    zip.finish().unwrap();
    output.into_inner()
}
