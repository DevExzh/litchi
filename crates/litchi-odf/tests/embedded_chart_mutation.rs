use litchi_odf::{
    ChartDefinition, ChartText, Document, OdfEmbeddedChartStorage, Presentation, Spreadsheet,
    constants,
};
use std::io::{Cursor, Write};
use zip::CompressionMethod;
use zip::write::SimpleFileOptions;

fn chart(title: &str) -> ChartDefinition {
    let mut chart = ChartDefinition::new("chart:bar");
    chart.title = Some(ChartText::new(title));
    chart
}

#[test]
fn odt_packaged_chart_roundtrip_replace_remove_and_atomic_failure() {
    let mut document =
        Document::from_bytes(host_package(constants::ODF_TEXT, "text", None)).unwrap();
    let index = document.add_embedded_chart(&chart("First")).unwrap();
    assert_eq!(document.embedded_chart(index).unwrap().text(), "First");
    document
        .replace_embedded_chart(index, &chart("Second"))
        .unwrap();
    assert_eq!(document.embedded_chart(index).unwrap().text(), "Second");
    let before = document.to_bytes().unwrap();
    assert!(document.replace_embedded_chart(99, &chart("Nope")).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    document.remove_embedded_chart(index).unwrap();
    assert!(document.embedded_objects().unwrap().is_empty());
}

#[test]
fn ods_inline_chart_roundtrip_replace_and_remove() {
    let mut sheet = Spreadsheet::from_bytes(host_package(
        constants::ODF_SPREADSHEET,
        "spreadsheet",
        None,
    ))
    .unwrap();
    let index = sheet
        .add_embedded_chart_with_storage(
            "Sheet1",
            &chart("Inline"),
            OdfEmbeddedChartStorage::InlineXml,
        )
        .unwrap();
    assert_eq!(sheet.embedded_chart(index).unwrap().text(), "Inline");
    sheet
        .replace_embedded_chart(index, &chart("Changed"))
        .unwrap();
    assert_eq!(sheet.embedded_chart(index).unwrap().text(), "Changed");
    sheet.remove_embedded_chart(index).unwrap();
    assert!(sheet.embedded_objects().unwrap().is_empty());
}

#[test]
fn odp_packaged_chart_roundtrip_and_remove() {
    let mut presentation = Presentation::from_bytes(host_package(
        constants::ODF_PRESENTATION,
        "presentation",
        None,
    ))
    .unwrap();
    let index = presentation
        .add_embedded_chart("Slide1", &chart("Slide chart"))
        .unwrap();
    assert_eq!(
        presentation.embedded_chart(index).unwrap().text(),
        "Slide chart"
    );
    presentation.remove_embedded_chart(index).unwrap();
    assert!(presentation.embedded_objects().unwrap().is_empty());
}

#[test]
fn malformed_manifest_and_external_targets_fail_atomically() {
    let object = r#"<draw:object xlink:href="./Object_1"/>"#;
    let mut document = Document::from_bytes(host_package(
        constants::ODF_TEXT,
        "text",
        Some((object, "application/octet-stream")),
    ))
    .unwrap();
    let before = document.to_bytes().unwrap();
    assert!(
        document
            .replace_embedded_chart(0, &chart("Rejected"))
            .is_err()
    );
    assert_eq!(document.to_bytes().unwrap(), before);

    let external = r#"<draw:object xlink:href="https://example.invalid/chart.odc"/>"#;
    let mut linked = Document::from_bytes(host_package_with_object(
        constants::ODF_TEXT,
        "text",
        external,
        None,
    ))
    .unwrap();
    let before = linked.to_bytes().unwrap();
    assert!(linked.remove_embedded_chart(0).is_err());
    assert_eq!(linked.to_bytes().unwrap(), before);

    for attributes in [
        r#"office:mimetype="application/vnd.oasis.opendocument.text""#,
        r#"mimetype="application/vnd.oasis.opendocument.chart""#,
        r#"office:mimetype="application/vnd.oasis.opendocument.chart" office:mimetype="application/vnd.oasis.opendocument.chart""#,
    ] {
        let inline = format!(
            r#"<draw:object><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" {attributes}><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area/></chart:chart></office:chart></office:body></office:document></draw:object>"#
        );
        let mut invalid_inline = Document::from_bytes(host_package_with_object(
            constants::ODF_TEXT,
            "text",
            &inline,
            None,
        ))
        .unwrap();
        let before = invalid_inline.to_bytes().unwrap();
        assert!(
            invalid_inline
                .replace_embedded_chart(0, &chart("Rejected"))
                .is_err()
        );
        assert_eq!(invalid_inline.to_bytes().unwrap(), before);
    }
}

fn host_package(mimetype: &str, family: &str, object: Option<(&str, &str)>) -> Vec<u8> {
    match object {
        Some((object, media)) => host_package_with_object(mimetype, family, object, Some(media)),
        None => host_package_with_object(mimetype, family, "", None),
    }
}

fn host_package_with_object(
    mimetype: &str,
    family: &str,
    object: &str,
    object_media: Option<&str>,
) -> Vec<u8> {
    let body = match family {
        "text" => format!("<office:text><text:p>host{object}</text:p></office:text>"),
        "spreadsheet" => format!(
            "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:table-row/><table:shapes>{object}</table:shapes></table:table></office:spreadsheet>"
        ),
        "presentation" => format!(
            "<office:presentation><draw:page draw:name=\"Slide1\">{object}</draw:page></office:presentation>"
        ),
        _ => unreachable!(),
    };
    let content = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"><office:body>{body}</office:body></office:document-content>"
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    let mut extra_manifest = String::new();
    if let Some(media) = object_media {
        zip.start_file("Object_1/content.xml", deflated).unwrap();
        zip.write_all(
            litchi_odf::serialize_chart_content(&chart("Existing"))
                .unwrap()
                .as_bytes(),
        )
        .unwrap();
        extra_manifest = format!(
            "<manifest:file-entry manifest:full-path=\"Object_1/\" manifest:media-type=\"{media}\"/><manifest:file-entry manifest:full-path=\"Object_1/content.xml\" manifest:media-type=\"text/xml\"/>"
        );
    }
    let manifest = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{mimetype}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>{extra_manifest}</manifest:manifest>"
    );
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}
