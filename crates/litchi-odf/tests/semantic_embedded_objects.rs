use litchi_odf::{
    Document, FlatOpenDocument, OdfEmbeddedObjectKind, OdfEmbeddedObjectSource,
    OdfInlineObjectRoot, OpenDocumentFamily, OpenDocumentPackage, Presentation, Spreadsheet,
};
use std::io::{Cursor, Write};
use zip::CompressionMethod;
use zip::write::SimpleFileOptions;

const ODT_REMOTE: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/odfimport/data/draw-object-link.fodt"
);
const ODS_REMOTE: &str =
    include_str!("../../../test-data/libreoffice-core/sc/qa/unit/data/draw-object-link.fods");
const ODP_REMOTE: &str =
    include_str!("../../../test-data/libreoffice-core/sd/qa/unit/data/draw-object-link.fodp");
const ODT_MATH: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/text-with-formula.fodt"
);
const ODP_MATH: &str =
    include_str!("../../../test-data/libreoffice-core/sd/qa/unit/data/odp/Math.fodp");
const ODT_CHART: &str =
    include_str!("../../../test-data/libreoffice-core/sw/qa/core/doc/data/tdf171549.fodt");

#[test]
fn libreoffice_remote_objects_remain_typed_and_inert_across_families() {
    for (xml, family) in [
        (ODT_REMOTE, OpenDocumentFamily::Text),
        (ODS_REMOTE, OpenDocumentFamily::Spreadsheet),
        (ODP_REMOTE, OpenDocumentFamily::Presentation),
    ] {
        let document = FlatOpenDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
        assert_eq!(document.family(), family);
        let objects = document.embedded_objects().unwrap();
        assert_eq!(objects.len(), 1);
        assert_eq!(objects[0].kind, OdfEmbeddedObjectKind::Object);
        assert!(matches!(
            &objects[0].source,
            OdfEmbeddedObjectSource::Linked { href }
                if href.starts_with("http://192.0.2.1:12345/")
        ));
    }
}

#[test]
fn libreoffice_inline_math_and_chart_payloads_are_retained_without_recursion() {
    for xml in [ODT_MATH, ODP_MATH] {
        let document = FlatOpenDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
        let objects = document.embedded_objects().unwrap();
        assert_eq!(objects.len(), 1);
        assert!(matches!(
            &objects[0].source,
            OdfEmbeddedObjectSource::InlineXml {
                root: OdfInlineObjectRoot::MathMl,
                xml,
                ..
            } if xml.contains("<math") && xml.contains("</math>")
        ));
    }

    let document = FlatOpenDocument::from_bytes(ODT_CHART.as_bytes().to_vec()).unwrap();
    let objects = document.embedded_objects().unwrap();
    let chart = objects
        .iter()
        .find(|object| object.notify_on_update_of_ranges.as_deref() == Some("Table1"))
        .unwrap();
    assert!(matches!(
        &chart.source,
        OdfEmbeddedObjectSource::InlineXml {
            root: OdfInlineObjectRoot::OpenDocument,
            xml,
            ..
        } if xml.contains("office:document") && xml.contains("office:chart")
    ));
}

#[test]
fn package_subdocuments_are_exactly_classified_for_specialized_families() {
    for (mimetype, family) in [
        (
            "application/vnd.oasis.opendocument.text",
            OpenDocumentFamily::Text,
        ),
        (
            "application/vnd.oasis.opendocument.spreadsheet",
            OpenDocumentFamily::Spreadsheet,
        ),
        (
            "application/vnd.oasis.opendocument.presentation",
            OpenDocumentFamily::Presentation,
        ),
    ] {
        let bytes = package(
            mimetype,
            "<draw:object xlink:href=\"./Object_1\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>",
            &[(
                "Object_1/content.xml",
                b"<office:document-content/>" as &[u8],
                "application/vnd.oasis.opendocument.chart",
            )],
            Some(("Object_1/", "application/vnd.oasis.opendocument.chart")),
        );
        let generic = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
        assert_eq!(generic.family(), family);
        assert_subdocument(&generic.embedded_objects().unwrap()[0]);

        let specialized = if family == OpenDocumentFamily::Text {
            Document::from_bytes(bytes)
                .unwrap()
                .embedded_objects()
                .unwrap()
        } else if family == OpenDocumentFamily::Spreadsheet {
            Spreadsheet::from_bytes(bytes)
                .unwrap()
                .embedded_objects()
                .unwrap()
        } else {
            Presentation::from_bytes(bytes)
                .unwrap()
                .embedded_objects()
                .unwrap()
        };
        assert_subdocument(&specialized[0]);
    }
}

#[test]
fn package_file_and_inline_ole_sources_are_inert_and_exact() {
    let bytes = package(
        "application/vnd.oasis.opendocument.text",
        "<draw:object-ole draw:class-id=\"clsid:deadbeef\" xlink:href=\"Object1\"/>",
        &[(
            "Object1",
            b"opaque-ole" as &[u8],
            "application/vnd.sun.star.oleobject",
        )],
        None,
    );
    let objects = OpenDocumentPackage::from_bytes(bytes)
        .unwrap()
        .embedded_objects()
        .unwrap();
    assert_eq!(objects[0].kind, OdfEmbeddedObjectKind::ObjectOle);
    assert_eq!(objects[0].class_id.as_deref(), Some("clsid:deadbeef"));
    assert!(matches!(
        &objects[0].source,
        OdfEmbeddedObjectSource::PackageFile {
            path,
            manifest_media_type: Some(media_type),
            ..
        } if path == "Object1" && media_type == "application/vnd.sun.star.oleobject"
    ));

    let body = "<office:text><text:p><draw:frame draw:name=\"OLE\"><draw:object-ole xlink:href=\"https://example.invalid/ignored\"><office:binary-data>Q0RG</office:binary-data></draw:object-ole><s:title>OLE title</s:title><s:desc>OLE description</s:desc></draw:frame></text:p></office:text>";
    let document = FlatOpenDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", body).into_bytes(),
    )
    .unwrap();
    let objects = document.embedded_objects().unwrap();
    assert!(matches!(
        &objects[0].source,
        OdfEmbeddedObjectSource::InlineBinary {
            bytes,
            ignored_href: Some(href),
        } if bytes == b"CDF" && href == "https://example.invalid/ignored"
    ));
    let frame = objects[0].frame.as_ref().unwrap();
    assert_eq!(frame.title.as_deref(), Some("OLE title"));
    assert_eq!(frame.description.as_deref(), Some("OLE description"));
}

#[test]
fn malformed_active_content_and_unsafe_package_paths_are_rejected() {
    for object in [
        "<draw:object-ole><office:binary-data>not!base64</office:binary-data></draw:object-ole>",
        "<draw:object-ole><office:binary-data><text:p/></office:binary-data></draw:object-ole>",
        "<draw:object-ole><office:binary-data/><office:binary-data/></draw:object-ole>",
        "<draw:object><math:math/><office:document/></draw:object>",
        "<draw:object><draw:object/></draw:object>",
    ] {
        let body = format!("<office:text><text:p>{object}</text:p></office:text>");
        let document = FlatOpenDocument::from_bytes(
            flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
        )
        .unwrap();
        assert!(document.embedded_objects().is_err());
    }

    let xml = flat_document(
        "application/vnd.oasis.opendocument.text",
        "<office:text><text:p><draw:object/></text:p></office:text>",
    )
    .replacen("<office:body>", "<!DOCTYPE x><office:body>", 1);
    let parsed = FlatOpenDocument::from_bytes(xml.into_bytes());
    assert!(
        parsed
            .and_then(|document| document.embedded_objects())
            .is_err()
    );

    let bytes = package(
        "application/vnd.oasis.opendocument.text",
        "<draw:object xlink:href=\"%2e%2e/secret\"/>",
        &[],
        None,
    );
    assert!(
        OpenDocumentPackage::from_bytes(bytes)
            .unwrap()
            .embedded_objects()
            .is_err()
    );
}

#[test]
fn depth_count_and_inline_xml_byte_limits_are_enforced() {
    let oversized = "x".repeat(16 * 1024 * 1024 + 1);
    let body = format!(
        "<office:text><text:p><draw:object><math:math>{oversized}</math:math></draw:object></text:p></office:text>"
    );
    let document = FlatOpenDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
    )
    .unwrap();
    assert!(document.embedded_objects().is_err());

    let nested = format!(
        "{}{}",
        "<text:span>".repeat(4_097),
        "</text:span>".repeat(4_097)
    );
    let body = format!("<office:text><text:p>{nested}</text:p></office:text>");
    let document = FlatOpenDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
    )
    .unwrap();
    assert!(document.embedded_objects().is_err());

    let objects = "<draw:object/>".repeat(100_001);
    let body = format!("<office:text><text:p>{objects}</text:p></office:text>");
    let document = FlatOpenDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
    )
    .unwrap();
    assert!(document.embedded_objects().is_err());
}

fn assert_subdocument(object: &litchi_odf::OdfEmbeddedObject) {
    assert!(matches!(
        &object.source,
        OdfEmbeddedObjectSource::PackageSubdocument {
            href,
            root_path,
            content_path,
            manifest_media_type: Some(media_type),
        } if href == "./Object_1"
            && root_path == "Object_1/"
            && content_path == "Object_1/content.xml"
            && media_type == "application/vnd.oasis.opendocument.chart"
    ));
}

fn flat_document(mimetype: &str, body: &str) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" office:mimetype=\"{mimetype}\"><office:body>{body}</office:body></office:document>"
    )
}

fn package(
    mimetype: &str,
    object: &str,
    entries: &[(&str, &[u8], &str)],
    directory: Option<(&str, &str)>,
) -> Vec<u8> {
    let body = if mimetype.ends_with(".text") {
        format!("<office:text><text:p><draw:frame>{object}</draw:frame></text:p></office:text>")
    } else if mimetype.ends_with(".spreadsheet") {
        format!(
            "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:shapes><draw:frame>{object}</draw:frame></table:shapes></table:table></office:spreadsheet>"
        )
    } else {
        format!(
            "<office:presentation><draw:page draw:name=\"Slide1\"><draw:frame>{object}</draw:frame></draw:page></office:presentation>"
        )
    };
    let content = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"><office:body>{body}</office:body></office:document-content>"
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    for (path, bytes, _) in entries {
        zip.start_file(path, deflated).unwrap();
        zip.write_all(bytes).unwrap();
    }

    let mut manifest_entries = String::new();
    if let Some((path, media_type)) = directory {
        manifest_entries.push_str(&format!(
            "<manifest:file-entry manifest:full-path=\"{path}\" manifest:media-type=\"{media_type}\"/>"
        ));
    }
    for (path, _, media_type) in entries {
        manifest_entries.push_str(&format!(
            "<manifest:file-entry manifest:full-path=\"{path}\" manifest:media-type=\"{media_type}\"/>"
        ));
    }
    let manifest = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{mimetype}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>{manifest_entries}</manifest:manifest>"
    );
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}
