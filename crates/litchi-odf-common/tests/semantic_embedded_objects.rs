use litchi_odf_common::{
    drawing::Part,
    embedded::{Kind, Object, Root, Source, scan_flat, scan_package},
    package::PackageLookup,
};

const ODT_REMOTE: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:p><draw:frame><draw:object xlink:href="http://192.0.2.1:12345/object"/></draw:frame></text:p></office:text></office:body></office:document>"#;
const ODS_REMOTE: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:frame><draw:object xlink:href="http://192.0.2.1:12345/object"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document>"#;
const ODP_REMOTE: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Slide1"><draw:frame><draw:object xlink:href="http://192.0.2.1:12345/object"/></draw:frame></draw:page></office:presentation></office:body></office:document>"#;
const ODT_MATH: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:text><text:p><draw:frame><draw:object><math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math></draw:object></draw:frame></text:p></office:text></office:body></office:document>"#;
const ODP_MATH: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page draw:name="Slide1"><draw:frame><draw:object><math xmlns="http://www.w3.org/1998/Math/MathML"><mi>y</mi></math></draw:object></draw:frame></draw:page></office:presentation></office:body></office:document>"#;
const ODT_CHART: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:text><text:p><draw:frame><draw:object draw:notify-on-update-of-ranges="Table1"><office:document><office:body><office:chart/></office:body></office:document></draw:object></draw:frame></text:p></office:text></office:body></office:document>"#;

struct Lookup<'a> {
    entries: &'a [(&'a str, Option<&'a str>)],
}

impl PackageLookup for Lookup<'_> {
    fn has_file(&self, path: &str) -> bool {
        self.entries.iter().any(|(entry, _)| *entry == path)
    }

    fn media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .iter()
            .find_map(|(entry, media_type)| (*entry == path).then_some(*media_type).flatten())
    }
}

#[test]
fn libreoffice_remote_objects_remain_typed_and_inert_across_families() {
    for xml in [ODT_REMOTE, ODS_REMOTE, ODP_REMOTE] {
        let objects = scan_flat(xml).unwrap();
        assert_eq!(objects.len(), 1);
        assert_eq!(objects[0].part, Part::FlatDocument);
        assert_eq!(objects[0].kind, Kind::Object);
        assert!(matches!(
            &objects[0].source,
            Source::Linked { href } if href.starts_with("http://192.0.2.1:12345/")
        ));
    }
}

#[test]
fn libreoffice_inline_math_and_chart_payloads_are_retained_without_recursion() {
    for xml in [ODT_MATH, ODP_MATH] {
        let objects = scan_flat(xml).unwrap();
        assert_eq!(objects.len(), 1);
        assert!(matches!(
            &objects[0].source,
            Source::InlineXml { root: Root::MathMl, xml, .. }
                if xml.contains("<math") && xml.contains("</math>")
        ));
    }

    let objects = scan_flat(ODT_CHART).unwrap();
    let chart = objects
        .iter()
        .find(|object| object.notify_on_update_of_ranges.as_deref() == Some("Table1"))
        .unwrap();
    assert!(matches!(
        &chart.source,
        Source::InlineXml { root: Root::OpenDocument, xml, .. }
            if xml.contains("office:document") && xml.contains("office:chart")
    ));
}

#[test]
fn package_subdocuments_are_exactly_classified() {
    let lookup = Lookup {
        entries: &[
            ("Object_1/content.xml", Some("text/xml")),
            (
                "Object_1/",
                Some("application/vnd.oasis.opendocument.chart"),
            ),
        ],
    };
    for body in [
        "<office:text><text:p>{object}</text:p></office:text>",
        "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:shapes>{object}</table:shapes></table:table></office:spreadsheet>",
        "<office:presentation><draw:page draw:name=\"Slide1\">{object}</draw:page></office:presentation>",
    ] {
        let object = "<draw:frame><draw:object xlink:href=\"./Object_1\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>";
        let xml = content_document(&body.replace("{object}", object));
        let objects = scan_package(&xml, None, &lookup).unwrap();
        assert_eq!(objects.len(), 1);
        assert_subdocument(&objects[0]);
    }
}

#[test]
fn package_file_and_inline_ole_sources_are_inert_and_exact() {
    let lookup = Lookup {
        entries: &[("Object1", Some("application/vnd.sun.star.oleobject"))],
    };
    let body = "<office:text><text:p><draw:frame><draw:object-ole draw:class-id=\"clsid:deadbeef\" xlink:href=\"Object1\"/></draw:frame></text:p></office:text>";
    let objects = scan_package(&content_document(body), None, &lookup).unwrap();
    assert_eq!(objects[0].kind, Kind::ObjectOle);
    assert_eq!(objects[0].class_id.as_deref(), Some("clsid:deadbeef"));
    assert!(matches!(
        &objects[0].source,
        Source::PackageFile {
            path,
            manifest_media_type: Some(media_type),
            ..
        } if path == "Object1" && media_type == "application/vnd.sun.star.oleobject"
    ));

    let body = "<office:text><text:p><draw:frame draw:name=\"OLE\"><draw:object-ole xlink:href=\"https://example.invalid/ignored\"><office:binary-data>Q0RG</office:binary-data></draw:object-ole><svg:title>OLE title</svg:title><svg:desc>OLE description</svg:desc></draw:frame></text:p></office:text>";
    let objects = scan_flat(&flat_document(body)).unwrap();
    assert!(matches!(
        &objects[0].source,
        Source::InlineBinary { bytes, ignored_href: Some(href) }
            if bytes == b"CDF" && href == "https://example.invalid/ignored"
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
        assert!(scan_flat(&flat_document(&body)).is_err(), "{object}");
    }

    let dtd = flat_document("<office:text><text:p><draw:object/></text:p></office:text>").replacen(
        "<office:body>",
        "<!DOCTYPE x><office:body>",
        1,
    );
    assert!(scan_flat(&dtd).is_err());

    let lookup = Lookup { entries: &[] };
    let body =
        "<office:text><text:p><draw:object xlink:href=\"%2e%2e/secret\"/></text:p></office:text>";
    assert!(scan_package(&content_document(body), None, &lookup).is_err());
}

#[test]
fn depth_count_and_inline_xml_byte_limits_are_enforced() {
    let oversized = "x".repeat(16 * 1024 * 1024 + 1);
    let body = format!(
        "<office:text><text:p><draw:object><math:math>{oversized}</math:math></draw:object></text:p></office:text>"
    );
    assert!(scan_flat(&flat_document(&body)).is_err());

    let nested = format!(
        "{}{}",
        "<text:span>".repeat(4_097),
        "</text:span>".repeat(4_097)
    );
    let body = format!("<office:text><text:p>{nested}</text:p></office:text>");
    assert!(scan_flat(&flat_document(&body)).is_err());

    let objects = "<draw:object/>".repeat(100_001);
    let body = format!("<office:text><text:p>{objects}</text:p></office:text>");
    assert!(scan_flat(&flat_document(&body)).is_err());
}

fn assert_subdocument(object: &Object) {
    assert!(matches!(
        &object.source,
        Source::PackageSubdocument {
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

fn content_document(body: &str) -> String {
    format!(
        "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\"><office:body>{body}</office:body></office:document-content>"
    )
}

fn flat_document(body: &str) -> String {
    format!(
        "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body>{body}</office:body></office:document>"
    )
}
