//! Small package fixtures shared by ODT integration tests.

use litchi_core::xml::escape_xml;
use soapberry_zip::office::StreamingArchiveWriter;

#[allow(
    dead_code,
    reason = "each integration-test crate imports only the fixture helpers it exercises"
)]
pub(crate) fn compact_xml_fixture(xml: &str) -> String {
    xml.lines()
        .map(str::trim)
        .collect::<Vec<_>>()
        .join(" ")
        .replace("> <", "><")
}

pub(crate) fn package(mimetype: &str, files: &[(&str, &[u8])]) -> Vec<u8> {
    // Parser fixtures deliberately cover producer whitespace and other
    // noncanonical XML. Bypass publication validation only in this test helper.
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("mimetype", mimetype.as_bytes())
        .unwrap();
    for (path, bytes) in files {
        writer.write_stored(path, bytes).unwrap();
    }
    let mut manifest = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/>"#,
        escape_xml(mimetype)
    );
    for (path, _) in files {
        manifest.push_str(&format!(
            r#"<manifest:file-entry manifest:full-path="{}" manifest:media-type="text/xml"/>"#,
            escape_xml(path)
        ));
    }
    manifest.push_str("</manifest:manifest>");
    writer
        .write_stored("META-INF/manifest.xml", manifest.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}
