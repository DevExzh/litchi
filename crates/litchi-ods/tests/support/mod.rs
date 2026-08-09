//! Test-only raw ODF archive construction for parser and provenance fixtures.

use std::io::{Cursor, Write};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

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

pub(crate) fn raw_package(entries: &[(&str, &[u8], &str)]) -> Vec<u8> {
    let mut manifest = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>",
    );
    for (path, _, media_type) in entries {
        manifest.push_str("<manifest:file-entry manifest:full-path=\"");
        manifest.push_str(path);
        manifest.push_str("\" manifest:media-type=\"");
        manifest.push_str(media_type);
        manifest.push_str("\"/>");
    }
    manifest.push_str("</manifest:manifest>");

    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored).expect("raw mimetype");
    zip.write_all(MIMETYPE.as_bytes())
        .expect("raw mimetype bytes");
    zip.start_file("META-INF/manifest.xml", stored)
        .expect("raw manifest");
    zip.write_all(manifest.as_bytes())
        .expect("raw manifest bytes");
    for (path, bytes, _) in entries {
        zip.start_file(*path, stored).expect("raw entry");
        zip.write_all(bytes).expect("raw entry bytes");
    }
    zip.finish().expect("raw package finish");
    output.into_inner()
}
