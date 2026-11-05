//! `AuthoredXmlFragment::relocated` (record 0278): balanced source markup
//! moves within its part with its producer spelling, while unbalanced or
//! out-of-range relocations are refused.

#![allow(
    clippy::unwrap_used,
    reason = "Fixed in-memory raw ZIP fixtures keep publication assertions concise."
)]

use litchi_odf_common::constants;
use litchi_odf_common::core::{
    AuthoredXmlFragment, OwnedPackage, XmlSourcePart, XmlSplicePublication,
    rebuild_package_with_xml_splices,
};
use std::io::{Cursor, Write};

const CONTENT: &[u8] = b"<document>\n  <first a=\"1\" />\n  <second><x/></second>\n</document>";

fn package() -> OwnedPackage {
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        constants::ODF_DATABASE
    );
    let mut output = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut output);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(constants::ODF_DATABASE.as_bytes()).unwrap();
        for (path, bytes) in [
            ("META-INF/manifest.xml", manifest.as_bytes()),
            ("content.xml", CONTENT),
        ] {
            zip.start_file(path, stored).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    OwnedPackage::from_bytes(output.into_inner()).unwrap()
}

fn find(part: &XmlSourcePart, needle: &[u8]) -> std::ops::Range<usize> {
    let start = part
        .bytes()
        .windows(needle.len())
        .position(|window| window == needle)
        .unwrap();
    start..start + needle.len()
}

#[test]
fn relocated_markup_keeps_its_producer_spelling() {
    let source = package();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let first = find(&part, b"<first a=\"1\" />");
    let second = find(&part, b"<second><x/></second>");
    // The producer's space before `/>` would fail the authored audit.
    assert!(AuthoredXmlFragment::markup(part.bytes()[first.clone()].to_vec()).is_err());
    let mut publication = XmlSplicePublication::new(part.clone());
    publication
        .replace(
            part.checked_range(first.clone(), &part.bytes()[first.clone()])
                .unwrap(),
            AuthoredXmlFragment::relocated(&part, second.clone()).unwrap(),
        )
        .unwrap();
    publication
        .replace(
            part.checked_range(second.clone(), &part.bytes()[second.clone()])
                .unwrap(),
            AuthoredXmlFragment::relocated(&part, first).unwrap(),
        )
        .unwrap();
    let bytes = rebuild_package_with_xml_splices(&source, vec![publication], 1 << 20).unwrap();
    let published = OwnedPackage::from_bytes(bytes).unwrap();
    assert_eq!(
        published.get_file("content.xml").unwrap(),
        b"<document>\n  <second><x/></second>\n  <first a=\"1\" />\n</document>"
    );
}

#[test]
fn unbalanced_or_outside_ranges_are_refused() {
    let source = package();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let open = find(&part, b"<second>");
    assert!(AuthoredXmlFragment::relocated(&part, open).is_err());
    let text = find(&part, b"\n  <first");
    assert!(AuthoredXmlFragment::relocated(&part, text).is_err());
    let length = part.bytes().len();
    assert!(AuthoredXmlFragment::relocated(&part, length..length + 1).is_err());
}
