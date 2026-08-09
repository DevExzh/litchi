#![allow(
    clippy::unwrap_used,
    reason = "Fixed in-memory raw ZIP fixtures keep publication assertions concise."
)]

use litchi_odf_common::{
    constants,
    core::{
        AuthoredXmlFragment, OwnedPackage, XmlSourcePart, XmlSplicePublication,
        rebuild_package_with_xml_splices,
    },
};
use std::{collections::HashMap, io::Cursor, io::Write, ops::Range};

const CONTENT: &[u8] = b"<?xml version=\"1.0\"?>\n<document>\n  <leaf/>\n</document>";
const RDF: &[u8] = b"<rdf>\n <leaf/>\n</rdf>";
const DECLARED: &[u8] = b"<declared>\n <leaf/>\n</declared>";
const PLUS_XML: &[u8] = b"<vendor>\n <leaf/>\n</vendor>";
const SIGNATURE: &[u8] = b"<signatures>\n <leaf/>\n</signatures>";

fn raw_package(marker: &str) -> Vec<u8> {
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="graph.rdf" manifest:media-type=""/><manifest:file-entry manifest:full-path="declared-part" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="vendor-part" manifest:media-type="application/vnd.example+xml"/><manifest:file-entry manifest:full-path="META-INF/documentsignatures.xml" manifest:media-type="application/vnd.oasis.opendocument.digital-signature"/><manifest:file-entry manifest:full-path="marker.bin" manifest:media-type="application/octet-stream"/></manifest:manifest>"#,
        constants::ODF_DATABASE
    );
    let mut output = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut output);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let deflated = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(constants::ODF_DATABASE.as_bytes()).unwrap();
        for (path, bytes) in [
            ("META-INF/manifest.xml", manifest.as_bytes()),
            ("content.xml", CONTENT),
            ("graph.rdf", RDF),
            ("declared-part", DECLARED),
            ("vendor-part", PLUS_XML),
            ("META-INF/documentsignatures.xml", SIGNATURE),
            ("marker.bin", marker.as_bytes()),
        ] {
            zip.start_file(path, deflated).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    output.into_inner()
}

fn insertion(part: &XmlSourcePart) -> Range<usize> {
    let position = part
        .bytes()
        .windows(2)
        .rposition(|window| window == b"</")
        .unwrap();
    position..position
}

#[test]
fn every_raw_xml_part_class_requires_an_audited_fragment() {
    let package = OwnedPackage::from_bytes(raw_package("one")).unwrap();
    for path in [
        "content.xml",
        "graph.rdf",
        "declared-part",
        "vendor-part",
        "META-INF/documentsignatures.xml",
    ] {
        let part = XmlSourcePart::load(&package, path).unwrap();
        let proof = part.checked_range(insertion(&part), b"").unwrap();
        let mut publication = XmlSplicePublication::new(part);
        let fragment = AuthoredXmlFragment::markup(b"\n <authored/>".to_vec());
        assert!(fragment.is_err(), "noncompact fragment accepted for {path}");
        let unclassified = AuthoredXmlFragment::markup(b"plain text".to_vec());
        assert!(
            unclassified.is_err(),
            "unclassified fragment accepted for {path}"
        );
        publication
            .replace(
                proof,
                AuthoredXmlFragment::markup(b"<authored/>".to_vec()).unwrap(),
            )
            .unwrap();
    }
}

#[test]
fn source_identity_stale_ranges_and_overlaps_are_rejected() {
    let identical_bytes = raw_package("same-bytes");
    let first = OwnedPackage::from_bytes(identical_bytes.clone()).unwrap();
    let second = OwnedPackage::from_bytes(identical_bytes).unwrap();
    let first_part = XmlSourcePart::load(&first, "content.xml").unwrap();
    let second_part = XmlSourcePart::load(&second, "content.xml").unwrap();
    assert!(first_part.checked_range(0..5, b"stale").is_err());

    let foreign = first_part
        .checked_range(insertion(&first_part), b"")
        .unwrap();
    let mut publication = XmlSplicePublication::new(second_part.clone());
    assert!(
        publication
            .replace(foreign, AuthoredXmlFragment::deletion())
            .is_err()
    );

    let range = 32..39;
    let expected = &second_part.bytes()[range.clone()];
    publication
        .replace(
            second_part.checked_range(range.clone(), expected).unwrap(),
            AuthoredXmlFragment::markup(b"<new/>".to_vec()).unwrap(),
        )
        .unwrap();
    assert!(
        publication
            .replace(
                second_part
                    .checked_range((range.start + 1)..range.end, &expected[1..])
                    .unwrap(),
                AuthoredXmlFragment::deletion(),
            )
            .is_err()
    );

    let foreign_part = XmlSourcePart::load(&first, "graph.rdf").unwrap();
    let foreign_publication = XmlSplicePublication::new(foreign_part);
    assert!(
        rebuild_package_with_xml_splices(&second, vec![foreign_publication], 2 * 1024 * 1024)
            .is_err()
    );
}

#[test]
fn rebuild_preserves_source_bytes_outside_splices_and_enumerates_every_member() {
    let source = OwnedPackage::from_bytes(raw_package("opaque-marker")).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let insert = insertion(&part);
    let before = part.bytes()[..insert.start].to_vec();
    let after = part.bytes()[insert.end..].to_vec();
    let proof = part.checked_range(insert, b"").unwrap();
    let mut publication = XmlSplicePublication::new(part);
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<authored/>".to_vec()).unwrap(),
        )
        .unwrap();

    let rebuilt = OwnedPackage::from_bytes(
        rebuild_package_with_xml_splices(&source, vec![publication], 2 * 1024 * 1024).unwrap(),
    )
    .unwrap();
    let content = rebuilt.get_file("content.xml").unwrap();
    assert!(content.starts_with(&before));
    assert!(content.ends_with(&after));
    assert_eq!(rebuilt.get_file("graph.rdf").unwrap(), RDF);
    assert_eq!(rebuilt.get_file("declared-part").unwrap(), DECLARED);
    assert_eq!(rebuilt.get_file("vendor-part").unwrap(), PLUS_XML);
    assert_eq!(rebuilt.get_file("marker.bin").unwrap(), b"opaque-marker");
    assert!(!rebuilt.has_file("META-INF/documentsignatures.xml").unwrap());

    let files: HashMap<_, _> = rebuilt
        .files()
        .unwrap()
        .into_iter()
        .map(|path| (path.clone(), rebuilt.get_file(&path).unwrap()))
        .collect();
    assert_eq!(files.len(), 7);
    for path in [
        "mimetype",
        "META-INF/manifest.xml",
        "content.xml",
        "graph.rdf",
        "declared-part",
        "vendor-part",
        "marker.bin",
    ] {
        assert!(files.contains_key(path), "missing rebuilt member {path}");
    }
}

#[test]
fn explicit_fragment_classes_accept_compact_bytes_only() {
    assert!(AuthoredXmlFragment::start_tag(b"<node value=\"one\">".to_vec()).is_ok());
    assert!(AuthoredXmlFragment::text(b"one &amp; two".to_vec()).is_ok());
    let _deletion = AuthoredXmlFragment::deletion();
    assert!(AuthoredXmlFragment::start_tag(b"<node  value=\"one\">".to_vec()).is_err());
    assert!(AuthoredXmlFragment::text(b"   ".to_vec()).is_err());
}

#[test]
fn bounded_rebuild_refuses_to_materialize_an_oversized_archive() {
    let source = OwnedPackage::from_bytes(raw_package("bounded")).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let publication = XmlSplicePublication::new(part);
    assert!(rebuild_package_with_xml_splices(&source, vec![publication], 64).is_err());
}
