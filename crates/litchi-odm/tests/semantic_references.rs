#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odm::{Master, subdocument::Target};

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const CONTENT: &str = include_str!("fixtures/libreoffice-master-document-content.xml");

fn package(content: &str, meta: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(meta) = meta {
        writer.add_file("meta.xml", meta.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn reference_projection_is_ordered_section_bound_and_inert() {
    let master = Master::from_bytes(package(CONTENT, None)).unwrap();
    let references = master.subdocuments();
    assert_eq!(references.len(), 3);
    assert_eq!(references[0].section(), "Chapter A");
    assert_eq!(references[0].href(), "Chapters/a.odt");
    assert!(matches!(references[0].target(), Target::Package(_)));
    assert_eq!(references[1].section(), "Nested C");
    assert!(matches!(references[1].target(), Target::Package(_)));
    assert_eq!(references[2].section(), "Chapter B");
    assert!(references[2].target().is_external());
    assert_eq!(references[2].href(), "https://example.test/b.odt");
}

#[test]
fn metadata_title_remains_projected_without_enabling_edits() {
    let meta = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>Master title</dc:title></office:meta></office:document-meta>"#;
    let master = Master::from_bytes(package(CONTENT, Some(meta))).unwrap();
    assert_eq!(master.title(), Some("Master title"));
}

#[test]
fn named_entities_are_rejected_before_semantic_projection() {
    let content = CONTENT.replacen("Master introduction", "&untrusted;", 1);
    assert!(Master::from_bytes(package(&content, None)).is_err());
}
