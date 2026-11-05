#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odm::{Package, subdocument::Target};

const MIME: &str = "application/vnd.oasis.opendocument.text-master";
const CONTENT: &str = include_str!("fixtures/libreoffice-master-document-content.xml");

fn package(content: &str, meta_xml: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(source_meta_xml) = meta_xml {
        writer
            .add_file("meta.xml", source_meta_xml.as_bytes())
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn reference_projection_is_ordered_section_bound_and_inert() {
    let master = Package::from_bytes(&package(CONTENT, None)).unwrap();
    let master_document = master.document();
    let references = master_document.subdocuments();
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
fn metadata_title_is_projected_from_the_source_snapshot() {
    let meta = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>Master title</dc:title></office:meta></office:document-meta>"#;
    let master = Package::from_bytes(&package(CONTENT, Some(meta))).unwrap();
    assert_eq!(master.document().title(), Some("Master title"));
}

#[test]
fn real_master_fixture_title_transaction_is_reversible_and_preserves_opaque_parts() {
    let meta = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><dc:title>Before</dc:title><meta:user-defined meta:name="opaque">keep</meta:user-defined></office:meta></office:document-meta>"#;
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("custom/cache.bin", b"cached", "application/octet-stream")
        .unwrap();
    writer.add_file("meta.xml", meta.as_bytes()).unwrap();
    let source = Package::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap();

    let mut edit = source.edit_title().unwrap();
    edit.set("After & retained").unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.package();
    assert_eq!(changed.document().title(), Some("After & retained"));
    assert_eq!(changed.document().content_xml(), CONTENT);
    assert!(
        changed
            .document()
            .files()
            .unwrap()
            .contains(&"custom/cache.bin".to_string())
    );
    let changed_archive = OwnedPackage::from_bytes(changed.as_bytes().to_vec()).unwrap();
    assert_eq!(
        changed_archive.get_file("custom/cache.bin").unwrap(),
        b"cached"
    );
    let changed_meta = String::from_utf8(changed_archive.get_file("meta.xml").unwrap()).unwrap();
    assert!(
        changed_meta.contains(r#"<meta:user-defined meta:name="opaque">keep</meta:user-defined>"#)
    );
    assert!(!changed_meta.contains('\n'));

    let reapplied = source.apply_title(commit.patch()).unwrap().into_package();
    assert_eq!(reapplied.as_bytes(), changed.as_bytes());
    let restored = changed
        .apply_title(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn title_noop_reuses_the_exact_source_artifact() {
    let meta = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>Still</dc:title></office:meta></office:document-meta>"#;
    let source = Package::from_bytes(&package(CONTENT, Some(meta))).unwrap();
    let commit = source.edit_title().unwrap().commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), source.as_bytes());
}

#[test]
fn title_edit_borrows_the_projected_title_until_first_mutation() {
    let meta = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>Borrowed</dc:title></office:meta></office:document-meta>"#;
    let source = Package::from_bytes(&package(CONTENT, Some(meta))).unwrap();
    let document = source.document();
    let source_title = document.title().unwrap();
    let mut edit = source.edit_title().unwrap();

    // The projected title is already retained by the immutable package.  A
    // title edit borrows that bounded value and only allocates after set().
    assert_eq!(edit.title().unwrap().as_ptr(), source_title.as_ptr());
    edit.set("Changed").unwrap();
    assert_ne!(edit.title().unwrap().as_ptr(), source_title.as_ptr());
}

#[test]
fn named_entities_are_rejected_before_semantic_projection() {
    let content = CONTENT.replacen("Master introduction", "&untrusted;", 1);
    assert!(Package::from_bytes(&package(&content, None)).is_err());
}
