#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "Test assertions intentionally unwrap known-valid package fixture construction failures."
)]

use super::{Archive, read_manifest};
use super::{Entry, Manifest, is_media_path, parse_manifest};

#[test]
fn archive_reads_a_real_package_without_family_models() {
    let data = include_bytes!("../../../../test-data/odf/odt/shape-text-in-paragraph.odt");
    let archive = Archive::new(data).expect("ODF archive");
    assert!(archive.contains("mimetype"));
    assert!(archive.contains("content.xml"));
    assert!(archive.is_stored("mimetype").unwrap());
    let manifest = read_manifest(&archive).expect("neutral manifest");
    assert_eq!(manifest.get_media_type("content.xml"), Some("text/xml"));
}

#[test]
fn archive_rejects_invalid_zip_bytes() {
    assert!(Archive::new(b"not a zip").is_err());
}

#[test]
fn parses_neutral_entries_without_interpreting_encryption() {
    let xml = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
      <m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/>
      <m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="12">
        <m:encryption-data><m:algorithm m:algorithm-name="vendor-cipher"/></m:encryption-data>
      </m:file-entry>
    </m:manifest>"#;

    let manifest = parse_manifest(xml).expect("neutral manifest");
    assert_eq!(manifest.mimetype, "application/vnd.oasis.opendocument.text");
    assert_eq!(manifest.get_media_type("content.xml"), Some("text/xml"));
    assert_eq!(
        manifest
            .get_entry("content.xml")
            .and_then(|entry| entry.size),
        Some(12)
    );
    assert!(manifest.has_path("/"));
}

#[test]
fn rejects_duplicate_entries_and_invalid_sizes() {
    let duplicate = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="x"/><m:file-entry m:full-path="x"/></m:manifest>"#;
    assert!(parse_manifest(duplicate).is_err());

    let invalid_size = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="x" m:size="nope"/></m:manifest>"#;
    assert!(parse_manifest(invalid_size).is_err());
}

#[test]
fn media_path_classification_is_format_neutral() {
    assert!(is_media_path("Pictures/image.png"));
    assert!(is_media_path("Object 1/preview.svg"));
    assert!(is_media_path("thumbnail.jpg"));
    assert!(!is_media_path("content.xml"));
}

#[test]
fn model_remains_plain_and_cloneable() {
    let manifest = Manifest {
        mimetype: "text/xml".to_string(),
        entries: [(
            "content.xml".to_string(),
            Entry {
                media_type: "text/xml".to_string(),
                size: None,
            },
        )]
        .into_iter()
        .collect(),
    };
    assert_eq!(manifest.clone().paths().count(), 1);
}
