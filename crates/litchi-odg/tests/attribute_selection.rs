#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::Drawing;
use soapberry_zip::office::StreamingArchiveWriter;

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn raw_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

#[test]
fn aliased_attributes_select_by_namespace_and_preserve_unrelated_qnames() {
    let content = r#"<office:document-content xmlns="urn:example:default" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:x="urn:example:foreign" office:version="1.4"><office:body><office:drawing><d:page d:name="Page"><d:rect d:name="before" x:name="wrong" name="default" u:name="unbound" x:opaque="keep" s:x="1cm" s:y="2cm" s:width="3cm" s:height="4cm"/></d:page></office:drawing></office:body></office:document-content>"#;
    let source = Drawing::from_bytes(package(content)).unwrap();
    assert_eq!(source.pages()[0].shapes()[0].name(), Some("before"));

    let mut edit = source.edit();
    edit.set_shape_name(0, 0, "after").unwrap();
    let commit = edit.commit().unwrap();

    let expected = content.replace(r#"d:name="before""#, r#"d:name="after""#);
    assert_eq!(commit.snapshot().content_xml(), expected);
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].name(),
        Some("after")
    );
    for attribute in [
        r#"x:name="wrong""#,
        r#"name="default""#,
        r#"u:name="unbound""#,
        r#"x:opaque="keep""#,
    ] {
        assert!(
            commit.snapshot().content_xml().contains(attribute),
            "unrelated attribute was not preserved: {attribute}"
        );
    }
}

#[test]
fn aliased_duplicate_target_attributes_are_rejected() {
    let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:draw-alias="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><d:page d:name="Page"><d:rect d:name="first" draw-alias:name="second"/></d:page></office:drawing></office:body></office:document-content>"#;

    assert!(Drawing::from_bytes(package(content)).is_err());
}

#[test]
fn malformed_attribute_after_selected_fields_is_rejected() {
    let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><d:page d:name="Page" xml:id="page-id" d:style-name="page-style" d:master-page-name="Master" malformed/></office:drawing></office:body></office:document-content>"#;

    let valid = content.replace(" malformed/>", "/>");
    assert!(Drawing::from_bytes(raw_package(&valid)).is_ok());
    assert!(Drawing::from_bytes(raw_package(content)).is_err());
}

#[test]
fn duplicate_unrelated_attribute_after_selected_fields_is_rejected() {
    let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="urn:example:foreign" office:version="1.4"><office:body><office:drawing><d:page d:name="Page" xml:id="page-id" d:style-name="page-style" d:master-page-name="Master" x:opaque="first" x:opaque="second"/></office:drawing></office:body></office:document-content>"#;

    let valid = content.replace(
        r#" x:opaque="first" x:opaque="second""#,
        r#" x:opaque="first""#,
    );
    assert!(Drawing::from_bytes(raw_package(&valid)).is_ok());
    assert!(Drawing::from_bytes(raw_package(content)).is_err());
}
