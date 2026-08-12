#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::edit;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const AUXILIARY: &[u8] = b"producer-owned-auxiliary-payload\0\xff";
const STYLES: &[u8] = br#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#;
const META: &[u8] = br#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#;

fn content(pages: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:presentation="{PRESENTATION}" xmlns:xlink="{XLINK}" xmlns:text="{TEXT}" xmlns:style="{STYLE}" xmlns:xml="{XML}" xmlns:mc="{MCE}" xmlns:vendor="urn:example:producer"><office:body><office:presentation>{pages}</office:presentation></office:body></office:document-content>"#
    )
}

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES).unwrap();
    writer.add_file("meta.xml", META).unwrap();
    writer
        .add_file_with_media_type("Producer/opaque.bin", AUXILIARY, "application/octet-stream")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn raw_package(content: &str) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored(
            "mimetype",
            b"application/vnd.oasis.opendocument.presentation",
        )
        .unwrap();
    writer
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    writer
        .write_deflated(
            "META-INF/manifest.xml",
            br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn package_content(bytes: &[u8]) -> String {
    let package = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    String::from_utf8(package.get_file("content.xml").unwrap()).unwrap()
}

#[test]
fn dependency_free_blank_copy_is_exact_named_reversible_and_source_checked() {
    let pages = concat!(
        r#"<draw:page draw:name="Blank"/>"#,
        r#"<draw:page draw:name="Blank Copy"/>"#,
        r#"<draw:page draw:name="Blank Copy 2"/>"#,
        r#"<draw:page draw:name="Untouched"/>"#,
    );
    let source_bytes = package(&content(pages));
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert_eq!(
        transaction.copy_dependency_free_blank_slide(0).unwrap(),
        Some(4)
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(source.bytes(), source_bytes);
    assert_eq!(commit.snapshot().slides().len(), 5);

    let expected_pages = format!("{pages}{}", r#"<draw:page draw:name="Blank Copy 3"/>"#);
    let published_content = package_content(commit.snapshot().bytes());
    assert!(published_content.contains(&expected_pages));
    let published = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        published.get_file("Producer/opaque.bin").unwrap(),
        AUXILIARY
    );
    assert_eq!(published.get_file("styles.xml").unwrap(), STYLES);
    assert_eq!(published.get_file("meta.xml").unwrap(), META);
    assert_eq!(
        commit
            .snapshot()
            .to_presentation()
            .unwrap()
            .pages()
            .unwrap()
            .pages()
            .iter()
            .map(|page| page.name.as_deref().unwrap())
            .collect::<Vec<_>>(),
        [
            "Blank",
            "Blank Copy",
            "Blank Copy 2",
            "Untouched",
            "Blank Copy 3"
        ]
    );

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    assert_eq!(
        durable.apply(&source).unwrap().bytes(),
        commit.snapshot().bytes()
    );
    assert_eq!(
        durable.inverse().apply(commit.snapshot()).unwrap().bytes(),
        source.bytes()
    );
    let stale =
        edit::Snapshot::from_bytes(package(&content(r#"<draw:page draw:name="Different"/>"#)))
            .unwrap();
    assert!(durable.apply(&stale).is_err());
}

#[test]
fn missing_copy_selector_is_an_exact_noop() {
    let source_bytes = package(&content(r#"<draw:page draw:name="Blank"/>"#));
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert_eq!(
        transaction
            .copy_dependency_free_blank_slide("Missing")
            .unwrap(),
        None
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}

#[test]
fn dependency_bearing_pages_are_refused_without_staging() {
    let cases = [
        ("nonempty", r#"<draw:page draw:name="Blank"></draw:page>"#),
        (
            "comment content",
            r#"<draw:page draw:name="Blank"><!--not blank--></draw:page>"#,
        ),
        (
            "cdata content",
            r#"<draw:page draw:name="Blank"><![CDATA[not blank]]></draw:page>"#,
        ),
        (
            "nested content",
            r#"<draw:page draw:name="Blank"><draw:rect/></draw:page>"#,
        ),
        (
            "style",
            r#"<draw:page draw:name="Blank" draw:style-name="dp1"/>"#,
        ),
        (
            "master",
            r#"<draw:page draw:name="Blank" draw:master-page-name="Default"/>"#,
        ),
        (
            "layout",
            r#"<draw:page draw:name="Blank" presentation:presentation-page-layout-name="AL1T0"/>"#,
        ),
        (
            "navigation",
            r#"<draw:page draw:name="Blank" draw:nav-order="shape1"/>"#,
        ),
        (
            "link",
            r#"<draw:page draw:name="Blank" xlink:href="https://example.invalid/"/>"#,
        ),
        (
            "event",
            r#"<draw:page draw:name="Blank"><presentation:event-listeners/></draw:page>"#,
        ),
        (
            "protection",
            r#"<draw:page draw:name="Blank" draw:protected="true"/>"#,
        ),
        (
            "MCE",
            r#"<draw:page draw:name="Blank" mc:Ignorable="future"/>"#,
        ),
        (
            "unknown attribute",
            r#"<draw:page draw:name="Blank" vendor:keep="opaque"/>"#,
        ),
        (
            "script child",
            r#"<draw:page draw:name="Blank"><office:script/></draw:page>"#,
        ),
    ];
    for (description, page) in cases {
        let source_bytes = package(&content(page));
        let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
        let mut transaction = source.transaction().unwrap();
        let error = transaction
            .copy_dependency_free_blank_slide(0)
            .expect_err(description);
        assert!(
            error.to_string().contains("dependency-free")
                || error.to_string().contains("cannot be reordered losslessly"),
            "unexpected {description} error: {error}"
        );
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed(), "{description}");
        assert_eq!(commit.snapshot().bytes(), source_bytes, "{description}");
    }
}

#[test]
fn noncompact_source_and_name_bound_are_refused_before_copy_allocation() {
    let noncompact =
        content("\n<draw:page draw:name=\"Blank\"/>\n<draw:page draw:name=\"Other\"/>\n");
    let source = edit::Snapshot::from_bytes(raw_package(&noncompact)).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .copy_dependency_free_blank_slide(0)
            .unwrap_err()
            .to_string()
            .contains("cannot be reordered losslessly")
    );

    let oversized_name = "x".repeat(4 * 1024 + 1);
    let source_bytes = package(&content(&format!(
        r#"<draw:page draw:name="{oversized_name}"/>"#
    )));
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .copy_dependency_free_blank_slide(0)
            .unwrap_err()
            .to_string()
            .contains("name exceeds")
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}
