#![allow(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::edit;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STYLES: &[u8] = br#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#;
const META: &[u8] = br#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#;
const OPAQUE: &[u8] = b"exact-opaque-payload";

fn content(pages: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:presentation="{PRESENTATION}" xmlns:xlink="{XLINK}" xmlns:mc="{MCE}" xmlns:vendor="urn:example:producer"><office:body><office:presentation>{pages}</office:presentation></office:body></office:document-content>"#
    )
}

fn package(content: &str) -> Vec<u8> {
    package_with_macro(content, false)
}

fn package_with_macro(content: &str, with_macro: bool) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES).unwrap();
    writer.add_file("meta.xml", META).unwrap();
    writer
        .add_file_with_media_type("Producer/opaque.bin", OPAQUE, "application/octet-stream")
        .unwrap();
    if with_macro {
        writer
            .add_file_with_media_type(
                "Basic/Standard/Module1.xml",
                br#"<?xml version="1.0"?><module/>"#,
                "text/xml",
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn package_with_auxiliary_xml(
    content: &str,
    path: &str,
    auxiliary_xml: &[u8],
    media_type: &str,
) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES).unwrap();
    writer.add_file("meta.xml", META).unwrap();
    writer
        .add_file_with_media_type(path, auxiliary_xml, media_type)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn content_xml(bytes: &[u8]) -> String {
    let package = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    String::from_utf8(package.get_file("content.xml").unwrap()).unwrap()
}

#[test]
fn dependency_free_blank_removal_is_exact_durable_reopenable_and_stale_checked() {
    let first = r#"<draw:page draw:name="First"/>"#;
    let selected = r#"<draw:page draw:name="Remove &amp; Me"/>"#;
    let last = r#"<draw:page draw:name="Last"/>"#;
    let source_bytes = package(&content(&format!("{first}{selected}{last}")));
    let source_package = OwnedPackage::from_bytes(source_bytes.clone()).unwrap();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();

    let mut transaction = source.transaction().unwrap();
    let removed = transaction
        .remove_dependency_free_blank_slide(1)
        .unwrap()
        .unwrap();
    assert_eq!(removed.index, 1);
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(source.bytes(), source_bytes);
    assert_eq!(commit.snapshot().slides().len(), 2);
    assert!(edit::Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).is_ok());

    let published_content = content_xml(commit.snapshot().bytes());
    assert!(published_content.contains(&format!("{first}{last}")));
    assert!(!published_content.contains(selected));
    let published = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    for path in ["styles.xml", "meta.xml", "Producer/opaque.bin"] {
        assert_eq!(
            published.get_file(path).unwrap(),
            source_package.get_file(path).unwrap(),
            "unrelated member payload changed: {path}"
        );
    }

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    let replayed = durable.apply(&source).unwrap();
    assert_eq!(replayed.bytes(), commit.snapshot().bytes());
    assert_eq!(
        durable.inverse().apply(&replayed).unwrap().bytes(),
        source_bytes
    );
    let stale = edit::Snapshot::from_bytes(package(&content(concat!(
        r#"<draw:page draw:name="Different"/>"#,
        r#"<draw:page draw:name="Last"/>"#,
    ))))
    .unwrap();
    assert!(durable.apply(&stale).is_err());
}

#[test]
fn missing_selector_is_an_exact_noop_and_final_slide_is_refused_atomically() {
    let source_bytes = package(&content(concat!(
        r#"<draw:page draw:name="First"/>"#,
        r#"<draw:page draw:name="Last"/>"#,
    )));
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .remove_dependency_free_blank_slide("Missing")
            .unwrap()
            .is_none()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);

    let single_bytes = package(&content(r#"<draw:page draw:name="Only"/>"#));
    let single = edit::Snapshot::from_bytes(single_bytes.clone()).unwrap();
    let mut transaction = single.transaction().unwrap();
    let error = transaction
        .remove_dependency_free_blank_slide(0)
        .unwrap_err();
    assert!(error.to_string().contains("final slide"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), single_bytes);
}

#[test]
fn selected_dependency_bearing_pages_are_refused_without_staging() {
    let cases = [
        (
            "children",
            r#"<draw:page draw:name="Unsafe"><draw:rect/></draw:page>"#,
        ),
        (
            "style",
            r#"<draw:page draw:name="Unsafe" draw:style-name="dp1"/>"#,
        ),
        (
            "master",
            r#"<draw:page draw:name="Unsafe" draw:master-page-name="Default"/>"#,
        ),
        (
            "layout",
            r#"<draw:page draw:name="Unsafe" presentation:presentation-page-layout-name="AL1T0"/>"#,
        ),
        (
            "navigation",
            r#"<draw:page draw:name="Unsafe" draw:nav-order="shape1"/>"#,
        ),
        (
            "identifier",
            r#"<draw:page draw:name="Unsafe" xml:id="slide-id"/>"#,
        ),
        (
            "link",
            r#"<draw:page draw:name="Unsafe" xlink:href="https://example.invalid/"/>"#,
        ),
        (
            "protection",
            r#"<draw:page draw:name="Unsafe" draw:protected="true"/>"#,
        ),
        (
            "MCE",
            r#"<draw:page draw:name="Unsafe" mc:Ignorable="future"/>"#,
        ),
        (
            "unknown",
            r#"<draw:page draw:name="Unsafe" vendor:keep="opaque"/>"#,
        ),
    ];
    for (description, page) in cases {
        let source_bytes = package(&content(&format!(r#"<draw:page draw:name="Keep"/>{page}"#)));
        let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
        let mut transaction = source.transaction().unwrap();
        transaction
            .remove_dependency_free_blank_slide(1)
            .expect_err(description);
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed(), "{description}");
        assert_eq!(commit.snapshot().bytes(), source_bytes, "{description}");
    }
}

#[test]
fn inbound_name_owner_and_macro_storage_are_refused_before_mutation() {
    let pages = concat!(
        r#"<vendor:owner vendor:value="Remove &amp; Me"/>"#,
        r#"<draw:page draw:name="Keep"/>"#,
        r#"<draw:page draw:name="Remove &amp; Me"/>"#,
    );
    for (description, bytes) in [
        ("inbound", package(&content(pages))),
        (
            "fragment link",
            package(&content(concat!(
                r##"<vendor:owner xlink:href="#Remove%20Me"/>"##,
                r#"<draw:page draw:name="Keep"/>"#,
                r#"<draw:page draw:name="Remove Me"/>"#,
            ))),
        ),
        (
            "auxiliary owner",
            package_with_auxiliary_xml(
                &content(concat!(
                    r#"<draw:page draw:name="Keep"/>"#,
                    r#"<draw:page draw:name="Remove Me"/>"#,
                )),
                "Producer/owner.XML",
                br#"<?xml version="1.0"?><owner value="Remove Me"/>"#,
                "application/octet-stream",
            ),
        ),
        (
            "RDF owner",
            package_with_auxiliary_xml(
                &content(concat!(
                    r#"<draw:page draw:name="Keep"/>"#,
                    r#"<draw:page draw:name="Remove Me"/>"#,
                )),
                "metadata.rdf",
                br#"<?xml version="1.0"?><owner value="Remove Me"/>"#,
                "application/rdf+xml",
            ),
        ),
        (
            "macro",
            package_with_macro(
                &content(concat!(
                    r#"<draw:page draw:name="Keep"/>"#,
                    r#"<draw:page draw:name="Remove"/>"#,
                )),
                true,
            ),
        ),
    ] {
        let source = edit::Snapshot::from_bytes(bytes.clone()).unwrap();
        let mut transaction = source.transaction().unwrap();
        let error = transaction
            .remove_dependency_free_blank_slide(1)
            .expect_err(description);
        assert!(
            error.to_string().contains("owner")
                || error.to_string().contains("macro")
                || error.to_string().contains("hyperlink"),
            "unexpected {description} error: {error}"
        );
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed(), "{description}");
        assert_eq!(commit.snapshot().bytes(), bytes, "{description}");
    }
}
