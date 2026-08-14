#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::edit;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const AUXILIARY: &[u8] = b"destination-owned-opaque\0\xff";

fn content(pages: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}"><office:body><office:presentation>{pages}</office:presentation></office:body></office:document-content>"#
    )
}

fn package(content_xml: &str, with_macro: bool) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file("content.xml", content_xml.as_bytes())
        .unwrap();
    writer
        .add_file(
            "styles.xml",
            br#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#,
        )
        .unwrap();
    writer
        .add_file(
            "meta.xml",
            br#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#,
        )
        .unwrap();
    writer
        .add_file_with_media_type("Producer/opaque.bin", AUXILIARY, "application/octet-stream")
        .unwrap();
    if with_macro {
        writer
            .add_file_with_media_type("Basic/Standard/Module1", b"macro", "text/plain")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn package_with_extra(content_xml: &str, path: &str, payload: &[u8], media_type: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file("content.xml", content_xml.as_bytes())
        .unwrap();
    writer
        .add_file(
            "styles.xml",
            br#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#,
        )
        .unwrap();
    writer
        .add_file(
            "meta.xml",
            br#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#,
        )
        .unwrap();
    writer
        .add_file_with_media_type(path, payload, media_type)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn content_bytes(bytes: &[u8]) -> Vec<u8> {
    OwnedPackage::from_bytes(bytes.to_vec())
        .unwrap()
        .get_file("content.xml")
        .unwrap()
}

#[test]
fn foreign_blank_transfer_is_source_bound_reversible_and_raw_preserving() {
    let donor_bytes = package(&content(r#"<draw:page draw:name="Donor"/>"#), false);
    let destination_bytes = package(
        &content(
            r#"<draw:page draw:name="Host"/><draw:page draw:name="Donor"/><draw:page draw:name="Donor Copy"/>"#,
        ),
        false,
    );
    let donor = edit::Snapshot::from_bytes(donor_bytes.clone()).unwrap();
    let destination = edit::Snapshot::from_bytes(destination_bytes.clone()).unwrap();
    let mut transaction = destination.transaction().unwrap();

    assert_eq!(
        transaction
            .transfer_dependency_free_blank_slide_from(&donor, 0)
            .unwrap(),
        Some(3)
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(donor.bytes(), donor_bytes);
    assert_eq!(destination.bytes(), destination_bytes);
    assert_eq!(commit.snapshot().slides().len(), 4);
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
        ["Host", "Donor", "Donor Copy", "Donor Copy 2"]
    );
    let published_content = content_bytes(commit.snapshot().bytes());
    assert!(
        String::from_utf8_lossy(&published_content)
            .contains(r#"<draw:page draw:name="Donor Copy 2"/>"#)
    );
    assert_eq!(
        OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec())
            .unwrap()
            .get_file("Producer/opaque.bin")
            .unwrap(),
        AUXILIARY
    );

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    assert_eq!(
        durable.apply(&destination).unwrap().bytes(),
        commit.snapshot().bytes()
    );
    assert_eq!(
        durable.inverse().apply(commit.snapshot()).unwrap().bytes(),
        destination.bytes()
    );
    let stale = edit::Snapshot::from_bytes(package(
        &content(r#"<draw:page draw:name="Different"/><draw:page draw:name="Host"/>"#),
        false,
    ))
    .unwrap();
    assert!(durable.apply(&stale).is_err());
    assert!(durable.inverse().apply(&stale).is_err());
}

#[test]
fn missing_foreign_selector_is_an_exact_noop() {
    let donor = edit::Snapshot::from_bytes(package(
        &content(r#"<draw:page draw:name="Donor"/>"#),
        false,
    ))
    .unwrap();
    let destination_bytes = package(&content(r#"<draw:page draw:name="Host"/>"#), false);
    let destination = edit::Snapshot::from_bytes(destination_bytes.clone()).unwrap();
    let mut transaction = destination.transaction().unwrap();
    assert_eq!(
        transaction
            .transfer_dependency_free_blank_slide_from(&donor, 99)
            .unwrap(),
        None
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().bytes(), destination_bytes);
}

#[test]
fn foreign_blank_transfer_supports_an_empty_retained_destination() {
    let donor = edit::Snapshot::from_bytes(package(
        &content(r#"<draw:page draw:name="Donor"/>"#),
        false,
    ))
    .unwrap();
    let destination = edit::Snapshot::from_bytes(package(&content(""), false)).unwrap();
    let mut transaction = destination.transaction().unwrap();
    assert_eq!(
        transaction
            .transfer_dependency_free_blank_slide_from(&donor, 0)
            .unwrap(),
        Some(0)
    );
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.snapshot().slides().len(), 1);
    assert!(
        String::from_utf8_lossy(&content_bytes(commit.snapshot().bytes()))
            .contains(r#"<draw:page draw:name="Donor"/>"#)
    );
}

#[test]
fn foreign_transfer_rejects_self_closing_presentation_destination() {
    let donor = edit::Snapshot::from_bytes(package(
        &content(r#"<draw:page draw:name="Donor"/>"#),
        false,
    ))
    .unwrap();
    let destination_content = format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}"><office:body><office:presentation/></office:body></office:document-content>"#
    );
    let destination_bytes = package(&destination_content, false);
    let destination = edit::Snapshot::from_bytes(destination_bytes.clone()).unwrap();
    let mut transaction = destination.transaction().unwrap();
    assert!(
        transaction
            .transfer_dependency_free_blank_slide_from(&donor, 0)
            .is_err()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), destination_bytes);
}

#[test]
fn foreign_transfer_rejects_nested_namespace_rebindings_at_all_insertions() {
    let donor = edit::Snapshot::from_bytes(package(
        &content(r#"<draw:page draw:name="Donor"/>"#),
        false,
    ))
    .unwrap();
    let cases = [
        format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}"><office:body xmlns:draw="urn:example:wrong"><office:presentation><draw:page draw:name="Host"/></office:presentation></office:body></office:document-content>"#
        ),
        format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:style="{STYLE}" xmlns:xlink="{XLINK}"><office:body><office:presentation xmlns:style="urn:example:wrong"><draw:page draw:name="Host"/></office:presentation></office:body></office:document-content>"#
        ),
        format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:style="{STYLE}" xmlns:xlink="{XLINK}"><office:automatic-styles xmlns:style="urn:example:wrong"/><office:body><office:presentation><draw:page draw:name="Host"/></office:presentation></office:body></office:document-content>"#
        ),
        format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}"><office:body xmlns="urn:example:wrong"><office:presentation><draw:page draw:name="Host"/></office:presentation></office:body></office:document-content>"#
        ),
    ];
    for destination_content in cases {
        let destination_bytes = package(&destination_content, false);
        let destination = edit::Snapshot::from_bytes(destination_bytes.clone()).unwrap();
        let mut transaction = destination.transaction().unwrap();
        assert!(
            transaction
                .transfer_dependency_free_blank_slide_from(&donor, 0)
                .is_err()
        );
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(commit.snapshot().bytes(), destination_bytes);
    }
}

#[test]
fn foreign_transfer_audits_case_insensitive_macro_and_all_xml_media_parts() {
    let donor_content = content(r#"<draw:page draw:name="Donor"/>"#);
    let cases = [
        (
            "bAsIc/Standard/Module",
            b"macro".as_slice(),
            "application/octet-stream",
        ),
        (
            "Scripts/Events.bin",
            b"macro".as_slice(),
            "application/octet-stream",
        ),
        (
            "settings.xml",
            br#"<settings><script:script xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"/></settings>"#.as_slice(),
            "text/xml",
        ),
        (
            "META-INF/custom.RDF",
            br##"<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description><script:script xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"/></rdf:Description></rdf:RDF>"##.as_slice(),
            "application/example+xml",
        ),
        (
            "META-INF/DocumentSignatures.xml",
            b"<signatures/>".as_slice(),
            "text/xml",
        ),
    ];
    for (path, payload, media_type) in cases {
        let donor_bytes = package_with_extra(&donor_content, path, payload, media_type);
        let donor = edit::Snapshot::from_bytes(donor_bytes).unwrap();
        let destination_bytes = package(&content(r#"<draw:page draw:name="Host"/>"#), false);
        let destination = edit::Snapshot::from_bytes(destination_bytes).unwrap();
        let mut transaction = destination.transaction().unwrap();
        assert!(
            transaction
                .transfer_dependency_free_blank_slide_from(&donor, 0)
                .is_err(),
            "unsafe donor part {path}"
        );
    }
}

#[test]
fn foreign_transfer_rejects_dependencies_macros_and_oversized_names_atomically() {
    let cases = [
        r#"<draw:page draw:name="Donor"><draw:rect/></draw:page>"#,
        r#"<draw:page draw:name="Donor" draw:master-page-name="Default"/>"#,
        r#"<draw:page draw:name="Donor" draw:protected="true"/>"#,
        r#"<draw:page draw:name="Donor" xlink:href="https://example.invalid/"/>"#,
    ];
    for page in cases {
        let donor = edit::Snapshot::from_bytes(package(&content(page), false)).unwrap();
        let destination_bytes = package(&content(r#"<draw:page draw:name="Host"/>"#), false);
        let destination = edit::Snapshot::from_bytes(destination_bytes.clone()).unwrap();
        let mut transaction = destination.transaction().unwrap();
        assert!(
            transaction
                .transfer_dependency_free_blank_slide_from(&donor, 0)
                .is_err()
        );
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(commit.snapshot().bytes(), destination_bytes);
    }

    let donor =
        edit::Snapshot::from_bytes(package(&content(r#"<draw:page draw:name="Donor"/>"#), true))
            .unwrap();
    let destination =
        edit::Snapshot::from_bytes(package(&content(r#"<draw:page draw:name="Host"/>"#), false))
            .unwrap();
    let mut transaction = destination.transaction().unwrap();
    assert!(
        transaction
            .transfer_dependency_free_blank_slide_from(&donor, 0)
            .is_err()
    );
    assert!(!transaction.commit().unwrap().changed());

    let oversized = "x".repeat(4 * 1024 + 1);
    let donor = edit::Snapshot::from_bytes(package(
        &content(&format!(r#"<draw:page draw:name="{oversized}"/>"#)),
        false,
    ))
    .unwrap();
    let destination_bytes = package(&content(r#"<draw:page draw:name="Host"/>"#), false);
    let destination = edit::Snapshot::from_bytes(destination_bytes.clone()).unwrap();
    let mut transaction = destination.transaction().unwrap();
    assert!(
        transaction
            .transfer_dependency_free_blank_slide_from(&donor, 0)
            .is_err()
    );
    assert_eq!(
        transaction.commit().unwrap().snapshot().bytes(),
        destination_bytes
    );
}
