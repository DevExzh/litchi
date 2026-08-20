#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odp::{Presentation, edit};

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Slide 1"><draw:frame draw:name="Photo"><draw:image xlink:href="Pictures/referenced.png"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

fn source_bytes() -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file(
            "styles.xml",
            br#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#,
        )
        .unwrap();
    writer
        .add_file(
            "meta.xml",
            br#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#,
        )
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/referenced.png", b"old-image", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/orphan.png", b"orphan-image", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type(
            "Vendor/opaque.bin",
            b"unknown-vendor-payload",
            "application/x-vendor",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn media_batch_replaces_adds_and_removes_with_untouched_member_preservation() {
    let source_bytes = source_bytes();
    let source_package = OwnedPackage::from_bytes(source_bytes.clone()).unwrap();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();

    let mut transaction = source.transaction().unwrap();
    assert_eq!(
        transaction
            .apply_media_changes(&[
                edit::MediaChange::replace("Pictures/referenced.png", b"new-image", "image/png"),
                edit::MediaChange::add("Pictures/new.gif", b"new-gif", "image/gif"),
                edit::MediaChange::remove("Pictures/orphan.png"),
            ])
            .unwrap(),
        3
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.patch().domains(), &[edit::Domain::Slides]);

    let output_package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        output_package.get_file("Pictures/referenced.png").unwrap(),
        b"new-image"
    );
    assert_eq!(
        output_package.get_file("Pictures/new.gif").unwrap(),
        b"new-gif"
    );
    assert!(!output_package.has_file("Pictures/orphan.png").unwrap());
    assert_eq!(
        output_package.get_file("Vendor/opaque.bin").unwrap(),
        source_package.get_file("Vendor/opaque.bin").unwrap()
    );
    assert_eq!(
        output_package.get_file("mimetype").unwrap(),
        source_package.get_file("mimetype").unwrap()
    );
    assert_eq!(
        output_package
            .manifest()
            .get_media_type("Vendor/opaque.bin"),
        Some("application/x-vendor")
    );
    assert!(
        !output_package
            .manifest()
            .entries
            .iter()
            .any(|(path, _)| path == "Pictures/orphan.png")
    );

    let reopened = Presentation::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        reopened
            .media_data(
                &litchi_odp::model::media::Reference::new("Pictures/referenced.png").unwrap(),
            )
            .unwrap(),
        Some(b"new-image".to_vec())
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
}

#[test]
fn referenced_removal_and_late_failure_are_failure_atomic() {
    let source_bytes = source_bytes();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let error = transaction
        .apply_media_changes(&[
            edit::MediaChange::replace("Pictures/referenced.png", b"would-change", "image/png"),
            edit::MediaChange::remove("Pictures/referenced.png"),
        ])
        .unwrap_err();
    assert!(error.to_string().contains("more than once"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);

    let mut transaction = source.transaction().unwrap();
    let error = transaction
        .apply_media_changes(&[edit::MediaChange::remove("Pictures/referenced.png")])
        .unwrap_err();
    assert!(error.to_string().contains("referenced member"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}

#[test]
fn exact_replacement_noop_and_batch_limits_preserve_source() {
    let source_bytes = source_bytes();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert_eq!(
        transaction
            .apply_media_changes(&[edit::MediaChange::replace(
                "Pictures/referenced.png",
                b"old-image",
                "image/png",
            )])
            .unwrap(),
        0
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);

    let too_many = (0..=edit::MAX_MEDIA_CHANGES)
        .map(|index| edit::MediaChange::remove(format!("Pictures/missing-{index}.png")))
        .collect::<Vec<_>>();
    let mut transaction = source.transaction().unwrap();
    assert!(transaction.apply_media_changes(&too_many).is_err());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}

#[test]
fn signed_sources_are_read_only_for_media_batches() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file(
            "META-INF/documentsignatures.xml",
            br#"<dsig:document-signatures xmlns:dsig="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
        )
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        source.security_policy().unwrap(),
        edit::SecurityPolicy::SignedReadOnly
    );
    assert!(source.transaction().is_err());
}
