#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odp::{Presentation, edit};
use soapberry_zip::office::StreamingArchiveWriter;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Slide 1"><draw:frame draw:name="Photo"><draw:image xlink:href="Pictures/referenced.png"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

fn mark_zip_encrypted(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .unwrap();
    let read_u16 = |bytes: &[u8], offset| u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
    let read_u32 = |bytes: &[u8], offset| {
        u32::from_le_bytes([
            bytes[offset],
            bytes[offset + 1],
            bytes[offset + 2],
            bytes[offset + 3],
        ])
    };
    let count = read_u16(&bytes, eocd + 10) as usize;
    let mut cursor = read_u32(&bytes, eocd + 16) as usize;
    for _ in 0..count {
        assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_length = read_u16(&bytes, cursor + 28) as usize;
        let extra_length = read_u16(&bytes, cursor + 30) as usize;
        let comment_length = read_u16(&bytes, cursor + 32) as usize;
        let name_start = cursor + 46;
        let name = &bytes[name_start..name_start + name_length];
        let local_offset = read_u32(&bytes, cursor + 42) as usize;
        if name == wanted.as_bytes() {
            let central_flags = read_u16(&bytes, cursor + 8) | 1;
            bytes[cursor + 8..cursor + 10].copy_from_slice(&central_flags.to_le_bytes());
            let local_flags = read_u16(&bytes, local_offset + 6) | 1;
            bytes[local_offset + 6..local_offset + 8].copy_from_slice(&local_flags.to_le_bytes());
            return bytes;
        }
        cursor += 46 + name_length + extra_length + comment_length;
    }
    panic!("missing ZIP member {wanted}");
}

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
            .package()
            .unwrap()
            .manifest()
            .get_media_type("Vendor/opaque.bin"),
        Some("application/x-vendor")
    );
    assert!(
        !output_package
            .package()
            .unwrap()
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

#[test]
fn every_signature_owner_name_is_read_only() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file("META-INF/vendor-signatures.xml", br#"<signatures/>"#)
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        source.security_policy().unwrap(),
        edit::SecurityPolicy::SignedReadOnly
    );
    assert!(source.transaction().is_err());
}

#[test]
fn zip_encryption_is_classified_by_the_owned_package() {
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Vendor/opaque.bin" manifest:media-type="application/octet-stream"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive.write_deflated("content.xml", CONTENT).unwrap();
    archive
        .write_stored("Vendor/opaque.bin", b"opaque")
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    let bytes = mark_zip_encrypted(archive.finish_to_bytes().unwrap(), "Vendor/opaque.bin");
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert!(package.has_zip_encrypted_entries());
    let source = edit::Snapshot::from_bytes(bytes).unwrap();
    assert_eq!(
        source.security_policy().unwrap(),
        edit::SecurityPolicy::EncryptedReadOnly
    );
    assert!(source.transaction().is_err());
}

#[test]
fn replace_then_remove_across_calls_removes_the_source_member() {
    let source_bytes = source_bytes();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();

    assert_eq!(
        transaction
            .apply_media_changes(&[edit::MediaChange::replace(
                "Pictures/orphan.png",
                b"replacement",
                "image/png",
            )])
            .unwrap(),
        1
    );
    assert_eq!(
        transaction
            .apply_media_changes(&[edit::MediaChange::remove("Pictures/orphan.png")])
            .unwrap(),
        1
    );

    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert!(!package.has_file("Pictures/orphan.png").unwrap());
}

#[test]
fn restoring_an_exact_source_payload_across_calls_is_a_noop() {
    let source_bytes = source_bytes();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();

    assert_eq!(
        transaction
            .apply_media_changes(&[edit::MediaChange::remove("Pictures/orphan.png")])
            .unwrap(),
        1
    );
    assert_eq!(
        transaction
            .apply_media_changes(&[edit::MediaChange::replace(
                "Pictures/orphan.png",
                b"orphan-image",
                "image/png",
            )])
            .unwrap(),
        1
    );

    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}

#[test]
fn reference_index_decodes_entities_percent_escapes_and_relationship_targets() {
    let content = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Slide 1"><draw:frame draw:name="Photo"><draw:image xlink:href="Pictures/entity&amp;name.png"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content).unwrap();
    writer
        .add_file_with_media_type("Pictures/entity&name.png", b"entity-image", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/relative.png", b"relative-image", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/rels.png", b"rels-image", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type(
            "deck/owner.xml",
            br#"<owner href="../Pictures/relative.png"/>"#,
            "application/xml",
        )
        .unwrap();
    writer
        .add_file_with_media_type(
            "word/_rels/document.xml.rels",
            br#"<Relationships><Relationship Target="../Pictures/rels.png" TargetMode="Internal"/></Relationships>"#,
            "application/vnd.openxmlformats-package.relationships+xml",
        )
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    for path in [
        "Pictures/entity&name.png",
        "Pictures/relative.png",
        "Pictures/rels.png",
    ] {
        let mut transaction = source.transaction().unwrap();
        let error = transaction
            .apply_media_changes(&[edit::MediaChange::remove(path)])
            .unwrap_err();
        assert!(
            error.to_string().contains("referenced member"),
            "{path}: {error}"
        );
    }
}

#[test]
fn invalid_path_in_a_later_change_is_rejected_before_staging() {
    let source_bytes = source_bytes();
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let error = transaction
        .apply_media_changes(&[
            edit::MediaChange::add("Pictures/new.gif", b"new-gif", "image/gif"),
            edit::MediaChange::add("Pictures/../escape.gif", b"escape", "image/gif"),
        ])
        .unwrap_err();
    assert!(error.to_string().contains("path"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source_bytes);
}

#[test]
fn repeated_replacements_account_only_the_current_staged_payload() {
    let source = edit::Snapshot::from_bytes(source_bytes()).unwrap();
    let payload = vec![0xAB; 32 * 1024 * 1024];
    let mut transaction = source.transaction().unwrap();
    for index in 0..4 {
        let mut replacement = payload.clone();
        replacement[0] = index;
        assert_eq!(
            transaction
                .apply_media_changes(&[edit::MediaChange::replace(
                    "Pictures/orphan.png",
                    replacement,
                    "image/png",
                )])
                .unwrap(),
            1
        );
    }
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        package.get_file("Pictures/orphan.png").unwrap().len(),
        32 * 1024 * 1024
    );
}

#[test]
fn dtd_in_an_xml_owner_fails_media_removal_closed() {
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Vendor/links.xml" manifest:media-type="application/xml"/><manifest:file-entry manifest:full-path="Pictures/orphan.png" manifest:media-type="image/png"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive.write_deflated("content.xml", CONTENT).unwrap();
    archive
        .write_deflated(
            "Vendor/links.xml",
            br#"<!DOCTYPE links [<!ENTITY media "Pictures/orphan.png">]><links/>"#,
        )
        .unwrap();
    archive
        .write_stored("Pictures/orphan.png", b"orphan-image")
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    let source = edit::Snapshot::from_bytes(archive.finish_to_bytes().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let error = transaction
        .apply_media_changes(&[edit::MediaChange::remove("Pictures/orphan.png")])
        .unwrap_err();
    assert!(error.to_string().contains("DTD") || error.to_string().contains("entity"));
}
