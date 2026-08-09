#![allow(
    clippy::shadow_reuse,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::{Builder, Presentation, edit};
use soapberry_zip::office::StreamingArchiveWriter;

#[test]
fn transaction_publishes_compact_xml_without_mutating_source() {
    let mut builder = Builder::new();
    builder.add_slide_with_title("One", "Body").unwrap();
    let source = edit::Snapshot::from_bytes(builder.build().unwrap()).unwrap();
    let source_bytes = source.bytes().to_vec();

    let mut transaction = source.transaction().unwrap();
    transaction.add("Two", "Second body").unwrap();
    let commit = transaction.commit().unwrap();

    assert!(commit.changed());
    assert_eq!(source.bytes(), source_bytes);
    assert_eq!(commit.snapshot().slides().len(), 2);
    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    let restored = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(restored.bytes(), source.bytes());
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    for path in ["content.xml", "styles.xml", "meta.xml"] {
        let xml = String::from_utf8(package.get_file(path).unwrap()).unwrap();
        assert!(!xml.contains('\n'), "{path} was not compact");
        assert!(
            !xml.contains("> <"),
            "{path} contains inter-element spacing"
        );
    }
}

#[test]
fn exact_noop_and_inverse_patch_preserve_source_bytes() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let commit = source.transaction().unwrap().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().bytes(), source.bytes());

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.bytes(), source.bytes());
}

#[test]
fn patch_refuses_a_different_source_snapshot() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.add("Added", "Body").unwrap();
    let commit = transaction.commit().unwrap();

    let mut other = Builder::new();
    other.add_slide_with_title("Other", "Deck").unwrap();
    let other = edit::Snapshot::from_bytes(other.build().unwrap()).unwrap();
    assert!(commit.patch().apply(&other).is_err());
}

#[test]
fn retained_unknown_slide_is_preserved_or_rewrite_is_refused() {
    const CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:vendor="urn:example:opaque"><office:body><office:presentation><draw:page draw:name="source"><vendor:opaque vendor:type="vendor-gear"/></draw:page></office:presentation></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    let mut refused = source.transaction().unwrap();
    let before = refused.slides().to_vec();
    assert!(refused.replace(0, "Changed", "Body").is_err());
    assert_eq!(refused.slides(), before);

    let mut transaction = source.transaction().unwrap();
    transaction.add("New", "Body").unwrap();
    let commit = transaction.commit().unwrap();
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains(r#"vendor:type="vendor-gear""#));
}

#[test]
fn presentation_exposes_the_source_checked_entry_point() {
    let presentation = Presentation::from_bytes(Builder::new().build().unwrap()).unwrap();
    let snapshot = presentation.snapshot().unwrap();
    assert!(snapshot.slides().is_empty());
}

#[test]
fn snapshot_clone_and_exact_noop_commit_share_backing_bytes() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<edit::Snapshot>();

    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let clone = source.clone();
    assert_eq!(clone.bytes().as_ptr(), source.bytes().as_ptr());

    let commit = source.transaction().unwrap().commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes().as_ptr(), source.bytes().as_ptr());
    assert_eq!(
        commit.patch().apply(&source).unwrap().bytes().as_ptr(),
        source.bytes().as_ptr()
    );
}

#[test]
fn package_input_limit_rejects_n_plus_one_bytes() {
    const MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;

    assert!(edit::Snapshot::from_bytes(vec![0; MAX_PACKAGE_BYTES + 1]).is_err());
}

#[test]
fn package_output_limit_rejects_n_plus_one_before_publication() {
    const MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let source_pointer = source.bytes().as_ptr();
    let path = "Media/n-plus-one.bin";
    let media_type = "application/octet-stream";
    let payload_len = MAX_PACKAGE_BYTES - source.bytes().len() + 1 - path.len() - media_type.len();
    let payload = vec![0; payload_len];
    let mut transaction = source.transaction().unwrap();

    assert!(transaction.embed_media(path, &payload, media_type).is_err());

    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(source.bytes().as_ptr(), source_pointer);
    assert_eq!(commit.snapshot().bytes().as_ptr(), source_pointer);
    assert_eq!(commit.snapshot().bytes(), source.bytes());
}

#[test]
fn changed_commit_snapshot_and_patch_output_share_backing_bytes() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.add("Added", "Body").unwrap();
    let commit = transaction.commit().unwrap();
    let published_pointer = commit.snapshot().bytes().as_ptr();

    assert!(commit.changed());
    assert_eq!(
        commit.patch().apply(&source).unwrap().bytes().as_ptr(),
        published_pointer
    );
    assert_eq!(commit.into_snapshot().bytes().as_ptr(), published_pointer);
}

#[test]
fn changed_publication_refuses_noncompact_referenced_xml() {
    const MIME: &str = "application/vnd.oasis.opendocument.presentation";
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#;
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive.write_deflated("content.xml", CONTENT).unwrap();
    archive
        .write_deflated("Object 1/content.xml", b"<object>\n  <opaque/>\n</object>")
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    let source = edit::Snapshot::from_bytes(archive.finish_to_bytes().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.add("New", "Body").unwrap();
    assert!(transaction.commit().is_err());
    assert!(source.slides().is_empty());
}

#[test]
fn changed_publication_accepts_semantic_xml_whitespace() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer
        .add_file(
            "content.xml",
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
        )
        .unwrap();
    writer
        .add_file_with_media_type(
            "Object 1/content.xml",
            b"<object xml:space=\"preserve\" note=\"line&#10;two\"><!--\nsemantic comment\n--><![CDATA[\nsemantic data\n]]>\nsemantic text\n</object>",
            "text/xml",
        )
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.add("New", "Body").unwrap();
    assert!(transaction.commit().is_ok());
}

#[test]
fn text_limit_failure_is_atomic_at_n_plus_one() {
    const N: usize = 16 * 1024 * 1024;
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let boundary = "a".repeat(N);
    transaction.add("", &boundary).unwrap();
    let before = transaction.slides().len();
    let over = "b".repeat(N + 1);
    assert!(transaction.add("", &over).is_err());
    assert_eq!(transaction.slides().len(), before);
}

#[test]
fn identical_transactions_emit_deterministic_bytes() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut first = source.transaction().unwrap();
    first.add("Same", "Content").unwrap();
    let first = first.commit().unwrap();
    let mut second = source.transaction().unwrap();
    second.add("Same", "Content").unwrap();
    let second = second.commit().unwrap();
    assert_eq!(first.snapshot().bytes(), second.snapshot().bytes());
}

#[test]
fn incomplete_source_page_coverage_is_refused_without_panicking() {
    let content = br#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:body><office:presentation><draw:page draw:name="one"/></office:presentation></office:body></office:document>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", content).unwrap();
    let snapshot = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(snapshot.slides().len(), 1);
    assert!(snapshot.transaction().is_err());
}

#[test]
fn failed_media_preflight_leaves_an_exact_noop_transaction() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .embed_media("../escape.bin", b"payload", "application/octet-stream")
            .is_err()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
}

#[test]
fn changed_commit_reopens_every_staged_media_part_and_manifest_entry() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .embed_media("Media/one.bin", b"first", "application/octet-stream")
        .unwrap();
    transaction
        .embed_media("Media/two.bin", b"second", "application/x-litchi-test")
        .unwrap();

    let commit = transaction.commit().unwrap();
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let reopened = package.package().unwrap();
    for (path, bytes, media_type) in [
        (
            "Media/one.bin",
            b"first".as_slice(),
            "application/octet-stream",
        ),
        (
            "Media/two.bin",
            b"second".as_slice(),
            "application/x-litchi-test",
        ),
    ] {
        assert_eq!(reopened.get_file(path).unwrap(), bytes);
        assert!(reopened.manifest().get_entry(path).is_some());
        assert_eq!(reopened.manifest().get_media_type(path), Some(media_type));
    }
}
