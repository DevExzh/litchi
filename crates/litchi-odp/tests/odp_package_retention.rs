#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::core::{PackageWriter, Profile};
use litchi_odp::{Builder, Package, edit};
use soapberry_zip::office::StreamingArchiveWriter;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PASSWORD: &str = "retention-password";

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn simple_package() -> Vec<u8> {
    package(&format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><office:body><office:presentation><draw:page draw:name="one"/></office:presentation></office:body></office:document-content>"#
    ))
}

fn password_package() -> Vec<u8> {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption(PASSWORD, Profile::compatible())
        .unwrap();
    writer
        .add_file(
            "content.xml",
            format!(
                r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:presentation/></office:body></office:document-content>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn snapshot_transaction_and_noop_commit_retain_one_archive_index() {
    let source = Package::from_bytes(&simple_package()).unwrap();
    let identity = source.prepared_index_identity();
    assert_ne!(identity, 0);
    assert_eq!(
        source.presentation().unwrap().prepared_index_identity(),
        identity
    );

    let transaction = source.edit().unwrap();
    assert_eq!(transaction.prepared_index_identity(), identity);
    let commit = transaction.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().prepared_index_identity(), identity);
    assert_eq!(commit.package().as_bytes(), source.as_bytes());
}

#[test]
fn changed_publication_retains_new_index_and_patch_handoffs_do_not_rebuild_it() {
    let source = Package::from_bytes(&simple_package()).unwrap();
    let source_identity = source.prepared_index_identity();
    let mut transaction = source.edit().unwrap();
    transaction.add("two", "body").unwrap();
    let commit = transaction.commit().unwrap();
    let target_identity = commit.package().prepared_index_identity();
    assert_ne!(target_identity, 0);
    assert_ne!(target_identity, source_identity);
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .prepared_index_identity(),
        target_identity
    );

    let applied = source.apply(commit.patch()).unwrap().into_package();
    assert_eq!(applied.prepared_index_identity(), target_identity);
    let restored = applied
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.prepared_index_identity(), source_identity);
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn stale_patch_still_fails_even_when_the_artifact_is_reopened() {
    let source = Package::from_bytes(&simple_package()).unwrap();
    let mut transaction = source.edit().unwrap();
    transaction.add("two", "body").unwrap();
    let commit = transaction.commit().unwrap();
    let reopened = Package::from_bytes(source.as_bytes()).unwrap();
    assert_ne!(
        reopened.prepared_index_identity(),
        source.prepared_index_identity()
    );
    assert!(reopened.apply(commit.patch()).is_ok());

    let stale = Package::from_bytes(&package(&format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><office:body><office:presentation><draw:page draw:name="different"/></office:presentation></office:body></office:document-content>"#
    )))
    .unwrap();
    assert!(stale.apply(commit.patch()).is_err());
}

#[test]
fn signed_and_encrypted_sources_retain_read_reopen_but_refuse_mutation() {
    let mut signed_writer =
        PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    signed_writer.set_mimetype(MIME).unwrap();
    signed_writer
        .add_file(
            "content.xml",
            format!(
                r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:presentation/></office:body></office:document-content>"#
            )
            .as_bytes(),
        )
        .unwrap();
    signed_writer
        .add_file(
            "META-INF/documentsignatures.xml",
            br#"<dsig:document-signatures xmlns:dsig="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
        )
        .unwrap();
    let signed = Package::from_bytes(&signed_writer.finish_to_bytes().unwrap()).unwrap();
    let signed_identity = signed.prepared_index_identity();
    // The constructor must not enumerate archive names to discover the
    // signature owner. Admission performs that census on first use.
    let signed_edit = signed.edit().unwrap();
    assert_eq!(
        signed.security_policy().unwrap(),
        edit::SecurityPolicy::SignedReadOnly
    );
    assert_eq!(
        signed.presentation().unwrap().prepared_index_identity(),
        signed_identity
    );
    assert!(signed_edit.commit().is_err());

    const MANIFEST: &[u8] = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.presentation"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="secret.bin" m:media-type="application/octet-stream" m:size="1"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#;
    let mut encrypted_writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    encrypted_writer
        .write_stored("mimetype", MIME.as_bytes())
        .unwrap();
    encrypted_writer
        .write_deflated(
            "content.xml",
            format!(
                r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:presentation/></office:body></office:document-content>"#
            )
            .as_bytes(),
        )
        .unwrap();
    encrypted_writer.write_deflated("secret.bin", b"x").unwrap();
    encrypted_writer
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    let encrypted = Package::from_bytes(&encrypted_writer.finish_to_bytes().unwrap()).unwrap();
    let encrypted_identity = encrypted.prepared_index_identity();
    // Encrypted-entry census is likewise deferred until first admission.
    let encrypted_edit = encrypted.edit().unwrap();
    assert_eq!(
        encrypted.security_policy().unwrap(),
        edit::SecurityPolicy::EncryptedReadOnly
    );
    assert_eq!(
        encrypted.presentation().unwrap().prepared_index_identity(),
        encrypted_identity
    );
    assert!(encrypted_edit.commit().is_err());
}

#[test]
fn noncompact_fallback_keeps_exact_noop_and_retained_index() {
    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive
        .write_deflated(
            "content.xml",
            format!(
                r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><office:body><office:presentation>
<draw:page draw:name="one"/>
</office:presentation></office:body></office:document-content>"#
            )
            .as_bytes(),
        )
        .unwrap();
    archive
        .write_deflated(
            "META-INF/manifest.xml",
            format!(
                r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#
            )
            .as_bytes(),
        )
        .unwrap();
    let source = Package::from_bytes(&archive.finish_to_bytes().unwrap()).unwrap();
    let identity = source.prepared_index_identity();
    let mut transaction = source.edit().unwrap();
    assert!(transaction.copy_dependency_free_blank_slide(0).is_err());
    let commit = transaction.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().prepared_index_identity(), identity);
    assert_eq!(commit.package().as_bytes(), source.as_bytes());
}

#[test]
fn ordinary_presentation_snapshot_transfers_the_existing_index() {
    let snapshot =
        Package::from_bytes(&Builder::new().build().unwrap().to_bytes().unwrap()).unwrap();
    let presentation = snapshot.presentation().unwrap();
    let identity = presentation.prepared_index_identity();
    assert_eq!(snapshot.prepared_index_identity(), identity);
}

#[test]
fn password_opened_editing_and_chart_snapshots_drop_credentials_after_validation() {
    let snapshot = Package::from_bytes_with_password(&password_package(), PASSWORD).unwrap();
    // The read projection keeps the caller's credential so encrypted parts
    // stay readable; the editing owner never receives it.
    let presentation = snapshot.presentation().unwrap();
    let identity = presentation.prepared_index_identity();
    assert_eq!(presentation.slide_count().unwrap(), 0);
    assert_eq!(snapshot.prepared_index_identity(), identity);
    assert_eq!(
        snapshot.security_policy().unwrap(),
        edit::SecurityPolicy::EncryptedReadOnly
    );
    let snapshot_edit = snapshot.edit().unwrap();
    assert!(snapshot_edit.commit().is_err());

    let charts = presentation.chart_snapshot().unwrap();
    assert_eq!(charts.prepared_index_identity(), identity);
    assert!(charts.charts().is_empty());
    assert!(charts.to_presentation().is_err());
}

#[test]
fn signed_policy_precedes_deferred_processing_instruction_validation() {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file(
            "content.xml",
            format!(
                r#"<office:document-content xmlns:office="{OFFICE}"><office:body><?producer retained?><office:presentation/></office:body></office:document-content>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .add_file(
            "META-INF/documentsignatures.xml",
            br#"<dsig:document-signatures xmlns:dsig="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
        )
        .unwrap();
    let source = Package::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap();

    let error = match source.edit().unwrap().commit() {
        Ok(_) => panic!("signed package with active processing instruction committed"),
        Err(error) => error,
    };
    assert!(error.to_string().contains("signed packages"));
}
