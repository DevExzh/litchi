#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::{CompositionLimits, Error, HistoryLimits, MergeChoice, SubEditJoinFailure};
use litchi_odf_common::{
    compact_xml,
    core::{OwnedPackage, PackageWriter},
};
use litchi_odg::{
    Builder, DocumentSigner, EncryptionProfile, FormControl, Package, SignatureAlgorithm,
    SignatureValidity, edit,
    page::Page,
    shape::{Shape, ShapeKind},
    style::Style,
    style_resource::{StyleResource, StyleResourceKind},
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.3"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1" draw:master-page-name="Default" xml:id="page1"><draw:layer-set><draw:layer draw:name="Foreground" draw:display="always" draw:protected="false"/><draw:layer draw:name="Background"/></draw:layer-set><draw:rect draw:name="Label" draw:layer="Foreground" draw:style-name="gr1" draw:text-style-name="P1" draw:z-index="7" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><svg:title>Label title</svg:title><svg:desc>Label description</svg:desc><text:p>Old label</text:p></draw:rect><draw:frame draw:name="Photo" draw:layer="Background" svg:width="2cm" svg:height="1cm"><svg:title>Photo title</svg:title><svg:desc>Photo description</svg:desc></draw:frame></draw:page></office:drawing></office:body></office:document-content>"#;

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn raw_negative_fixture_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

fn raw_security_fixture(encrypted: bool, signed: bool) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    let encrypted_entry = if encrypted {
        r#"<manifest:file-entry manifest:full-path="Pictures/protected.bin" manifest:media-type="application/octet-stream" manifest:size="1"><manifest:encryption-data manifest:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" manifest:checksum="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="><manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" manifest:initialisation-vector="AAAAAAAAAAAAAAAAAAAAAA=="/><manifest:start-key-generation manifest:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" manifest:key-size="32"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:salt="AAAAAAAAAAAAAAAAAAAAAA==" manifest:iteration-count="1000" manifest:key-size="32"/></manifest:encryption-data></manifest:file-entry>"#
    } else {
        ""
    };
    let manifest = format!(
        r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>{encrypted_entry}</manifest:manifest>"#
    );
    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", CONTENT.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", manifest.as_bytes())
        .unwrap();
    if encrypted {
        archive
            .write_stored("Pictures/protected.bin", b"x")
            .unwrap();
    }
    if signed {
        archive
            .write_deflated(
                "META-INF/documentsignatures.xml",
                br"<document-signatures/>",
            )
            .unwrap();
    }
    archive.finish_to_bytes().unwrap()
}

#[test]
fn semantic_package_views_use_layers_shapes_and_shared_frame_context() {
    let drawing = Package::from_bytes(&package(CONTENT)).unwrap();
    assert_eq!(
        drawing.page(0usize).unwrap().unwrap().name(),
        Some("Page 1")
    );
    assert_eq!(
        drawing.page("Page 1").unwrap().unwrap().name(),
        Some("Page 1")
    );
    assert!(drawing.page(9usize).unwrap().is_none());
    assert_eq!(drawing.pages().len(), 1);
    assert_eq!(drawing.pages()[0].name(), Some("Page 1"));
    assert_eq!(drawing.pages()[0].xml_id(), Some("page1"));
    assert_eq!(drawing.pages()[0].style_name(), Some("dp1"));
    assert_eq!(drawing.pages()[0].master_page_name(), Some("Default"));
    assert!(drawing.layers().is_empty());
    assert_eq!(drawing.pages()[0].layers().len(), 2);
    assert_eq!(drawing.pages()[0].layers()[0].name(), "Foreground");
    assert_eq!(drawing.pages()[0].layers()[0].display(), Some("always"));
    assert_eq!(drawing.pages()[0].layers()[0].protected(), Some(false));
    let shape = &drawing.pages()[0].shapes()[0];
    assert_eq!(drawing.pages()[0].shape("Label").unwrap(), Some(shape));
    assert_eq!(drawing.pages()[0].shape(0usize).unwrap(), Some(shape));
    assert!(drawing.pages()[0].shape("Missing").unwrap().is_none());
    assert_eq!(shape.name(), Some("Label"));
    assert_eq!(shape.layer(), Some("Foreground"));
    assert_eq!(shape.kind(), ShapeKind::Rectangle);
    assert_eq!(shape.style_name(), Some("gr1"));
    assert_eq!(shape.text_style_name(), Some("P1"));
    assert_eq!(shape.z_index(), Some(7));
    assert_eq!(shape.x(), Some("1cm"));
    assert_eq!(shape.y(), Some("2cm"));
    assert_eq!(shape.width(), Some("3cm"));
    assert_eq!(shape.height(), Some("4cm"));
    assert_eq!(shape.title(), Some("Label title"));
    assert_eq!(shape.description(), Some("Label description"));
    assert_eq!(shape.text(), "Old label");
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .width
            .as_deref(),
        Some("2cm")
    );
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .title
            .as_deref(),
        Some("Photo title")
    );
}

#[test]
fn exact_shape_name_selector_rejects_ambiguity() {
    let duplicate = CONTENT.replace("draw:name=\"Photo\"", "draw:name=\"Label\"");
    let drawing = Package::from_bytes(&package(&duplicate)).unwrap();
    assert!(drawing.pages()[0].shape("Label").is_err());
}

#[test]
fn package_edit_is_source_checked_reversible_and_compact() {
    let drawing = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_text(0, 0, "New <label>").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.package().pages()[0].shapes()[0].text(),
        "New <label>"
    );
    assert!(
        commit
            .package()
            .content_xml()
            .contains("<text:p>New &lt;label&gt;</text:p>")
    );
    assert!(!commit.package().content_xml().contains(">\n<"));
    let output = OwnedPackage::from_bytes(commit.package().as_bytes().to_vec()).unwrap();
    let archive = output.package().unwrap();
    for path in ["content.xml", "META-INF/manifest.xml"] {
        compact_xml::validate(
            &archive.get_file(path).unwrap(),
            compact_xml::Limits::default(),
        )
        .unwrap();
    }
    assert!(!commit.package().content_xml().contains(">\r\n<"));
    assert!(commit.patch().is_applicable_to(&drawing));
    let reapplied = drawing.apply(commit.patch()).unwrap();
    assert_eq!(
        reapplied.into_package().as_bytes(),
        commit.package().as_bytes()
    );
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.into_package().as_bytes(), drawing.as_bytes());
}

#[test]
fn package_shape_name_edit_is_source_checked_reversible_and_compact() {
    let drawing = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_name(0, 0, "Renamed & exact").unwrap();
    let commit = transaction.commit().unwrap();

    let shape = &commit.package().pages()[0].shapes()[0];
    assert_eq!(shape.name(), Some("Renamed & exact"));
    assert!(
        commit
            .package()
            .content_xml()
            .contains("draw:name=\"Renamed &amp; exact\"")
    );
    assert!(
        commit
            .package()
            .content_xml()
            .contains("draw:layer=\"Foreground\"")
    );
    assert!(
        commit
            .package()
            .content_xml()
            .contains("draw:name=\"Photo\"")
    );
    assert!(!commit.package().content_xml().contains(">\n<"));
    assert!(!commit.package().content_xml().contains(">\r\n<"));

    let change = commit.patch().name_change().unwrap();
    assert_eq!(change.before(), "Label");
    assert_eq!(change.after(), "Renamed & exact");
    assert!(commit.patch().is_applicable_to(&drawing));
    let different =
        Package::from_bytes(&package(&CONTENT.replace("Photo", "Different photo"))).unwrap();
    assert!(!commit.patch().is_applicable_to(&different));
    assert!(different.apply(commit.patch()).is_err());
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.into_package().as_bytes(), drawing.as_bytes());
}

#[test]
fn page_rename_is_durable_unique_and_dependency_checked() {
    let drawing = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut edit = drawing.edit();
    edit.set_page_name(0, "Renamed page").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.package().pages()[0].name(), Some("Renamed page"));
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "page.name.set");
    assert_eq!(
        durable.apply(&drawing).unwrap().as_bytes(),
        commit.package().as_bytes()
    );

    let referenced = CONTENT.replace(
        "</draw:page></office:drawing>",
        r#"</draw:page><draw:page draw:name="Reference"><draw:page-thumbnail draw:page-name="Page 1"/></draw:page></office:drawing>"#,
    );
    let referenced_drawing = Package::from_bytes(&package(&referenced)).unwrap();
    let mut refused = referenced_drawing.edit();
    assert!(refused.set_page_name(0, "Unsafe rename").is_err());
    let mut removal = referenced_drawing.edit();
    assert!(removal.remove_page(0usize).is_err());
}

#[test]
fn page_style_change_requires_a_declared_style_and_roundtrips() {
    let styled = CONTENT
        .replace(
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace""#,
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0""#,
        )
        .replace(
            "<office:body>",
            r#"<office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"/><style:style style:name="dp2" style:family="drawing-page"/></office:automatic-styles><office:body>"#,
        );
    let drawing = Package::from_bytes(&package(&styled)).unwrap();
    let mut edit = drawing.edit();
    edit.set_page_style_name(0, "dp2").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.package().pages()[0].style_name(), Some("dp2"));
    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        drawing.as_bytes()
    );

    let mut missing = drawing.edit();
    assert!(missing.set_page_style_name(0, "missing").is_err());
}

#[test]
fn package_layer_edit_is_declared_source_checked_and_reversible() {
    let drawing = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_layer(0, 0, "Background").unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.package().pages()[0].shapes()[0].layer(),
        Some("Background")
    );
    assert!(
        commit
            .package()
            .content_xml()
            .contains("draw:layer=\"Background\"")
    );
    assert!(!commit.package().content_xml().contains(">\n<"));
    let change = commit.patch().layer_change().unwrap();
    assert_eq!(change.before(), "Foreground");
    assert_eq!(change.after(), "Background");
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.into_package().as_bytes(), drawing.as_bytes());

    let mut invalid = drawing.edit();
    assert!(invalid.set_shape_layer(0, 0, "Missing").is_err());
}

#[test]
fn ended_page_scope_does_not_capture_shapes_outside_a_page() {
    let malformed = CONTENT.replace(
        "</draw:page></office:drawing>",
        "</draw:page><draw:g><draw:rect draw:name=\"outside\"/></draw:g></office:drawing>",
    );
    assert!(Package::from_bytes(&package(&malformed)).is_err());
}

#[test]
fn dtd_is_refused_and_noncompact_source_requires_checked_splices() {
    let dtd = CONTENT.replacen("<office:body>", "<!DOCTYPE drawing><office:body>", 1);
    assert!(Package::from_bytes(&raw_negative_fixture_package(&dtd)).is_err());

    let noncompact = CONTENT.replacen("<office:body>", "\n<office:body>", 1);
    let drawing = Package::from_bytes(&raw_negative_fixture_package(&noncompact)).unwrap();
    let mut checked_splice = drawing.edit();
    checked_splice.set_shape_text(0, 0, "New label").unwrap();
    let commit = checked_splice.commit().unwrap();
    assert_eq!(
        commit.package().content_xml(),
        noncompact.replace("Old label", "New label")
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        drawing.as_bytes()
    );

    // Page insertion is a checked splice too (record 0288): the producer's
    // whitespace survives.
    let mut inserted = drawing.edit();
    inserted.add_page(Page::new("Page 2")).unwrap();
    let inserted = inserted.commit().unwrap();
    assert!(inserted.package().content_xml().contains("\n<office:body>"));
    assert_eq!(inserted.package().drawing().pages().len(), 2);

    // Edits without a splice form still rewrite the whole part, which a
    // noncompact source refuses.
    let mut whole_source_rewrite = drawing.edit();
    whole_source_rewrite
        .add_layer(0, litchi_odg::layer::Layer::new("extra"))
        .unwrap();
    assert!(whole_source_rewrite.commit().is_err());
}

#[test]
fn signed_and_encrypted_security_state_is_inert_and_rewrite_is_refused() {
    for (encrypted, signed) in [(true, false), (false, true), (true, true)] {
        let drawing = Package::from_bytes(&raw_security_fixture(encrypted, signed)).unwrap();
        assert_eq!(drawing.security().is_encrypted(), encrypted);
        assert_eq!(drawing.security().is_signed(), signed);
        assert!(!drawing.security().allows_rewrite());
        let mut edit = drawing.edit();
        edit.set_shape_text(0, 0, "must refuse").unwrap();
        assert!(edit.commit().is_err());
    }

    let signed = Package::from_bytes(&raw_security_fixture(false, true)).unwrap();
    let mut explicit =
        signed.edit_with_security_policy(edit::SecurityWritePolicy::RemoveSignatures);
    explicit.set_shape_text(0, 0, "signature removed").unwrap();
    let commit = explicit.commit().unwrap();
    assert!(!commit.package().security().is_signed());
    let durable = commit.patch().durable().unwrap();
    assert!(durable.apply(&signed).is_err());
    assert_eq!(
        durable
            .apply_with_security_policy(&signed, edit::SecurityWritePolicy::RemoveSignatures,)
            .unwrap()
            .as_bytes(),
        commit.package().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        signed.as_bytes()
    );

    let encrypted = Package::from_bytes(&raw_security_fixture(true, false)).unwrap();
    let mut still_refused =
        encrypted.edit_with_security_policy(edit::SecurityWritePolicy::RemoveSignatures);
    still_refused.set_shape_text(0, 0, "no decryption").unwrap();
    assert!(still_refused.commit().is_err());
}

#[test]
fn active_content_is_inventoried_preserved_or_explicitly_refused() {
    let active = CONTENT
        .replace(
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace""#,
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ev="http://www.w3.org/2001/xml-events" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink""#,
        )
        .replace(
            "</draw:page>",
            r#"<script:event-listener xlink:href="https://example.invalid/script"/><ev:event-listener/><presentation:show/><office:dde-source/><draw:object xlink:href="https://example.invalid/object"/></draw:page>"#,
        );
    let source = Package::from_bytes(&package(&active)).unwrap();
    let status = source.active_content();
    assert!(status.is_present());
    assert_eq!(status.scripts(), 1);
    assert_eq!(status.events(), 1);
    assert_eq!(status.actions(), 1);
    assert_eq!(status.dde(), 1);
    assert_eq!(status.external_links(), 2);
    assert_eq!(status.embedded_objects(), 1);

    let mut preserved = source.edit();
    preserved.set_shape_text(0, 0, "preserved inertly").unwrap();
    let commit = preserved.commit().unwrap();
    assert_eq!(commit.package().active_content(), status);
    assert!(
        commit
            .package()
            .content_xml()
            .contains("script:event-listener")
    );

    let mut refused =
        source.edit_with_active_content_policy(edit::ActiveContentWritePolicy::Refuse);
    refused.set_shape_text(0, 0, "refused").unwrap();
    assert!(refused.commit().is_err());

    let durable = commit.patch().durable().unwrap();
    assert!(
        durable
            .apply_with_policies(
                &source,
                edit::SecurityWritePolicy::Refuse,
                edit::ActiveContentWritePolicy::Refuse,
            )
            .is_err()
    );
    let limits = CompositionLimits::new(4, 4, 8, 8);
    let mut joined = source.joined_edits(limits);
    joined
        .join(commit.patch().prepare("active", limits).unwrap())
        .unwrap();
    let mut joined_edit = source.edit_with_policies(
        edit::SecurityWritePolicy::Refuse,
        edit::ActiveContentWritePolicy::Refuse,
    );
    joined_edit.join(joined).unwrap();
    assert!(joined_edit.commit().is_err());

    let safe = Package::from_bytes(&package(CONTENT)).unwrap();
    let external_fill = StyleResource::new(StyleResourceKind::FillImage, "external-fill")
        .with_attribute("xlink:href", "https://example.invalid/fill.png")
        .with_attribute("xlink:type", "simple");
    let mut target_refused =
        safe.edit_with_active_content_policy(edit::ActiveContentWritePolicy::Refuse);
    target_refused.put_style_resource(&external_fill).unwrap();
    assert!(target_refused.commit().is_err());
    let mut target_preserved = safe.edit();
    target_preserved.put_style_resource(&external_fill).unwrap();
    let target_commit = target_preserved.commit().unwrap();
    assert_eq!(target_commit.package().active_content().external_links(), 1);
    assert!(
        target_commit
            .package()
            .content_xml()
            .contains("https://example.invalid/fill.png")
    );
}

#[test]
fn advanced_named_style_resources_are_dependency_checked_and_durable() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut edit = source.edit();
    edit.put_style_resource(
        &StyleResource::new(StyleResourceKind::Gradient, "advanced-gradient")
            .with_attribute("draw:style", "linear")
            .with_attribute("draw:start-color", "#112233")
            .with_attribute("draw:end-color", "#ddeeff"),
    )
    .unwrap();
    edit.put_style(
        &Style::new("advanced-style", "graphic")
            .with_property("draw:fill", "gradient")
            .with_property("draw:fill-gradient-name", "advanced-gradient"),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.package().style_resources().iter().any(|resource| {
        resource.kind() == StyleResourceKind::Gradient && resource.name() == "advanced-gradient"
    }));
    let mut blocked = commit.package().edit();
    assert!(
        blocked
            .remove_style_resource(StyleResourceKind::Gradient, "advanced-gradient")
            .is_err()
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn fresh_password_and_signing_lifecycle_is_explicit_and_verifiable() {
    const RSA_KEY: &[u8] =
        include_bytes!("../../litchi-odf-common/tests/fixtures/signatures/rsa-key.pk8");
    const RSA_CERT: &[u8] =
        include_bytes!("../../litchi-odf-common/tests/fixtures/signatures/rsa-cert.der");
    let signer = DocumentSigner::from_pkcs8_der(
        SignatureAlgorithm::RsaSha256,
        RSA_KEY,
        vec![RSA_CERT.to_vec()],
        "2026-08-10T12:00:00Z",
    )
    .unwrap();
    let built = Builder::new()
        .build_encrypted_and_signed("drawing-password", EncryptionProfile::compatible(), signer)
        .unwrap();
    let bytes = built.as_bytes();
    assert!(Package::from_bytes_with_password(bytes, "wrong-password").is_err());
    let drawing = Package::from_bytes_with_password(bytes, "drawing-password").unwrap();
    assert_eq!(drawing.as_bytes(), bytes);
    assert!(drawing.security().is_encrypted());
    assert!(drawing.security().is_signed());
    let exact = drawing.edit().commit().unwrap();
    assert!(!exact.diagnostics().changed());
    assert_eq!(exact.package().as_bytes(), drawing.as_bytes());
    let mut changed_encrypted = drawing.edit();
    changed_encrypted
        .add_page(Page::new("Refused change"))
        .unwrap();
    assert!(changed_encrypted.commit().is_err());
    let capabilities = drawing.security_capabilities();
    assert_eq!(capabilities.source(), drawing.security());
    assert!(capabilities.can_open_with_password());
    assert!(capabilities.can_encrypt_new());
    assert!(!capabilities.can_reencrypt_existing());
    assert!(capabilities.can_verify_signatures());
    assert!(capabilities.can_sign_new());
    assert!(!capabilities.can_resign_existing());
    assert!(capabilities.can_remove_invalidated_signatures());
    assert_eq!(
        capabilities.disposition(edit::SecurityLifecycleOperation::RewriteEncryptedExisting),
        edit::SecurityLifecycleDisposition::ExactSourceOnly
    );
    assert_eq!(
        capabilities.disposition(edit::SecurityLifecycleOperation::ChangeExistingPassword),
        edit::SecurityLifecycleDisposition::Unsupported(
            edit::SecurityLifecycleRefusal::ExistingPackageReencryptionUnavailable
        )
    );
    assert_eq!(
        capabilities.disposition(edit::SecurityLifecycleOperation::ResignExisting),
        edit::SecurityLifecycleDisposition::Unsupported(
            edit::SecurityLifecycleRefusal::ExistingPackageResigningUnavailable
        )
    );
    assert_eq!(
        drawing
            .digital_signatures()
            .unwrap()
            .document_signatures
            .len(),
        1
    );
    let verification = drawing.verify_document_signatures().unwrap();
    assert_eq!(verification.len(), 1);
    assert_eq!(verification[0].validity, SignatureValidity::Valid);
}

#[test]
fn hostile_raw_package_corpus_is_bounded_and_never_partially_opened() {
    let duplicate_drawing = CONTENT.replace(
        "</office:body>",
        "<office:drawing></office:drawing></office:body>",
    );
    let nested_page = CONTENT.replace(
        "</draw:page>",
        "<draw:page draw:name=\"nested\"/></draw:page>",
    );
    let layer_outside_page = CONTENT.replace(
        "</office:drawing>",
        "<draw:layer-set><draw:layer draw:name=\"outside\"/></draw:layer-set></office:drawing>",
    );
    let active_style = CONTENT
        .replace(
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace""#,
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0""#,
        )
        .replace(
            "<office:body>",
            r#"<office:automatic-styles><style:style style:name="active" style:family="graphic"><script:event-listener/></style:style></office:automatic-styles><office:body>"#,
        );
    for malformed in [duplicate_drawing, nested_page, layer_outside_page] {
        assert!(Package::from_bytes(&raw_negative_fixture_package(&malformed)).is_err());
    }
    assert!(Package::from_bytes(&raw_negative_fixture_package(&active_style)).is_err());

    let prefix = CONTENT.replace("</draw:page>", "");
    let deep = format!(
        "{}{}{}{}",
        prefix.trim_end_matches("</office:drawing></office:body></office:document-content>"),
        "<draw:g>".repeat(300),
        "</draw:g>".repeat(300),
        "</draw:page></office:drawing></office:body></office:document-content>"
    );
    assert!(Package::from_bytes(&raw_negative_fixture_package(&deep)).is_err());

    let valid_raw = raw_negative_fixture_package(CONTENT);
    let step = (valid_raw.len() / 64).max(1);
    for end in (0..valid_raw.len()).step_by(step) {
        assert!(Package::from_bytes(&valid_raw[..end]).is_err());
    }
}

#[test]
fn durable_semantic_patch_roundtrips_inverts_and_refuses_stale_sources() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut edit = source.edit();
    edit.set_shape_text(0, 0, "durable text").unwrap();
    edit.set_shape_geometry(0, 0, "2cm", "3cm", "5cm", "6cm")
        .unwrap();
    edit.set_shape_style_name(0, 0, "gr2").unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    assert_eq!(wire, durable.to_deterministic_json().unwrap());
    let reopened = edit::DurablePatch::from_deterministic_json(&wire).unwrap();
    assert_eq!(reopened.operations().len(), 3);
    assert_eq!(
        reopened.apply(&source).unwrap().as_bytes(),
        commit.package().as_bytes()
    );
    assert_eq!(
        reopened
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );

    let stale = Package::from_bytes(&package(&CONTENT.replace("Photo", "Changed"))).unwrap();
    assert!(matches!(reopened.apply(&stale), Err(Error::PatchConflict)));
}

#[test]
fn durable_structural_patch_reopens_full_package_and_inverts_exactly() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut edit = source.edit();
    edit.add_page(Page::new("Added")).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "package.replace");
    let reopened =
        edit::DurablePatch::from_deterministic_json(&durable.to_deterministic_json().unwrap())
            .unwrap();
    let applied = reopened.apply(&source).unwrap();
    assert_eq!(applied.as_bytes(), commit.package().as_bytes());
    assert_eq!(applied.pages().len(), 2);
    assert_eq!(
        reopened.inverse().apply(&applied).unwrap().as_bytes(),
        source.as_bytes()
    );

    let limits = CompositionLimits::new(4, 4, 8, 8);
    let mut joined = source.joined_edits(limits);
    joined
        .join(commit.patch().prepare("structural", limits).unwrap())
        .unwrap();
    let mut joined_edit = source.edit();
    joined_edit.join(joined).unwrap();
    let joined_commit = joined_edit.commit().unwrap();
    assert_eq!(joined_commit.package().as_bytes(), applied.as_bytes());
    assert_eq!(
        source
            .apply(joined_commit.patch())
            .unwrap()
            .package()
            .as_bytes(),
        applied.as_bytes()
    );
    assert_eq!(
        joined_commit
            .package()
            .apply(&joined_commit.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn path_data_is_typed_losslessly_editable_and_durable() {
    let content = CONTENT.replace(
        "</draw:page>",
        r#"<draw:path draw:name="Route" svg:d="M 0 0 L 10 10"/></draw:page>"#,
    );
    let source = Package::from_bytes(&package(&content)).unwrap();
    assert_eq!(
        source.pages()[0].shapes()[2].path_data(),
        Some("M 0 0 L 10 10")
    );
    let mut edit = source.edit();
    edit.set_shape_path_data(0, 2, "M 1 2 C 3 4 5 6 7 8")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.package().pages()[0].shapes()[2].path_data(),
        Some("M 1 2 C 3 4 5 6 7 8")
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "shape.path.set");
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.package().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn transforms_points_and_line_endpoints_insert_reopen_and_invert() {
    let content = CONTENT.replace(
        "</draw:page>",
        r#"<draw:polyline draw:name="Route"/><draw:line draw:name="Axis"/></draw:page>"#,
    );
    let source = Package::from_bytes(&package(&content)).unwrap();
    let mut refused = source.edit();
    assert!(
        refused
            .set_shape_points(0, 0, "0 0 1 1", "0,0 1,1")
            .is_err()
    );
    assert_eq!(
        refused.commit().unwrap().package().as_bytes(),
        source.as_bytes()
    );
    let mut edit = source.edit();
    edit.set_shape_transform(0, 0, "rotate (15) translate (1cm 2cm)")
        .unwrap();
    edit.set_shape_points(0, 2, "0 0 1000 1000", "0,0 500,1000 1000,0")
        .unwrap();
    edit.set_shape_line_geometry(0, 3, "1cm", "2cm", "8cm", "9cm")
        .unwrap();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Polygon)
            .with_name("Authored polygon")
            .with_points("0 0 10 10", "0,0 5,10 10,0")
            .with_transform("translate (2cm 3cm)"),
    )
    .unwrap();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Connector)
            .with_name("Authored connector")
            .with_line_geometry("0cm", "0cm", "4cm", "5cm"),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let shapes = commit.package().pages()[0].shapes();
    assert_eq!(
        shapes[0].transform(),
        Some("rotate (15) translate (1cm 2cm)")
    );
    assert_eq!(shapes[2].view_box(), Some("0 0 1000 1000"));
    assert_eq!(shapes[2].points(), Some("0,0 500,1000 1000,0"));
    assert_eq!(
        shapes[3].line_geometry(),
        [Some("1cm"), Some("2cm"), Some("8cm"), Some("9cm")]
    );
    assert_eq!(shapes[4].points(), Some("0,0 5,10 10,0"));
    assert_eq!(shapes[4].transform(), Some("translate (2cm 3cm)"));
    assert_eq!(
        shapes[5].line_geometry(),
        [Some("0cm"), Some("0cm"), Some("4cm"), Some("5cm")]
    );
    let reopened = Package::from_bytes(commit.package().as_bytes()).unwrap();
    assert_eq!(reopened.as_bytes(), commit.package().as_bytes());
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
    assert!(!commit.package().active_content().is_present());
}

#[test]
fn form_controls_remain_inert_but_their_references_are_editable() {
    let content = CONTENT.replace(
        "</draw:page>",
        r#"<draw:control draw:name="Button" draw:control="button1"/></draw:page>"#,
    );
    let source = Package::from_bytes(&package(&content)).unwrap();
    let control = &source.pages()[0].shapes()[2];
    assert_eq!(control.kind(), ShapeKind::Control);
    assert_eq!(control.control_reference(), Some("button1"));
    let mut edit = source.edit();
    edit.set_shape_control_reference(0, 2, "button2").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.package().pages()[0].shapes()[2].control_reference(),
        Some("button2")
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "shape.control.set");
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.package().as_bytes()
    );
}

#[test]
fn form_model_crud_is_inert_dependency_checked_and_durable() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut edit = source.edit();
    edit.add_form_control(
        &FormControl::new("control1")
            .with_name("Button")
            .with_element("checkbox")
            .with_attribute("form:label", "Arbitrary label")
            .with_attribute("form:disabled", "true"),
    )
    .unwrap();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Control)
            .with_name("Button shape")
            .with_control_reference("control1"),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.package().form_controls()[0].id(), "control1");
    assert_eq!(commit.package().form_controls()[0].element(), "checkbox");
    assert_eq!(
        commit.package().form_controls()[0]
            .attributes()
            .get("form:label")
            .map(String::as_str),
        Some("Arbitrary label")
    );
    assert_eq!(
        commit.package().pages()[0].shapes()[2].control_reference(),
        Some("control1")
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.package().as_bytes()
    );

    let mut blocked = commit.package().edit();
    assert!(blocked.remove_form_control("control1").is_err());

    let mut removal = commit.package().edit();
    removal.remove_shape(0, 2).unwrap();
    removal.remove_form_control("control1").unwrap();
    let removed = removal.commit().unwrap();
    assert!(removed.package().form_controls().is_empty());
}

#[test]
fn arbitrary_style_properties_have_dependency_checked_crud_and_exact_inverse() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let style = Style::new("custom-fill", "graphic")
        .with_property("draw:fill", "solid")
        .with_property("draw:fill-color", "#12ab34")
        .with_property("fo:opacity", "75%");
    let mut edit = source.edit();
    edit.put_style(&style).unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit
            .package()
            .style_definitions()
            .iter()
            .any(|candidate| candidate == &style)
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );

    let mut removal = commit.package().edit();
    assert_eq!(removal.remove_style("custom-fill").unwrap(), style);
    assert!(
        removal
            .commit()
            .unwrap()
            .package()
            .style_definitions()
            .is_empty()
    );
}

#[test]
fn joined_edits_are_deterministic_disjoint_and_conflict_aware() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let limits = CompositionLimits::new(8, 8, 32, 16);

    let mut text_edit = source.edit();
    text_edit.set_shape_text(0, 0, "joined").unwrap();
    let text_commit = text_edit.commit().unwrap();
    let mut name_edit = source.edit();
    name_edit.set_shape_name(0, 1, "Joined photo").unwrap();
    let name_commit = name_edit.commit().unwrap();

    let mut joined = source.joined_edits(limits);
    joined
        .join(text_commit.patch().prepare("z-text", limits).unwrap())
        .unwrap();
    joined
        .join(name_commit.patch().prepare("a-name", limits).unwrap())
        .unwrap();
    assert_eq!(
        joined
            .sub_edits()
            .map(edit::SubEdit::id)
            .collect::<Vec<_>>(),
        ["a-name", "z-text"]
    );
    let mut joined_edit = source.edit();
    joined_edit.join(joined).unwrap();
    let joined_commit = joined_edit.commit().unwrap();
    let applied = joined_commit.package().clone();
    assert_eq!(applied.pages()[0].shapes()[0].text(), "joined");
    assert_eq!(applied.pages()[0].shapes()[1].name(), Some("Joined photo"));
    assert_eq!(
        source
            .apply(joined_commit.patch())
            .unwrap()
            .package()
            .as_bytes(),
        applied.as_bytes()
    );
    assert_eq!(
        applied
            .apply(&joined_commit.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );

    let mut conflicting_edit = source.edit();
    conflicting_edit.set_shape_text(0, 0, "other").unwrap();
    let conflicting_commit = conflicting_edit.commit().unwrap();
    let mut conflicts = source.joined_edits(limits);
    conflicts
        .join(text_commit.patch().prepare("left", limits).unwrap())
        .unwrap();
    let error = conflicts
        .join(conflicting_commit.patch().prepare("right", limits).unwrap())
        .unwrap_err();
    assert!(matches!(error.failure(), SubEditJoinFailure::Overlap(_)));

    let foreign = Package::from_bytes(&package(&CONTENT.replace("Photo", "Changed"))).unwrap();
    let mut stale_joined = source.joined_edits(limits);
    stale_joined
        .join(text_commit.patch().prepare("stale", limits).unwrap())
        .unwrap();
    let mut foreign_edit = foreign.edit();
    let stale_error = match foreign_edit.join(stale_joined) {
        Ok(_) => panic!("stale joined work was accepted"),
        Err(error) => error,
    };
    assert!(matches!(stale_error.error(), Error::PatchConflict));
    assert!(matches!(
        text_commit.patch().then(name_commit.patch()),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn joined_staging_preserves_order_and_rejects_without_mutating_the_edit() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let limits = CompositionLimits::new(8, 8, 32, 16);

    let mut text_edit = source.edit();
    text_edit.set_shape_text(0, 0, "joined text").unwrap();
    let text_commit = text_edit.commit().unwrap();

    let mut ordinary = source.edit();
    ordinary.set_shape_name(0, 1, "ordinary name").unwrap();
    let mut joined = source.joined_edits(limits);
    joined
        .join(text_commit.patch().prepare("joined-text", limits).unwrap())
        .unwrap();
    ordinary.join(joined).unwrap();
    let commit = ordinary.commit().unwrap();
    assert_eq!(commit.diagnostics().change_count(), 2);
    assert_eq!(
        commit.package().pages()[0].shapes()[0].text(),
        "joined text"
    );
    assert_eq!(
        commit.package().pages()[0].shapes()[1].name(),
        Some("ordinary name")
    );

    let mut conflicting = source.edit();
    conflicting.set_shape_text(0, 0, "ordinary text").unwrap();
    let mut rejected_plan = source.joined_edits(limits);
    rejected_plan
        .join(text_commit.patch().prepare("conflict", limits).unwrap())
        .unwrap();
    let error = match conflicting.join(rejected_plan) {
        Ok(_) => panic!("overlapping joined work was accepted"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    let recovered = error.into_rejected();
    let ordinary_commit = conflicting.commit().unwrap();
    assert_eq!(
        ordinary_commit.package().pages()[0].shapes()[0].text(),
        "ordinary text"
    );
    let mut retry = source.edit();
    retry.join(recovered).unwrap();
    assert_eq!(
        retry.commit().unwrap().package().pages()[0].shapes()[0].text(),
        "joined text"
    );
}

#[test]
fn joined_staging_checks_payload_source_preconditions() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let foreign = Package::from_bytes(&package(&CONTENT.replace("Photo", "Changed"))).unwrap();
    let limits = CompositionLimits::new(4, 4, 8, 8);
    let mut foreign_edit = foreign.edit();
    foreign_edit.set_shape_text(0, 0, "foreign text").unwrap();
    let foreign_commit = foreign_edit.commit().unwrap();

    let plan = source.joined_edits(limits);
    let mismatched = edit::SubEdit::new(
        plan.lineage().clone(),
        limits,
        "mismatched-source",
        Vec::<String>::new(),
        vec!["page/0/shape/0#shape.text.set".to_owned()],
        foreign_commit.patch().durable().unwrap(),
    )
    .unwrap();
    let mut plan = plan;
    plan.join(mismatched).unwrap();
    let mut edit = source.edit();
    let error = match edit.join(plan) {
        Ok(_) => panic!("mismatched durable payload was accepted"),
        Err(error) => error,
    };
    assert!(matches!(error.error(), Error::PatchConflict));
    assert_eq!(
        edit.commit().unwrap().package().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn three_way_plan_does_not_mutate_until_conflicts_are_resolved() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let original = source.as_bytes().to_vec();
    let limits = CompositionLimits::new(8, 8, 32, 16);
    let mut left_edit = source.edit();
    left_edit.set_shape_name(0, 0, "Left").unwrap();
    let left_commit = left_edit.commit().unwrap();
    let mut right_edit = source.edit();
    right_edit.set_shape_name(0, 0, "Right").unwrap();
    let right_commit = right_edit.commit().unwrap();
    let mut left = source.joined_edits(limits);
    left.join(left_commit.patch().prepare("left", limits).unwrap())
        .unwrap();
    let mut right = source.joined_edits(limits);
    right
        .join(right_commit.patch().prepare("right", limits).unwrap())
        .unwrap();

    let plan = edit::MergePlan::new(left, right).unwrap();
    assert_eq!(source.as_bytes(), original);
    assert_eq!(plan.conflicts().len(), 1);
    let mut unresolved = plan.finish().unwrap_err();
    unresolved.resolve(MergeChoice::Right);
    let staged = unresolved.finish().unwrap();
    let mut staged_edit = source.edit();
    staged_edit.join(staged).unwrap();
    let applied = staged_edit.commit().unwrap().into_package();
    assert_eq!(applied.pages()[0].shapes()[0].name(), Some("Right"));
}

#[test]
fn history_enforces_step_and_weight_bounds_without_mutating_on_refusal() {
    let source = Package::from_bytes(&package(CONTENT)).unwrap();
    let mut first_edit = source.edit();
    first_edit.set_shape_name(0, 0, "First").unwrap();
    let first = first_edit.commit().unwrap().into_package();
    let mut second_edit = first.edit();
    second_edit.set_shape_name(0, 0, "Second").unwrap();
    let second = second_edit.commit().unwrap().into_package();
    let mut history = source.history(HistoryLimits::new(2, 10));
    history.record(first.clone(), 6).unwrap();
    history.record(second.clone(), 6).unwrap();
    assert_eq!(history.retained_weight(), 6);
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), first.as_bytes());
    assert!(!history.undo());
    assert!(history.record(second, 11).is_err());
    assert_eq!(history.current().as_bytes(), first.as_bytes());
}
