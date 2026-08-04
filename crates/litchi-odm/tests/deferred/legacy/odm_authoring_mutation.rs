use litchi_odf::{
    MasterDocument, MasterDocumentBuilder, MasterDocumentElement, MasterSection, MasterSubdocument,
    MutableMasterDocument, OdfDocumentSigner, OdfEncryptionProfile, OdfSignatureAlgorithm,
    OdfSignatureValidity, OwnedPackage, PackageWriter, Section, SectionDisplay,
    TableOfContentsSource, TextIndex, TextIndexBody,
};

const ODM: &str = "application/vnd.oasis.opendocument.text-master";
const OTM: &str = "application/vnd.oasis.opendocument.text-master-template";
const RSA_KEY: &[u8] = include_bytes!("fixtures/signatures/rsa-key.pk8");
const RSA_CERT: &[u8] = include_bytes!("fixtures/signatures/rsa-cert.der");

fn local_section(name: &str, text: &str) -> Section {
    Section {
        name: name.to_string(),
        style: None,
        protected: false,
        xml_id: None,
        protection_key: None,
        protection_key_digest_algorithm: None,
        display: SectionDisplay::Visible,
        condition: None,
        source: None,
        dde_source: None,
        content: text.to_string(),
    }
}

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(ODM).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer
        .add_file("styles.xml", b"<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>")
        .unwrap();
    writer
        .add_file("settings.xml", b"<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>")
        .unwrap();
    writer
        .add_file_with_media_type("custom/cache.bin", b"cached", "application/octet-stream")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn builder_authors_odm_and_otm_with_nested_mixed_content() {
    let mut nested = MasterSection::new(local_section("Outer & 文档", "Local <cached>")).unwrap();
    nested
        .push(MasterDocumentElement::Paragraph(
            "Nested & text".to_string(),
        ))
        .unwrap();
    nested
        .push(MasterDocumentElement::Subdocument(
            MasterSubdocument::new("Nested link", "Chapters/章.odt").unwrap(),
        ))
        .unwrap();
    let index = TextIndex::table_of_contents(
        "Master Contents",
        TableOfContentsSource::new(),
        TextIndexBody::new(),
    )
    .unwrap();

    for template in [false, true] {
        let mut builder = if template {
            MasterDocumentBuilder::template()
        } else {
            MasterDocumentBuilder::new()
        };
        builder
            .add_paragraph("Résumé & overview")
            .unwrap()
            .add_section(nested.clone())
            .unwrap()
            .add_index(index.clone())
            .unwrap()
            .add_subdocument(MasterSubdocument::new("Appendix", "../appendix.odt").unwrap())
            .unwrap();
        builder
            .set_settings_xml(Some("<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>".to_string()))
            .unwrap();
        let bytes = builder.build().unwrap();
        let master = MasterDocument::from_bytes(bytes).unwrap();
        assert_eq!(master.mimetype(), if template { OTM } else { ODM });
        assert_eq!(master.subdocuments().len(), 2);
        assert!(master.text().unwrap().contains("Résumé & overview"));
        assert!(master.text().unwrap().contains("Local <cached>"));
    }
}

#[test]
fn libreoffice_fixture_mutates_losslessly_and_preserves_auxiliary_parts() {
    let fixture = include_str!("fixtures/libreoffice-master-document-content.xml");
    let master = MasterDocument::from_bytes(package(fixture)).unwrap();
    assert_eq!(master.subdocuments().len(), 3);
    let mut mutable = MutableMasterDocument::from_document(master).unwrap();
    let original_extension = r#"<loext:opaque loext:value="keep &amp; exact"/>"#;

    mutable
        .reorder_subdocuments(&["Chapter B".to_string(), "Chapter A".to_string()])
        .unwrap();
    let mut updated = MasterSubdocument::new("Chapter B", "Chapters/new-b.odt").unwrap();
    updated
        .set_source_section_name(Some("Body & Notes".to_string()))
        .unwrap();
    mutable.update_subdocument("Chapter B", &updated).unwrap();
    let removed = mutable.remove_subdocument("Chapter A").unwrap();
    assert_eq!(removed.href(), Some("Chapters/a.odt"));
    mutable.add_paragraph("尾声 & end").unwrap();

    let bytes = mutable.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.get_file("custom/cache.bin").unwrap(), b"cached");
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains(original_extension));
    assert!(content.contains("Chapters/new-b.odt"));
    assert!(content.contains("Body &amp; Notes"));
    assert!(content.contains("尾声 &amp; end"));
    let reopened = MasterDocument::from_bytes(bytes).unwrap();
    assert_eq!(reopened.subdocuments().len(), 2);
    assert_eq!(reopened.subdocuments()[0].section_name(), "Chapter B");
    assert_eq!(reopened.subdocuments()[1].section_name(), "Nested C");
}

#[test]
fn local_section_and_index_mutation_round_trip() {
    let fixture = include_str!("fixtures/libreoffice-master-document-content.xml");
    let mut mutable = MutableMasterDocument::from_bytes(package(fixture)).unwrap();
    mutable
        .update_section("Local", &local_section("Renamed Local", "更新 & local"))
        .unwrap();
    let replacement = TextIndex::table_of_contents(
        "New Contents",
        TableOfContentsSource::new(),
        TextIndexBody::new(),
    )
    .unwrap();
    mutable.update_index("Contents", &replacement).unwrap();
    mutable.move_body_element(0, 4).unwrap();
    let reopened = MasterDocument::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert!(reopened.text().unwrap().contains("更新 & local"));
    assert!(reopened.text().unwrap().contains("Master introduction"));
}

#[test]
fn invalid_links_duplicates_depth_and_reorders_roll_back() {
    assert!(MasterSubdocument::new("bad", "https://example.invalid/a b.odt").is_err());

    let fixture = include_str!("fixtures/libreoffice-master-document-content.xml");
    let mut mutable = MutableMasterDocument::from_bytes(package(fixture)).unwrap();
    let before = mutable.content_xml().to_string();
    assert!(
        mutable
            .add_subdocument(&MasterSubdocument::new("Chapter B", "other.odt").unwrap())
            .is_err()
    );
    assert_eq!(mutable.content_xml(), before);
    assert!(
        mutable
            .reorder_subdocuments(&["Chapter A".to_string(), "missing".to_string()])
            .is_err()
    );
    assert_eq!(mutable.content_xml(), before);
    assert!(mutable.move_body_element(99, 0).is_err());
    assert_eq!(mutable.content_xml(), before);

    let mut builder = MasterDocumentBuilder::new();
    assert!(
        builder
            .add_auxiliary_file("../escape.bin", vec![], None)
            .is_err()
    );
    builder.add_paragraph("still valid").unwrap();
    assert!(builder.build().is_ok());

    let mut element = MasterDocumentElement::Paragraph("deep".to_string());
    for index in 0..129 {
        let mut section = MasterSection::new(local_section(&format!("s{index}"), "")).unwrap();
        section.push(element).unwrap();
        element = MasterDocumentElement::Section(section);
    }
    let mut builder = MasterDocumentBuilder::new();
    builder.push(element).unwrap();
    assert!(builder.build().is_err());
}

#[test]
fn encrypted_and_signed_master_packages_compose() {
    let signer = OdfDocumentSigner::from_pkcs8_der(
        OdfSignatureAlgorithm::RsaSha256,
        RSA_KEY,
        vec![RSA_CERT.to_vec()],
        "2026-07-19T12:00:00Z",
    )
    .unwrap();
    let mut builder = MasterDocumentBuilder::new();
    builder
        .add_paragraph("secured master")
        .unwrap()
        .add_subdocument(MasterSubdocument::new("Secure", "secure.odt").unwrap())
        .unwrap();
    builder
        .set_encryption("master-password", OdfEncryptionProfile::compatible())
        .unwrap();
    builder.set_document_signer(signer);
    let bytes = builder.build().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert!(
        package
            .verify_document_signatures()
            .unwrap()
            .iter()
            .all(|result| result.validity == OdfSignatureValidity::Valid)
    );
    let opened = MasterDocument::from_bytes_with_password(bytes, "master-password").unwrap();
    assert_eq!(opened.subdocuments().len(), 1);
}

#[test]
fn odm_authoring_contains_no_unsafe_code() {
    for source in [
        include_str!("../../../src/migration/legacy/document.rs"),
        include_str!("../../../src/migration/legacy/builder.rs"),
        include_str!("../../../src/migration/legacy/mutable.rs"),
    ] {
        assert!(!source.contains("unsafe {"));
    }
}
