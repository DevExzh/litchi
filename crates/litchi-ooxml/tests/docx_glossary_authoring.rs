use litchi_ooxml::docx::{
    DocPartCategory, DocPartGallery, DocPartName, DocPartProperties, DocPartType,
    GlossaryAuxiliaryPart, GlossaryDocument, GlossaryEntry, GlossaryPackage,
    GlossaryRelationship, InsertionBehavior, Package,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn entry(name: &str, relationship_id: Option<&str>) -> GlossaryEntry {
    let relationship = relationship_id
        .map(|id| format!(r#" r:embed="{id}""#))
        .unwrap_or_default();
    GlossaryEntry {
        properties: Some(DocPartProperties {
            name: Some(DocPartName {
                value: name.to_string(),
                decorated: Some(false),
            }),
            category: Some(DocPartCategory {
                name: "General".to_string(),
                gallery: DocPartGallery::parse("autoTxt").unwrap(),
            }),
            types: vec![DocPartType::Normal],
            behaviors: vec![InsertionBehavior::Content],
            description: Some(format!("{name} description")),
            guid: Some("{12345678-1234-4ABC-8DEF-1234567890AB}".to_string()),
            ..DocPartProperties::default()
        }),
        body_xml: Some(
            format!(
                r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r><w:drawing{relationship}/><w:t>{name}</w:t></w:r></w:p></w:docPartBody>"#
            )
            .into_bytes(),
        ),
    }
}

#[test]
fn typed_entries_preserve_order_and_reject_invalid_updates_atomically() {
    let mut glossary = GlossaryDocument::default();
    glossary.add_entry(entry("First", None)).unwrap();
    glossary.add_entry(entry("Second", None)).unwrap();
    glossary.move_entry(1, 0).unwrap();
    assert_eq!(glossary.find_entry("second").unwrap().0, 0);

    let before = glossary.clone();
    assert!(glossary.add_entry(entry("SECOND", None)).is_err());
    assert_eq!(glossary, before);
    assert!(
        glossary
            .update_entry(0, |entry| {
                entry.properties.as_mut().unwrap().guid = Some("not-a-guid".into());
                Ok(())
            })
            .is_err()
    );
    assert_eq!(glossary, before);
}

#[test]
fn transitional_package_roundtrip_and_document_accessor() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("glossary.docx");
    let mut package = Package::new().unwrap();
    package.add_glossary_entry(entry("One", None)).unwrap();
    package.add_glossary_entry(entry("Two", None)).unwrap();
    package.save(&path).unwrap();

    let reopened = Package::open(&path).unwrap();
    let glossary = reopened.glossary_document().unwrap().unwrap();
    assert_eq!(glossary.entries().len(), 2);
    assert_eq!(
        reopened
            .document()
            .unwrap()
            .glossary_document()
            .unwrap()
            .unwrap()
            .entries()
            .len(),
        2
    );
}

#[test]
fn ordinary_mutation_preserves_libreoffice_test_glossary() {
    let directory = tempfile::tempdir().unwrap();
    let output = directory.path().join("libreoffice-preserved.docx");
    let mut package = Package::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx"
    ))
    .unwrap();
    assert!(package.glossary_document().unwrap().unwrap().entries().is_empty());
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("unrelated mutation");
    package.save(&output).unwrap();

    let reopened = Package::open(&output).unwrap();
    assert!(reopened.glossary_document().unwrap().unwrap().entries().is_empty());
}

#[test]
fn strict_auxiliary_styles_and_media_roundtrip_then_remove_cleanly() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("strict-glossary.docx");
    let removed_path = directory.path().join("removed.docx");
    let mut package = Package::new().unwrap();
    let mut glossary = GlossaryPackage::new(
        GlossaryDocument {
            background_xml: None,
            entries: vec![entry("Picture", Some("rIdImage"))],
        },
        true,
    );
    glossary.relationships = vec![
        GlossaryRelationship {
            id: "rIdStyles".into(),
            relationship_type: rt::STYLES.into(),
            target: "styles.xml".into(),
            external: false,
        },
        GlossaryRelationship {
            id: "rIdImage".into(),
            relationship_type: rt::IMAGE.into(),
            target: "media/image1.png".into(),
            external: false,
        },
    ];
    glossary.auxiliary_parts = vec![
        GlossaryAuxiliaryPart {
            part_name: "/word/glossary/styles.xml".into(),
            content_type: ct::WML_STYLES.into(),
            data: format!(r#"<w:styles xmlns:w="{W}"/>"#).into_bytes(),
            relationships: Vec::new(),
        },
        GlossaryAuxiliaryPart {
            part_name: "/word/glossary/media/image1.png".into(),
            content_type: "image/png".into(),
            data: vec![0x89, b'P', b'N', b'G'],
            relationships: Vec::new(),
        },
    ];
    package.set_glossary_package(glossary).unwrap();
    package
        .update_glossary_document(|document| {
            document
                .add_entry(entry("Second", None))
                .map(|_| ())
                .map_err(|error| litchi_ooxml::OoxmlError::InvalidFormat(error.to_string()))
        })
        .unwrap();
    package.save(&path).unwrap();

    let mut reopened = Package::open(&path).unwrap();
    let graph = reopened.glossary_package().unwrap().unwrap();
    assert!(graph.strict);
    assert_eq!(graph.auxiliary_parts.len(), 2);
    assert_eq!(graph.relationships.len(), 2);
    assert!(reopened.remove_glossary_document().unwrap().is_some());
    assert!(reopened.document().unwrap().styles().is_ok());
    reopened.save(&removed_path).unwrap();
    assert!(
        Package::open(&removed_path)
            .unwrap()
            .glossary_document()
            .unwrap()
            .is_none()
    );
}

#[test]
fn invalid_graph_updates_roll_back_without_losing_existing_glossary() {
    let mut package = Package::new().unwrap();
    package.add_glossary_entry(entry("Keep", None)).unwrap();

    for relationship in [
        GlossaryRelationship {
            id: "rIdMissing".into(),
            relationship_type: rt::IMAGE.into(),
            target: "media/missing.png".into(),
            external: false,
        },
        GlossaryRelationship {
            id: "rIdExternal".into(),
            relationship_type: rt::IMAGE.into(),
            target: "https://example.invalid/image.png".into(),
            external: true,
        },
        GlossaryRelationship {
            id: "rIdSpoof".into(),
            relationship_type: "urn:spoof".into(),
            target: "media/image.png".into(),
            external: false,
        },
    ] {
        let mut invalid = GlossaryPackage::new(
            GlossaryDocument {
                background_xml: None,
                entries: vec![entry("Invalid", None)],
            },
            false,
        );
        invalid.relationships.push(relationship);
        assert!(package.set_glossary_package(invalid).is_err());
        assert!(package
            .glossary_document()
            .unwrap()
            .unwrap()
            .find_entry("Keep")
            .is_some());
    }
}
