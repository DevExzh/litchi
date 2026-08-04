use litchi_docx::glossary::{
    self, Catalog, Category, Conformance, Entry, Gallery, Id, Insert, Kind, Name, Props, raw,
};
use litchi_ooxml::docx::Package;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{PackURI, part::BlobPart};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn entry(name: &str, relationship_id: Option<&str>) -> Entry {
    let relationship = relationship_id
        .map(|id| format!(r#" r:embed="{id}""#))
        .unwrap_or_default();
    let body = format!(
        r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r><w:drawing{relationship}/><w:t>{name}</w:t></w:r></w:p></w:docPartBody>"#
    )
    .into_bytes();
    let props = Props {
        category: Some(
            Category::new("General", Gallery::new("autoTxt").expect("gallery")).expect("category"),
        ),
        kinds: Kind::NORMAL,
        inserts: Insert::CONTENT,
        description: Some(format!("{name} description")),
        id: Some(Id::new("{12345678-1234-4ABC-8DEF-1234567890AB}").expect("canonical ID")),
        ..Props::new(Name::new(name).expect("name").with_decorated(false))
    };
    Entry::new(name, body)
        .and_then(|entry| entry.with_props(props))
        .expect("entry")
}

fn mark_signed(package: &mut Package) {
    package
        .edit_opc(|opc| {
            opc.try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
            )))?;
            opc.rels_mut().add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rSignature".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
}

#[test]
fn semantic_crud_is_name_first_and_numeric_fallback_is_checked() {
    let mut catalog = Catalog::new();
    catalog.add(entry("First", None)).unwrap();
    catalog.add(entry("Straße", None)).unwrap();

    assert_eq!(
        catalog.get("STRASSE").unwrap().unwrap().name(),
        Some("Straße")
    );
    assert_eq!(catalog.at(0).unwrap().name(), Some("First"));
    assert!(catalog.at(9).is_err());
    assert!(catalog.add(entry("strasse", None)).is_err());

    let previous = catalog.put(entry("FIRST", None)).unwrap().unwrap();
    assert_eq!(previous.name(), Some("First"));
    assert!(catalog.move_to("strasse", 0).unwrap());
    assert_eq!(catalog.at(0).unwrap().name(), Some("Straße"));
    assert_eq!(
        catalog.remove("STRASSE").unwrap().unwrap().name(),
        Some("Straße")
    );
    assert!(catalog.remove_at(9).is_err());
}

#[test]
fn host_facade_and_document_accessor_round_trip() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("glossary.docx");
    let mut catalog = Catalog::new();
    catalog.add(entry("One", None)).unwrap();
    catalog.add(entry("Two", None)).unwrap();

    let mut package = Package::new().unwrap();
    package
        .put_glossary(catalog, Conformance::Transitional)
        .unwrap();
    package.save(&path).unwrap();

    let reopened = Package::open(&path).unwrap();
    let (catalog, conformance) = reopened.glossary().unwrap().unwrap();
    assert_eq!(conformance, Conformance::Transitional);
    assert_eq!(catalog.len(), 2);
    let graph = reopened.glossary_graph().unwrap().unwrap();
    assert_eq!(graph.parts.len(), 4);
    for required in [
        "styles.xml",
        "settings.xml",
        "fontTable.xml",
        "webSettings.xml",
    ] {
        assert!(
            graph.parts.iter().any(|part| part.name.ends_with(required)),
            "missing glossary {required}"
        );
    }
    assert_eq!(
        reopened
            .document()
            .unwrap()
            .glossary()
            .unwrap()
            .unwrap()
            .0
            .len(),
        2
    );
}

#[test]
fn unrelated_document_mutation_preserves_producer_glossary() {
    let directory = tempfile::tempdir().unwrap();
    let output = directory.path().join("libreoffice-preserved.docx");
    let mut package = Package::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx"
    ))
    .unwrap();
    assert!(package.glossary().unwrap().unwrap().0.is_empty());
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("unrelated mutation");
    package.save(&output).unwrap();
    assert!(
        Package::open(&output)
            .unwrap()
            .glossary()
            .unwrap()
            .unwrap()
            .0
            .is_empty()
    );
}

#[test]
fn complete_raw_graph_round_trips_and_can_be_moved_out() {
    let mut catalog = Catalog::new();
    catalog.add(entry("Picture", Some("rIdImage"))).unwrap();
    let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
    graph.rels = vec![
        raw::Rel {
            id: "rIdStyles".into(),
            kind: rt::STYLES.into(),
            target: "styles.xml".into(),
            external: false,
        },
        raw::Rel {
            id: "rIdImage".into(),
            kind: rt::IMAGE.into(),
            target: "media/image1.png".into(),
            external: false,
        },
    ];
    graph.parts = vec![
        raw::Part::new(
            "/word/glossary/styles.xml",
            ct::WML_STYLES,
            format!(r#"<w:styles xmlns:w="{W}"/>"#).into_bytes(),
        )
        .unwrap(),
        raw::Part::new(
            "/word/glossary/media/image1.png",
            "image/png",
            vec![0x89, b'P', b'N', b'G'],
        )
        .unwrap(),
    ];

    let mut package = Package::new().unwrap();
    assert!(package.put_glossary_graph(&graph).unwrap());
    let loaded = package.glossary_graph().unwrap().unwrap();
    assert_eq!(loaded.parts.len(), 2);
    assert_eq!(
        loaded.catalog.get("picture").unwrap().unwrap().name(),
        Some("Picture")
    );

    let removed = package.take_glossary_graph().unwrap().unwrap();
    assert_eq!(removed.parts.len(), 2);
    assert!(package.glossary().unwrap().is_none());
    assert!(package.document().unwrap().styles().is_ok());
}

#[test]
fn invalid_graph_update_is_failure_atomic() {
    let mut package = Package::new().unwrap();
    let mut existing = Catalog::new();
    existing.add(entry("Keep", None)).unwrap();
    package
        .put_glossary(existing, Conformance::Transitional)
        .unwrap();

    for relationship in [
        raw::Rel {
            id: "rIdMissing".into(),
            kind: rt::IMAGE.into(),
            target: "media/missing.png".into(),
            external: false,
        },
        raw::Rel {
            id: "rIdExternal".into(),
            kind: rt::STYLES.into(),
            target: "https://example.invalid/styles.xml".into(),
            external: true,
        },
        raw::Rel {
            id: "rIdSpoof".into(),
            kind: "urn:spoof".into(),
            target: "media/image.png".into(),
            external: false,
        },
    ] {
        let mut invalid_catalog = Catalog::new();
        invalid_catalog.add(entry("Invalid", None)).unwrap();
        let mut invalid = raw::Graph::new(invalid_catalog, Conformance::Transitional);
        invalid.rels.push(relationship);
        assert!(package.put_glossary_graph(&invalid).is_err());
        assert!(
            package
                .glossary()
                .unwrap()
                .unwrap()
                .0
                .get("keep")
                .unwrap()
                .is_some()
        );
    }

    assert!(
        package
            .edit_opc(|opc| Ok::<_, litchi_ooxml::error::OoxmlError>(glossary::remove(opc)?))
            .unwrap()
    );
    assert!(
        !package
            .edit_opc(|opc| Ok::<_, litchi_ooxml::error::OoxmlError>(glossary::remove(opc)?))
            .unwrap()
    );
}

#[test]
fn host_graph_noop_preserves_signature_and_real_change_unsigns() {
    let mut package = Package::new().unwrap();
    let mut catalog = Catalog::new();
    catalog.add(entry("Keep", None)).unwrap();
    assert!(
        package
            .put_glossary(catalog, Conformance::Transitional)
            .unwrap()
    );
    mark_signed(&mut package);
    assert!(package.is_signed());
    let signature_origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    assert!(package.opc_package().contains_part(&signature_origin));

    let graph = package.glossary_graph().unwrap().unwrap();
    assert!(!package.put_glossary_graph(&graph).unwrap());
    assert!(package.is_signed());
    assert!(package.opc_package().contains_part(&signature_origin));

    let mut changed = package.glossary_graph().unwrap().unwrap();
    changed.catalog.add(entry("Changed", None)).unwrap();
    assert!(package.put_glossary_graph(&changed).unwrap());
    assert!(!package.is_signed());
    assert!(
        package
            .glossary()
            .unwrap()
            .unwrap()
            .0
            .get("changed")
            .unwrap()
            .is_some()
    );
}

#[test]
fn host_graph_failure_preserves_existing_graph_and_signature() {
    let mut package = Package::new().unwrap();
    let mut catalog = Catalog::new();
    catalog.add(entry("Keep", None)).unwrap();
    package
        .put_glossary(catalog, Conformance::Transitional)
        .unwrap();
    mark_signed(&mut package);

    let mut invalid = package.glossary_graph().unwrap().unwrap();
    invalid.rels.push(raw::Rel {
        id: "rIdMissing".to_owned(),
        kind: rt::IMAGE.to_owned(),
        target: "media/missing.png".to_owned(),
        external: false,
    });
    assert!(package.put_glossary_graph(&invalid).is_err());
    assert!(package.is_signed());
    assert!(
        package
            .glossary()
            .unwrap()
            .unwrap()
            .0
            .get("keep")
            .unwrap()
            .is_some()
    );
}
