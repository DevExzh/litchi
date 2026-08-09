#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::rdf::{Object, Subject, Triple};
use litchi_odp::{Builder, edit};

fn literal(value: &str) -> Triple {
    Triple {
        subject: Subject::Iri("https://example.test/deck".to_string()),
        predicate: "https://example.test/title".to_string(),
        object: Object::Literal {
            value: value.to_string(),
            datatype: None,
            language: Some("en".to_string()),
        },
    }
}

#[test]
fn slide_and_rdf_edits_publish_as_one_reversible_package_commit() {
    let mut builder = Builder::new();
    builder.add_slide_with_title("Source", "Body").unwrap();
    let source = edit::Snapshot::from_bytes(builder.build().unwrap()).unwrap();
    let source_bytes = source.bytes().to_vec();

    let mut transaction = source.transaction().unwrap();
    transaction
        .add("Second", "Slide and graph are atomic")
        .unwrap();
    let graph_path = transaction
        .add_rdf_graph(None, &[literal("First")])
        .unwrap();
    assert_eq!(graph_path, "Metadata/metadata_1.rdf");
    assert_eq!(
        transaction
            .add_rdf_triple(&graph_path, &literal("Second"))
            .unwrap(),
        1
    );
    transaction.move_rdf_triple(&graph_path, 1, 0).unwrap();
    transaction
        .replace_rdf_triple(&graph_path, 0, &literal("Published"))
        .unwrap();

    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(source.bytes(), source_bytes);
    assert_eq!(commit.snapshot().slides().len(), 2);

    let presentation = commit.snapshot().to_presentation().unwrap();
    let graphs = presentation.rdf_graphs().unwrap();
    assert_eq!(graphs.len(), 1);
    assert_eq!(graphs[0].path, graph_path);
    assert_eq!(graphs[0].triples.len(), 2);
    assert_eq!(graphs[0].triples[0], literal("Published"));
    assert_eq!(graphs[0].triples[1], literal("First"));

    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let rdf = String::from_utf8(package.get_file(&graph_path).unwrap()).unwrap();
    assert!(!rdf.contains('\n'));
    assert!(!rdf.contains("> <"));

    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    let restored = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(restored.bytes(), source.bytes());
}

#[test]
fn failed_checked_rdf_selector_leaves_an_exact_noop() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let graph_path = transaction
        .add_rdf_graph(Some("Metadata/deck.rdf"), &[literal("Kept")])
        .unwrap();
    let before = transaction.rdf_graphs().unwrap().to_vec();

    assert!(
        transaction
            .remove_rdf_triple(&graph_path, usize::MAX)
            .is_err()
    );
    assert_eq!(transaction.rdf_graphs().unwrap(), before);

    transaction.remove_rdf_graph(&graph_path).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
}

#[test]
fn signed_packages_are_refused_before_a_transaction_can_stage_changes() {
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file(
            "META-INF/documentsignatures.xml",
            br#"<dsig:document-signatures xmlns:dsig="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
        )
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    let Err(error) = source.transaction() else {
        panic!("signed package unexpectedly admitted an editing transaction");
    };
    assert!(error.to_string().contains("signed packages"));
}
