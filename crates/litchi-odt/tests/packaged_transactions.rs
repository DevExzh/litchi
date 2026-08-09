use litchi_odt::form::TextControl;
use litchi_odt::{
    Document, ScriptResourceKind, ScriptResourceSpec,
    core::{PackageWriter, Profile},
    mutable::MutableDocument,
    package::{
        embedded::{EmbeddedResource, EmbeddedResourceKind, EmbeddedResourceSource},
        forms::{AuthoredForm, AuthoredFormControl},
    },
    protection::Policy,
    rdf::{Object, Subject, Triple},
    transaction::{OperationResult, ParagraphSelector},
};

fn source() -> Document {
    let mut document = MutableDocument::new();
    document.add_paragraph("Unique paragraph").unwrap();
    Document::from_bytes(document.to_bytes().unwrap()).unwrap()
}

fn real_producer_source() -> Document {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer
        .add_file(
            "content.xml",
            include_bytes!("fixtures/libreoffice-master-document-content.xml"),
        )
        .unwrap();
    writer
        .add_file(
            "styles.xml",
            br#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:styles/></office:document-styles>"#,
        )
        .unwrap();
    writer
        .add_file(
            "meta.xml",
            br#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:meta/></office:document-meta>"#,
        )
        .unwrap();
    Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
}

#[test]
fn packaged_transaction_is_source_checked_reversible_and_exact_for_noop() {
    let source = source();
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();

    let no_op = snapshot.edit().commit().unwrap();
    assert_eq!(no_op.snapshot().as_bytes(), snapshot.as_bytes());

    let mut edit = snapshot.edit();
    edit.append_line_break(ParagraphSelector::exact_text("Unique paragraph"))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit.snapshot().document().unwrap().paragraphs().unwrap()[0]
            .text()
            .unwrap()
            .contains('\n')
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        snapshot.as_bytes()
    );
    assert!(commit.patch().apply(commit.snapshot()).is_err());
}

#[test]
fn packaged_transaction_rejects_ambiguous_text_selectors() {
    let mut document = MutableDocument::new();
    document.add_paragraph("duplicate").unwrap();
    document.add_paragraph("duplicate").unwrap();
    let source = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();

    assert!(
        snapshot
            .edit()
            .append_line_break(ParagraphSelector::exact_text("duplicate"))
            .is_err()
    );
}

#[test]
fn rdf_transaction_preserves_a_real_producer_package_and_is_reversible() {
    let source = real_producer_source();
    let before = source.snapshot().unwrap();
    let triple = Triple {
        subject: Subject::Iri("urn:litchi:document".to_string()),
        predicate: "urn:litchi:review#reviewed".to_string(),
        object: Object::Literal {
            value: "true".to_string(),
            datatype: None,
            language: None,
        },
    };

    let mut edit = source.edit().unwrap();
    edit.add_rdf_graph(Some("metadata/litchi.rdf"), &[triple.clone()])
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.snapshot().document().unwrap();
    let graphs = after.rdf_graphs().unwrap();
    assert_eq!(graphs.len(), 1);
    assert_eq!(graphs[0].path, "metadata/litchi.rdf");
    assert_eq!(graphs[0].triples, vec![triple]);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        before.as_bytes()
    );
}

#[test]
fn changed_transactions_refuse_signed_and_encrypted_envelopes() {
    let mut signed_writer = PackageWriter::new();
    signed_writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    signed_writer
        .add_file(
            "content.xml",
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>signed</text:p></office:text></office:body></office:document-content>"#,
        )
        .unwrap();
    signed_writer
        .add_file("META-INF/documentsignatures.xml", b"<signatures/>")
        .unwrap();
    let signed = Document::from_bytes(signed_writer.finish_to_bytes().unwrap()).unwrap();
    let signed_snapshot = signed.snapshot().unwrap();
    assert_eq!(
        signed_snapshot
            .edit()
            .commit()
            .unwrap()
            .snapshot()
            .as_bytes(),
        signed_snapshot.as_bytes()
    );
    let mut signed_edit = signed.edit().unwrap();
    signed_edit
        .append_line_break(ParagraphSelector::at(0))
        .unwrap();
    assert!(signed_edit.commit().is_err());

    let mut encrypted_writer = PackageWriter::new();
    encrypted_writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    encrypted_writer
        .set_encryption("secret", Profile::compatible())
        .unwrap();
    encrypted_writer
        .add_file(
            "content.xml",
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>encrypted</text:p></office:text></office:body></office:document-content>"#,
        )
        .unwrap();
    let encrypted_bytes = encrypted_writer.finish_to_bytes().unwrap();
    let encrypted = Document::from_bytes_with_password(encrypted_bytes, "secret").unwrap();
    let encrypted_snapshot = encrypted.snapshot().unwrap();
    assert_eq!(
        encrypted_snapshot
            .edit()
            .commit()
            .unwrap()
            .snapshot()
            .as_bytes(),
        encrypted_snapshot.as_bytes()
    );
    let mut encrypted_edit = encrypted.edit().unwrap();
    encrypted_edit
        .append_line_break(ParagraphSelector::at(0))
        .unwrap();
    assert!(encrypted_edit.commit().is_err());
}

#[test]
fn form_and_linked_resource_operations_commit_as_one_atomic_package() {
    let source = source();
    let before = source.snapshot().unwrap();
    let form = AuthoredForm::new("ReviewForm");
    let resource = EmbeddedResource {
        kind: EmbeddedResourceKind::Object,
        source: EmbeddedResourceSource::Linked {
            href: "https://example.invalid/inert-object".to_string(),
        },
        frame_name: Some("ReviewObject".to_string()),
        xml_id: Some("review-object".to_string()),
        class_id: None,
    };

    let mut edit = source.edit().unwrap();
    edit.add_form(0, &form)
        .unwrap()
        .add_embedded_resource(&resource)
        .unwrap();
    let commit = edit.commit().unwrap();
    let document = commit.snapshot().document().unwrap();
    assert_eq!(document.forms().unwrap().groups.len(), 1);
    assert_eq!(document.embedded_objects().unwrap().len(), 1);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        before.as_bytes()
    );
}

#[test]
fn residual_semantic_families_share_one_reversible_transaction() {
    let source = source();
    let triple = Triple {
        subject: Subject::Iri("urn:litchi:transaction".to_string()),
        predicate: "urn:litchi:review#status".to_string(),
        object: Object::Literal {
            value: "closed".to_string(),
            datatype: None,
            language: None,
        },
    };
    let form = AuthoredForm::new("Parent");
    let child = AuthoredForm::new("Child");
    let control = AuthoredFormControl::from(TextControl::text("Review", "review-control"));
    let script = ScriptResourceSpec {
        kind: ScriptResourceKind::Opaque,
        preferred_path: Some("Scripts/review.bin".to_string()),
        media_type: "application/octet-stream".to_string(),
        bytes: b"inert-not-executed".to_vec(),
    };

    let mut edit = source.edit().unwrap();
    edit.add_rdf_graph(Some("metadata/review.rdf"), &[triple.clone()])
        .unwrap()
        .add_rdf_triple("metadata/review.rdf", &triple)
        .unwrap()
        .set_protection(&Policy::default().with_read_only(Some(true)))
        .unwrap()
        .add_form(0, &form)
        .unwrap()
        .add_nested_form(0, &child)
        .unwrap()
        .add_form_control(0, &control)
        .unwrap()
        .add_script_resource(&script)
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(
        commit.results(),
        &[
            OperationResult::Path("metadata/review.rdf".to_string()),
            OperationResult::Index(1),
            OperationResult::Unit,
            OperationResult::Index(0),
            OperationResult::Index(1),
            OperationResult::Index(0),
            OperationResult::Path("Scripts/review.bin".to_string()),
        ]
    );
    let document = commit.snapshot().document().unwrap();
    assert_eq!(document.rdf_graphs().unwrap()[0].triples.len(), 2);
    assert_eq!(document.protection().unwrap().read_only, Some(true));
    assert_eq!(document.forms().unwrap().groups[0].forms.len(), 1);
    assert_eq!(document.script_resources().unwrap()[0].bytes, script.bytes);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.original_bytes()
    );
}
