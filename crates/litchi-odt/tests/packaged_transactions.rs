use litchi_odf_common::package::raw_identical_members;
use litchi_odt::form::TextControl;
use litchi_odt::{
    Document, ScriptResourceKind, ScriptResourceSpec,
    core::{OwnedPackage as CoreOwnedPackage, PackageWriter, Profile},
    elements::field::DynamicTextField,
    mutable::MutableDocument,
    note::{Note, NoteClass},
    odc::{ChartClass, Definition},
    package::{
        embedded::{EmbeddedResource, EmbeddedResourceKind, EmbeddedResourceSource},
        forms::{AuthoredForm, AuthoredFormControl},
    },
    protection::Policy,
    rdf::{Object, Subject, Triple},
    ruby_family::{
        Alignment, Annotation as RubyAnnotation, Base as RubyBase, Properties as RubyProperties,
        Style as RubyStyle,
    },
    transaction::{
        EnvelopeKind, HistoryLimits, MergeChoice, MergePlan, OperationResult, ParagraphSelector,
        Position, SubEditJoinFailure, TransferDependencyKind,
    },
};
mod support;

fn source() -> Document {
    let mut document = MutableDocument::new();
    document.add_paragraph("Unique paragraph").unwrap();
    Document::from_bytes(document.to_bytes().unwrap()).unwrap()
}

#[test]
fn paragraph_replacement_raw_preserves_unchanged_media_and_metadata() {
    const MEDIA_PATH: &str = "Pictures/opaque.bin";
    let mut base = MutableDocument::new();
    base.add_paragraph("Before").unwrap();
    let base = CoreOwnedPackage::from_bytes(base.to_bytes().unwrap()).unwrap();
    let package = base.package().unwrap();
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    for path in package.files().unwrap() {
        if matches!(path.as_str(), "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        writer
            .add_file_with_media_type(
                &path,
                &package.get_file(&path).unwrap(),
                package.manifest().get_media_type(&path).unwrap_or_default(),
            )
            .unwrap();
    }
    writer.add_manifest_directory("Pictures/", "").unwrap();
    writer
        .add_file_with_media_type(
            MEDIA_PATH,
            &vec![0x5a; 1024 * 1024],
            "application/octet-stream",
        )
        .unwrap();
    let source_bytes = writer.finish_to_bytes().unwrap();
    let source = Document::from_bytes(source_bytes.clone()).unwrap();
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();

    let mut edit = snapshot.edit();
    edit.replace_paragraph(Position::new(0), "After").unwrap();
    let commit = edit.commit().unwrap();
    let identical = raw_identical_members(&source_bytes, commit.snapshot().as_bytes()).unwrap();

    assert!(!identical.contains("content.xml"));
    for path in [
        "mimetype",
        "styles.xml",
        "meta.xml",
        "META-INF/manifest.xml",
        MEDIA_PATH,
    ] {
        assert!(identical.contains(path), "{path}");
    }
    let reopened = commit.snapshot().document().unwrap();
    assert_eq!(reopened.paragraphs().unwrap()[0].text().unwrap(), "After");
    assert_eq!(
        reopened.get_file(MEDIA_PATH).unwrap(),
        vec![0x5a; 1024 * 1024]
    );
    assert_eq!(
        commit.patch().apply(&snapshot).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source_bytes
    );
}

#[test]
fn paragraph_replacement_above_common_raw_limit_uses_existing_rebuild_path() {
    const PARAGRAPH_COUNT: usize = 16_120;
    const TARGET: usize = PARAGRAPH_COUNT / 2;
    let text = "x".repeat(1024);
    let mut document = MutableDocument::new();
    for _ in 0..PARAGRAPH_COUNT {
        document.add_paragraph(&text).unwrap();
    }
    let source = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert!(
        source.get_file("content.xml").unwrap().len()
            > litchi_odf_common::package::MAX_CONTENT_REPLACEMENT_BYTES
    );
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();

    let mut edit = snapshot.edit();
    edit.replace_paragraph(Position::new(TARGET), "fallback replacement")
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = commit.snapshot().document().unwrap();
    let paragraphs = reopened.paragraphs().unwrap();

    assert_eq!(paragraphs.len(), PARAGRAPH_COUNT);
    assert_eq!(paragraphs[TARGET].text().unwrap(), "fallback replacement");
    assert_eq!(paragraphs[TARGET - 1].text().unwrap(), text);
    assert_eq!(paragraphs[TARGET + 1].text().unwrap(), text);
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

fn assert_all_package_xml_is_compact(bytes: &[u8]) {
    let package = litchi_odt::core::package::OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    let archive = package.package().unwrap();
    for path in archive.files().unwrap() {
        if path.ends_with(".xml") || path.ends_with(".rdf") {
            litchi_odf_common::compact_xml::validate(&archive.get_file(&path).unwrap()).unwrap();
        }
    }
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
fn checked_positions_are_resolved_against_the_source_snapshot() {
    let source = source();
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();
    let position = Position::new(0);

    let mut edit = snapshot.edit();
    edit.append_line_break(ParagraphSelector::position(position))
        .unwrap();
    assert!(edit.commit().is_ok());

    assert!(
        snapshot
            .edit()
            .append_line_break(ParagraphSelector::position(Position::new(1)))
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
    edit.add_rdf_graph(Some("metadata/litchi.rdf"), std::slice::from_ref(&triple))
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
        signed_snapshot.envelope_kind().unwrap(),
        EnvelopeKind::Signed
    );
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
        encrypted_snapshot.envelope_kind().unwrap(),
        EnvelopeKind::Encrypted
    );
    let encrypted_noop = encrypted_snapshot.edit().commit().unwrap();
    assert_eq!(
        encrypted_noop.snapshot().as_bytes(),
        encrypted_snapshot.as_bytes()
    );
    assert!(encrypted_noop.patch().durable().is_err());
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
    edit.add_rdf_graph(Some("metadata/review.rdf"), std::slice::from_ref(&triple))
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
    let durable = commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    for operation in ["form.add", "form.add_nested", "form.control.add"] {
        assert!(wire_text.contains(operation), "missing {operation}");
    }
    assert_eq!(
        durable
            .apply(&source.snapshot().unwrap())
            .unwrap()
            .as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_all_package_xml_is_compact(commit.snapshot().as_bytes());
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

#[test]
fn form_replacement_reorder_and_removal_are_durable_and_transferable() {
    let source = source().snapshot().unwrap();
    let first = AuthoredForm::new("First");
    let second = AuthoredForm::new("Second");
    let first_control = AuthoredFormControl::from(TextControl::text("One", "one"));
    let second_control = AuthoredFormControl::from(TextControl::text("Two", "two"));
    let mut setup = source.edit();
    setup
        .add_form(0, &first)
        .unwrap()
        .add_form(0, &second)
        .unwrap()
        .add_form_control(0, &first_control)
        .unwrap()
        .add_form_control(0, &second_control)
        .unwrap();
    let setup = setup.commit().unwrap().into_snapshot();

    let replacement_form = AuthoredForm::new("Replacement");
    let replacement_control =
        AuthoredFormControl::from(TextControl::text("Replacement", "replacement"));
    let mut edit = setup.edit();
    edit.replace_form_control(0, &replacement_control)
        .unwrap()
        .move_form_control_to(Position::new(0), Position::new(1))
        .unwrap()
        .remove_form_control_at(Position::new(1))
        .unwrap()
        .replace_form(1, &replacement_form)
        .unwrap()
        .move_form_to(Position::new(0), Position::new(1))
        .unwrap()
        .remove_form_at(Position::new(1))
        .unwrap();
    let committed = edit.commit().unwrap();
    let durable = committed.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    for operation in [
        "form.control.replace",
        "form.control.move",
        "form.control.remove",
        "form.replace",
        "form.move",
        "form.remove",
    ] {
        assert!(wire_text.contains(operation), "missing {operation}");
    }
    assert_eq!(
        durable.apply(&setup).unwrap().as_bytes(),
        committed.snapshot().as_bytes()
    );
    assert_eq!(
        committed
            .patch()
            .plan_transfer(&setup)
            .unwrap()
            .commit()
            .unwrap()
            .snapshot()
            .as_bytes(),
        committed.snapshot().as_bytes()
    );
}

#[test]
fn durable_patch_round_trips_deterministically_and_remains_exactly_reversible() {
    let source = source();
    let snapshot = source.snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.append_line_break(ParagraphSelector::position(Position::new(0)))
        .unwrap();
    let commit = edit.commit().unwrap();

    let durable = commit.patch().durable().unwrap();
    let first_json = durable.to_deterministic_json().unwrap();
    let second_json = durable.to_deterministic_json().unwrap();
    assert_eq!(first_json, second_json);

    let decoded =
        litchi_odt::transaction::DurablePatch::from_deterministic_json(&first_json).unwrap();
    let applied = decoded.apply(&snapshot).unwrap();
    assert_eq!(applied.as_bytes(), commit.snapshot().as_bytes());
    assert_all_package_xml_is_compact(applied.as_bytes());
    let restored = decoded.inverse().apply(&applied).unwrap();
    assert_eq!(restored.as_bytes(), snapshot.as_bytes());
    assert!(decoded.apply(&applied).is_err());
}

#[test]
fn semantic_durable_patch_covers_styles_fields_revisions_rdf_protection_and_script_blobs() {
    let source = source();
    let snapshot = source.snapshot().unwrap();
    let ruby = RubyStyle::new(
        "RubyReview",
        Some(RubyProperties {
            position: None,
            alignment: Some(Alignment::Center),
        }),
    )
    .unwrap();
    let field = DynamicTextField::DdeConnection {
        connection_name: "review-cache".to_string(),
        display_text: " \t ".to_string(),
    };
    let triple = Triple {
        subject: Subject::Iri("urn:litchi:durable".to_string()),
        predicate: "urn:litchi:review#status".to_string(),
        object: Object::Literal {
            value: "ready".to_string(),
            datatype: None,
            language: Some("en".to_string()),
        },
    };
    let script = ScriptResourceSpec {
        kind: ScriptResourceKind::Opaque,
        preferred_path: Some("Scripts/durable.bin".to_string()),
        media_type: "application/octet-stream".to_string(),
        bytes: b"durable inert payload".to_vec(),
    };

    let mut edit = snapshot.edit();
    edit.set_ruby_style(&ruby)
        .unwrap()
        .insert_dynamic_text_field(Position::new(0), &field)
        .unwrap()
        .set_tracked_change_policy(
            Some(true),
            Some("c2FsdGVkLWtleQ=="),
            Some("urn:litchi:digest:test"),
        )
        .unwrap()
        .add_rdf_graph(Some("metadata/durable.rdf"), std::slice::from_ref(&triple))
        .unwrap()
        .set_protection(&Policy::default().with_read_only(Some(true)))
        .unwrap()
        .add_script_resource(&script)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    for operation in [
        "style.ruby.set",
        "field.dynamic.insert",
        "revision.policy.set",
        "rdf.graph.add",
        "protection.set",
        "resource.script.add",
    ] {
        assert!(wire_text.contains(operation), "missing {operation}");
    }
    assert!(!wire_text.contains("durable inert payload"));

    let decoded = litchi_odt::transaction::DurablePatch::from_deterministic_json(&wire).unwrap();
    let applied = decoded.apply(&snapshot).unwrap();
    assert_eq!(applied.as_bytes(), commit.snapshot().as_bytes());
    let reopened = applied.document().unwrap();
    assert_eq!(reopened.ruby_styles().unwrap().styles, vec![ruby]);
    assert_eq!(reopened.dynamic_text_fields().unwrap(), vec![field]);
    assert_eq!(
        reopened.tracked_changes().unwrap().track_changes,
        Some(true)
    );
    assert_eq!(reopened.rdf_graphs().unwrap()[0].triples, vec![triple]);
    assert_eq!(reopened.protection().unwrap().read_only, Some(true));
    assert_eq!(reopened.script_resources().unwrap()[0].bytes, script.bytes);

    let package =
        litchi_odt::core::package::OwnedPackage::from_bytes(applied.as_bytes().to_vec()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("&#32;&#9;&#32;"));
}

#[test]
fn rich_note_and_ruby_edits_use_typed_durable_fragments() {
    let snapshot = source().snapshot().unwrap();
    let note = Note::new(NoteClass::Footnote, "1", "Durable footnote").unwrap();
    let ruby = RubyAnnotation::new(None, RubyBase::from_text("漢").unwrap(), "kan", None).unwrap();
    let mut edit = snapshot.edit();
    edit.insert_note(Position::new(0), &note)
        .unwrap()
        .insert_ruby_annotation(Position::new(0), &ruby)
        .unwrap();
    let committed = edit.commit().unwrap();
    let durable = committed.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    assert!(wire_text.contains("note.insert"));
    assert!(wire_text.contains("ruby.annotation.insert"));
    assert!(!wire_text.contains("document.replace"));

    let decoded = litchi_odt::transaction::DurablePatch::from_deterministic_json(&wire).unwrap();
    let inserted = decoded.apply(&snapshot).unwrap();
    assert_eq!(inserted.as_bytes(), committed.snapshot().as_bytes());
    let reopened = inserted.document().unwrap();
    assert_eq!(reopened.notes().unwrap(), vec![note]);
    assert_eq!(reopened.ruby_annotations().unwrap().annotations, vec![ruby]);

    let replacement_note = Note::new(NoteClass::Endnote, "A", "Replacement").unwrap();
    let replacement_ruby =
        RubyAnnotation::new(None, RubyBase::from_text("字").unwrap(), "ji", None).unwrap();
    let mut replacement = inserted.edit();
    replacement
        .replace_note(Position::new(0), &replacement_note)
        .unwrap()
        .replace_ruby_annotation(Position::new(0), &replacement_ruby)
        .unwrap();
    let replacement = replacement.commit().unwrap();
    let replayed = replacement
        .patch()
        .durable()
        .unwrap()
        .apply(&inserted)
        .unwrap();
    assert_eq!(replayed.as_bytes(), replacement.snapshot().as_bytes());

    let mut removal = replayed.edit();
    removal
        .remove_note(Position::new(0))
        .unwrap()
        .remove_ruby_annotation(Position::new(0))
        .unwrap();
    let removal = removal.commit().unwrap();
    let removed = removal
        .patch()
        .durable()
        .unwrap()
        .apply(&replayed)
        .unwrap()
        .document()
        .unwrap();
    assert!(removed.notes().unwrap().is_empty());
    assert!(removed.ruby_annotations().unwrap().annotations.is_empty());
}

#[test]
fn chart_and_resource_payloads_replay_and_transfer_with_explicit_dependencies() {
    let snapshot = source().snapshot().unwrap();
    let chart = Definition::new(ChartClass::line());
    let resource = EmbeddedResource {
        kind: EmbeddedResourceKind::Image,
        source: EmbeddedResourceSource::InlineBinary {
            bytes: b"durable-resource-payload".to_vec(),
            media_type: Some("image/png".to_string()),
        },
        frame_name: Some("Durable Resource".to_string()),
        xml_id: Some("durable-resource".to_string()),
        class_id: None,
    };
    let mut edit = snapshot.edit();
    edit.add_embedded_chart(&chart)
        .unwrap()
        .add_embedded_resource(&resource)
        .unwrap();
    let committed = edit.commit().unwrap();
    let durable = committed.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    assert!(wire_text.contains("chart.add"));
    assert!(wire_text.contains("resource.embedded.add"));
    assert!(!wire_text.contains("durable-resource-payload"));

    let decoded = litchi_odt::transaction::DurablePatch::from_deterministic_json(&wire).unwrap();
    let applied = decoded.apply(&snapshot).unwrap();
    assert_eq!(applied.as_bytes(), committed.snapshot().as_bytes());
    applied.document().unwrap().embedded_chart(0).unwrap();

    let destination = source().snapshot().unwrap();
    let transfer = committed.patch().plan_transfer(&destination).unwrap();
    assert!(transfer.dependencies().iter().any(|dependency| {
        dependency.kind() == TransferDependencyKind::ResourcePayload && dependency.is_satisfied()
    }));
    let transferred = transfer.commit().unwrap();
    transferred
        .snapshot()
        .document()
        .unwrap()
        .embedded_chart(0)
        .unwrap();

    let mut styled = Definition::new(ChartClass::line());
    styled.style_name = Some("MissingChartStyle".to_string());
    let mut styled_edit = snapshot.edit();
    styled_edit.add_embedded_chart(&styled).unwrap();
    let styled_commit = styled_edit.commit().unwrap();
    let styled_transfer = styled_commit.patch().plan_transfer(&destination).unwrap();
    assert!(styled_transfer.dependencies().iter().any(|dependency| {
        dependency.kind() == TransferDependencyKind::ChartStyle
            && dependency.key() == "MissingChartStyle"
            && dependency.is_satisfied()
    }));
    styled_transfer.commit().unwrap();
}

#[test]
fn chart_and_resource_replacement_removal_and_moves_use_semantic_wire() {
    let source = source().snapshot().unwrap();
    let linked = |name: &str| EmbeddedResource {
        kind: EmbeddedResourceKind::Object,
        source: EmbeddedResourceSource::Linked {
            href: format!("https://example.invalid/{name}"),
        },
        frame_name: Some(name.to_string()),
        xml_id: None,
        class_id: None,
    };
    let image = |name: &str, bytes: &[u8]| EmbeddedResource {
        kind: EmbeddedResourceKind::Image,
        source: EmbeddedResourceSource::InlineBinary {
            bytes: bytes.to_vec(),
            media_type: Some("image/png".to_string()),
        },
        frame_name: Some(name.to_string()),
        xml_id: None,
        class_id: None,
    };
    let mut setup = source.edit();
    setup
        .add_embedded_resource(&linked("object-one"))
        .unwrap()
        .add_embedded_resource(&linked("object-two"))
        .unwrap()
        .add_embedded_chart(&Definition::new(ChartClass::line()))
        .unwrap()
        .add_embedded_chart(&Definition::new(ChartClass::bar()))
        .unwrap()
        .add_embedded_resource(&image("image-one", b"one"))
        .unwrap()
        .add_embedded_resource(&image("image-two", b"two"))
        .unwrap();
    let setup = setup.commit().unwrap().into_snapshot();

    let mut edit = setup.edit();
    edit.replace_embedded_chart(2, &Definition::new(ChartClass::ring()))
        .unwrap()
        .remove_embedded_chart_at(Position::new(3))
        .unwrap()
        .replace_embedded_object(0, &linked("object-replacement"))
        .unwrap()
        .move_embedded_object_to(Position::new(0), Position::new(1))
        .unwrap()
        .remove_embedded_object_at(Position::new(1))
        .unwrap()
        .replace_embedded_image(0, &image("image-replacement", b"replacement"))
        .unwrap()
        .move_embedded_image_to(Position::new(0), Position::new(1))
        .unwrap()
        .remove_embedded_image_at(Position::new(1))
        .unwrap();
    let committed = edit.commit().unwrap();
    let durable = committed.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    for operation in [
        "chart.replace",
        "chart.remove",
        "resource.embedded.object.replace",
        "resource.embedded.object.move",
        "resource.embedded.object.remove",
        "resource.embedded.image.replace",
        "resource.embedded.image.move",
        "resource.embedded.image.remove",
    ] {
        assert!(wire_text.contains(operation), "missing {operation}");
    }
    assert_eq!(
        durable.apply(&setup).unwrap().as_bytes(),
        committed.snapshot().as_bytes()
    );
    assert_eq!(
        committed
            .patch()
            .plan_transfer(&setup)
            .unwrap()
            .commit()
            .unwrap()
            .snapshot()
            .as_bytes(),
        committed.snapshot().as_bytes()
    );
}

#[test]
fn cross_document_transfer_is_dependency_checked_and_refuses_source_local_edits() {
    let source_snapshot = source().snapshot().unwrap();
    let mut portable = source_snapshot.edit();
    portable
        .insert_paragraph(Position::new(1), "portable")
        .unwrap()
        .append_run(Position::new(1), " addition", None)
        .unwrap();
    let portable = portable.commit().unwrap();

    let destination = source().snapshot().unwrap();
    let plan = portable.patch().plan_transfer(&destination).unwrap();
    assert_eq!(plan.operation_count(), 2);
    assert_eq!(plan.dependencies().len(), 1);
    assert_eq!(
        plan.dependencies()[0].kind(),
        TransferDependencyKind::Paragraph
    );
    assert!(plan.dependencies()[0].is_satisfied());
    assert_eq!(
        plan.commit()
            .unwrap()
            .snapshot()
            .document()
            .unwrap()
            .text()
            .unwrap(),
        "Unique paragraph\nportable addition"
    );

    let mut local = source_snapshot.edit();
    local
        .replace_paragraph(Position::new(0), "source-local")
        .unwrap();
    let local = local.commit().unwrap();
    assert!(local.patch().plan_transfer(&destination).is_err());

    let empty = {
        let document = MutableDocument::new();
        Document::from_bytes(document.to_bytes().unwrap())
            .unwrap()
            .snapshot()
            .unwrap()
    };
    let unresolved = portable.patch().plan_transfer(&empty).unwrap();
    assert!(
        unresolved
            .dependencies()
            .iter()
            .any(|dependency| !dependency.is_satisfied())
    );
    assert!(unresolved.commit().is_err());
}

#[test]
fn genuine_libreoffice_package_survives_transaction_reopen_and_full_resave() {
    let producer = Document::from_bytes(
        include_bytes!(
            "../../../test-data/libreoffice-core/sw/qa/extras/odfexport/data/Formcontrol needs high z-index.odt"
        )
        .to_vec(),
    )
    .unwrap();
    let snapshot = producer.snapshot().unwrap();
    let triple = Triple {
        subject: Subject::Iri("urn:litchi:libreoffice".to_string()),
        predicate: "urn:litchi:review#reopened".to_string(),
        object: Object::Literal {
            value: "true".to_string(),
            datatype: None,
            language: None,
        },
    };
    let chart = Definition::new(ChartClass::line());
    let resource = EmbeddedResource {
        kind: EmbeddedResourceKind::Object,
        source: EmbeddedResourceSource::Linked {
            href: "https://example.invalid/writer-resource".to_string(),
        },
        frame_name: Some("Writer Resource".to_string()),
        xml_id: Some("writer-resource".to_string()),
        class_id: None,
    };
    let mut edit = snapshot.edit();
    edit.add_rdf_graph(
        Some("metadata/litchi-reopen.rdf"),
        std::slice::from_ref(&triple),
    )
    .unwrap()
    .add_embedded_chart(&chart)
    .unwrap()
    .add_embedded_resource(&resource)
    .unwrap();
    let committed = edit.commit().unwrap();
    let chart_index = match committed.results().get(1) {
        Some(OperationResult::Index(index)) => *index,
        other => panic!("unexpected embedded chart result: {other:?}"),
    };

    let first_reopen = Document::from_bytes(committed.snapshot().as_bytes().to_vec()).unwrap();
    let resaved = first_reopen.to_bytes().unwrap();
    let second_reopen = Document::from_bytes(resaved).unwrap();
    let graph = second_reopen
        .rdf_graphs()
        .unwrap()
        .into_iter()
        .find(|graph| graph.path == "metadata/litchi-reopen.rdf")
        .unwrap();
    assert_eq!(graph.triples, vec![triple]);
    assert!(!second_reopen.forms().unwrap().groups.is_empty());
    second_reopen.embedded_chart(chart_index).unwrap();
    assert!(second_reopen.embedded_objects().unwrap().len() >= 2);
}

#[test]
fn raw_zip_signature_marker_is_classified_and_never_mutated() {
    let content = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>signed raw fixture</text:p></office:text></office:body></office:document-content>"#;
    let bytes = support::package(
        "application/vnd.oasis.opendocument.text",
        &[
            ("content.xml", content.as_slice()),
            (
                "META-INF/documentsignatures.xml",
                b"not parsed or trusted signature bytes",
            ),
        ],
    );
    let snapshot = Document::from_bytes(bytes).unwrap().snapshot().unwrap();
    assert_eq!(snapshot.envelope_kind().unwrap(), EnvelopeKind::Signed);
    let mut edit = snapshot.edit();
    edit.append_line_break(ParagraphSelector::position(Position::new(0)))
        .unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn sealed_durable_patch_discards_reverse_artifact_and_checks_source_digest() {
    let source = source();
    let snapshot = source.snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.append_line_break(ParagraphSelector::position(Position::new(0)))
        .unwrap();
    let commit = edit.commit().unwrap();
    let reversible_json = commit
        .patch()
        .durable()
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let sealed_json = commit
        .patch()
        .durable()
        .unwrap()
        .seal()
        .to_deterministic_json()
        .unwrap();
    assert!(sealed_json.len() < reversible_json.len());

    let sealed =
        litchi_odt::transaction::SealedPatch::from_deterministic_json(&sealed_json).unwrap();
    assert_eq!(
        sealed.apply(&snapshot).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert!(sealed.apply(commit.snapshot()).is_err());
}

#[test]
fn durable_patch_parser_rejects_noncanonical_json() {
    let source = source();
    let snapshot = source.snapshot().unwrap();
    let json = snapshot
        .edit()
        .commit()
        .unwrap()
        .patch()
        .durable()
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let mut noncanonical = Vec::with_capacity(json.len() + 1);
    noncanonical.push(b' ');
    noncanonical.extend_from_slice(&json);
    assert!(litchi_odt::transaction::DurablePatch::from_deterministic_json(&noncanonical).is_err());

    let mut foreign_format = json;
    let format = foreign_format
        .windows(b"litchi.odt".len())
        .position(|window| window == b"litchi.odt")
        .unwrap();
    foreign_format[format + b"litchi.odt".len() - 1] = b'x';
    assert!(
        litchi_odt::transaction::DurablePatch::from_deterministic_json(&foreign_format).is_err()
    );
}

#[test]
fn semantic_durable_patch_replays_paragraph_runs_and_hyperlinks() {
    let source = source();
    let snapshot = source.snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.insert_paragraph(Position::new(1), "second")
        .unwrap()
        .replace_paragraph(Position::new(0), "first")
        .unwrap()
        .append_run(Position::new(0), " styled", Some("Emphasis"))
        .unwrap()
        .append_hyperlink(Position::new(1), "https://example.invalid/review", "review")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut repeated = snapshot.edit();
    repeated
        .insert_paragraph(Position::new(1), "second")
        .unwrap()
        .replace_paragraph(Position::new(0), "first")
        .unwrap()
        .append_run(Position::new(0), " styled", Some("Emphasis"))
        .unwrap()
        .append_hyperlink(Position::new(1), "https://example.invalid/review", "review")
        .unwrap();
    assert_eq!(
        repeated.commit().unwrap().snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    let durable = commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let wire_text = std::str::from_utf8(&wire).unwrap();
    assert!(wire_text.contains("paragraph.insert"));
    assert!(wire_text.contains("paragraph.replace"));
    assert!(wire_text.contains("run.append"));
    assert!(wire_text.contains("hyperlink.append"));
    assert!(!wire_text.contains("document.replace"));

    let reopened = durable.apply(&snapshot).unwrap().document().unwrap();
    let paragraphs = reopened.paragraphs().unwrap();
    assert_eq!(paragraphs[0].text().unwrap(), "first styled");
    assert_eq!(paragraphs[1].text().unwrap(), "secondreview");
    assert_eq!(
        reopened.hyperlinks().unwrap()[0].1,
        "https://example.invalid/review"
    );
    assert_eq!(
        durable
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        snapshot.as_bytes()
    );
}

#[test]
fn deterministic_join_reports_exact_conflicts_and_commits_in_identifier_order() {
    let mut document = MutableDocument::new();
    document.add_paragraph("zero").unwrap();
    document.add_paragraph("one").unwrap();
    let snapshot = Document::from_bytes(document.to_bytes().unwrap())
        .unwrap()
        .snapshot()
        .unwrap();

    let mut joined = snapshot.joined_edit();
    let mut second = snapshot.edit();
    second.replace_paragraph(Position::new(1), "ONE").unwrap();
    joined.join("z-second", second).unwrap();
    let mut first = snapshot.edit();
    first.replace_paragraph(Position::new(0), "ZERO").unwrap();
    joined.join("a-first", first).unwrap();

    let mut overlapping = snapshot.edit();
    overlapping.append_run(Position::new(0), "!", None).unwrap();
    // A whole-paragraph replacement and an inline append intentionally share
    // one paragraph effect key through conservative conflict ownership.
    let conflict = match joined.join("overlap", overlapping) {
        Ok(_) => panic!("overlapping ODT edits must conflict"),
        Err(conflict) => conflict,
    };
    assert!(matches!(conflict.failure(), SubEditJoinFailure::Overlap(_)));
    assert!(!conflict.conflicts().unwrap().is_empty());

    let committed = joined.commit().unwrap();
    let paragraphs = committed
        .snapshot()
        .document()
        .unwrap()
        .paragraphs()
        .unwrap();
    assert_eq!(paragraphs[0].text().unwrap(), "ZERO");
    assert_eq!(paragraphs[1].text().unwrap(), "ONE");
}

#[test]
fn three_way_plan_is_non_applying_and_requires_explicit_conflict_resolution() {
    let snapshot = source().snapshot().unwrap();
    let mut left = snapshot.joined_edit();
    let mut left_edit = snapshot.edit();
    left_edit
        .replace_paragraph(Position::new(0), "left")
        .unwrap();
    left.join("left", left_edit).unwrap();

    let mut right = snapshot.joined_edit();
    let mut right_edit = snapshot.edit();
    right_edit
        .append_run(Position::new(0), " right", None)
        .unwrap();
    right.join("right", right_edit).unwrap();

    let mut plan = MergePlan::new(left, right).unwrap();
    assert_eq!(plan.conflicts().len(), 1);
    assert_eq!(
        snapshot.document().unwrap().text().unwrap(),
        "Unique paragraph"
    );
    plan.resolve(MergeChoice::Left);
    let joined = match plan.finish() {
        Ok(joined) => joined,
        Err(_) => panic!("resolved ODT merge plan must finish"),
    };
    let merged = joined.commit().unwrap();
    assert_eq!(
        merged.snapshot().document().unwrap().text().unwrap(),
        "left"
    );
}

#[test]
fn bounded_history_rejects_stale_edits_and_fully_reopens_undo_redo() {
    let snapshot = source().snapshot().unwrap();
    let mut history = snapshot.history(HistoryLimits::new(4, 64 * 1024 * 1024));
    let stale = history.edit();
    let mut current = history.edit();
    current
        .replace_paragraph(Position::new(0), "committed")
        .unwrap();
    history.commit(current).unwrap();
    assert_eq!(history.document().unwrap().text().unwrap(), "committed");
    assert!(history.commit(stale).is_err());
    assert_eq!(history.document().unwrap().text().unwrap(), "committed");
    assert!(history.undo());
    assert_eq!(
        history.document().unwrap().text().unwrap(),
        "Unique paragraph"
    );
    assert!(history.redo());
    assert_eq!(history.document().unwrap().text().unwrap(), "committed");
}
