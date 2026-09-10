use litchi_core::Position;
use litchi_oth::{Builder, Patch, Template};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:foo="urn:future:example"><office:body><office:text><text:section text:name="s"><text:p>section &amp; old</text:p><foo:future foo:value="keep"/></text:section><text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>note old</text:p></text:note-body></text:note><office:annotation office:name="a"><text:p>annotation old</text:p></office:annotation><foo:section><text:p>foreign owner</text:p></foo:section></office:text></office:body></office:document-content>"#;

fn template() -> Template {
    Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap()
}

#[test]
fn section_note_annotation_splices_are_source_bound_and_reversible() {
    let source = template();
    let mut edit = source.edit();
    edit.set_section_text(Position::new(0), "section new & text")
        .unwrap();
    edit.set_note_body(Position::new(0), "note new").unwrap();
    edit.set_annotation_text(Position::new(0), "annotation new")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(
        commit
            .template()
            .content_xml()
            .contains("section new &amp; text")
    );
    assert!(commit.template().content_xml().contains("note new"));
    assert!(commit.template().content_xml().contains("annotation new"));
    assert!(
        commit
            .template()
            .content_xml()
            .contains("foo:future foo:value=\"keep\"")
    );
    assert!(commit.template().content_xml().contains("foreign owner"));
    assert_eq!(commit.patch().structure_changes().len(), 3);

    let wire = commit.patch().to_bytes().unwrap();
    let durable = Patch::from_bytes(&wire).unwrap();
    assert_eq!(durable.structure_changes().len(), 3);
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.template().as_bytes()
    );
    let inverse = Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    assert_eq!(
        inverse.apply(commit.template()).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn structural_noop_and_rich_structure_refusal_are_atomic() {
    let source = template();
    let mut noop = source.edit();
    noop.set_section_text(Position::new(0), "section & old")
        .unwrap();
    let commit = noop.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.template().as_bytes(), source.as_bytes());

    let rich = Builder::new()
        .content_xml(
            CONTENT
                .replace(
                    "<text:p>section &amp; old</text:p>",
                    "<text:p><text:span>rich</text:span></text:p>",
                )
                .as_str(),
        )
        .build()
        .unwrap();
    let rich = Template::from_bytes(rich).unwrap();
    assert!(
        rich.edit()
            .set_section_text(Position::new(0), "refuse")
            .is_err()
    );
}

#[test]
fn foreign_same_named_structure_is_not_selectable() {
    let source = template();
    assert!(
        source
            .edit()
            .set_section_text(Position::new(1), "foreign")
            .is_err()
    );
}

#[test]
fn structure_lifecycle_and_table_append_are_durable_and_exactly_reversible() {
    let source_xml = CONTENT.replace(
        "<text:note text:note-class=\"footnote\">",
        "<text:note text:id=\"n0\" text:note-class=\"footnote\">",
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(&source_xml).build().unwrap()).unwrap();
    let mut edit = source.edit();
    edit.append_section("created", "new section").unwrap();
    edit.append_note("n1", litchi_oth::note::NoteClass::Endnote, "2", "new note")
        .unwrap();
    edit.append_annotation("created-a", "new annotation")
        .unwrap();
    edit.append_table("created-t", "new cell").unwrap();
    edit.remove_section(Position::new(0)).unwrap();
    edit.remove_note(Position::new(0)).unwrap();
    edit.remove_annotation(Position::new(0)).unwrap();
    let commit = edit.commit().unwrap();
    let body = commit.template().text_body().unwrap();
    assert_eq!(body.sections().unwrap().len(), 1);
    assert_eq!(body.sections().unwrap()[0].name(), Some("created"));
    assert_eq!(body.notes().unwrap().len(), 1);
    assert_eq!(body.notes().unwrap()[0].id(), Some("n1"));
    assert_eq!(body.annotations().unwrap().len(), 1);
    assert_eq!(body.annotations().unwrap()[0].name(), Some("created-a"));
    assert_eq!(body.tables().unwrap().len(), 1);
    assert_eq!(body.tables().unwrap()[0].name(), Some("created-t"));
    assert_eq!(body.tables().unwrap()[0].column_count(), 1);
    assert_eq!(
        body.tables().unwrap()[0].rows()[0].cells()[0].text(),
        "new cell"
    );
    assert_eq!(commit.patch().structure_lifecycle_changes().len(), 4 + 3);
    let durable = Patch::from_bytes(&commit.patch().to_bytes().unwrap()).unwrap();
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.template().as_bytes()
    );
    let inverse = Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    assert_eq!(
        inverse.apply(commit.template()).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn structural_lifecycle_rejects_ambiguous_or_referenced_identity() {
    let source_xml = CONTENT.replace(
        "<foo:section><text:p>foreign owner</text:p></foo:section>",
        "<foo:section xlink:href=\"s\"><text:p>foreign owner</text:p></foo:section>",
    );
    let source = Template::from_bytes(
        Builder::new()
            .content_xml(source_xml.replace(
                "xmlns:foo=\"urn:future:example\"",
                "xmlns:foo=\"urn:future:example\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
            ))
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(source.edit().remove_section(Position::new(0)).is_err());
}

#[test]
fn semantic_references_decode_xml_and_uri_forms_before_removal() {
    let source_xml = CONTENT
        .replace(
            "xmlns:foo=\"urn:future:example\"",
            "xmlns:foo=\"urn:future:example\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        )
        .replace(
            "</text:section>",
            "</text:section><foo:reference xlink:href=\"#s%31\"/>",
        )
        .replace("text:name=\"s\"", "text:name=\"s1\"");
    let source =
        Template::from_bytes(Builder::new().content_xml(source_xml).build().unwrap()).unwrap();
    assert!(source.edit().remove_section(Position::new(0)).is_err());

    let source_xml = CONTENT
        .replace(
            "xmlns:foo=\"urn:future:example\"",
            "xmlns:foo=\"urn:future:example\" xmlns:text2=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"",
        )
        .replace(
            "</text:section>",
            "</text:section><foo:reference text2:section-name=\"s&#x31;\"/>",
        )
        .replace("text:name=\"s\"", "text:name=\"s1\"");
    let source =
        Template::from_bytes(Builder::new().content_xml(source_xml).build().unwrap()).unwrap();
    assert!(source.edit().remove_section(Position::new(0)).is_err());
}

#[test]
fn removing_an_earlier_structure_remaps_later_text_identity_and_patch_wire() {
    let source_xml = CONTENT.replace(
        "</text:section>",
        "</text:section><text:section text:name=\"later\"><text:p>later old</text:p></text:section>",
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(source_xml).build().unwrap()).unwrap();
    let mut edit = source.edit();
    edit.remove_section(Position::new(0)).unwrap();
    edit.set_section_text(Position::new(1), "later new")
        .unwrap();
    let commit = edit.commit().unwrap();
    let body = commit.template().text_body().unwrap();
    let sections = body.sections().unwrap();
    assert_eq!(sections.len(), 1);
    assert_eq!(sections[0].name(), Some("later"));
    assert_eq!(sections[0].text(), "later new");
    let change = &commit.patch().structure_changes()[0];
    assert_eq!(change.selector().position(), Position::new(1));
    assert_eq!(change.target_selector().position(), Position::new(0));
    assert_eq!(change.identity(), Some("later"));

    let durable = Patch::from_bytes(&commit.patch().to_bytes().unwrap()).unwrap();
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.template().as_bytes()
    );
    let inverse = Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    assert_eq!(
        inverse.apply(commit.template()).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn fresh_notes_reject_untyped_other_classes() {
    let source = template();
    assert!(
        source
            .edit()
            .append_note(
                "n-other",
                litchi_oth::note::NoteClass::Other("vendor-note".to_string()),
                "1",
                "body",
            )
            .is_err()
    );
}

#[test]
fn lifecycle_join_refuses_same_tail_and_delete_text_overlap() {
    let source_xml = CONTENT.replace(
        "<text:note text:note-class=\"footnote\">",
        "<text:note text:id=\"n0\" text:note-class=\"footnote\">",
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(&source_xml).build().unwrap()).unwrap();
    let mut left = source.edit();
    left.append_table("left", "one").unwrap();
    let mut right = source.edit();
    right.append_table("right", "two").unwrap();
    assert!(
        matches!(left.join(right), Err(error) if error.failure() == litchi_oth::JoinFailure::Append)
    );

    let mut removing = source.edit();
    removing.remove_section(Position::new(0)).unwrap();
    let mut editing = source.edit();
    editing.set_section_text(Position::new(0), "new").unwrap();
    assert!(removing.join(editing).is_err());
}
