use litchi_odt::{
    Block, ChangeType, Document, DocumentBuilder, MutableDocument, Position, Section,
    SectionDdeSource, SectionDisplay, SectionSource, Story, TrackChange, TrackedChanges,
    wrap_section_xml,
};

fn section(name: &str) -> Section {
    Section {
        name: name.to_string(),
        style: Some("Section Style".to_string()),
        protected: false,
        xml_id: Some(format!("id_{name}")),
        protection_key: None,
        protection_key_digest_algorithm: None,
        display: SectionDisplay::Visible,
        condition: None,
        source: None,
        dde_source: None,
        content: "Unicode 😀 <& content".to_string(),
    }
}

fn insertion() -> TrackChange {
    TrackChange {
        id: "insert_1".to_string(),
        xml_id: Some("insert_1".to_string()),
        author: Some("Author".to_string()),
        date: Some("2026-07-19".to_string()),
        comment: Some("inside section".to_string()),
        change_type: ChangeType::Insertion,
        style_name: None,
        merge_last_paragraph: None,
        content: String::new(),
    }
}

#[test]
fn builder_authors_linked_conditional_protected_and_dde_sections() {
    let mut linked = section("Linked");
    linked.protected = true;
    linked.protection_key = Some("YWJj".to_string());
    linked.protection_key_digest_algorithm = Some("urn:example:sha256".to_string());
    linked.display = SectionDisplay::Condition;
    linked.condition = Some("ooow:visible()".to_string());
    linked.source = Some(SectionSource {
        href: Some("https://example.invalid/source.odt".to_string()),
        section_name: Some("Remote".to_string()),
        filter_name: Some("writer8".to_string()),
    });
    let mut dde = section("Dde");
    dde.xml_id = Some("dde_id".to_string());
    dde.content.clear();
    dde.dde_source = Some(SectionDdeSource {
        name: Some("Feed".to_string()),
        conversion_mode: Some("keep-text".to_string()),
        automatic_update: Some(false),
    });

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("before").unwrap();
    builder.add_section(linked).unwrap();
    builder.add_section(dde).unwrap();
    let document = Document::from_bytes(builder.build().unwrap()).unwrap();
    let sections = document.sections().unwrap();
    assert_eq!(sections.len(), 2);
    assert_eq!(sections[0].display, SectionDisplay::Condition);
    assert_eq!(sections[0].content, "Unicode 😀 <& content");
    assert_eq!(
        sections[0].source.as_ref().unwrap().href.as_deref(),
        Some("https://example.invalid/source.odt")
    );
    assert_eq!(
        sections[1].dde_source.as_ref().unwrap().name.as_deref(),
        Some("Feed")
    );
}

#[test]
fn mutable_wraps_nested_sections_and_preserves_tracked_changes() {
    let mut document = MutableDocument::new();
    document.add_paragraph("A😀B").unwrap();
    document.add_paragraph("Second").unwrap();
    document
        .set_tracked_changes(TrackedChanges {
            changes: vec![insertion()],
            ..TrackedChanges::default()
        })
        .unwrap();
    document
        .mark_tracked_change_range(
            "insert_1",
            Position {
                story: Story::Paragraph(0),
                character: 1,
            },
            Position {
                story: Story::Paragraph(0),
                character: 2,
            },
        )
        .unwrap();

    let mut outer = section("Outer");
    outer.content.clear();
    document
        .wrap_section(&outer, Block::BodyParagraph(0), Block::BodyParagraph(1))
        .unwrap();
    let mut inner = section("Inner");
    inner.xml_id = Some("inner_id".to_string());
    inner.content.clear();
    document
        .wrap_section(&inner, Block::BodyParagraph(1), Block::BodyParagraph(1))
        .unwrap();
    assert_eq!(document.sections().unwrap().len(), 2);
    assert_eq!(document.tracked_changes().unwrap().changes[0].content, "😀");

    let mut replacement = section("OuterRenamed");
    replacement.xml_id = Some("outer_new".to_string());
    replacement.display = SectionDisplay::Hidden;
    replacement.content.clear();
    document.update_section("Outer", &replacement).unwrap();
    assert_eq!(document.sections().unwrap()[0].name, "OuterRenamed");
    assert_eq!(document.tracked_changes().unwrap().changes[0].content, "😀");

    document.unwrap_section("Inner").unwrap();
    assert_eq!(document.sections().unwrap().len(), 1);
    let reopened = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.sections().unwrap()[0].name, "OuterRenamed");
    assert_eq!(reopened.tracked_changes().unwrap().changes[0].content, "😀");
}

#[test]
fn mixed_body_and_table_cell_ranges_are_lossless_and_errors_roll_back() {
    let xml = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:text><text:p>A<text:span text:style-name="S">😀</text:span></text:p><table:table table:name="T"><table:table-row><table:table-cell><text:p>Cell 1</text:p><text:p>Cell 2</text:p></table:table-cell></table:table-row></table:table><text:p>Tail</text:p></office:text></office:body></office:document-content>"#;
    let mut body = section("BodyMixed");
    body.content.clear();
    let wrapped =
        wrap_section_xml(xml, &body, &Block::BodyParagraph(0), &Block::BodyTable(0)).unwrap();
    assert!(wrapped.contains(r#"<text:span text:style-name="S">😀</text:span>"#));
    assert!(wrapped.contains("<table:table table:name=\"T\">"));

    let mut cell = section("CellNested");
    cell.xml_id = Some("cell_nested".to_string());
    cell.content.clear();
    let wrapped = wrap_section_xml(
        &wrapped,
        &cell,
        &Block::TableCellParagraph {
            table: 0,
            row: 0,
            cell: 0,
            paragraph: 0,
        },
        &Block::TableCellParagraph {
            table: 0,
            row: 0,
            cell: 0,
            paragraph: 1,
        },
    )
    .unwrap();
    assert!(wrapped.contains("Cell 1"));
    assert!(wrapped.contains("Cell 2"));

    let duplicate = section("BodyMixed");
    assert!(litchi_odt::add_section_xml(&wrapped, &duplicate).is_err());
    let crossing = wrap_section_xml(
        &wrapped,
        &section("Crossing"),
        &Block::BodyTable(0),
        &Block::BodyParagraph(1),
    );
    assert!(crossing.is_err());
}

#[test]
fn remove_deletes_content_while_unwrap_and_clear_retain_it() {
    let mut unwrap = MutableDocument::new();
    unwrap.add_paragraph("keep me").unwrap();
    let mut wrapper = section("Wrapper");
    wrapper.content.clear();
    unwrap
        .wrap_section(&wrapper, Block::BodyParagraph(0), Block::BodyParagraph(0))
        .unwrap();
    unwrap.unwrap_section("Wrapper").unwrap();
    assert_eq!(
        Document::from_bytes(unwrap.to_bytes().unwrap())
            .unwrap()
            .text()
            .unwrap(),
        "keep me"
    );

    let mut remove = MutableDocument::new();
    remove.add_paragraph("delete me").unwrap();
    remove
        .wrap_section(&wrapper, Block::BodyParagraph(0), Block::BodyParagraph(0))
        .unwrap();
    remove.remove_section("Wrapper").unwrap();
    assert!(
        Document::from_bytes(remove.to_bytes().unwrap())
            .unwrap()
            .text()
            .unwrap()
            .is_empty()
    );
}
