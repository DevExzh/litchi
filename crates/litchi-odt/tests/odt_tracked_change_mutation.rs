use litchi_odt::{
    ChangeType, Document, MutableDocument, OdtTrackedPosition, OdtTrackedStory, TrackChange,
    TrackedChanges, mark_tracked_change_range_xml, set_tracked_changes_xml,
};

fn change(id: &str, kind: ChangeType) -> TrackChange {
    TrackChange {
        id: id.to_string(),
        xml_id: Some(id.to_string()),
        author: Some("Zoë & 李".to_string()),
        date: Some("2026-07-19T09:00:00+08:00".to_string()),
        comment: Some("review <note> & two  spaces".to_string()),
        change_type: kind,
        style_name: (kind == ChangeType::FormatChange).then(|| "Changed Style".to_string()),
        merge_last_paragraph: (kind == ChangeType::Deletion).then_some(false),
        content: if kind == ChangeType::Deletion {
            "deleted 😀 text\nnext".to_string()
        } else {
            Default::default()
        },
    }
}

#[test]
fn marks_a_styled_span_inside_a_table_cell_without_rewriting_markup() {
    let xml = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:text><table:table table:name="T"><table:table-row><table:table-cell><text:p>A<text:span text:style-name="Em">😀B</text:span>C</text:p></table:table-cell></table:table-row></table:table></office:text></office:body></office:document-content>"#;
    let tracked = TrackedChanges {
        changes: vec![change("insert_1", ChangeType::Insertion)],
        ..TrackedChanges::default()
    };
    let with_declaration = set_tracked_changes_xml(xml, Some(&tracked)).unwrap();
    let start = OdtTrackedPosition {
        story: OdtTrackedStory::TableCell {
            table: 0,
            row: 0,
            cell: 0,
            paragraph: 0,
        },
        character: 1,
    };
    let mut end = start.clone();
    end.character = 3;
    let marked =
        mark_tracked_change_range_xml(&with_declaration, "insert_1", &start, &end).unwrap();
    assert!(marked.contains(r#"<text:span text:style-name="Em">😀B"#));
    assert!(marked.contains("text:change-start"));
    assert!(marked.contains("text:change-end"));

    let duplicate_xml_id = marked.replace("<text:p>", "<text:p xml:id=\"insert_1\">");
    assert!(set_tracked_changes_xml(&duplicate_xml_id, Some(&tracked)).is_err());
}

fn position(paragraph: usize, character: usize) -> OdtTrackedPosition {
    OdtTrackedPosition {
        story: OdtTrackedStory::Paragraph(paragraph),
        character,
    }
}

fn declarations() -> TrackedChanges {
    TrackedChanges {
        track_changes: Some(true),
        protection_key: Some("YWJj".to_string()),
        protection_key_digest_algorithm: Some("urn:example:sha256".to_string()),
        changes: vec![
            change("insert_1", ChangeType::Insertion),
            change("format_1", ChangeType::FormatChange),
            change("delete_1", ChangeType::Deletion),
        ],
    }
}

#[test]
fn marks_unicode_ranges_points_and_reopens_package() {
    let mut document = MutableDocument::new();
    document.add_paragraph("A😀B<&").unwrap();
    document.add_paragraph("again").unwrap();
    document.set_tracked_changes(declarations()).unwrap();
    document
        .mark_tracked_change_range("insert_1", position(0, 1), position(0, 2))
        .unwrap();
    document
        .mark_tracked_change_range("insert_1", position(1, 0), position(1, 2))
        .unwrap();
    document
        .mark_tracked_change_range("format_1", position(0, 3), position(0, 5))
        .unwrap();
    document
        .mark_tracked_deletion("delete_1", position(0, 0))
        .unwrap();

    let parsed = document.tracked_changes().unwrap();
    assert_eq!(parsed.changes[0].content, "😀\nag");
    assert_eq!(parsed.changes[1].content, "<&");
    assert_eq!(
        parsed.changes[0].comment.as_deref(),
        Some("review <note> & two  spaces")
    );

    let reopened = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    let parsed = reopened.tracked_changes().unwrap();
    assert_eq!(parsed.track_changes, Some(true));
    assert_eq!(parsed.protection_key.as_deref(), Some("YWJj"));
    assert_eq!(parsed.changes.len(), 3);
    assert_eq!(parsed.changes[2].content, "deleted 😀 text\nnext");
}

#[test]
fn declaration_and_marker_mutations_roll_back_atomically() {
    let mut document = MutableDocument::new();
    document.add_paragraph("abcdef").unwrap();
    document.set_tracked_changes(declarations()).unwrap();
    document
        .mark_tracked_change_range("insert_1", position(0, 1), position(0, 4))
        .unwrap();

    assert!(
        document
            .mark_tracked_change_range("format_1", position(0, 2), position(0, 5))
            .is_err()
    );
    assert_eq!(
        document.tracked_changes().unwrap().changes[0].content,
        "bcd"
    );

    assert!(
        document
            .mark_tracked_deletion("unknown", position(0, 0))
            .is_err()
    );
    let mut duplicate = declarations();
    duplicate
        .changes
        .push(change("insert_1", ChangeType::Insertion));
    assert!(document.set_tracked_changes(duplicate).is_err());
    assert_eq!(document.tracked_changes().unwrap().changes.len(), 3);

    assert!(
        document
            .set_tracked_change_policy(Some(true), Some("not base64".to_string()), None)
            .is_err()
    );
    assert_eq!(
        document
            .tracked_changes()
            .unwrap()
            .protection_key
            .as_deref(),
        Some("YWJj")
    );

    document.unmark_tracked_change("insert_1").unwrap();
    assert!(
        document.tracked_changes().unwrap().changes[0]
            .content
            .is_empty()
    );
    let removed = document.remove_tracked_change("format_1").unwrap();
    assert_eq!(removed.id, "format_1");
    document.clear_tracked_changes().unwrap();
    assert!(document.tracked_changes().unwrap().changes.is_empty());
}
