use super::super::support::*;

#[test]
fn tracked_text_revisions_round_trip_through_both_output_paths() {
    let timestamp = DateTime {
        year: 2026,
        month: 7,
        day: 15,
        hour: 14,
        minute: 30,
        weekday: 3,
    };
    let mut writer = Writer::new();
    writer.set_section_formatting_revision(
        FormattingRevision::new("Section Editor").with_timestamp(timestamp),
    );
    writer
        .add_paragraph_runs(
            vec![
                (
                    "inserted ".to_string(),
                    CharacterFormatting {
                        insertion_revision: Some(
                            TextRevision::new("Alice 😀")
                                .with_timestamp(timestamp)
                                .with_reason(crate::RevisionReason::from_raw(42).unwrap())
                                .with_revision_save_id(0x11223344),
                        ),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "deleted".to_string(),
                    CharacterFormatting {
                        deletion_revision: Some(
                            TextRevision::new("Bob")
                                .with_id(7)
                                .with_revision_save_id(0x55667788),
                        ),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    " formatted".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        formatting_revision: Some(
                            FormattingRevision::new("张三")
                                .with_timestamp(timestamp)
                                .with_reason(crate::RevisionReason::APPLIED_STYLE)
                                .with_revision_save_id(0x99AABBCC),
                        ),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "\u{13}".to_string(),
                    CharacterFormatting {
                        special: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    " LISTNUM ".to_string(),
                    CharacterFormatting {
                        field_vanish: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "\u{14}".to_string(),
                    CharacterFormatting {
                        special: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "12.".to_string(),
                    CharacterFormatting {
                        display_field_revision: Some(
                            DisplayFieldRevision::new("Field Editor", "11.")
                                .with_timestamp(timestamp),
                        ),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "\u{15}".to_string(),
                    CharacterFormatting {
                        special: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
            ],
            ParagraphFormatting {
                alignment: Some(1),
                formatting_revision: Some(
                    FormattingRevision::new("Paragraph Editor").with_timestamp(timestamp),
                ),
                numbering_revision_list_applied: Some(true),
                numbering_revision: Some(NumberingRevision {
                    was_numbered: true,
                    placeholder_positions: [1, 0, 0, 0, 0, 0, 0, 0, 0],
                    numbers: [12, 0, 0, 0, 0, 0, 0, 0, 0],
                    ..NumberingRevision::new("Numbering Editor", "%.").with_timestamp(timestamp)
                }),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    assert_eq!(
        document.revision_authors(),
        [
            "Unknown",
            "Section Editor",
            "Paragraph Editor",
            "Numbering Editor",
            "Alice 😀",
            "Bob",
            "张三",
            "Field Editor"
        ]
    );
    let section_revision = &document.section_revisions()[0];
    assert_eq!(section_revision.start, 0);
    assert!(section_revision.end > section_revision.start);
    assert_eq!(section_revision.author, "Section Editor");
    assert_eq!(section_revision.timestamp, Some(timestamp));
    let paragraphs = document.paragraphs().unwrap();
    let paragraph_revision = paragraphs[0].formatting_revision().unwrap();
    assert_eq!(paragraph_revision.author, "Paragraph Editor");
    assert_eq!(paragraph_revision.timestamp, Some(timestamp));
    assert_eq!(paragraphs[0].numbering_revision_list_applied(), Some(true));
    let numbering_revision = paragraphs[0].numbering_revision().unwrap();
    assert_eq!(numbering_revision.author, "Numbering Editor");
    assert_eq!(numbering_revision.timestamp, Some(timestamp));
    assert!(numbering_revision.was_numbered);
    assert_eq!(numbering_revision.placeholder_positions[0], 1);
    assert_eq!(numbering_revision.numbers[0], 12);
    assert_eq!(numbering_revision.format_string, "%.");
    let runs = paragraphs[0].runs().unwrap();
    let insertion = runs
        .iter()
        .find_map(|run| run.insertion_revision())
        .unwrap();
    assert_eq!(insertion.author, "Alice 😀");
    assert_eq!(insertion.timestamp, Some(timestamp));
    assert_eq!(insertion.reason.unwrap().raw(), 42);
    assert_eq!(insertion.revision_id, Some(42));
    assert_eq!(insertion.revision_save_id, Some(0x11223344));
    let deletion = runs.iter().find_map(|run| run.deletion_revision()).unwrap();
    assert_eq!(deletion.author, "Bob");
    assert_eq!(deletion.timestamp, None);
    assert_eq!(deletion.reason.unwrap().raw(), 7);
    assert_eq!(deletion.revision_id, Some(7));
    assert_eq!(deletion.revision_save_id, Some(0x55667788));
    let formatting = runs
        .iter()
        .find_map(|run| run.formatting_revision())
        .unwrap();
    assert_eq!(formatting.kind, crate::RevisionKind::Formatting);
    assert_eq!(formatting.author, "张三");
    assert_eq!(formatting.timestamp, Some(timestamp));
    assert_eq!(
        formatting.reason,
        Some(crate::RevisionReason::APPLIED_STYLE)
    );
    assert_eq!(formatting.revision_id, Some(1));
    assert_eq!(formatting.revision_save_id, Some(0x99AABBCC));
    let display_field = runs
        .iter()
        .find_map(|run| run.display_field_revision())
        .unwrap();
    assert_eq!(display_field.author, "Field Editor");
    assert_eq!(display_field.timestamp, Some(timestamp));
    assert_eq!(display_field.previous_result, "11.");

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-revisions-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    let document = package.document().unwrap();
    assert_eq!(
        document.revision_authors(),
        [
            "Unknown",
            "Section Editor",
            "Paragraph Editor",
            "Numbering Editor",
            "Alice 😀",
            "Bob",
            "张三",
            "Field Editor"
        ]
    );
    assert_eq!(document.section_revisions()[0].author, "Section Editor");
    assert!(
        document.paragraphs().unwrap()[0]
            .formatting_revision()
            .is_some()
    );
    assert!(
        document.paragraphs().unwrap()[0]
            .numbering_revision()
            .is_some()
    );
    assert!(
        document.paragraphs().unwrap()[0]
            .runs()
            .unwrap()
            .iter()
            .any(|run| run.deletion_revision().is_some())
    );
    assert!(
        document.paragraphs().unwrap()[0]
            .runs()
            .unwrap()
            .iter()
            .any(|run| run.formatting_revision().is_some())
    );
    assert!(
        document.paragraphs().unwrap()[0]
            .runs()
            .unwrap()
            .iter()
            .any(|run| run.display_field_revision().is_some())
    );
    std::fs::remove_file(path).unwrap();
}
