use super::super::support::*;

#[test]
fn rejects_invalid_writer_revision_metadata() {
    let error_for = |formatting: CharacterFormatting| {
        let mut writer = Writer::new();
        writer
            .add_paragraph_runs(
                vec![("text".to_string(), formatting)],
                ParagraphFormatting::default(),
            )
            .unwrap();
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
    };
    let both = CharacterFormatting {
        insertion_revision: Some(TextRevision::new("Alice")),
        deletion_revision: Some(TextRevision::new("Alice")),
        ..CharacterFormatting::default()
    };
    assert!(error_for(both).contains("both an insertion and a deletion"));

    let nested = CharacterFormatting {
        preserved_properties_for_revision: Some(Box::new(CharacterFormatting {
            preserved_properties_for_revision: Some(Box::new(CharacterFormatting::default())),
            ..CharacterFormatting::default()
        })),
        ..CharacterFormatting::default()
    };
    assert!(error_for(nested).contains("nested preserved states"));

    let invalid_reason = CharacterFormatting {
        insertion_revision: Some(TextRevision::new("Alice").with_id(0x002C)),
        ..CharacterFormatting::default()
    };
    assert!(error_for(invalid_reason).contains("reason code is undefined"));

    let conflicting_reason = CharacterFormatting {
        insertion_revision: Some(
            TextRevision::new("Alice")
                .with_id(1)
                .with_reason(crate::RevisionReason::NORMAL_EDIT),
        ),
        ..CharacterFormatting::default()
    };
    assert!(error_for(conflicting_reason).contains("conflicting"));

    let conflicting_formatting_reason = CharacterFormatting {
        insertion_revision: Some(
            TextRevision::new("Alice").with_reason(crate::RevisionReason::NORMAL_EDIT),
        ),
        formatting_revision: Some(
            FormattingRevision::new("Alice").with_reason(crate::RevisionReason::APPLIED_STYLE),
        ),
        ..CharacterFormatting::default()
    };
    assert!(error_for(conflicting_formatting_reason).contains("insertion and formatting"));

    let invalid_time = CharacterFormatting {
        insertion_revision: Some(TextRevision::new("Alice").with_timestamp(CommentDateTime {
            year: 2026,
            month: 13,
            day: 1,
            hour: 0,
            minute: 0,
            weekday: 0,
        })),
        ..CharacterFormatting::default()
    };
    assert!(error_for(invalid_time).contains("timestamp"));

    let mut writer = Writer::new();
    writer.set_section_formatting_revision(FormattingRevision::new("Editor").with_timestamp(
        CommentDateTime {
            year: 2026,
            month: 0,
            day: 1,
            hour: 0,
            minute: 0,
            weekday: 0,
        },
    ));
    writer.add_paragraph("text").unwrap();
    assert!(
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
            .contains("timestamp")
    );

    let mut writer = Writer::new();
    writer
        .add_paragraph_runs(
            vec![("text".to_string(), CharacterFormatting::default())],
            ParagraphFormatting {
                numbering_revision: Some(NumberingRevision::new("Alice", "x".repeat(32))),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();
    assert!(
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
            .contains("NumRM")
    );

    let invalid_display = CharacterFormatting {
        display_field_revision: Some(DisplayFieldRevision::new("Alice", "x".repeat(16))),
        ..CharacterFormatting::default()
    };
    assert!(error_for(invalid_display).contains("LISTNUM"));
}
