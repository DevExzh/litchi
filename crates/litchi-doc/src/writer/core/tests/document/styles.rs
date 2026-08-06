use super::super::support::*;

#[test]
fn writes_custom_styles_into_document_stylesheet() {
    let mut writer = Writer::new();
    let paragraph_style = writer
        .add_style(crate::writer::stylesheet::StyleDefinition::new(
            crate::StyleKind::Paragraph,
            "Custom Body",
        ))
        .unwrap();
    let character_style = writer
        .add_style(crate::writer::stylesheet::StyleDefinition::new(
            crate::StyleKind::Character,
            "Custom Emphasis",
        ))
        .unwrap();
    let table_style = writer
        .add_style(crate::writer::stylesheet::StyleDefinition::new(
            crate::StyleKind::Table,
            "Custom Grid",
        ))
        .unwrap();
    assert_eq!(
        (paragraph_style, character_style, table_style),
        (15, 16, 17)
    );
    writer
        .add_paragraph_with_format(
            "Styled document",
            CharacterFormatting {
                style_index: Some(character_style),
                ..CharacterFormatting::default()
            },
            ParagraphFormatting {
                style_index: Some(paragraph_style),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();
    let table = writer.add_table(1, 1).unwrap();
    writer
        .set_table_row_formatting(
            table,
            0,
            crate::writer::tap::TableRow {
                cells: vec![crate::writer::tap::TableCell::default()],
                table_style_index: Some(table_style),
                ..crate::writer::tap::TableRow::default()
            },
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let stylesheet = document.stylesheet().unwrap();
    assert_eq!(stylesheet.styles().len(), 18);
    assert_eq!(stylesheet.get(paragraph_style).unwrap().name, "Custom Body");
    assert_eq!(
        stylesheet.get(character_style).unwrap().name,
        "Custom Emphasis"
    );
    assert_eq!(stylesheet.get(table_style).unwrap().name, "Custom Grid");
    assert_eq!(
        stylesheet.get(table_style).unwrap().kind,
        crate::StyleKind::Table
    );
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(
        paragraphs[0].properties().style_index,
        Some(paragraph_style)
    );
    assert_eq!(
        paragraphs[0].runs().unwrap()[0].properties().style_index,
        Some(character_style)
    );
    assert_eq!(
        document.tables().unwrap()[0].rows().unwrap()[0]
            .properties()
            .unwrap()
            .table_style_index,
        Some(table_style)
    );
}

#[test]
fn writes_revision_marked_style_and_author_table() {
    let timestamp = CommentDateTime {
        year: 2026,
        month: 7,
        day: 16,
        hour: 11,
        minute: 45,
        weekday: 4,
    };
    let previous_papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[0]].concat();
    let previous_chpx = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
    let mut writer = Writer::new();
    let style_index = writer
        .add_style(
            crate::writer::stylesheet::StyleDefinition::new(
                crate::StyleKind::Paragraph,
                "Tracked Body",
            )
            .with_revision(
                crate::writer::stylesheet::StyleRevision::paragraph(
                    "Style Editor",
                    previous_papx.clone(),
                    previous_chpx.clone(),
                )
                .with_timestamp(timestamp),
            ),
        )
        .unwrap();
    writer
        .add_formatted_paragraph(
            "Tracked style",
            ParagraphFormatting {
                style_index: Some(style_index),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.revision_authors(), ["Unknown", "Style Editor"]);
    let stylesheet = document.stylesheet().unwrap();
    let revision = stylesheet
        .get(style_index)
        .unwrap()
        .revision
        .as_ref()
        .unwrap();
    assert_eq!(revision.author_index, 1);
    assert_eq!(revision.author.as_deref(), Some("Style Editor"));
    assert_eq!(revision.timestamp, Some(timestamp));
    assert_eq!(
        revision.paragraph_properties.as_deref(),
        Some(previous_papx.as_slice())
    );
    assert_eq!(revision.character_properties, previous_chpx);
    assert_eq!(
        document.paragraphs().unwrap()[0].properties().style_index,
        Some(style_index)
    );
}

#[test]
fn rejects_undefined_or_wrong_kind_style_references() {
    let error_for_paragraph_style = |style_index| {
        let mut writer = Writer::new();
        writer
            .add_formatted_paragraph(
                "text",
                ParagraphFormatting {
                    style_index: Some(style_index),
                    ..ParagraphFormatting::default()
                },
            )
            .unwrap();
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
    };
    assert!(error_for_paragraph_style(14).contains("undefined DOC style index 14"));

    let mut writer = Writer::new();
    let character_style = writer
        .add_style(crate::writer::stylesheet::StyleDefinition::new(
            crate::StyleKind::Character,
            "Wrong Kind",
        ))
        .unwrap();
    writer
        .add_formatted_paragraph(
            "text",
            ParagraphFormatting {
                style_index: Some(character_style),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();
    let error = writer
        .write_to(&mut Cursor::new(Vec::new()))
        .unwrap_err()
        .to_string();
    assert!(error.contains("Character DOC style 15, expected Paragraph"));
}
