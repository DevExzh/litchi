use super::super::super::support::*;

#[test]
fn supplementary_unicode_uses_utf16_code_unit_character_positions() {
    assert_eq!(utf16_code_unit_len("A😀𝄞").unwrap(), 5);

    let mut writer = Writer::new();
    writer
        .add_paragraph_runs(
            vec![
                (
                    "A😀".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "B𝄞C".to_string(),
                    CharacterFormatting {
                        italic: Some(true),
                        ..CharacterFormatting::default()
                    },
                ),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    writer.add_paragraph("After 🦀").unwrap();
    writer
        .add_paragraph("😀\u{13} HYPERLINK \"https://example.test\" \u{14}link\u{15}")
        .unwrap();
    writer.set_odd_header("Header 😀");
    writer.set_odd_footer("Footer 𝄞");
    writer.add_footnote(FootnoteEntry::new(1, "Footnote 🦀", 1));
    writer.add_endnote(FootnoteEntry::new(2, "Endnote 😀", 1));

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();

    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(paragraphs[0].text().unwrap(), "A😀B𝄞C\u{2}\u{2}");
    assert_eq!(paragraphs[1].text().unwrap(), "After 🦀");
    let fields = document.fields_table().unwrap().main_document_fields();
    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].start_cp, 21);
    assert_eq!(
        fields[0].field_type,
        crate::parts::fields::FieldType::Hyperlink
    );
    let field_text = document.fields().unwrap();
    assert_eq!(field_text.len(), 1);
    assert_eq!(
        field_text[0].instruction.trim(),
        r#"HYPERLINK "https://example.test""#
    );
    assert_eq!(field_text[0].result.as_deref(), Some("link"));
    let headers = document.headers().unwrap();
    assert_eq!(headers.len(), 1, "{headers:?}");
    assert!(
        headers
            .iter()
            .any(|header| header.text().contains("Header 😀")),
        "{headers:?}"
    );
    let footers = document.footers().unwrap();
    assert_eq!(footers.len(), 1, "{footers:?}");
    assert!(
        footers
            .iter()
            .any(|footer| footer.text().contains("Footer 𝄞")),
        "{footers:?}"
    );
    let footnotes = document.footnotes().unwrap();
    assert_eq!(footnotes[0].number, 1);
    assert!(footnotes[0].text().contains("Footnote 🦀"));
    let endnotes = document.endnotes().unwrap();
    assert_eq!(endnotes[0].number, 1);
    assert!(endnotes[0].text().contains("Endnote 😀"));
}
