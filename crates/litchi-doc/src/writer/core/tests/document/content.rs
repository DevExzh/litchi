use super::super::support::*;

#[test]
fn test_add_paragraph() {
    let mut writer = Writer::new();
    writer.add_paragraph("Test").unwrap();
    assert_eq!(writer.paragraphs.len(), 1);
    assert_eq!(writer.paragraphs[0].runs[0].text, "Test");
}

#[test]
fn test_add_multiple_paragraphs() {
    let mut writer = Writer::new();
    writer.add_paragraph("First paragraph").unwrap();
    writer.add_paragraph("Second paragraph").unwrap();
    writer.add_paragraph("Third paragraph").unwrap();
    assert_eq!(writer.paragraphs.len(), 3);
    assert_eq!(writer.paragraphs[0].runs[0].text, "First paragraph");
    assert_eq!(writer.paragraphs[1].runs[0].text, "Second paragraph");
    assert_eq!(writer.paragraphs[2].runs[0].text, "Third paragraph");
}

#[test]
fn test_add_formatted_paragraph() {
    let mut writer = Writer::new();
    let para_fmt = ParagraphFormatting {
        alignment: Some(1), // Center
        space_before: Some(240),
        space_after: Some(120),
        ..Default::default()
    };
    writer
        .add_formatted_paragraph("Formatted text", para_fmt)
        .unwrap();
    assert_eq!(writer.paragraphs.len(), 1);
    assert_eq!(writer.paragraphs[0].runs[0].text, "Formatted text");
    assert_eq!(writer.paragraphs[0].formatting.alignment, Some(1));
}

#[test]
fn test_add_paragraph_with_character_formatting() {
    let mut writer = Writer::new();
    let char_fmt = CharacterFormatting {
        bold: Some(true),
        italic: Some(true),
        font_size: Some(24),
        ..Default::default()
    };
    let para_fmt = ParagraphFormatting::default();
    writer
        .add_paragraph_with_format("Bold italic text", char_fmt, para_fmt)
        .unwrap();
    assert_eq!(writer.paragraphs.len(), 1);
    assert_eq!(writer.paragraphs[0].runs[0].text, "Bold italic text");
    assert_eq!(writer.paragraphs[0].runs[0].formatting.bold, Some(true));
    assert_eq!(writer.paragraphs[0].runs[0].formatting.italic, Some(true));
    assert_eq!(writer.paragraphs[0].runs[0].formatting.font_size, Some(24));
}

#[test]
fn test_add_paragraph_runs() {
    let mut writer = Writer::new();
    let runs = vec![
        (
            "Bold ".to_string(),
            CharacterFormatting {
                bold: Some(true),
                ..Default::default()
            },
        ),
        (
            "Italic".to_string(),
            CharacterFormatting {
                italic: Some(true),
                ..Default::default()
            },
        ),
    ];
    writer
        .add_paragraph_runs(runs, ParagraphFormatting::default())
        .unwrap();
    assert_eq!(writer.paragraphs.len(), 1);
    assert_eq!(writer.paragraphs[0].runs.len(), 2);
    assert_eq!(writer.paragraphs[0].runs[0].text, "Bold ");
    assert_eq!(writer.paragraphs[0].runs[1].text, "Italic");
}
