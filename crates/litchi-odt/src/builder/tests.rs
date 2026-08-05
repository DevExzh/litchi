//! Regression coverage for the layered ODT builder owner.

use super::model::DocumentElement;
use super::*;
use crate::elements::table::Table;
use crate::elements::text::{Heading, Hyperlink, List, ListItem, Paragraph};
use litchi_core::Metadata;
use tempfile::tempdir;

#[test]
fn test_builder_new() {
    let builder = Builder::new();
    assert!(builder.elements.is_empty());
}

#[test]
fn test_builder_default() {
    let builder: Builder = Default::default();
    assert!(builder.elements.is_empty());
}

#[test]
fn test_add_paragraph() {
    let mut builder = Builder::new();
    builder.add_paragraph("Hello, World!").unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn hyperlink_authoring_round_trips_through_an_odt_package() {
    let mut builder = Builder::new();
    builder
        .add_hyperlink("https://example.test/a?x=1&y=2", "Example & link")
        .unwrap();
    let mut configured = Hyperlink::with_href("#bookmark", "Jump").unwrap();
    configured.set_target_frame_name("_self");
    configured.set_show(Some(crate::TextHyperlinkShow::Replace));
    configured.set_actuate(Some(crate::TextHyperlinkActuate::OnRequest));
    configured.set_title("Jump to bookmark");
    builder.add_hyperlink_element(configured).unwrap();

    let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        document.hyperlinks().unwrap(),
        vec![
            (
                "Example & link".to_string(),
                "https://example.test/a?x=1&y=2".to_string(),
            ),
            ("Jump".to_string(), "#bookmark".to_string()),
        ]
    );
    let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("xlink:type=\"simple\""));
    assert!(content.contains("office:target-frame-name=\"_self\""));
    assert!(content.contains("xlink:show=\"replace\""));
    assert!(content.contains("xlink:actuate=\"onRequest\""));
    assert!(content.contains("xlink:href=\"https://example.test/a?x=1&amp;y=2\""));

    let mut invalid = Builder::new();
    assert!(invalid.add_hyperlink("", "missing target").is_err());
    assert!(invalid.elements.is_empty());
}

#[test]
fn ruby_annotation_authoring_round_trips_through_an_odt_package() {
    let style = crate::ruby_family::Style::new(
        "RubyAbove",
        Some(crate::ruby_family::Properties {
            position: Some(crate::ruby_family::Position::Above),
            alignment: Some(crate::ruby_family::Alignment::Center),
        }),
    )
    .unwrap();
    let annotation = crate::ruby_family::Annotation::new(
        Some(style.name.clone()),
        crate::ruby_family::Base::from_text("漢").unwrap(),
        "かん",
        None,
    )
    .unwrap();

    let mut builder = Builder::new();
    builder.add_paragraph("Read ").unwrap();
    builder.add_ruby_style(style.clone()).unwrap();
    assert!(builder.add_ruby_style(style.clone()).is_err());
    builder.add_ruby_annotation(0, &annotation).unwrap();

    let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(document.ruby_styles().unwrap().styles, vec![style]);
    assert_eq!(
        document.ruby_annotations().unwrap().annotations,
        vec![annotation.clone()]
    );
    let rubies = document.rubies().unwrap();
    let ruby = rubies.first().unwrap();
    assert_eq!(ruby.base(), "漢");
    assert_eq!(ruby.text(), "かん");

    let mut invalid = Builder::new();
    assert!(invalid.add_ruby_annotation(0, &annotation).is_err());
    assert!(invalid.elements.is_empty());
}

#[test]
fn ruby_range_annotation_authoring_round_trips_through_an_odt_package() {
    let annotation = crate::ruby_family::Annotation::new(
        None,
        crate::ruby_family::Base::from_text("字").unwrap(),
        "じ",
        None,
    )
    .unwrap();
    let mut builder = Builder::new();
    builder.add_paragraph("Read 漢字").unwrap();
    let start = "Read 漢".len();
    builder
        .wrap_ruby_annotation(0, start..start + "字".len(), &annotation)
        .unwrap();

    let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        document.ruby_annotations().unwrap().annotations,
        vec![annotation]
    );

    let structured_base = crate::ruby_family::Base::from_xml_fragment(
            r#"<text:span text:style-name="Em">漢</text:span><text:span text:style-name="Strong">字</text:span>"#,
        )
        .unwrap();
    let mut builder = Builder::new();
    builder
        .add_rich_paragraph(vec![
            ("A", None),
            ("漢", Some("Em")),
            ("字", Some("Strong")),
            ("Z", None),
        ])
        .unwrap();
    let structured =
        crate::ruby_family::Annotation::new(None, structured_base, "かんじ", None).unwrap();
    builder
        .wrap_ruby_annotation(0, 1..1 + "漢字".len(), &structured)
        .unwrap();

    let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        document.ruby_annotations().unwrap().annotations,
        vec![structured]
    );
}

#[test]
fn note_authoring_round_trips_through_an_odt_package() {
    let mut note = crate::Note::new(crate::NoteClass::Footnote, "1", "First\nSecond").unwrap();
    note.set_id(Some("note-1".to_string())).unwrap();
    note.set_label(Some("*".to_string())).unwrap();

    let mut builder = Builder::new();
    builder.add_paragraph("Body text").unwrap();
    builder.add_note(0, &note).unwrap();

    let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(document.notes().unwrap(), vec![note.clone()]);
    assert_eq!(document.footnotes().unwrap(), vec![note]);
    assert!(document.endnotes().unwrap().is_empty());

    let mut invalid = Builder::new();
    invalid.add_paragraph("Only paragraph").unwrap();
    assert!(
        invalid
            .add_note(
                1,
                &crate::Note::new(crate::NoteClass::Endnote, "i", "No").unwrap()
            )
            .is_err()
    );
    assert!(invalid.notes.is_empty());
}

#[test]
fn test_add_heading() {
    let mut builder = Builder::new();
    builder.add_heading("Chapter 1", 1).unwrap();
    builder.add_heading("Section 1.1", 2).unwrap();
    assert_eq!(builder.elements.len(), 2);
}

#[test]
fn test_add_heading_invalid_level() {
    let mut builder = Builder::new();
    let result = builder.add_heading("Invalid", 0);
    assert!(result.is_err());

    let result = builder.add_heading("Invalid", 7);
    assert!(result.is_err());
}

#[test]
fn test_add_rich_paragraph() {
    let mut builder = Builder::new();
    builder
        .add_rich_paragraph(vec![
            ("This is ", None),
            ("bold", Some("Bold")),
            (" text.", None),
        ])
        .unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_add_bulleted_list() {
    let mut builder = Builder::new();
    builder
        .add_bulleted_list(vec!["Item 1", "Item 2", "Item 3"])
        .unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_add_numbered_list() {
    let mut builder = Builder::new();
    builder
        .add_numbered_list(vec!["First", "Second", "Third"])
        .unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_add_paragraph_element() {
    let mut builder = Builder::new();
    let mut para = Paragraph::new();
    para.set_text("Custom paragraph");
    builder.add_paragraph_element(para).unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_add_heading_element() {
    let mut builder = Builder::new();
    let mut heading = Heading::new(1);
    heading.set_text("Custom heading");
    builder.add_heading_element(heading).unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_add_list_element() {
    let mut builder = Builder::new();
    let mut list = List::new();
    let mut item = ListItem::new();
    item.set_text("Item");
    list.add_item(item);
    builder.add_list_element(list).unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_add_table() {
    let mut builder = Builder::new();
    let mut table = Table::new();
    table.set_name("Table1");
    builder.add_table(table).unwrap();
    assert_eq!(builder.elements.len(), 1);
}

#[test]
fn test_set_metadata() {
    let mut builder = Builder::new();
    let metadata = Metadata {
        title: Some("Test Title".to_string()),
        author: Some("Test Author".to_string()),
        subject: Some("Test Subject".to_string()),
        description: Some("Test Description".to_string()),
        keywords: Some("test, keywords".to_string()),
        ..Metadata::default()
    };
    builder.set_metadata(metadata);

    assert_eq!(builder.metadata.title, Some("Test Title".to_string()));
    assert_eq!(builder.metadata.author, Some("Test Author".to_string()));
}

#[test]
fn test_generate_content_body() {
    let mut builder = Builder::new();
    builder.add_paragraph("Paragraph 1").unwrap();
    builder.add_heading("Heading", 1).unwrap();

    let body = builder.generate_content_body();
    assert!(body.contains("Paragraph 1"));
    assert!(body.contains("Heading"));
}

#[test]
fn test_generate_content_xml() {
    let mut builder = Builder::new();
    builder.add_paragraph("Test").unwrap();

    let xml = builder.generate_content_xml();
    assert!(xml.starts_with(r#"<?xml version="1.0" encoding="UTF-8"?"#));
    assert!(xml.contains("office:document-content"));
    assert!(xml.contains("office:text"));
    assert!(xml.contains("Test"));
}

#[test]
fn test_generate_meta_xml() {
    let mut builder = Builder::new();
    builder.metadata.title = Some("My Title".to_string());
    builder.metadata.author = Some("My Author".to_string());
    builder.metadata.subject = Some("My Subject".to_string());
    builder.metadata.description = Some("My Description".to_string());
    builder.metadata.keywords = Some("my, keywords".to_string());

    let meta_xml = builder.generate_meta_xml();
    assert!(meta_xml.contains("office:document-meta"));
    assert!(meta_xml.contains("Litchi/"));
    assert!(meta_xml.contains("My Title"));
    assert!(meta_xml.contains("My Author"));
    assert!(meta_xml.contains("My Subject"));
    assert!(meta_xml.contains("My Description"));
    assert!(meta_xml.contains("my, keywords"));
}

#[test]
fn test_generate_styles_xml() {
    let builder = Builder::new();
    let styles_xml = builder.generate_styles_xml();
    assert!(styles_xml.contains("office:document-styles"));
    assert!(styles_xml.contains("L1")); // Numbered list style
}

#[test]
fn line_numbering_configuration_round_trips_through_an_odt_package() {
    let configuration = crate::line_numbering::Configuration {
        number_lines: Some(true),
        number_format: Some(crate::line_numbering::Format::UpperAlpha),
        letter_sync: Some(true),
        style_name: Some("LineNumbers".to_string()),
        increment: Some(5),
        number_position: Some(crate::line_numbering::Position::Outer),
        offset: Some(crate::line_numbering::NonNegativeLength::new("0.25in").unwrap()),
        count_empty_lines: Some(true),
        count_in_text_boxes: Some(false),
        restart_on_page: Some(true),
        separator: Some(crate::line_numbering::Separator {
            increment: Some(10),
            text: " / ".to_string(),
        }),
    };

    let mut builder = Builder::new();
    assert!(builder.line_numbering_configuration().is_none());
    builder
        .set_line_numbering_configuration(configuration.clone())
        .unwrap();
    assert_eq!(builder.line_numbering_configuration(), Some(&configuration));
    builder.clear_line_numbering_configuration();
    assert!(builder.line_numbering_configuration().is_none());
    builder
        .set_line_numbering_configuration(configuration.clone())
        .unwrap();

    let bytes = builder.build().unwrap();
    let document = crate::Document::from_bytes(bytes.clone()).unwrap();
    assert_eq!(
        document.line_numbering_configuration().unwrap(),
        Some(configuration.clone())
    );
    let package = crate::Package::from_bytes(bytes).unwrap();
    assert_eq!(
        package.line_numbering_configuration().unwrap(),
        Some(configuration)
    );
}

#[test]
fn test_build() {
    let mut builder = Builder::new();
    builder.add_paragraph("Test content").unwrap();

    let result = builder.build();
    assert!(result.is_ok());
    let bytes = result.unwrap();
    assert!(!bytes.is_empty());
    // Check it's a valid ZIP (starts with PK)
    assert_eq!(&bytes[0..2], b"PK");
}

#[test]
fn test_save() {
    let dir = tempdir().unwrap();
    let path = dir.path().join("test.odt");

    let mut builder = Builder::new();
    builder.add_paragraph("Test content").unwrap();

    let result = builder.save(&path);
    assert!(result.is_ok());
    assert!(path.exists());

    // Verify the file is a valid ZIP
    let bytes = std::fs::read(&path).unwrap();
    assert_eq!(&bytes[0..2], b"PK");
}

#[test]
fn test_chained_builder_api() {
    let mut builder = Builder::new();
    builder
        .add_heading("Title", 1)
        .unwrap()
        .add_paragraph("Introduction")
        .unwrap()
        .add_bulleted_list(vec!["Point 1", "Point 2"])
        .unwrap()
        .add_numbered_list(vec!["Step 1", "Step 2"])
        .unwrap();

    assert_eq!(builder.elements.len(), 4);
}

#[test]
fn test_document_element_clone() {
    let mut builder = Builder::new();
    builder.add_paragraph("Test").unwrap();

    let cloned = builder.elements[0].clone();
    match (&builder.elements[0], &cloned) {
        (DocumentElement::Paragraph(_), DocumentElement::Paragraph(_)) => {},
        _ => panic!("Clone mismatch"),
    }
}

#[test]
fn test_document_element_debug() {
    let mut builder = Builder::new();
    builder.add_paragraph("Test").unwrap();

    let debug_str = format!("{:?}", builder.elements[0]);
    assert!(debug_str.contains("Paragraph"));
}

#[test]
fn test_complete_document() {
    let mut builder = Builder::new();

    // Set metadata
    let metadata = Metadata {
        title: Some("Complete Document".to_string()),
        author: Some("Test Author".to_string()),
        ..Metadata::default()
    };
    builder.set_metadata(metadata);

    // Add various elements
    builder.add_heading("Title", 1).unwrap();
    builder.add_paragraph("This is a paragraph.").unwrap();
    builder
        .add_rich_paragraph(vec![
            ("Normal ", None),
            ("styled", Some("Emphasis")),
            (" text", None),
        ])
        .unwrap();
    builder
        .add_bulleted_list(vec!["Bullet 1", "Bullet 2"])
        .unwrap();
    builder
        .add_numbered_list(vec!["Number 1", "Number 2"])
        .unwrap();

    // Build and verify
    let result = builder.build();
    assert!(result.is_ok());
}

#[test]
fn test_empty_document_build() {
    let builder = Builder::new();
    let result = builder.build();
    assert!(result.is_ok());
}

#[test]
fn test_heading_levels() {
    let mut builder = Builder::new();
    for level in 1..=6 {
        builder
            .add_heading(&format!("Level {}", level), level)
            .unwrap();
    }
    assert_eq!(builder.elements.len(), 6);
}

#[test]
fn test_list_with_empty_items() {
    let mut builder = Builder::new();
    builder.add_bulleted_list(vec![]).unwrap();
    assert_eq!(builder.elements.len(), 1);
}
