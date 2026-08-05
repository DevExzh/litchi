use super::codec::*;
use super::model::*;
use crate::CommentDateTime;
use crate::SmartTagRecognizerRange;
use crate::parts::pap::{
    AutoNumberAlignment, Border as ParagraphBorder, BorderStyle as ParagraphBorderStyle,
    Borders as ParagraphBorders, DropCap, FontAlignment, FrameAnchor, FrameHeight,
    FrameHorizontalAnchor, FrameHorizontalPosition, FrameTextFlow, FrameTextWrap,
    FrameVerticalAnchor, FrameVerticalPosition, LegacyAutoNumbering, LegacyBorderPosition,
    LegacyBorderStyle, PhysicalJustification, Shading as ParagraphShading, TabAlignment, TabLeader,
    TabStop, TextBoxTightWrap,
};
use crate::parts::{list_names::ListNamesTable, list_templates::ListTemplateTable};
use crate::sprm_operations::*;
use crate::writer::bookmarks::BookmarkEntry;
use crate::writer::comments::CommentEntry;
use crate::writer::font_table::FontTableBuilder;
use crate::writer::footnotes::FootnoteEntry;
use crate::writer::numbering::{ListFormatOverride, ListStructure};
use crate::writer::revisions::{
    DisplayFieldRevision, FormattingRevision, NumberingRevision, TextRevision,
};
use crate::writer::smart_tags::SmartTagEntry;

use crate::parts::numbering::NumberFormat;
use std::io::Cursor;

#[test]
fn test_create_writer() {
    let writer = Writer::new();
    assert_eq!(writer.paragraphs.len(), 0);
    assert_eq!(writer.tables.len(), 0);
}

#[test]
fn writes_custom_styles_into_document_stylesheet() {
    let mut writer = Writer::new();
    let paragraph_style = writer
        .add_style(super::super::stylesheet::StyleDefinition::new(
            crate::StyleKind::Paragraph,
            "Custom Body",
        ))
        .unwrap();
    let character_style = writer
        .add_style(super::super::stylesheet::StyleDefinition::new(
            crate::StyleKind::Character,
            "Custom Emphasis",
        ))
        .unwrap();
    let table_style = writer
        .add_style(super::super::stylesheet::StyleDefinition::new(
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
            super::super::tap::TableRow {
                cells: vec![super::super::tap::TableCell::default()],
                table_style_index: Some(table_style),
                ..super::super::tap::TableRow::default()
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
            super::super::stylesheet::StyleDefinition::new(
                crate::StyleKind::Paragraph,
                "Tracked Body",
            )
            .with_revision(
                super::super::stylesheet::StyleRevision::paragraph(
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
        .add_style(super::super::stylesheet::StyleDefinition::new(
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

#[test]
fn test_add_table() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 3).unwrap();
    assert_eq!(idx, 0);
    assert_eq!(writer.tables[0].rows.len(), 2);
    assert_eq!(writer.tables[0].rows[0].cells.len(), 3);
}

#[test]
fn test_set_table_cell() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 2).unwrap();
    writer.set_table_cell_text(idx, 0, 0, "Cell").unwrap();
    assert_eq!(
        writer.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text,
        "Cell"
    );
}

#[test]
fn test_set_table_cell_multiple() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 2).unwrap();
    writer.set_table_cell_text(idx, 0, 0, "A").unwrap();
    writer.set_table_cell_text(idx, 0, 1, "B").unwrap();
    writer.set_table_cell_text(idx, 1, 0, "C").unwrap();
    writer.set_table_cell_text(idx, 1, 1, "D").unwrap();
    assert_eq!(
        writer.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text,
        "A"
    );
    assert_eq!(
        writer.tables[0].rows[0].cells[1].paragraphs[0].runs[0].text,
        "B"
    );
    assert_eq!(
        writer.tables[0].rows[1].cells[0].paragraphs[0].runs[0].text,
        "C"
    );
    assert_eq!(
        writer.tables[0].rows[1].cells[1].paragraphs[0].runs[0].text,
        "D"
    );
}

#[test]
fn tables_round_trip_through_both_output_paths() {
    let mut writer = Writer::new();
    writer.add_paragraph("Before table").unwrap();
    let table = writer.add_table(2, 2).unwrap();
    writer
        .set_table_cell_paragraph_runs(
            table,
            0,
            0,
            vec![(
                "A😀".to_string(),
                CharacterFormatting {
                    bold: Some(true),
                    ..CharacterFormatting::default()
                },
            )],
            ParagraphFormatting::default(),
        )
        .unwrap();
    writer
        .append_table_cell_paragraph_runs(
            table,
            0,
            0,
            vec![(
                "continued".to_string(),
                CharacterFormatting {
                    italic: Some(true),
                    ..CharacterFormatting::default()
                },
            )],
            ParagraphFormatting::default(),
        )
        .unwrap();
    writer.set_table_cell_text(table, 0, 1, "B").unwrap();
    writer
        .set_table_row_formatting(
            table,
            0,
            crate::writer::TableRow {
                cells: vec![
                    crate::writer::TableCell {
                        width: 2880,
                        merged: false,
                        vertical_merge: crate::parts::tap::VerticalMergeStatus::First,
                        vertical_alignment: crate::parts::tap::VerticalAlignment::Center,
                        text_direction: crate::parts::tap::TextDirection::TbRl,
                        fit_text: true,
                        no_wrap: true,
                        hide_mark: true,
                        borders: crate::parts::tap::CellBorders {
                            top: Some(crate::parts::tap::BorderStyle {
                                width: 8,
                                color: Some((1, 2, 3)),
                                border_type: crate::parts::tap::BorderType::Single,
                                spacing: 2,
                                shadow: true,
                                frame: false,
                            }),
                            diagonal_down: Some(crate::parts::tap::BorderStyle {
                                width: 4,
                                color: Some((10, 20, 30)),
                                border_type: crate::parts::tap::BorderType::Outset,
                                spacing: 1,
                                shadow: false,
                                frame: true,
                            }),
                            ..crate::parts::tap::CellBorders::default()
                        },
                        border_type_overrides: crate::parts::tap::CellBorderTypes::default(),
                        shading: Some(crate::parts::tap::CellShading {
                            foreground_color: Some((1, 2, 3)),
                            background_color: Some((250, 240, 230)),
                            pattern: crate::parts::tap::ShadingPattern::DarkCross,
                        }),
                        padding_top: Some(120),
                        padding_left: Some(240),
                        padding_bottom: Some(120),
                        padding_right: Some(240),
                    },
                    crate::writer::TableCell {
                        width: 5760,
                        merged: true,
                        ..crate::writer::TableCell::default()
                    },
                ],
                height: 360,
                is_header: true,
                allow_break: true,
                borders: crate::writer::TableBorders {
                    vertical: Some(crate::parts::tap::BorderStyle {
                        width: 6,
                        color: Some((40, 50, 60)),
                        border_type: crate::parts::tap::BorderType::Double,
                        spacing: 0,
                        shadow: false,
                        frame: false,
                    }),
                    ..crate::writer::TableBorders::default()
                },
                ..crate::writer::TableRow::default()
            },
        )
        .unwrap();
    writer.set_table_cell_text(table, 1, 0, "C").unwrap();
    writer
        .set_table_row_formatting(
            table,
            1,
            crate::writer::TableRow {
                cells: vec![
                    crate::writer::TableCell {
                        width: 4320,
                        merged: false,
                        vertical_merge: crate::parts::tap::VerticalMergeStatus::Merged,
                        ..crate::writer::TableCell::default()
                    },
                    crate::writer::TableCell {
                        width: 4320,
                        merged: false,
                        ..crate::writer::TableCell::default()
                    },
                ],
                height: -480,
                allow_break: false,
                ..crate::writer::TableRow::default()
            },
        )
        .unwrap();
    let second_table = writer.add_table(1, 1).unwrap();
    writer
        .set_table_cell_text(second_table, 0, 0, "Separate")
        .unwrap();

    let assert_document = |document: crate::Document| {
        let stylesheet = document.stylesheet().unwrap();
        assert_eq!(stylesheet.styles().len(), 15);
        assert_eq!(stylesheet.get(0).unwrap().name, "Normal");
        assert_eq!(stylesheet.get(10).unwrap().invariant_id, 65);
        let tables = document.tables().unwrap();
        assert_eq!(tables.len(), 2);
        assert_eq!(tables[0].row_count().unwrap(), 2);
        assert_eq!(tables[0].column_count().unwrap(), 2);
        let rows = tables[0].rows().unwrap();
        assert_eq!(rows[0].properties().unwrap().cell_count, 2);
        assert_eq!(
            rows[0].properties().unwrap().cell_boundaries,
            [0, 2880, 8640]
        );
        assert_eq!(rows[0].properties().unwrap().row_height, Some(360));
        assert!(rows[0].properties().unwrap().is_header_row);
        assert!(!rows[0].properties().unwrap().allow_row_break);
        assert_eq!(
            rows[0].cells().unwrap()[0].text().unwrap(),
            "A😀\ncontinued"
        );
        assert_eq!(rows[0].cells().unwrap()[0].paragraphs().unwrap().len(), 2);
        let cell_paragraphs = rows[0].cells().unwrap()[0].paragraphs().unwrap();
        assert_eq!(cell_paragraphs[0].runs().unwrap()[0].bold(), Some(true));
        assert_eq!(cell_paragraphs[1].runs().unwrap()[0].italic(), Some(true));
        assert_eq!(rows[0].cells().unwrap()[1].text().unwrap(), "B");
        let first_cell_properties = rows[0].cells().unwrap()[0].properties().unwrap().clone();
        assert_eq!(
            first_cell_properties.merge_status,
            crate::parts::tap::CellMergeStatus::First
        );
        assert_eq!(first_cell_properties.preferred_width.unwrap().value, 2880);
        assert_eq!(
            first_cell_properties.vertical_merge_status,
            crate::parts::tap::VerticalMergeStatus::First
        );
        assert_eq!(
            first_cell_properties.vertical_alignment,
            crate::parts::tap::VerticalAlignment::Center
        );
        assert_eq!(
            first_cell_properties.text_direction,
            crate::parts::tap::TextDirection::TbRl
        );
        assert!(first_cell_properties.fit_text);
        assert!(first_cell_properties.no_wrap);
        assert!(first_cell_properties.hide_mark);
        let top_border = first_cell_properties.borders.top.unwrap();
        assert_eq!(top_border.width, 8);
        assert_eq!(top_border.color, Some((1, 2, 3)));
        assert_eq!(
            top_border.border_type,
            crate::parts::tap::BorderType::Single
        );
        assert_eq!(top_border.spacing, 2);
        assert!(top_border.shadow);
        assert!(!top_border.frame);
        let diagonal = first_cell_properties.borders.diagonal_down.unwrap();
        assert_eq!(diagonal.color, Some((10, 20, 30)));
        assert_eq!(diagonal.border_type, crate::parts::tap::BorderType::Outset);
        assert!(diagonal.frame);
        assert_eq!(
            rows[0].properties().unwrap().border_vertical.unwrap().color,
            Some((40, 50, 60))
        );
        assert_eq!(
            first_cell_properties.shading,
            Some(crate::parts::tap::CellShading {
                foreground_color: Some((1, 2, 3)),
                background_color: Some((250, 240, 230)),
                pattern: crate::parts::tap::ShadingPattern::DarkCross,
            })
        );
        assert_eq!(
            first_cell_properties.background_color,
            Some((250, 240, 230))
        );
        assert_eq!(first_cell_properties.padding_top, Some(120));
        assert_eq!(first_cell_properties.padding_left, Some(240));
        assert_eq!(first_cell_properties.padding_bottom, Some(120));
        assert_eq!(first_cell_properties.padding_right, Some(240));
        let first_cell = &rows[0].cells().unwrap()[0];
        assert_eq!(first_cell.shading(), first_cell_properties.shading);
        assert_eq!(first_cell.shading_inherits_from_style(), Some(false));
        assert_eq!(first_cell.background_color(), Some((250, 240, 230)));
        assert_eq!(first_cell.padding_top(), Some(120));
        assert_eq!(first_cell.padding_left(), Some(240));
        assert_eq!(first_cell.padding_bottom(), Some(120));
        assert_eq!(first_cell.padding_right(), Some(240));
        assert_eq!(
            rows[0].cells().unwrap()[1]
                .properties()
                .unwrap()
                .merge_status,
            crate::parts::tap::CellMergeStatus::Merged
        );
        assert_eq!(rows[1].cells().unwrap()[0].text().unwrap(), "C");
        assert_eq!(rows[1].cells().unwrap()[1].text().unwrap(), "");
        assert_eq!(rows[1].properties().unwrap().row_height, Some(-480));
        assert!(!rows[1].properties().unwrap().allow_row_break);
        assert_eq!(
            rows[1].cells().unwrap()[0]
                .properties()
                .unwrap()
                .vertical_merge_status,
            crate::parts::tap::VerticalMergeStatus::Merged
        );
        assert_eq!(
            tables[1].rows().unwrap()[0].cells().unwrap()[0]
                .text()
                .unwrap(),
            "Separate"
        );
        assert!(document.text().unwrap().ends_with('\r'));
        let element_table = document
            .elements()
            .unwrap()
            .into_iter()
            .find_map(|element| match element {
                crate::Element::Table(table) => Some(table),
                crate::Element::Paragraph(_) => None,
            })
            .unwrap();
        assert_eq!(
            element_table.properties().unwrap().cell_boundaries,
            [0, 2880, 8640]
        );
    };

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    assert_document(package.document().unwrap());

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-table-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    assert_document(package.document().unwrap());
    std::fs::remove_file(path).unwrap();
}

#[test]
fn file_and_seekable_outputs_are_byte_identical() {
    let mut writer = Writer::new();
    writer.set_property("Title", "Canonical output");
    writer.add_paragraph("One output assembly path").unwrap();

    let mut memory = Cursor::new(Vec::new());
    writer.write_to(&mut memory).unwrap();

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-output-equivalence-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let file = std::fs::read(&path).unwrap();
    std::fs::remove_file(path).unwrap();

    assert_eq!(file, memory.into_inner());
}

#[test]
fn test_set_property() {
    let mut writer = Writer::new();
    writer.set_property("Title", "Test Document");
    writer.set_property("Author", "Test Author");
    assert_eq!(
        writer.properties.get("Title"),
        Some(&"Test Document".to_string())
    );
    assert_eq!(
        writer.properties.get("Author"),
        Some(&"Test Author".to_string())
    );
}

#[test]
fn test_headers_and_footers() {
    let mut writer = Writer::new();
    writer.set_odd_header("Odd Header");
    writer.set_even_header("Even Header");
    writer.set_first_header("First Header");
    writer.set_odd_footer("Odd Footer");
    writer.set_even_footer("Even Footer");
    writer.set_first_footer("First Footer");
    assert_eq!(
        writer.header_odd.as_ref().unwrap()[0].runs[0].0,
        "Odd Header"
    );
    assert_eq!(
        writer.header_even.as_ref().unwrap()[0].runs[0].0,
        "Even Header"
    );
    assert_eq!(
        writer.header_first.as_ref().unwrap()[0].runs[0].0,
        "First Header"
    );
    assert_eq!(
        writer.footer_odd.as_ref().unwrap()[0].runs[0].0,
        "Odd Footer"
    );
    assert_eq!(
        writer.footer_even.as_ref().unwrap()[0].runs[0].0,
        "Even Footer"
    );
    assert_eq!(
        writer.footer_first.as_ref().unwrap()[0].runs[0].0,
        "First Footer"
    );
}

#[test]
fn test_footnotes() {
    let mut writer = Writer::new();
    let entry = FootnoteEntry::new(0u32, "This is a footnote", 1u16);
    writer.add_footnote(entry);
    assert_eq!(writer.footnotes.len(), 1);
    assert_eq!(writer.footnotes[0].text, "This is a footnote");
}

#[test]
fn test_endnotes() {
    let mut writer = Writer::new();
    let entry = FootnoteEntry::new(0u32, "This is an endnote", 1u16);
    writer.add_endnote(entry);
    assert_eq!(writer.endnotes.len(), 1);
    assert_eq!(writer.endnotes[0].text, "This is an endnote");
}

#[test]
fn test_write_to_memory() {
    let mut writer = Writer::new();
    writer.add_paragraph("Test paragraph").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    let result = writer.write_to(&mut cursor);
    assert!(result.is_ok());
    assert!(!cursor.into_inner().is_empty());
}

#[test]
fn test_empty_document_write() {
    let mut writer = Writer::new();
    let mut cursor = Cursor::new(Vec::new());
    let result = writer.write_to(&mut cursor);
    assert!(result.is_ok());
    let data = cursor.into_inner();
    assert!(!data.is_empty());
    let mut package = crate::Package::from_reader(Cursor::new(data)).unwrap();
    assert_eq!(package.document().unwrap().text().unwrap(), "\r");
}

#[test]
fn test_character_formatting_default() {
    let fmt = CharacterFormatting::default();
    assert!(fmt.bold.is_none());
    assert!(fmt.italic.is_none());
    assert!(fmt.underline.is_none());
    assert!(fmt.font_size.is_none());
}

#[test]
fn test_paragraph_formatting_default() {
    let fmt = ParagraphFormatting::default();
    assert!(fmt.alignment.is_none());
    assert!(fmt.left_indent.is_none());
    assert!(fmt.right_indent.is_none());
    assert!(fmt.space_before.is_none());
    assert!(fmt.space_after.is_none());
}

#[test]
fn test_line_spacing_default() {
    let ls = LineSpacing::default();
    assert_eq!(ls, LineSpacing::single());
    assert_eq!(ls.dya_line, 240);
    assert!(ls.is_multiple);
}

#[test]
fn encodes_physical_and_logical_paragraph_justification_compatibly() {
    let physical = build_papx_grpprl(&ParagraphFormatting {
        physical_justification: Some(PhysicalJustification::HighCompression),
        ..ParagraphFormatting::default()
    });
    assert_eq!(
        physical,
        [SPRM_P_JC.to_le_bytes().as_slice(), &[5]].concat()
    );

    let logical = build_papx_grpprl(&ParagraphFormatting {
        alignment: Some(7),
        ..ParagraphFormatting::default()
    });
    assert_eq!(
        logical,
        [
            SPRM_P_JC.to_le_bytes().as_slice(),
            &[5],
            SPRM_P_JC_LOGICAL.to_le_bytes().as_slice(),
            &[7],
        ]
        .concat()
    );
}

#[test]
fn encodes_paragraph_revision_save_id() {
    let encoded = build_papx_grpprl(&ParagraphFormatting {
        revision_save_id: Some(0x1122_3344),
        ..ParagraphFormatting::default()
    });
    assert_eq!(
        encoded,
        [
            SPRM_P_RSID.to_le_bytes().as_slice(),
            0x1122_3344u32.to_le_bytes().as_slice(),
        ]
        .concat()
    );
}

#[test]
fn preserves_ordered_paragraph_property_revision_state() {
    let formatting = ParagraphFormatting {
        right_indent: Some(200),
        preserved_properties_for_revision: Some(Box::new(ParagraphFormatting {
            left_indent: Some(100),
            ..ParagraphFormatting::default()
        })),
        ..ParagraphFormatting::default()
    };
    let grpprl = build_revision_papx_grpprl(&formatting, None).unwrap();
    let properties = crate::parts::pap::ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.indent_left, Some(100));
    assert_eq!(properties.indent_right, Some(200));
    let previous = properties.preserved_properties_for_revision.unwrap();
    assert_eq!(previous.indent_left, Some(100));
    assert_eq!(previous.indent_right, None);

    let mut writer = Writer::new();
    writer
        .add_formatted_paragraph("Tracked formatting", formatting)
        .unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    let properties = paragraphs[0].properties();
    assert_eq!(properties.indent_left, Some(100));
    assert_eq!(properties.indent_right, Some(200));
    let previous = properties
        .preserved_properties_for_revision
        .as_ref()
        .unwrap();
    assert_eq!(previous.indent_left, Some(100));
    assert_eq!(previous.indent_right, None);
}

#[test]
fn test_line_spacing_constructors_and_sprm_encoding() {
    let cases = [
        (LineSpacing::single(), [0xf0, 0x00, 0x01, 0x00]),
        (LineSpacing::one_and_half(), [0x68, 0x01, 0x01, 0x00]),
        (LineSpacing::double(), [0xe0, 0x01, 0x01, 0x00]),
        (
            LineSpacing::multiple_240ths(300).unwrap(),
            [0x2c, 0x01, 0x01, 0x00],
        ),
        (
            LineSpacing::at_least_twips(240).unwrap(),
            [0xf0, 0x00, 0x00, 0x00],
        ),
        (
            LineSpacing::exact_twips(240).unwrap(),
            [0x10, 0xff, 0x00, 0x00],
        ),
    ];

    for (line_spacing, operand) in cases {
        let formatting = ParagraphFormatting {
            line_spacing: Some(line_spacing),
            ..ParagraphFormatting::default()
        };
        let mut expected = SPRM_P_DYA_LINE.to_le_bytes().to_vec();
        expected.extend_from_slice(&operand);
        assert_eq!(build_papx_grpprl(&formatting), expected);
    }

    assert!(LineSpacing::multiple_240ths(0).is_err());
    assert!(LineSpacing::multiple_240ths(31_681).is_err());
    assert!(LineSpacing::at_least_twips(0).is_err());
    assert!(LineSpacing::at_least_twips(31_681).is_err());
    assert!(LineSpacing::exact_twips(0).is_err());
    assert!(LineSpacing::exact_twips(31_681).is_err());
}

#[test]
fn test_paragraph_formatting_writer_reader_round_trip() {
    let legacy_autonumbering = LegacyAutoNumbering {
        number_format: NumberFormat::RussianUpper,
        alignment: AutoNumberAlignment::Justified,
        include_previous_levels: true,
        hanging_indent: true,
        set_bold: true,
        set_italic: true,
        set_small_caps: true,
        set_caps: true,
        set_strike: true,
        set_underline: true,
        prefix_space: true,
        bold: true,
        italic: true,
        small_caps: true,
        caps: true,
        strike: true,
        underline: 3,
        color_index: 6,
        font_index: 4,
        font_size_half_points: 24,
        start_at: 3,
        indent_twips: -360,
        space_twips: 180,
        number_once_per_cell: true,
        number_across_cells: false,
        restart_each_section: true,
        prefix: "§(".to_string(),
        suffix: ")".to_string(),
    };
    let mut writer = Writer::new();
    writer
        .add_formatted_paragraph(
            "Exactly spaced",
            ParagraphFormatting {
                alignment: Some(1),
                left_indent_chars: Some(250),
                right_indent_chars: Some(-125),
                first_line_indent_chars: Some(-50),
                space_before: Some(120),
                space_after: Some(240),
                no_line_numbering: Some(true),
                space_before_lines: Some(-20),
                space_after_lines: Some(31_680),
                space_before_auto: Some(true),
                space_after_auto: Some(true),
                open_table_cell_mark: Some(true),
                frame_anchor_locked: Some(true),
                kinsoku: Some(true),
                word_wrap: Some(false),
                overflow_punctuation: Some(true),
                top_line_punctuation: Some(true),
                auto_space_east_asian_latin: Some(true),
                auto_space_east_asian_numbers: Some(false),
                font_alignment: Some(FontAlignment::Bottom),
                frame_text_flow: Some(FrameTextFlow {
                    vertical: true,
                    backwards: true,
                    rotate_font: false,
                }),
                frame_horizontal_position: Some(FrameHorizontalPosition::Right),
                frame_vertical_position: Some(FrameVerticalPosition::Offset(300)),
                frame_width: Some(1_440),
                frame_anchor: Some(FrameAnchor {
                    vertical: FrameVerticalAnchor::Paragraph,
                    horizontal: FrameHorizontalAnchor::Margin,
                }),
                in_table: Some(false),
                table_terminating_paragraph: Some(false),
                frame_text_wrap: Some(FrameTextWrap::Through),
                frame_height: Some(FrameHeight {
                    height_twips: 720,
                    minimum: true,
                }),
                frame_horizontal_text_distance: Some(480),
                frame_vertical_text_distance: Some(240),
                drop_cap: Some(DropCap {
                    kind: crate::parts::pap::DropCapType::Margin,
                    lines: 3,
                }),
                no_auto_hyphenation: Some(true),
                side_by_side: Some(true),
                use_page_setup_settings: Some(true),
                adjust_right_indent: Some(false),
                no_allow_overlap: Some(true),
                contextual_spacing: Some(true),
                mirror_indents: Some(true),
                text_box_tight_wrap: Some(TextBoxTightWrap::FirstAndLastLine),
                borders: ParagraphBorders {
                    top: Some(ParagraphBorder {
                        style: ParagraphBorderStyle::Inset,
                        width: 12,
                        color: Some((0x11, 0x22, 0x33)),
                        spacing: 7,
                        shadow: true,
                        frame: true,
                    }),
                    left: Some(ParagraphBorder {
                        style: ParagraphBorderStyle::DoubleWave,
                        width: 8,
                        color: None,
                        spacing: 4,
                        shadow: false,
                        frame: false,
                    }),
                    ..ParagraphBorders::default()
                },
                legacy_border_style: Some(LegacyBorderStyle::Shadow),
                legacy_border_position: Some(LegacyBorderPosition::LeftBar),
                shading: Some(ParagraphShading {
                    foreground_color: Some((1, 2, 3)),
                    background_color: None,
                    pattern: crate::parts::tap::ShadingPattern::DiagonalCross,
                }),
                line_spacing: Some(LineSpacing::exact_twips(360).unwrap()),
                tab_stops_to_delete: vec![720],
                tab_stops_to_add: vec![
                    TabStop {
                        position: 1_440,
                        alignment: TabAlignment::List,
                        leader: TabLeader::DefaultLeader,
                    },
                    TabStop {
                        position: 720,
                        alignment: TabAlignment::Decimal,
                        leader: TabLeader::Dots,
                    },
                ],
                ilvl: Some(8),
                ilfo: Some(1),
                legacy_autonumbering: Some(legacy_autonumbering.clone()),
                revision_save_id: Some(0x1122_3344),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();
    writer
        .add_formatted_paragraph(
            "Double spaced",
            ParagraphFormatting {
                line_spacing: Some(LineSpacing::double()),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();
    writer
        .add_formatted_paragraph(
            "Thai distributed",
            ParagraphFormatting {
                alignment: Some(9),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package =
        super::super::super::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();

    assert_eq!(paragraphs.len(), 3);
    assert_eq!(paragraphs[0].text().unwrap(), "Exactly spaced");
    assert_eq!(
        paragraphs[0].properties().justification,
        crate::parts::pap::Justification::Center
    );
    assert_eq!(paragraphs[0].properties().space_before, Some(120));
    assert_eq!(paragraphs[0].properties().space_after, Some(240));
    assert!(paragraphs[0].properties().no_line_numbering);
    assert_eq!(paragraphs[0].properties().list_level, Some(8));
    assert_eq!(paragraphs[0].properties().list_format_override, Some(1));
    assert_eq!(
        paragraphs[0].properties().legacy_autonumbering,
        Some(legacy_autonumbering)
    );
    assert_eq!(
        paragraphs[0].properties().revision_save_id,
        Some(0x1122_3344)
    );
    assert_eq!(
        paragraphs[0].properties().tab_stops,
        vec![
            TabStop {
                position: 720,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::Dots,
            },
            TabStop {
                position: 1_440,
                alignment: TabAlignment::List,
                leader: TabLeader::DefaultLeader,
            },
        ]
    );
    assert_eq!(paragraphs[0].properties().indent_left_chars, Some(250));
    assert_eq!(paragraphs[0].properties().indent_right_chars, Some(-125));
    assert_eq!(
        paragraphs[0].properties().indent_first_line_chars,
        Some(-50)
    );
    assert_eq!(paragraphs[0].properties().space_before_lines, Some(-20));
    assert_eq!(paragraphs[0].properties().space_after_lines, Some(31_680));
    assert!(paragraphs[0].properties().space_before_auto);
    assert!(paragraphs[0].properties().space_after_auto);
    assert!(paragraphs[0].properties().open_table_cell_mark);
    assert!(paragraphs[0].properties().locked);
    assert!(paragraphs[0].properties().kinsoku);
    assert!(!paragraphs[0].properties().word_wrap);
    assert!(paragraphs[0].properties().overflow_punct);
    assert!(paragraphs[0].properties().top_line_punct);
    assert!(paragraphs[0].properties().auto_space_de);
    assert!(!paragraphs[0].properties().auto_space_dn);
    assert_eq!(
        paragraphs[0].properties().font_align,
        Some(FontAlignment::Bottom)
    );
    assert_eq!(
        paragraphs[0].properties().frame_text_flow,
        Some(FrameTextFlow {
            vertical: true,
            backwards: true,
            rotate_font: false,
        })
    );
    assert_eq!(
        paragraphs[0].properties().frame_horizontal_position,
        Some(FrameHorizontalPosition::Right)
    );
    assert_eq!(
        paragraphs[0].properties().frame_vertical_position,
        Some(FrameVerticalPosition::Offset(300))
    );
    assert_eq!(paragraphs[0].properties().frame_width, Some(1_440));
    assert_eq!(
        paragraphs[0].properties().frame_anchor,
        Some(FrameAnchor {
            vertical: FrameVerticalAnchor::Paragraph,
            horizontal: FrameHorizontalAnchor::Margin,
        })
    );
    assert!(!paragraphs[0].properties().in_table);
    assert!(!paragraphs[0].properties().is_table_row_end);
    assert_eq!(
        paragraphs[0].properties().text_wrap,
        Some(FrameTextWrap::Through)
    );
    assert_eq!(
        paragraphs[0].properties().frame_height,
        Some(FrameHeight {
            height_twips: 720,
            minimum: true,
        })
    );
    assert_eq!(paragraphs[0].properties().dxa_from_text, Some(480));
    assert_eq!(paragraphs[0].properties().dya_from_text, Some(240));
    assert_eq!(
        paragraphs[0].properties().drop_cap,
        Some(DropCap {
            kind: crate::parts::pap::DropCapType::Margin,
            lines: 3,
        })
    );
    assert!(paragraphs[0].properties().no_auto_hyph);
    assert!(paragraphs[0].properties().side_by_side);
    assert_eq!(
        paragraphs[0].properties().use_page_setup_settings,
        Some(true)
    );
    assert_eq!(paragraphs[0].properties().adjust_right_indent, Some(false));
    assert!(paragraphs[0].properties().no_allow_overlap);
    assert!(paragraphs[0].properties().contextual_spacing);
    assert!(paragraphs[0].properties().mirror_indents);
    assert_eq!(
        paragraphs[0].properties().text_box_tight_wrap,
        Some(TextBoxTightWrap::FirstAndLastLine)
    );
    assert_eq!(
        paragraphs[0].properties().borders,
        ParagraphBorders {
            top: Some(ParagraphBorder {
                style: ParagraphBorderStyle::Inset,
                width: 12,
                color: Some((0x11, 0x22, 0x33)),
                spacing: 7,
                shadow: true,
                frame: true,
            }),
            left: Some(ParagraphBorder {
                style: ParagraphBorderStyle::DoubleWave,
                width: 8,
                color: None,
                spacing: 4,
                shadow: false,
                frame: false,
            }),
            ..ParagraphBorders::default()
        }
    );
    assert_eq!(
        paragraphs[0].properties().legacy_border_style,
        Some(LegacyBorderStyle::Shadow)
    );
    assert_eq!(
        paragraphs[0].properties().legacy_border_position,
        Some(LegacyBorderPosition::LeftBar)
    );
    assert_eq!(
        paragraphs[0].properties().shading,
        Some(ParagraphShading {
            foreground_color: Some((1, 2, 3)),
            background_color: None,
            pattern: crate::parts::tap::ShadingPattern::DiagonalCross,
        })
    );
    assert_eq!(paragraphs[0].properties().line_spacing, Some(-360));
    assert_eq!(
        paragraphs[0].properties().line_spacing_type,
        crate::parts::pap::LineSpacingType::Exactly
    );
    assert_eq!(paragraphs[1].text().unwrap(), "Double spaced");
    assert_eq!(paragraphs[1].properties().line_spacing, Some(480));
    assert_eq!(
        paragraphs[1].properties().line_spacing_type,
        crate::parts::pap::LineSpacingType::Double
    );
    assert_eq!(
        paragraphs[2].properties().justification,
        crate::parts::pap::Justification::ThaiDistributed
    );
}

#[test]
fn rejects_invalid_current_paragraph_layout_values() {
    for formatting in [
        ParagraphFormatting {
            alignment: Some(10),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            outline_level: Some(10),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            ilvl: Some(9),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            ilfo: Some(0x07FF),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            preserved_properties_for_revision: Some(Box::new(ParagraphFormatting {
                left_indent: Some(31_681),
                ..ParagraphFormatting::default()
            })),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            preserved_properties_for_revision: Some(Box::new(ParagraphFormatting {
                preserved_properties_for_revision: Some(Box::new(ParagraphFormatting::default())),
                ..ParagraphFormatting::default()
            })),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            legacy_autonumbering: Some(LegacyAutoNumbering {
                prefix: "x".repeat(33),
                ..LegacyAutoNumbering::default()
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            legacy_autonumbering: Some(LegacyAutoNumbering {
                underline: 8,
                ..LegacyAutoNumbering::default()
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            legacy_autonumbering: Some(LegacyAutoNumbering {
                color_index: 17,
                ..LegacyAutoNumbering::default()
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            legacy_autonumbering: Some(LegacyAutoNumbering {
                indent_twips: i16::MIN,
                ..LegacyAutoNumbering::default()
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            legacy_autonumbering: Some(LegacyAutoNumbering {
                space_twips: 31_681,
                ..LegacyAutoNumbering::default()
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            left_indent: Some(31_681),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            space_after: Some(31_681),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            line_spacing: Some(LineSpacing {
                dya_line: i16::MIN,
                is_multiple: false,
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            tab_stops_to_add: vec![TabStop {
                position: 31_681,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }],
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            tab_stops_to_delete: vec![720, 720],
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_text_flow: Some(FrameTextFlow {
                vertical: false,
                backwards: true,
                rotate_font: false,
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_height: Some(FrameHeight {
                height_twips: 32_768,
                minimum: false,
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            drop_cap: Some(DropCap {
                kind: crate::parts::pap::DropCapType::Regular,
                lines: 0,
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_horizontal_text_distance: Some(-1),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_horizontal_position: Some(FrameHorizontalPosition::Offset(i16::MAX)),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_horizontal_position: Some(FrameHorizontalPosition::Offset(-5)),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_width: Some(31_681),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            table_terminating_paragraph: Some(true),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            frame_text_flow: Some(FrameTextFlow {
                vertical: true,
                backwards: false,
                rotate_font: false,
            }),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            space_before_lines: Some(-21),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            space_after_lines: Some(31_681),
            ..ParagraphFormatting::default()
        },
        ParagraphFormatting {
            borders: ParagraphBorders {
                top: Some(ParagraphBorder {
                    style: ParagraphBorderStyle::Single,
                    width: 8,
                    color: None,
                    spacing: 32,
                    shadow: false,
                    frame: false,
                }),
                ..ParagraphBorders::default()
            },
            ..ParagraphFormatting::default()
        },
    ] {
        assert!(build_revision_papx_grpprl(&formatting, None).is_err());
    }
}

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
    let mut package =
        super::super::super::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
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

#[test]
fn comments_round_trip_with_other_subdocuments() {
    let mut writer = Writer::new();
    writer.add_paragraph("Main 😀").unwrap();
    writer.add_footnote(FootnoteEntry::new(0, "Footnote", 1));
    writer.add_comment(
        CommentEntry::new(1, "Review 🦀", "Alice 😀", "A😀")
            .with_range(2, 6)
            .with_extended_metadata(crate::CommentExtendedMetadata {
                modified_at: Some(CommentDateTime {
                    year: 2026,
                    month: 7,
                    day: 15,
                    hour: 14,
                    minute: 30,
                    weekday: 3,
                }),
                depth: 0,
                parent_index: None,
                is_ink: false,
            }),
    );
    writer.add_comment(
        CommentEntry::new(3, "Second review", "Alice 😀", "AL")
            .with_range(0, 7)
            .with_extended_metadata(crate::CommentExtendedMetadata {
                modified_at: None,
                depth: 1,
                parent_index: Some(0),
                is_ink: true,
            }),
    );
    writer.add_endnote(FootnoteEntry::new(2, "Endnote", 1));
    writer.set_odd_header("Header");

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();

    assert_eq!(document.footnotes().unwrap().len(), 1);
    assert_eq!(document.headers().unwrap().len(), 1);
    assert_eq!(document.endnotes().unwrap().len(), 1);
    let comments = document.comments().unwrap();
    assert_eq!(comments.len(), 2);
    assert_eq!(comments[0].author, "Alice 😀");
    assert_eq!(comments[0].initials, "A😀");
    assert_eq!(comments[0].bookmark_tag, Some(0));
    assert_eq!(
        (comments[0].range_start, comments[0].range_end),
        (Some(2), Some(6))
    );
    let first_metadata = comments[0].extended_metadata.unwrap();
    assert_eq!(first_metadata.depth, 0);
    assert_eq!(first_metadata.parent_index, None);
    assert_eq!(
        first_metadata.modified_at,
        Some(CommentDateTime {
            year: 2026,
            month: 7,
            day: 15,
            hour: 14,
            minute: 30,
            weekday: 3,
        })
    );
    assert!(comments[0].text().contains("Review 🦀"));
    assert_eq!(comments[0].paragraphs().unwrap().len(), 1);
    assert_eq!(comments[1].author, "Alice 😀");
    assert_eq!(comments[1].initials, "AL");
    assert_eq!(
        (comments[1].range_start, comments[1].range_end),
        (Some(0), Some(7))
    );
    assert_eq!(comments[1].extended_metadata.unwrap().parent_index, Some(0));
    assert!(comments[1].extended_metadata.unwrap().is_ink);
    assert!(comments[1].text().contains("Second review"));

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-comments-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    let comments = package.document().unwrap().comments().unwrap();
    assert_eq!(comments.len(), 2);
    assert_eq!(
        (comments[0].range_start, comments[0].range_end),
        (Some(2), Some(6))
    );
    assert_eq!(comments[1].extended_metadata.unwrap().parent_index, Some(0));
    std::fs::remove_file(path).unwrap();
}

#[test]
fn rejects_comment_metadata_outside_binary_limits() {
    let mut writer = Writer::new();
    writer.add_paragraph("Main").unwrap();
    writer.add_comment(CommentEntry::new(0, "Body", "Author", "0123456789"));

    let error = writer.write_to(&mut Cursor::new(Vec::new())).unwrap_err();
    assert!(error.to_string().contains("at most nine"));
}

#[test]
fn rejects_invalid_comment_ranges_timestamps_and_reply_trees() {
    let write_error = |entry: CommentEntry| {
        let mut writer = Writer::new();
        writer.add_paragraph("Main").unwrap();
        writer.add_comment(entry);
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
    };

    let error = write_error(CommentEntry::new(0, "Body", "Author", "A").with_range(4, 2));
    assert!(error.contains("range must be ordered"));

    let error = write_error(
        CommentEntry::new(0, "Body", "Author", "A").with_extended_metadata(
            crate::CommentExtendedMetadata {
                modified_at: Some(CommentDateTime {
                    year: 2026,
                    month: 13,
                    day: 1,
                    hour: 0,
                    minute: 0,
                    weekday: 0,
                }),
                depth: 0,
                parent_index: None,
                is_ink: false,
            },
        ),
    );
    assert!(error.contains("DTTM"));

    let error = write_error(
        CommentEntry::new(0, "Body", "Author", "A").with_extended_metadata(
            crate::CommentExtendedMetadata {
                modified_at: None,
                depth: 1,
                parent_index: Some(0),
                is_ink: false,
            },
        ),
    );
    assert!(error.contains("pre-order"));
}

#[test]
fn standard_bookmarks_round_trip_through_both_output_paths() {
    let mut writer = Writer::new();
    writer.add_paragraph("Main text").unwrap();
    writer.add_bookmark(BookmarkEntry::new("Outer", 2, 5));
    writer.add_bookmark(
        BookmarkEntry::new("_Cell", 0, 8)
            .with_native_export(false)
            .with_column_range(1, 3),
    );

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let bookmarks = package.document().unwrap().bookmarks().unwrap();
    assert_eq!(bookmarks.len(), 2);
    assert_eq!(bookmarks[0].name, "_Cell");
    assert_eq!((bookmarks[0].start, bookmarks[0].end), (0, 8));
    assert_eq!(bookmarks[0].column_range, Some((1, 3)));
    assert!(!bookmarks[0].is_native);
    assert_eq!(bookmarks[1].name, "Outer");
    assert_eq!((bookmarks[1].start, bookmarks[1].end), (2, 5));

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-bookmarks-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    assert_eq!(package.document().unwrap().bookmarks().unwrap(), bookmarks);
    std::fs::remove_file(path).unwrap();
}

#[test]
fn smart_tags_round_trip_through_both_output_paths() {
    let mut writer = Writer::new();
    writer.add_paragraph("abcdefghijklmnopqrst").unwrap();
    writer.add_smart_tag(
        SmartTagEntry::new(0, 10, "urn:example:geo", "place")
            .with_origin(crate::SmartTagOrigin::ExternalRecognizer)
            .with_native_export(true)
            .with_property("city", "東京"),
    );
    writer.add_smart_tag(
        SmartTagEntry::new(5, 15, "urn:example:geo", "place")
            .with_sub_entity(true)
            .with_property("city", "Paris"),
    );
    writer.add_smart_tag(SmartTagEntry::new(5, 5, "urn:example:point", "cursor"));
    writer.add_smart_tag_recognizer_range(SmartTagRecognizerRange {
        start: 0,
        end: 5,
        state: crate::SmartTagRecognizerState::Dirty,
    });
    writer.add_smart_tag_recognizer_range(SmartTagRecognizerRange {
        start: 5,
        end: 20,
        state: crate::SmartTagRecognizerState::Clean,
    });

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    for index in [114usize, 115, 117, 118, 132] {
        assert!(document.fib().get_table_pointer(index).unwrap().1 > 0);
    }
    let smart_tags = document.smart_tags().unwrap().clone();
    assert_eq!(smart_tags.tags.len(), 3);
    assert_eq!(smart_tags.store.as_ref().unwrap().types.len(), 2);
    assert_eq!(
        smart_tags.tags[0].info.origin,
        crate::SmartTagOrigin::ExternalRecognizer
    );
    assert!(smart_tags.tags[0].is_native);
    assert_eq!(
        smart_tags
            .store
            .as_ref()
            .unwrap()
            .resolve_property(smart_tags.tags[0].property_bag.properties[0]),
        Some(("city", "東京"))
    );
    assert_eq!(
        (smart_tags.tags[1].start_depth, smart_tags.tags[1].end_depth),
        (3, 0)
    );
    assert_eq!(smart_tags.recognizer_ranges.len(), 2);

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-smart-tags-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    assert_eq!(package.document().unwrap().smart_tags(), Some(&smart_tags));
    std::fs::remove_file(path).unwrap();
}

#[test]
fn rejects_invalid_standard_bookmarks() {
    let write_error = |entries: Vec<BookmarkEntry>| {
        let mut writer = Writer::new();
        writer.add_paragraph("Main").unwrap();
        for entry in entries {
            writer.add_bookmark(entry);
        }
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
    };
    assert!(write_error(vec![BookmarkEntry::new("", 0, 1)]).contains("names"));
    assert!(
        write_error(vec![
            BookmarkEntry::new("Same", 0, 1),
            BookmarkEntry::new("Same", 1, 2),
        ])
        .contains("unique")
    );
    assert!(write_error(vec![BookmarkEntry::new("Range", 4, 2)]).contains("range"));
    assert!(
        write_error(vec![
            BookmarkEntry::new("Column", 0, 1).with_column_range(3, 2)
        ])
        .contains("column")
    );
}

#[test]
fn tracked_text_revisions_round_trip_through_both_output_paths() {
    let timestamp = CommentDateTime {
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

#[test]
fn preserves_ordered_character_property_revision_state() {
    let formatting = CharacterFormatting {
        italic: Some(true),
        preserved_properties_for_revision: Some(Box::new(CharacterFormatting {
            bold: Some(true),
            ..CharacterFormatting::default()
        })),
        ..CharacterFormatting::default()
    };
    let mut fonts = FontTableBuilder::new();
    let grpprl = build_revision_chpx_grpprl(&formatting, &mut fonts, None).unwrap();
    let properties = crate::parts::chp::CharacterProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.is_bold, Some(true));
    assert_eq!(properties.is_italic, Some(true));
    let previous = properties.preserved_properties_for_revision.unwrap();
    assert_eq!(previous.is_bold, Some(true));
    assert_eq!(previous.is_italic, None);

    let mut writer = Writer::new();
    writer
        .add_paragraph_runs(
            vec![("Tracked".to_string(), formatting)],
            ParagraphFormatting::default(),
        )
        .unwrap();
    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    let runs = paragraphs[0].runs().unwrap();
    let properties = runs[0].properties();
    assert_eq!(properties.is_bold, Some(true));
    assert_eq!(properties.is_italic, Some(true));
    let previous = properties
        .preserved_properties_for_revision
        .as_ref()
        .unwrap();
    assert_eq!(previous.is_bold, Some(true));
    assert_eq!(previous.is_italic, None);
}

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

#[test]
fn list_tables_round_trip_through_fib_indices() {
    let mut writer = Writer::new();
    let mut list = ListStructure::new(42);
    let mut level = crate::writer::numbering::ListLevel::new(3, NumberFormat::Decimal);
    level.number_text = "%1.😀".to_string();
    list.add_level(level);
    writer.add_list(list);
    writer.add_list_override(ListFormatOverride::new(42, 1));
    writer.set_list_names(ListNamesTable::try_new(vec!["Outline".to_string()]).unwrap());
    let template = crate::ListTemplateCode::BuiltIn {
        format: crate::BuiltInListTemplate::ArabicPeriod,
        language: crate::ListTemplateLanguageId::new(0x0409),
    };
    writer.set_list_templates(ListTemplateTable::try_new(vec![Some([template; 9])]).unwrap());
    writer
        .add_paragraph_runs(
            vec![("List item".to_string(), CharacterFormatting::default())],
            ParagraphFormatting {
                ilvl: Some(0),
                ilfo: Some(1),
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    let tables = document.list_tables().unwrap();
    assert_eq!(tables.structures().len(), 1);
    assert_eq!(tables.overrides().len(), 1);
    assert_eq!(tables.structures()[0].levels[0].number_text, "%1.😀");
    assert_eq!(document.list_names().unwrap().name(0), Some("Outline"));
    assert_eq!(
        document.list_templates().unwrap().get(0).unwrap(),
        &[template; 9]
    );

    let paragraphs = document.paragraphs().unwrap();
    let info = document.paragraph_list_info(&paragraphs[0]).unwrap();
    assert_eq!(info.start_at, 3);
    assert_eq!(info.number_text, "%1.😀");
}

#[test]
fn test_add_table_invalid_dimensions() {
    let mut writer = Writer::new();
    assert!(writer.add_table(0, 3).is_err());
    assert!(writer.add_table(2, 0).is_err());
    assert!(writer.add_table(0, 0).is_err());
    assert!(writer.add_table(1, 64).is_err());
}

#[test]
fn test_set_table_cell_invalid_indices() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 2).unwrap();
    assert!(writer.set_table_cell_text(idx, 2, 0, "Invalid").is_err());
    assert!(writer.set_table_cell_text(idx, 0, 2, "Invalid").is_err());
    assert!(writer.set_table_cell_text(999, 0, 0, "Invalid").is_err());
}

#[test]
fn rejects_invalid_table_row_formatting() {
    let mut writer = Writer::new();
    let table = writer.add_table(2, 2).unwrap();
    let one_cell = crate::writer::TableRow {
        cells: vec![crate::writer::TableCell {
            width: 1000,
            merged: false,
            ..crate::writer::TableCell::default()
        }],
        ..crate::writer::TableRow::default()
    };
    assert!(writer.set_table_row_formatting(table, 0, one_cell).is_err());

    let invalid_merge = crate::writer::TableRow {
        cells: vec![
            crate::writer::TableCell {
                width: 1000,
                merged: true,
                ..crate::writer::TableCell::default()
            },
            crate::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::writer::TableCell::default()
            },
        ],
        ..crate::writer::TableRow::default()
    };
    assert!(
        writer
            .set_table_row_formatting(table, 0, invalid_merge)
            .is_err()
    );

    let late_header = crate::writer::TableRow {
        cells: vec![
            crate::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::writer::TableCell::default()
            },
            crate::writer::TableCell {
                width: 1000,
                merged: false,
                ..crate::writer::TableCell::default()
            },
        ],
        is_header: true,
        ..crate::writer::TableRow::default()
    };
    writer
        .set_table_row_formatting(table, 1, late_header)
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());

    let mut writer = Writer::new();
    let table = writer.add_table(1, 1).unwrap();
    writer
        .set_table_row_formatting(
            table,
            0,
            crate::writer::TableRow {
                cells: vec![crate::writer::TableCell {
                    width: 1000,
                    vertical_merge: crate::parts::tap::VerticalMergeStatus::Merged,
                    ..crate::writer::TableCell::default()
                }],
                ..crate::writer::TableRow::default()
            },
        )
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
}
#[cfg(test)]
mod header_kind_tests {
    use super::*;

    #[test]
    fn header_kinds_map_to_plcfhdd_slots() {
        assert_eq!(HeaderKind::Odd.slot(), HEADER_SLOT_ODD);
        assert_eq!(HeaderKind::Even.slot(), HEADER_SLOT_EVEN);
        assert_eq!(HeaderKind::FirstPage.slot(), HEADER_SLOT_FIRST);
        // The writer's slot assignment matches the MS-DOC PlcfHdd layout:
        // even header 6, odd header 7, first-page header 10.
        assert_eq!(
            (HEADER_SLOT_EVEN, HEADER_SLOT_ODD, HEADER_SLOT_FIRST),
            (6, 7, 10)
        );
    }

    #[test]
    fn header_shape_ids_use_the_header_cluster() {
        let mut writer = Writer::new();
        writer
            .insert_header_picture(
                HeaderKind::Odd,
                crate::writer::images::Picture::from_parts(
                    vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
                    480,
                    240,
                )
                .unwrap(),
                crate::writer::images::FloatingPosition::new(0, 0),
            )
            .unwrap();
        writer
            .insert_header_text_box(
                HeaderKind::Even,
                crate::writer::shapes::Shape::new(
                    crate::writer::shapes::Kind::Rectangle,
                    1440,
                    720,
                )
                .unwrap(),
                crate::writer::images::FloatingPosition::new(0, 0),
                "box",
            )
            .unwrap();
        // One shared cluster for both kinds, in insertion order.
        assert_eq!(writer.header_pictures[0].shape_id, 2049);
        assert_eq!(writer.header_shapes[0].shape_id, 2050);
        // Anchors landed in the right header paragraph lists.
        assert_eq!(writer.header_odd.as_ref().unwrap().len(), 1);
        assert_eq!(writer.header_even.as_ref().unwrap().len(), 1);
        assert!(writer.header_first.is_none());
    }
}

#[cfg(test)]
mod chpx_position_hresi_effect_writer_tests {
    use super::*;
    use crate::parts::chp::{CharacterPosition, HresiOperand, HyphenationMode, TextEffect};
    use std::io::Cursor;

    #[test]
    fn emits_canonical_typed_sprms_and_round_trips_package() {
        let position = CharacterPosition::new(-3168).unwrap();
        let hyphenation =
            HresiOperand::with_character(HyphenationMode::DeleteAndChange, b'Z').unwrap();
        let formatting = CharacterFormatting {
            position: Some(position),
            hyphenation: Some(hyphenation),
            text_effect: Some(TextEffect::Shimmer),
            ..CharacterFormatting::default()
        };
        let mut fonts = FontTableBuilder::new();
        let grpprl = build_chpx_grpprl(&formatting, &mut fonts);
        let mut expected = Vec::new();
        expected.extend_from_slice(&SPRM_C_HPS_POS.to_le_bytes());
        expected.extend_from_slice(&(-3168i16).to_le_bytes());
        expected.extend_from_slice(&SPRM_C_HRESI.to_le_bytes());
        expected.extend_from_slice(&[6, b'Z']);
        expected.extend_from_slice(&SPRM_C_SFXT_TEXT.to_le_bytes());
        expected.push(6);
        assert_eq!(grpprl, expected);

        let properties = crate::parts::chp::CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.position, position);
        assert_eq!(properties.hyphenation, hyphenation);
        assert_eq!(properties.text_effect, TextEffect::Shimmer);

        let mut writer = Writer::new();
        writer
            .add_paragraph_runs(
                vec![("effects".to_string(), formatting)],
                ParagraphFormatting::default(),
            )
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let mut package = crate::Package::from_reader(Cursor::new(output.into_inner())).unwrap();
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();
        let runs = paragraphs[0].runs().unwrap();
        let properties = runs[0].properties();
        assert_eq!(properties.position, position);
        assert_eq!(properties.hyphenation, hyphenation);
        assert_eq!(properties.text_effect, TextEffect::Shimmer);
    }
}
