use super::super::support::*;

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
