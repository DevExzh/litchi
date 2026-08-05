//! Focused TAP model and codec regression tests.

use super::*;
use crate::parts::tap::{
    BorderStyle, BorderType, CellBorderTypes, CellBorders, CellShading, CellSpacing,
    CellSpacingSource, ShadingPattern, TableConditionalFormatting, TableHorizontalAnchor,
    TableHorizontalPosition, TableJustification, TableLook, TableLookFlags, TablePositioning,
    TableStyleBorder, TableStyleCondition, TableStyleDefaults, TableStyleShading,
    TableVerticalAnchor, TableVerticalPosition, TableWidth, TextDirection, VerticalAlignment,
    VerticalMergeStatus, WidthType,
};

#[test]
fn test_tap_builder() {
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![
            TableCell {
                width: 1000,
                merged: false,
                vertical_merge: VerticalMergeStatus::First,
                vertical_alignment: VerticalAlignment::Center,
                text_direction: TextDirection::TbRl,
                fit_text: true,
                no_wrap: true,
                hide_mark: true,
                borders: CellBorders {
                    top: Some(BorderStyle {
                        width: 8,
                        color: Some((1, 2, 3)),
                        border_type: BorderType::Single,
                        spacing: 2,
                        shadow: true,
                        frame: false,
                    }),
                    diagonal_down: Some(BorderStyle {
                        width: 4,
                        color: Some((10, 20, 30)),
                        border_type: BorderType::Outset,
                        spacing: 1,
                        shadow: false,
                        frame: true,
                    }),
                    ..CellBorders::default()
                },
                border_type_overrides: CellBorderTypes::default(),
                shading: Some(CellShading {
                    foreground_color: Some((0, 0, 255)),
                    background_color: Some((255, 255, 0)),
                    pattern: ShadingPattern::DarkCross,
                }),
                padding_top: Some(120),
                padding_left: Some(240),
                padding_bottom: Some(120),
                padding_right: Some(240),
            },
            TableCell {
                width: 1000,
                merged: false,
                ..TableCell::default()
            },
        ],
        height: -200,
        is_header: true,
        allow_break: false,
        borders: TableBorders {
            vertical: Some(BorderStyle {
                width: 6,
                color: Some((40, 50, 60)),
                border_type: BorderType::Double,
                spacing: 0,
                shadow: false,
                frame: false,
            }),
            ..TableBorders::default()
        },
        ..TableRow::default()
    });

    let sprms = builder.try_generate_row_sprms(0).unwrap();
    let opcodes = crate::sprm::parse_sprms(&sprms)
        .unwrap()
        .into_iter()
        .map(|sprm| sprm.opcode)
        .collect::<Vec<_>>();
    let legacy_cant_split = opcodes.iter().position(|opcode| *opcode == 0x3403).unwrap();
    let modern_cant_split = opcodes.iter().position(|opcode| *opcode == 0x3466).unwrap();
    assert!(legacy_cant_split < modern_cant_split);
    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.cell_boundaries, [0, 1000, 2000]);
    assert_eq!(tap.row_height, Some(-200));
    assert!(tap.is_header_row);
    assert!(!tap.allow_row_break);
    assert_eq!(
        tap.cell_properties[0].vertical_merge_status,
        VerticalMergeStatus::First
    );
    assert_eq!(
        tap.cell_properties[0].vertical_alignment,
        VerticalAlignment::Center
    );
    assert_eq!(tap.cell_properties[0].text_direction, TextDirection::TbRl);
    assert!(tap.cell_properties[0].fit_text);
    assert!(tap.cell_properties[0].no_wrap);
    assert!(tap.cell_properties[0].hide_mark);
    let border = tap.cell_properties[0].borders.top.unwrap();
    assert_eq!(border.color, Some((1, 2, 3)));
    assert_eq!(border.spacing, 2);
    assert!(border.shadow);
    let diagonal = tap.cell_properties[0].borders.diagonal_down.unwrap();
    assert_eq!(diagonal.color, Some((10, 20, 30)));
    assert_eq!(diagonal.border_type, BorderType::Outset);
    assert!(diagonal.frame);
    assert_eq!(tap.border_vertical.unwrap().color, Some((40, 50, 60)));
    assert_eq!(
        tap.cell_properties[0].shading,
        Some(CellShading {
            foreground_color: Some((0, 0, 255)),
            background_color: Some((255, 255, 0)),
            pattern: ShadingPattern::DarkCross,
        })
    );
    assert_eq!(tap.cell_properties[0].padding_top, Some(120));
    assert_eq!(tap.cell_properties[0].padding_left, Some(240));
    assert_eq!(tap.cell_properties[0].padding_bottom, Some(120));
    assert_eq!(tap.cell_properties[0].padding_right, Some(240));
    assert_eq!(
        tap.cell_properties[0].preferred_width.unwrap().width_type,
        WidthType::Twips
    );
}

#[test]
fn test_tap_builder_empty() {
    let builder = TapBuilder::new();
    assert_eq!(builder.row_count(), 0);
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::RowOutOfBounds(0))
    );
}

#[test]
fn writes_full_color_shading_across_all_cell_chunks() {
    let shading = CellShading {
        foreground_color: Some((1, 2, 3)),
        background_color: Some((250, 240, 230)),
        pattern: ShadingPattern::Percent42Point5,
    };
    let mut cells = vec![TableCell::default(); 45];
    cells[0].shading = Some(shading);
    cells[22].shading = Some(shading);
    cells[44].shading = Some(shading);
    let row = TableRow {
        cells,
        ..TableRow::default()
    };

    let sprms = generate_row_sprms(&row).unwrap();
    let opcodes = crate::sprm::parse_sprms(&sprms)
        .unwrap()
        .into_iter()
        .map(|sprm| sprm.opcode)
        .collect::<Vec<_>>();
    assert!(!opcodes.contains(&0xD609));
    assert!(opcodes.contains(&0xD612));
    assert!(opcodes.contains(&0xD616));
    assert!(opcodes.contains(&0xD60C));
    assert!(opcodes.contains(&0xD670));
    assert!(opcodes.contains(&0xD671));
    assert!(opcodes.contains(&0xD672));

    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.cell_properties[0].shading, Some(shading));
    assert_eq!(tap.cell_properties[22].shading, Some(shading));
    assert_eq!(tap.cell_properties[44].shading, Some(shading));
    assert!(tap.cell_properties[43].shading.is_none());
}

#[test]
fn round_trips_scalar_table_style_defaults() {
    let defaults = TableStyleDefaults {
        padding_top: Some(120),
        padding_left: Some(240),
        padding_bottom: Some(120),
        padding_right: Some(240),
        vertical_alignment: Some(VerticalAlignment::Bottom),
        no_wrap: Some(false),
        horizontal_band_size: Some(2),
        vertical_band_size: Some(3),
        ..TableStyleDefaults::default()
    };
    let sprms = generate_table_style_sprms(&defaults).unwrap();
    let parsed_sprms = crate::sprm::parse_sprms(&sprms).unwrap();
    assert_eq!(
        parsed_sprms
            .iter()
            .filter(|sprm| sprm.opcode == 0xD63E)
            .count(),
        2
    );

    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.style_defaults, defaults);
}

#[test]
fn rejects_invalid_scalar_table_style_defaults() {
    assert_eq!(
        generate_table_style_sprms(&TableStyleDefaults {
            padding_top: Some(31_681),
            ..TableStyleDefaults::default()
        }),
        Err(TapBuildError::InvalidCellPadding(31_681))
    );
    assert_eq!(
        generate_table_style_sprms(&TableStyleDefaults {
            horizontal_band_size: Some(0),
            ..TableStyleDefaults::default()
        }),
        Err(TapBuildError::InvalidStyleBandSize("horizontal", 0))
    );
    assert_eq!(
        generate_table_style_sprms(&TableStyleDefaults {
            vertical_band_size: Some(4),
            ..TableStyleDefaults::default()
        }),
        Err(TapBuildError::InvalidStyleBandSize("vertical", 4))
    );
}

#[test]
fn round_trips_visual_table_style_defaults() {
    let border = BorderStyle {
        width: 8,
        color: Some((1, 2, 3)),
        border_type: BorderType::Outset,
        spacing: 2,
        shadow: true,
        frame: false,
    };
    let shading = CellShading {
        foreground_color: Some((4, 5, 6)),
        background_color: Some((7, 8, 9)),
        pattern: ShadingPattern::DarkCross,
    };
    let defaults = TableStyleDefaults {
        shading: Some(TableStyleShading::Shading(shading)),
        ..TableStyleDefaults::default()
    };
    let conditional_properties = TableStyleDefaults {
        border_top: Some(TableStyleBorder::NoBorder),
        border_inside_vertical: Some(TableStyleBorder::Border(border)),
        ..TableStyleDefaults::default()
    };
    let conditional = TableConditionalFormatting {
        condition: TableStyleCondition::HeaderRow,
        properties: conditional_properties,
        raw_grpprl: Vec::new(),
    };
    let sprms = generate_table_style_sprms_with_conditionals(&defaults, &[conditional]).unwrap();
    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.style_defaults, defaults);
    assert_eq!(tap.conditional_formats.len(), 1);
    assert_eq!(
        tap.conditional_formats[0].condition,
        TableStyleCondition::HeaderRow
    );
    assert_eq!(
        tap.conditional_formats[0].properties,
        conditional_properties
    );
    assert!(!tap.conditional_formats[0].raw_grpprl.is_empty());

    let clear = TableStyleDefaults {
        shading: Some(TableStyleShading::NoShading),
        ..TableStyleDefaults::default()
    };
    let sprms = generate_table_style_sprms(&clear).unwrap();
    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.style_defaults, clear);
}

#[test]
fn rejects_invalid_visual_table_style_defaults() {
    let defaults = TableStyleDefaults {
        border_top: Some(TableStyleBorder::Border(BorderStyle {
            width: 8,
            color: None,
            border_type: BorderType::Single,
            spacing: 32,
            shadow: false,
            frame: false,
        })),
        ..TableStyleDefaults::default()
    };
    assert_eq!(
        generate_table_style_sprms(&defaults),
        Err(TapBuildError::StyleBorderOutsideConditional)
    );
    let conditional = TableConditionalFormatting {
        condition: TableStyleCondition::HeaderRow,
        properties: defaults,
        raw_grpprl: Vec::new(),
    };
    assert_eq!(
        generate_table_style_sprms_with_conditionals(
            &TableStyleDefaults::default(),
            &[conditional],
        ),
        Err(TapBuildError::InvalidBorderSpacing(32))
    );

    let mut recursive = Vec::new();
    recursive.extend_from_slice(&0xD66Au16.to_le_bytes());
    recursive.push(2);
    recursive.extend_from_slice(&1u16.to_le_bytes());
    let recursive = TableConditionalFormatting {
        condition: TableStyleCondition::HeaderRow,
        properties: TableStyleDefaults::default(),
        raw_grpprl: recursive,
    };
    assert!(matches!(
        generate_table_style_sprms_with_conditionals(&TableStyleDefaults::default(), &[recursive],),
        Err(TapBuildError::InvalidConditionalProperties(_))
    ));

    let oversized = TableConditionalFormatting {
        condition: TableStyleCondition::HeaderRow,
        properties: TableStyleDefaults::default(),
        raw_grpprl: vec![0; 254],
    };
    assert_eq!(
        generate_table_style_sprms_with_conditionals(&TableStyleDefaults::default(), &[oversized],),
        Err(TapBuildError::ConditionalPropertiesTooLong(254))
    );
}

#[test]
fn test_tap_builder_multiple_rows() {
    let mut builder = TapBuilder::new();
    for i in 0..5 {
        builder.add_row(TableRow {
            cells: vec![
                TableCell {
                    width: 1000,
                    merged: false,
                    ..TableCell::default()
                },
                TableCell {
                    width: 1000,
                    merged: false,
                    ..TableCell::default()
                },
                TableCell {
                    width: 1000,
                    merged: false,
                    ..TableCell::default()
                },
            ],
            height: 200 + (i as i16 * 50),
            is_header: i == 0,
            ..TableRow::default()
        });
    }

    assert_eq!(builder.row_count(), 5);
    let sprms = builder.generate_row_sprms(0);
    assert!(!sprms.is_empty());
}

#[test]
fn round_trips_cell_border_type_prefix_overrides() {
    let overrides = CellBorderTypes {
        top: Some(BorderType::Double),
        left: Some(BorderType::Dotted),
        bottom: Some(BorderType::None),
        right: Some(BorderType::Outset),
    };
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![
            TableCell {
                width: 1000,
                border_type_overrides: overrides,
                ..TableCell::default()
            },
            TableCell {
                width: 1000,
                ..TableCell::default()
            },
        ],
        ..TableRow::default()
    });

    let sprms = builder.try_generate_row_sprms(0).unwrap();
    let border_type_sprm = crate::sprm::parse_sprms(&sprms)
        .unwrap()
        .into_iter()
        .find(|sprm| sprm.opcode == 0xD662)
        .unwrap();
    assert_eq!(border_type_sprm.operand_bytes(), &[3, 6, 0, 0x1A]);

    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.cell_properties[0].border_type_overrides, overrides);
    assert_eq!(
        tap.cell_properties[1].border_type_overrides,
        CellBorderTypes::default()
    );
}

#[test]
fn round_trips_row_identity_metadata() {
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        paragraph_group_id: Some(0x1020_3040),
        revision_save_id: Some(0xA1B2_C3D4),
        ..TableRow::default()
    });

    let sprms = builder.try_generate_row_sprms(0).unwrap();
    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.paragraph_group_id, Some(0x1020_3040));
    assert_eq!(tap.revision_save_id, Some(0xA1B2_C3D4));
}

#[test]
fn round_trips_logical_justification_for_rtl_tables() {
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        justification: TableJustification::Right,
        right_to_left: true,
        ..TableRow::default()
    });

    let sprms = builder.try_generate_row_sprms(0).unwrap();
    let parsed = crate::sprm::parse_sprms(&sprms).unwrap();
    let legacy = parsed.iter().find(|sprm| sprm.opcode == 0x5400).unwrap();
    let modern = parsed.iter().find(|sprm| sprm.opcode == 0x548A).unwrap();
    assert_eq!(legacy.operand_bytes(), &[0, 0]);
    assert_eq!(modern.operand_bytes(), &[2, 0]);

    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.justification, TableJustification::Right);
    assert_eq!(
        tap.legacy_physical_justification,
        Some(TableJustification::Left)
    );
    assert_eq!(
        tap.modern_logical_justification,
        Some(TableJustification::Right)
    );
}

#[test]
fn round_trips_table_row_revision_state() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        justification: TableJustification::Center,
        preserved_properties_for_revision: Some(Box::new(TableRow {
            cells: vec![TableCell {
                width: 1000,
                ..TableCell::default()
            }],
            justification: TableJustification::Right,
            formatting_revision: Some(TableRevisionMark {
                active: true,
                author_index: 12,
                timestamp,
            }),
            ..TableRow::default()
        })),
        ..TableRow::default()
    });

    let sprms = builder.try_generate_row_sprms(0).unwrap();
    let parsed = crate::sprm::parse_sprms(&sprms).unwrap();
    let position = |opcode| {
        parsed
            .iter()
            .position(|sprm| sprm.opcode == opcode)
            .unwrap()
    };
    let wall = position(0x3668);
    assert!(position(0xD667) < wall);
    for opcode in [0x5400, 0x548A] {
        let positions = parsed
            .iter()
            .enumerate()
            .filter_map(|(index, sprm)| (sprm.opcode == opcode).then_some(index))
            .collect::<Vec<_>>();
        assert!(positions[0] < wall);
        assert!(wall < positions[1]);
    }

    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.has_formatting_revision, Some(true));
    assert_eq!(tap.formatting_revision_author_index, Some(12));
    assert_eq!(tap.formatting_revision_timestamp, Some(timestamp));
    assert!(tap.properties_preserved_for_revision);
    assert_eq!(tap.justification, TableJustification::Center);
    let previous = tap.preserved_properties_for_revision.unwrap();
    assert_eq!(previous.justification, TableJustification::Right);
    assert_eq!(previous.formatting_revision_author_index, Some(12));
}

#[test]
fn round_trips_table_sizing_and_fit_properties() {
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![
            TableCell {
                width: 1000,
                ..TableCell::default()
            },
            TableCell {
                width: 1000,
                ..TableCell::default()
            },
        ],
        preferred_width: Some(TableWidth {
            value: 7_500,
            width_type: WidthType::Percentage,
        }),
        auto_fit: true,
        width_before: Some(TableWidth {
            value: 250,
            width_type: WidthType::Percentage,
        }),
        width_after: Some(TableWidth {
            value: 400,
            width_type: WidthType::Twips,
        }),
        preferred_indent: Some(TableWidth {
            value: -120,
            width_type: WidthType::Twips,
        }),
        keep_with_next: true,
        table_look: Some(TableLook {
            autoformat_index: -1,
            flags: TableLookFlags::BORDERS
                | TableLookFlags::HEADER_COLUMN
                | TableLookFlags::NO_COLUMN_BANDING,
        }),
        table_style_index: Some(0x1234),
        right_to_left: true,
        allow_overlap: false,
        positioning: Some(TablePositioning {
            vertical_anchor: TableVerticalAnchor::Paragraph,
            horizontal_anchor: TableHorizontalAnchor::Page,
        }),
        horizontal_position: TableHorizontalPosition::Center,
        vertical_position: TableVerticalPosition::Offset(720),
        distance_from_text_left: 120,
        distance_from_text_top: 240,
        distance_from_text_right: 360,
        distance_from_text_bottom: 480,
        cell_spacing: Some(CellSpacing {
            width: 240,
            source: CellSpacingSource::TableBorder,
        }),
        ..TableRow::default()
    });

    let sprms = builder.try_generate_row_sprms(0).unwrap();
    let opcodes = crate::sprm::parse_sprms(&sprms)
        .unwrap()
        .into_iter()
        .map(|sprm| sprm.opcode)
        .collect::<Vec<_>>();
    assert_eq!(opcodes[0], 0x563A);
    assert!(
        opcodes.iter().position(|opcode| *opcode == 0x560B)
            < opcodes.iter().position(|opcode| *opcode == 0x5664)
    );
    let tap = crate::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
    assert_eq!(tap.preferred_width, builder.rows()[0].preferred_width);
    assert!(tap.auto_fit);
    assert_eq!(tap.width_before, builder.rows()[0].width_before);
    assert_eq!(tap.width_after, builder.rows()[0].width_after);
    assert_eq!(tap.preferred_indent, builder.rows()[0].preferred_indent);
    assert!(tap.keep_with_next);
    assert_eq!(tap.table_look, builder.rows()[0].table_look);
    assert_eq!(tap.table_style_index, Some(0x1234));
    assert!(tap.right_to_left);
    assert!(!tap.allow_overlap);
    assert_eq!(tap.positioning, builder.rows()[0].positioning);
    assert_eq!(
        tap.horizontal_position,
        builder.rows()[0].horizontal_position
    );
    assert_eq!(tap.vertical_position, builder.rows()[0].vertical_position);
    assert_eq!(tap.distance_from_text_left, 120);
    assert_eq!(tap.distance_from_text_top, 240);
    assert_eq!(tap.distance_from_text_right, 360);
    assert_eq!(tap.distance_from_text_bottom, 480);
    assert_eq!(tap.cell_spacing, builder.rows()[0].cell_spacing);
}

#[test]
fn test_create_simple_table() {
    let table = create_simple_table(3, 4, 1440); // 3 rows, 4 cols, 1 inch cells
    assert_eq!(table.row_count(), 3);
}

#[test]
fn test_create_simple_table_single_cell() {
    let table = create_simple_table(1, 1, 1000);
    assert_eq!(table.row_count(), 1);
    assert_eq!(table.rows()[0].cells.len(), 1);
}

#[test]
fn test_create_simple_table_large() {
    let table = create_simple_table(10, 10, 500);
    assert_eq!(table.row_count(), 10);
    assert_eq!(table.rows()[0].cells.len(), 10);
}

#[test]
fn test_table_row_count() {
    let table = create_simple_table(5, 3, 1000);
    assert_eq!(table.row_count(), 5);
}

#[test]
fn rejects_unrepresentable_rows() {
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        preserved_properties_for_revision: Some(Box::new(TableRow {
            properties_preserved_for_revision: true,
            ..TableRow::default()
        })),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::NestedPreservedState)
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            merged: true,
            ..TableCell::default()
        }],
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::MergeWithoutPrecedingCell)
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: u16::MAX,
            merged: false,
            ..TableCell::default()
        }],
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::CellWidthsOverflow)
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        height: i16::MIN,
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidRowHeight(i16::MIN))
    );

    let invalid_width = TableWidth {
        value: 30_001,
        width_type: WidthType::Percentage,
    };
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        preferred_width: Some(invalid_width),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidPreferredWidth(
            "table width",
            invalid_width
        ))
    );

    let invalid_indent = TableWidth {
        value: 31_000,
        width_type: WidthType::Twips,
    };
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        preferred_width: Some(TableWidth {
            value: 1_000,
            width_type: WidthType::Twips,
        }),
        preferred_indent: Some(invalid_indent),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidPreferredWidth(
            "table indent",
            invalid_indent
        ))
    );

    let invalid_flags = 0x8000;
    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        table_look: Some(TableLook {
            autoformat_index: 0,
            flags: TableLookFlags::from_bits_retain(invalid_flags),
        }),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidTableLookFlags(invalid_flags))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        horizontal_position: TableHorizontalPosition::Center,
        ..TableRow::default()
    });
    assert!(builder.try_generate_row_sprms(0).is_ok());

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        positioning: Some(TablePositioning {
            vertical_anchor: TableVerticalAnchor::Margin,
            horizontal_anchor: TableHorizontalAnchor::Column,
        }),
        horizontal_position: TableHorizontalPosition::Offset(-1),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidTablePosition("horizontal", -1))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        positioning: Some(TablePositioning {
            vertical_anchor: TableVerticalAnchor::Margin,
            horizontal_anchor: TableHorizontalAnchor::Column,
        }),
        distance_from_text_right: 31_681,
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidWrapDistance("right", 31_681))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        cell_spacing: Some(CellSpacing {
            width: 15_841,
            source: CellSpacingSource::Explicit,
        }),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidCellSpacing(15_841))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        paragraph_group_id: Some(0),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidParagraphGroupId)
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        formatting_revision: Some(TableRevisionMark {
            active: true,
            author_index: 0x8000,
            timestamp: 0,
        }),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidRevisionAuthorIndex(0x8000))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            ..TableCell::default()
        }],
        formatting_revision: Some(TableRevisionMark {
            active: true,
            author_index: 0,
            timestamp: 0x3F,
        }),
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidRevisionTimestamp(0x3F))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            border_type_overrides: CellBorderTypes {
                top: Some(BorderType::Single),
                ..CellBorderTypes::default()
            },
            ..TableCell::default()
        }],
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::IncompleteCellBorderTypes(0))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![
            TableCell {
                width: 1000,
                ..TableCell::default()
            },
            TableCell {
                width: 1000,
                border_type_overrides: CellBorderTypes {
                    top: Some(BorderType::Single),
                    left: Some(BorderType::Single),
                    bottom: Some(BorderType::Single),
                    right: Some(BorderType::Single),
                },
                ..TableCell::default()
            },
        ],
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::IncompleteCellBorderTypes(0))
    );

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            borders: CellBorders {
                top: Some(BorderStyle {
                    width: 8,
                    color: Some((1, 2, 3)),
                    border_type: BorderType::Single,
                    spacing: 0,
                    shadow: false,
                    frame: false,
                }),
                ..CellBorders::default()
            },
            ..TableCell::default()
        }],
        ..TableRow::default()
    });
    assert!(builder.try_generate_row_sprms(0).is_ok());

    let mut builder = TapBuilder::new();
    builder.add_row(TableRow {
        cells: vec![TableCell {
            width: 1000,
            padding_left: Some(31_681),
            ..TableCell::default()
        }],
        ..TableRow::default()
    });
    assert_eq!(
        builder.try_generate_row_sprms(0),
        Err(TapBuildError::InvalidCellPadding(31_681))
    );
}
