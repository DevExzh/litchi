use super::super::super::support::*;

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
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
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
