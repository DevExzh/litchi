use super::super::super::support::*;

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
