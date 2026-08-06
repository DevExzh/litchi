//! Focused PAP parser/model/codec regression tests.

use super::*;
use crate::parts::numbering::NumberFormat;
use crate::parts::tap::TableStyleCondition;
use crate::sprm::parse_sprms;
use crate::sprm_operations::*;

#[test]
fn test_default_pap() {
    let pap = ParagraphProperties::new();
    assert_eq!(pap.justification, Justification::Left);
    assert!(!pap.keep_on_page);
    assert!(!pap.has_formatting());
}

#[test]
fn parses_legacy_paragraph_border_controls_strictly() {
    for (encoded, expected) in [
        (0, LegacyBorderStyle::Single),
        (1, LegacyBorderStyle::Thick),
        (2, LegacyBorderStyle::Double),
        (3, LegacyBorderStyle::Shadow),
    ] {
        let properties = ParagraphProperties::from_sprm(
            &[SPRM_P_BRCL.to_le_bytes().as_slice(), &[encoded]].concat(),
        )
        .unwrap();
        assert_eq!(properties.legacy_border_style, Some(expected));
        assert!(properties.has_formatting());
    }
    for (encoded, expected) in [
        (0, LegacyBorderPosition::None),
        (1, LegacyBorderPosition::Above),
        (2, LegacyBorderPosition::Below),
        (15, LegacyBorderPosition::Box),
        (16, LegacyBorderPosition::LeftBar),
    ] {
        let properties = ParagraphProperties::from_sprm(
            &[SPRM_P_BRCP.to_le_bytes().as_slice(), &[encoded]].concat(),
        )
        .unwrap();
        assert_eq!(properties.legacy_border_position, Some(expected));
        assert!(properties.has_formatting());
    }

    for encoded in 4..=u8::MAX {
        assert!(
            ParagraphProperties::from_sprm(
                &[SPRM_P_BRCL.to_le_bytes().as_slice(), &[encoded]].concat(),
            )
            .is_err()
        );
    }
    for encoded in 0..=u8::MAX {
        if !matches!(encoded, 0 | 1 | 2 | 15 | 16) {
            assert!(
                ParagraphProperties::from_sprm(
                    &[SPRM_P_BRCP.to_le_bytes().as_slice(), &[encoded]].concat(),
                )
                .is_err()
            );
        }
    }
}

#[test]
fn parses_conditional_table_style_paragraph_formatting_strictly() {
    let wrap = |condition: u16, nested: &[u8]| {
        let mut grpprl = SPRM_P_CNF.to_le_bytes().to_vec();
        grpprl.push((nested.len() + 2) as u8);
        grpprl.extend_from_slice(&condition.to_le_bytes());
        grpprl.extend_from_slice(nested);
        grpprl
    };
    let mut nested = SPRM_P_DYA_BEFORE.to_le_bytes().to_vec();
    nested.extend_from_slice(&120u16.to_le_bytes());
    nested.extend_from_slice(&SPRM_P_F_KEEP.to_le_bytes());
    nested.push(1);

    let properties = ParagraphProperties::from_sprm(&wrap(0x0001, &nested)).unwrap();
    assert_eq!(properties.conditional_formats.len(), 1);
    let conditional = &properties.conditional_formats[0];
    assert_eq!(conditional.condition, TableStyleCondition::HeaderRow);
    assert_eq!(conditional.raw_grpprl, nested);
    assert_eq!(conditional.properties.space_before, Some(120));
    assert!(conditional.properties.keep_on_page);

    let recursive = wrap(0x0002, &[]);
    let character = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
    let truncated = SPRM_P_DYA_BEFORE.to_le_bytes();
    for invalid in [
        [SPRM_P_CNF.to_le_bytes().as_slice(), &[0]].concat(),
        wrap(0x0003, &[]),
        wrap(0x0001, &recursive),
        wrap(0x0001, &character),
        wrap(0x0001, &truncated),
    ] {
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }
}

#[test]
fn parses_legacy_autonumber_descriptors_and_text_widths() {
    let grpprl = |operand: &[u8]| {
        let mut bytes = SPRM_P_ANLD.to_le_bytes().to_vec();
        bytes.push(operand.len() as u8);
        bytes.extend_from_slice(operand);
        bytes
    };
    let mut operand = [0u8; 84];
    operand[0] = NumberFormat::RussianUpper as u8;
    operand[1] = 1;
    operand[2] = 1;
    operand[3] = 0xFE;
    operand[4] = 0xFF;
    operand[5] = 3 | (6 << 3);
    operand[6..8].copy_from_slice(&4u16.to_le_bytes());
    operand[8..10].copy_from_slice(&24u16.to_le_bytes());
    operand[10..12].copy_from_slice(&3u16.to_le_bytes());
    operand[12..14].copy_from_slice(&(-360i16).to_le_bytes());
    operand[14..16].copy_from_slice(&180u16.to_le_bytes());
    operand[16..19].copy_from_slice(&[1, 0, 1]);
    operand[20..22].copy_from_slice(&('(' as u16).to_le_bytes());
    operand[22..24].copy_from_slice(&(')' as u16).to_le_bytes());

    let expected = LegacyAutoNumbering {
        number_format: NumberFormat::RussianUpper,
        alignment: AutoNumberAlignment::Right,
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
        prefix: "(".to_string(),
        suffix: ")".to_string(),
    };

    for width in [52, 84] {
        let properties = ParagraphProperties::from_sprm(&grpprl(&operand[..width])).unwrap();
        assert_eq!(properties.legacy_autonumbering, Some(expected.clone()));
        assert!(properties.has_formatting());
    }
}

#[test]
fn rejects_malformed_legacy_autonumber_descriptors() {
    let grpprl = |operand: &[u8]| {
        let mut bytes = SPRM_P_ANLD.to_le_bytes().to_vec();
        bytes.push(operand.len() as u8);
        bytes.extend_from_slice(operand);
        bytes
    };
    let valid = [0u8; 84];
    let mut cases = vec![valid[..51].to_vec(), valid[..53].to_vec()];
    for mutation in [
        |value: &mut [u8; 84]| value[0] = 60,
        |value: &mut [u8; 84]| value[5] = 17 << 3,
        |value: &mut [u8; 84]| value[16] = 2,
        |value: &mut [u8; 84]| value[17] = 2,
        |value: &mut [u8; 84]| value[18] = 2,
        |value: &mut [u8; 84]| value[19] = 1,
        |value: &mut [u8; 84]| {
            value[12..14].copy_from_slice(&i16::MIN.to_le_bytes());
        },
        |value: &mut [u8; 84]| {
            value[14..16].copy_from_slice(&31_681u16.to_le_bytes());
        },
        |value: &mut [u8; 84]| value[1] = 33,
        |value: &mut [u8; 84]| {
            value[1] = 1;
            value[20..22].copy_from_slice(&0xD800u16.to_le_bytes());
        },
    ] {
        let mut operand = valid;
        mutation(&mut operand);
        cases.push(operand.to_vec());
    }
    for operand in cases {
        assert!(ParagraphProperties::from_sprm(&grpprl(&operand)).is_err());
    }
}

#[test]
fn test_justification() {
    let left = Justification::Left;
    let center = Justification::Center;
    assert_ne!(left, center);
    assert_eq!(left, Justification::Left);
}

#[test]
fn test_line_spacing_type() {
    let single = LineSpacingType::Single;
    let double = LineSpacingType::Double;
    assert_ne!(single, double);
}

#[test]
fn parses_lists_indents_spacing_and_ordered_tab_changes_strictly() {
    let mut grpprl = Vec::new();
    grpprl.extend_from_slice(&SPRM_P_ILVL.to_le_bytes());
    grpprl.push(8);
    grpprl.extend_from_slice(&SPRM_P_ILFO.to_le_bytes());
    grpprl.extend_from_slice(&1u16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_F_NO_LINE_NUMB.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_DXA_RIGHT.to_le_bytes());
    grpprl.extend_from_slice(&(-720i16).to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DXA_LEFT.to_le_bytes());
    grpprl.extend_from_slice(&1_440i16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_NEST.to_le_bytes());
    grpprl.extend_from_slice(&(-240i16).to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DXA_LEFT1.to_le_bytes());
    grpprl.extend_from_slice(&360i16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DYA_LINE.to_le_bytes());
    grpprl.extend_from_slice(&(-360i16).to_le_bytes());
    grpprl.extend_from_slice(&0u16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DYA_BEFORE.to_le_bytes());
    grpprl.extend_from_slice(&120u16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DYA_AFTER.to_le_bytes());
    grpprl.extend_from_slice(&240u16.to_le_bytes());

    // Add a dotted right tab and a bar tab whose ignored leader bits are reserved.
    grpprl.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
    grpprl.extend_from_slice(&[8, 0, 2]);
    grpprl.extend_from_slice(&100i16.to_le_bytes());
    grpprl.extend_from_slice(&200i16.to_le_bytes());
    grpprl.extend_from_slice(&[0x0A, 0x34]);
    // Delete within 25 twips of 100 and add a list tab with the default leader.
    grpprl.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
    grpprl.extend_from_slice(&[7, 1]);
    grpprl.extend_from_slice(&110i16.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&300i16.to_le_bytes());
    grpprl.push(0x3E);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.list_level, Some(8));
    assert_eq!(properties.list_format_override, Some(1));
    assert!(properties.no_line_numbering);
    assert_eq!(properties.indent_right, Some(-720));
    assert_eq!(properties.indent_left, Some(1_200));
    assert_eq!(properties.indent_first_line, Some(360));
    assert_eq!(properties.line_spacing, Some(-360));
    assert_eq!(properties.line_spacing_type, LineSpacingType::Exactly);
    assert_eq!(properties.space_before, Some(120));
    assert_eq!(properties.space_after, Some(240));
    assert_eq!(
        properties.tab_stops,
        vec![
            TabStop {
                position: 200,
                alignment: TabAlignment::Bar,
                leader: TabLeader::None,
            },
            TabStop {
                position: 300,
                alignment: TabAlignment::List,
                leader: TabLeader::DefaultLeader,
            },
        ]
    );

    // Fast-save deletion operands carry a parallel close-distance array.
    let mut fast_save = SPRM_P_CHG_TABS.to_le_bytes().to_vec();
    fast_save.extend_from_slice(&[6, 1]);
    fast_save.extend_from_slice(&300i16.to_le_bytes());
    fast_save.extend_from_slice(&26i16.to_le_bytes());
    fast_save.push(0);
    let mut properties = properties;
    for sprm in parse_sprms(&fast_save).unwrap() {
        ParagraphProperties::apply_sprm(&mut properties, &sprm).unwrap();
    }
    assert_eq!(properties.tab_stops.len(), 1);
    assert_eq!(properties.tab_stops[0].position, 200);
}

#[test]
fn rejects_invalid_list_indent_spacing_and_tab_operands() {
    let fixed = |opcode: u16, operand: &[u8]| {
        let mut grpprl = opcode.to_le_bytes().to_vec();
        grpprl.extend_from_slice(operand);
        grpprl
    };
    let variable = |opcode: u16, operand: &[u8]| {
        let mut grpprl = opcode.to_le_bytes().to_vec();
        grpprl.push(operand.len() as u8);
        grpprl.extend_from_slice(operand);
        grpprl
    };
    for grpprl in [
        fixed(SPRM_P_ILVL, &[9]),
        fixed(SPRM_P_ILFO, &0x07FFu16.to_le_bytes()),
        fixed(SPRM_P_F_NO_LINE_NUMB, &[2]),
        fixed(SPRM_P_DXA_LEFT, &31_681i16.to_le_bytes()),
        fixed(SPRM_P_DYA_LINE, &[0xC1, 0x7B, 0, 0]),
        fixed(SPRM_P_DYA_LINE, &[240, 0, 2, 0]),
        fixed(SPRM_P_DYA_BEFORE, &31_681u16.to_le_bytes()),
        variable(SPRM_P_CHG_TABS_PAPX, &[0, 1, 100, 0, 5]),
        variable(SPRM_P_CHG_TABS_PAPX, &[0, 1, 100, 0, 0x30]),
        variable(SPRM_P_CHG_TABS_PAPX, &[0, 2, 200, 0, 100, 0, 0, 0]),
        variable(SPRM_P_CHG_TABS, &[1, 100, 0, 0, 0x80, 0]),
    ] {
        assert!(ParagraphProperties::from_sprm(&grpprl).is_err());
    }
}

#[test]
fn test_indent_conversion() {
    let mut pap = ParagraphProperties::new();
    pap.indent_left = Some(1440); // 1 inch in twips
    assert_eq!(pap.get_indent_left_inches(), 1.0);
}

#[test]
fn parses_all_paragraph_formatting_revision_sprms_strictly() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    for opcode in [
        SPRM_P_PROP_RMARK,
        SPRM_P_PROP_RMARK90,
        SPRM_P_PROP_RMARK_CURRENT,
    ] {
        let mut grpprl = opcode.to_le_bytes().to_vec();
        grpprl.push(7);
        grpprl.push(1);
        grpprl.extend_from_slice(&2i16.to_le_bytes());
        grpprl.extend_from_slice(&timestamp.to_le_bytes());
        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.has_formatting_revision, Some(true));
        assert_eq!(properties.formatting_revision_author_index, Some(2));
        assert_eq!(properties.formatting_revision_timestamp, Some(timestamp));
    }

    for operand in [
        vec![2, 0, 0, 0, 0, 0, 0],
        vec![1, 0xFF, 0xFF, 0, 0, 0, 0],
        vec![1, 0, 0, 0, 0, 0],
        vec![1, 0, 0, 0x3F, 0, 0, 0],
    ] {
        let mut grpprl = SPRM_P_PROP_RMARK_CURRENT.to_le_bytes().to_vec();
        grpprl.push(operand.len() as u8);
        grpprl.extend_from_slice(&operand);
        assert!(ParagraphProperties::from_sprm(&grpprl).is_err());
    }
}

#[test]
fn parses_table_row_revision_state_strictly() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut grpprl = 0xD667u16.to_le_bytes().to_vec();
    grpprl.push(7);
    grpprl.push(1);
    grpprl.extend_from_slice(&2i16.to_le_bytes());
    grpprl.extend_from_slice(&timestamp.to_le_bytes());
    grpprl.extend_from_slice(&0x3668u16.to_le_bytes());
    grpprl.push(1);
    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.has_table_formatting_revision, Some(true));
    assert_eq!(properties.table_formatting_revision_author_index, Some(2));
    assert_eq!(
        properties.table_formatting_revision_timestamp,
        Some(timestamp)
    );
    assert!(properties.table_properties_preserved_for_revision);

    for operand in [
        vec![2, 0, 0, 0, 0, 0, 0],
        vec![1, 0xFF, 0xFF, 0, 0, 0, 0],
        vec![1, 0, 0, 0, 0, 0],
        vec![1, 0, 0, 0x3F, 0, 0, 0],
    ] {
        let mut invalid = 0xD667u16.to_le_bytes().to_vec();
        invalid.push(operand.len() as u8);
        invalid.extend_from_slice(&operand);
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }

    let invalid_wall = [0x68, 0x36, 2];
    assert!(ParagraphProperties::from_sprm(&invalid_wall).is_err());
}

#[test]
fn parses_numbering_revision_state_strictly() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut numrm = [0u8; 128];
    numrm[0] = 1;
    numrm[2..4].copy_from_slice(&1i16.to_le_bytes());
    numrm[4..8].copy_from_slice(&timestamp.to_le_bytes());
    numrm[8] = 1;
    numrm[17] = 0;
    numrm[28..32].copy_from_slice(&12u32.to_le_bytes());
    numrm[64..66].copy_from_slice(&2u16.to_le_bytes());
    numrm[66..68].copy_from_slice(&('%' as u16).to_le_bytes());
    numrm[68..70].copy_from_slice(&('.' as u16).to_le_bytes());

    let mut grpprl = SPRM_P_F_NUM_RM_INS.to_le_bytes().to_vec();
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_NUM_RM.to_le_bytes());
    grpprl.push(128);
    grpprl.extend_from_slice(&numrm);
    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.numbering_revision_list_applied, Some(true));
    let revision = properties.numbering_revision.unwrap();
    assert!(revision.was_numbered);
    assert_eq!(revision.author_index, 1);
    assert_eq!(revision.timestamp, timestamp);
    assert_eq!(revision.placeholder_positions[0], 1);
    assert_eq!(revision.numbers[0], 12);
    assert_eq!(revision.format_string, "%.");

    let mut invalid_bool = SPRM_P_F_NUM_RM_INS.to_le_bytes().to_vec();
    invalid_bool.push(2);
    assert!(ParagraphProperties::from_sprm(&invalid_bool).is_err());

    for mutate in [0usize, 2, 8, 64] {
        let mut invalid = numrm;
        match mutate {
            0 => invalid[0] = 2,
            2 => invalid[2..4].copy_from_slice(&(-1i16).to_le_bytes()),
            8 => invalid[8] = 3,
            64 => invalid[64..66].copy_from_slice(&32u16.to_le_bytes()),
            _ => unreachable!(),
        }
        let mut grpprl = SPRM_P_NUM_RM.to_le_bytes().to_vec();
        grpprl.push(128);
        grpprl.extend_from_slice(&invalid);
        assert!(ParagraphProperties::from_sprm(&grpprl).is_err());
    }
}

#[test]
fn parses_current_paragraph_identity_and_revision_state_strictly() {
    let mut grpprl = Vec::new();
    grpprl.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_IPGP.to_le_bytes());
    grpprl.extend_from_slice(&9u32.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_RSID.to_le_bytes());
    grpprl.extend_from_slice(&0x1122_3344u32.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_F_NO_ALLOW_OVERLAP.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_F_CONTEXTUAL_SPACING.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_F_MIRROR_INDENTS.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_TTWO.to_le_bytes());
    grpprl.push(4);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert!(properties.properties_preserved_for_revision);
    assert!(properties.preserved_properties_for_revision.is_some());
    assert_eq!(properties.paragraph_group_id, Some(9));
    assert_eq!(properties.revision_save_id, Some(0x1122_3344));
    assert!(properties.no_allow_overlap);
    assert!(properties.contextual_spacing);
    assert!(properties.mirror_indents);
    assert_eq!(
        properties.text_box_tight_wrap,
        Some(TextBoxTightWrap::LastLineOnly)
    );

    let invalid_bool = [SPRM_P_WALL.to_le_bytes().as_slice(), &[2]].concat();
    assert!(ParagraphProperties::from_sprm(&invalid_bool).is_err());

    let invalid_group = [
        SPRM_P_IPGP.to_le_bytes().as_slice(),
        0u32.to_le_bytes().as_slice(),
    ]
    .concat();
    assert!(ParagraphProperties::from_sprm(&invalid_group).is_err());

    let truncated_rsid = [SPRM_P_RSID.to_le_bytes().as_slice(), &[1, 2]].concat();
    assert!(ParagraphProperties::from_sprm(&truncated_rsid).is_err());

    let invalid_tight_wrap = [SPRM_P_TTWO.to_le_bytes().as_slice(), &[5]].concat();
    assert!(ParagraphProperties::from_sprm(&invalid_tight_wrap).is_err());
}

#[test]
fn preserves_and_resets_ordered_paragraph_revision_state() {
    let mut grpprl = Vec::new();
    grpprl.extend_from_slice(&SPRM_P_DXA_LEFT.to_le_bytes());
    grpprl.extend_from_slice(&100i16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_P_DXA_RIGHT.to_le_bytes());
    grpprl.extend_from_slice(&200i16.to_le_bytes());

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert!(properties.properties_preserved_for_revision);
    assert_eq!(properties.indent_left, Some(100));
    assert_eq!(properties.indent_right, Some(200));
    let previous = properties.preserved_properties_for_revision.unwrap();
    assert_eq!(previous.indent_left, Some(100));
    assert_eq!(previous.indent_right, None);
    assert!(!previous.properties_preserved_for_revision);
    assert!(previous.preserved_properties_for_revision.is_none());

    grpprl.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
    grpprl.push(0);
    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert!(!properties.properties_preserved_for_revision);
    assert!(properties.preserved_properties_for_revision.is_none());
}

#[test]
fn parses_current_character_relative_paragraph_layout_strictly() {
    let mut grpprl = Vec::new();
    for (opcode, value) in [
        (SPRM_P_DXC_RIGHT, -125i16),
        (SPRM_P_DXC_LEFT, 250),
        (SPRM_P_DXC_LEFT1, -50),
        (SPRM_P_DYL_BEFORE, -20),
        (SPRM_P_DYL_AFTER, 31_680),
        (SPRM_P_DXA_LEFT_2000, 100),
        (SPRM_P_NEST_2000, -20),
    ] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(&value.to_le_bytes());
    }
    for opcode in [
        SPRM_P_F_OPEN_TCH,
        SPRM_P_F_DYA_BEFORE_AUTO,
        SPRM_P_F_DYA_AFTER_AUTO,
    ] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(1);
    }
    grpprl.extend_from_slice(&SPRM_P_JC_LOGICAL.to_le_bytes());
    grpprl.push(9);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.indent_right_chars, Some(-125));
    assert_eq!(properties.indent_left_chars, Some(250));
    assert_eq!(properties.indent_first_line_chars, Some(-50));
    assert_eq!(properties.space_before_lines, Some(-20));
    assert_eq!(properties.space_after_lines, Some(31_680));
    assert_eq!(properties.indent_left, Some(80));
    assert!(properties.open_table_cell_mark);
    assert!(properties.space_before_auto);
    assert!(properties.space_after_auto);
    assert_eq!(properties.justification, Justification::ThaiDistributed);

    for (opcode, value) in [(SPRM_P_DYL_BEFORE, -21i16), (SPRM_P_DYL_AFTER, 31_681)] {
        let invalid = [opcode.to_le_bytes(), value.to_le_bytes()].concat();
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }

    let invalid_bool = [SPRM_P_F_OPEN_TCH.to_le_bytes().as_slice(), &[2]].concat();
    assert!(ParagraphProperties::from_sprm(&invalid_bool).is_err());

    let invalid_logical_jc = [SPRM_P_JC_LOGICAL.to_le_bytes().as_slice(), &[10]].concat();
    assert!(ParagraphProperties::from_sprm(&invalid_logical_jc).is_err());

    let invalid_legacy_jc = [SPRM_P_JC.to_le_bytes().as_slice(), &[6]].concat();
    assert!(ParagraphProperties::from_sprm(&invalid_legacy_jc).is_err());
}

#[test]
fn parses_current_and_word97_paragraph_borders_strictly() {
    let mut grpprl = SPRM_P_BRC_TOP.to_le_bytes().to_vec();
    grpprl.push(8);
    grpprl.extend_from_slice(&[0x11, 0x22, 0x33, 0, 12, 27, 0x67, 0]);
    grpprl.extend_from_slice(&SPRM_P_BRC_LEFT80.to_le_bytes());
    grpprl.extend_from_slice(&[8, 3, 2, 0x24]);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(
        properties.borders.top,
        Some(Border {
            style: BorderStyle::Inset,
            width: 12,
            color: Some((0x11, 0x22, 0x33)),
            spacing: 7,
            shadow: true,
            frame: true,
        })
    );
    assert_eq!(
        properties.borders.left,
        Some(Border {
            style: BorderStyle::Double,
            width: 8,
            color: Some((0, 0, 255)),
            spacing: 4,
            shadow: true,
            frame: false,
        })
    );

    for operand in [
        vec![0, 0, 0, 1, 8, 1, 0, 0],
        vec![0, 0, 0, 0, 8, 2, 0, 0],
        vec![0; 7],
    ] {
        let mut invalid = SPRM_P_BRC_TOP.to_le_bytes().to_vec();
        invalid.push(operand.len() as u8);
        invalid.extend_from_slice(&operand);
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }

    for operand in [[8, 1, 17, 0], [8, 26, 1, 0]] {
        let invalid = [SPRM_P_BRC_TOP80.to_le_bytes().as_slice(), &operand].concat();
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }
}

#[test]
fn parses_word6_paragraph_borders_and_pagination_controls_strictly() {
    let single = 2u16 | (1 << 3) | (1 << 5) | (6 << 6) | (7 << 11);
    let dotted = 6u16 | (2 << 6) | (3 << 11);
    let mut grpprl = Vec::new();
    for (opcode, border) in [
        (SPRM_P_BRC_TOP10, single),
        (SPRM_P_BRC_LEFT10, dotted),
        (SPRM_P_BRC_BOTTOM10, single),
        (SPRM_P_BRC_RIGHT10, single),
        (SPRM_P_BRC_BETWEEN10, single),
        (SPRM_P_BRC_BAR10, single),
    ] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(&border.to_le_bytes());
    }
    grpprl.extend_from_slice(&SPRM_P_DXA_FROM_TEXT10.to_le_bytes());
    grpprl.extend_from_slice(&720i16.to_le_bytes());
    for opcode in [
        SPRM_P_F_SIDE_BY_SIDE,
        SPRM_P_F_KEEP,
        SPRM_P_F_KEEP_FOLLOW,
        SPRM_P_F_PAGE_BREAK_BEFORE,
    ] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(1);
    }

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(
        properties.borders.top,
        Some(Border {
            style: BorderStyle::Single,
            width: 12,
            color: Some((255, 0, 0)),
            spacing: 7,
            shadow: true,
            frame: false,
        })
    );
    assert_eq!(
        properties.borders.left,
        Some(Border {
            style: BorderStyle::Dotted,
            width: 6,
            color: Some((0, 0, 255)),
            spacing: 3,
            shadow: false,
            frame: false,
        })
    );
    assert_eq!(properties.dxa_from_text, Some(720));
    assert!(properties.side_by_side);
    assert!(properties.keep_on_page);
    assert!(properties.keep_with_next);
    assert!(properties.page_break_before);

    for raw in [1u16 << 3, 2 | (1 << 3) | (17 << 6)] {
        let invalid = [
            SPRM_P_BRC_TOP10.to_le_bytes().as_slice(),
            raw.to_le_bytes().as_slice(),
        ]
        .concat();
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }
    let negative_distance = [
        SPRM_P_DXA_FROM_TEXT10.to_le_bytes().as_slice(),
        (-1i16).to_le_bytes().as_slice(),
    ]
    .concat();
    assert!(ParagraphProperties::from_sprm(&negative_distance).is_err());
    for opcode in [
        SPRM_P_F_SIDE_BY_SIDE,
        SPRM_P_F_KEEP,
        SPRM_P_F_KEEP_FOLLOW,
        SPRM_P_F_PAGE_BREAK_BEFORE,
    ] {
        let invalid = [opcode.to_le_bytes().as_slice(), &[2]].concat();
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }
}

#[test]
fn parses_current_grid_table_depth_and_shading_strictly() {
    let mut grpprl = Vec::new();
    for opcode in [SPRM_P_F_USE_PGSU_SETTINGS, SPRM_P_F_ADJUST_RIGHT] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(1);
    }
    grpprl.extend_from_slice(&SPRM_P_ITAP.to_le_bytes());
    grpprl.extend_from_slice(&3i32.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DTAP.to_le_bytes());
    grpprl.extend_from_slice(&(-1i32).to_le_bytes());
    for opcode in [SPRM_P_F_INNER_TABLE_CELL, SPRM_P_F_INNER_TTP] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(1);
    }
    grpprl.extend_from_slice(&SPRM_P_SHD.to_le_bytes());
    grpprl.push(10);
    grpprl.extend_from_slice(&[1, 2, 3, 0, 4, 5, 6, 0, 0x19, 0]);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.use_page_setup_settings, Some(true));
    assert_eq!(properties.adjust_right_indent, Some(true));
    assert_eq!(properties.table_nesting_level, 2);
    assert!(properties.inner_table_cell);
    assert!(properties.inner_table_row_end);
    assert_eq!(
        properties.shading,
        Some(Shading {
            foreground_color: Some((1, 2, 3)),
            background_color: Some((4, 5, 6)),
            pattern: ShadingPattern::DiagonalCross,
        })
    );

    let invalid_i32 = |opcode: u16, value: i32| {
        [
            opcode.to_le_bytes().as_slice(),
            value.to_le_bytes().as_slice(),
        ]
        .concat()
    };
    assert!(ParagraphProperties::from_sprm(&invalid_i32(SPRM_P_ITAP, -1)).is_err());
    assert!(ParagraphProperties::from_sprm(&invalid_i32(SPRM_P_DTAP, -1)).is_err());
    assert!(
        ParagraphProperties::from_sprm(
            &[SPRM_P_F_USE_PGSU_SETTINGS.to_le_bytes().as_slice(), &[2]].concat()
        )
        .is_err()
    );

    let inner_at_depth_one = [
        SPRM_P_ITAP.to_le_bytes().as_slice(),
        1i32.to_le_bytes().as_slice(),
        SPRM_P_F_INNER_TABLE_CELL.to_le_bytes().as_slice(),
        &[1],
    ]
    .concat();
    assert!(ParagraphProperties::from_sprm(&inner_at_depth_one).is_err());

    for operand in [
        vec![1, 2, 3, 2, 4, 5, 6, 0, 1, 0],
        vec![1, 2, 3, 0, 4, 5, 6, 0, 0x1A, 0],
        vec![0; 9],
    ] {
        let mut invalid = SPRM_P_SHD.to_le_bytes().to_vec();
        invalid.push(operand.len() as u8);
        invalid.extend_from_slice(&operand);
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }
}

#[test]
fn parses_outline_and_bidi_controls_strictly() {
    let grpprl = [
        SPRM_P_OUT_LVL.to_le_bytes().as_slice(),
        &[9],
        SPRM_P_F_BI_DI.to_le_bytes().as_slice(),
        &[1],
    ]
    .concat();
    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.outline_level, Some(9));
    assert!(properties.bi_directional);

    assert!(
        ParagraphProperties::from_sprm(&[SPRM_P_OUT_LVL.to_le_bytes().as_slice(), &[10]].concat())
            .is_err()
    );
    assert!(
        ParagraphProperties::from_sprm(&[SPRM_P_F_BI_DI.to_le_bytes().as_slice(), &[2]].concat())
            .is_err()
    );
}

#[test]
fn applies_style_level_increments_and_physical_justification_in_order() {
    let heading = [
        SPRM_P_ISTD.to_le_bytes().as_slice(),
        3u16.to_le_bytes().as_slice(),
        SPRM_P_INC_LVL.to_le_bytes().as_slice(),
        &[(-5i8) as u8],
    ]
    .concat();
    let properties = ParagraphProperties::from_sprm(&heading).unwrap();
    assert_eq!(properties.style_index, Some(1));
    assert_eq!(properties.outline_level, Some(0));

    let body = [
        SPRM_P_ISTD.to_le_bytes().as_slice(),
        10u16.to_le_bytes().as_slice(),
        SPRM_P_OUT_LVL.to_le_bytes().as_slice(),
        &[5],
        SPRM_P_INC_LVL.to_le_bytes().as_slice(),
        &[(-10i8) as u8],
    ]
    .concat();
    let properties = ParagraphProperties::from_sprm(&body).unwrap();
    assert_eq!(properties.style_index, Some(10));
    assert_eq!(properties.outline_level, Some(0));

    for (code, physical, normalized) in [
        (
            3,
            PhysicalJustification::LowCompression,
            Justification::Justified,
        ),
        (
            4,
            PhysicalJustification::MediumCompression,
            Justification::MediumKashida,
        ),
        (
            5,
            PhysicalJustification::HighCompression,
            Justification::HighKashida,
        ),
    ] {
        let grpprl = [SPRM_P_JC.to_le_bytes().as_slice(), &[code]].concat();
        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.physical_justification, Some(physical));
        assert_eq!(properties.justification, normalized);
    }

    let logical_supersedes_physical = [
        SPRM_P_JC.to_le_bytes().as_slice(),
        &[5],
        SPRM_P_JC_LOGICAL.to_le_bytes().as_slice(),
        &[4],
    ]
    .concat();
    let properties = ParagraphProperties::from_sprm(&logical_supersedes_physical).unwrap();
    assert_eq!(properties.justification, Justification::Distributed);
    assert_eq!(properties.physical_justification, None);
}

#[test]
fn applies_style_permutations_and_rejects_malformed_spp_operands() {
    let permutation = [
        0, 2, 0, 4, 0, // fLong, first, last
        20, 0, 21, 0, 22, 0, // mappings for styles 2 through 4
    ];
    let mut grpprl = SPRM_P_ISTD.to_le_bytes().to_vec();
    grpprl.extend_from_slice(&3u16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_ISTD_PERMUTE.to_le_bytes());
    grpprl.push(permutation.len() as u8);
    grpprl.extend_from_slice(&permutation);
    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.style_index, Some(21));

    for operand in [
        vec![1, 2, 0, 4, 0, 20, 0, 21, 0, 22, 0],
        vec![0, 4, 0, 2, 0, 20, 0],
        vec![0, 2, 0, 4, 0, 20, 0, 21, 0],
    ] {
        let mut invalid = SPRM_P_ISTD_PERMUTE.to_le_bytes().to_vec();
        invalid.push(operand.len() as u8);
        invalid.extend_from_slice(&operand);
        assert!(ParagraphProperties::from_sprm(&invalid).is_err());
    }
}

#[test]
fn parses_asian_and_frame_controls_strictly() {
    let mut grpprl = Vec::new();
    for opcode in [
        SPRM_P_F_LOCKED,
        SPRM_P_F_WIDOW_CONTROL,
        SPRM_P_F_KINSOKU,
        SPRM_P_F_WORD_WRAP,
        SPRM_P_F_OVERFLOW_PUNCT,
        SPRM_P_F_TOP_LINE_PUNCT,
        SPRM_P_F_AUTO_SPACE_DE,
        SPRM_P_F_AUTO_SPACE_DN,
    ] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(1);
    }
    grpprl.extend_from_slice(&SPRM_P_W_ALIGN_FONT.to_le_bytes());
    grpprl.extend_from_slice(&4u16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_FRAME_TEXT_FLOW.to_le_bytes());
    grpprl.extend_from_slice(&7u16.to_le_bytes());

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert!(properties.locked);
    assert!(properties.widow_control);
    assert!(properties.kinsoku);
    assert!(properties.word_wrap);
    assert!(properties.overflow_punct);
    assert!(properties.top_line_punct);
    assert!(properties.auto_space_de);
    assert!(properties.auto_space_dn);
    assert_eq!(properties.font_align, Some(FontAlignment::Auto));
    assert_eq!(
        properties.frame_text_flow,
        Some(FrameTextFlow {
            vertical: true,
            backwards: true,
            rotate_font: true,
        })
    );

    for opcode in [SPRM_P_F_LOCKED, SPRM_P_F_AUTO_SPACE_DN] {
        assert!(
            ParagraphProperties::from_sprm(&[opcode.to_le_bytes().as_slice(), &[2]].concat())
                .is_err()
        );
    }
    assert!(
        ParagraphProperties::from_sprm(
            &[SPRM_P_W_ALIGN_FONT.to_le_bytes(), 5u16.to_le_bytes()].concat()
        )
        .is_err()
    );
    assert!(
        ParagraphProperties::from_sprm(
            &[SPRM_P_FRAME_TEXT_FLOW.to_le_bytes(), 2u16.to_le_bytes()].concat()
        )
        .is_err()
    );
}

#[test]
fn parses_frame_wrap_height_drop_cap_and_distances_strictly() {
    let mut grpprl = Vec::new();
    grpprl.extend_from_slice(&SPRM_P_WR.to_le_bytes());
    grpprl.push(5);
    grpprl.extend_from_slice(&SPRM_P_W_HEIGHT_ABS.to_le_bytes());
    grpprl.extend_from_slice(&(0x8000u16 | 720).to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DCS.to_le_bytes());
    grpprl.extend_from_slice(&(2u16 | (3u16 << 3)).to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DYA_FROM_TEXT.to_le_bytes());
    grpprl.extend_from_slice(&240i16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DXA_FROM_TEXT.to_le_bytes());
    grpprl.extend_from_slice(&480i16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_F_NO_AUTO_HYPH.to_le_bytes());
    grpprl.push(1);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.text_wrap, Some(FrameTextWrap::Through));
    assert_eq!(
        properties.frame_height,
        Some(FrameHeight {
            height_twips: 720,
            minimum: true,
        })
    );
    assert_eq!(
        properties.drop_cap,
        Some(DropCap {
            kind: DropCapType::Margin,
            lines: 3,
        })
    );
    assert_eq!(properties.dya_from_text, Some(240));
    assert_eq!(properties.dxa_from_text, Some(480));
    assert!(properties.no_auto_hyph);

    let invalid_word =
        |opcode: u16, value: u16| [opcode.to_le_bytes(), value.to_le_bytes()].concat();
    assert!(ParagraphProperties::from_sprm(&[0x23, 0x24, 6]).is_err());
    assert!(ParagraphProperties::from_sprm(&invalid_word(SPRM_P_W_HEIGHT_ABS, 0x8000)).is_err());
    assert!(ParagraphProperties::from_sprm(&invalid_word(SPRM_P_DCS, 3 | (2 << 3))).is_err());
    assert!(ParagraphProperties::from_sprm(&invalid_word(SPRM_P_DCS, 1 | (11 << 3))).is_err());
    assert!(ParagraphProperties::from_sprm(&invalid_word(SPRM_P_DYA_FROM_TEXT, 31_681)).is_err());
    assert!(ParagraphProperties::from_sprm(&invalid_word(SPRM_P_DXA_FROM_TEXT, u16::MAX)).is_err());
    assert!(ParagraphProperties::from_sprm(&[0x2A, 0x24, 2]).is_err());
}

#[test]
fn parses_table_markers_and_frame_positioning_strictly() {
    let mut grpprl = Vec::new();
    for (opcode, value) in [(SPRM_P_F_IN_TABLE, 1), (SPRM_P_F_TTP, 0)] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.push(value);
    }
    grpprl.extend_from_slice(&SPRM_P_DXA_ABS.to_le_bytes());
    grpprl.extend_from_slice(&301i16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DYA_ABS.to_le_bytes());
    grpprl.extend_from_slice(&(-20i16).to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_DXA_WIDTH.to_le_bytes());
    grpprl.extend_from_slice(&31_680u16.to_le_bytes());
    grpprl.extend_from_slice(&SPRM_P_PC.to_le_bytes());
    grpprl.push(0x90);

    let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
    assert!(properties.in_table);
    assert!(!properties.is_table_row_end);
    assert_eq!(
        properties.frame_horizontal_position,
        Some(FrameHorizontalPosition::Offset(300))
    );
    assert_eq!(
        properties.frame_vertical_position,
        Some(FrameVerticalPosition::Outside)
    );
    assert_eq!(properties.frame_width, Some(31_680));
    assert_eq!(
        properties.frame_anchor,
        Some(FrameAnchor {
            vertical: FrameVerticalAnchor::Page,
            horizontal: FrameHorizontalAnchor::Page,
        })
    );

    assert!(ParagraphProperties::from_sprm(&[0x16, 0x24, 2]).is_err());
    assert!(ParagraphProperties::from_sprm(&[0x17, 0x24, 1]).is_err());
    assert!(
        ParagraphProperties::from_sprm(
            &[SPRM_P_DXA_ABS.to_le_bytes(), i16::MIN.to_le_bytes()].concat()
        )
        .is_err()
    );
    assert!(
        ParagraphProperties::from_sprm(
            &[SPRM_P_DXA_WIDTH.to_le_bytes(), 31_681u16.to_le_bytes()].concat()
        )
        .is_err()
    );
    assert!(ParagraphProperties::from_sprm(&[0x1B, 0x26, 0x01]).is_err());
}
