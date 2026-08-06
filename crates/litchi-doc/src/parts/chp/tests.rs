use super::super::tap::TableStyleCondition;
use super::*;
use crate::sprm_operations::*;

#[test]
fn parses_insertion_and_deletion_revision_sprms() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut grpprl = Vec::new();
    for (opcode, operand) in [
        (SPRM_C_F_RMARK, vec![1]),
        (SPRM_C_IBST_RMARK, 1u16.to_le_bytes().to_vec()),
        (SPRM_C_DTTM_RMARK, timestamp.to_le_bytes().to_vec()),
        (SPRM_C_IDSL_RMARK, 42u16.to_le_bytes().to_vec()),
        (SPRM_C_RSID_PROP, 0x11223344u32.to_le_bytes().to_vec()),
        (SPRM_C_RSID_TEXT, 0x55667788u32.to_le_bytes().to_vec()),
        (SPRM_C_F_RMARK_DEL, vec![1]),
        (SPRM_C_IBST_RMARK_DEL, 0u16.to_le_bytes().to_vec()),
        (SPRM_C_DTTM_RMARK_DEL, 0u32.to_le_bytes().to_vec()),
        (SPRM_C_IDSL_RMARK_DEL, 7u16.to_le_bytes().to_vec()),
        (SPRM_C_RSID_RM_DEL, 0x99AABBCCu32.to_le_bytes().to_vec()),
    ] {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(&operand);
    }
    let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.is_revision_inserted, Some(true));
    assert_eq!(properties.revision_author_index, Some(1));
    assert_eq!(properties.revision_timestamp, Some(timestamp));
    assert_eq!(properties.revision_id, Some(42));
    assert_eq!(properties.formatting_revision_save_id, Some(0x11223344));
    assert_eq!(properties.insertion_revision_save_id, Some(0x55667788));
    assert_eq!(properties.is_revision_deleted, Some(true));
    assert_eq!(properties.deletion_author_index, Some(0));
    assert_eq!(properties.deletion_timestamp, Some(0));
    assert_eq!(properties.deletion_revision_id, Some(7));
    assert_eq!(properties.deletion_revision_save_id, Some(0x99AABBCC));

    let mut malformed = Vec::new();
    malformed.extend_from_slice(&SPRM_C_IBST_RMARK.to_le_bytes());
    malformed.extend_from_slice(&(-1i16).to_le_bytes());
    assert!(CharacterProperties::from_sprm(&malformed).is_err());

    let mut undefined_reason = SPRM_C_IDSL_RMARK.to_le_bytes().to_vec();
    undefined_reason.extend_from_slice(&0x002Cu16.to_le_bytes());
    assert!(CharacterProperties::from_sprm(&undefined_reason).is_err());
}

#[test]
fn parses_both_character_formatting_revision_sprms_strictly() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    for opcode in [SPRM_C_PROP_RMARK90, SPRM_C_PROP_RMARK_CURRENT] {
        let mut grpprl = opcode.to_le_bytes().to_vec();
        grpprl.push(7);
        grpprl.push(1);
        grpprl.extend_from_slice(&2i16.to_le_bytes());
        grpprl.extend_from_slice(&timestamp.to_le_bytes());
        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.has_formatting_revision, Some(true));
        assert_eq!(properties.formatting_revision_author_index, Some(2));
        assert_eq!(properties.formatting_revision_timestamp, Some(timestamp));
    }

    for operand in [
        vec![2, 0, 0, 0, 0, 0, 0],
        vec![1, 0xFF, 0xFF, 0, 0, 0, 0],
        vec![1, 0, 0, 0, 0, 0],
    ] {
        let mut grpprl = SPRM_C_PROP_RMARK_CURRENT.to_le_bytes().to_vec();
        grpprl.push(operand.len() as u8);
        grpprl.extend_from_slice(&operand);
        assert!(CharacterProperties::from_sprm(&grpprl).is_err());
    }
}

#[test]
fn parses_display_field_revision_strictly() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut operand = [0u8; 39];
    operand[0] = 2; // Any nonzero value means active.
    operand[1..3].copy_from_slice(&2u16.to_le_bytes());
    operand[3..7].copy_from_slice(&timestamp.to_le_bytes());
    operand[7..9].copy_from_slice(&3u16.to_le_bytes());
    for (index, unit) in "12.".encode_utf16().enumerate() {
        let offset = 9 + index * 2;
        operand[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
    }
    let mut grpprl = SPRM_C_DISP_FLD_RMARK.to_le_bytes().to_vec();
    grpprl.push(39);
    grpprl.extend_from_slice(&operand);
    let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
    let revision = properties.display_field_revision.unwrap();
    assert!(revision.active);
    assert_eq!(revision.author_index, 2);
    assert_eq!(revision.timestamp, timestamp);
    assert_eq!(revision.previous_result, "12.");

    for length in [16u8, 38] {
        let mut invalid = operand.to_vec();
        if length == 16 {
            invalid[7..9].copy_from_slice(&16u16.to_le_bytes());
        } else {
            invalid.truncate(38);
        }
        let mut grpprl = SPRM_C_DISP_FLD_RMARK.to_le_bytes().to_vec();
        grpprl.push(length);
        grpprl.extend_from_slice(&invalid);
        assert!(CharacterProperties::from_sprm(&grpprl).is_err());
    }
}

#[test]
fn test_default_chp() {
    let chp = CharacterProperties::new();
    assert_eq!(chp.is_bold, None);
    assert_eq!(chp.is_italic, None);
    assert_eq!(chp.underline, UnderlineStyle::None);
    assert!(!chp.has_formatting());
}

#[test]
fn preserves_ordered_character_revision_state() {
    let mut grpprl = SPRM_C_F_BOLD.to_le_bytes().to_vec();
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_C_WALL.to_le_bytes());
    grpprl.push(1);
    grpprl.extend_from_slice(&SPRM_C_F_ITALIC.to_le_bytes());
    grpprl.push(1);

    let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
    assert_eq!(properties.is_bold, Some(true));
    assert_eq!(properties.is_italic, Some(true));
    assert!(properties.properties_preserved_for_revision);
    let previous = properties.preserved_properties_for_revision.unwrap();
    assert_eq!(previous.is_bold, Some(true));
    assert_eq!(previous.is_italic, None);

    grpprl.extend_from_slice(&SPRM_C_WALL.to_le_bytes());
    grpprl.push(0);
    let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
    assert!(!properties.properties_preserved_for_revision);
    assert!(properties.preserved_properties_for_revision.is_none());

    let invalid = [SPRM_C_WALL.to_le_bytes().as_slice(), &[2]].concat();
    assert!(CharacterProperties::from_sprm(&invalid).is_err());
}

#[test]
fn parses_conditional_table_style_character_formatting_strictly() {
    let wrap = |condition: u16, nested: &[u8]| {
        let mut grpprl = SPRM_C_CNF.to_le_bytes().to_vec();
        grpprl.push((nested.len() + 2) as u8);
        grpprl.extend_from_slice(&condition.to_le_bytes());
        grpprl.extend_from_slice(nested);
        grpprl
    };
    let nested = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
    let properties = CharacterProperties::from_sprm(&wrap(0x0008, &nested)).unwrap();
    let conditional = &properties.conditional_formats[0];
    assert_eq!(conditional.condition, TableStyleCondition::LastColumn);
    assert_eq!(conditional.raw_grpprl, nested);
    assert_eq!(conditional.properties.is_bold, Some(true));

    let recursive = wrap(0x0002, &[]);
    let paragraph = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
    let truncated = SPRM_C_F_BOLD.to_le_bytes();
    for invalid in [
        [SPRM_C_CNF.to_le_bytes().as_slice(), &[0]].concat(),
        wrap(0x0003, &[]),
        wrap(0x0001, &recursive),
        wrap(0x0001, &paragraph),
        wrap(0x0001, &truncated),
    ] {
        assert!(CharacterProperties::from_sprm(&invalid).is_err());
    }
}

#[test]
fn test_underline_style() {
    let single = UnderlineStyle::Single;
    let double = UnderlineStyle::Double;
    assert_ne!(single, double);
    assert_eq!(single, UnderlineStyle::Single);
}

#[test]
fn test_vertical_position() {
    let normal = VerticalPosition::Normal;
    let super_pos = VerticalPosition::Superscript;
    assert_ne!(normal, super_pos);
}

#[test]
fn test_toggle_value() {
    // Test basic values
    assert!(!CharacterProperties::get_toggle_value(0, None));
    assert!(CharacterProperties::get_toggle_value(1, None));

    // Test preserve old value
    assert!(CharacterProperties::get_toggle_value(0x80, Some(true)));
    assert!(!CharacterProperties::get_toggle_value(0x80, Some(false)));

    // Test toggle old value
    assert!(!CharacterProperties::get_toggle_value(0x81, Some(true)));
    assert!(CharacterProperties::get_toggle_value(0x81, Some(false)));
}

#[test]
fn ignores_non_character_sprms_in_mixed_piece_modifier() {
    let properties = CharacterProperties::from_sprm(&[
        0x03, 0x24, 0x02, // paragraph justification
        0x35, 0x08, 0x01, // character bold
    ])
    .unwrap();
    assert_eq!(properties.is_bold, Some(true));
    assert_eq!(properties.is_strikethrough, None);
}

#[test]
fn parses_complex_script_language_and_proofing_sprms() {
    let properties = CharacterProperties::from_sprm(&[
        0x5A, 0x08, 0x01, // sprmCFBiDi
        0x5C, 0x08, 0x01, // sprmCFBoldBi
        0x5C, 0x08, 0x81, // toggle complex-script bold back off
        0x5D, 0x08, 0x01, // sprmCFItalicBi
        0x5E, 0x4A, 0x34, 0x12, // sprmCFtcBi
        0x5F, 0x48, 0x01, 0x04, // sprmCLidBi
        0x60, 0x4A, 0x0D, 0x00, // sprmCIcoBi
        0x61, 0x4A, 0x1C, 0x00, // sprmCHpsBi
        0x6D, 0x48, 0x09, 0x04, // sprmCRgLid0_80
        0x6E, 0x48, 0x11, 0x04, // sprmCRgLid1_80
        0x73, 0x48, 0x0C, 0x04, // sprmCRgLid0 supersedes legacy
        0x74, 0x48, 0x12, 0x04, // sprmCRgLid1 supersedes legacy
        0x6F, 0x28, 0x02, // sprmCIdctHint
        0x75, 0x08, 0x01, // sprmCFNoProof
    ])
    .unwrap();

    assert_eq!(properties.is_bidi, Some(true));
    assert_eq!(properties.is_bold_bidi, Some(false));
    assert_eq!(properties.is_italic_bidi, Some(true));
    assert_eq!(properties.font_index_bidi, Some(0x1234));
    assert_eq!(properties.language_id_bidi, Some(0x0401));
    assert_eq!(properties.color_index_bidi, Some(13));
    assert_eq!(properties.font_size_bidi, Some(28));
    assert_eq!(properties.language_id, Some(0x040C));
    assert_eq!(properties.language_id_fe, Some(0x0412));
    assert_eq!(
        properties.script_hint,
        Some(CharacterScriptHint::ComplexScript)
    );
    assert_eq!(properties.is_no_proof, Some(true));
    assert!(properties.has_formatting());
}

#[test]
fn preserves_reserved_script_hint_values() {
    let properties = CharacterProperties::from_sprm(&[0x6F, 0x28, 0xFF]).unwrap();
    assert_eq!(
        properties.script_hint,
        Some(CharacterScriptHint::Reserved(0xFF))
    );
}

#[test]
fn parses_palette_character_border_and_shading() {
    let properties = CharacterProperties::from_sprm(&[
        0x65, 0x68, // sprmCBrc80
        0x10, 0x14, 0x06, 0x65, // 2pt wave, red, 5pt, shadow and frame
        0x66, 0x48, // sprmCShd80
        0xE6, 0x04, // solid red foreground on yellow
    ])
    .unwrap();

    assert_eq!(
        properties.border,
        Some(CharacterBorder {
            color: CharacterColor::Rgb(255, 0, 0),
            width: 0x10,
            style: CharacterBorderStyle::Wave,
            spacing: 5,
            has_shadow: true,
            has_frame: true,
        })
    );
    assert_eq!(
        properties.shading,
        Some(CharacterShading {
            foreground_color: CharacterColor::Rgb(255, 0, 0),
            background_color: CharacterColor::Rgb(255, 255, 0),
            pattern: CharacterShadingPattern::Solid,
        })
    );
}

#[test]
fn parses_rgb_character_border_shading_and_underline() {
    let properties = CharacterProperties::from_sprm(&[
        0x71, 0xCA, // sprmCShd
        0x0A, // SHDOperand byte count
        0x12, 0x34, 0x56, 0x00, // foreground COLORREF
        0x00, 0x00, 0x00, 0xFF, // automatic background COLORREF
        0x25, 0x00, // 12.5 percent pattern
        0x72, 0xCA, // sprmCBrc
        0x08, // BrcOperand byte count
        0xAA, 0xBB, 0xCC, 0x00, // border COLORREF
        0x10, 0x18, 0x43, 0x00, // width, emboss, 3pt, frame
        0x77, 0x68, // sprmCCvUl
        0x01, 0x02, 0x03, 0x00, // underline COLORREF
        0x82, 0x08, 0x01, // sprmCFComplexScripts
    ])
    .unwrap();

    assert_eq!(
        properties.shading,
        Some(CharacterShading {
            foreground_color: CharacterColor::Rgb(0x12, 0x34, 0x56),
            background_color: CharacterColor::Automatic,
            pattern: CharacterShadingPattern::Percent12_5,
        })
    );
    assert_eq!(
        properties.border,
        Some(CharacterBorder {
            color: CharacterColor::Rgb(0xAA, 0xBB, 0xCC),
            width: 0x10,
            style: CharacterBorderStyle::ThreeDEmboss,
            spacing: 3,
            has_shadow: false,
            has_frame: true,
        })
    );
    assert_eq!(
        properties.underline_color,
        Some(CharacterColor::Rgb(1, 2, 3))
    );
    assert_eq!(properties.is_complex_scripts, Some(true));
    assert!(properties.has_formatting());
}

#[test]
fn preserves_reserved_character_format_values() {
    let properties = CharacterProperties::from_sprm(&[
        0x65, 0x68, // sprmCBrc80
        0x02, 0x02, 0x11, 0x00, // reserved style and palette index
        0x71, 0xCA, // sprmCShd
        0x0A, 0x01, 0x02, 0x03, 0x00, 0x04, 0x05, 0x06, 0x00, 0x1A, 0x00,
    ])
    .unwrap();

    let border = properties.border.unwrap();
    assert_eq!(border.style, CharacterBorderStyle::Reserved(0x02));
    assert_eq!(border.color, CharacterColor::ReservedPaletteIndex(0x11));
    assert_eq!(
        properties.shading.unwrap().pattern,
        CharacterShadingPattern::Reserved(0x001A)
    );
}

#[cfg(test)]
mod chpx_position_hresi_effect_tests {
    use super::*;

    fn append(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(operand);
    }

    #[test]
    fn decodes_all_values_boundaries_defaults_and_later_wins() {
        let modes = [
            (HyphenationMode::Normal, [1, 0]),
            (HyphenationMode::AddBefore, [2, b'A']),
            (HyphenationMode::ChangeBefore, [3, b'B']),
            (HyphenationMode::DeleteBefore, [4, b' ']),
            (HyphenationMode::ChangeAfter, [5, b'Y']),
            (HyphenationMode::DeleteAndChange, [6, b'Z']),
        ];
        for (mode, bytes) in modes {
            let mut grpprl = Vec::new();
            append(&mut grpprl, SPRM_C_HRESI, &bytes);
            let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
            assert_eq!(properties.hyphenation.mode(), mode);
            assert_eq!(properties.hyphenation.bytes(), bytes);
        }
        for (raw, effect) in [
            (0, TextEffect::None),
            (1, TextEffect::LasVegasLights),
            (2, TextEffect::BlinkingBackground),
            (3, TextEffect::SparkleText),
            (4, TextEffect::MarchingBlackAnts),
            (5, TextEffect::MarchingRedAnts),
            (6, TextEffect::Shimmer),
        ] {
            let mut grpprl = Vec::new();
            append(&mut grpprl, SPRM_C_SFXT_TEXT, &[raw]);
            assert_eq!(
                CharacterProperties::from_sprm(&grpprl).unwrap().text_effect,
                effect
            );
        }

        let mut grpprl = Vec::new();
        append(&mut grpprl, SPRM_C_HPS_POS, &(-3168i16).to_le_bytes());
        append(&mut grpprl, SPRM_C_HPS_POS, &3168i16.to_le_bytes());
        append(&mut grpprl, SPRM_C_HRESI, &[1, 0]);
        append(&mut grpprl, SPRM_C_HRESI, &[5, b'Q']);
        append(&mut grpprl, SPRM_C_SFXT_TEXT, &[1]);
        append(&mut grpprl, SPRM_C_SFXT_TEXT, &[6]);
        let properties = CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.position.half_points(), 3168);
        assert_eq!(properties.hyphenation.mode(), HyphenationMode::ChangeAfter);
        assert_eq!(properties.hyphenation.replacement_character(), Some(b'Q'));
        assert_eq!(properties.text_effect, TextEffect::Shimmer);
        assert!(properties.has_formatting());

        let defaults = CharacterProperties::default();
        assert_eq!(defaults.position, CharacterPosition::NORMAL);
        assert_eq!(defaults.hyphenation, HresiOperand::normal());
        assert_eq!(defaults.text_effect, TextEffect::None);
    }

    #[test]
    fn rejects_out_of_range_and_dependent_operands() {
        assert!(CharacterPosition::new(-3169).is_err());
        assert!(CharacterPosition::new(3169).is_err());
        assert!(CharacterPosition::new(-3168).is_ok());
        assert!(CharacterPosition::new(3168).is_ok());
        assert!(HresiOperand::with_character(HyphenationMode::Normal, b'A').is_err());
        for byte in [0x00, 0x1F, 0x7F, 0x80, 0xFF] {
            assert!(HresiOperand::with_character(HyphenationMode::AddBefore, byte).is_err());
        }

        for (opcode, operand) in [
            (SPRM_C_HPS_POS, 3169i16.to_le_bytes()),
            (SPRM_C_HRESI, [0, 0]),
            (SPRM_C_HRESI, [7, b'A']),
            (SPRM_C_HRESI, [1, b'A']),
            (SPRM_C_HRESI, [2, 0]),
        ] {
            let mut grpprl = Vec::new();
            append(&mut grpprl, opcode, &operand);
            assert!(CharacterProperties::from_sprm(&grpprl).is_err());
        }
        let mut grpprl = Vec::new();
        append(&mut grpprl, SPRM_C_SFXT_TEXT, &[7]);
        assert!(CharacterProperties::from_sprm(&grpprl).is_err());
    }
}
