#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

mod tests {
    use super::super::*;

    // =============================================================================
    // TxMasterStyleAtom Instance Type Tests
    // =============================================================================

    #[test]
    fn test_tx_style_instance_values() {
        assert_eq!(tx_style_instance::TITLE, 0);
        assert_eq!(tx_style_instance::BODY, 1);
        assert_eq!(tx_style_instance::NOTES, 2);
        assert_eq!(tx_style_instance::OTHER, 4);
        assert_eq!(tx_style_instance::CENTER_BODY, 5);
        assert_eq!(tx_style_instance::CENTER_TITLE, 6);
        assert_eq!(tx_style_instance::HALF_BODY, 7);
        assert_eq!(tx_style_instance::QUARTER_BODY, 8);
    }

    // =============================================================================
    // ParagraphMask Tests
    // =============================================================================

    #[test]
    fn test_paragraph_mask_values() {
        assert_eq!(ParagraphMask::HAS_BULLET.bits(), 0x0001);
        assert_eq!(ParagraphMask::BULLET_HAS_FONT.bits(), 0x0002);
        assert_eq!(ParagraphMask::BULLET_HAS_COLOR.bits(), 0x0004);
        assert_eq!(ParagraphMask::BULLET_HAS_SIZE.bits(), 0x0008);
        assert_eq!(ParagraphMask::BULLET_FONT.bits(), 0x0010);
        assert_eq!(ParagraphMask::BULLET_COLOR.bits(), 0x0020);
        assert_eq!(ParagraphMask::BULLET_SIZE.bits(), 0x0040);
        assert_eq!(ParagraphMask::BULLET_CHAR.bits(), 0x0080);
        assert_eq!(ParagraphMask::LEFT_MARGIN.bits(), 0x0100);
        assert_eq!(ParagraphMask::INDENT.bits(), 0x0400);
        assert_eq!(ParagraphMask::ALIGN.bits(), 0x0800);
        assert_eq!(ParagraphMask::LINE_SPACING.bits(), 0x1000);
        assert_eq!(ParagraphMask::SPACE_BEFORE.bits(), 0x2000);
        assert_eq!(ParagraphMask::SPACE_AFTER.bits(), 0x4000);
        assert_eq!(ParagraphMask::DEFAULT_TAB_SIZE.bits(), 0x8000);
        assert_eq!(ParagraphMask::FONT_ALIGN.bits(), 0x0001_0000);
        assert_eq!(ParagraphMask::WRAP_FLAGS.bits(), 0x0002_0000);
        assert_eq!(ParagraphMask::TEXT_DIRECTION.bits(), 0x0004_0000);
    }

    #[test]
    fn test_pf_mask_module_values() {
        assert_eq!(pf_mask::HAS_BULLET, 0x0001);
        assert_eq!(pf_mask::BULLET_HAS_FONT, 0x0002);
        assert_eq!(pf_mask::LEFT_MARGIN, 0x0100);
        assert_eq!(pf_mask::INDENT, 0x0400);
        assert_eq!(pf_mask::ALIGN, 0x0800);
    }

    // =============================================================================
    // CharacterMask Tests
    // =============================================================================

    #[test]
    fn test_character_mask_values() {
        assert_eq!(CharacterMask::BOLD.bits(), 0x0001);
        assert_eq!(CharacterMask::ITALIC.bits(), 0x0002);
        assert_eq!(CharacterMask::UNDERLINE.bits(), 0x0004);
        assert_eq!(CharacterMask::SHADOW.bits(), 0x0010);
        assert_eq!(CharacterMask::FEHINT.bits(), 0x0020);
        assert_eq!(CharacterMask::KUMI.bits(), 0x0080);
        assert_eq!(CharacterMask::EMBOSS.bits(), 0x0200);
        assert_eq!(CharacterMask::STYLE_INDEX.bits(), 0x0800);
        assert_eq!(CharacterMask::HAS_SCHEME_COLOR.bits(), 0x1000);
        assert_eq!(CharacterMask::HAS_SHADOW_COLOR.bits(), 0x2000);
    }

    #[test]
    fn test_cf_mask_module_values() {
        assert_eq!(cf_mask::BOLD, 0x0001);
        assert_eq!(cf_mask::ITALIC, 0x0002);
        assert_eq!(cf_mask::UNDERLINE, 0x0004);
        assert_eq!(cf_mask::SHADOW, 0x0010);
        assert_eq!(cf_mask::EMBOSS, 0x0200);
        assert_eq!(cf_mask::STYLE_INDEX, 0x0800);
    }

    // =============================================================================
    // Font Size Tests
    // =============================================================================

    #[test]
    fn test_font_size_values() {
        assert_eq!(font_size::PT_44, 4400);
        assert_eq!(font_size::PT_32, 3200);
        assert_eq!(font_size::PT_28, 2800);
        assert_eq!(font_size::PT_24, 2400);
        assert_eq!(font_size::PT_20, 2000);
        assert_eq!(font_size::PT_18, 1800);
        assert_eq!(font_size::PT_16, 1600);
        assert_eq!(font_size::PT_14, 1400);
        assert_eq!(font_size::PT_12, 1200);
    }

    // =============================================================================
    // Indent Tests
    // =============================================================================

    #[test]
    fn test_indent_values() {
        assert_eq!(indent::LEVEL_0, 0x0000);
        assert_eq!(indent::LEVEL_1, 0x0120); // 288 = 0.5 inch
        assert_eq!(indent::LEVEL_2, 0x0240); // 576 = 1.0 inch
        assert_eq!(indent::LEVEL_3, 0x0360); // 864 = 1.5 inch
        assert_eq!(indent::LEVEL_4, 0x0480); // 1152 = 2.0 inch
    }

    // =============================================================================
    // Bullet Tests
    // =============================================================================

    #[test]
    fn test_bullet_values() {
        assert_eq!(bullet::FLAGS_DEFAULT, 0x2022);
        assert_eq!(bullet::FONT_INDEX, 0x0064);
        assert_eq!(bullet::CHAR_NONE, 0x0000);
        assert_eq!(bullet::SIZE_DEFAULT, 0x0000);
        assert_eq!(bullet::COLOR_SCHEME, 0x0001_FF00);
        assert_eq!(bullet::COLOR_ALPHA, 0x0000_FF00);
    }

    // =============================================================================
    // Alignment Tests
    // =============================================================================

    #[test]
    fn test_align_values() {
        assert_eq!(align::LEFT, 0x0000);
        assert_eq!(align::CENTER, 0x0001);
        assert_eq!(align::RIGHT, 0x0002);
        assert_eq!(align::JUSTIFY, 0x0003);
        assert_eq!(align::DEFAULT, 0x0064);
    }

    // =============================================================================
    // Spacing Tests
    // =============================================================================

    #[test]
    fn test_spacing_values() {
        assert_eq!(spacing::DEFAULT_LINE, 0x0000);
        assert_eq!(spacing::LINE_120_PCT, 0x0014);
        assert_eq!(spacing::LINE_150_PCT, 0x001E);
        assert_eq!(spacing::SPACE_AFTER_216, 0x00D8);
    }

    // =============================================================================
    // Tab Tests
    // =============================================================================

    #[test]
    fn test_tab_values() {
        assert_eq!(tab::DEFAULT_SIZE, 0x0240); // 576 = 1 inch
    }

    // =============================================================================
    // Font Flags Tests
    // =============================================================================

    #[test]
    fn test_font_flags_values() {
        assert_eq!(font_flags::NONE, 0x0000);
        assert_eq!(font_flags::INHERIT, 0xFFFF);
    }

    // =============================================================================
    // Position Tests
    // =============================================================================

    #[test]
    fn test_position_values() {
        assert_eq!(position::TITLE_DEFAULT, 0x002C); // 44
        assert_eq!(position::BODY_DEFAULT, 0x0020); // 32
        assert_eq!(position::NOTES_DEFAULT, 0x000C); // 12
        assert_eq!(position::OTHER_DEFAULT, 0x0012); // 18
    }

    // =============================================================================
    // SimpleLevelEntry Tests
    // =============================================================================

    #[test]
    fn test_simple_level_entry_new() {
        let entry = SimpleLevelEntry::new(0x1234_5678, 0xABCD, 0x1234);
        // Copy to local variables to avoid unaligned reference warnings
        let pf_mask = entry.pf_mask;
        let cf_mask = entry.cf_mask;
        let font_size = entry.font_size;
        assert_eq!(pf_mask, 0x1234_5678);
        assert_eq!(cf_mask, 0xABCD);
        assert_eq!(font_size, 0x1234);
    }

    #[test]
    fn test_simple_level_entry_empty() {
        let entry = SimpleLevelEntry::EMPTY;
        let pf_mask = entry.pf_mask;
        let cf_mask = entry.cf_mask;
        let font_size = entry.font_size;
        assert_eq!(pf_mask, 0);
        assert_eq!(cf_mask, 0);
        assert_eq!(font_size, 0);
    }

    #[test]
    fn test_simple_level_entry_with_font_size() {
        let entry = SimpleLevelEntry::with_font_size(2400);
        let pf_mask = entry.pf_mask;
        let cf_mask = entry.cf_mask;
        let font_size = entry.font_size;
        assert_eq!(pf_mask, 0);
        assert_eq!(cf_mask, cf_mask::BOLD);
        assert_eq!(font_size, 2400);
    }

    // =============================================================================
    // IndentedLevelEntry Tests
    // =============================================================================

    #[test]
    fn test_indented_level_entry_new() {
        let entry = IndentedLevelEntry::new(0x0120, 0x0120);
        let pf_mask = entry.pf_mask;
        let left_margin = entry.left_margin;
        let indent = entry.indent;
        let cf_mask = entry.cf_mask;
        let cf_flags = entry.cf_flags;
        assert_eq!(pf_mask, pf_mask::LEFT_MARGIN | pf_mask::INDENT);
        assert_eq!(left_margin, 0x0120);
        assert_eq!(indent, 0x0120);
        assert_eq!(cf_mask, 0);
        assert_eq!(cf_flags, 0);
    }

    // =============================================================================
    // TxMasterStyleBuilder Tests
    // =============================================================================

    #[test]
    fn test_tx_master_style_builder_new() {
        let builder = TxMasterStyleBuilder::new(5);
        let data = builder.build();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    #[test]
    fn generic_master_style_builder_round_trips_strictly() {
        let mut builder = TxMasterStyleBuilder::for_text_type(5, 1);
        builder.add_simple_level(&SimpleStyleLevel {
            pf_mask: pf_mask::LEFT_MARGIN | pf_mask::INDENT,
            left_margin: 720,
            indent: 360,
            cf_mask: u32::from(cf_mask::BOLD),
            cf_flags: cf_mask::BOLD,
            font_size: 24,
            has_font_size: true,
        });

        let style = crate::TextMasterStyle::parse(&builder.build(), 5).unwrap();

        assert_eq!(style.levels[0].explicit_level, Some(0));
        assert_eq!(
            style.levels[0].paragraph.get_value("text.offset"),
            Some(720)
        );
        assert_eq!(
            style.levels[0].paragraph.get_value("bullet.offset"),
            Some(360)
        );
        assert_eq!(style.levels[0].character.get_value("char.flags"), Some(1));
        assert_eq!(style.levels[0].character.get_value("font.size"), Some(24));
    }

    #[test]
    fn prebuilt_master_styles_parse_strictly() {
        for (text_type, data) in [
            (0, build_tx_master_style_title()),
            (1, build_tx_master_style_body()),
            (2, build_tx_master_style_notes()),
            (4, build_tx_master_style_other()),
            (5, build_tx_master_style_center_body()),
            (6, build_tx_master_style_center_title()),
            (7, build_tx_master_style_half_body()),
            (8, build_tx_master_style_quarter_body()),
        ] {
            crate::TextMasterStyle::parse(&data, text_type)
                .unwrap_or_else(|error| panic!("text type {text_type}: {error}"));
        }
    }

    // =============================================================================
    // Body Bullet Tests
    // =============================================================================

    #[test]
    fn test_body_bullet_levels() {
        let (char1, flags1) = body_bullet::LEVEL_1;
        assert_eq!(char1, 0x2013); // dash bullet
        assert_eq!(flags1, 0x01D4);

        let (char2, flags2) = body_bullet::LEVEL_2;
        assert_eq!(char2, 0x2022); // round bullet
        assert_eq!(flags2, 0x02D0);

        let (char3, flags3) = body_bullet::LEVEL_3;
        assert_eq!(char3, 0x2013); // dash bullet
        assert_eq!(flags3, 0x03F0);

        let (char4, flags4) = body_bullet::LEVEL_4;
        assert_eq!(char4, 0x00BB); // right angle bullet
        assert_eq!(flags4, 0x0510);
    }

    // =============================================================================
    // Font Size Arrays Tests
    // =============================================================================

    #[test]
    fn test_half_body_font_sizes() {
        assert_eq!(HALF_BODY_FONT_SIZES.len(), 5);
        assert_eq!(HALF_BODY_FONT_SIZES[0], font_size::PT_28);
        assert_eq!(HALF_BODY_FONT_SIZES[1], font_size::PT_24);
        assert_eq!(HALF_BODY_FONT_SIZES[2], font_size::PT_20);
        assert_eq!(HALF_BODY_FONT_SIZES[3], font_size::PT_18);
        assert_eq!(HALF_BODY_FONT_SIZES[4], font_size::PT_18);
    }

    #[test]
    fn test_quarter_body_font_sizes() {
        assert_eq!(QUARTER_BODY_FONT_SIZES.len(), 5);
        assert_eq!(QUARTER_BODY_FONT_SIZES[0], font_size::PT_24);
        assert_eq!(QUARTER_BODY_FONT_SIZES[1], font_size::PT_20);
        assert_eq!(QUARTER_BODY_FONT_SIZES[2], font_size::PT_18);
        assert_eq!(QUARTER_BODY_FONT_SIZES[3], font_size::PT_16);
        assert_eq!(QUARTER_BODY_FONT_SIZES[4], font_size::PT_16);
    }

    // =============================================================================
    // TxMasterStyle Constants Tests
    // =============================================================================

    #[test]
    fn test_tx_master_style_title_length() {
        assert_eq!(TX_MASTER_STYLE_TITLE.len(), 62);
        // First two bytes should be level count (1)
        assert_eq!(TX_MASTER_STYLE_TITLE[0], 0x01);
        assert_eq!(TX_MASTER_STYLE_TITLE[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_body_length() {
        assert_eq!(TX_MASTER_STYLE_BODY.len(), 124);
        // First two bytes should be level count (5)
        assert_eq!(TX_MASTER_STYLE_BODY[0], 0x05);
        assert_eq!(TX_MASTER_STYLE_BODY[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_notes_length() {
        assert_eq!(TX_MASTER_STYLE_NOTES.len(), 110);
        // First two bytes should be level count (5)
        assert_eq!(TX_MASTER_STYLE_NOTES[0], 0x05);
        assert_eq!(TX_MASTER_STYLE_NOTES[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_other_length() {
        assert_eq!(TX_MASTER_STYLE_OTHER.len(), 110);
        // First two bytes should be level count (5)
        assert_eq!(TX_MASTER_STYLE_OTHER[0], 0x05);
        assert_eq!(TX_MASTER_STYLE_OTHER[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_center_body_length() {
        assert_eq!(TX_MASTER_STYLE_CENTER_BODY.len(), 82);
        // First two bytes should be level count (5)
        assert_eq!(TX_MASTER_STYLE_CENTER_BODY[0], 0x05);
        assert_eq!(TX_MASTER_STYLE_CENTER_BODY[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_center_title_length() {
        assert_eq!(TX_MASTER_STYLE_CENTER_TITLE.len(), 12);
        // First two bytes should be level count (1)
        assert_eq!(TX_MASTER_STYLE_CENTER_TITLE[0], 0x01);
        assert_eq!(TX_MASTER_STYLE_CENTER_TITLE[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_half_body_length() {
        assert_eq!(TX_MASTER_STYLE_HALF_BODY.len(), 62);
        // First two bytes should be level count (5)
        assert_eq!(TX_MASTER_STYLE_HALF_BODY[0], 0x05);
        assert_eq!(TX_MASTER_STYLE_HALF_BODY[1], 0x00);
    }

    #[test]
    fn test_tx_master_style_quarter_body_length() {
        assert_eq!(TX_MASTER_STYLE_QUARTER_BODY.len(), 62);
        // First two bytes should be level count (5)
        assert_eq!(TX_MASTER_STYLE_QUARTER_BODY[0], 0x05);
        assert_eq!(TX_MASTER_STYLE_QUARTER_BODY[1], 0x00);
    }

    // =============================================================================
    // Style Function Tests
    // =============================================================================

    #[test]
    fn test_tx_master_style_title_fn() {
        let data = tx_master_style_title();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 1);
    }

    #[test]
    fn test_tx_master_style_body_fn() {
        let data = tx_master_style_body();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    #[test]
    fn test_tx_master_style_notes_fn() {
        let data = tx_master_style_notes();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    #[test]
    fn test_tx_master_style_other_fn() {
        let data = tx_master_style_other();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    #[test]
    fn test_tx_master_style_center_body_fn() {
        let data = tx_master_style_center_body();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    #[test]
    fn test_tx_master_style_center_title_fn() {
        let data = tx_master_style_center_title();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 1);
    }

    #[test]
    fn test_tx_master_style_half_body_fn() {
        let data = tx_master_style_half_body();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    #[test]
    fn test_tx_master_style_quarter_body_fn() {
        let data = tx_master_style_quarter_body();
        assert!(!data.is_empty());
        assert_eq!(u16::from_le_bytes([data[0], data[1]]), 5);
    }

    // =============================================================================
    // PF/CF Mask Constants Tests
    // =============================================================================

    #[test]
    fn test_title_pf_mask() {
        assert_eq!(TITLE_PF_MASK, 0x003F_FDFF);
    }

    #[test]
    fn test_title_cf_mask() {
        assert_eq!(TITLE_CF_MASK, 0x0007);
    }

    #[test]
    fn test_body_level_pf_mask() {
        assert_eq!(BODY_LEVEL_PF_MASK, 0x0580);
    }

    #[test]
    fn test_indent_only_pf_mask() {
        assert_eq!(INDENT_ONLY_PF_MASK, pf_mask::LEFT_MARGIN | pf_mask::INDENT);
    }

    #[test]
    fn test_center_body_pf_mask() {
        assert_eq!(CENTER_BODY_PF_MASK, 0x0901);
    }

    #[test]
    fn test_build_tx_master_style_title() {
        let built = build_tx_master_style_title();
        let style = crate::TextMasterStyle::parse(&built, 0).unwrap();
        assert_eq!(style.levels.len(), 1);
        assert_eq!(style.levels[0].paragraph.get_value("alignment"), Some(0));
        assert_eq!(style.levels[0].character.get_value("font.size"), Some(44));
    }

    #[test]
    fn test_build_tx_master_style_center_title() {
        let built = build_tx_master_style_center_title();
        let style = crate::TextMasterStyle::parse(&built, 6).unwrap();
        assert_eq!(style.levels.len(), 1);
        assert_eq!(style.levels[0].explicit_level, Some(0));
        assert_eq!(style.levels[0].paragraph.get_value("alignment"), Some(1));
        assert_eq!(style.levels[0].character.get_value("font.size"), Some(44));
    }
}
