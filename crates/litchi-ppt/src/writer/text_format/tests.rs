#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;

#[test]
fn test_font_style() {
    let style = FontStyle::bold_italic();
    assert!(style.bold);
    assert!(style.italic);
    assert_eq!(style.to_flags(), 0x0003);
}

#[test]
fn test_text_color() {
    // Red = RGB(255, 0, 0) -> PPT format: R | G<<8 | B<<16 | 0xFE<<24 = 0xFE0000FF
    let red = TextColor::RED;
    assert_eq!(red.to_ppt_color(), 0xFE00_00FF);

    // Scheme color index occupies the fourth byte.
    let scheme = TextColor::scheme(4);
    assert_eq!(scheme.to_ppt_color(), 0x0400_0000);
}

#[test]
fn test_text_run() {
    let run = TextRun::new("Hello").bold().size(24);
    assert!(run.style.bold);
    assert_eq!(run.font_size, 24);
    assert_eq!(run.char_count(), 5);
}

#[test]
fn text_run_counts_utf16_code_units() {
    let run = TextRun::new("😀x");
    assert_eq!(run.char_count(), 3);
}

#[test]
fn test_paragraph() {
    let para = Paragraph::new("Test").center();
    assert_eq!(para.alignment, TextAlign::Center);
    assert_eq!(para.char_count(), 5); // 4 chars + 1 end marker
}

#[test]
fn empty_rich_paragraph_has_character_style_coverage() {
    let mut builder = TextPropsBuilder::new();
    builder.add_paragraph(Paragraph::with_runs(Vec::new()));

    let style = builder.build_style_text_prop().unwrap();
    let (paragraphs, characters) = crate::text_prop::parse_style_text_prop_atom(&style, 0);

    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].characters_covered, 1);
    assert_eq!(characters.len(), 1);
    assert_eq!(characters[0].characters_covered, 1);
}

#[test]
fn empty_rich_text_has_complete_style_coverage() {
    let style = TextPropsBuilder::new().build_style_text_prop().unwrap();
    let (paragraphs, characters) = crate::text_prop::parse_style_text_prop_atom(&style, 0);

    assert_eq!(paragraphs[0].characters_covered, 1);
    assert_eq!(characters[0].characters_covered, 1);
}

#[test]
fn rejects_invalid_text_cf_values() {
    let mut invalid_size = TextPropsBuilder::new();
    invalid_size.add_paragraph(Paragraph::with_runs(vec![TextRun::new("x").size(0)]));
    assert!(invalid_size.build_style_text_prop().is_err());

    let mut invalid_scheme = TextPropsBuilder::new();
    invalid_scheme.add_paragraph(Paragraph::with_runs(vec![
        TextRun::new("x").color_scheme(8),
    ]));
    assert!(invalid_scheme.build_style_text_prop().is_err());

    let mut invalid_bullet = TextPropsBuilder::new();
    invalid_bullet.add_paragraph(Paragraph::new("x").with_bullet('😀'));
    assert!(invalid_bullet.build_style_text_prop().is_err());

    let mut invalid_position = TextPropsBuilder::new();
    invalid_position.add_paragraph(Paragraph::with_runs(vec![
        TextRun::new("x").baseline_position(101),
    ]));
    assert!(invalid_position.build_style_text_prop().is_err());

    let mut invalid_indent = TextPropsBuilder::new();
    invalid_indent.add_paragraph(Paragraph::new("x").indent_level(5));
    assert!(invalid_indent.build_style_text_prop().is_err());

    let mut invalid_bullet_size = TextPropsBuilder::new();
    invalid_bullet_size.add_paragraph(Paragraph::new("x").bullet_size(0));
    assert!(invalid_bullet_size.build_style_text_prop().is_err());

    let mut invalid_bullet_scheme = TextPropsBuilder::new();
    invalid_bullet_scheme.add_paragraph(Paragraph::new("x").bullet_color_scheme(8));
    assert!(invalid_bullet_scheme.build_style_text_prop().is_err());

    let mut invalid_pp9_run = TextPropsBuilder::new();
    invalid_pp9_run.add_paragraph(Paragraph::with_runs(vec![TextRun::new("x").pp9_run_id(16)]));
    assert!(invalid_pp9_run.build_style_text_prop().is_err());

    let mut reserved_style = TextRun::new("x");
    reserved_style.style.specified_mask = 0x0008;
    let mut invalid_reserved_style = TextPropsBuilder::new();
    invalid_reserved_style.add_paragraph(Paragraph::with_runs(vec![reserved_style]));
    assert!(invalid_reserved_style.build_style_text_prop().is_err());
}

#[test]
fn character_flags_preserve_values_and_presence() {
    let run = TextRun::new("x")
        .bold_value(false)
        .italic()
        .underline_value(false)
        .shadow()
        .fe_hint(true)
        .kumi(false)
        .strikethrough(true)
        .embossed_value(false)
        .pp9_run_id(13)
        .font(65_535)
        .asian_font(65_534)
        .ansi_font(32_768)
        .symbol_font(60_000);
    let mut builder = TextPropsBuilder::new();
    builder.add_paragraph(Paragraph::with_runs(vec![run]));
    let style = builder.build_style_text_prop().unwrap();
    let (_, character_styles) =
        crate::text_prop::parse_style_text_prop_atom_strict(&style, 1).unwrap();
    assert_eq!(character_styles[0].property_mask & 0xFFFF, 0x3FB7);
    assert_eq!(character_styles[0].get_value("char.flags"), Some(0x3532));
    assert_eq!(character_styles[0].get_value("font.index"), Some(65_535));
    assert_eq!(
        character_styles[0].get_value("asian.font.index"),
        Some(65_534)
    );
    assert_eq!(
        character_styles[0].get_value("ansi.font.index"),
        Some(32_768)
    );
    assert_eq!(
        character_styles[0].get_value("symbol.font.index"),
        Some(60_000)
    );

    let text_record = crate::Record {
        record_type: crate::consts::RecordType::TextBytesAtom,
        record_type_raw: 4008,
        version: 0,
        instance: 0,
        data_length: 1,
        data: b"x".to_vec(),
        children: Vec::new(),
    };
    let style_record = crate::Record {
        record_type: crate::consts::RecordType::StyleTextPropAtom,
        record_type_raw: 4001,
        version: 0,
        instance: 0,
        data_length: u32::try_from(style.len()).unwrap(),
        data: style,
        children: Vec::new(),
    };
    let mut extractor = crate::TextRunExtractor::new();
    extractor
        .extract_from_records(&[text_record, style_record])
        .unwrap();
    let formatting = &extractor.runs()[0].formatting;
    assert_eq!(formatting.font_style_raw, Some(0x3532));
    assert_eq!(formatting.bold_explicit, Some(false));
    assert_eq!(formatting.italic_explicit, Some(true));
    assert_eq!(formatting.underline_explicit, Some(false));
    assert_eq!(formatting.shadow_explicit, Some(true));
    assert_eq!(formatting.fe_hint, Some(true));
    assert_eq!(formatting.kumi, Some(false));
    assert_eq!(formatting.legacy_strikethrough, Some(true));
    assert_eq!(formatting.embossed_explicit, Some(false));
    assert_eq!(formatting.pp9_run_id, Some(13));
    assert_eq!(formatting.font_index, Some(65_535));
    assert_eq!(formatting.asian_font_index, Some(65_534));
    assert_eq!(formatting.ansi_font_index, Some(32_768));
    assert_eq!(formatting.symbol_font_index, Some(60_000));
}

#[test]
fn paragraph_properties_round_trip_in_spec_order() {
    let mut paragraph = Paragraph::new("x")
        .with_bullet('•')
        .bullet_font(65_535)
        .bullet_size(-24)
        .bullet_color_rgb(1, 2, 3)
        .align(TextAlign::Distributed)
        .line_spacing(120)
        .space_before(-10)
        .space_after(20)
        .indent_level(2)
        .default_tab_size(144)
        .tab_stops(vec![
            TabStop::new(-20, TabAlign::Center),
            TabStop::new(720, TabAlign::Decimal),
        ])
        .font_alignment(TextFontAlign::UpholdFixed)
        .character_wrap(true)
        .word_wrap(false)
        .overflow(true)
        .text_direction(TextDirection::RightToLeft);
    paragraph.left_margin = 720;
    paragraph.indent = -360;
    let mut builder = TextPropsBuilder::new();
    builder.add_paragraph(paragraph);

    let style = builder.build_style_text_prop().unwrap();
    let (paragraphs, _) = crate::text_prop::parse_style_text_prop_atom_strict(&style, 1).unwrap();
    let properties = &paragraphs[0];

    assert_eq!(properties.indent_level, 2);
    assert_eq!(properties.get_value("paragraph.flags"), Some(0x000F));
    assert_eq!(properties.get_value("bullet.char"), Some(0x2022));
    assert_eq!(properties.get_value("bullet.font"), Some(65_535));
    assert_eq!(properties.get_value("bullet.size"), Some(-24));
    assert_eq!(
        properties.get_value("bullet.color"),
        Some(0xFE03_0201u32.cast_signed())
    );
    assert_eq!(properties.get_value("alignment"), Some(4));
    assert_eq!(properties.get_value("linespacing"), Some(120));
    assert_eq!(properties.get_value("spacebefore"), Some(-10));
    assert_eq!(properties.get_value("spaceafter"), Some(20));
    assert_eq!(properties.get_value("text.offset"), Some(720));
    assert_eq!(properties.get_value("bullet.offset"), Some(-360));
    assert_eq!(properties.get_value("defaultTabSize"), Some(144));
    assert_eq!(properties.get_value("tabStops"), Some(2));
    assert_eq!(properties.tab_stops[0].position, -20);
    assert_eq!(properties.tab_stops[0].alignment, 1);
    assert_eq!(properties.tab_stops[1].position, 720);
    assert_eq!(properties.tab_stops[1].alignment, 3);
    assert_eq!(properties.get_value("fontAlignment"), Some(3));
    assert_eq!(properties.get_value("wrapFlags"), Some(5));
    assert_eq!(properties.get_value("textDirection"), Some(1));
}

#[test]
fn paragraph_writer_preserves_explicit_false_and_default_values() {
    let paragraph = Paragraph::new("x")
        .bullet_enabled(false)
        .bullet_font_enabled(false)
        .bullet_color_enabled(false)
        .bullet_size_enabled(false)
        .align(TextAlign::Left)
        .line_spacing(100)
        .space_before(0)
        .space_after(0)
        .left_margin(0)
        .first_line_indent(0)
        .tab_stops(Vec::new())
        .character_wrap(false)
        .word_wrap(false)
        .overflow(false)
        .text_direction(TextDirection::LeftToRight);
    let mut builder = TextPropsBuilder::new();
    builder.add_paragraph(paragraph);

    let style = builder.build_style_text_prop().unwrap();
    let (paragraphs, _) = crate::text_prop::parse_style_text_prop_atom_strict(&style, 1).unwrap();
    let properties = &paragraphs[0];

    assert_eq!(properties.get_value("paragraph.flags"), Some(0));
    assert_eq!(properties.get_value("alignment"), Some(0));
    assert_eq!(properties.get_value("linespacing"), Some(100));
    assert_eq!(properties.get_value("spacebefore"), Some(0));
    assert_eq!(properties.get_value("spaceafter"), Some(0));
    assert_eq!(properties.get_value("text.offset"), Some(0));
    assert_eq!(properties.get_value("bullet.offset"), Some(0));
    assert_eq!(properties.get_value("tabStops"), Some(0));
    assert!(properties.tab_stops.is_empty());
    assert_eq!(properties.get_value("wrapFlags"), Some(0));
    assert_eq!(properties.get_value("textDirection"), Some(0));
}

#[test]
fn test_font_entity() {
    let font = FontEntity::arial();
    let data = font.build();
    assert_eq!(data.len(), 68);
    // Check "Arial" in UTF-16LE
    assert_eq!(data[0], b'A');
    assert_eq!(data[1], 0);
    assert_eq!(&data[10..12], &[0, 0]);
    assert_eq!(data[64], 0x00);
    assert_eq!(data[65], 0x00);
    assert_eq!(data[66], 0x04);
    assert_eq!(data[67], 0x22);

    let exact_32 = FontEntity::new("12345678901234567890123456789012").build();
    assert_eq!(&exact_32[62..64], &[b'2', 0]);

    let canonical_31 = FontEntity::new("1234567890123456789012345678901").build();
    assert_eq!(&canonical_31[62..64], &[0, 0]);
}
