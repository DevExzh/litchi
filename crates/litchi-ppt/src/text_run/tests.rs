use super::*;

#[test]
fn test_text_run_creation() {
    let run = TextRun::new("Hello".to_string(), 0);
    assert_eq!(run.text, "Hello");
    assert_eq!(run.start_index, 0);
    assert_eq!(run.length, 5);
}

#[test]
fn test_text_run_extractor() {
    let mut extractor = TextRunExtractor::new();

    // Create a simple TextCharsAtom record
    let text_data = vec![
        0x48, 0x00, // 'H'
        0x65, 0x00, // 'e'
        0x6C, 0x00, // 'l'
        0x6C, 0x00, // 'l'
        0x6F, 0x00, // 'o'
        0x00, 0x00, // null terminator
    ];

    let record = Record {
        record_type: RecordType::TextCharsAtom,
        record_type_raw: 4000,
        version: 0,
        instance: 0,
        data_length: text_data.len() as u32,
        data: text_data,
        children: Vec::new(),
    };

    extractor.extract_from_records(&[record]).unwrap();
    assert_eq!(extractor.text(), "Hello");
    assert_eq!(extractor.run_count(), 1);
}

#[test]
fn text_run_extractor_uses_ppt_unicode_encodings_and_character_offsets() {
    let unicode_record = Record {
        record_type: RecordType::TextCharsAtom,
        record_type_raw: 4000,
        version: 0,
        instance: 0,
        data_length: 4,
        data: vec![0x3D, 0xD8, 0x00, 0xDE],
        children: Vec::new(),
    };
    let byte_record = Record {
        record_type: RecordType::TextBytesAtom,
        record_type_raw: 4008,
        version: 0,
        instance: 0,
        data_length: 2,
        data: vec![0x80, 0xE9],
        children: Vec::new(),
    };
    let mut extractor = TextRunExtractor::new();

    extractor
        .extract_from_records(&[unicode_record, byte_record])
        .unwrap();

    assert_eq!(extractor.text(), "😀\u{80}é");
    assert_eq!(extractor.runs()[0].start_index, 0);
    assert_eq!(extractor.runs()[0].length, 1);
    assert_eq!(extractor.runs()[1].start_index, 1);
    assert_eq!(extractor.runs()[1].length, 2);
}

#[test]
fn style_atom_splits_text_into_character_runs() {
    let text_record = Record {
        record_type: RecordType::TextBytesAtom,
        record_type_raw: 4008,
        version: 0,
        instance: 0,
        data_length: 4,
        data: b"abcd".to_vec(),
        children: Vec::new(),
    };
    let mut style_data = Vec::new();
    style_data.extend_from_slice(&5u32.to_le_bytes());
    style_data.extend_from_slice(&0i16.to_le_bytes());
    style_data.extend_from_slice(&0u32.to_le_bytes());
    style_data.extend_from_slice(&2u32.to_le_bytes());
    style_data.extend_from_slice(&0x0001u32.to_le_bytes());
    style_data.extend_from_slice(&0x0001i16.to_le_bytes());
    style_data.extend_from_slice(&3u32.to_le_bytes());
    style_data.extend_from_slice(&0x0002u32.to_le_bytes());
    style_data.extend_from_slice(&0x0002i16.to_le_bytes());
    let style_record = Record {
        record_type: RecordType::StyleTextPropAtom,
        record_type_raw: 4001,
        version: 0,
        instance: 0,
        data_length: style_data.len() as u32,
        data: style_data,
        children: Vec::new(),
    };
    let mut extractor = TextRunExtractor::new();

    extractor
        .extract_from_records(&[text_record, style_record])
        .unwrap();

    assert_eq!(extractor.text(), "abcd");
    assert_eq!(extractor.run_count(), 2);
    assert_eq!(extractor.runs()[0].text, "ab");
    assert!(extractor.runs()[0].formatting.bold);
    assert!(!extractor.runs()[0].formatting.italic);
    assert_eq!(extractor.runs()[1].text, "cd");
    assert!(!extractor.runs()[1].formatting.bold);
    assert!(extractor.runs()[1].formatting.italic);
    assert_eq!(extractor.runs()[1].start_index, 2);
}

#[test]
fn style_spans_count_utf16_code_units_without_splitting_surrogates() {
    assert_eq!(utf16_prefix("😀x", 2), ("😀".len(), 1));
    assert_eq!(utf16_prefix("😀x", 3), ("😀x".len(), 2));
}

#[test]
fn exposes_complete_paragraph_runs_with_utf16_coverage() {
    let text = "😀a\rb";
    let text_record = Record {
        record_type: RecordType::TextCharsAtom,
        record_type_raw: 4000,
        version: 0,
        instance: 0,
        data_length: (text.encode_utf16().count() * 2) as u32,
        data: text.encode_utf16().flat_map(u16::to_le_bytes).collect(),
        children: Vec::new(),
    };
    let mask: u32 = 0x000F
        | 0x0010
        | 0x0020
        | 0x0040
        | 0x0080
        | 0x0100
        | 0x0400
        | 0x0800
        | 0x1000
        | 0x2000
        | 0x4000
        | 0x8000
        | 0x0001_0000
        | 0x0002_0000
        | 0x0004_0000
        | 0x0008_0000
        | 0x0010_0000
        | 0x0020_0000;
    let mut style_data = Vec::new();
    style_data.extend_from_slice(&4u32.to_le_bytes());
    style_data.extend_from_slice(&2i16.to_le_bytes());
    style_data.extend_from_slice(&mask.to_le_bytes());
    style_data.extend_from_slice(&0x000Fu16.to_le_bytes());
    style_data.extend_from_slice(&0x2022u16.to_le_bytes());
    style_data.extend_from_slice(&65_535u16.to_le_bytes());
    style_data.extend_from_slice(&(-24i16).to_le_bytes());
    style_data.extend_from_slice(&0xFE33_2211u32.to_le_bytes());
    style_data.extend_from_slice(&4u16.to_le_bytes());
    style_data.extend_from_slice(&120i16.to_le_bytes());
    style_data.extend_from_slice(&(-10i16).to_le_bytes());
    style_data.extend_from_slice(&20i16.to_le_bytes());
    style_data.extend_from_slice(&720i16.to_le_bytes());
    style_data.extend_from_slice(&(-360i16).to_le_bytes());
    style_data.extend_from_slice(&144i16.to_le_bytes());
    style_data.extend_from_slice(&2u16.to_le_bytes());
    style_data.extend_from_slice(&(-20i16).to_le_bytes());
    style_data.extend_from_slice(&1u16.to_le_bytes());
    style_data.extend_from_slice(&720i16.to_le_bytes());
    style_data.extend_from_slice(&3u16.to_le_bytes());
    style_data.extend_from_slice(&3u16.to_le_bytes());
    style_data.extend_from_slice(&5u16.to_le_bytes());
    style_data.extend_from_slice(&1u16.to_le_bytes());
    style_data.extend_from_slice(&2u32.to_le_bytes());
    style_data.extend_from_slice(&1i16.to_le_bytes());
    style_data.extend_from_slice(&0x0800u32.to_le_bytes());
    style_data.extend_from_slice(&2u16.to_le_bytes());
    style_data.extend_from_slice(&6u32.to_le_bytes());
    style_data.extend_from_slice(&0u32.to_le_bytes());
    let style_record = Record {
        record_type: RecordType::StyleTextPropAtom,
        record_type_raw: 4001,
        version: 0,
        instance: 0,
        data_length: style_data.len() as u32,
        data: style_data,
        children: Vec::new(),
    };
    let mut extractor = TextRunExtractor::new();

    extractor
        .extract_from_records(&[text_record, style_record])
        .unwrap();

    assert_eq!(extractor.paragraph_runs().len(), 2);
    let first = &extractor.paragraph_runs()[0];
    assert_eq!(first.text, "😀a\r");
    assert_eq!(first.start_index, 0);
    assert_eq!(first.length, 3);
    assert_eq!(first.formatting.property_mask, mask);
    assert_eq!(first.formatting.indent_level, 2);
    assert_eq!(first.formatting.bullet_flags_raw, Some(0x000F));
    assert_eq!(first.formatting.bullet_enabled, Some(true));
    assert_eq!(first.formatting.bullet_font_enabled, Some(true));
    assert_eq!(first.formatting.bullet_color_enabled, Some(true));
    assert_eq!(first.formatting.bullet_size_enabled, Some(true));
    assert_eq!(first.formatting.bullet_character, Some(0x2022));
    assert_eq!(first.formatting.bullet_font_index, Some(65_535));
    assert_eq!(first.formatting.bullet_size, Some(-24));
    assert_eq!(first.formatting.bullet_color, Some(0x0011_2233));
    assert_eq!(first.formatting.bullet_color_raw, Some(0xFE33_2211));
    assert_eq!(
        first.formatting.alignment,
        Some(ParagraphAlignment::Distributed)
    );
    assert_eq!(first.formatting.line_spacing, Some(120));
    assert_eq!(first.formatting.space_before, Some(-10));
    assert_eq!(first.formatting.space_after, Some(20));
    assert_eq!(first.formatting.left_margin, Some(720));
    assert_eq!(first.formatting.indent, Some(-360));
    assert_eq!(first.formatting.default_tab_size, Some(144));
    assert_eq!(
        first.formatting.tab_stops,
        Some(vec![
            ParagraphTabStop {
                position: -20,
                alignment: ParagraphTabAlignment::Center,
            },
            ParagraphTabStop {
                position: 720,
                alignment: ParagraphTabAlignment::Decimal,
            },
        ])
    );
    assert_eq!(
        first.formatting.font_alignment,
        Some(ParagraphFontAlignment::UpholdFixed)
    );
    assert_eq!(first.formatting.wrap_flags_raw, Some(5));
    assert_eq!(first.formatting.character_wrap, Some(true));
    assert_eq!(first.formatting.word_wrap, Some(false));
    assert_eq!(first.formatting.overflow, Some(true));
    assert_eq!(
        first.formatting.text_direction,
        Some(ParagraphTextDirection::RightToLeft)
    );

    let second = &extractor.paragraph_runs()[1];
    assert_eq!(second.text, "b");
    assert_eq!(second.start_index, 3);
    assert_eq!(second.length, 1);
    assert_eq!(second.formatting.indent_level, 1);
    assert_eq!(second.formatting.alignment, Some(ParagraphAlignment::Right));
}

#[test]
fn rejects_invalid_paragraph_enumerations() {
    let mut style = super::super::text_prop::TextPropCollection::new(
        1,
        super::super::text_prop::TextPropType::Paragraph,
    );
    style.property_mask = 0x0010_0000;
    style.tab_stops.push(super::super::text_prop::TextTabStop {
        position: 0,
        alignment: 4,
    });

    let error = paragraph_formatting_from_style(&style).unwrap_err();
    assert!(error.to_string().contains("TextTabTypeEnum"));

    style.indent_level = 5;
    let error = paragraph_formatting_from_style(&style).unwrap_err();
    assert!(error.to_string().contains("indent level"));
}

#[test]
fn preserves_formatting_for_an_empty_paragraph() {
    let text_record = Record {
        record_type: RecordType::TextBytesAtom,
        record_type_raw: 4008,
        version: 0,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: Vec::new(),
    };
    let mut style_data = Vec::new();
    style_data.extend_from_slice(&1u32.to_le_bytes());
    style_data.extend_from_slice(&0i16.to_le_bytes());
    style_data.extend_from_slice(&0x0800u32.to_le_bytes());
    style_data.extend_from_slice(&1u16.to_le_bytes());
    style_data.extend_from_slice(&1u32.to_le_bytes());
    style_data.extend_from_slice(&0u32.to_le_bytes());
    let style_record = Record {
        record_type: RecordType::StyleTextPropAtom,
        record_type_raw: 4001,
        version: 0,
        instance: 0,
        data_length: style_data.len() as u32,
        data: style_data,
        children: Vec::new(),
    };
    let mut extractor = TextRunExtractor::new();

    extractor
        .extract_from_records(&[text_record, style_record])
        .unwrap();

    assert!(extractor.runs().is_empty());
    assert_eq!(extractor.paragraph_runs().len(), 1);
    assert_eq!(extractor.paragraph_runs()[0].text, "");
    assert_eq!(
        extractor.paragraph_runs()[0].formatting.alignment,
        Some(ParagraphAlignment::Center)
    );
}

#[test]
fn decodes_direct_and_scheme_color_index_structs() {
    assert_eq!(
        decode_color_index_struct(0xFE33_2211),
        (Some(0x0011_2233), None)
    );
    assert_eq!(decode_color_index_struct(0x0400_0000), (None, Some(4)));
    assert_eq!(decode_color_index_struct(0xFF00_0000), (None, None));
}

#[test]
fn rejects_invalid_text_cf_font_sizes() {
    let text_record = Record {
        record_type: RecordType::TextBytesAtom,
        record_type_raw: 4008,
        version: 0,
        instance: 0,
        data_length: 1,
        data: b"x".to_vec(),
        children: Vec::new(),
    };
    let mut style_data = Vec::new();
    style_data.extend_from_slice(&2u32.to_le_bytes());
    style_data.extend_from_slice(&0i16.to_le_bytes());
    style_data.extend_from_slice(&0u32.to_le_bytes());
    style_data.extend_from_slice(&2u32.to_le_bytes());
    style_data.extend_from_slice(&0x0002_0000u32.to_le_bytes());
    style_data.extend_from_slice(&0i16.to_le_bytes());
    let style_record = Record {
        record_type: RecordType::StyleTextPropAtom,
        record_type_raw: 4001,
        version: 0,
        instance: 0,
        data_length: style_data.len() as u32,
        data: style_data,
        children: Vec::new(),
    };
    let mut extractor = TextRunExtractor::new();

    let error = extractor
        .extract_from_records(&[text_record, style_record])
        .unwrap_err();
    assert!(error.to_string().contains("font size"));
}

#[test]
fn rejects_invalid_text_cf_color_and_baseline_values() {
    let mut invalid_color = super::super::text_prop::TextPropCollection::new(
        1,
        super::super::text_prop::TextPropType::Character,
    );
    let mut color = super::super::text_prop::TextProp::new("font.color", 4, 0x40000);
    color.value = 0x0800_0000;
    invalid_color.properties.push(color);
    assert!(formatting_from_style(&invalid_color).is_err());

    let mut invalid_position = super::super::text_prop::TextPropCollection::new(
        1,
        super::super::text_prop::TextPropType::Character,
    );
    let mut position = super::super::text_prop::TextProp::new("superscript", 2, 0x80000);
    position.value = 101;
    invalid_position.properties.push(position);
    assert!(formatting_from_style(&invalid_position).is_err());
}
