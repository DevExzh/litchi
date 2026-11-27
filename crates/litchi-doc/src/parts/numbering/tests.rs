//! Focused regression tests for the DOC numbering owner.

use super::*;

#[test]
fn test_number_format() {
    assert_eq!(NumberFormat::try_from(0), Ok(NumberFormat::Arabic));
    assert_eq!(NumberFormat::try_from(23), Ok(NumberFormat::Bullet));
    assert_eq!(NumberFormat::try_from(255), Ok(NumberFormat::None));
}

#[test]
fn test_list_alignment() {
    assert_eq!(ListAlignment::try_from(0), Ok(ListAlignment::Left));
    assert_eq!(ListAlignment::try_from(1), Ok(ListAlignment::Center));
    assert_eq!(ListAlignment::try_from(2), Ok(ListAlignment::Right));
    assert_eq!(ListAlignment::try_from(3), Err(3));
}

#[test]
fn test_number_format_all_variants() {
    for value in 0..=59 {
        assert_eq!(NumberFormat::try_from(value).unwrap() as u8, value);
    }
    assert_eq!(NumberFormat::try_from(255), Ok(NumberFormat::None));
}

#[test]
fn test_number_format_rejects_unknown() {
    for value in 60..=254 {
        assert_eq!(NumberFormat::try_from(value), Err(value));
    }
}

#[test]
fn test_number_format_clone() {
    let fmt = NumberFormat::Bullet;
    let cloned = fmt;
    assert_eq!(fmt, cloned);
}

#[test]
fn test_number_format_copy() {
    let fmt = NumberFormat::UpperRoman;
    let copied = fmt;
    assert_eq!(fmt, copied);
}

#[test]
fn test_number_format_debug() {
    let fmt = NumberFormat::Arabic;
    let debug_str = format!("{:?}", fmt);
    assert!(debug_str.contains("Arabic"));
}

#[test]
fn test_number_format_equality() {
    assert_eq!(NumberFormat::Arabic, NumberFormat::Arabic);
    assert_ne!(NumberFormat::Arabic, NumberFormat::Bullet);
}

#[test]
fn test_list_alignment_rejects_unknown() {
    assert_eq!(ListAlignment::try_from(3), Err(3));
    assert_eq!(ListAlignment::try_from(100), Err(100));
}

#[test]
fn test_list_alignment_clone() {
    let align = ListAlignment::Center;
    let cloned = align;
    assert_eq!(align, cloned);
}

#[test]
fn test_list_level_creation() {
    let level = ListLevel {
        start_at: 1,
        number_format: NumberFormat::Arabic,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 720,
        indent_hanging: 360,
        number_text: "%1.".to_string(),
    };

    assert_eq!(level.start_at, 1);
    assert_eq!(level.number_format, NumberFormat::Arabic);
    assert_eq!(level.alignment, ListAlignment::Left);
    assert_eq!(level.level, 0);
    assert_eq!(level.indent_left, 720);
    assert_eq!(level.indent_hanging, 360);
    assert_eq!(level.number_text, "%1.");
    assert!(level.is_numbered());
    assert!(!level.is_bullet());
}

#[test]
fn test_list_level_bullet() {
    let level = ListLevel {
        start_at: 1,
        number_format: NumberFormat::Bullet,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 720,
        indent_hanging: 360,
        number_text: "\u{2022}".to_string(),
    };

    assert!(level.is_bullet());
    assert!(!level.is_numbered());
}

#[test]
fn test_list_level_none() {
    let level = ListLevel {
        start_at: 0,
        number_format: NumberFormat::None,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 0,
        indent_hanging: 0,
        number_text: String::new(),
    };

    assert!(!level.is_bullet());
    assert!(!level.is_numbered());
}

#[test]
fn test_list_level_clone() {
    let level = ListLevel {
        start_at: 1,
        number_format: NumberFormat::LowerRoman,
        alignment: ListAlignment::Right,
        level: 2,
        follow_char: 1,
        indent_left: 1440,
        indent_hanging: 720,
        number_text: "(%2)".to_string(),
    };
    let cloned = level.clone();

    assert_eq!(cloned.start_at, level.start_at);
    assert_eq!(cloned.number_format, level.number_format);
    assert_eq!(cloned.alignment, level.alignment);
    assert_eq!(cloned.level, level.level);
    assert_eq!(cloned.number_text, level.number_text);
}

#[test]
fn test_list_level_debug() {
    let level = ListLevel {
        start_at: 1,
        number_format: NumberFormat::Arabic,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 720,
        indent_hanging: 360,
        number_text: "%1.".to_string(),
    };
    let debug_str = format!("{:?}", level);
    assert!(debug_str.contains("ListLevel"));
    assert!(debug_str.contains("Arabic"));
}

#[test]
fn test_list_structure_creation() {
    let levels = vec![ListLevel {
        start_at: 1,
        number_format: NumberFormat::Arabic,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 720,
        indent_hanging: 360,
        number_text: "%1.".to_string(),
    }];

    let lst = ListStructure {
        list_id: 12345,
        template_id: 67890,
        is_simple: false,
        levels,
    };

    assert_eq!(lst.list_id, 12345);
    assert_eq!(lst.template_id, 67890);
    assert!(!lst.is_simple);
    assert_eq!(lst.levels.len(), 1);
}

#[test]
fn test_list_structure_simple() {
    let lst = ListStructure {
        list_id: 1,
        template_id: 1,
        is_simple: true,
        levels: Vec::new(),
    };

    assert!(lst.is_simple);
}

#[test]
fn test_list_structure_level_accessor() {
    let levels = vec![
        ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        },
        ListLevel {
            start_at: 1,
            number_format: NumberFormat::LowerLetter,
            alignment: ListAlignment::Left,
            level: 1,
            follow_char: 0,
            indent_left: 1440,
            indent_hanging: 360,
            number_text: "%1.%2.".to_string(),
        },
    ];

    let lst = ListStructure {
        list_id: 1,
        template_id: 1,
        is_simple: false,
        levels,
    };

    assert!(lst.level(0).is_some());
    assert!(lst.level(1).is_some());
    assert!(lst.level(2).is_none());
    assert_eq!(lst.level(0).unwrap().number_format, NumberFormat::Arabic);
    assert_eq!(
        lst.level(1).unwrap().number_format,
        NumberFormat::LowerLetter
    );
}

#[test]
fn test_list_structure_clone() {
    let lst = ListStructure {
        list_id: 100,
        template_id: 200,
        is_simple: false,
        levels: vec![ListLevel {
            start_at: 1,
            number_format: NumberFormat::Bullet,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "\u{2022}".to_string(),
        }],
    };
    let cloned = lst.clone();

    assert_eq!(cloned.list_id, lst.list_id);
    assert_eq!(cloned.template_id, lst.template_id);
    assert_eq!(cloned.levels.len(), lst.levels.len());
}

#[test]
fn test_list_structure_debug() {
    let lst = ListStructure {
        list_id: 1,
        template_id: 2,
        is_simple: false,
        levels: Vec::new(),
    };
    let debug_str = format!("{:?}", lst);
    assert!(debug_str.contains("ListStructure"));
}

#[test]
fn test_list_format_override_creation() {
    let lfo = ListFormatOverride {
        list_id: 12345,
        override_count: 1,
        lfo_id: 1,
        level_overrides: Vec::new(),
    };

    assert_eq!(lfo.list_id, 12345);
    assert_eq!(lfo.override_count, 1);
    assert_eq!(lfo.lfo_id, 1);
}

#[test]
fn test_list_format_override_clone() {
    let lfo = ListFormatOverride {
        list_id: 100,
        override_count: 2,
        lfo_id: 5,
        level_overrides: Vec::new(),
    };
    let cloned = lfo.clone();

    assert_eq!(cloned.list_id, lfo.list_id);
    assert_eq!(cloned.override_count, lfo.override_count);
    assert_eq!(cloned.lfo_id, lfo.lfo_id);
}

#[test]
fn test_list_format_override_debug() {
    let lfo = ListFormatOverride {
        list_id: 1,
        override_count: 0,
        lfo_id: 1,
        level_overrides: Vec::new(),
    };
    let debug_str = format!("{:?}", lfo);
    assert!(debug_str.contains("ListFormatOverride"));
}

#[test]
fn test_list_tables_empty() {
    let tables = ListTables {
        list_structures: Vec::new(),
        list_overrides: Vec::new(),
        metadata: ListTablesMetadata::default(),
    };

    assert!(tables.structures().is_empty());
    assert!(tables.overrides().is_empty());
}

#[test]
fn test_list_tables_with_data() {
    let tables = ListTables {
        list_structures: vec![ListStructure {
            list_id: 1,
            template_id: 1,
            is_simple: false,
            levels: Vec::new(),
        }],
        list_overrides: vec![ListFormatOverride {
            list_id: 1,
            override_count: 0,
            lfo_id: 1,
            level_overrides: Vec::new(),
        }],
        metadata: ListTablesMetadata::default(),
    };

    assert_eq!(tables.structures().len(), 1);
    assert_eq!(tables.overrides().len(), 1);
}

#[test]
fn test_list_tables_find_structure() {
    let tables = ListTables {
        list_structures: vec![
            ListStructure {
                list_id: 100,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            },
            ListStructure {
                list_id: 200,
                template_id: 2,
                is_simple: true,
                levels: Vec::new(),
            },
        ],
        list_overrides: Vec::new(),
        metadata: ListTablesMetadata::default(),
    };

    assert!(tables.find_structure(100).is_some());
    assert!(tables.find_structure(200).is_some());
    assert!(tables.find_structure(999).is_none());
}

#[test]
fn test_list_tables_find_override() {
    let tables = ListTables {
        list_structures: Vec::new(),
        list_overrides: vec![
            ListFormatOverride {
                list_id: 1,
                override_count: 0,
                lfo_id: 10,
                level_overrides: Vec::new(),
            },
            ListFormatOverride {
                list_id: 2,
                override_count: 1,
                lfo_id: 20,
                level_overrides: Vec::new(),
            },
        ],
        metadata: ListTablesMetadata::default(),
    };

    assert!(tables.find_override(10).is_some());
    assert!(tables.find_override(20).is_some());
    assert!(tables.find_override(999).is_none());
}

#[test]
fn test_list_tables_get_list_for_lfo() {
    let tables = ListTables {
        list_structures: vec![ListStructure {
            list_id: 100,
            template_id: 1,
            is_simple: false,
            levels: Vec::new(),
        }],
        list_overrides: vec![ListFormatOverride {
            list_id: 100,
            override_count: 0,
            lfo_id: 1,
            level_overrides: Vec::new(),
        }],
        metadata: ListTablesMetadata::default(),
    };

    let lst = tables.get_list_for_lfo(1);
    assert!(lst.is_some());
    assert_eq!(lst.unwrap().list_id, 100);

    assert!(tables.get_list_for_lfo(999).is_none());
}

#[test]
fn test_list_tables_get_list_for_lfo_no_override() {
    let tables = ListTables {
        list_structures: vec![ListStructure {
            list_id: 100,
            template_id: 1,
            is_simple: false,
            levels: Vec::new(),
        }],
        list_overrides: Vec::new(),
        metadata: ListTablesMetadata::default(),
    };

    assert!(tables.get_list_for_lfo(1).is_none());
}

#[test]
fn paragraph_binding_borrows_base_level_and_applies_start_override() {
    let base = ListLevel {
        start_at: 1,
        number_format: NumberFormat::Arabic,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 0,
        indent_hanging: 0,
        number_text: "%1.".to_string(),
    };
    let tables = ListTables {
        list_structures: vec![ListStructure {
            list_id: 42,
            template_id: 42,
            is_simple: true,
            levels: vec![base],
        }],
        list_overrides: vec![ListFormatOverride {
            list_id: 42,
            override_count: 1,
            lfo_id: 1,
            level_overrides: vec![ListLevelOverride {
                level: 0,
                start_at: Some(7),
                format: None,
            }],
        }],
        metadata: ListTablesMetadata::default(),
    };

    let binding = tables.bind_paragraph(-1, 0).unwrap();
    assert!(binding.preserve_indents);
    assert_eq!(binding.definition.list_id, 42);
    assert_eq!(binding.format_override.lfo_id, 1);
    assert!(std::ptr::eq(binding.effective_level(), binding.base_level));
    assert_eq!(binding.effective_start_at(), 7);
    assert!(binding.has_start_at_override());
    assert!(!binding.has_formatting_override());
}

#[test]
fn paragraph_binding_borrows_formatting_override_and_rejects_sentinels() {
    let level = |start_at, text: &str| ListLevel {
        start_at,
        number_format: NumberFormat::Arabic,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 0,
        indent_hanging: 0,
        number_text: text.to_string(),
    };
    let tables = ListTables {
        list_structures: vec![ListStructure {
            list_id: 9,
            template_id: 9,
            is_simple: true,
            levels: vec![level(1, "%1.")],
        }],
        list_overrides: vec![ListFormatOverride {
            list_id: 9,
            override_count: 1,
            lfo_id: 1,
            level_overrides: vec![ListLevelOverride {
                level: 0,
                start_at: None,
                format: Some(level(3, "(%1)")),
            }],
        }],
        metadata: ListTablesMetadata::default(),
    };

    let binding = tables.bind_paragraph(1, 0).unwrap();
    let replacement = binding.level_override.unwrap().format.as_ref().unwrap();
    assert!(std::ptr::eq(binding.effective_level(), replacement));
    assert_eq!(binding.effective_start_at(), 3);
    assert!(binding.has_formatting_override());
    assert!(tables.bind_paragraph(0, 0).is_none());
    assert!(tables.bind_paragraph(i16::MIN, 0).is_none());
    assert!(tables.bind_paragraph(1, 12).is_none());
}

#[test]
fn test_list_level_from_bytes_too_short() {
    let data = vec![0u8; 10];
    let result = ListLevel::from_bytes(&data, 0);
    assert!(result.is_err());
}

#[test]
fn test_list_level_from_bytes_minimal() {
    // A minimal LVL is a 28-byte LVLF plus an empty two-byte XST.
    let mut data = vec![0u8; 30];
    // start_at at offset 0
    data[0] = 1; // start_at = 1
    // number_format at offset 4
    data[4] = 0; // Arabic
    // alignment at offset 5
    data[5] = 0; // Left
    // follow_char at offset 15
    data[15] = 0;
    // dxaIndentSav at offset 16
    data[16] = 0xD0; // 720 in little-endian
    data[17] = 0x02;

    let result = ListLevel::from_bytes(&data, 0);
    assert!(result.is_ok());

    let level = result.unwrap();
    assert_eq!(level.start_at, 1);
    assert_eq!(level.number_format, NumberFormat::Arabic);
    assert_eq!(level.alignment, ListAlignment::Left);
    assert_eq!(level.level, 0);
    assert_eq!(level.indent_left, 720);
    assert_eq!(level.indent_hanging, 0);
    assert_eq!(level.number_text, "");
}

#[test]
fn list_level_preserves_exotic_msonfc_and_rejects_reserved_values() {
    let mut data = vec![0u8; 30];
    data[4] = NumberFormat::RussianUpper as u8;
    let level = ListLevel::from_bytes(&data, 0).unwrap();
    assert_eq!(level.number_format, NumberFormat::RussianUpper);

    data[4] = 0x3C;
    assert!(ListLevel::from_bytes(&data, 0).is_err());
    data[4] = NumberFormat::Hex as u8;
    assert!(ListLevel::from_bytes(&data, 0).is_err());
    data[4] = NumberFormat::Arabic as u8;
    data[5] = 3;
    assert!(ListLevel::from_bytes(&data, 0).is_err());
    data[5] = 0;
    data[15] = 3;
    assert!(ListLevel::from_bytes(&data, 0).is_err());
    data[15] = 0;
    data[..4].copy_from_slice(&32_768u32.to_le_bytes());
    assert!(ListLevel::from_bytes(&data, 0).is_err());
}

#[test]
fn test_list_level_from_bytes_bullet() {
    let mut data = vec![0u8; 32];
    data[4] = 23; // Bullet format
    data[28..30].copy_from_slice(&1u16.to_le_bytes());
    data[30..32].copy_from_slice(&0x2022u16.to_le_bytes());

    let level = ListLevel::from_bytes(&data, 0).unwrap();
    assert!(level.is_bullet());
    assert!(!level.is_numbered());
}

#[test]
fn test_list_level_from_bytes_with_text() {
    let mut data = vec![0u8; 34];
    // Fixed part
    data[0] = 1; // start_at
    data[4] = 0; // Arabic
    data[5] = 0; // Left
    data[15] = 0; // follow_char
    data[28..30].copy_from_slice(&2u16.to_le_bytes());
    data[30..32].copy_from_slice(&0u16.to_le_bytes()); // level 0 placeholder
    data[32..34].copy_from_slice(&('.' as u16).to_le_bytes());

    let level = ListLevel::from_bytes(&data, 0).unwrap();
    assert_eq!(level.number_text, "%1.");
}

#[test]
fn test_list_structure_from_bytes_too_short() {
    let data = vec![0u8; 10];
    let result = ListStructure::from_bytes(&data);
    assert!(result.is_err());
}

#[test]
fn test_list_structure_from_bytes_minimal() {
    let mut data = vec![0u8; 28];
    // list_id at offset 0
    data[0] = 0x39; // 57 in little-endian
    data[1] = 0x00;
    data[2] = 0x00;
    data[3] = 0x00;
    // template_id at offset 4
    data[4] = 0x30; // 48 in little-endian
    data[5] = 0x00;
    // flags at offset 26 - simple flag
    data[26] = 0x01; // is_simple = true

    let result = ListStructure::from_bytes(&data);
    assert!(result.is_ok());

    let lst = result.unwrap();
    assert_eq!(lst.list_id, 57);
    assert_eq!(lst.template_id, 48);
    assert!(lst.is_simple);
    assert!(lst.levels.is_empty());
}

#[test]
fn test_list_format_override_from_bytes_too_short() {
    let data = vec![0u8; 5];
    let result = ListFormatOverride::from_bytes(&data);
    assert!(result.is_err());
}

#[test]
fn test_list_format_override_from_bytes_valid() {
    let mut data = vec![0u8; 16];
    // list_id at offset 0
    data[0] = 0x39;
    data[1] = 0x00;
    data[2] = 0x00;
    data[3] = 0x00;
    // override_count at offset 12
    data[12] = 2;

    let result = ListFormatOverride::from_bytes(&data);
    assert!(result.is_ok());

    let lfo = result.unwrap();
    assert_eq!(lfo.list_id, 57);
    assert_eq!(lfo.override_count, 2);
}

#[test]
fn test_numbering_with_unicode_number_text() {
    let level = ListLevel {
        start_at: 1,
        number_format: NumberFormat::Bullet,
        alignment: ListAlignment::Left,
        level: 0,
        follow_char: 0,
        indent_left: 720,
        indent_hanging: 360,
        number_text: "\u{2022} \u{25ba} \u{2192}".to_string(), // bullet, pointer, arrow
    };

    assert_eq!(level.number_text, "\u{2022} \u{25ba} \u{2192}");
}

#[test]
fn test_list_level_negative_indent() {
    let mut data = vec![0u8; 30];
    // dxaIndentSav at offset 16 (signed 32-bit)
    data[16] = 0xF0; // -16 in little-endian two's complement
    data[17] = 0xFF;
    data[18] = 0xFF;
    data[19] = 0xFF;

    let level = ListLevel::from_bytes(&data, 0).unwrap();
    assert_eq!(level.indent_left, -16);
}

#[test]
fn parses_split_plflst_header_and_level_array() {
    let mut writer = crate::writer::numbering::NumberingWriter::new();
    let mut list = crate::writer::numbering::ListStructure::new(42);
    let mut first = crate::writer::numbering::ListLevel::new(3, NumberFormat::Decimal);
    first.number_text = "%1.😀".to_string();
    list.add_level(first);
    list.add_level(crate::writer::numbering::ListLevel::new(
        1,
        NumberFormat::LowerLetter,
    ));
    writer.add_list(list);
    let (header, levels) = writer.build_plflst().unwrap();

    let parsed = ListTables::parse_plflst(&header, &levels).unwrap();
    assert_eq!(parsed.len(), 1);
    assert_eq!(parsed[0].list_id, 42);
    assert!(!parsed[0].is_simple);
    assert_eq!(parsed[0].levels.len(), 9);
    assert_eq!(parsed[0].levels[0].start_at, 3);
    assert_eq!(parsed[0].levels[0].number_text, "%1.😀");
    assert_eq!(parsed[0].levels[1].number_format, NumberFormat::LowerLetter);
}

#[test]
fn parses_parallel_lfo_and_lfo_data_arrays() {
    let mut writer = crate::writer::numbering::NumberingWriter::new();
    writer.add_override(crate::writer::numbering::ListFormatOverride::new(100, 1));
    writer.add_override(crate::writer::numbering::ListFormatOverride::new(200, 2));

    let parsed = ListTables::parse_plflfo(&writer.build_plflfo()).unwrap();
    assert_eq!(parsed.len(), 2);
    assert_eq!((parsed[0].list_id, parsed[0].lfo_id), (100, 1));
    assert_eq!((parsed[1].list_id, parsed[1].lfo_id), (200, 2));
}

#[test]
fn rejects_truncated_list_tables() {
    assert!(ListTables::parse_plflst(&[1, 0], &[]).is_err());
    assert!(ListTables::parse_plflst(&[0, 0, 0], &[]).is_err());

    let mut truncated_lfo = vec![0u8; 20];
    truncated_lfo[..4].copy_from_slice(&1u32.to_le_bytes());
    assert!(ListTables::parse_plflfo(&truncated_lfo).is_err());
}
