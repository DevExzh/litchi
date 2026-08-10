#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::float_cmp,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

use super::CellsReader;
use crate::conditional_formatting::{
    AxisPosition14, Bar14, Color, Formatting, IconSet14, RecordKind, Rule, RuleMetadata, RuleType,
    Value,
};
use crate::package::cell::CellHeader;
use crate::package::error::Error;
use crate::package::formula::Context;
use crate::package::records::Stream;
use crate::raw::{Writer, kind};
use litchi_core::sheet::{Cell, CellValue};

fn rich_string_worksheet() -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
    writer.write_record(kind::BEGIN_COL_INFOS, &[]).unwrap();
    let mut column = 2u32.to_le_bytes().to_vec();
    column.extend_from_slice(&4u32.to_le_bytes());
    column.extend_from_slice(&4096u32.to_le_bytes());
    column.extend_from_slice(&0u32.to_le_bytes());
    column.extend_from_slice(&0x130Fu16.to_le_bytes());
    writer.write_record(kind::COL_INFO, &column).unwrap();
    writer.write_record(kind::END_COL_INFOS, &[]).unwrap();
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();

    let mut row = 4u32.to_le_bytes().to_vec();
    row.extend_from_slice(&0u32.to_le_bytes());
    row.extend_from_slice(&500u16.to_le_bytes());
    row.extend_from_slice(&[3, 0x7A, 1]);
    row.extend_from_slice(&1u32.to_le_bytes());
    row.extend_from_slice(&2u32.to_le_bytes());
    row.extend_from_slice(&2u32.to_le_bytes());
    writer.write_record(kind::ROW_HDR, &row).unwrap();

    let mut cell = 2u32.to_le_bytes().to_vec();
    cell.extend_from_slice(&[0, 0, 0, 1]);
    cell.push(1);
    cell.extend_from_slice(&2u32.to_le_bytes());
    cell.extend_from_slice(&[b'A', 0, b'B', 0]);
    cell.extend_from_slice(&2u32.to_le_bytes());
    cell.extend_from_slice(&[0, 0, 3, 0, 1, 0, 5, 0]);
    writer.write_record(kind::CELL_R_STRING, &cell).unwrap();
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();
    bytes
}

#[test]
fn decodes_and_validates_cell_style_header() {
    type Reader<'a> = CellsReader<'a, std::io::Cursor<&'a [u8]>>;

    let header = [7, 0, 0, 0, 0x34, 0x12, 0, 1];
    assert_eq!(
        Reader::decode_cell_header(&header, 0x1235).unwrap(),
        CellHeader {
            col: 7,
            style_id: 0x1234,
            show_phonetic: true,
        }
    );
    assert!(matches!(
        Reader::decode_cell_header(&header, 0x1234),
        Err(Error::Unrecognized { .. })
    ));

    let reserved = [0, 0, 0, 0, 0, 0, 0, 2];
    assert!(matches!(
        Reader::decode_cell_header(&reserved, 1),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn parses_iso_sheet_protection_metadata_and_matching_flags() {
    type Reader<'a> = CellsReader<'a, std::io::Cursor<&'a [u8]>>;
    let flags = [
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true, false,
    ];
    let mut iso = 250_000u32.to_le_bytes().to_vec();
    for flag in flags {
        iso.extend_from_slice(&u32::from(flag).to_le_bytes());
    }
    iso.extend_from_slice(&3u32.to_le_bytes());
    iso.extend_from_slice(&[1, 2, 3]);
    iso.extend_from_slice(&2u32.to_le_bytes());
    iso.extend_from_slice(&[4, 5]);
    iso.extend_from_slice(&7u32.to_le_bytes());
    for unit in "SHA-512".encode_utf16() {
        iso.extend_from_slice(&unit.to_le_bytes());
    }

    let (strong, parsed_flags) = Reader::parse_strong_sheet_protection(&iso).unwrap();
    assert_eq!(strong.spin_count, 250_000);
    assert_eq!(strong.hash, vec![1, 2, 3]);
    assert_eq!(strong.salt, vec![4, 5]);
    assert_eq!(strong.algorithm, "SHA-512");
    assert_eq!(parsed_flags, flags);

    let mut base = 0u16.to_le_bytes().to_vec();
    for flag in flags {
        base.extend_from_slice(&u32::from(flag).to_le_bytes());
    }
    let protection = Reader::parse_sheet_protection(&base).unwrap();
    assert_eq!(Reader::sheet_protection_flags(&protection), flags);

    let mut worksheet = Vec::new();
    let mut writer = Writer::new(&mut worksheet);
    writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    writer
        .write_record(kind::SHEET_PROTECTION_ISO, &iso)
        .unwrap();
    writer.write_record(kind::SHEET_PROTECTION, &base).unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();
    let formula_context = Context::default();
    let iter = Stream::new(std::io::Cursor::new(worksheet));
    let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
    assert!(reader.next_cell().unwrap().is_none());
    assert_eq!(reader.sheet_protection.unwrap(), protection);
    assert_eq!(reader.strong_sheet_protection.unwrap(), strong);
}

#[test]
fn reads_inline_rich_string_cells() {
    let bytes = rich_string_worksheet();
    let formula_context = Context::default();
    let iter = Stream::new(std::io::Cursor::new(bytes));
    let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();

    let cell = reader.next_cell().unwrap().unwrap();
    assert_eq!(cell.row(), 4);
    assert_eq!(cell.column(), 2);
    assert_eq!(cell.value(), &CellValue::String("AB".to_string()));
    assert!(cell.show_phonetic());
    let rich = cell.rich_string().unwrap();
    assert_eq!(rich.text, "AB");
    assert_eq!(rich.runs.len(), 2);
    assert_eq!(rich.runs[0].font_id, 3);
    assert_eq!(rich.runs[1].font_id, 5);
    assert!(reader.next_cell().unwrap().is_none());

    assert_eq!(reader.column_infos.len(), 1);
    let column = &reader.column_infos[0];
    assert_eq!((column.first_column, column.last_column), (2, 4));
    assert_eq!(column.width, 16.0);
    assert!(column.hidden);
    assert!(column.user_set_width);
    assert!(column.best_fit);
    assert!(column.show_phonetic);
    assert_eq!(column.outline_level, 3);
    assert!(column.collapsed);

    assert_eq!(reader.row_infos.len(), 1);
    let row = &reader.row_infos[0];
    assert_eq!(row.row, 4);
    assert_eq!(row.style_id, Some(0));
    assert_eq!(row.height, Some(25.0));
    assert!(row.extra_ascender);
    assert!(row.extra_descender);
    assert_eq!(row.outline_level, 2);
    assert!(row.collapsed);
    assert!(row.hidden);
    assert!(row.show_phonetic);
    assert_eq!(row.column_spans, vec![(2, 2)]);
}

#[test]
fn rejects_a_validation_collection_with_a_mismatched_count() {
    let mut worksheet = Vec::new();
    let mut writer = Writer::new(&mut worksheet);
    writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    let mut begin = vec![0; 14];
    begin.extend_from_slice(&1u32.to_le_bytes());
    writer.write_record(kind::BEGIN_D_VALS, &begin).unwrap();
    writer.write_record(kind::END_D_VALS, &[]).unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();

    let formula_context = Context::default();
    let iter = Stream::new(std::io::Cursor::new(worksheet));
    let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
    assert!(matches!(
        reader.next_cell(),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn rejects_conditional_formatting_with_a_mismatched_rule_count() {
    let mut worksheet = Vec::new();
    let mut writer = Writer::new(&mut worksheet);
    writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    let mut begin = 1u32.to_le_bytes().to_vec();
    begin.extend_from_slice(&0u32.to_le_bytes());
    begin.extend_from_slice(&1u32.to_le_bytes());
    begin.extend_from_slice(&[0; 16]);
    writer
        .write_record(kind::BEGIN_COND_FORMATTING, &begin)
        .unwrap();
    writer.write_record(kind::END_COND_FORMATTING, &[]).unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();

    let formula_context = Context::default();
    let iter = Stream::new(std::io::Cursor::new(worksheet));
    let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
    assert!(matches!(
        reader.next_cell(),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn reads_office_2013_conditional_formatting_visualizations() {
    fn metadata(priority: i32) -> RuleMetadata {
        RuleMetadata {
            priority,
            unused: priority as u32,
            guid: [priority as u8; 16],
            guid_present: true,
            linked_classic_priority: None,
        }
    }

    let mut color_rule = Rule::new(RuleType::ColorScale, 1);
    color_rule.extension14 = Some(metadata(1));
    let color_thresholds = [Value::new(2, None), Value::new(3, None)];
    let color_records = [Color::from_argb(0xffff_0000), Color::from_argb(0xff00_ff00)];

    let mut bar_rule = Rule::new(RuleType::DataBar, 2);
    bar_rule.extension14 = Some(metadata(2));
    let bar = Bar14::new(
        Value::new(8, None),
        Value::new(9, None),
        Color::from_argb(0xff44_72c4),
    );

    let mut icon_rule = Rule::new(RuleType::IconSet, 3);
    icon_rule.extension14 = Some(metadata(3));
    let mut icon_thresholds = vec![
        Value::new(1, Some("0".to_string())),
        Value::new(1, Some("33".to_string())),
        Value::new(1, Some("67".to_string())),
    ];
    for threshold in &mut icon_thresholds {
        threshold.save_greater_than_or_equal = true;
    }
    let icon_set = IconSet14::new(18, icon_thresholds.clone());

    let mut formatting = Formatting::new(vec!["A1:A10".to_string()]);
    formatting.record_kind = RecordKind::Extension14;
    formatting.rules = vec![color_rule.clone(), bar_rule.clone(), icon_rule.clone()];

    let mut worksheet = Vec::new();
    let mut writer = Writer::new(&mut worksheet);
    writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    writer
        .write_record(
            kind::BEGIN_COND_FORMATTING14,
            &formatting.serialize_extension14_header().unwrap(),
        )
        .unwrap();

    writer
        .write_record(
            kind::BEGIN_CF_RULE14,
            &color_rule.serialize_extension14().unwrap(),
        )
        .unwrap();
    writer.write_record(kind::BEGIN_COLOR_SCALE14, &[]).unwrap();
    for threshold in &color_thresholds {
        writer
            .write_record(kind::CFVO14, &threshold.serialize_extension14().unwrap())
            .unwrap();
    }
    for color in color_records {
        writer
            .write_record(kind::COLOR14, &color.serialize_extension14().unwrap())
            .unwrap();
    }
    writer.write_record(kind::END_COLOR_SCALE14, &[]).unwrap();
    writer.write_record(kind::END_CF_RULE14, &[]).unwrap();

    writer
        .write_record(
            kind::BEGIN_CF_RULE14,
            &bar_rule.serialize_extension14().unwrap(),
        )
        .unwrap();
    writer
        .write_record(kind::BEGIN_DATABAR14, &bar.serialize_header().unwrap())
        .unwrap();
    for threshold in [&bar.min_cfvo, &bar.max_cfvo] {
        writer
            .write_record(kind::CFVO14, &threshold.serialize_extension14().unwrap())
            .unwrap();
    }
    for color in [bar.positive_color, bar.axis_color].into_iter().flatten() {
        writer
            .write_record(kind::COLOR14, &color.serialize_extension14().unwrap())
            .unwrap();
    }
    writer.write_record(kind::END_DATABAR14, &[]).unwrap();
    writer.write_record(kind::END_CF_RULE14, &[]).unwrap();

    writer
        .write_record(
            kind::BEGIN_CF_RULE14,
            &icon_rule.serialize_extension14().unwrap(),
        )
        .unwrap();
    writer
        .write_record(
            kind::BEGIN_ICON_SET14,
            &icon_set.serialize_header().unwrap(),
        )
        .unwrap();
    for threshold in &icon_thresholds {
        writer
            .write_record(kind::CFVO14, &threshold.serialize_extension14().unwrap())
            .unwrap();
    }
    writer.write_record(kind::END_ICON_SET14, &[]).unwrap();
    writer.write_record(kind::END_CF_RULE14, &[]).unwrap();
    writer
        .write_record(kind::END_COND_FORMATTING14, &[])
        .unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();

    let formula_context = Context::default();
    let iter = Stream::new(std::io::Cursor::new(worksheet));
    let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
    assert!(reader.next_cell().unwrap().is_none());
    let parsed = &reader.conditional_formattings[0];
    assert_eq!(parsed.record_kind, RecordKind::Extension14);
    assert_eq!(parsed.rules.len(), 3);
    assert_eq!(
        parsed.rules[0]
            .color_scale14
            .as_ref()
            .unwrap()
            .min_cfvo
            .cfvo_type,
        2
    );
    let parsed_bar = parsed.rules[1].data_bar14.as_ref().unwrap();
    assert_eq!(parsed_bar.axis_position, AxisPosition14::Automatic);
    assert_eq!(parsed_bar.positive_color.unwrap().argb, Some(0xff44_72c4));
    let parsed_icons = parsed.rules[2].icon_set14.as_ref().unwrap();
    assert_eq!(parsed_icons.icon_set_type, 18);
    assert_eq!(parsed_icons.cfvos.len(), 3);
}

#[test]
fn rejects_incomplete_color_scale_collection() {
    let mut worksheet = Vec::new();
    let mut writer = Writer::new(&mut worksheet);
    writer.write_record(kind::WS_DIM, &[0; 16]).unwrap();
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    let mut begin = 1u32.to_le_bytes().to_vec();
    begin.extend_from_slice(&0u32.to_le_bytes());
    begin.extend_from_slice(&1u32.to_le_bytes());
    begin.extend_from_slice(&[0; 16]);
    writer
        .write_record(kind::BEGIN_COND_FORMATTING, &begin)
        .unwrap();

    let mut rule = 3u32.to_le_bytes().to_vec();
    rule.extend_from_slice(&2u32.to_le_bytes());
    rule.extend_from_slice(&u32::MAX.to_le_bytes());
    rule.extend_from_slice(&1u32.to_le_bytes());
    rule.extend_from_slice(&[0; 10]);
    rule.extend_from_slice(&[0; 12]);
    rule.extend_from_slice(&u32::MAX.to_le_bytes());
    writer.write_record(kind::BEGIN_CF_RULE, &rule).unwrap();
    writer.write_record(kind::BEGIN_COLOR_SCALE, &[]).unwrap();
    writer.write_record(kind::END_COLOR_SCALE, &[]).unwrap();
    writer.write_record(kind::END_CF_RULE, &[]).unwrap();
    writer.write_record(kind::END_COND_FORMATTING, &[]).unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();

    let formula_context = Context::default();
    let iter = Stream::new(std::io::Cursor::new(worksheet));
    let mut reader = CellsReader::new(iter, &[], &formula_context, 1).unwrap();
    assert!(reader.next_cell().is_err());
}
