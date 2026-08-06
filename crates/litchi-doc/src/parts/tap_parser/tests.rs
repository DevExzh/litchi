//! TAP parser regression and [MS-DOC] validation tests.

use super::super::tap::{
    BorderStyle, BorderType, CellBorderTypes, CellMergeStatus, CellShading, CellSpacing,
    CellSpacingSource, ShadingPattern, TableHorizontalAnchor, TableHorizontalPosition,
    TableJustification, TableLook, TableLookFlags, TablePositioning, TableStyleBorder,
    TableStyleCondition, TableStyleDefaults, TableStyleShading, TableVerticalAnchor,
    TableVerticalPosition, TableWidth, TextDirection, VerticalAlignment, VerticalMergeStatus,
    WidthType,
};
use super::TapParser;
use bumpalo::Bump;

fn table_definition_grpprl(operand: &[u8]) -> Vec<u8> {
    let mut grpprl = 0xD608u16.to_le_bytes().to_vec();
    grpprl.extend_from_slice(&u16::try_from(operand.len() + 1).unwrap().to_le_bytes());
    grpprl.extend_from_slice(operand);
    grpprl
}

fn single_cell_definition_grpprl(flags: u16, width: i16) -> Vec<u8> {
    let mut operand = vec![1, 0, 0];
    operand.extend_from_slice(&width.to_le_bytes());
    operand.extend_from_slice(&flags.to_le_bytes());
    operand.extend_from_slice(&width.to_le_bytes());
    operand.extend_from_slice(&[0; 16]);
    table_definition_grpprl(&operand)
}

fn append_variable_sprm(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
    grpprl.extend_from_slice(&opcode.to_le_bytes());
    grpprl.push(u8::try_from(operand.len()).unwrap());
    grpprl.extend_from_slice(operand);
}

fn append_fixed_sprm(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
    grpprl.extend_from_slice(&opcode.to_le_bytes());
    grpprl.extend_from_slice(operand);
}

fn full_shading(foreground: [u8; 4], background: [u8; 4], pattern: u16) -> Vec<u8> {
    let mut shading = foreground.to_vec();
    shading.extend_from_slice(&background);
    shading.extend_from_slice(&pattern.to_le_bytes());
    shading
}

fn full_border(color: [u8; 4], width: u8, border_type: u8, effects: u8) -> Vec<u8> {
    let mut border = color.to_vec();
    border.extend_from_slice(&[width, border_type, effects, 0]);
    border
}

#[test]
fn test_tap_parser_creation() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);

    // Simple SPRM data: sprmTDefTable with 2 cells
    // Format: opcode(2) + size(2) + itcMac(1) + boundaries(3*2)
    let sprm_data = vec![
        0x08, 0xD6, // sprmTDefTable (0xD608)
        0x08, 0x00, // size = 8 bytes (after this size field)
        0x02, // itcMac = 2 cells
        0x00, 0x00, // boundary 0 = 0 twips
        0x64, 0x00, // boundary 1 = 100 twips
        0xC8, 0x00, // boundary 2 = 200 twips
    ];

    let tap = parser.parse_tap(&sprm_data).unwrap();
    assert_eq!(tap.cell_count, 2);
    // For 2 cells, we should have 3 boundaries (start, middle, end)
    // But if initialization adds more, we just check the count is correct
    assert_eq!(tap.cell_boundaries.len(), 3);
}

#[test]
fn rejects_malformed_table_definitions() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);

    assert!(parser.parse_tap(&table_definition_grpprl(&[])).is_err());
    assert!(parser.parse_tap(&table_definition_grpprl(&[0])).is_err());
    assert!(parser.parse_tap(&table_definition_grpprl(&[64])).is_err());
    assert!(
        parser
            .parse_tap(&table_definition_grpprl(&[2, 0, 0, 100, 0]))
            .is_err()
    );
    assert!(
        parser
            .parse_tap(&table_definition_grpprl(&[2, 0, 0, 200, 0, 100, 0]))
            .is_err()
    );
    assert!(
        parser
            .parse_tap(&table_definition_grpprl(&[1, 0x3F, 0x84, 0, 0]))
            .is_err()
    );

    let empty = parser
        .parse_tap(&table_definition_grpprl(&[0, 0, 0]))
        .unwrap();
    assert_eq!(empty.cell_count, 0);
    assert_eq!(empty.cell_boundaries, [0]);

    let mut excess_descriptors = vec![1, 0, 0, 100, 0];
    excess_descriptors.extend_from_slice(&[0; 40]);
    let tap = parser
        .parse_tap(&table_definition_grpprl(&excess_descriptors))
        .unwrap();
    assert_eq!(tap.cell_properties.len(), 1);

    let mut partial_descriptors = vec![2, 0, 0, 100, 0, 200, 0];
    partial_descriptors.extend_from_slice(&[0; 20]);
    let tap = parser
        .parse_tap(&table_definition_grpprl(&partial_descriptors))
        .unwrap();
    assert_eq!(tap.cell_properties.len(), 2);
}

#[test]
fn decodes_tc80_layout_bits_and_width_type() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let flags = 2u16 | (3 << 2) | (3 << 5) | (2 << 7) | (3 << 9) | 0x7000;

    let tap = parser
        .parse_tap(&single_cell_definition_grpprl(flags, 1440))
        .unwrap();
    let cell = &tap.cell_properties[0];
    assert_eq!(cell.merge_status, CellMergeStatus::First);
    assert_eq!(cell.vertical_merge_status, VerticalMergeStatus::First);
    assert_eq!(cell.text_direction, TextDirection::BtLr);
    assert_eq!(cell.vertical_alignment, VerticalAlignment::Bottom);
    assert!(cell.fit_text);
    assert!(cell.no_wrap);
    assert!(cell.hide_mark);
    assert!(cell.direct_style.vertical_alignment);
    assert!(cell.direct_style.no_wrap);
    assert!(cell.direct_style.border_top);
    assert!(cell.direct_style.border_left);
    assert!(cell.direct_style.border_bottom);
    assert!(cell.direct_style.border_right);
    let width = cell.preferred_width.unwrap();
    assert_eq!(width.value, 1440);
    assert_eq!(width.width_type, WidthType::Twips);
}

#[test]
fn rejects_invalid_tc80_layout_values() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    for flags in [2 << 2, 2 << 5, 3 << 7, 4 << 9] {
        assert!(
            parser
                .parse_tap(&single_cell_definition_grpprl(flags, 1440))
                .is_err(),
            "flags {flags:#06x} should be rejected"
        );
    }
    assert!(
        parser
            .parse_tap(&single_cell_definition_grpprl(3 << 9, -1))
            .is_err()
    );
}

#[test]
fn parses_default_and_cell_range_padding() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
    append_variable_sprm(&mut grpprl, 0xD634, &[0, 1, 0x0F, 3, 108, 0]);
    append_variable_sprm(&mut grpprl, 0xD632, &[1, 2, 0x08, 3, 240, 0]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.cell_properties[0].padding_top, Some(108));
    assert_eq!(tap.cell_properties[0].padding_right, Some(108));
    assert_eq!(tap.cell_properties[1].padding_left, Some(108));
    assert_eq!(tap.cell_properties[1].padding_right, Some(240));
}

#[test]
fn parses_and_resets_default_cell_spacing() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
    append_variable_sprm(&mut grpprl, 0xD633, &[0, 1, 0x0F, 0x13, 240, 0]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.cell_spacing,
        Some(CellSpacing {
            width: 240,
            source: CellSpacingSource::TableBorder,
        })
    );

    append_variable_sprm(&mut grpprl, 0xD633, &[0, 1, 0x0F, 3, 120, 0]);
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.cell_spacing,
        Some(CellSpacing {
            width: 120,
            source: CellSpacingSource::Explicit,
        })
    );

    append_variable_sprm(&mut grpprl, 0xD633, &[0, 1, 0x0F, 0, 0, 0]);
    assert!(parser.parse_tap(&grpprl).unwrap().cell_spacing.is_none());
}

#[test]
fn rejects_malformed_cell_padding_and_shading() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
        append_variable_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };

    assert!(parse_with(0xD634, &[0, 2, 0x0F, 3, 0, 0]).is_err());
    assert!(parse_with(0xD632, &[2, 2, 0x0F, 3, 0, 0]).is_err());
    assert!(parse_with(0xD632, &[0, 2, 0x10, 3, 0, 0]).is_err());
    assert!(parse_with(0xD632, &[0, 2, 0x0F, 1, 0, 0]).is_err());
    assert!(parse_with(0xD632, &[0, 2, 0x0F, 0, 1, 0]).is_err());
    assert!(parse_with(0xD632, &[0, 2, 0x0F, 3, 0xC1, 0x7B]).is_err());
    assert!(parse_with(0xD633, &[1, 1, 0x0F, 3, 0, 0]).is_err());
    assert!(parse_with(0xD633, &[0, 1, 0x0E, 3, 0, 0]).is_err());
    assert!(parse_with(0xD633, &[0, 1, 0x0F, 0, 1, 0]).is_err());
    assert!(parse_with(0xD633, &[0, 1, 0x0F, 1, 0, 0]).is_err());
    assert!(parse_with(0xD633, &[0, 1, 0x0F, 3, 0xE1, 0x3D]).is_err());
    assert!(parse_with(0xD609, &[0]).is_err());
    assert!(parse_with(0xD609, &[0, 0, 0, 0, 0, 0]).is_err());
    assert!(parse_with(0xD609, &[17, 0]).is_err());
    assert!(parse_with(0xD609, &[0, 0x68]).is_err());

    // sprmTTlp has operation 0x0A and must not be interpreted as Shd80.
    let mut grpprl = table_definition_grpprl(&[1, 0, 0, 100, 0]);
    grpprl.extend_from_slice(&0x740Au16.to_le_bytes());
    grpprl.extend_from_slice(&0u32.to_le_bytes());
    assert!(parser.parse_tap(&grpprl).is_ok());
}

#[test]
fn parses_full_color_range_and_raw_shading() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[4, 0, 0, 100, 0, 200, 0, 44, 1, 144, 1]);
    let blue_on_red = full_shading([0, 0, 255, 0], [255, 0, 0, 0], 0x12);
    let green = full_shading([0, 0, 0, 0xFF], [0, 255, 0, 0], 0);
    let nil = full_shading([0xFF; 4], [0xFF; 4], 0);
    let mut range_shading = vec![1, 3];
    range_shading.extend_from_slice(&blue_on_red);
    append_variable_sprm(&mut grpprl, 0xD62D, &range_shading);
    let mut odd_shading = vec![0, 4];
    odd_shading.extend_from_slice(&green);
    append_variable_sprm(&mut grpprl, 0xD62E, &odd_shading);
    append_variable_sprm(&mut grpprl, 0xD670, &nil);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert!(tap.cell_properties[0].shading_inherits_from_style);
    assert!(!tap.cell_properties[0].direct_style.shading);
    assert_eq!(tap.cell_properties[1].background_color, Some((255, 0, 0)));
    assert!(tap.cell_properties[1].direct_style.shading);
    assert_eq!(
        tap.cell_properties[1].shading.unwrap().pattern,
        ShadingPattern::DarkCross
    );
    assert_eq!(tap.cell_properties[2].background_color, Some((0, 255, 0)));
    assert!(tap.cell_properties[3].shading.is_none());

    let mut whole_table = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
    append_variable_sprm(&mut whole_table, 0xD660, &blue_on_red);
    let tap = parser.parse_tap(&whole_table).unwrap();
    assert_eq!(
        tap.cell_properties[0].shading,
        tap.cell_properties[1].shading
    );
    assert_eq!(tap.cell_properties[0].background_color, Some((255, 0, 0)));
}

#[test]
fn rejects_malformed_full_color_shading() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
        append_variable_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };
    assert!(parse_with(0xD612, &[0]).is_err());
    assert!(parse_with(0xD612, &[0; 30]).is_err());
    assert!(parse_with(0xD62D, &[0; 11]).is_err());
    assert!(parse_with(0xD660, &[0; 9]).is_err());
    assert!(parse_with(0xD62D, &[2, 2, 0, 0, 0, 0xFF, 0, 0, 0, 0xFF, 0, 0]).is_err());
    assert!(parse_with(0xD62D, &[0, 2, 0, 0, 0, 1, 0, 0, 0, 0xFF, 0, 0]).is_err());
    assert!(parse_with(0xD62D, &[0, 2, 0, 0, 0, 0xFF, 0, 0, 0, 0xFF, 0x1A, 0]).is_err());
}

#[test]
fn parses_row_range_diagonal_and_color_borders() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);

    append_variable_sprm(&mut grpprl, 0xD620, &[0, 2, 0x01, 8, 1, 1, 0]);
    append_variable_sprm(&mut grpprl, 0xD61A, &[1, 2, 3, 0, 0xFF, 0xFF, 0xFF, 0xFF]);
    let diagonal = full_border([10, 20, 30, 0], 4, 0x1A, 0x41);
    let mut range = vec![0, 2, 0x30];
    range.extend_from_slice(&diagonal);
    append_variable_sprm(&mut grpprl, 0xD62F, &range);
    let mut row_borders = full_border([40, 50, 60, 0], 6, 3, 0);
    row_borders.resize(48, 0);
    append_variable_sprm(&mut grpprl, 0xD613, &row_borders);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.cell_properties[0].borders.top.unwrap().color,
        Some((1, 2, 3))
    );
    assert!(tap.cell_properties[1].borders.top.is_none());
    let diagonal = tap.cell_properties[0].borders.diagonal_down.unwrap();
    assert_eq!(diagonal.border_type, BorderType::Outset);
    assert_eq!(diagonal.color, Some((10, 20, 30)));
    assert_eq!(tap.cell_properties[1].borders.diagonal_up, Some(diagonal));
    assert_eq!(tap.border_top.unwrap().color, Some((40, 50, 60)));
}

#[test]
fn rejects_malformed_modern_table_borders() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
        append_variable_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };
    assert!(parse_with(0xD605, &[0; 23]).is_err());
    assert!(parse_with(0xD613, &[0; 47]).is_err());
    assert!(parse_with(0xD61A, &[0; 4]).is_err());
    assert!(parse_with(0xD620, &[0; 6]).is_err());
    assert!(parse_with(0xD62F, &[2, 2, 1, 0, 0, 0, 0xFF, 8, 1, 0, 0]).is_err());
    assert!(parse_with(0xD62F, &[0, 2, 0x40, 0, 0, 0, 0xFF, 8, 1, 0, 0]).is_err());
    assert!(parse_with(0xD62F, &[0, 2, 1, 0, 0, 0, 1, 8, 1, 0, 0]).is_err());
    assert!(parse_with(0xD62F, &[0, 2, 1, 0, 0, 0, 0xFF, 8, 2, 0, 0]).is_err());
    assert!(TapParser::parse_border_code(&[8, 0x1A, 1, 0], 0).is_err());
}

#[test]
fn parses_cell_border_type_prefix_overrides() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
    let mut top_border = vec![0, 1, 0x01];
    top_border.extend_from_slice(&full_border([1, 2, 3, 0], 8, 1, 0));
    append_variable_sprm(&mut grpprl, 0xD62F, &top_border);
    append_variable_sprm(&mut grpprl, 0xD662, &[3, 6, 0, 0x1A]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.cell_properties[0].border_type_overrides,
        CellBorderTypes {
            top: Some(BorderType::Double),
            left: Some(BorderType::Dotted),
            bottom: Some(BorderType::None),
            right: Some(BorderType::Outset),
        }
    );
    assert_eq!(
        tap.cell_properties[0].borders.top.unwrap().border_type,
        BorderType::Double
    );
    assert_eq!(
        tap.cell_properties[1].border_type_overrides,
        CellBorderTypes::default()
    );
}

#[test]
fn rejects_malformed_cell_border_type_prefixes() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
        append_variable_sprm(&mut grpprl, 0xD662, operand);
        parser.parse_tap(&grpprl)
    };
    assert!(parse_with(&[1, 1, 1]).is_err());
    assert!(parse_with(&[1; 12]).is_err());
    assert!(parse_with(&[2, 1, 1, 1]).is_err());
    assert!(parse_with(&[0x1C, 1, 1, 1]).is_err());
}

#[test]
fn parses_cell_range_layout_overrides() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[3, 0, 0, 100, 0, 200, 0, 44, 1]);
    append_fixed_sprm(&mut grpprl, 0x7629, &[0, 2, 5, 0]);
    append_variable_sprm(&mut grpprl, 0xD62B, &[1, 3]);
    append_variable_sprm(&mut grpprl, 0xD62C, &[1, 3, 2]);
    append_variable_sprm(&mut grpprl, 0xD635, &[0, 2, 2, 0xC4, 0x09]);
    append_fixed_sprm(&mut grpprl, 0xF636, &[0, 3, 1]);
    append_variable_sprm(&mut grpprl, 0xD639, &[1, 2, 1]);
    append_variable_sprm(&mut grpprl, 0xD642, &[2, 3, 1]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.cell_properties[0].text_direction, TextDirection::TbLr);
    assert_eq!(tap.cell_properties[1].text_direction, TextDirection::TbLr);
    assert_eq!(
        tap.cell_properties[1].vertical_merge_status,
        VerticalMergeStatus::First
    );
    assert_eq!(
        tap.cell_properties[2].vertical_alignment,
        VerticalAlignment::Bottom
    );
    let width = tap.cell_properties[0].preferred_width.unwrap();
    assert_eq!(width.width_type, WidthType::Percentage);
    assert_eq!(width.value, 2500);
    assert!(tap.cell_properties.iter().all(|cell| cell.fit_text));
    assert!(tap.cell_properties[1].no_wrap);
    assert!(tap.cell_properties[1].direct_style.no_wrap);
    assert!(tap.cell_properties[2].direct_style.vertical_alignment);
    assert!(tap.cell_properties[2].hide_mark);
}

#[test]
fn rejects_malformed_cell_range_layout_overrides() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_variable = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
        append_variable_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };
    let parse_fixed = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };
    assert!(parse_fixed(0x7629, &[0, 2, 2, 0]).is_err());
    assert!(parse_fixed(0x7629, &[2, 2, 0, 0]).is_err());
    assert!(parse_variable(0xD62B, &[2, 0]).is_err());
    assert!(parse_variable(0xD62B, &[0, 2]).is_err());
    assert!(parse_variable(0xD62C, &[0, 2, 3]).is_err());
    assert!(parse_variable(0xD635, &[0, 2, 2, 0x89, 0x13]).is_err());
    assert!(parse_variable(0xD635, &[0, 2, 3, 0xC1, 0x7B]).is_err());
    assert!(parse_fixed(0xF636, &[0, 2, 2]).is_err());
    assert!(parse_variable(0xD639, &[0, 3, 1]).is_err());
    assert!(parse_variable(0xD642, &[0, 2]).is_err());
}

#[test]
fn applies_structural_cell_modifiers_in_sequence() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut merged = table_definition_grpprl(&[3, 0, 0, 100, 0, 200, 0, 44, 1]);

    // Mark the original last cell so insertion/deletion property movement is observable.
    append_variable_sprm(&mut merged, 0xD639, &[2, 3, 1]);
    append_fixed_sprm(&mut merged, 0x7621, &[1, 2, 50, 0]);
    append_fixed_sprm(&mut merged, 0x7623, &[2, 4, 75, 0]);
    append_fixed_sprm(&mut merged, 0x5624, &[1, 4]);

    let tap = parser.parse_tap(&merged).unwrap();
    assert_eq!(tap.cell_count, 5);
    assert_eq!(tap.cell_boundaries, [0, 100, 150, 225, 300, 400]);
    assert_eq!(tap.cell_properties.len(), 5);
    assert!(tap.cell_properties[4].no_wrap);
    assert_eq!(
        tap.cell_properties
            .iter()
            .map(|cell| cell.merge_status)
            .collect::<Vec<_>>(),
        [
            CellMergeStatus::None,
            CellMergeStatus::First,
            CellMergeStatus::Merged,
            CellMergeStatus::Merged,
            CellMergeStatus::None,
        ]
    );

    append_fixed_sprm(&mut merged, 0x5625, &[1, 4]);
    append_fixed_sprm(&mut merged, 0x5622, &[1, 3]);
    let tap = parser.parse_tap(&merged).unwrap();
    assert_eq!(tap.cell_count, 3);
    assert_eq!(tap.cell_boundaries, [0, 225, 300, 400]);
    assert_eq!(tap.cell_properties.len(), 3);
    assert!(
        tap.cell_properties
            .iter()
            .all(|cell| cell.merge_status == CellMergeStatus::None)
    );
    assert!(!tap.cell_properties[0].no_wrap);
    assert!(!tap.cell_properties[1].no_wrap);
    assert!(tap.cell_properties[2].no_wrap);

    let mut empty = table_definition_grpprl(&[0, 0, 0]);
    append_fixed_sprm(&mut empty, 0x7621, &[0, 2, 120, 0]);
    let tap = parser.parse_tap(&empty).unwrap();
    assert_eq!(tap.cell_count, 2);
    assert_eq!(tap.cell_boundaries, [0, 120, 240]);
}

#[test]
fn rejects_malformed_structural_cell_modifiers() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[3, 0, 0, 100, 0, 200, 0, 44, 1]);
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };

    assert!(parse_with(0x7621, &[4, 1, 10, 0]).is_err());
    assert!(parse_with(0x7621, &[1, 0, 10, 0]).is_err());
    assert!(parse_with(0x7621, &[1, 61, 10, 0]).is_err());
    assert!(parse_with(0x7621, &[1, 1, 0xC1, 0x7B]).is_err());
    assert!(parse_with(0x7621, &[1, 2, 0x80, 0x3E]).is_err());

    assert!(parse_with(0x5622, &[0, 3]).is_err());
    assert!(parse_with(0x5622, &[2, 1]).is_err());
    assert!(parse_with(0x5622, &[3, 3]).is_err());

    assert!(parse_with(0x7623, &[2, 1, 10, 0]).is_err());
    assert!(parse_with(0x7623, &[0, 1, 0xC1, 0x7B]).is_err());
    let mut overflowing = table_definition_grpprl(&[1, 0x30, 0x75, 0x94, 0x75]);
    append_fixed_sprm(&mut overflowing, 0x7623, &[0, 1, 0xD0, 0x07]);
    assert!(parser.parse_tap(&overflowing).is_err());

    assert!(parse_with(0x5624, &[2, 1]).is_err());
    assert!(parse_with(0x5624, &[3, 3]).is_err());
    assert!(parse_with(0x5625, &[2, 4]).is_err());

    // ItcFirstLim explicitly permits empty and one-cell ranges.
    assert!(parse_with(0x5624, &[1, 1]).is_ok());
    let tap = parse_with(0x5624, &[1, 2]).unwrap();
    assert_eq!(tap.cell_properties[1].merge_status, CellMergeStatus::First);
}

#[test]
fn applies_core_row_geometry_and_pagination_strictly() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[2, 0, 0, 100, 0, 200, 0]);
    append_fixed_sprm(&mut grpprl, 0x5400, &[2, 0]);
    append_fixed_sprm(&mut grpprl, 0x9602, &[20, 0]);
    append_fixed_sprm(&mut grpprl, 0x9601, &[200, 0]);
    append_fixed_sprm(&mut grpprl, 0x3403, &[0]);
    append_fixed_sprm(&mut grpprl, 0x3404, &[1]);
    append_fixed_sprm(&mut grpprl, 0x9407, &(-240i16).to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x3466, &[1]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.justification, TableJustification::Right);
    assert_eq!(tap.indent_left, 200);
    assert_eq!(tap.gap_half, 20);
    assert_eq!(tap.cell_boundaries, [180, 300, 400]);
    assert_eq!(tap.row_height, Some(-240));
    assert!(tap.is_header_row);
    assert!(!tap.allow_row_break);

    append_fixed_sprm(&mut grpprl, 0x9407, &[0, 0]);
    append_fixed_sprm(&mut grpprl, 0x3466, &[0]);
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.row_height, None);
    assert!(tap.allow_row_break);
}

#[test]
fn resolves_physical_and_logical_table_justification() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut physical = table_definition_grpprl(&[1, 0, 0, 100, 0]);
    append_fixed_sprm(&mut physical, 0x5400, &[2, 0]);
    append_fixed_sprm(&mut physical, 0x560B, &[1, 0]);

    let tap = parser.parse_tap(&physical).unwrap();
    assert_eq!(tap.justification, TableJustification::Left);
    assert_eq!(
        tap.legacy_physical_justification,
        Some(TableJustification::Right)
    );
    assert_eq!(tap.modern_logical_justification, None);

    let mut logical = physical;
    append_fixed_sprm(&mut logical, 0x548A, &[2, 0]);
    // A later compatibility property cannot override the modern logical one.
    append_fixed_sprm(&mut logical, 0x5400, &[1, 0]);
    let tap = parser.parse_tap(&logical).unwrap();
    assert_eq!(tap.justification, TableJustification::Right);
    assert_eq!(
        tap.legacy_physical_justification,
        Some(TableJustification::Center)
    );
    assert_eq!(
        tap.modern_logical_justification,
        Some(TableJustification::Right)
    );
}

#[test]
fn rejects_malformed_core_row_geometry_and_pagination() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[1, 0, 0, 100, 0]);
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };

    assert!(parse_with(0x5400, &[3, 0]).is_err());
    assert!(parse_with(0x548A, &[3, 0]).is_err());
    assert!(parse_with(0x9601, &(-31_681i16).to_le_bytes()).is_err());
    assert!(parse_with(0x9602, &(-1i16).to_le_bytes()).is_err());
    assert!(parse_with(0x9602, &31_681i16.to_le_bytes()).is_err());
    assert!(parse_with(0x3403, &[2]).is_err());
    assert!(parse_with(0x3404, &[2]).is_err());
    assert!(parse_with(0x3466, &[2]).is_err());
    assert!(parse_with(0x9407, &(-31_681i16).to_le_bytes()).is_err());
    assert!(parse_with(0x9407, &31_681i16.to_le_bytes()).is_err());

    let mut shifted = table_definition_grpprl(&[1, 0x30, 0x75, 0x94, 0x75]);
    append_fixed_sprm(&mut shifted, 0x9601, &31_680i16.to_le_bytes());
    assert!(parser.parse_tap(&shifted).is_err());

    let mut reordered = table_definition_grpprl(&[1, 0, 0, 100, 0]);
    append_fixed_sprm(&mut reordered, 0x9602, &[200, 0]);
    reordered.extend_from_slice(&table_definition_grpprl(&[1, 0, 0, 100, 0]));
    append_fixed_sprm(&mut reordered, 0x9602, &[0, 0]);
    assert!(parser.parse_tap(&reordered).is_err());
}

#[test]
fn parses_table_sizing_and_fit_properties() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[2, 0, 0, 232, 3, 208, 7]);
    append_fixed_sprm(&mut grpprl, 0xF614, &[2, 0x30, 0x75]);
    append_fixed_sprm(&mut grpprl, 0x3615, &[1]);
    append_fixed_sprm(&mut grpprl, 0xF617, &[2, 0x88, 0x13]);
    append_fixed_sprm(&mut grpprl, 0xF618, &[3, 200, 0]);
    append_fixed_sprm(&mut grpprl, 0x3619, &[1]);
    append_fixed_sprm(&mut grpprl, 0xF661, &[3, 0x9C, 0xFF]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.preferred_width,
        Some(TableWidth {
            value: 30_000,
            width_type: WidthType::Percentage,
        })
    );
    assert!(tap.auto_fit);
    assert_eq!(
        tap.width_before,
        Some(TableWidth {
            value: 5_000,
            width_type: WidthType::Percentage,
        })
    );
    assert_eq!(
        tap.width_after,
        Some(TableWidth {
            value: 200,
            width_type: WidthType::Twips,
        })
    );
    assert_eq!(
        tap.preferred_indent,
        Some(TableWidth {
            value: -100,
            width_type: WidthType::Twips,
        })
    );
    assert!(tap.keep_with_next);

    append_fixed_sprm(&mut grpprl, 0xF614, &[0, 0, 0]);
    append_fixed_sprm(&mut grpprl, 0x3615, &[0]);
    append_fixed_sprm(&mut grpprl, 0xF617, &[0, 0xFF, 0xFF]);
    append_fixed_sprm(&mut grpprl, 0x3619, &[0]);
    append_fixed_sprm(&mut grpprl, 0xF661, &[1, 0, 0]);
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert!(tap.preferred_width.is_none());
    assert!(!tap.auto_fit);
    assert!(tap.width_before.is_none());
    assert!(!tap.keep_with_next);
    assert_eq!(
        tap.preferred_indent,
        Some(TableWidth {
            value: 0,
            width_type: WidthType::Auto,
        })
    );
}

#[test]
fn rejects_malformed_table_sizing_and_fit_properties() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[1, 0, 0, 232, 3]);
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };

    for operand in [
        [0, 1, 0],
        [1, 1, 0],
        [2, 0x31, 0x75],
        [3, 0xC1, 0x7B],
        [0x13, 0, 0],
    ] {
        assert!(parse_with(0xF614, &operand).is_err());
    }
    for operand in [[1, 1, 0], [2, 0x89, 0x13], [3, 0xC1, 0x7B], [0x13, 0, 0]] {
        assert!(parse_with(0xF617, &operand).is_err());
    }
    for operand in [
        [0, 1, 0],
        [1, 1, 0],
        [2, 0, 0],
        [3, 0xB7, 0x84],
        [0x13, 0, 0],
    ] {
        assert!(parse_with(0xF661, &operand).is_err());
    }
    assert!(parse_with(0x3615, &[2]).is_err());
    assert!(parse_with(0x3619, &[2]).is_err());

    let mut beyond_right_edge = table_definition_grpprl(&[1, 0, 0, 232, 3]);
    append_fixed_sprm(&mut beyond_right_edge, 0xF661, &[3, 0xF8, 0x77]);
    append_fixed_sprm(&mut beyond_right_edge, 0xF614, &[3, 0xE8, 0x03]);
    assert!(parser.parse_tap(&beyond_right_edge).is_err());
}

#[test]
fn parses_scalar_table_style_defaults() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = Vec::new();
    append_variable_sprm(&mut grpprl, 0xD63E, &[0, 1, 0x05, 3, 120, 0]);
    append_variable_sprm(&mut grpprl, 0xD63E, &[0, 1, 0x0A, 3, 240, 0]);
    append_fixed_sprm(&mut grpprl, 0x347C, &[2]);
    append_fixed_sprm(&mut grpprl, 0x347D, &[0]);
    append_fixed_sprm(&mut grpprl, 0x3488, &[2]);
    append_fixed_sprm(&mut grpprl, 0x3489, &[3]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.style_defaults,
        TableStyleDefaults {
            padding_top: Some(120),
            padding_left: Some(240),
            padding_bottom: Some(120),
            padding_right: Some(240),
            vertical_alignment: Some(VerticalAlignment::Bottom),
            no_wrap: Some(false),
            horizontal_band_size: Some(2),
            vertical_band_size: Some(3),
            ..TableStyleDefaults::default()
        }
    );
}

#[test]
fn rejects_malformed_scalar_table_style_defaults() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_variable = |operand: &[u8]| {
        let mut grpprl = Vec::new();
        append_variable_sprm(&mut grpprl, 0xD63E, operand);
        parser.parse_tap(&grpprl)
    };
    assert!(parse_variable(&[1, 1, 1, 3, 0, 0]).is_err());
    assert!(parse_variable(&[0, 1, 0, 3, 0, 0]).is_err());
    assert!(parse_variable(&[0, 1, 1, 0, 0, 0]).is_err());
    assert!(parse_variable(&[0, 1, 1, 3, 0xC1, 0x7B]).is_err());

    let parse_fixed = |opcode, operand: &[u8]| {
        let mut grpprl = Vec::new();
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };
    assert!(parse_fixed(0x347C, &[3]).is_err());
    assert!(parse_fixed(0x347D, &[2]).is_err());
    assert!(parse_fixed(0x3488, &[0]).is_err());
    assert!(parse_fixed(0x3489, &[4]).is_err());
}

#[test]
fn parses_visual_table_style_defaults() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut nested = Vec::new();
    append_variable_sprm(&mut nested, 0xD47F, &[0; 8]);
    for (opcode, border_type) in [
        (0xD680, 1),
        (0xD681, 3),
        (0xD682, 6),
        (0xD683, 7),
        (0xD684, 8),
        (0xD685, 0x1A),
        (0xD686, 0x1B),
    ] {
        append_variable_sprm(
            &mut nested,
            opcode,
            &full_border([1, 2, 3, 0], 8, border_type, 0),
        );
    }
    let mut grpprl = Vec::new();
    let mut conditional = 0x0001u16.to_le_bytes().to_vec();
    conditional.extend_from_slice(&nested);
    append_variable_sprm(&mut grpprl, 0xD66A, &conditional);
    let shading = CellShading {
        foreground_color: Some((4, 5, 6)),
        background_color: Some((7, 8, 9)),
        pattern: ShadingPattern::DarkCross,
    };
    append_variable_sprm(
        &mut grpprl,
        0xD687,
        &full_shading([4, 5, 6, 0], [7, 8, 9, 0], shading.pattern as u16),
    );
    // ShdNil is ignored inside a table style and does not clear the prior value.
    append_variable_sprm(&mut grpprl, 0xD687, &full_shading([0xFF; 4], [0xFF; 4], 0));

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.conditional_formats.len(), 1);
    assert_eq!(
        tap.conditional_formats[0].condition,
        TableStyleCondition::HeaderRow
    );
    assert_eq!(tap.conditional_formats[0].raw_grpprl, nested);
    let conditional = &tap.conditional_formats[0].properties;
    assert_eq!(conditional.border_top, Some(TableStyleBorder::NoBorder));
    assert!(matches!(
        conditional.border_bottom,
        Some(TableStyleBorder::Border(BorderStyle {
            border_type: BorderType::Single,
            ..
        }))
    ));
    assert!(matches!(
        conditional.border_diagonal_down,
        Some(TableStyleBorder::Border(BorderStyle {
            border_type: BorderType::Outset,
            ..
        }))
    ));
    assert!(matches!(
        conditional.border_diagonal_up,
        Some(TableStyleBorder::Border(BorderStyle {
            border_type: BorderType::Inset,
            ..
        }))
    ));
    assert_eq!(
        tap.style_defaults.shading,
        Some(TableStyleShading::Shading(shading))
    );

    let mut auto = Vec::new();
    append_variable_sprm(
        &mut auto,
        0xD687,
        &full_shading([0, 0, 0, 0xFF], [0, 0, 0, 0xFF], 0),
    );
    assert_eq!(
        parser.parse_tap(&auto).unwrap().style_defaults.shading,
        Some(TableStyleShading::NoShading)
    );
}

#[test]
fn rejects_malformed_visual_table_style_defaults() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = Vec::new();
        append_variable_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };
    // Style borders are invalid in the outer UpxTapx property list.
    assert!(parse_with(0xD47F, &[0; 8]).is_err());
    let parse_border = |operand: &[u8]| {
        let mut nested = Vec::new();
        append_variable_sprm(&mut nested, 0xD47F, operand);
        let mut conditional = 0x0001u16.to_le_bytes().to_vec();
        conditional.extend_from_slice(&nested);
        parse_with(0xD66A, &conditional)
    };
    assert!(parse_border(&[0; 7]).is_err());
    assert!(parse_border(&[0, 0, 0, 0, 0xFF, 0xFF, 0xFF, 0xFF]).is_err());
    assert!(parse_border(&full_border([0; 4], 8, 2, 0)).is_err());
    assert!(parse_with(0xD687, &[0; 9]).is_err());
    assert!(parse_with(0xD687, &full_shading([0; 4], [0; 4], 99)).is_err());

    assert!(parse_with(0xD66A, &[0, 0]).is_err());
    assert!(parse_with(0xD66A, &[1]).is_err());
    let mut malformed_nested = 0x0001u16.to_le_bytes().to_vec();
    append_variable_sprm(&mut malformed_nested, 0xD47F, &[0; 7]);
    assert!(parse_with(0xD66A, &malformed_nested).is_err());
    let mut recursive = 0x0001u16.to_le_bytes().to_vec();
    append_variable_sprm(&mut recursive, 0xD66A, &[1, 0]);
    assert!(parse_with(0xD66A, &recursive).is_err());
}

#[test]
fn parses_spec_cn_operand_length_and_rejects_disallowed_nested_properties() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);

    // CNFOperand.cb excludes itself and covers cnfc plus grpprl.
    let mut spec_form = Vec::new();
    append_variable_sprm(&mut spec_form, 0xD66A, &[2, 1, 0]);
    let tap = parser.parse_tap(&spec_form).unwrap();
    assert_eq!(tap.conditional_formats.len(), 1);
    assert_eq!(
        tap.conditional_formats[0].condition,
        TableStyleCondition::HeaderRow
    );
    assert!(tap.conditional_formats[0].raw_grpprl.is_empty());

    let mut nested = Vec::new();
    append_variable_sprm(&mut nested, 0xD608, &[1, 0, 0]);
    let mut invalid = 1u16.to_le_bytes().to_vec();
    invalid.extend_from_slice(&nested);
    let mut grpprl = Vec::new();
    append_variable_sprm(&mut grpprl, 0xD66A, &invalid);
    assert!(parser.parse_tap(&grpprl).is_err());
}

#[test]
fn parses_all_table_style_conditions() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let expected: [(u16, TableStyleCondition); 12] = [
        (0x0001, TableStyleCondition::HeaderRow),
        (0x0002, TableStyleCondition::FooterRow),
        (0x0004, TableStyleCondition::FirstColumn),
        (0x0008, TableStyleCondition::LastColumn),
        (0x0010, TableStyleCondition::OddColumnBand),
        (0x0020, TableStyleCondition::EvenColumnBand),
        (0x0040, TableStyleCondition::OddRowBand),
        (0x0080, TableStyleCondition::EvenRowBand),
        (0x0100, TableStyleCondition::TopRightCell),
        (0x0200, TableStyleCondition::TopLeftCell),
        (0x0400, TableStyleCondition::BottomRightCell),
        (0x0800, TableStyleCondition::BottomLeftCell),
    ];
    let mut grpprl = Vec::new();
    for (code, _) in expected {
        append_variable_sprm(&mut grpprl, 0xD66A, &code.to_le_bytes());
    }
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.conditional_formats
            .iter()
            .map(|format| format.condition)
            .collect::<Vec<_>>(),
        expected
            .into_iter()
            .map(|(_, condition)| condition)
            .collect::<Vec<_>>()
    );
}

#[test]
fn parses_table_look_style_direction_and_overlap() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[1, 0, 0, 232, 3]);
    append_fixed_sprm(&mut grpprl, 0x740A, &[0xFF, 0xFF, 0xFF, 0x07]);
    append_fixed_sprm(&mut grpprl, 0x560B, &[1, 0]);
    append_fixed_sprm(&mut grpprl, 0x5664, &[0, 0]);
    append_fixed_sprm(&mut grpprl, 0x563A, &[0x34, 0x12]);
    append_fixed_sprm(&mut grpprl, 0x3465, &[1]);

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.table_look,
        Some(TableLook {
            autoformat_index: -1,
            flags: TableLookFlags::all(),
        })
    );
    assert_eq!(tap.table_style_index, Some(0x1234));
    assert!(tap.legacy_right_to_left);
    assert!(!tap.modern_right_to_left);
    assert!(tap.right_to_left);
    assert!(!tap.allow_overlap);

    append_fixed_sprm(&mut grpprl, 0x560B, &[0, 0]);
    append_fixed_sprm(&mut grpprl, 0x3465, &[0]);
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert!(!tap.right_to_left);
    assert!(tap.allow_overlap);
}

#[test]
fn rejects_malformed_table_look_and_direction() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[1, 0, 0, 232, 3]);
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };

    assert!(parse_with(0x740A, &[0, 0, 0, 0x08]).is_err());
    assert!(parse_with(0x560B, &[2, 0]).is_err());
    assert!(parse_with(0x5664, &[2, 0]).is_err());
    assert!(parse_with(0x3465, &[2]).is_err());
}

#[test]
fn parses_floating_table_position_and_wrap_distances() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let mut grpprl = table_definition_grpprl(&[1, 0, 0, 232, 3]);
    append_fixed_sprm(&mut grpprl, 0x360D, &[0xA0]);
    append_fixed_sprm(&mut grpprl, 0x940E, &(-4i16).to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x940F, &721i16.to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x9410, &120u16.to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x9411, &240u16.to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x941E, &360u16.to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x941F, &480u16.to_le_bytes());

    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(
        tap.positioning,
        Some(TablePositioning {
            vertical_anchor: TableVerticalAnchor::Paragraph,
            horizontal_anchor: TableHorizontalAnchor::Page,
        })
    );
    assert_eq!(tap.horizontal_position, TableHorizontalPosition::Center);
    assert_eq!(tap.vertical_position, TableVerticalPosition::Offset(720));
    assert_eq!(tap.distance_from_text_left, 120);
    assert_eq!(tap.distance_from_text_top, 240);
    assert_eq!(tap.distance_from_text_right, 360);
    assert_eq!(tap.distance_from_text_bottom, 480);

    append_fixed_sprm(&mut grpprl, 0x940E, &(-16i16).to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x940F, &(-20i16).to_le_bytes());
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.horizontal_position, TableHorizontalPosition::Outside);
    assert_eq!(tap.vertical_position, TableVerticalPosition::Outside);
}

#[test]
fn rejects_malformed_floating_table_position() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let parse_with = |opcode, operand: &[u8]| {
        let mut grpprl = table_definition_grpprl(&[1, 0, 0, 232, 3]);
        append_fixed_sprm(&mut grpprl, opcode, operand);
        parser.parse_tap(&grpprl)
    };

    assert!(parse_with(0x360D, &[1]).is_err());
    assert!(parse_with(0x940E, &i16::MIN.to_le_bytes()).is_err());
    assert!(parse_with(0x940E, &i16::MAX.to_le_bytes()).is_err());
    assert!(parse_with(0x940F, &i16::MIN.to_le_bytes()).is_err());
    assert!(parse_with(0x940F, &i16::MAX.to_le_bytes()).is_err());
    for opcode in [0x9410, 0x9411, 0x941E, 0x941F] {
        assert!(parse_with(opcode, &31_681u16.to_le_bytes()).is_err());
    }
}

#[test]
fn test_border_code_parsing() {
    let data = vec![
        0x08, // width = 8 (1 point)
        0x06, // type = dotted
        0x06, // color = red
        0x62, // 2pt spacing, shadow, and frame
    ];

    let border = TapParser::parse_border_code(&data, 0).unwrap();
    assert!(border.is_some());
    let border = border.unwrap();
    assert_eq!(border.width, 8);
    assert_eq!(border.border_type, BorderType::Dotted);
    assert_eq!(border.color, Some((255, 0, 0)));
    assert_eq!(border.spacing, 2);
    assert!(border.shadow);
    assert!(border.frame);

    assert!(
        TapParser::parse_border_code(&[0xFF; 4], 0)
            .unwrap()
            .is_none()
    );
    assert!(TapParser::parse_border_code(&[8, 2, 1, 0], 0).is_err());
    assert!(TapParser::parse_border_code(&[8, 1, 17, 0], 0).is_err());
}

#[test]
fn parses_table_row_revision_state_strictly() {
    let arena = Bump::new();
    let parser = TapParser::new(&arena);
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut grpprl = 0xD667u16.to_le_bytes().to_vec();
    grpprl.push(7);
    grpprl.push(1);
    grpprl.extend_from_slice(&1i16.to_le_bytes());
    grpprl.extend_from_slice(&timestamp.to_le_bytes());
    grpprl.extend_from_slice(&0x3668u16.to_le_bytes());
    grpprl.push(1);
    append_fixed_sprm(&mut grpprl, 0x7469, &0x1020_3040u32.to_le_bytes());
    append_fixed_sprm(&mut grpprl, 0x7479, &0xA1B2_C3D4u32.to_le_bytes());
    let tap = parser.parse_tap(&grpprl).unwrap();
    assert_eq!(tap.has_formatting_revision, Some(true));
    assert_eq!(tap.formatting_revision_author_index, Some(1));
    assert_eq!(tap.formatting_revision_timestamp, Some(timestamp));
    assert!(tap.properties_preserved_for_revision);
    let previous = tap.preserved_properties_for_revision.as_ref().unwrap();
    assert_eq!(previous.has_formatting_revision, Some(true));
    assert_eq!(previous.paragraph_group_id, None);
    assert_eq!(tap.paragraph_group_id, Some(0x1020_3040));
    assert_eq!(tap.revision_save_id, Some(0xA1B2_C3D4));

    let invalid_wall = [0x68, 0x36, 2];
    assert!(parser.parse_tap(&invalid_wall).is_err());
    let reset_wall = [0x68, 0x36, 1, 0x68, 0x36, 0];
    let reset = parser.parse_tap(&reset_wall).unwrap();
    assert!(!reset.properties_preserved_for_revision);
    assert!(reset.preserved_properties_for_revision.is_none());
    let zero_ipgp = [0x69, 0x74, 0, 0, 0, 0];
    assert!(parser.parse_tap(&zero_ipgp).is_err());
    let truncated_rsid = [0x79, 0x74, 1, 2, 3];
    assert!(parser.parse_tap(&truncated_rsid).is_err());
    let invalid_dttm = [0x67, 0xD6, 7, 1, 0, 0, 0x3F, 0, 0, 0];
    assert!(parser.parse_tap(&invalid_dttm).is_err());
}
