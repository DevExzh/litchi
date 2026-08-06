//! Focused stylesheet parser, resolver, and preservation coverage.

use super::codec::*;
use super::model::*;
use super::semantic::*;
use crate::leniency::Leniency;
use crate::package::Result;

fn parse(data: &[u8]) -> Result<StyleSheet> {
    StyleSheet::parse_data(data, 0, Leniency::Strict)
}

fn std_record(
    invariant_id: u16,
    kind: u16,
    base: u16,
    next: u16,
    name: &str,
    property_sets: &[&[u8]],
) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&invariant_id.to_le_bytes());
    data.extend_from_slice(&(kind | (base << 4)).to_le_bytes());
    data.extend_from_slice(&((property_sets.len() as u16) | (next << 4)).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    let units = name.encode_utf16().collect::<Vec<_>>();
    data.extend_from_slice(&(units.len() as u16).to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    data.extend_from_slice(&0u16.to_le_bytes());
    for property_set in property_sets {
        data.extend_from_slice(&(property_set.len() as u16).to_le_bytes());
        data.extend_from_slice(property_set);
        if property_set.len() % 2 != 0 {
            data.push(0);
        }
    }
    let size = data.len() as u16;
    data[6..8].copy_from_slice(&size.to_le_bytes());
    data
}

fn stylesheet(mut slots: Vec<Option<Vec<u8>>>) -> Vec<u8> {
    if slots.len() < 15 {
        slots.resize(15, None);
    }
    let mut data = Vec::new();
    data.extend_from_slice(&18u16.to_le_bytes());
    data.extend_from_slice(&(slots.len() as u16).to_le_bytes());
    data.extend_from_slice(&10u16.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&15u16.to_le_bytes());
    data.extend_from_slice(&15u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0i16.to_le_bytes());
    data.extend_from_slice(&0i16.to_le_bytes());
    data.extend_from_slice(&0i16.to_le_bytes());
    for slot in slots {
        if let Some(std) = slot {
            data.extend_from_slice(&(std.len() as u16).to_le_bytes());
            data.extend_from_slice(&std);
            if std.len() % 2 != 0 {
                data.push(0xA5);
            }
        } else {
            data.extend_from_slice(&0u16.to_le_bytes());
        }
    }
    data
}

fn with_post_2000(std: Vec<u8>, info1: u16, revision_id: u32, info3: u16) -> Vec<u8> {
    let mut extended = Vec::with_capacity(std.len() + 8);
    extended.extend_from_slice(&std[..10]);
    extended.extend_from_slice(&info1.to_le_bytes());
    extended.extend_from_slice(&revision_id.to_le_bytes());
    extended.extend_from_slice(&info3.to_le_bytes());
    extended.extend_from_slice(&std[10..]);
    let size = extended.len() as u16;
    extended[6..8].copy_from_slice(&size.to_le_bytes());
    extended
}

fn valid_stylesheet() -> Vec<u8> {
    let normal = std_record(0, 1, NIL_STYLE, 0, "Normal,正文", &[&[], &[]]);
    let default_font = std_record(65, 2, NIL_STYLE, 10, "Default Paragraph Font", &[&[]]);
    let mut slots = vec![None; 15];
    slots[0] = Some(normal);
    slots[10] = Some(default_font);
    stylesheet(slots)
}

#[test]
fn parses_styles_and_preserves_raw_upx() {
    let tapx = [
        crate::sprm_operations::SPRM_T_JC.to_le_bytes().as_slice(),
        &1u16.to_le_bytes(),
    ]
    .concat();
    let papx = [
        crate::sprm_operations::SPRM_P_F_KEEP
            .to_le_bytes()
            .as_slice(),
        &[1],
    ]
    .concat();
    let chpx = [
        crate::sprm_operations::SPRM_C_F_BOLD
            .to_le_bytes()
            .as_slice(),
        &[1],
    ]
    .concat();
    let mut slots = vec![None; 15];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Default Font", &[&[]]));
    slots[11] = Some(std_record(
        105,
        3,
        0,
        0,
        "Table Normal,Grid Alias",
        &[&tapx, &papx, &chpx],
    ));
    let parsed = parse(&stylesheet(slots)).unwrap();
    assert_eq!(parsed.header().style_count, 15);
    assert_eq!(parsed.styles().len(), 15);
    let normal = parsed.get(0).unwrap();
    assert_eq!(normal.name, "Normal");
    assert_eq!(normal.paragraph_properties(), Some([].as_slice()));
    let table = parsed.get(11).unwrap();
    assert_eq!(table.kind, StyleKind::Table);
    assert_eq!(table.base_style, Some(0));
    assert_eq!(table.aliases, ["Grid Alias"]);
    assert_eq!(table.table_properties(), Some(tapx.as_slice()));
    assert_eq!(table.paragraph_properties(), Some(papx.as_slice()));
    assert_eq!(table.character_properties(), Some(chpx.as_slice()));
}

#[test]
fn parses_the_writer_stylesheet() {
    let data = crate::writer::stylesheet::generate_minimal_stylesheet();
    let parsed = parse(&data).unwrap();
    assert_eq!(parsed.styles().len(), 15);
    assert_eq!(parsed.get(0).unwrap().invariant_id, 0);
    assert_eq!(parsed.get(10).unwrap().invariant_id, 65);
    assert!(parsed.get(13).is_none());
    assert!(parsed.get(14).is_none());
}

#[test]
fn parses_post_2000_metadata_and_preserves_stshi_extensions() {
    let revision = [
        6, 0, // LPUpxRm.cbUpx
        0, 0, 0, 0, // UpxRm.date
        0, 0, // UpxRm.ibstAuthor
        0, 0, // LPUpxPapxRM.cbUpx
        0, 0, // LPUpxChpxRM.cbUpx
    ];
    let normal = std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[], &revision]);
    let normal = with_post_2000(normal, 10 | 0x1000, 0x1234_5678, (42 << 4) | 5);
    let default_font = std_record(65, 2, NIL_STYLE, 10, "Default Font", &[&[]]);
    let default_font = with_post_2000(default_font, 0, 0, 0);
    let mut slots = vec![None; 15];
    slots[0] = Some(normal);
    slots[10] = Some(default_font);
    let mut data = stylesheet(slots);
    data[4..6].copy_from_slice(&18u16.to_le_bytes());
    data.splice(20..20, [4, 0, 0xAA, 0x55]);
    data[0..2].copy_from_slice(&22u16.to_le_bytes());

    let parsed = parse(&data).unwrap();
    assert_eq!(parsed.stshi_tail(), [4, 0, 0xAA, 0x55]);
    let post = parsed.get(0).unwrap().post_2000.as_ref().unwrap();
    assert_eq!(post.linked_style, Some(10));
    assert!(post.has_original_style);
    assert_eq!(post.revision_id, 0x1234_5678);
    assert_eq!(post.html_font_category, 5);
    assert_eq!(post.priority, 42);
}

#[test]
fn validates_revision_marked_style_nested_records() {
    use crate::sprm_operations::{SPRM_C_F_BOLD, SPRM_P_F_KEEP};

    let papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
    let chpx = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
    let mut revision = Vec::new();
    revision.extend_from_slice(&6u16.to_le_bytes());
    revision.extend_from_slice(&0u32.to_le_bytes());
    revision.extend_from_slice(&2i16.to_le_bytes());
    revision.extend_from_slice(&(papx.len() as u16).to_le_bytes());
    revision.extend_from_slice(&papx);
    revision.push(0);
    revision.extend_from_slice(&(chpx.len() as u16).to_le_bytes());
    revision.extend_from_slice(&chpx);
    revision.push(0);

    let parsed = parse_style_revision(&revision, StyleKind::Paragraph, 15).unwrap();
    assert_eq!(parsed.author_index, 2);
    assert_eq!(parsed.author, None);
    assert_eq!(parsed.timestamp, None);
    assert_eq!(parsed.paragraph_properties, Some(papx));
    assert_eq!(parsed.character_properties, chpx);

    let mut wrong_rm_size = revision.clone();
    wrong_rm_size[0..2].copy_from_slice(&5u16.to_le_bytes());
    assert!(parse_style_revision(&wrong_rm_size, StyleKind::Paragraph, 15).is_err());

    let mut bad_inner_padding = revision.clone();
    bad_inner_padding[8 + 2 + 3] = 0xA5;
    assert!(parse_style_revision(&bad_inner_padding, StyleKind::Paragraph, 15).is_err());

    let mut trailing = revision;
    trailing.extend_from_slice(&[0, 0]);
    assert!(parse_style_revision(&trailing, StyleKind::Paragraph, 15).is_err());
}

#[test]
fn enforces_kind_specific_style_sprm_restrictions() {
    use crate::sprm_operations::{
        SPRM_C_CNF, SPRM_C_F_BOLD, SPRM_C_ISTD, SPRM_P_F_IN_TABLE, SPRM_P_F_KEEP, SPRM_P_ILFO,
    };

    let bold = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
    assert!(validate_character_style_sprms(&bold, false).is_ok());
    let character_style = [SPRM_C_ISTD.to_le_bytes().as_slice(), &15u16.to_le_bytes()].concat();
    assert!(validate_character_style_sprms(&character_style, false).is_err());
    let conditional_character = [
        SPRM_C_CNF.to_le_bytes().as_slice(),
        &[2],
        &1u16.to_le_bytes(),
    ]
    .concat();
    assert!(validate_character_style_sprms(&conditional_character, false).is_err());
    assert!(validate_character_style_sprms(&conditional_character, true).is_ok());

    let keep = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
    assert!(validate_paragraph_style_sprms(&keep, 15, false).is_ok());
    let table_state = [SPRM_P_F_IN_TABLE.to_le_bytes().as_slice(), &[1]].concat();
    assert!(validate_paragraph_style_sprms(&table_state, 15, true).is_err());
    let list = [SPRM_P_ILFO.to_le_bytes().as_slice(), &1u16.to_le_bytes()].concat();
    assert!(validate_numbering_style_sprms(&list, 15).is_ok());
    assert!(validate_numbering_style_sprms(&keep, 15).is_err());

    let forbidden_table_position =
        [0x9601u16.to_le_bytes().as_slice(), &0i16.to_le_bytes()].concat();
    assert!(validate_table_style_sprms(&forbidden_table_position, 15, false).is_err());
    let width_before = [0xF617u16.to_le_bytes().as_slice(), &[3, 0, 0]].concat();
    assert!(validate_table_style_sprms(&width_before, 15, false).is_err());
    assert!(validate_table_style_sprms(&width_before, 11, false).is_ok());

    let border = [0xD47Fu16.to_le_bytes().as_slice(), &[8], &[0; 8]].concat();
    assert!(validate_table_style_sprms(&border, 15, false).is_err());
    let conditional_table = [
        0xD66Au16.to_le_bytes().as_slice(),
        &[(border.len() + 2) as u8],
        &1u16.to_le_bytes(),
        border.as_slice(),
    ]
    .concat();
    assert!(validate_table_style_sprms(&conditional_table, 15, false).is_ok());
    assert!(validate_table_style_sprms(&conditional_table, 15, true).is_err());
}

#[test]
fn rejects_malformed_record_framing() {
    let valid = valid_stylesheet();
    assert!(parse(&valid[..valid.len() - 1]).is_err());

    let mut short_header = valid.clone();
    short_header[0..2].copy_from_slice(&16u16.to_le_bytes());
    assert!(parse(&short_header).is_err());

    let mut negative_std = valid.clone();
    negative_std[20..22].copy_from_slice(&0x8000u16.to_le_bytes());
    assert!(parse(&negative_std).is_err());

    let mut wrong_bch = valid.clone();
    wrong_bch[28..30].copy_from_slice(&0u16.to_le_bytes());
    assert!(parse(&wrong_bch).is_err());

    let mut trailing = valid;
    trailing.push(0);
    assert!(parse(&trailing).is_err());
}

#[test]
fn rejects_invalid_semantics_and_padding() {
    let mut slots = vec![None; 15];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Same", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Same", &[&[]]));
    assert!(parse(&stylesheet(slots)).is_err());

    let mut slots = vec![None; 15];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    slots[11] = Some(std_record(105, 3, 11, 0, "Table", &[&[], &[], &[]]));
    assert!(parse(&stylesheet(slots)).is_err());

    let mut slots = vec![None; 15];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    let mut table = std_record(105, 3, 0, 0, "Table", &[&[1], &[], &[]]);
    let padding = table
        .windows(4)
        .rposition(|bytes| bytes == [1, 0, 1, 0])
        .unwrap()
        + 3;
    table[padding] = 1;
    slots[11] = Some(table);
    assert!(parse(&stylesheet(slots)).is_err());
}

#[test]
fn resolves_table_style_inheritance_and_default_fallback() {
    let mut slots = vec![None; 17];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    slots[11] = Some(std_record(
        105,
        3,
        NIL_STYLE,
        0,
        "Normal Table",
        &[&[0x00, 0x54, 0x01, 0x00], &[], &[]],
    ));
    slots[15] = Some(std_record(
        0x0FFE,
        3,
        11,
        0,
        "Base Table",
        &[&[0x7D, 0x34, 0x01], &[], &[]],
    ));
    slots[16] = Some(std_record(
        0x0FFE,
        3,
        15,
        0,
        "Derived Table",
        &[&[0x00, 0x54, 0x02, 0x00], &[], &[]],
    ));
    let parsed = parse(&stylesheet(slots)).unwrap();

    let (effective, properties) = parsed.resolve_table_properties(16).unwrap();
    assert_eq!(effective, 16);
    assert_eq!(
        properties.justification,
        super::super::tap::TableJustification::Right
    );
    assert_eq!(properties.style_defaults.no_wrap, Some(true));

    let (effective, fallback) = parsed.resolve_table_properties(999).unwrap();
    assert_eq!(effective, 11);
    assert_eq!(
        fallback.justification,
        super::super::tap::TableJustification::Center
    );
}

#[test]
fn resolves_table_text_style_inheritance_conditions_and_fallback() {
    use crate::sprm_operations::{
        SPRM_C_CNF, SPRM_C_F_BOLD, SPRM_C_F_ITALIC, SPRM_P_CNF, SPRM_P_F_KEEP, SPRM_P_F_KEEP_FOLLOW,
    };

    fn append(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(operand);
    }

    fn conditional(opcode: u16, condition: u16, nested: &[u8]) -> Vec<u8> {
        let mut grpprl = opcode.to_le_bytes().to_vec();
        grpprl.push((nested.len() + 2) as u8);
        grpprl.extend_from_slice(&condition.to_le_bytes());
        grpprl.extend_from_slice(nested);
        grpprl
    }

    let mut normal_papx = Vec::new();
    append(&mut normal_papx, SPRM_P_F_KEEP, &[1]);
    let normal_conditional = [SPRM_P_F_KEEP_FOLLOW.to_le_bytes().as_slice(), &[1]].concat();
    normal_papx.extend_from_slice(&conditional(SPRM_P_CNF, 0x0001, &normal_conditional));
    let mut normal_chpx = Vec::new();
    append(&mut normal_chpx, SPRM_C_F_BOLD, &[1]);
    let normal_character_conditional = [SPRM_C_F_ITALIC.to_le_bytes().as_slice(), &[1]].concat();
    normal_chpx.extend_from_slice(&conditional(
        SPRM_C_CNF,
        0x0001,
        &normal_character_conditional,
    ));

    let mut derived_papx = 15u16.to_le_bytes().to_vec();
    append(&mut derived_papx, SPRM_P_F_KEEP, &[0]);
    let derived_conditional = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
    derived_papx.extend_from_slice(&conditional(SPRM_P_CNF, 0x0008, &derived_conditional));
    let mut derived_chpx = Vec::new();
    append(&mut derived_chpx, SPRM_C_F_BOLD, &[0]);
    let derived_character_conditional = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
    derived_chpx.extend_from_slice(&conditional(
        SPRM_C_CNF,
        0x0008,
        &derived_character_conditional,
    ));

    let mut slots = vec![None; 16];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    slots[11] = Some(std_record(
        105,
        3,
        NIL_STYLE,
        0,
        "Normal Table",
        &[&[], &normal_papx, &normal_chpx],
    ));
    slots[15] = Some(std_record(
        0x0FFE,
        3,
        11,
        0,
        "Derived Table",
        &[&[], &derived_papx, &derived_chpx],
    ));
    let stylesheet = parse(&stylesheet(slots)).unwrap();

    let (effective, paragraph, character) = stylesheet.resolve_table_text_properties(15).unwrap();
    assert_eq!(effective, 15);
    assert!(!paragraph.keep_on_page);
    assert_eq!(paragraph.conditional_formats.len(), 2);
    assert_eq!(
        paragraph.conditional_formats[0].condition,
        super::super::tap::TableStyleCondition::HeaderRow
    );
    assert!(paragraph.conditional_formats[0].properties.keep_with_next);
    assert_eq!(
        paragraph.conditional_formats[1].condition,
        super::super::tap::TableStyleCondition::LastColumn
    );
    assert!(paragraph.conditional_formats[1].properties.keep_on_page);
    assert_eq!(character.is_bold, Some(false));
    assert_eq!(character.conditional_formats.len(), 2);
    assert_eq!(
        character.conditional_formats[0].condition,
        super::super::tap::TableStyleCondition::HeaderRow
    );
    assert_eq!(
        character.conditional_formats[0].properties.is_italic,
        Some(true)
    );
    assert_eq!(
        character.conditional_formats[1].condition,
        super::super::tap::TableStyleCondition::LastColumn
    );
    assert_eq!(
        character.conditional_formats[1].properties.is_bold,
        Some(true)
    );

    let (_, table_papx, _) = stylesheet.resolve_table_text_style_sprms(15).unwrap();
    let direct_papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
    let cascaded = super::super::pap::ParagraphProperties::cascade_table_style(
        &table_papx,
        Some(0),
        &direct_papx,
        &stylesheet,
    )
    .unwrap();
    assert!(cascaded.keep_on_page);
    assert_eq!(cascaded.style_index, Some(0));
    assert_eq!(cascaded.conditional_formats.len(), 2);

    let (effective, paragraph, character) = stylesheet.resolve_table_text_properties(999).unwrap();
    assert_eq!(effective, 11);
    assert!(paragraph.keep_on_page);
    assert_eq!(paragraph.conditional_formats.len(), 1);
    assert_eq!(character.is_bold, Some(true));
    assert_eq!(character.conditional_formats.len(), 1);
}

#[test]
fn rejects_malformed_table_text_style_property_sets_when_resolved() {
    let mut slots = vec![None; 16];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    slots[11] = Some(std_record(
        105,
        3,
        NIL_STYLE,
        0,
        "Normal Table",
        &[&[], &[], &[]],
    ));
    slots[15] = Some(std_record(
        0x0FFE,
        3,
        11,
        0,
        "Malformed Table",
        &[&[], &[0x35, 0x08, 1], &[]],
    ));
    let stylesheet = parse(&stylesheet(slots)).unwrap();
    assert!(stylesheet.resolve_table_text_properties(15).is_err());
}

#[test]
fn flattens_table_text_conditions_in_position_precedence_order() {
    use crate::sprm_operations::{SPRM_C_CNF, SPRM_C_F_BOLD};

    fn conditional(condition: u16, value: u8) -> Vec<u8> {
        let nested = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[value]].concat();
        let mut grpprl = SPRM_C_CNF.to_le_bytes().to_vec();
        grpprl.push((nested.len() + 2) as u8);
        grpprl.extend_from_slice(&condition.to_le_bytes());
        grpprl.extend_from_slice(&nested);
        grpprl
    }

    // Source order is deliberately the reverse of positional precedence.
    let source = [conditional(0x0001, 0), conditional(0x0040, 1)].concat();
    let flattened = flatten_conditional_style_sprms(
        &source,
        SPRM_C_CNF,
        &[
            super::super::tap::TableStyleCondition::OddRowBand,
            super::super::tap::TableStyleCondition::HeaderRow,
        ],
    )
    .unwrap();
    let properties = super::super::chp::CharacterProperties::from_sprm(&flattened).unwrap();
    assert_eq!(properties.is_bold, Some(false));
    assert!(properties.conditional_formats.is_empty());
}

#[test]
fn applies_table_styles_in_direct_sprm_order_and_preserves_sizing() {
    fn append(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
        grpprl.extend_from_slice(&opcode.to_le_bytes());
        grpprl.extend_from_slice(operand);
    }

    let mut slots = vec![None; 16];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    slots[11] = Some(std_record(
        105,
        3,
        NIL_STYLE,
        0,
        "Normal Table",
        &[&[], &[], &[]],
    ));
    let mut style_tapx = Vec::new();
    append(&mut style_tapx, 0x548A, &[2, 0]);
    append(&mut style_tapx, 0x3404, &[1]);
    append(&mut style_tapx, 0x3403, &[1]);
    append(&mut style_tapx, 0x347D, &[1]);
    slots[15] = Some(std_record(
        0x0FFE,
        3,
        11,
        0,
        "Applied Table",
        &[&style_tapx, &[], &[]],
    ));
    let stylesheet = parse(&stylesheet(slots)).unwrap();

    let mut direct = Vec::new();
    append(&mut direct, 0x548A, &[1, 0]);
    append(&mut direct, 0x3404, &[0]);
    append(&mut direct, 0x3403, &[0]);
    append(&mut direct, 0x9407, &[0x20, 0x03]);
    append(&mut direct, 0xF614, &[3, 0xE8, 0x03]);
    append(&mut direct, 0x3615, &[1]);
    append(&mut direct, 0x5664, &[1, 0]);
    append(&mut direct, 0x7479, &[0x78, 0x56, 0x34, 0x12]);
    append(&mut direct, 0x563A, &[15, 0]);

    let arena = bumpalo::Bump::new();
    let parser = super::super::tap_parser::TapParser::new(&arena);
    let styled = parser
        .parse_tap_with_stylesheet(&direct, &stylesheet)
        .unwrap();
    assert_eq!(styled.table_style_index, Some(15));
    assert_eq!(
        styled.justification,
        super::super::tap::TableJustification::Right
    );
    assert!(styled.is_header_row);
    assert!(!styled.allow_row_break);
    assert_eq!(styled.style_defaults.no_wrap, Some(true));
    assert_eq!(styled.row_height, Some(800));
    assert_eq!(styled.preferred_width.unwrap().value, 1000);
    assert!(styled.auto_fit);
    assert!(styled.right_to_left);
    assert_eq!(styled.revision_save_id, Some(0x1234_5678));

    append(&mut direct, 0x548A, &[0, 0]);
    append(&mut direct, 0x3404, &[0]);
    append(&mut direct, 0x3403, &[0]);
    let overridden = parser
        .parse_tap_with_stylesheet(&direct, &stylesheet)
        .unwrap();
    assert_eq!(
        overridden.justification,
        super::super::tap::TableJustification::Left
    );
    assert!(!overridden.is_header_row);
    assert!(overridden.allow_row_break);
    assert_eq!(overridden.row_height, Some(800));
    assert_eq!(overridden.preferred_width.unwrap().value, 1000);
    assert!(overridden.auto_fit);
    assert!(overridden.right_to_left);
    assert_eq!(overridden.revision_save_id, Some(0x1234_5678));
}

#[test]
fn resolves_paragraph_and_character_style_property_arrays() {
    let mut slots = vec![None; 19];
    slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
    slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
    slots[15] = Some(std_record(
        0x0FFE,
        1,
        0,
        0,
        "Base Paragraph",
        &[&[15, 0, 0x03, 0x24, 2], &[0x35, 0x08, 1]],
    ));
    slots[16] = Some(std_record(
        0x0FFE,
        1,
        15,
        0,
        "Derived Paragraph",
        &[&[0x03, 0x24, 1], &[0x36, 0x08, 1]],
    ));
    slots[17] = Some(std_record(
        0x0FFE,
        2,
        10,
        17,
        "Base Character",
        &[&[0x35, 0x08, 1]],
    ));
    slots[18] = Some(std_record(
        0x0FFE,
        2,
        17,
        18,
        "Derived Character",
        &[&[0x36, 0x08, 1]],
    ));
    let stylesheet = parse(&stylesheet(slots)).unwrap();

    let (effective, paragraph, character) = stylesheet.resolve_paragraph_style_sprms(16).unwrap();
    assert_eq!(effective, Some(16));
    assert_eq!(paragraph, [0x03, 0x24, 2, 0x03, 0x24, 1]);
    assert_eq!(character, [0x35, 0x08, 1, 0x36, 0x08, 1]);
    let styled = super::super::pap::ParagraphProperties::from_sprm(&paragraph).unwrap();
    assert_eq!(
        styled.justification,
        super::super::pap::Justification::Center
    );
    let mut direct = paragraph.clone();
    direct.extend_from_slice(&[0x03, 0x24, 0]);
    let overridden = super::super::pap::ParagraphProperties::from_sprm(&direct).unwrap();
    assert_eq!(
        overridden.justification,
        super::super::pap::Justification::Left
    );

    let (effective, character) = stylesheet.resolve_character_style_sprms(18).unwrap();
    assert_eq!(effective, Some(18));
    assert_eq!(character, [0x35, 0x08, 1, 0x36, 0x08, 1]);

    let direct_grpprl = [0x30, 0x4A, 18, 0, 0x35, 0x08, 0];
    let direct = super::super::chp::CharacterProperties::from_sprm(&direct_grpprl).unwrap();
    let cascaded = super::super::paragraph_extractor::cascade_character_properties(
        Some(&stylesheet),
        &[0x35, 0x08, 1, 0x36, 0x08, 0],
        &direct,
        &direct_grpprl,
    )
    .unwrap();
    assert_eq!(cascaded.style_index, Some(18));
    assert_eq!(cascaded.is_bold, Some(false));
    assert_eq!(cascaded.is_italic, Some(true));

    let ordered_grpprl = [
        0x35, 0x08, 0, // direct bold before the style: reset by sprmCIstd
        0x0C, 0x2A, 7, // highlight: explicitly preserved by sprmCIstd
        0x55, 0x08, 1, // fSpec: explicitly preserved by sprmCIstd
        0x30, 0x4A, 18, 0, // derived character style
        0x36, 0x08, 0, // direct italic after the style: authoritative
    ];
    let ordered_direct =
        super::super::chp::CharacterProperties::from_sprm(&ordered_grpprl).unwrap();
    let ordered = super::super::paragraph_extractor::cascade_character_properties(
        Some(&stylesheet),
        &[0x35, 0x08, 1],
        &ordered_direct,
        &ordered_grpprl,
    )
    .unwrap();
    assert_eq!(ordered.style_index, Some(18));
    assert_eq!(ordered.is_bold, Some(true));
    assert_eq!(ordered.is_italic, Some(false));
    assert!(ordered.is_spec);
    assert_eq!(
        ordered.highlight,
        Some(super::super::chp::HighlightColor::Yellow)
    );

    assert_eq!(
        stylesheet.resolve_paragraph_style_sprms(18).unwrap(),
        (None, Vec::new(), Vec::new())
    );
    assert_eq!(
        stylesheet.resolve_character_style_sprms(16).unwrap(),
        (None, Vec::new())
    );

    let mut ordered_papx = vec![
        0x03, 0x24, 0, // direct left alignment before style switch
        0x16, 0x24, 1, // table membership is preserved
        0x5A, 0x24, 1, // open cell-mark display state is preserved
        0x64, 0x26, 1, // paragraph revision wall is preserved
        0x65, 0x64, 7, 0, 0, 0, // PGPInfo identity is preserved
        0x67, 0x64, 0x78, 0x56, 0x34, 0x12, // paragraph RSID is preserved
        0x00, 0x46, 15, 0, // switch back to Base Paragraph (right)
    ];
    let switched = super::super::pap::ParagraphProperties::cascade_styles(
        Some(16),
        &ordered_papx,
        &stylesheet,
    )
    .unwrap();
    assert_eq!(switched.style_index, Some(15));
    assert_eq!(
        switched.justification,
        super::super::pap::Justification::Right
    );
    assert!(switched.in_table);
    assert!(switched.open_table_cell_mark);
    assert!(switched.properties_preserved_for_revision);
    assert_eq!(switched.paragraph_group_id, Some(7));
    assert_eq!(switched.revision_save_id, Some(0x1234_5678));

    ordered_papx.extend_from_slice(&[0x03, 0x24, 1]);
    let overridden = super::super::pap::ParagraphProperties::cascade_styles(
        Some(16),
        &ordered_papx,
        &stylesheet,
    )
    .unwrap();
    assert_eq!(
        overridden.justification,
        super::super::pap::Justification::Center
    );

    let permuted = super::super::pap::ParagraphProperties::cascade_styles(
        Some(16),
        &[
            0x01, 0xC6, 7, // sprmPIstdPermute and SPPOperand length
            0, 16, 0, 16, 0, 15, 0, // style 16 maps to style 15
        ],
        &stylesheet,
    )
    .unwrap();
    assert_eq!(permuted.style_index, Some(15));
    assert_eq!(
        permuted.justification,
        super::super::pap::Justification::Right
    );
}
