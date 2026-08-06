//! Structural and semantic validation for stylesheet property sets.

use std::collections::HashSet;

use super::super::codec::{corrupted, read_u16};
use super::super::model::{StyleDefinition, StyleKind};
use crate::leniency::{Leniency, StylesheetDefect, ToleranceReport};
use crate::package::Result;
use crate::sprm_operations::{get_sprm_operation, get_sprm_type};

pub(in crate::parts::styles) fn strip_paragraph_style_index(
    properties: &[u8],
    style_index: u16,
) -> Result<&[u8]> {
    if properties.len() >= 2 {
        let prefix = read_u16(properties, 0, "UpxPapx.istd")?;
        if prefix == style_index {
            return Ok(&properties[2..]);
        }
    }
    Ok(properties)
}

pub(in crate::parts::styles) fn validate_style_sprms(
    properties: &[u8],
    expected_type: u8,
    structure: &str,
) -> Result<Vec<crate::sprm::Sprm>> {
    let sprms = crate::sprm::parse_sprms(properties)?;
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != properties.len()
        || sprms
            .iter()
            .any(|sprm| get_sprm_type(sprm.opcode) != expected_type)
    {
        return Err(corrupted(&format!(
            "{structure} contains malformed or wrong-type SPRMs"
        )));
    }
    Ok(sprms)
}

pub(crate) fn validate_character_style_sprms(
    properties: &[u8],
    conditional_table_style: bool,
) -> Result<()> {
    let sprms = validate_style_sprms(properties, 2, "UpxChpx")?;
    // [MS-DOC] UpxChpx: explicit exclusions plus every property that sprmCIstd preserves.
    const FORBIDDEN: &[u16] = &[
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x09, 0x0A, 0x0C, 0x11, 0x15, 0x16, 0x17,
        0x18, 0x1A, 0x30, 0x31, 0x33, 0x47, 0x55, 0x56, 0x57, 0x5A, 0x62, 0x63, 0x64, 0x67, 0x6F,
        0x79, 0x82, 0x83, 0x86, 0x87, 0x88, 0x89, 0x90,
    ];
    if let Some(sprm) = sprms.iter().find(|sprm| {
        let operation = get_sprm_operation(sprm.opcode);
        FORBIDDEN.contains(&operation) || (!conditional_table_style && operation == 0x85)
    }) {
        return Err(corrupted(&format!(
            "UpxChpx contains disallowed style SPRM {:#06x}",
            sprm.opcode
        )));
    }
    if sprms
        .iter()
        .any(|sprm| get_sprm_operation(sprm.opcode) == 0x85)
    {
        crate::parts::chp::CharacterProperties::from_sprm(properties)?;
    }
    Ok(())
}
pub(crate) fn validate_paragraph_style_sprms(
    properties: &[u8],
    style_index: u16,
    conditional_table_style: bool,
) -> Result<()> {
    let properties = strip_paragraph_style_index(properties, style_index)?;
    let sprms = validate_style_sprms(properties, 1, "UpxPapx")?;
    // [MS-DOC] UpxPapx: explicit exclusions plus every property that sprmPIstd preserves.
    const FORBIDDEN: &[u16] = &[
        0x00, 0x01, 0x02, 0x10, 0x15, 0x16, 0x17, 0x2C, 0x3F, 0x43, 0x45, 0x46, 0x49, 0x4B, 0x4C,
        0x5A, 0x5F, 0x62, 0x64, 0x65, 0x67, 0x69, 0x6B, 0x6C, 0x6F,
    ];
    if let Some(sprm) = sprms.iter().find(|sprm| {
        let operation = get_sprm_operation(sprm.opcode);
        FORBIDDEN.contains(&operation) || (!conditional_table_style && operation == 0x66)
    }) {
        return Err(corrupted(&format!(
            "UpxPapx contains disallowed style SPRM {:#06x}",
            sprm.opcode
        )));
    }
    if sprms
        .iter()
        .any(|sprm| get_sprm_operation(sprm.opcode) == 0x66)
    {
        crate::parts::pap::ParagraphProperties::from_sprm(properties)?;
    }
    Ok(())
}

pub(crate) fn validate_numbering_style_sprms(properties: &[u8], style_index: u16) -> Result<()> {
    let properties = strip_paragraph_style_index(properties, style_index)?;
    let sprms = validate_style_sprms(properties, 1, "numbering-style UpxPapx")?;
    if sprms
        .iter()
        .any(|sprm| sprm.opcode != crate::sprm_operations::SPRM_P_ILFO)
    {
        return Err(corrupted(
            "numbering-style UpxPapx contains an SPRM other than sprmPIlfo",
        ));
    }
    Ok(())
}

pub(crate) fn validate_table_style_sprms(
    properties: &[u8],
    style_index: u16,
    inside_conditional: bool,
) -> Result<()> {
    let sprms = crate::sprm::parse_sprms(properties)?;
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != properties.len() || sprms.iter().any(|sprm| get_sprm_type(sprm.opcode) != 5) {
        return Err(corrupted("UpxTapx contains malformed or wrong-type SPRMs"));
    }
    // [MS-DOC] UpxTapx: explicit exclusions plus every property that sprmTIstd preserves.
    const FORBIDDEN: &[u16] = &[
        0x01, 0x02, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F, 0x10, 0x11, 0x12, 0x14,
        0x15, 0x16, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E, 0x1F, 0x20, 0x21, 0x22, 0x23, 0x24,
        0x25, 0x29, 0x2B, 0x2C, 0x2F, 0x32, 0x35, 0x36, 0x39, 0x42, 0x60, 0x62, 0x64, 0x65, 0x67,
        0x68, 0x69, 0x70, 0x71, 0x72, 0x79,
    ];
    let has_conditional = sprms
        .iter()
        .any(|sprm| get_sprm_operation(sprm.opcode) == 0x6A);
    for sprm in sprms {
        let operation = get_sprm_operation(sprm.opcode);
        let conditional_border = (0x7F..=0x84).contains(&operation);
        if FORBIDDEN.contains(&operation) || (conditional_border && !inside_conditional) {
            return Err(corrupted("UpxTapx contains a disallowed style SPRM"));
        }
        if operation == 0x17
            && (inside_conditional || style_index != 11 || sprm.operand_bytes() != [3, 0, 0])
        {
            return Err(corrupted(
                "sprmTWidthBefore is invalid for this table style",
            ));
        }
        if operation == 0x6A {
            if inside_conditional {
                return Err(corrupted("sprmTCnf cannot be nested recursively"));
            }
            let operand = sprm.operand_bytes();
            let nested = operand
                .get(2..)
                .ok_or_else(|| corrupted("sprmTCnf operand is truncated"))?;
            validate_table_style_sprms(nested, style_index, true)?;
        }
    }
    if has_conditional {
        let arena = bumpalo::Bump::new();
        crate::parts::tap_parser::TapParser::new(&arena).parse_tap(properties)?;
    }
    Ok(())
}

pub(in crate::parts::styles) fn validate_styles(
    styles: &[Option<StyleDefinition>],
    leniency: Leniency,
    tolerance: &mut ToleranceReport,
) -> Result<()> {
    for required_empty in [13usize, 14] {
        if styles.get(required_empty).is_some_and(Option::is_some) {
            return Err(corrupted("reserved fixed-index style is not empty"));
        }
    }
    const FIXED_IDS: [u16; 13] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 65, 105, 107];
    for (index, expected_id) in FIXED_IDS.into_iter().enumerate() {
        if let Some(style) = styles.get(index).and_then(Option::as_ref) {
            if style.invariant_id != expected_id {
                return Err(corrupted("fixed-index style has the wrong invariant ID"));
            }
            let expected_kind = match index {
                0..=9 => StyleKind::Paragraph,
                10 => StyleKind::Character,
                11 => StyleKind::Table,
                12 => StyleKind::Numbering,
                _ => unreachable!(),
            };
            if style.kind != expected_kind {
                return Err(corrupted("fixed-index style has the wrong kind"));
            }
        }
    }

    let mut names = HashSet::new();
    for style in styles.iter().flatten() {
        for name in std::iter::once(&style.name).chain(style.aliases.iter()) {
            if !names.insert(name.as_str()) {
                // MS-DOC 2.9 requires uniqueness, but the name only labels the
                // style: every stored reference resolves by index, so a
                // duplicate cannot make a lookup ambiguous. Rejecting costs the
                // caller the whole document.
                if !leniency.tolerates_stylesheet_defects() {
                    return Err(corrupted("style names and aliases must be unique"));
                }
                tolerance.record(StylesheetDefect::DuplicateStyleName, style.index);
            }
        }
        if let Some(base) = style.base_style
            && (base == style.index || styles.get(usize::from(base)).is_none_or(Option::is_none))
        {
            return Err(corrupted("style has an invalid base style"));
        }
        if styles
            .get(usize::from(style.next_style))
            .is_none_or(Option::is_none)
        {
            return Err(corrupted("style has an invalid next style"));
        }
        if let Some(linked) = style.post_2000.as_ref().and_then(|post| post.linked_style)
            && styles.get(usize::from(linked)).is_none_or(Option::is_none)
        {
            return Err(corrupted("style has an invalid linked style"));
        }
    }

    for style in styles.iter().flatten() {
        let mut visited = HashSet::new();
        let mut current = Some(style.index);
        while let Some(index) = current {
            if !visited.insert(index) {
                return Err(corrupted("style inheritance contains a cycle"));
            }
            current = styles
                .get(usize::from(index))
                .and_then(Option::as_ref)
                .and_then(|definition| definition.base_style);
        }
    }
    Ok(())
}
