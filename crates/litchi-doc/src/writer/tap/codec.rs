//! DOC TAP binary encoding.
//!
//! This layer lowers the semantic table model into Word SPRMs. It does not
//! own the public table model or validation error definitions.

use crate::parts::tap::{
    BorderStyle, BorderType, CellShading, CellSpacing, CellSpacingSource,
    TableConditionalFormatting, TableHorizontalAnchor, TableHorizontalPosition, TableJustification,
    TablePositioning, TableStyleBorder, TableStyleDefaults, TableStyleShading, TableVerticalAnchor,
    TableVerticalPosition, TableWidth, TextDirection, VerticalAlignment, VerticalMergeStatus,
    WidthType,
};
use crate::sprm_operations::{
    SPRM_T_C_HORZ_BANDS, SPRM_T_C_VERT_BANDS, SPRM_T_CELL_BRC_BOTTOM_STYLE,
    SPRM_T_CELL_BRC_INSIDE_H_STYLE, SPRM_T_CELL_BRC_INSIDE_V_STYLE, SPRM_T_CELL_BRC_LEFT_STYLE,
    SPRM_T_CELL_BRC_RIGHT_STYLE, SPRM_T_CELL_BRC_TL2BR_STYLE, SPRM_T_CELL_BRC_TOP_STYLE,
    SPRM_T_CELL_BRC_TR2BL_STYLE, SPRM_T_CELL_NO_WRAP_STYLE, SPRM_T_CELL_PADDING_STYLE,
    SPRM_T_CELL_SHD_STYLE, SPRM_T_CELL_VERT_ALIGN_STYLE, SPRM_T_CNF, SPRM_T_DXA_ABS,
    SPRM_T_DXA_FROM_TEXT, SPRM_T_DXA_FROM_TEXT_RIGHT, SPRM_T_DYA_ABS, SPRM_T_DYA_FROM_TEXT,
    SPRM_T_DYA_FROM_TEXT_BOTTOM, SPRM_T_F_AUTOFIT, SPRM_T_F_BI_DI, SPRM_T_F_BI_DI90,
    SPRM_T_F_CANT_SPLIT, SPRM_T_F_CANT_SPLIT90, SPRM_T_F_KEEP_FOLLOW, SPRM_T_F_NO_ALLOW_OVERLAP,
    SPRM_T_IPGP, SPRM_T_ISTD, SPRM_T_JC, SPRM_T_JC90, SPRM_T_PC, SPRM_T_PROP_RMARK, SPRM_T_RSID,
    SPRM_T_TABLE_HEADER, SPRM_T_TABLE_WIDTH, SPRM_T_TLP, SPRM_T_WALL, SPRM_T_WIDTH_AFTER,
    SPRM_T_WIDTH_BEFORE, SPRM_T_WIDTH_INDENT,
};
use crate::writer::sprm::SprmBuilder;

use super::model::{TableBorders, TableCell, TableRow};
use super::validation::TapBuildError;

#[derive(Debug, Clone, Copy)]
enum PreferredWidthUsage {
    Table,
    TablePart,
    Indent,
}

fn encode_preferred_width(
    property: &'static str,
    width: Option<TableWidth>,
    usage: PreferredWidthUsage,
) -> Result<Option<[u8; 3]>, TapBuildError> {
    let Some(width) = width else {
        return Ok(None);
    };
    let units = match width.width_type {
        WidthType::Auto if width.value == 0 => 1,
        WidthType::Percentage
            if matches!(usage, PreferredWidthUsage::Table)
                && (0..=30_000).contains(&width.value) =>
        {
            2
        },
        WidthType::Percentage
            if matches!(usage, PreferredWidthUsage::TablePart)
                && (0..=5_000).contains(&width.value) =>
        {
            2
        },
        WidthType::Twips
            if matches!(
                usage,
                PreferredWidthUsage::Table | PreferredWidthUsage::TablePart
            ) && (0..=31_680).contains(&width.value) =>
        {
            3
        },
        WidthType::Twips
            if matches!(usage, PreferredWidthUsage::Indent)
                && (-31_560..=31_680).contains(&width.value) =>
        {
            3
        },
        _ => return Err(TapBuildError::InvalidPreferredWidth(property, width)),
    };
    let value = width.value.to_le_bytes();
    Ok(Some([units, value[0], value[1]]))
}

fn encode_horizontal_position(position: TableHorizontalPosition) -> Result<i16, TapBuildError> {
    Ok(match position {
        TableHorizontalPosition::Left => 0,
        TableHorizontalPosition::Center => -4,
        TableHorizontalPosition::Right => -8,
        TableHorizontalPosition::Inside => -12,
        TableHorizontalPosition::Outside => -16,
        TableHorizontalPosition::Offset(value) => {
            let stored = i32::from(value) + 1;
            if !(-31_679..=31_681).contains(&i32::from(value))
                || matches!(stored, 0 | -4 | -8 | -12 | -16)
            {
                return Err(TapBuildError::InvalidTablePosition("horizontal", value));
            }
            stored as i16
        },
    })
}

fn encode_vertical_position(position: TableVerticalPosition) -> Result<i16, TapBuildError> {
    Ok(match position {
        TableVerticalPosition::Inline => 0,
        TableVerticalPosition::Top => -4,
        TableVerticalPosition::Center => -8,
        TableVerticalPosition::Bottom => -12,
        TableVerticalPosition::Inside => -16,
        TableVerticalPosition::Outside => -20,
        TableVerticalPosition::Offset(value) => {
            let stored = i32::from(value) + 1;
            if !(-31_679..=31_681).contains(&i32::from(value))
                || matches!(stored, 0 | -4 | -8 | -12 | -16 | -20)
            {
                return Err(TapBuildError::InvalidTablePosition("vertical", value));
            }
            stored as i16
        },
    })
}

fn encode_positioning(positioning: TablePositioning) -> u8 {
    let vertical = match positioning.vertical_anchor {
        TableVerticalAnchor::Margin => 0,
        TableVerticalAnchor::Page => 1,
        TableVerticalAnchor::Paragraph => 2,
        TableVerticalAnchor::None => 3,
    };
    let horizontal = match positioning.horizontal_anchor {
        TableHorizontalAnchor::Column => 0,
        TableHorizontalAnchor::Margin => 1,
        TableHorizontalAnchor::Page => 2,
        TableHorizontalAnchor::None => 3,
    };
    (vertical << 4) | (horizontal << 6)
}

fn justification_code(justification: TableJustification) -> u16 {
    match justification {
        TableJustification::Left => 0,
        TableJustification::Center => 1,
        TableJustification::Right => 2,
    }
}

fn physical_justification(logical: TableJustification, right_to_left: bool) -> TableJustification {
    if right_to_left {
        match logical {
            TableJustification::Left => TableJustification::Right,
            TableJustification::Center => TableJustification::Center,
            TableJustification::Right => TableJustification::Left,
        }
    } else {
        logical
    }
}

/// Serialize non-conditional table-style defaults for an `UpxTapx`.
pub fn generate_table_style_sprms(defaults: &TableStyleDefaults) -> Result<Vec<u8>, TapBuildError> {
    generate_table_style_properties(defaults, false)
}

/// Serialize table-style defaults and their conditional `sprmTCnf` entries.
pub fn generate_table_style_sprms_with_conditionals(
    defaults: &TableStyleDefaults,
    conditional_formats: &[TableConditionalFormatting],
) -> Result<Vec<u8>, TapBuildError> {
    let mut sprms = generate_table_style_properties(defaults, false)?;
    for conditional in conditional_formats {
        if conditional.raw_grpprl.len() > 253 {
            return Err(TapBuildError::ConditionalPropertiesTooLong(
                conditional.raw_grpprl.len(),
            ));
        }
        let nested = if conditional.raw_grpprl.is_empty() {
            generate_table_style_properties(&conditional.properties, true)?
        } else {
            let arena = bumpalo::Bump::new();
            crate::parts::tap_parser::TapParser::new(&arena)
                .parse_conditional_tap(&conditional.raw_grpprl)
                .map_err(|error| TapBuildError::InvalidConditionalProperties(error.to_string()))?;
            conditional.raw_grpprl.clone()
        };
        if nested.len() > 253 {
            return Err(TapBuildError::ConditionalPropertiesTooLong(nested.len()));
        }
        sprms.extend_from_slice(&SPRM_T_CNF.to_le_bytes());
        sprms.push((nested.len() + 2) as u8);
        sprms.extend_from_slice(&conditional.condition.code().to_le_bytes());
        sprms.extend_from_slice(&nested);
    }
    Ok(sprms)
}

fn generate_table_style_properties(
    defaults: &TableStyleDefaults,
    inside_conditional: bool,
) -> Result<Vec<u8>, TapBuildError> {
    let has_style_border = [
        defaults.border_top,
        defaults.border_bottom,
        defaults.border_left,
        defaults.border_right,
        defaults.border_inside_horizontal,
        defaults.border_inside_vertical,
        defaults.border_diagonal_down,
        defaults.border_diagonal_up,
    ]
    .iter()
    .any(Option::is_some);
    if has_style_border && !inside_conditional {
        return Err(TapBuildError::StyleBorderOutsideConditional);
    }
    for (axis, size) in [
        ("horizontal", defaults.horizontal_band_size),
        ("vertical", defaults.vertical_band_size),
    ] {
        if let Some(size) = size
            && !(1..=3).contains(&size)
        {
            return Err(TapBuildError::InvalidStyleBandSize(axis, size));
        }
    }

    let mut padding_groups = Vec::<(u16, u8)>::with_capacity(4);
    for (mask, padding) in [
        (0x01, defaults.padding_top),
        (0x02, defaults.padding_left),
        (0x04, defaults.padding_bottom),
        (0x08, defaults.padding_right),
    ] {
        let Some(padding) = padding else {
            continue;
        };
        if padding > 31_680 {
            return Err(TapBuildError::InvalidCellPadding(padding));
        }
        if let Some((_, sides)) = padding_groups
            .iter_mut()
            .find(|(width, _)| *width == padding)
        {
            *sides |= mask;
        } else {
            padding_groups.push((padding, mask));
        }
    }

    let mut sprms = Vec::with_capacity(padding_groups.len() * 9 + 12);
    for (width, sides) in padding_groups {
        sprms.extend_from_slice(&SPRM_T_CELL_PADDING_STYLE.to_le_bytes());
        sprms.push(6);
        sprms.extend_from_slice(&[0, 1, sides, 3]);
        sprms.extend_from_slice(&width.to_le_bytes());
    }
    if let Some(alignment) = defaults.vertical_alignment {
        let value = match alignment {
            VerticalAlignment::Top => 0,
            VerticalAlignment::Center => 1,
            VerticalAlignment::Bottom => 2,
        };
        sprms.extend_from_slice(&SPRM_T_CELL_VERT_ALIGN_STYLE.to_le_bytes());
        sprms.push(value);
    }
    if let Some(no_wrap) = defaults.no_wrap {
        sprms.extend_from_slice(&SPRM_T_CELL_NO_WRAP_STYLE.to_le_bytes());
        sprms.push(u8::from(no_wrap));
    }
    for (opcode, border) in [
        (SPRM_T_CELL_BRC_TOP_STYLE, defaults.border_top),
        (SPRM_T_CELL_BRC_BOTTOM_STYLE, defaults.border_bottom),
        (SPRM_T_CELL_BRC_LEFT_STYLE, defaults.border_left),
        (SPRM_T_CELL_BRC_RIGHT_STYLE, defaults.border_right),
        (
            SPRM_T_CELL_BRC_INSIDE_H_STYLE,
            defaults.border_inside_horizontal,
        ),
        (
            SPRM_T_CELL_BRC_INSIDE_V_STYLE,
            defaults.border_inside_vertical,
        ),
        (SPRM_T_CELL_BRC_TL2BR_STYLE, defaults.border_diagonal_down),
        (SPRM_T_CELL_BRC_TR2BL_STYLE, defaults.border_diagonal_up),
    ] {
        let Some(border) = border else {
            continue;
        };
        sprms.extend_from_slice(&opcode.to_le_bytes());
        sprms.push(8);
        append_full_border(
            &mut sprms,
            match border {
                TableStyleBorder::NoBorder => None,
                TableStyleBorder::Border(border) => Some(border),
            },
            false,
        )?;
    }
    if let Some(shading) = defaults.shading {
        sprms.extend_from_slice(&SPRM_T_CELL_SHD_STYLE.to_le_bytes());
        sprms.push(10);
        match shading {
            TableStyleShading::NoShading => append_shading(&mut sprms, None, true),
            TableStyleShading::Shading(shading) => {
                append_shading(&mut sprms, Some(shading), false);
            },
        }
    }
    if let Some(size) = defaults.horizontal_band_size {
        sprms.extend_from_slice(&SPRM_T_C_HORZ_BANDS.to_le_bytes());
        sprms.push(size);
    }
    if let Some(size) = defaults.vertical_band_size {
        sprms.extend_from_slice(&SPRM_T_C_VERT_BANDS.to_le_bytes());
        sprms.push(size);
    }
    Ok(sprms)
}

pub(crate) fn generate_row_sprms(row: &TableRow) -> Result<Vec<u8>, TapBuildError> {
    if let Some(previous) = &row.preserved_properties_for_revision {
        if previous.preserved_properties_for_revision.is_some()
            || previous.properties_preserved_for_revision
        {
            return Err(TapBuildError::NestedPreservedState);
        }
        let mut sprms = generate_current_row_sprms(previous)?;
        sprms.extend_from_slice(&0x3668u16.to_le_bytes());
        sprms.push(1);
        let mut current = row.clone();
        current.properties_preserved_for_revision = false;
        current.preserved_properties_for_revision = None;
        sprms.extend_from_slice(&generate_current_row_sprms(&current)?);
        return Ok(sprms);
    }
    generate_current_row_sprms(row)
}

fn generate_current_row_sprms(row: &TableRow) -> Result<Vec<u8>, TapBuildError> {
    let cell_count = row.cells.len();
    if !(1..=63).contains(&cell_count) {
        return Err(TapBuildError::InvalidCellCount(cell_count));
    }
    if row.cells.first().is_some_and(|cell| cell.merged) {
        return Err(TapBuildError::MergeWithoutPrecedingCell);
    }
    if !(-31_680..=31_680).contains(&row.height) {
        return Err(TapBuildError::InvalidRowHeight(row.height));
    }
    let effective_widths = if row.cells.iter().all(|cell| cell.width == 0) {
        const DEFAULT_TABLE_WIDTH: u32 = 8640;
        (0..cell_count)
            .map(|index| {
                let left = DEFAULT_TABLE_WIDTH * index as u32 / cell_count as u32;
                let right = DEFAULT_TABLE_WIDTH * (index + 1) as u32 / cell_count as u32;
                (right - left) as u16
            })
            .collect::<Vec<_>>()
    } else {
        row.cells.iter().map(|cell| cell.width).collect()
    };

    let mut boundaries = Vec::with_capacity(cell_count + 1);
    boundaries.push(0i16);
    let mut boundary = 0u32;
    for width in &effective_widths {
        boundary = boundary
            .checked_add(u32::from(*width))
            .ok_or(TapBuildError::CellWidthsOverflow)?;
        if boundary > 31_680 {
            return Err(TapBuildError::CellWidthsOverflow);
        }
        boundaries.push(boundary as i16);
    }

    let preferred_width = encode_preferred_width(
        "table width",
        row.preferred_width,
        PreferredWidthUsage::Table,
    )?;
    let width_before = encode_preferred_width(
        "leading table-part width",
        row.width_before,
        PreferredWidthUsage::TablePart,
    )?;
    let width_after = encode_preferred_width(
        "trailing table-part width",
        row.width_after,
        PreferredWidthUsage::TablePart,
    )?;
    let preferred_indent = encode_preferred_width(
        "table indent",
        row.preferred_indent,
        PreferredWidthUsage::Indent,
    )?;
    if let Some(TableWidth {
        value: indent,
        width_type: WidthType::Twips,
    }) = row.preferred_indent
    {
        let layout_width = match row.preferred_width {
            Some(TableWidth {
                value,
                width_type: WidthType::Twips,
            }) => i32::from(value),
            _ => boundary as i32,
        };
        if i32::from(indent) + layout_width > 31_680 {
            return Err(TapBuildError::InvalidPreferredWidth(
                "table indent",
                TableWidth {
                    value: indent,
                    width_type: WidthType::Twips,
                },
            ));
        }
    }
    for (side, distance) in [
        ("left", row.distance_from_text_left),
        ("top", row.distance_from_text_top),
        ("right", row.distance_from_text_right),
        ("bottom", row.distance_from_text_bottom),
    ] {
        if distance > 31_680 {
            return Err(TapBuildError::InvalidWrapDistance(side, distance));
        }
    }
    let horizontal_position = encode_horizontal_position(row.horizontal_position)?;
    let vertical_position = encode_vertical_position(row.vertical_position)?;
    if row.paragraph_group_id == Some(0) {
        return Err(TapBuildError::InvalidParagraphGroupId);
    }
    if let Some(revision) = row.formatting_revision {
        if revision.author_index > i16::MAX as u16 {
            return Err(TapBuildError::InvalidRevisionAuthorIndex(
                revision.author_index,
            ));
        }
        if crate::revision::decode_dttm(revision.timestamp).is_err() {
            return Err(TapBuildError::InvalidRevisionTimestamp(revision.timestamp));
        }
    }

    let mut builder = SprmBuilder::new();
    // Apply the style first so later SPRMs remain direct row formatting.
    if let Some(style_index) = row.table_style_index {
        builder.add_word(SPRM_T_ISTD, style_index);
    }
    if row.justification != TableJustification::Left || row.right_to_left {
        builder.add_word(
            SPRM_T_JC90,
            justification_code(physical_justification(row.justification, row.right_to_left)),
        );
        builder.add_word(SPRM_T_JC, justification_code(row.justification));
    }
    if let Some(positioning) = row.positioning {
        builder.add_byte(SPRM_T_PC, encode_positioning(positioning));
    }
    if row.horizontal_position != TableHorizontalPosition::Left {
        builder.add_signed_word(SPRM_T_DXA_ABS, horizontal_position);
    }
    if row.vertical_position != TableVerticalPosition::Inline {
        builder.add_signed_word(SPRM_T_DYA_ABS, vertical_position);
    }
    if row.distance_from_text_left != 0 {
        builder.add_word(SPRM_T_DXA_FROM_TEXT, row.distance_from_text_left);
    }
    if row.distance_from_text_top != 0 {
        builder.add_word(SPRM_T_DYA_FROM_TEXT, row.distance_from_text_top);
    }
    if row.distance_from_text_right != 0 {
        builder.add_word(SPRM_T_DXA_FROM_TEXT_RIGHT, row.distance_from_text_right);
    }
    if row.distance_from_text_bottom != 0 {
        builder.add_word(SPRM_T_DYA_FROM_TEXT_BOTTOM, row.distance_from_text_bottom);
    }
    if !row.allow_break
        || row
            .cells
            .iter()
            .any(|cell| cell.merged || cell.vertical_merge != VerticalMergeStatus::None)
    {
        // Emit the legacy form first for older readers, followed by the
        // authoritative modern form as required for equivalent SPRMs.
        builder.add_bool(SPRM_T_F_CANT_SPLIT90, true);
        builder.add_bool(SPRM_T_F_CANT_SPLIT, true);
    }
    if row.is_header {
        builder.add_bool(SPRM_T_TABLE_HEADER, true);
    }
    if row.height != 0 {
        builder.add_signed_word(0x9407, row.height);
    }
    if let Some(width) = preferred_width {
        builder.add_three_byte(SPRM_T_TABLE_WIDTH, width);
    }
    if row.auto_fit {
        builder.add_bool(SPRM_T_F_AUTOFIT, true);
    }
    if let Some(width) = width_before {
        builder.add_three_byte(SPRM_T_WIDTH_BEFORE, width);
    }
    if let Some(width) = width_after {
        builder.add_three_byte(SPRM_T_WIDTH_AFTER, width);
    }
    if row.keep_with_next {
        builder.add_bool(SPRM_T_F_KEEP_FOLLOW, true);
    }
    if let Some(width) = preferred_indent {
        builder.add_three_byte(SPRM_T_WIDTH_INDENT, width);
    }
    if let Some(look) = row.table_look {
        let flags = look.flags.bits();
        if flags & !0x07FF != 0 {
            return Err(TapBuildError::InvalidTableLookFlags(flags));
        }
        builder.add_dword(
            SPRM_T_TLP,
            u32::from_le_bytes([
                look.autoformat_index.to_le_bytes()[0],
                look.autoformat_index.to_le_bytes()[1],
                flags.to_le_bytes()[0],
                flags.to_le_bytes()[1],
            ]),
        );
    }
    if row.right_to_left {
        builder.add_word(SPRM_T_F_BI_DI, 1);
        builder.add_word(SPRM_T_F_BI_DI90, 1);
    }
    if !row.allow_overlap {
        builder.add_bool(SPRM_T_F_NO_ALLOW_OVERLAP, true);
    }
    if let Some(identifier) = row.paragraph_group_id {
        builder.add_dword(SPRM_T_IPGP, identifier);
    }
    if let Some(identifier) = row.revision_save_id {
        builder.add_dword(SPRM_T_RSID, identifier);
    }
    let mut sprms = Vec::new();
    if let Some(revision) = row.formatting_revision {
        sprms.extend_from_slice(&SPRM_T_PROP_RMARK.to_le_bytes());
        sprms.push(7);
        sprms.push(u8::from(revision.active));
        sprms.extend_from_slice(&revision.author_index.to_le_bytes());
        sprms.extend_from_slice(&revision.timestamp.to_le_bytes());
    }
    if row.properties_preserved_for_revision {
        sprms.extend_from_slice(&SPRM_T_WALL.to_le_bytes());
        sprms.push(1);
    }
    sprms.extend_from_slice(&builder.build());

    let mut operand = Vec::with_capacity(1 + (cell_count + 1) * 2 + cell_count * 20);
    operand.push(cell_count as u8);
    for boundary in boundaries {
        operand.extend_from_slice(&boundary.to_le_bytes());
    }
    for (index, width) in effective_widths.into_iter().enumerate() {
        let horizontal_merge = if row.cells[index].merged {
            1
        } else if row.cells.get(index + 1).is_some_and(|cell| cell.merged) {
            2
        } else {
            0
        };
        let text_flow = match row.cells[index].text_direction {
            TextDirection::LrTb => 0,
            TextDirection::TbRl => 1,
            TextDirection::BtLr => 3,
            TextDirection::LrBt => 4,
            TextDirection::TbLr => 5,
        };
        let vertical_merge = match row.cells[index].vertical_merge {
            VerticalMergeStatus::None => 0,
            VerticalMergeStatus::Merged => 1,
            VerticalMergeStatus::First => 3,
        };
        let vertical_alignment = match row.cells[index].vertical_alignment {
            VerticalAlignment::Top => 0,
            VerticalAlignment::Center => 1,
            VerticalAlignment::Bottom => 2,
        };
        let mut flags = horizontal_merge
            | (text_flow << 2)
            | (vertical_merge << 5)
            | (vertical_alignment << 7)
            | (3u16 << 9); // ftsDxa
        if row.cells[index].fit_text {
            flags |= 0x1000;
        }
        if row.cells[index].no_wrap {
            flags |= 0x2000;
        }
        if row.cells[index].hide_mark {
            flags |= 0x4000;
        }
        operand.extend_from_slice(&flags.to_le_bytes());
        operand.extend_from_slice(&width.to_le_bytes());
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.top)?);
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.left)?);
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.bottom)?);
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.right)?);
    }

    let encoded_size =
        u16::try_from(operand.len() + 1).map_err(|_| TapBuildError::CellWidthsOverflow)?;
    sprms.extend_from_slice(&0xD608u16.to_le_bytes());
    sprms.extend_from_slice(&encoded_size.to_le_bytes());
    sprms.extend_from_slice(&operand);

    append_table_borders(&mut sprms, row.borders)?;
    append_cell_borders(&mut sprms, &row.cells)?;
    append_cell_border_types(&mut sprms, &row.cells)?;
    append_cell_shading(&mut sprms, &row.cells)?;
    append_cell_spacing(&mut sprms, row.cell_spacing)?;
    append_cell_padding(&mut sprms, &row.cells)?;
    Ok(sprms)
}

fn append_cell_shading(sprms: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    if !cells.iter().any(|cell| cell.shading.is_some()) {
        return Ok(());
    }
    let last_shaded = cells
        .iter()
        .rposition(|cell| cell.shading.is_some())
        .expect("at least one shaded cell was checked above");
    let legacy_cells = &cells[..=last_shaded];
    let legacy = legacy_cells
        .iter()
        .map(|cell| encode_shading80(cell.shading))
        .collect::<Option<Vec<_>>>();
    if let Some(descriptors) = legacy {
        sprms.extend_from_slice(&0xD609u16.to_le_bytes());
        sprms.push((legacy_cells.len() * 2) as u8);
        for descriptor in descriptors {
            sprms.extend_from_slice(&descriptor.to_le_bytes());
        }
    }

    append_full_shading_chunks(sprms, cells, false);
    append_full_shading_chunks(sprms, cells, true);
    Ok(())
}

fn append_full_shading_chunks(sprms: &mut Vec<u8>, cells: &[TableCell], raw: bool) {
    for (chunk_index, chunk) in cells.chunks(22).enumerate() {
        let Some(last_shaded) = chunk.iter().rposition(|cell| cell.shading.is_some()) else {
            continue;
        };
        let chunk = &chunk[..=last_shaded];
        let opcode = match chunk_index {
            0 if raw => 0xD670u16,
            1 if raw => 0xD671u16,
            2 if raw => 0xD672u16,
            0 => 0xD612u16,
            1 => 0xD616u16,
            2 => 0xD60Cu16,
            _ => unreachable!("DOC rows contain at most 63 cells"),
        };
        sprms.extend_from_slice(&opcode.to_le_bytes());
        sprms.push((chunk.len() * 10) as u8);
        for cell in chunk {
            append_shading(sprms, cell.shading, raw);
        }
    }
}

fn encode_shading80(shading: Option<CellShading>) -> Option<u16> {
    let Some(shading) = shading else {
        return Some(u16::MAX);
    };
    let foreground = rgb_to_ico(shading.foreground_color).ok()?;
    let background = rgb_to_ico(shading.background_color).ok()?;
    Some(u16::from(foreground) | (u16::from(background) << 5) | ((shading.pattern as u16) << 10))
}

fn append_shading(output: &mut Vec<u8>, shading: Option<CellShading>, raw: bool) {
    let Some(shading) = shading else {
        if raw {
            output.extend_from_slice(&[0, 0, 0, 0xFF, 0, 0, 0, 0xFF]);
        } else {
            output.extend_from_slice(&[0xFF; 8]);
        }
        output.extend_from_slice(&0u16.to_le_bytes());
        return;
    };
    append_colorref(output, shading.foreground_color);
    append_colorref(output, shading.background_color);
    output.extend_from_slice(&(shading.pattern as u16).to_le_bytes());
}

fn append_colorref(output: &mut Vec<u8>, color: Option<(u8, u8, u8)>) {
    match color {
        Some((red, green, blue)) => output.extend_from_slice(&[red, green, blue, 0]),
        None => output.extend_from_slice(&[0, 0, 0, 0xFF]),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct PaddingRun {
    first: u8,
    limit: u8,
    sides: u8,
    width: u16,
}

type PaddingGetter = fn(&TableCell) -> Option<u16>;

fn append_cell_padding(sprms: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    let sides: [(u8, PaddingGetter); 4] = [
        (0x01, |cell| cell.padding_top),
        (0x02, |cell| cell.padding_left),
        (0x04, |cell| cell.padding_bottom),
        (0x08, |cell| cell.padding_right),
    ];
    let mut runs = Vec::<PaddingRun>::new();
    for (side, get_width) in sides {
        let mut first = 0;
        while first < cells.len() {
            let Some(width) = get_width(&cells[first]) else {
                first += 1;
                continue;
            };
            if width > 31_680 {
                return Err(TapBuildError::InvalidCellPadding(width));
            }
            let mut limit = first + 1;
            while limit < cells.len() && get_width(&cells[limit]) == Some(width) {
                limit += 1;
            }
            if let Some(existing) = runs.iter_mut().find(|run| {
                run.first == first as u8 && run.limit == limit as u8 && run.width == width
            }) {
                existing.sides |= side;
            } else {
                runs.push(PaddingRun {
                    first: first as u8,
                    limit: limit as u8,
                    sides: side,
                    width,
                });
            }
            first = limit;
        }
    }

    for run in runs {
        sprms.extend_from_slice(&0xD632u16.to_le_bytes());
        sprms.push(6);
        sprms.extend_from_slice(&[run.first, run.limit, run.sides, 0x03]);
        sprms.extend_from_slice(&run.width.to_le_bytes());
    }
    Ok(())
}

fn append_cell_spacing(
    sprms: &mut Vec<u8>,
    spacing: Option<CellSpacing>,
) -> Result<(), TapBuildError> {
    let Some(spacing) = spacing else {
        return Ok(());
    };
    if spacing.width > 15_840 {
        return Err(TapBuildError::InvalidCellSpacing(spacing.width));
    }
    let units = match spacing.source {
        CellSpacingSource::Explicit => 0x03,
        CellSpacingSource::TableBorder => 0x13,
    };
    sprms.extend_from_slice(&0xD633u16.to_le_bytes());
    sprms.push(6);
    sprms.extend_from_slice(&[0, 1, 0x0F, units]);
    sprms.extend_from_slice(&spacing.width.to_le_bytes());
    Ok(())
}

fn append_cell_border_types(sprms: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    let Some(last) = cells.iter().rposition(|cell| {
        let types = cell.border_type_overrides;
        types.top.is_some()
            || types.left.is_some()
            || types.bottom.is_some()
            || types.right.is_some()
    }) else {
        return Ok(());
    };
    let mut operand = Vec::with_capacity((last + 1) * 4);
    for (index, cell) in cells[..=last].iter().enumerate() {
        let types = cell.border_type_overrides;
        let (Some(top), Some(left), Some(bottom), Some(right)) =
            (types.top, types.left, types.bottom, types.right)
        else {
            return Err(TapBuildError::IncompleteCellBorderTypes(index));
        };
        operand.extend_from_slice(&[
            border_type_code(top),
            border_type_code(left),
            border_type_code(bottom),
            border_type_code(right),
        ]);
    }
    sprms.extend_from_slice(&0xD662u16.to_le_bytes());
    sprms.push(operand.len() as u8);
    sprms.extend_from_slice(&operand);
    Ok(())
}

fn encode_border80_fallback(border: Option<BorderStyle>) -> Result<[u8; 4], TapBuildError> {
    let Some(border) = border else {
        return Ok([0; 4]);
    };
    if border.border_type == BorderType::None {
        return Ok([0; 4]);
    }
    if border.spacing > 31 {
        return Err(TapBuildError::InvalidBorderSpacing(border.spacing));
    }
    let border_type = border_type_code(border.border_type);
    if matches!(border.border_type, BorderType::Outset | BorderType::Inset) {
        return Ok([0; 4]);
    }
    let ico = rgb_to_ico(border.color).unwrap_or(0);
    let effects = border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6);
    Ok([border.width, border_type, ico, effects])
}

fn border_type_code(border_type: BorderType) -> u8 {
    match border_type {
        BorderType::None => 0,
        BorderType::Single => 1,
        BorderType::Thick => 5,
        BorderType::Double => 3,
        BorderType::Dotted => 6,
        BorderType::Dashed => 7,
        BorderType::DotDash => 8,
        BorderType::DotDotDash => 9,
        BorderType::Triple => 10,
        BorderType::ThinThickSmall => 11,
        BorderType::ThickThinSmall => 12,
        BorderType::ThinThickThinSmall => 13,
        BorderType::ThinThickMedium => 14,
        BorderType::ThickThinMedium => 15,
        BorderType::ThinThickThinMedium => 16,
        BorderType::ThinThickLarge => 17,
        BorderType::ThickThinLarge => 18,
        BorderType::ThinThickThinLarge => 19,
        BorderType::Wave => 20,
        BorderType::DoubleWave => 21,
        BorderType::DashSmall => 22,
        BorderType::DashDotStroked => 23,
        BorderType::Emboss => 24,
        BorderType::Engrave => 25,
        BorderType::Outset => 26,
        BorderType::Inset => 27,
    }
}

fn append_full_border(
    output: &mut Vec<u8>,
    border: Option<BorderStyle>,
    nil: bool,
) -> Result<(), TapBuildError> {
    let Some(border) = border else {
        if nil {
            output.extend_from_slice(&[0; 4]);
            output.extend_from_slice(&[0xFF; 4]);
        } else {
            output.extend_from_slice(&[0; 8]);
        }
        return Ok(());
    };
    if border.border_type == BorderType::None {
        if nil {
            output.extend_from_slice(&[0; 4]);
            output.extend_from_slice(&[0xFF; 4]);
        } else {
            output.extend_from_slice(&[0; 8]);
        }
        return Ok(());
    }
    if border.spacing > 31 {
        return Err(TapBuildError::InvalidBorderSpacing(border.spacing));
    }
    append_colorref(output, border.color);
    output.push(border.width);
    output.push(border_type_code(border.border_type));
    output.push(border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6));
    output.push(0);
    Ok(())
}

fn append_table_borders(output: &mut Vec<u8>, borders: TableBorders) -> Result<(), TapBuildError> {
    let values = [
        borders.top,
        borders.left,
        borders.bottom,
        borders.right,
        borders.horizontal,
        borders.vertical,
    ];
    if values.iter().all(Option::is_none) {
        return Ok(());
    }
    if let Some(legacy) = values
        .iter()
        .map(|border| encode_border80_exact(*border))
        .collect::<Option<Vec<_>>>()
    {
        output.extend_from_slice(&0xD605u16.to_le_bytes());
        output.push(24);
        for border in legacy {
            output.extend_from_slice(&border);
        }
    }
    output.extend_from_slice(&0xD613u16.to_le_bytes());
    output.push(48);
    for border in values {
        append_full_border(output, border, false)?;
    }
    Ok(())
}

fn encode_border80_exact(border: Option<BorderStyle>) -> Option<[u8; 4]> {
    let Some(border) = border else {
        return Some([0; 4]);
    };
    if border.border_type == BorderType::None {
        return Some([0; 4]);
    }
    if border.spacing > 31 || matches!(border.border_type, BorderType::Outset | BorderType::Inset) {
        return None;
    }
    let ico = rgb_to_ico(border.color).ok()?;
    Some([
        border.width,
        border_type_code(border.border_type),
        ico,
        border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6),
    ])
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BorderRun {
    first: u8,
    limit: u8,
    sides: u8,
    border: BorderStyle,
}

type BorderGetter = fn(&TableCell) -> Option<BorderStyle>;

fn append_cell_borders(output: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    let sides: [(u8, BorderGetter); 6] = [
        (0x01, |cell| cell.borders.top),
        (0x02, |cell| cell.borders.left),
        (0x04, |cell| cell.borders.bottom),
        (0x08, |cell| cell.borders.right),
        (0x10, |cell| cell.borders.diagonal_down),
        (0x20, |cell| cell.borders.diagonal_up),
    ];
    let mut runs = Vec::<BorderRun>::new();
    for (side, get_border) in sides {
        let mut first = 0;
        while first < cells.len() {
            let Some(border) = get_border(&cells[first]) else {
                first += 1;
                continue;
            };
            let mut limit = first + 1;
            while limit < cells.len() && get_border(&cells[limit]) == Some(border) {
                limit += 1;
            }
            if let Some(run) = runs.iter_mut().find(|run| {
                run.first == first as u8 && run.limit == limit as u8 && run.border == border
            }) {
                run.sides |= side;
            } else {
                runs.push(BorderRun {
                    first: first as u8,
                    limit: limit as u8,
                    sides: side,
                    border,
                });
            }
            first = limit;
        }
    }
    for run in runs {
        output.extend_from_slice(&0xD62Fu16.to_le_bytes());
        output.push(11);
        output.extend_from_slice(&[run.first, run.limit, run.sides]);
        append_full_border(output, Some(run.border), true)?;
    }
    Ok(())
}

fn rgb_to_ico(color: Option<(u8, u8, u8)>) -> Result<u8, (u8, u8, u8)> {
    Ok(match color {
        None => 0,
        Some((0, 0, 0)) => 1,
        Some((0, 0, 255)) => 2,
        Some((0, 255, 255)) => 3,
        Some((0, 255, 0)) => 4,
        Some((255, 0, 255)) => 5,
        Some((255, 0, 0)) => 6,
        Some((255, 255, 0)) => 7,
        Some((255, 255, 255)) => 8,
        Some((0, 0, 128)) => 9,
        Some((0, 128, 128)) => 10,
        Some((0, 128, 0)) => 11,
        Some((128, 0, 128)) => 12,
        Some((128, 0, 0)) => 13,
        Some((128, 128, 0)) => 14,
        Some((128, 128, 128)) => 15,
        Some((192, 192, 192)) => 16,
        Some(color) => return Err(color),
    })
}
