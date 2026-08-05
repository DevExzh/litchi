//! Binary text, character-property, and paragraph-property encoders.

use crate::parts::pap::{
    AutoNumberAlignment, Border as ParagraphBorder, BorderStyle as ParagraphBorderStyle,
    FrameHorizontalAnchor, FrameHorizontalPosition, FrameTextWrap, FrameVerticalAnchor,
    FrameVerticalPosition, LegacyAutoNumbering, PhysicalJustification, TabAlignment, TabLeader,
    TabStop,
};
use crate::sprm_operations::*;
use crate::writer::font_table::FontTableBuilder;
use crate::writer::revisions::TextRevision;

use super::model::*;
pub(super) fn write_textbox_story_text(
    texts: &[&str],
    text_stream: &mut Vec<u8>,
    current_cp: &mut u32,
) -> Result<(Vec<u32>, u32), WriteError> {
    let story_start_cp = *current_cp;
    let mut start_cps = Vec::with_capacity(texts.len());
    for text in texts {
        start_cps.push(*current_cp - story_start_cp);
        for paragraph in text.replace("\r\n", "\n").replace('\r', "\n").split('\n') {
            let para_len = utf16_code_unit_len(paragraph)?;
            for unit in paragraph.encode_utf16() {
                text_stream.extend_from_slice(&unit.to_le_bytes());
            }
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            *current_cp += para_len + 1;
        }
        // Trailing CR of this text box's text, as Word writes.
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        *current_cp += 1;
    }
    // Story-final CR, included in the ccp count.
    text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
    *current_cp += 1;
    Ok((start_cps, *current_cp - story_start_cp))
}

/// Build a CHPX grpprl (group of SPRMs) from CharacterFormatting
pub(super) fn build_chpx_grpprl(
    fmt: &CharacterFormatting,
    font_builder: &mut FontTableBuilder,
) -> Vec<u8> {
    let mut grp = Vec::with_capacity(16);

    #[inline]
    fn push_byte(grp: &mut Vec<u8>, opcode: u16, val: u8) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.push(val);
    }

    #[inline]
    fn push_word(grp: &mut Vec<u8>, opcode: u16, val: u16) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&val.to_le_bytes());
    }

    #[inline]
    fn push_dword(grp: &mut Vec<u8>, opcode: u16, val: u32) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&val.to_le_bytes());
    }

    if let Some(style_index) = fmt.style_index {
        push_word(&mut grp, SPRM_C_ISTD, style_index);
    }
    // Bold
    if let Some(b) = fmt.bold {
        push_byte(&mut grp, SPRM_C_F_BOLD, if b { 1 } else { 0 });
    }
    // Italic
    if let Some(i) = fmt.italic {
        push_byte(&mut grp, SPRM_C_F_ITALIC, if i { 1 } else { 0 });
    }
    // Underline (1 = single, 0 = none)
    if let Some(u) = fmt.underline {
        push_byte(&mut grp, SPRM_C_KUL, if u { 1 } else { 0 });
    }
    // Strikethrough
    if let Some(s) = fmt.strike {
        push_byte(&mut grp, SPRM_C_F_STRIKE, if s { 1 } else { 0 });
    }
    // Double strikethrough
    if let Some(ds) = fmt.double_strike {
        push_byte(&mut grp, SPRM_C_F_D_STRIKE, if ds { 1 } else { 0 });
    }
    // Superscript/Subscript via sprmCIss (0=none,1=super,2=sub)
    let mut iss: Option<u8> = None;
    if let Some(true) = fmt.superscript {
        iss = Some(1);
    } else if let Some(true) = fmt.subscript {
        iss = Some(2);
    }
    if let Some(v) = iss {
        push_byte(&mut grp, SPRM_C_ISS, v);
    }
    // Small caps / All caps / Hidden
    if let Some(sc) = fmt.small_caps {
        push_byte(&mut grp, SPRM_C_F_SMALL_CAPS, if sc { 1 } else { 0 });
    }
    if let Some(ac) = fmt.all_caps {
        push_byte(&mut grp, SPRM_C_F_CAPS, if ac { 1 } else { 0 });
    }
    if let Some(h) = fmt.hidden {
        push_byte(&mut grp, SPRM_C_F_VANISH, if h { 1 } else { 0 });
    }
    // Special/Field vanish (for field codes and control chars)
    if let Some(sp) = fmt.special {
        push_byte(&mut grp, SPRM_C_F_SPEC, if sp { 1 } else { 0 });
    }
    if let Some(vn) = fmt.field_vanish {
        push_byte(&mut grp, SPRM_C_F_FLD_VANISH, if vn { 1 } else { 0 });
    }
    // Font size (half-points)
    if let Some(hps) = fmt.font_size {
        push_word(&mut grp, SPRM_C_HPS, hps);
    }
    if let Some(position) = fmt.position {
        grp.extend_from_slice(&SPRM_C_HPS_POS.to_le_bytes());
        grp.extend_from_slice(&position.half_points().to_le_bytes());
    }
    if let Some(hyphenation) = fmt.hyphenation {
        grp.extend_from_slice(&SPRM_C_HRESI.to_le_bytes());
        grp.extend_from_slice(&hyphenation.bytes());
    }
    // Font name -> map to ftc index via FontTableBuilder and set default font
    if let Some(name) = &fmt.font_name {
        let idx = font_builder.get_or_add(name);
        push_word(&mut grp, SPRM_C_FTC_DEFAULT, idx);
    }
    // Color (RGB) -> sprmCCv expects a 4-byte value
    if let Some((r, g, b)) = fmt.color {
        let cv: u32 = (r as u32) | ((g as u32) << 8) | ((b as u32) << 16);
        push_dword(&mut grp, SPRM_C_CV, cv);
    }
    if let Some(effect) = fmt.text_effect {
        push_byte(&mut grp, SPRM_C_SFXT_TEXT, effect.into());
    }

    grp
}

pub(super) fn build_revision_chpx_grpprl(
    fmt: &CharacterFormatting,
    font_builder: &mut FontTableBuilder,
    revisions: Option<&RevisionWriterData>,
) -> Result<Vec<u8>, WriteError> {
    if fmt
        .preserved_properties_for_revision
        .as_ref()
        .is_some_and(|previous| previous.preserved_properties_for_revision.is_some())
    {
        return Err(WriteError::InvalidData(
            "DOC character property revisions cannot contain nested preserved states".to_string(),
        ));
    }
    if fmt.insertion_revision.is_some() && fmt.deletion_revision.is_some() {
        return Err(WriteError::InvalidData(
            "a DOC character run cannot be both an insertion and a deletion".to_string(),
        ));
    }
    let mut grp = if let Some(previous) = &fmt.preserved_properties_for_revision {
        let mut grp = build_revision_chpx_grpprl(previous, font_builder, revisions)?;
        grp.extend_from_slice(&SPRM_C_WALL.to_le_bytes());
        grp.push(1);
        grp.extend_from_slice(&build_chpx_grpprl(fmt, font_builder));
        grp
    } else {
        build_chpx_grpprl(fmt, font_builder)
    };
    let Some(revisions) = revisions else {
        return Ok(grp);
    };
    let mut append = |revision: &TextRevision,
                      flag_opcode: u16,
                      author_opcode: u16,
                      time_opcode: u16,
                      reason_opcode: u16,
                      rsid_opcode: u16|
     -> Result<(), WriteError> {
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            WriteError::InvalidData("DOC revision author was not indexed".to_string())
        })?;
        grp.extend_from_slice(&flag_opcode.to_le_bytes());
        grp.push(1);
        grp.extend_from_slice(&author_opcode.to_le_bytes());
        grp.extend_from_slice(&author_index.to_le_bytes());
        if revision.timestamp.is_some() {
            grp.extend_from_slice(&time_opcode.to_le_bytes());
            grp.extend_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        }
        let structured_reason = revision.reason.map(crate::RevisionReason::raw);
        if let (Some(raw), Some(structured)) = (revision.revision_id, structured_reason)
            && raw != structured
        {
            return Err(WriteError::InvalidData(
                "DOC revision contains conflicting raw and structured reason codes".to_string(),
            ));
        }
        if let Some(reason) = structured_reason.or(revision.revision_id) {
            if reason > crate::RevisionReason::MAX_VALUE {
                return Err(WriteError::InvalidData(
                    "DOC revision reason code is undefined".to_string(),
                ));
            }
            grp.extend_from_slice(&reason_opcode.to_le_bytes());
            grp.extend_from_slice(&reason.to_le_bytes());
        }
        if let Some(revision_save_id) = revision.revision_save_id {
            grp.extend_from_slice(&rsid_opcode.to_le_bytes());
            grp.extend_from_slice(&revision_save_id.to_le_bytes());
        }
        Ok(())
    };
    if let Some(revision) = &fmt.insertion_revision {
        append(
            revision,
            SPRM_C_F_RMARK,
            SPRM_C_IBST_RMARK,
            SPRM_C_DTTM_RMARK,
            SPRM_C_IDSL_RMARK,
            SPRM_C_RSID_TEXT,
        )?;
    }
    if let Some(revision) = &fmt.deletion_revision {
        append(
            revision,
            SPRM_C_F_RMARK_DEL,
            SPRM_C_IBST_RMARK_DEL,
            SPRM_C_DTTM_RMARK_DEL,
            SPRM_C_IDSL_RMARK_DEL,
            SPRM_C_RSID_RM_DEL,
        )?;
    }
    if let Some(revision) = &fmt.formatting_revision {
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            WriteError::InvalidData("DOC revision author was not indexed".to_string())
        })?;
        grp.extend_from_slice(&SPRM_C_PROP_RMARK_CURRENT.to_le_bytes());
        grp.push(7);
        grp.push(1);
        grp.extend_from_slice(&author_index.to_le_bytes());
        grp.extend_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        if let Some(reason) = revision.reason {
            let insertion_reason = fmt.insertion_revision.as_ref().and_then(|insertion| {
                insertion
                    .reason
                    .map(crate::RevisionReason::raw)
                    .or(insertion.revision_id)
            });
            if insertion_reason.is_some_and(|value| value != reason.raw()) {
                return Err(WriteError::InvalidData(
                    "DOC insertion and formatting revisions have conflicting reason codes"
                        .to_string(),
                ));
            }
            grp.extend_from_slice(&SPRM_C_IDSL_RMARK.to_le_bytes());
            grp.extend_from_slice(&reason.raw().to_le_bytes());
        }
        if let Some(revision_save_id) = revision.revision_save_id {
            grp.extend_from_slice(&SPRM_C_RSID_PROP.to_le_bytes());
            grp.extend_from_slice(&revision_save_id.to_le_bytes());
        }
    }
    if let Some(revision) = &fmt.display_field_revision {
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            WriteError::InvalidData("DOC display-field revision author was not indexed".to_string())
        })?;
        let units = revision.previous_result.encode_utf16().collect::<Vec<_>>();
        if units.len() > 15 {
            return Err(WriteError::InvalidData(
                "DOC LISTNUM previous result exceeds its 15-code-unit XST".to_string(),
            ));
        }
        let mut operand = [0u8; 39];
        operand[0] = 1;
        operand[1..3].copy_from_slice(&author_index.to_le_bytes());
        operand[3..7].copy_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        operand[7..9].copy_from_slice(&(units.len() as u16).to_le_bytes());
        for (index, unit) in units.into_iter().enumerate() {
            let offset = 9 + index * 2;
            operand[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
        }
        grp.extend_from_slice(&SPRM_C_DISP_FLD_RMARK.to_le_bytes());
        grp.push(39);
        grp.extend_from_slice(&operand);
    }
    Ok(grp)
}

/// Build a PAPX grpprl (group of SPRMs) from ParagraphFormatting
pub(super) fn build_papx_grpprl(fmt: &ParagraphFormatting) -> Vec<u8> {
    let mut grp = Vec::with_capacity(16);

    #[inline]
    fn push_byte(grp: &mut Vec<u8>, opcode: u16, val: u8) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.push(val);
    }

    #[inline]
    fn push_i16(grp: &mut Vec<u8>, opcode: u16, val: i16) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&(val as u16).to_le_bytes());
    }

    #[inline]
    fn push_u16(grp: &mut Vec<u8>, opcode: u16, val: u16) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.extend_from_slice(&val.to_le_bytes());
    }

    #[inline]
    fn push_bool(grp: &mut Vec<u8>, opcode: u16, val: bool) {
        grp.extend_from_slice(&opcode.to_le_bytes());
        grp.push(if val { 1 } else { 0 });
    }

    if let Some(style_index) = fmt.style_index {
        push_u16(&mut grp, SPRM_P_ISTD, style_index);
    }
    // Alignment. Emit a compatible physical value before the authoritative logical value.
    if let Some(jc) = fmt.alignment {
        let physical = match jc {
            0..=3 => Some(jc),
            4 | 5 => Some(4),
            7 | 8 => Some(5),
            9 => Some(3),
            _ => None,
        };
        if let Some(physical) = physical {
            push_byte(&mut grp, SPRM_P_JC, physical);
        }
        push_byte(&mut grp, SPRM_P_JC_LOGICAL, jc);
    } else if let Some(physical) = fmt.physical_justification {
        let code = match physical {
            PhysicalJustification::Left => 0,
            PhysicalJustification::Center => 1,
            PhysicalJustification::Right => 2,
            PhysicalJustification::LowCompression => 3,
            PhysicalJustification::MediumCompression => 4,
            PhysicalJustification::HighCompression => 5,
        };
        push_byte(&mut grp, SPRM_P_JC, code);
    }
    if let Some(style) = fmt.legacy_border_style {
        push_byte(&mut grp, SPRM_P_BRCL, style as u8);
    }
    if let Some(position) = fmt.legacy_border_position {
        push_byte(&mut grp, SPRM_P_BRCP, position as u8);
    }
    // Indents (twips). Emit legacy and modern variants. Values are signed twips.
    if let Some(dxa_left) = fmt.left_indent {
        let v = dxa_left.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
        push_i16(&mut grp, SPRM_P_DXA_LEFT, v);
        push_i16(&mut grp, SPRM_P_DXA_LEFT_2000, v);
    }
    if let Some(dxa_right) = fmt.right_indent {
        let v = dxa_right.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
        push_i16(&mut grp, SPRM_P_DXA_RIGHT, v);
        push_i16(&mut grp, SPRM_P_DXA_RIGHT_2000, v);
    }
    if let Some(dxa_first) = fmt.first_line_indent {
        let v = dxa_first.clamp(i16::MIN as i32, i16::MAX as i32) as i16;
        push_i16(&mut grp, SPRM_P_DXA_LEFT1, v);
        push_i16(&mut grp, SPRM_P_DXA_LEFT1_2000, v);
    }
    if let Some(dxc_left) = fmt.left_indent_chars {
        push_i16(&mut grp, SPRM_P_DXC_LEFT, dxc_left);
    }
    if let Some(dxc_right) = fmt.right_indent_chars {
        push_i16(&mut grp, SPRM_P_DXC_RIGHT, dxc_right);
    }
    if let Some(dxc_first) = fmt.first_line_indent_chars {
        push_i16(&mut grp, SPRM_P_DXC_LEFT1, dxc_first);
    }
    // Spacing (twips)
    if let Some(dya_before) = fmt.space_before {
        push_u16(&mut grp, SPRM_P_DYA_BEFORE, dya_before);
    }
    if let Some(dya_after) = fmt.space_after {
        push_u16(&mut grp, SPRM_P_DYA_AFTER, dya_after);
    }
    if let Some(disabled) = fmt.no_line_numbering {
        push_bool(&mut grp, SPRM_P_F_NO_LINE_NUMB, disabled);
    }
    if let Some(dyl_before) = fmt.space_before_lines {
        push_i16(&mut grp, SPRM_P_DYL_BEFORE, dyl_before);
    }
    if let Some(dyl_after) = fmt.space_after_lines {
        push_i16(&mut grp, SPRM_P_DYL_AFTER, dyl_after);
    }

    // Auto spacing flags
    if let Some(auto) = fmt.space_before_auto {
        push_bool(&mut grp, SPRM_P_F_DYA_BEFORE_AUTO, auto);
    }
    if let Some(auto) = fmt.space_after_auto {
        push_bool(&mut grp, SPRM_P_F_DYA_AFTER_AUTO, auto);
    }
    if let Some(open) = fmt.open_table_cell_mark {
        push_bool(&mut grp, SPRM_P_F_OPEN_TCH, open);
    }

    // Side-by-side and pagination controls
    if let Some(side_by_side) = fmt.side_by_side {
        push_bool(&mut grp, SPRM_P_F_SIDE_BY_SIDE, side_by_side);
    }
    if let Some(keep) = fmt.keep {
        push_bool(&mut grp, SPRM_P_F_KEEP, keep);
    }
    if let Some(keep_next) = fmt.keep_with_next {
        push_bool(&mut grp, SPRM_P_F_KEEP_FOLLOW, keep_next);
    }
    if let Some(pbb) = fmt.page_break_before {
        push_bool(&mut grp, SPRM_P_F_PAGE_BREAK_BEFORE, pbb);
    }

    // Widow/orphan control
    if let Some(wc) = fmt.widow_control {
        push_bool(&mut grp, SPRM_P_F_WIDOW_CONTROL, wc);
    }
    for (opcode, value) in [
        (SPRM_P_F_LOCKED, fmt.frame_anchor_locked),
        (SPRM_P_F_KINSOKU, fmt.kinsoku),
        (SPRM_P_F_WORD_WRAP, fmt.word_wrap),
        (SPRM_P_F_OVERFLOW_PUNCT, fmt.overflow_punctuation),
        (SPRM_P_F_TOP_LINE_PUNCT, fmt.top_line_punctuation),
        (SPRM_P_F_AUTO_SPACE_DE, fmt.auto_space_east_asian_latin),
        (SPRM_P_F_AUTO_SPACE_DN, fmt.auto_space_east_asian_numbers),
    ] {
        if let Some(value) = value {
            push_bool(&mut grp, opcode, value);
        }
    }
    if let Some(alignment) = fmt.font_alignment {
        push_u16(&mut grp, SPRM_P_W_ALIGN_FONT, alignment as u16);
    }
    if let Some(flow) = fmt.frame_text_flow {
        let value = u16::from(flow.vertical)
            | (u16::from(flow.backwards) << 1)
            | (u16::from(flow.rotate_font) << 2);
        push_u16(&mut grp, SPRM_P_FRAME_TEXT_FLOW, value);
    }
    if let Some(position) = fmt.frame_horizontal_position {
        let value = match position {
            FrameHorizontalPosition::Left => 0,
            FrameHorizontalPosition::Center => -4,
            FrameHorizontalPosition::Right => -8,
            FrameHorizontalPosition::Inside => -12,
            FrameHorizontalPosition::Outside => -16,
            FrameHorizontalPosition::Offset(offset) => offset + 1,
        };
        push_i16(&mut grp, SPRM_P_DXA_ABS, value);
    }
    if let Some(position) = fmt.frame_vertical_position {
        let value = match position {
            FrameVerticalPosition::Inline => 0,
            FrameVerticalPosition::Top => -4,
            FrameVerticalPosition::Center => -8,
            FrameVerticalPosition::Bottom => -12,
            FrameVerticalPosition::Inside => -16,
            FrameVerticalPosition::Outside => -20,
            FrameVerticalPosition::Offset(offset) => offset + 1,
        };
        push_i16(&mut grp, SPRM_P_DYA_ABS, value);
    }
    if let Some(width) = fmt.frame_width {
        push_u16(&mut grp, SPRM_P_DXA_WIDTH, width);
    }
    if let Some(anchor) = fmt.frame_anchor {
        let vertical = match anchor.vertical {
            FrameVerticalAnchor::Margin => 0,
            FrameVerticalAnchor::Page => 1,
            FrameVerticalAnchor::Paragraph => 2,
            FrameVerticalAnchor::None => 3,
        };
        let horizontal = match anchor.horizontal {
            FrameHorizontalAnchor::Column => 0,
            FrameHorizontalAnchor::Margin => 1,
            FrameHorizontalAnchor::Page => 2,
            FrameHorizontalAnchor::None => 3,
        };
        push_byte(&mut grp, SPRM_P_PC, (vertical << 4) | (horizontal << 6));
    }
    if let Some(in_table) = fmt.in_table {
        push_bool(&mut grp, SPRM_P_F_IN_TABLE, in_table);
    }
    if let Some(terminating) = fmt.table_terminating_paragraph {
        push_bool(&mut grp, SPRM_P_F_TTP, terminating);
    }
    if let Some(wrap) = fmt.frame_text_wrap {
        push_byte(&mut grp, SPRM_P_WR, wrap as u8);
    }
    if let Some(height) = fmt.frame_height {
        push_u16(
            &mut grp,
            SPRM_P_W_HEIGHT_ABS,
            height.height_twips | (u16::from(height.minimum) << 15),
        );
    }
    if let Some(distance) = fmt.frame_horizontal_text_distance {
        push_i16(&mut grp, SPRM_P_DXA_FROM_TEXT, distance);
    }
    if let Some(distance) = fmt.frame_vertical_text_distance {
        push_i16(&mut grp, SPRM_P_DYA_FROM_TEXT, distance);
    }
    if let Some(drop_cap) = fmt.drop_cap {
        let kind = match drop_cap.kind {
            crate::parts::pap::DropCapType::Regular => 1u16,
            crate::parts::pap::DropCapType::Margin => 2,
        };
        push_u16(
            &mut grp,
            SPRM_P_DCS,
            kind | (u16::from(drop_cap.lines) << 3),
        );
    }
    if let Some(disabled) = fmt.no_auto_hyphenation {
        push_bool(&mut grp, SPRM_P_F_NO_AUTO_HYPH, disabled);
    }

    // BiDi paragraph
    if let Some(bidi) = fmt.bidi {
        push_bool(&mut grp, SPRM_P_F_BI_DI, bidi);
    }
    if let Some(use_grid) = fmt.use_page_setup_settings {
        push_bool(&mut grp, SPRM_P_F_USE_PGSU_SETTINGS, use_grid);
    }
    if let Some(adjust) = fmt.adjust_right_indent {
        push_bool(&mut grp, SPRM_P_F_ADJUST_RIGHT, adjust);
    }

    // Outline level
    if let Some(lvl) = fmt.outline_level {
        grp.extend_from_slice(&SPRM_P_OUT_LVL.to_le_bytes());
        grp.push(lvl);
    }

    // Floating-object overlap and text-box wrapping behavior
    if let Some(no_overlap) = fmt.no_allow_overlap {
        push_bool(&mut grp, SPRM_P_F_NO_ALLOW_OVERLAP, no_overlap);
    }
    if let Some(cs) = fmt.contextual_spacing {
        push_bool(&mut grp, SPRM_P_F_CONTEXTUAL_SPACING, cs);
    }
    if let Some(mi) = fmt.mirror_indents {
        push_bool(&mut grp, SPRM_P_F_MIRROR_INDENTS, mi);
    }
    if let Some(tight_wrap) = fmt.text_box_tight_wrap {
        push_byte(&mut grp, SPRM_P_TTWO, tight_wrap as u8);
    }
    for (opcode, border) in [
        (SPRM_P_BRC_TOP, fmt.borders.top),
        (SPRM_P_BRC_LEFT, fmt.borders.left),
        (SPRM_P_BRC_BOTTOM, fmt.borders.bottom),
        (SPRM_P_BRC_RIGHT, fmt.borders.right),
        (SPRM_P_BRC_BETWEEN, fmt.borders.between),
        (SPRM_P_BRC_BAR, fmt.borders.bar),
    ] {
        if let Some(border) = border {
            append_paragraph_border(&mut grp, opcode, border);
        }
    }
    if let Some(shading) = fmt.shading {
        grp.extend_from_slice(&SPRM_P_SHD.to_le_bytes());
        grp.push(10);
        for color in [shading.foreground_color, shading.background_color] {
            match color {
                Some((red, green, blue)) => grp.extend_from_slice(&[red, green, blue, 0]),
                None => grp.extend_from_slice(&[0, 0, 0, 0xFF]),
            }
        }
        grp.extend_from_slice(&(shading.pattern as u16).to_le_bytes());
    }
    if let Some(applied) = fmt.numbering_revision_list_applied {
        push_bool(&mut grp, SPRM_P_F_NUM_RM_INS, applied);
    }

    // List numbering: ilvl (list level) and ilfo (list format override)
    if let Some(ilvl) = fmt.ilvl {
        push_byte(&mut grp, SPRM_P_ILVL, ilvl);
    }
    if let Some(ilfo) = fmt.ilfo {
        push_u16(&mut grp, SPRM_P_ILFO, ilfo);
    }
    if let Some(autonumbering) = &fmt.legacy_autonumbering {
        append_legacy_autonumbering(&mut grp, autonumbering);
    }
    if let Some(revision_save_id) = fmt.revision_save_id {
        grp.extend_from_slice(&SPRM_P_RSID.to_le_bytes());
        grp.extend_from_slice(&revision_save_id.to_le_bytes());
    }

    // Line spacing (LSPD: 4 bytes = dyaLine (i16 LE), fMulti (i16 LE))
    if let Some(ls) = fmt.line_spacing {
        let mut bytes = [0u8; 4];
        let f_multi: u16 = if ls.is_multiple { 1 } else { 0 };
        bytes[0..2].copy_from_slice(&(ls.dya_line as u16).to_le_bytes());
        bytes[2..4].copy_from_slice(&f_multi.to_le_bytes());
        grp.extend_from_slice(&SPRM_P_DYA_LINE.to_le_bytes());
        grp.extend_from_slice(&bytes);
    }
    append_tab_changes(&mut grp, &fmt.tab_stops_to_delete, &fmt.tab_stops_to_add);

    grp
}

pub(super) fn append_tab_changes(output: &mut Vec<u8>, deletes: &[i32], additions: &[TabStop]) {
    let mut deletes = deletes.to_vec();
    deletes.sort_unstable();
    for chunk in deletes.chunks(64) {
        output.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
        output.push((2 + chunk.len() * 2) as u8);
        output.push(chunk.len() as u8);
        for position in chunk {
            output.extend_from_slice(&(*position as i16).to_le_bytes());
        }
        output.push(0);
    }

    let mut additions = additions.to_vec();
    additions.sort_unstable_by_key(|tab| tab.position);
    for chunk in additions.chunks(64) {
        output.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
        output.push((2 + chunk.len() * 3) as u8);
        output.push(0);
        output.push(chunk.len() as u8);
        for tab in chunk {
            output.extend_from_slice(&(tab.position as i16).to_le_bytes());
        }
        for tab in chunk {
            let alignment = match tab.alignment {
                TabAlignment::Left => 0,
                TabAlignment::Center => 1,
                TabAlignment::Right => 2,
                TabAlignment::Decimal => 3,
                TabAlignment::Bar => 4,
                TabAlignment::List => 6,
            };
            let leader = if tab.alignment == TabAlignment::Bar {
                0
            } else {
                match tab.leader {
                    TabLeader::None => 0,
                    TabLeader::Dots => 1,
                    TabLeader::Hyphens => 2,
                    TabLeader::Underline => 3,
                    TabLeader::Heavy => 4,
                    TabLeader::MiddleDot => 5,
                    TabLeader::DefaultLeader => 7,
                }
            };
            output.push(alignment | (leader << 3));
        }
    }
}

pub(super) fn append_legacy_autonumbering(output: &mut Vec<u8>, value: &LegacyAutoNumbering) {
    let mut operand = [0u8; 84];
    operand[0] = value.number_format as u8;
    let prefix = value.prefix.encode_utf16().collect::<Vec<_>>();
    let suffix = value.suffix.encode_utf16().collect::<Vec<_>>();
    operand[1] = prefix.len() as u8;
    operand[2] = suffix.len() as u8;
    operand[3] = match value.alignment {
        AutoNumberAlignment::Left => 0,
        AutoNumberAlignment::Center => 1,
        AutoNumberAlignment::Right => 2,
        AutoNumberAlignment::Justified => 3,
    } | (u8::from(value.include_previous_levels) << 2)
        | (u8::from(value.hanging_indent) << 3)
        | (u8::from(value.set_bold) << 4)
        | (u8::from(value.set_italic) << 5)
        | (u8::from(value.set_small_caps) << 6)
        | (u8::from(value.set_caps) << 7);
    operand[4] = u8::from(value.set_strike)
        | (u8::from(value.set_underline) << 1)
        | (u8::from(value.prefix_space) << 2)
        | (u8::from(value.bold) << 3)
        | (u8::from(value.italic) << 4)
        | (u8::from(value.small_caps) << 5)
        | (u8::from(value.caps) << 6)
        | (u8::from(value.strike) << 7);
    operand[5] = value.underline | (value.color_index << 3);
    operand[6..8].copy_from_slice(&value.font_index.to_le_bytes());
    operand[8..10].copy_from_slice(&value.font_size_half_points.to_le_bytes());
    operand[10..12].copy_from_slice(&value.start_at.to_le_bytes());
    operand[12..14].copy_from_slice(&value.indent_twips.to_le_bytes());
    operand[14..16].copy_from_slice(&value.space_twips.to_le_bytes());
    operand[16] = u8::from(value.number_once_per_cell);
    operand[17] = u8::from(value.number_across_cells);
    operand[18] = u8::from(value.restart_each_section);
    for (index, unit) in prefix.into_iter().chain(suffix).enumerate() {
        let offset = 20 + index * 2;
        operand[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
    }
    output.extend_from_slice(&SPRM_P_ANLD.to_le_bytes());
    output.push(operand.len() as u8);
    output.extend_from_slice(&operand);
}

pub(super) fn append_paragraph_border(output: &mut Vec<u8>, opcode: u16, border: ParagraphBorder) {
    output.extend_from_slice(&opcode.to_le_bytes());
    output.push(8);
    match border.color {
        Some((red, green, blue)) => output.extend_from_slice(&[red, green, blue, 0]),
        None => output.extend_from_slice(&[0, 0, 0, 0xFF]),
    }
    output.push(border.width);
    output.push(match border.style {
        ParagraphBorderStyle::None => 0,
        ParagraphBorderStyle::Single => 1,
        ParagraphBorderStyle::Double => 3,
        ParagraphBorderStyle::Thick => 5,
        ParagraphBorderStyle::Dotted => 6,
        ParagraphBorderStyle::Dashed => 7,
        ParagraphBorderStyle::DotDash => 8,
        ParagraphBorderStyle::DotDotDash => 9,
        ParagraphBorderStyle::Triple => 10,
        ParagraphBorderStyle::ThinThickSmallGap => 11,
        ParagraphBorderStyle::ThickThinSmallGap => 12,
        ParagraphBorderStyle::ThinThickThinSmallGap => 13,
        ParagraphBorderStyle::ThinThickMediumGap => 14,
        ParagraphBorderStyle::ThickThinMediumGap => 15,
        ParagraphBorderStyle::ThinThickThinMediumGap => 16,
        ParagraphBorderStyle::ThinThickLargeGap => 17,
        ParagraphBorderStyle::ThickThinLargeGap => 18,
        ParagraphBorderStyle::ThinThickThinLargeGap => 19,
        ParagraphBorderStyle::Wave => 20,
        ParagraphBorderStyle::DoubleWave => 21,
        ParagraphBorderStyle::DashSmallGap => 22,
        ParagraphBorderStyle::DashDotStroked => 23,
        ParagraphBorderStyle::ThreeDEmboss => 24,
        ParagraphBorderStyle::ThreeDEngrave => 25,
        ParagraphBorderStyle::Outset => 26,
        ParagraphBorderStyle::Inset => 27,
    });
    output.push(border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6));
    output.push(0);
}

pub(super) fn build_revision_papx_grpprl(
    fmt: &ParagraphFormatting,
    revisions: Option<&RevisionWriterData>,
) -> Result<Vec<u8>, WriteError> {
    if fmt
        .preserved_properties_for_revision
        .as_ref()
        .is_some_and(|previous| previous.preserved_properties_for_revision.is_some())
    {
        return Err(WriteError::InvalidData(
            "DOC paragraph property revisions cannot contain nested preserved states".to_string(),
        ));
    }
    if let Some(alignment) = fmt.alignment
        && alignment > 9
    {
        return Err(WriteError::InvalidData(format!(
            "DOC paragraph alignment {alignment} is outside 0..=9"
        )));
    }
    if let Some(outline_level) = fmt.outline_level
        && outline_level > 9
    {
        return Err(WriteError::InvalidData(format!(
            "DOC paragraph outline level {outline_level} is outside 0..=9"
        )));
    }
    if let Some(level) = fmt.ilvl
        && level > 8
        && level != 0x0C
    {
        return Err(WriteError::InvalidData(format!(
            "DOC paragraph list level {level} is neither 0..=8 nor the skip value 12"
        )));
    }
    if let Some(ilfo) = fmt.ilfo
        && (0x07FF..=0xF800).contains(&ilfo)
    {
        return Err(WriteError::InvalidData(format!(
            "DOC paragraph list override {ilfo:#06x} is reserved"
        )));
    }
    if let Some(value) = &fmt.legacy_autonumbering {
        let prefix_units = value.prefix.encode_utf16().count();
        let suffix_units = value.suffix.encode_utf16().count();
        if prefix_units + suffix_units > 32 {
            return Err(WriteError::InvalidData(format!(
                "DOC legacy autonumber label uses {} UTF-16 units; maximum is 32",
                prefix_units + suffix_units
            )));
        }
        if value.underline > 7 {
            return Err(WriteError::InvalidData(format!(
                "DOC legacy autonumber underline {} exceeds 7",
                value.underline
            )));
        }
        if value.color_index > 16 {
            return Err(WriteError::InvalidData(format!(
                "DOC legacy autonumber color index {} exceeds 16",
                value.color_index
            )));
        }
        if !(-31_680..=31_680).contains(&value.indent_twips) {
            return Err(WriteError::InvalidData(format!(
                "DOC legacy autonumber indent {} is outside -31680..=31680",
                value.indent_twips
            )));
        }
        if value.space_twips > 31_680 {
            return Err(WriteError::InvalidData(format!(
                "DOC legacy autonumber spacing {} exceeds 31680",
                value.space_twips
            )));
        }
    }
    for (name, value) in [
        ("left_indent", fmt.left_indent),
        ("right_indent", fmt.right_indent),
        ("first_line_indent", fmt.first_line_indent),
    ] {
        if let Some(value) = value
            && !(-31_680..=31_680).contains(&value)
        {
            return Err(WriteError::InvalidData(format!(
                "DOC paragraph {name} value {value} is outside -31680..=31680"
            )));
        }
    }
    for (name, value) in [
        ("space_before", fmt.space_before),
        ("space_after", fmt.space_after),
    ] {
        if let Some(value) = value
            && value > 31_680
        {
            return Err(WriteError::InvalidData(format!(
                "DOC paragraph {name} value {value} exceeds 31680"
            )));
        }
    }
    if let Some(spacing) = fmt.line_spacing
        && !(-31_680..=31_680).contains(&spacing.dya_line)
    {
        return Err(WriteError::InvalidData(format!(
            "DOC paragraph line spacing {} is outside the LSPD range",
            spacing.dya_line
        )));
    }
    let added_tab_positions = fmt
        .tab_stops_to_add
        .iter()
        .map(|tab| tab.position)
        .collect::<Vec<_>>();
    for (kind, positions) in [
        ("deleted", fmt.tab_stops_to_delete.as_slice()),
        ("added", added_tab_positions.as_slice()),
    ] {
        let mut sorted = positions.to_vec();
        sorted.sort_unstable();
        if sorted
            .iter()
            .any(|position| !(-31_680..=31_680).contains(position))
        {
            return Err(WriteError::InvalidData(format!(
                "DOC {kind} tab position is outside -31680..=31680"
            )));
        }
        if sorted.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(WriteError::InvalidData(format!(
                "DOC {kind} tab positions contain a duplicate"
            )));
        }
    }
    if let Some(flow) = fmt.frame_text_flow
        && flow.backwards
        && !flow.vertical
    {
        return Err(WriteError::InvalidData(
            "DOC backwards frame text flow requires vertical flow".to_string(),
        ));
    }
    if let Some(height) = fmt.frame_height
        && (height.height_twips > 0x7FFF || (height.minimum && height.height_twips == 0))
    {
        return Err(WriteError::InvalidData(
            "DOC paragraph frame height is outside the WHeightAbs range".to_string(),
        ));
    }
    if let Some(drop_cap) = fmt.drop_cap
        && !(1..=10).contains(&drop_cap.lines)
    {
        return Err(WriteError::InvalidData(format!(
            "DOC drop-cap line count {} is outside 1..=10",
            drop_cap.lines
        )));
    }
    for (name, distance) in [
        ("horizontal", fmt.frame_horizontal_text_distance),
        ("vertical", fmt.frame_vertical_text_distance),
    ] {
        if let Some(distance) = distance
            && !(0..=31_680).contains(&distance)
        {
            return Err(WriteError::InvalidData(format!(
                "DOC {name} frame text distance {distance} is outside 0..=31680"
            )));
        }
    }
    for (name, offset) in [
        (
            "horizontal",
            match fmt.frame_horizontal_position {
                Some(FrameHorizontalPosition::Offset(value)) => Some(value),
                _ => None,
            },
        ),
        (
            "vertical",
            match fmt.frame_vertical_position {
                Some(FrameVerticalPosition::Offset(value)) => Some(value),
                _ => None,
            },
        ),
    ] {
        if let Some(offset) = offset
            && !(-31_679..=31_681).contains(&offset)
        {
            return Err(WriteError::InvalidData(format!(
                "DOC {name} frame offset {offset} is outside the plus-one range"
            )));
        }
        if let Some(offset) = offset {
            let stored = offset + 1;
            let is_special =
                matches!(stored, 0 | -4 | -8 | -12 | -16) || (name == "vertical" && stored == -20);
            if is_special {
                return Err(WriteError::InvalidData(format!(
                    "DOC {name} frame offset {offset} encodes a reserved alignment value"
                )));
            }
        }
    }
    if let Some(width) = fmt.frame_width
        && width > 31_680
    {
        return Err(WriteError::InvalidData(format!(
            "DOC paragraph frame width {width} exceeds 31680"
        )));
    }
    if fmt.table_terminating_paragraph == Some(true) && fmt.in_table != Some(true) {
        return Err(WriteError::InvalidData(
            "DOC table-terminating paragraph requires in_table=true".to_string(),
        ));
    }
    if fmt.frame_text_flow.is_some()
        && !matches!(fmt.frame_text_wrap, Some(wrap) if wrap != FrameTextWrap::Auto)
        && !matches!(fmt.frame_height, Some(height) if height.height_twips != 0)
        && fmt.frame_horizontal_position.is_none()
        && fmt.frame_vertical_position.is_none()
        && fmt.frame_width.is_none()
        && fmt.frame_anchor.is_none()
    {
        return Err(WriteError::InvalidData(
            "DOC frame text flow requires a non-default frame property".to_string(),
        ));
    }
    for (name, value) in [
        ("space_before_lines", fmt.space_before_lines),
        ("space_after_lines", fmt.space_after_lines),
    ] {
        if let Some(value) = value
            && !(-20..=31_680).contains(&value)
        {
            return Err(WriteError::InvalidData(format!(
                "DOC paragraph {name} value {value} is outside -20..=31680"
            )));
        }
    }
    for border in [
        fmt.borders.top,
        fmt.borders.left,
        fmt.borders.bottom,
        fmt.borders.right,
        fmt.borders.between,
        fmt.borders.bar,
    ]
    .into_iter()
    .flatten()
    {
        if border.spacing > 31 {
            return Err(WriteError::InvalidData(format!(
                "DOC paragraph border spacing {} exceeds 31 points",
                border.spacing
            )));
        }
    }

    let mut grp = if let Some(previous) = &fmt.preserved_properties_for_revision {
        let mut grp = build_revision_papx_grpprl(previous, revisions)?;
        grp.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
        grp.push(1);
        grp.extend_from_slice(&build_papx_grpprl(fmt));
        grp
    } else {
        build_papx_grpprl(fmt)
    };
    if let Some(revision) = &fmt.formatting_revision {
        let revisions = revisions.ok_or_else(|| {
            WriteError::InvalidData("DOC paragraph revision author was not indexed".to_string())
        })?;
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            WriteError::InvalidData("DOC paragraph revision author was not indexed".to_string())
        })?;
        grp.extend_from_slice(&SPRM_P_PROP_RMARK_CURRENT.to_le_bytes());
        grp.push(7);
        grp.push(1);
        grp.extend_from_slice(&author_index.to_le_bytes());
        grp.extend_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
    }
    if let Some(revision) = &fmt.numbering_revision {
        let revisions = revisions.ok_or_else(|| {
            WriteError::InvalidData("DOC numbering revision author was not indexed".to_string())
        })?;
        let author_index = revisions.indexes.get(&revision.author).ok_or_else(|| {
            WriteError::InvalidData("DOC numbering revision author was not indexed".to_string())
        })?;
        let units = revision.format_string.encode_utf16().collect::<Vec<_>>();
        if units.len() > 31
            || revision
                .placeholder_positions
                .iter()
                .any(|position| usize::from(*position) > units.len())
        {
            return Err(WriteError::InvalidData(
                "DOC numbering revision format or placeholder exceeds NumRM limits".to_string(),
            ));
        }
        let mut numrm = [0u8; 128];
        numrm[0] = u8::from(revision.was_numbered);
        numrm[2..4].copy_from_slice(&author_index.to_le_bytes());
        numrm[4..8].copy_from_slice(&pack_dttm(revision.timestamp)?.to_le_bytes());
        numrm[8..17].copy_from_slice(&revision.placeholder_positions);
        numrm[17..26].copy_from_slice(&revision.number_formats);
        for (index, number) in revision.numbers.iter().enumerate() {
            let offset = 28 + index * 4;
            numrm[offset..offset + 4].copy_from_slice(&number.to_le_bytes());
        }
        numrm[64..66].copy_from_slice(&(units.len() as u16).to_le_bytes());
        for (index, unit) in units.into_iter().enumerate() {
            let offset = 66 + index * 2;
            numrm[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
        }
        grp.extend_from_slice(&SPRM_P_NUM_RM.to_le_bytes());
        grp.push(128);
        grp.extend_from_slice(&numrm);
    }
    Ok(grp)
}

pub(super) fn append_table_depth_sprms(grp: &mut Vec<u8>) {
    grp.extend_from_slice(&SPRM_P_F_IN_TABLE.to_le_bytes());
    grp.push(1);
    grp.extend_from_slice(&SPRM_P_ITAP.to_le_bytes());
    grp.extend_from_slice(&1u32.to_le_bytes());
}

pub(super) fn build_table_row_papx_grpprl(
    formatting: &crate::writer::tap::TableRow,
) -> Result<Vec<u8>, WriteError> {
    let mut grp = Vec::new();
    append_table_depth_sprms(&mut grp);
    grp.extend_from_slice(&SPRM_P_F_TTP.to_le_bytes());
    grp.push(1);
    grp.extend_from_slice(
        &crate::writer::tap::generate_row_sprms(formatting)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?,
    );
    Ok(grp)
}
