//! High-level paragraph-property parsing and style cascading.

use super::model::{
    DropCap, DropCapType, FontAlignment, FrameAnchor, FrameHeight, FrameHorizontalAnchor,
    FrameHorizontalPosition, FrameTextFlow, FrameTextWrap, FrameVerticalAnchor,
    FrameVerticalPosition, Justification, LegacyBorderPosition, LegacyBorderStyle, LineSpacingType,
    ParagraphProperties, PhysicalJustification, TextBoxTightWrap,
};
use crate::package::{Error as PackageError, Result};
use crate::parts::{styles::StyleSheet, tap_parser::TapParser};
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::{
    SPRM_P_BRC_BAR, SPRM_P_BRC_BETWEEN, SPRM_P_BRC_BOTTOM, SPRM_P_BRC_LEFT, SPRM_P_BRC_RIGHT,
    SPRM_P_BRC_TOP, SPRM_P_CNF, SPRM_P_DXA_LEFT_2000, SPRM_P_DXA_LEFT1_2000, SPRM_P_DXA_RIGHT_2000,
    SPRM_P_DXC_LEFT, SPRM_P_DXC_LEFT1, SPRM_P_DXC_RIGHT, SPRM_P_DYL_AFTER, SPRM_P_DYL_BEFORE,
    SPRM_P_F_CONTEXTUAL_SPACING, SPRM_P_F_DYA_AFTER_AUTO, SPRM_P_F_DYA_BEFORE_AUTO,
    SPRM_P_F_MIRROR_INDENTS, SPRM_P_F_NO_ALLOW_OVERLAP, SPRM_P_F_OPEN_TCH, SPRM_P_IPGP,
    SPRM_P_ISTD, SPRM_P_ISTD_PERMUTE, SPRM_P_JC_LOGICAL, SPRM_P_NEST_2000, SPRM_P_RSID,
    SPRM_P_TTWO, SPRM_P_WALL, get_sprm_operation, get_sprm_type,
};
use litchi_core::binary::read_u16_le;

impl ParagraphProperties {
    /// Create a new `ParagraphProperties` with default values.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse paragraph properties from SPRM (Single Property Modifier) data.
    ///
    /// SPRMs are variable-length records that modify properties.
    ///
    /// Based on Apache POI's `ParagraphSprmUncompressor`.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        Self::from_sprm_context(grpprl, None)
    }

    pub(crate) fn from_sprm_with_stylesheet(
        grpprl: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        Self::from_sprm_context(grpprl, Some(stylesheet))
    }

    fn from_sprm_context(grpprl: &[u8], stylesheet: Option<&StyleSheet>) -> Result<Self> {
        let mut pap = Self::default();
        let sprms = parse_sprms(grpprl)?;
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != grpprl.len() {
            return Err(PackageError::Corrupted(
                "PAP grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }

        for sprm in &sprms {
            // Only process PAP SPRMs (type = 1)
            if get_sprm_type(sprm.opcode) == 1 {
                Self::apply_sprm(&mut pap, sprm)?;
            } else if get_sprm_type(sprm.opcode) == 5 {
                Self::apply_table_revision_sprm(&mut pap, sprm)?;
            }
        }

        if sprms.iter().any(|sprm| get_sprm_type(sprm.opcode) == 5) {
            let arena = bumpalo::Bump::new();
            let parser = TapParser::new(&arena);
            pap.table_properties = Some(if let Some(stylesheet) = stylesheet {
                parser.parse_tap_with_stylesheet(grpprl, stylesheet)?
            } else {
                parser.parse_tap(grpprl)?
            });
        }

        Ok(pap)
    }

    pub(crate) fn cascade_styles(
        initial_style_index: Option<u16>,
        direct_sprms: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        Self::cascade_styles_on_baseline(
            &Self::default(),
            initial_style_index,
            direct_sprms,
            stylesheet,
        )
    }

    /// Resolve the inherited paragraph-style state before direct PAPX SPRMs.
    ///
    /// Callers parsing source-ordered PAPX runs may reuse this immutable
    /// baseline for adjacent runs with the same initial style. Direct SPRMs
    /// must still be applied independently for every run.
    pub(crate) fn resolve_style_baseline(
        initial_style_index: Option<u16>,
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        Self::paragraph_style_on_baseline(&Self::default(), initial_style_index, stylesheet)
    }

    /// Apply one run's direct SPRMs to an already resolved initial style.
    ///
    /// A direct `sprmPIstd` or permutation still resolves from the document
    /// baseline, exactly as [`Self::cascade_styles`] does.
    pub(crate) fn cascade_styles_from_resolved_baseline(
        resolved_initial: &Self,
        direct_sprms: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        Self::apply_direct_sprms(
            resolved_initial.clone(),
            &Self::default(),
            direct_sprms,
            stylesheet,
        )
    }

    pub(crate) fn cascade_table_style(
        table_style_sprms: &[u8],
        initial_style_index: Option<u16>,
        direct_sprms: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        let table_baseline = Self::from_sprm(table_style_sprms)?;
        Self::cascade_styles_on_baseline(
            &table_baseline,
            initial_style_index,
            direct_sprms,
            stylesheet,
        )
    }

    fn cascade_styles_on_baseline(
        baseline: &Self,
        initial_style_index: Option<u16>,
        direct_sprms: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        let current = Self::paragraph_style_on_baseline(baseline, initial_style_index, stylesheet)?;
        Self::apply_direct_sprms(current, baseline, direct_sprms, stylesheet)
    }

    fn apply_direct_sprms(
        mut current: Self,
        style_baseline: &Self,
        direct_sprms: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        let sprms = parse_sprms(direct_sprms)?;
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != direct_sprms.len() {
            return Err(PackageError::Corrupted(
                "PAPX grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }
        for sprm in &sprms {
            if get_sprm_type(sprm.opcode) != 1 {
                continue;
            }
            let requested_style = if sprm.opcode == SPRM_P_ISTD {
                Some(sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmPIstd is missing its style index".to_string())
                })?)
            } else if sprm.opcode == SPRM_P_ISTD_PERMUTE {
                Self::permuted_style(sprm, current.style_index)?
            } else {
                None
            };
            if let Some(requested) = requested_style {
                let mut styled =
                    Self::paragraph_style_on_baseline(style_baseline, Some(requested), stylesheet)?;
                styled.style_index = Some(requested);
                Self::preserve_style_state(&current, &mut styled);
                current = styled;
            } else {
                Self::apply_sprm(&mut current, sprm)?;
            }
        }

        let table_state = Self::from_sprm_with_stylesheet(direct_sprms, stylesheet)?;
        current.table_properties = table_state.table_properties;
        current.has_table_formatting_revision = table_state.has_table_formatting_revision;
        current.table_formatting_revision_author_index =
            table_state.table_formatting_revision_author_index;
        current.table_formatting_revision_timestamp =
            table_state.table_formatting_revision_timestamp;
        current.table_properties_preserved_for_revision =
            table_state.table_properties_preserved_for_revision;
        Ok(current)
    }

    fn paragraph_style_on_baseline(
        baseline: &Self,
        style_index: Option<u16>,
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        let Some(requested) = style_index else {
            return Ok(baseline.clone());
        };
        let (effective, paragraph, _) = stylesheet.resolve_paragraph_style_sprms(requested)?;
        let mut styled = baseline.clone();
        for sprm in parse_sprms(&paragraph)? {
            Self::apply_sprm(&mut styled, &sprm)?;
        }
        styled.style_index = Some(requested);
        if effective.is_some() && (1..=9).contains(&requested) {
            styled.outline_level = Some((requested - 1) as u8);
        }
        Ok(styled)
    }

    fn preserve_style_state(previous: &Self, styled: &mut Self) {
        styled.in_table = previous.in_table;
        styled.is_table_row_end = previous.is_table_row_end;
        styled.table_nesting_level = previous.table_nesting_level;
        styled.inner_table_cell = previous.inner_table_cell;
        styled.inner_table_row_end = previous.inner_table_row_end;
        styled.is_table_cell_end = previous.is_table_cell_end;
        styled.open_table_cell_mark = previous.open_table_cell_mark;
        styled.table_properties = previous.table_properties.clone();
        styled.paragraph_group_id = previous.paragraph_group_id;
        styled.properties_preserved_for_revision = previous.properties_preserved_for_revision;
        styled.preserved_properties_for_revision =
            previous.preserved_properties_for_revision.clone();
        styled.revision_save_id = previous.revision_save_id;
        styled.has_formatting_revision = previous.has_formatting_revision;
        styled.formatting_revision_author_index = previous.formatting_revision_author_index;
        styled.formatting_revision_timestamp = previous.formatting_revision_timestamp;
        styled.numbering_revision_list_applied = previous.numbering_revision_list_applied;
        styled.numbering_revision = previous.numbering_revision.clone();
    }

    fn apply_table_revision_sprm(pap: &mut ParagraphProperties, sprm: &Sprm) -> Result<()> {
        match sprm.opcode {
            0xD667 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 7 {
                    return Err(PackageError::Corrupted(
                        "sprmTPropRMark operand must contain exactly 7 bytes".to_string(),
                    ));
                }
                pap.has_table_formatting_revision = Some(match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(PackageError::Corrupted(
                            "sprmTPropRMark must begin with a Boolean8 value".to_string(),
                        ));
                    },
                });
                let author = i16::from_le_bytes([operand[1], operand[2]]);
                pap.table_formatting_revision_author_index =
                    Some(u16::try_from(author).map_err(|_| {
                        PackageError::Corrupted(
                            "sprmTPropRMark author index is negative".to_string(),
                        )
                    })?);
                let timestamp =
                    u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
                crate::revision::decode_dttm(timestamp)?;
                pap.table_formatting_revision_timestamp = Some(timestamp);
            },
            0x3668 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 1 {
                    return Err(PackageError::Corrupted(
                        "sprmTWall operand must contain exactly 1 byte".to_string(),
                    ));
                }
                pap.table_properties_preserved_for_revision = match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(PackageError::Corrupted(
                            "sprmTWall must contain a Boolean8 value".to_string(),
                        ));
                    },
                };
            },
            _ => {},
        }
        Ok(())
    }
}

/// Apply a single SPRM operation to paragraph properties.
///
/// Based on Apache POI's `ParagraphSprmUncompressor.unCompressPAPOperation()`.
///
/// # Arguments
///
/// * `pap` - The paragraph properties to modify
/// * `sprm` - The SPRM operation to apply
// This is the specification-indexed SPRM dispatch table. Keeping opcode
// handling together makes overlap and precedence reviewable.
#[allow(
    clippy::cognitive_complexity,
    reason = "the parser keeps branch order aligned with the MS-DOC record grammar"
)]
impl ParagraphProperties {
    pub(super) fn apply_sprm(pap: &mut ParagraphProperties, sprm: &Sprm) -> Result<()> {
        match sprm.opcode {
            SPRM_P_DXC_RIGHT => {
                pap.indent_right_chars = Some(Self::required_i16(sprm, "sprmPDxcRight")?);
                return Ok(());
            },
            SPRM_P_DXC_LEFT => {
                pap.indent_left_chars = Some(Self::required_i16(sprm, "sprmPDxcLeft")?);
                return Ok(());
            },
            SPRM_P_DXC_LEFT1 => {
                pap.indent_first_line_chars = Some(Self::required_i16(sprm, "sprmPDxcLeft1")?);
                return Ok(());
            },
            SPRM_P_DYL_BEFORE => {
                pap.space_before_lines = Some(Self::line_hundredths(sprm, "sprmPDylBefore")?);
                return Ok(());
            },
            SPRM_P_DYL_AFTER => {
                pap.space_after_lines = Some(Self::line_hundredths(sprm, "sprmPDylAfter")?);
                return Ok(());
            },
            SPRM_P_F_OPEN_TCH => {
                pap.open_table_cell_mark = Self::strict_bool8(sprm, "sprmPFOpenTch")?;
                return Ok(());
            },
            SPRM_P_F_DYA_BEFORE_AUTO => {
                pap.space_before_auto = Self::strict_bool8(sprm, "sprmPFDyaBeforeAuto")?;
                return Ok(());
            },
            SPRM_P_F_DYA_AFTER_AUTO => {
                pap.space_after_auto = Self::strict_bool8(sprm, "sprmPFDyaAfterAuto")?;
                return Ok(());
            },
            SPRM_P_DXA_RIGHT_2000 => {
                pap.indent_right = Some(i32::from(Self::xas(sprm, "sprmPDxaRight")?));
                return Ok(());
            },
            SPRM_P_DXA_LEFT_2000 => {
                pap.indent_left = Some(i32::from(Self::xas(sprm, "sprmPDxaLeft")?));
                return Ok(());
            },
            SPRM_P_NEST_2000 => {
                let delta = i32::from(Self::xas(sprm, "sprmPNest")?);
                pap.indent_left = Some(pap.indent_left.unwrap_or(0) + delta);
                return Ok(());
            },
            SPRM_P_DXA_LEFT1_2000 => {
                pap.indent_first_line = Some(i32::from(Self::xas(sprm, "sprmPDxaLeft1")?));
                return Ok(());
            },
            SPRM_P_JC_LOGICAL => {
                let code = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPJc is missing its justification".to_string())
                })?;
                pap.justification = Justification::try_from(code).map_err(|invalid| {
                    PackageError::Corrupted(format!(
                        "sprmPJc has invalid logical justification {invalid}"
                    ))
                })?;
                pap.physical_justification = None;
                return Ok(());
            },
            SPRM_P_BRC_TOP => {
                pap.borders.top = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_LEFT => {
                pap.borders.left = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_BOTTOM => {
                pap.borders.bottom = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_RIGHT => {
                pap.borders.right = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_BETWEEN => {
                pap.borders.between = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_BAR => {
                pap.borders.bar = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_F_NO_ALLOW_OVERLAP => {
                pap.no_allow_overlap = Self::strict_bool8(sprm, "sprmPFNoAllowOverlap")?;
                return Ok(());
            },
            SPRM_P_WALL => {
                let enabled = Self::strict_bool8(sprm, "sprmPWall")?;
                pap.preserved_properties_for_revision = if enabled {
                    let mut previous = pap.clone();
                    previous.properties_preserved_for_revision = false;
                    previous.preserved_properties_for_revision = None;
                    Some(Box::new(previous))
                } else {
                    None
                };
                pap.properties_preserved_for_revision = enabled;
                return Ok(());
            },
            SPRM_P_IPGP => {
                let group_id = sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmPIpgp is missing its PGPInfo index".to_string())
                })?;
                if group_id == 0 {
                    return Err(PackageError::Corrupted(
                        "sprmPIpgp must contain a nonzero PGPInfo index".to_string(),
                    ));
                }
                pap.paragraph_group_id = Some(group_id);
                return Ok(());
            },
            SPRM_P_CNF => {
                pap.conditional_formats
                    .push(Self::parse_conditional_formatting(sprm)?);
                return Ok(());
            },
            SPRM_P_RSID => {
                pap.revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    PackageError::Corrupted("sprmPRsid is missing its revision save ID".to_string())
                })?);
                return Ok(());
            },
            SPRM_P_F_CONTEXTUAL_SPACING => {
                pap.contextual_spacing = Self::strict_bool8(sprm, "sprmPFContextualSpacing")?;
                return Ok(());
            },
            SPRM_P_F_MIRROR_INDENTS => {
                pap.mirror_indents = Self::strict_bool8(sprm, "sprmPFMirrorIndents")?;
                return Ok(());
            },
            SPRM_P_TTWO => {
                let tight_wrap = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPTtwo is missing its tight-wrap mode".to_string())
                })?;
                pap.text_box_tight_wrap =
                    Some(TextBoxTightWrap::try_from(tight_wrap).map_err(|invalid| {
                        PackageError::Corrupted(format!(
                            "sprmPTtwo has invalid tight-wrap mode {invalid}"
                        ))
                    })?);
                return Ok(());
            },
            _ => {},
        }
        let operation = get_sprm_operation(sprm.opcode);

        match operation {
            // Operation 0x00: sprmPIstd - Paragraph style
            0x00 => {
                if let Some(istd) = sprm.operand_word() {
                    pap.style_index = Some(istd);
                }
            },
            // Operation 0x01: sprmPIstdPermute - Style permutation
            0x01 => {
                if let Some(style) = Self::permuted_style(sprm, pap.style_index)? {
                    pap.style_index = Some(style);
                    if (1..=9).contains(&style) {
                        pap.outline_level = Some((style - 1) as u8);
                    }
                }
            },
            // Operation 0x02: sprmPIncLvl - Increment outline level
            0x02 => {
                let delta = i16::from(sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPIncLvl is missing its signed offset".to_string())
                })? as i8);
                if let Some(style @ 1..=9) = pap.style_index {
                    let style = (style as i16 + delta).clamp(1, 9) as u16;
                    pap.style_index = Some(style);
                    pap.outline_level = Some((style - 1) as u8);
                } else if let Some(level @ 0..=8) = pap.outline_level {
                    pap.outline_level = Some((i16::from(level) + delta).clamp(0, 9) as u8);
                }
            },
            // Operation 0x03: sprmPJc - Paragraph justification
            0x03 => {
                let jc = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPJc80 is missing its justification".to_string())
                })?;
                let physical = PhysicalJustification::try_from(jc).map_err(|invalid| {
                    PackageError::Corrupted(format!(
                        "sprmPJc80 has invalid physical justification {invalid}"
                    ))
                })?;
                pap.physical_justification = Some(physical);
                pap.justification = physical.normalized();
            },
            // Operation 0x04: sprmPFSideBySide - Side-by-side
            0x04 => {
                pap.side_by_side = Self::strict_bool8(sprm, "sprmPFSideBySide")?;
            },
            // Operation 0x05: sprmPFKeep - Keep paragraph intact
            0x05 => {
                pap.keep_on_page = Self::strict_bool8(sprm, "sprmPFKeep")?;
            },
            // Operation 0x06: sprmPFKeepFollow - Keep with next
            0x06 => {
                pap.keep_with_next = Self::strict_bool8(sprm, "sprmPFKeepFollow")?;
            },
            // Operation 0x07: sprmPFPageBreakBefore - Page break before
            0x07 => {
                pap.page_break_before = Self::strict_bool8(sprm, "sprmPFPageBreakBefore")?;
            },
            // Operation 0x08: sprmPBrcl - Border location
            0x08 => {
                let value = *sprm.operand_bytes().first().ok_or_else(|| {
                    PackageError::Corrupted("sprmPBrcl is missing its line style".to_string())
                })?;
                pap.legacy_border_style =
                    Some(LegacyBorderStyle::try_from(value).map_err(|invalid| {
                        PackageError::Corrupted(format!(
                            "sprmPBrcl has invalid legacy border style {invalid}"
                        ))
                    })?);
            },
            // Operation 0x09: sprmPBrcp - Border position
            0x09 => {
                let value = *sprm.operand_bytes().first().ok_or_else(|| {
                    PackageError::Corrupted("sprmPBrcp is missing its placement".to_string())
                })?;
                pap.legacy_border_position =
                    Some(LegacyBorderPosition::try_from(value).map_err(|invalid| {
                        PackageError::Corrupted(format!(
                            "sprmPBrcp has invalid legacy border placement {invalid}"
                        ))
                    })?);
            },
            // Operation 0x0A: sprmPIlvl - List level
            0x0A => {
                let ilvl = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPIlvl is missing its list level".to_string())
                })?;
                if ilvl > 8 && ilvl != 0x0C {
                    return Err(PackageError::Corrupted(format!(
                        "sprmPIlvl has invalid list level {ilvl}"
                    )));
                }
                pap.list_level = Some(ilvl);
            },
            // Operation 0x0B: sprmPIlfo - List format override
            0x0B => {
                let raw = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmPIlfo is missing its list override".to_string())
                })?;
                if (0x07FF..=0xF800).contains(&raw) {
                    return Err(PackageError::Corrupted(format!(
                        "sprmPIlfo has reserved list override {raw:#06x}"
                    )));
                }
                pap.list_format_override = Some(i16::from_le_bytes(raw.to_le_bytes()));
            },
            // Operation 0x0C: sprmPFNoLineNumb - No line numbering
            0x0C => {
                pap.no_line_numbering = Self::strict_bool8(sprm, "sprmPFNoLineNumb")?;
            },
            // Operation 0x0D: sprmPChgTabsPapx - Tab stops
            0x0D => {
                Self::handle_tabs(pap, sprm, false)?;
            },
            // Operation 0x0E: sprmPDxaRight - Right indent
            0x0E => {
                pap.indent_right = Some(i32::from(Self::xas(sprm, "sprmPDxaRight")?));
            },
            // Operation 0x0F: sprmPDxaLeft - Left indent
            0x0F => {
                pap.indent_left = Some(i32::from(Self::xas(sprm, "sprmPDxaLeft")?));
            },
            // Operation 0x10: sprmPNest - Nested indent
            0x10 => {
                let delta = i32::from(Self::xas(sprm, "sprmPNest")?);
                pap.indent_left = Some(pap.indent_left.unwrap_or(0) + delta);
            },
            // Operation 0x11: sprmPDxaLeft1 - First line indent
            0x11 => {
                pap.indent_first_line = Some(i32::from(Self::xas(sprm, "sprmPDxaLeft1")?));
            },
            // Operation 0x12: sprmPDyaLine - Line spacing
            0x12 => {
                let bytes = sprm.operand_bytes();
                if bytes.len() != 4 {
                    return Err(PackageError::Corrupted(
                        "sprmPDyaLine must contain exactly 4 bytes".to_string(),
                    ));
                }
                let raw_dya = read_u16_le(bytes, 0).map_err(|error| {
                    PackageError::Corrupted(format!("invalid sprmPDyaLine spacing: {error}"))
                })?;
                if raw_dya > 0x7BC0 && raw_dya < 0x8440 {
                    return Err(PackageError::Corrupted(format!(
                        "sprmPDyaLine value {raw_dya:#06x} is outside the LSPD range"
                    )));
                }
                let dya_line = i16::from_le_bytes(raw_dya.to_le_bytes());
                let f_mult = read_u16_le(bytes, 2).map_err(|error| {
                    PackageError::Corrupted(format!("invalid sprmPDyaLine mode: {error}"))
                })?;
                if f_mult > 1 {
                    return Err(PackageError::Corrupted(format!(
                        "sprmPDyaLine has invalid multiple-line flag {f_mult}"
                    )));
                }
                pap.line_spacing = Some(dya_line);
                pap.line_spacing_type = if f_mult == 1 {
                    match dya_line {
                        240 => LineSpacingType::Single,
                        360 => LineSpacingType::OnePointFive,
                        480 => LineSpacingType::Double,
                        _ => LineSpacingType::Multiple,
                    }
                } else if dya_line > 0 {
                    LineSpacingType::AtLeast
                } else {
                    LineSpacingType::Exactly
                };
            },
            // Operation 0x13: sprmPDyaBefore - Space before
            0x13 => {
                pap.space_before = Some(Self::unsigned_twips(sprm, "sprmPDyaBefore")?);
            },
            // Operation 0x14: sprmPDyaAfter - Space after
            0x14 => {
                pap.space_after = Some(Self::unsigned_twips(sprm, "sprmPDyaAfter")?);
            },
            // Operation 0x15: sprmPChgTabs - Change tabs (fast saved)
            0x15 => {
                Self::handle_tabs(pap, sprm, true)?;
            },
            // Operation 0x16: sprmPFInTable - In table flag
            0x16 => {
                pap.in_table = Self::strict_bool8(sprm, "sprmPFInTable")?;
            },
            // Operation 0x17: sprmPFTtp - Table row end
            0x17 => {
                let value = Self::strict_bool8(sprm, "sprmPFTtp")?;
                if value && !pap.in_table {
                    return Err(PackageError::Corrupted(
                        "sprmPFTtp requires sprmPFInTable to be enabled".to_string(),
                    ));
                }
                pap.is_table_row_end = value;
            },
            // Operation 0x18: sprmPDxaAbs - Absolute horizontal position
            0x18 => {
                let raw = Self::required_i16(sprm, "sprmPDxaAbs")?;
                pap.frame_horizontal_position = Some(match raw {
                    0 => FrameHorizontalPosition::Left,
                    -4 => FrameHorizontalPosition::Center,
                    -8 => FrameHorizontalPosition::Right,
                    -12 => FrameHorizontalPosition::Inside,
                    -16 => FrameHorizontalPosition::Outside,
                    value if (-31_678..=31_682).contains(&value) => {
                        FrameHorizontalPosition::Offset(value - 1)
                    },
                    invalid => {
                        return Err(PackageError::Corrupted(format!(
                            "sprmPDxaAbs stored position {invalid} is outside XAS_plusOne"
                        )));
                    },
                });
            },
            // Operation 0x19: sprmPDyaAbs - Absolute vertical position
            0x19 => {
                let raw = Self::required_i16(sprm, "sprmPDyaAbs")?;
                pap.frame_vertical_position = Some(match raw {
                    0 => FrameVerticalPosition::Inline,
                    -4 => FrameVerticalPosition::Top,
                    -8 => FrameVerticalPosition::Center,
                    -12 => FrameVerticalPosition::Bottom,
                    -16 => FrameVerticalPosition::Inside,
                    -20 => FrameVerticalPosition::Outside,
                    value if (-31_678..=31_682).contains(&value) => {
                        FrameVerticalPosition::Offset(value - 1)
                    },
                    invalid => {
                        return Err(PackageError::Corrupted(format!(
                            "sprmPDyaAbs stored position {invalid} is outside YAS_plusOne"
                        )));
                    },
                });
            },
            // Operation 0x1A: sprmPDxaWidth - Absolute width
            0x1A => {
                let width = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmPDxaWidth is missing its frame width".to_string())
                })?;
                if width > 31_680 {
                    return Err(PackageError::Corrupted(format!(
                        "sprmPDxaWidth value {width} exceeds 31680"
                    )));
                }
                pap.frame_width = Some(width);
            },
            // Operation 0x1B: sprmPPc - Positioning code
            0x1B => {
                let value = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPPc is missing its anchor code".to_string())
                })?;
                if value & 0x0F != 0 {
                    return Err(PackageError::Corrupted(
                        "sprmPPc padding bits must be zero".to_string(),
                    ));
                }
                pap.frame_anchor = Some(FrameAnchor {
                    vertical: match (value >> 4) & 0x03 {
                        0 => FrameVerticalAnchor::Margin,
                        1 => FrameVerticalAnchor::Page,
                        2 => FrameVerticalAnchor::Paragraph,
                        _ => FrameVerticalAnchor::None,
                    },
                    horizontal: match value >> 6 {
                        0 => FrameHorizontalAnchor::Column,
                        1 => FrameHorizontalAnchor::Margin,
                        2 => FrameHorizontalAnchor::Page,
                        _ => FrameHorizontalAnchor::None,
                    },
                });
            },
            // Operations 0x1C-0x21: Old border formats (Word 6.0)
            0x1C => pap.borders.top = Self::parse_border10(sprm)?,
            0x1D => pap.borders.left = Self::parse_border10(sprm)?,
            0x1E => pap.borders.bottom = Self::parse_border10(sprm)?,
            0x1F => pap.borders.right = Self::parse_border10(sprm)?,
            0x20 => pap.borders.between = Self::parse_border10(sprm)?,
            0x21 => pap.borders.bar = Self::parse_border10(sprm)?,
            // Operation 0x22: sprmPDxaFromText10 - Distance from text (Word 6.0)
            0x22 => {
                pap.dxa_from_text = Some(Self::nonnegative_distance(sprm, "sprmPDxaFromText10")?);
            },
            // Operation 0x23: sprmPWr - Text wrapping
            0x23 => {
                let value = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPWr is missing its wrap mode".to_string())
                })?;
                pap.text_wrap = Some(FrameTextWrap::try_from(value).map_err(|invalid| {
                    PackageError::Corrupted(format!("sprmPWr has invalid wrap mode {invalid}"))
                })?);
            },
            // Operations 0x24-0x29: Word 97 Brc80 borders
            0x24 => pap.borders.top = Self::parse_border80(sprm)?,
            0x25 => pap.borders.left = Self::parse_border80(sprm)?,
            0x26 => pap.borders.bottom = Self::parse_border80(sprm)?,
            0x27 => pap.borders.right = Self::parse_border80(sprm)?,
            0x28 => pap.borders.between = Self::parse_border80(sprm)?,
            0x29 => pap.borders.bar = Self::parse_border80(sprm)?,
            // Operation 0x2A: sprmPFNoAutoHyph - No auto hyphenation
            0x2A => {
                pap.no_auto_hyph = Self::strict_bool8(sprm, "sprmPFNoAutoHyph")?;
            },
            // Operation 0x2B: sprmPWHeightAbs - Frame height
            0x2B => {
                let value = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted(
                        "sprmPWHeightAbs is missing its frame height".to_string(),
                    )
                })?;
                let frame_height = FrameHeight {
                    height_twips: value & 0x7FFF,
                    minimum: value & 0x8000 != 0,
                };
                if frame_height.minimum && frame_height.height_twips == 0 {
                    return Err(PackageError::Corrupted(
                        "sprmPWHeightAbs minimum frame height cannot be zero".to_string(),
                    ));
                }
                pap.frame_height = Some(frame_height);
            },
            // Operation 0x2C: sprmPDcs - Drop cap
            0x2C => {
                let value = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted(
                        "sprmPDcs is missing its drop-cap descriptor".to_string(),
                    )
                })?;
                if value == 0 {
                    pap.drop_cap = None;
                } else {
                    let kind = match value & 0x07 {
                        1 => DropCapType::Regular,
                        2 => DropCapType::Margin,
                        invalid => {
                            return Err(PackageError::Corrupted(format!(
                                "sprmPDcs has invalid drop-cap type {invalid}"
                            )));
                        },
                    };
                    let lines = ((value >> 3) & 0x1F) as u8;
                    if !(1..=10).contains(&lines) {
                        return Err(PackageError::Corrupted(format!(
                            "sprmPDcs has invalid drop-cap line count {lines}"
                        )));
                    }
                    pap.drop_cap = Some(DropCap { kind, lines });
                }
            },
            // Operation 0x2D: sprmPShd80 - Shading (Word 97-2000)
            0x2D => {
                let shd = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmPShd80 is missing its Shd80".to_string())
                })?;
                pap.shading = Self::parse_shd80(shd)?;
            },
            // Operation 0x2E: sprmPDyaFromText - Vertical distance from text
            0x2E => {
                pap.dya_from_text = Some(Self::nonnegative_distance(sprm, "sprmPDyaFromText")?);
            },
            // Operation 0x2F: sprmPDxaFromText - Horizontal distance from text
            0x2F => {
                pap.dxa_from_text = Some(Self::nonnegative_distance(sprm, "sprmPDxaFromText")?);
            },
            // Operation 0x30: sprmPFLocked - Locked paragraph
            0x30 => {
                pap.locked = Self::strict_bool8(sprm, "sprmPFLocked")?;
            },
            // Operation 0x31: sprmPFWidowControl - Widow/orphan control
            0x31 => {
                pap.widow_control = Self::strict_bool8(sprm, "sprmPFWidowControl")?;
            },
            // Operation 0x33: sprmPFKinsoku - Kinsoku
            0x33 => {
                pap.kinsoku = Self::strict_bool8(sprm, "sprmPFKinsoku")?;
            },
            // Operation 0x34: sprmPFWordWrap - Word wrap
            0x34 => {
                pap.word_wrap = Self::strict_bool8(sprm, "sprmPFWordWrap")?;
            },
            // Operation 0x35: sprmPFOverflowPunct - Overflow punctuation
            0x35 => {
                pap.overflow_punct = Self::strict_bool8(sprm, "sprmPFOverflowPunct")?;
            },
            // Operation 0x36: sprmPFTopLinePunct - Top line punctuation
            0x36 => {
                pap.top_line_punct = Self::strict_bool8(sprm, "sprmPFTopLinePunct")?;
            },
            // Operation 0x37: sprmPFAutoSpaceDE - Auto space DE
            0x37 => {
                pap.auto_space_de = Self::strict_bool8(sprm, "sprmPFAutoSpaceDE")?;
            },
            // Operation 0x38: sprmPFAutoSpaceDN - Auto space DN
            0x38 => {
                pap.auto_space_dn = Self::strict_bool8(sprm, "sprmPFAutoSpaceDN")?;
            },
            // Operation 0x39: sprmPWAlignFont - Font alignment
            0x39 => {
                let value = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmPWAlignFont is missing its alignment".to_string())
                })?;
                pap.font_align = Some(FontAlignment::try_from(value).map_err(|invalid| {
                    PackageError::Corrupted(format!(
                        "sprmPWAlignFont has invalid alignment {invalid}"
                    ))
                })?);
            },
            // Operation 0x3A: sprmPFrameTextFlow - Frame text flow
            0x3A => {
                let value = sprm.operand_word().ok_or_else(|| {
                    PackageError::Corrupted("sprmPFrameTextFlow is missing its flags".to_string())
                })?;
                let flow = FrameTextFlow {
                    vertical: value & 1 != 0,
                    backwards: value & 2 != 0,
                    rotate_font: value & 4 != 0,
                };
                if flow.backwards && !flow.vertical {
                    return Err(PackageError::Corrupted(
                        "sprmPFrameTextFlow backwards flow requires vertical flow".to_string(),
                    ));
                }
                pap.frame_text_flow = Some(flow);
            },
            // Operation 0x3B: sprmPISnapBaseLine - Snap to baseline
            0x3B => {
                // Not commonly used
            },
            // Operation 0x3E: sprmPAnld - Autonumber list data
            0x3E => {
                pap.legacy_autonumbering = Some(Self::parse_legacy_autonumbering(sprm)?);
            },
            // Versioned sprmPPropRMark property revision marks.
            0x3F | 0x65 | 0x6F => Self::apply_property_revision(pap, sprm)?,
            // Operation 0x40: sprmPOutLvl - Outline level
            0x40 => {
                let level = sprm.operand_byte().ok_or_else(|| {
                    PackageError::Corrupted("sprmPOutLvl is missing its outline level".to_string())
                })?;
                if level > 9 {
                    return Err(PackageError::Corrupted(format!(
                        "sprmPOutLvl has invalid outline level {level}"
                    )));
                }
                pap.outline_level = Some(level);
            },
            // Operation 0x41: sprmPFBiDi - Bi-directional paragraph
            0x41 => {
                pap.bi_directional = Self::strict_bool8(sprm, "sprmPFBiDi")?;
            },
            // Operation 0x43: sprmPFNumRMIns - Numbering revision insert
            0x43 => {
                pap.numbering_revision_list_applied =
                    Some(Self::strict_bool8(sprm, "sprmPFNumRMIns")?);
            },
            // Operation 0x44: sprmPCrLf - CR/LF
            0x44 => {
                // Not commonly used
            },
            // Operation 0x45: sprmPNumRM - Numbering revision mark
            0x45 => pap.numbering_revision = Some(Self::parse_numbering_revision(sprm)?),
            // Operation 0x47: sprmPFUsePgsuSettings - Use page setup settings
            0x47 => {
                pap.use_page_setup_settings =
                    Some(Self::strict_bool8(sprm, "sprmPFUsePgsuSettings")?);
            },
            // Operation 0x48: sprmPFAdjustRight - Adjust right
            0x48 => {
                pap.adjust_right_indent = Some(Self::strict_bool8(sprm, "sprmPFAdjustRight")?);
            },
            // Operation 0x49: sprmPItap - Table nesting level
            0x49 => {
                let depth = Self::required_i32(sprm, "sprmPItap")?;
                if depth < 0 {
                    return Err(PackageError::Corrupted(
                        "sprmPItap table depth must be non-negative".to_string(),
                    ));
                }
                pap.table_nesting_level = depth;
            },
            // Operation 0x4A: sprmPDtap - Table nesting delta
            0x4A => {
                let delta = Self::required_i32(sprm, "sprmPDtap")?;
                let depth = pap.table_nesting_level.checked_add(delta).ok_or_else(|| {
                    PackageError::Corrupted("sprmPDtap table depth overflowed".to_string())
                })?;
                if depth < 0 {
                    return Err(PackageError::Corrupted(
                        "sprmPDtap produced a negative table depth".to_string(),
                    ));
                }
                pap.table_nesting_level = depth;
            },
            // Operation 0x4B: sprmPFInnerTableCell - Inner table cell
            0x4B => {
                let value = Self::strict_bool8(sprm, "sprmPFInnerTableCell")?;
                if value && pap.table_nesting_level <= 1 {
                    return Err(PackageError::Corrupted(
                        "sprmPFInnerTableCell requires table depth greater than 1".to_string(),
                    ));
                }
                pap.inner_table_cell = value;
            },
            // Operation 0x4C: sprmPFInnerTtp - Inner table row end
            0x4C => {
                let value = Self::strict_bool8(sprm, "sprmPFInnerTtp")?;
                if value && pap.table_nesting_level <= 1 {
                    return Err(PackageError::Corrupted(
                        "sprmPFInnerTtp requires table depth greater than 1".to_string(),
                    ));
                }
                pap.inner_table_row_end = value;
            },
            // Operation 0x4D: sprmPShd - Shading (Word 2002+)
            0x4D => pap.shading = Self::parse_shading_descriptor(sprm)?,
            // Operation 0x67: sprmPRsid - Revision save ID
            0x67 => {
                // Revision save ID - not commonly used
            },
            // Default: Unknown or unsupported SPRM
            _ => {
                // Silently ignore unknown SPRMs
            },
        }
        Ok(())
    }
}
