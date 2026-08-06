//! Strict low-level codecs for paragraph-property SPRM operands.

use super::model::*;
use crate::package::{Error as PackageError, Result};
use crate::parts::numbering::NumberFormat;
use crate::parts::tap::TableStyleCondition;
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::*;
use litchi_core::binary::{read_i16_le, read_u16_le};

impl ParagraphProperties {
    pub(super) fn parse_legacy_autonumbering(sprm: &Sprm) -> Result<LegacyAutoNumbering> {
        let data = sprm.operand_bytes();
        if !matches!(data.len(), 52 | 84) {
            return Err(PackageError::Corrupted(format!(
                "sprmPAnld has {} bytes; expected 52 or 84",
                data.len()
            )));
        }
        let number_format = NumberFormat::try_from(data[0]).map_err(|invalid| {
            PackageError::Corrupted(format!("sprmPAnld has invalid MSONFC {invalid:#04x}"))
        })?;
        let alignment = match data[3] & 0x03 {
            0 => AutoNumberAlignment::Left,
            1 => AutoNumberAlignment::Center,
            2 => AutoNumberAlignment::Right,
            _ => AutoNumberAlignment::Justified,
        };
        let bool8 = |value, name| match value {
            0 => Ok(false),
            1 => Ok(true),
            invalid => Err(PackageError::Corrupted(format!(
                "sprmPAnld {name} has invalid Boolean8 value {invalid}"
            ))),
        };
        let color_index = data[5] >> 3;
        if color_index > 16 {
            return Err(PackageError::Corrupted(format!(
                "sprmPAnld has invalid color index {color_index}"
            )));
        }
        let indent_twips = read_i16_le(data, 12).map_err(|error| {
            PackageError::Corrupted(format!("sprmPAnld has invalid indent: {error}"))
        })?;
        if !(-31_680..=31_680).contains(&indent_twips) {
            return Err(PackageError::Corrupted(format!(
                "sprmPAnld indent {indent_twips} is outside -31680..=31680"
            )));
        }
        let space_twips = read_u16_le(data, 14).map_err(|error| {
            PackageError::Corrupted(format!("sprmPAnld has invalid spacing: {error}"))
        })?;
        if space_twips > 31_680 {
            return Err(PackageError::Corrupted(format!(
                "sprmPAnld spacing {space_twips} exceeds 31680"
            )));
        }
        if data[19] != 0 {
            return Err(PackageError::Corrupted(
                "sprmPAnld reserved flag byte must be zero".to_string(),
            ));
        }
        let before = usize::from(data[1]);
        let after = usize::from(data[2]);
        let text_count = before + after;
        let capacity = (data.len() - 20) / 2;
        if text_count > capacity {
            return Err(PackageError::Corrupted(format!(
                "sprmPAnld requests {text_count} label characters; capacity is {capacity}"
            )));
        }
        let mut text = Vec::with_capacity(text_count);
        for index in 0..text_count {
            text.push(read_u16_le(data, 20 + index * 2).map_err(|error| {
                PackageError::Corrupted(format!("sprmPAnld has invalid label text: {error}"))
            })?);
        }
        let prefix = String::from_utf16(&text[..before]).map_err(|_| {
            PackageError::Corrupted("sprmPAnld prefix is not valid UTF-16".to_string())
        })?;
        let suffix = String::from_utf16(&text[before..]).map_err(|_| {
            PackageError::Corrupted("sprmPAnld suffix is not valid UTF-16".to_string())
        })?;

        Ok(LegacyAutoNumbering {
            number_format,
            alignment,
            include_previous_levels: data[3] & 0x04 != 0,
            hanging_indent: data[3] & 0x08 != 0,
            set_bold: data[3] & 0x10 != 0,
            set_italic: data[3] & 0x20 != 0,
            set_small_caps: data[3] & 0x40 != 0,
            set_caps: data[3] & 0x80 != 0,
            set_strike: data[4] & 0x01 != 0,
            set_underline: data[4] & 0x02 != 0,
            prefix_space: data[4] & 0x04 != 0,
            bold: data[4] & 0x08 != 0,
            italic: data[4] & 0x10 != 0,
            small_caps: data[4] & 0x20 != 0,
            caps: data[4] & 0x40 != 0,
            strike: data[4] & 0x80 != 0,
            underline: data[5] & 0x07,
            color_index,
            font_index: read_u16_le(data, 6).map_err(|error| {
                PackageError::Corrupted(format!("sprmPAnld has invalid font index: {error}"))
            })?,
            font_size_half_points: read_u16_le(data, 8).map_err(|error| {
                PackageError::Corrupted(format!("sprmPAnld has invalid font size: {error}"))
            })?,
            start_at: read_u16_le(data, 10).map_err(|error| {
                PackageError::Corrupted(format!("sprmPAnld has invalid start value: {error}"))
            })?,
            indent_twips,
            space_twips,
            number_once_per_cell: bool8(data[16], "fNumber1")?,
            number_across_cells: bool8(data[17], "fNumberAcross")?,
            restart_each_section: bool8(data[18], "fRestartHdn")?,
            prefix,
            suffix,
        })
    }

    pub(super) fn parse_conditional_formatting(
        sprm: &Sprm,
    ) -> Result<ParagraphConditionalFormatting> {
        let operand = sprm.operand_bytes();
        if operand.len() < 2 {
            return Err(PackageError::Corrupted(
                "sprmPCnf must contain a 2-byte condition".to_string(),
            ));
        }
        let code = read_u16_le(operand, 0).map_err(|error| {
            PackageError::Corrupted(format!("sprmPCnf has an invalid condition: {error}"))
        })?;
        let condition = TableStyleCondition::from_code(code).ok_or_else(|| {
            PackageError::Corrupted(format!("sprmPCnf contains invalid condition {code:#06x}"))
        })?;
        let raw_grpprl = operand[2..].to_vec();
        let nested = parse_sprms(&raw_grpprl)?;
        let consumed = nested
            .last()
            .map_or(0, |nested| nested.offset + nested.size);
        if consumed != raw_grpprl.len() {
            return Err(PackageError::Corrupted(
                "sprmPCnf nested grpprl is truncated".to_string(),
            ));
        }
        if nested.iter().any(|nested| nested.opcode == SPRM_P_CNF) {
            return Err(PackageError::Corrupted(
                "sprmPCnf cannot be nested inside another sprmPCnf".to_string(),
            ));
        }
        if nested
            .iter()
            .any(|nested| get_sprm_type(nested.opcode) != 1)
        {
            return Err(PackageError::Corrupted(
                "sprmPCnf can contain only paragraph SPRMs".to_string(),
            ));
        }
        let properties = Box::new(Self::from_sprm(&raw_grpprl)?);
        Ok(ParagraphConditionalFormatting {
            condition,
            properties,
            raw_grpprl,
        })
    }

    pub(super) fn apply_property_revision(
        pap: &mut ParagraphProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 7 {
            return Err(PackageError::Corrupted(
                "sprmPPropRMark operand must contain exactly 7 bytes".to_string(),
            ));
        }
        pap.has_formatting_revision = Some(match operand[0] {
            0 => false,
            1 => true,
            _ => {
                return Err(PackageError::Corrupted(
                    "sprmPPropRMark must begin with a Boolean8 value".to_string(),
                ));
            },
        });
        let author = i16::from_le_bytes([operand[1], operand[2]]);
        pap.formatting_revision_author_index = Some(u16::try_from(author).map_err(|_| {
            PackageError::Corrupted("sprmPPropRMark author index is negative".to_string())
        })?);
        let timestamp = u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
        crate::revision::decode_dttm(timestamp)?;
        pap.formatting_revision_timestamp = Some(timestamp);
        Ok(())
    }

    pub(super) fn permuted_style(sprm: &Sprm, current: Option<u16>) -> Result<Option<u16>> {
        let operand = sprm.operand_bytes();
        if operand.len() < 7 {
            return Err(PackageError::Corrupted(
                "sprmPIstdPermute SPPOperand is too short".to_string(),
            ));
        }
        if operand[0] != 0 {
            return Err(PackageError::Corrupted(
                "sprmPIstdPermute fLong must be zero".to_string(),
            ));
        }
        let first = read_u16_le(operand, 1).map_err(|error| {
            PackageError::Corrupted(format!("invalid sprmPIstdPermute first style: {error}"))
        })?;
        let last = read_u16_le(operand, 3).map_err(|error| {
            PackageError::Corrupted(format!("invalid sprmPIstdPermute last style: {error}"))
        })?;
        if last < first {
            return Err(PackageError::Corrupted(
                "sprmPIstdPermute last style precedes its first style".to_string(),
            ));
        }
        let count = usize::from(last - first) + 1;
        let expected = 5 + count * 2;
        if operand.len() != expected {
            return Err(PackageError::Corrupted(format!(
                "sprmPIstdPermute has {} bytes; expected {expected}",
                operand.len()
            )));
        }
        let Some(current) = current.filter(|style| (first..=last).contains(style)) else {
            return Ok(None);
        };
        let offset = 5 + usize::from(current - first) * 2;
        read_u16_le(operand, offset).map(Some).map_err(|error| {
            PackageError::Corrupted(format!("invalid sprmPIstdPermute mapped style: {error}"))
        })
    }

    pub(super) fn strict_bool8(sprm: &Sprm, name: &str) -> Result<bool> {
        match sprm.operand_byte() {
            Some(0) => Ok(false),
            Some(1) => Ok(true),
            _ => Err(PackageError::Corrupted(format!(
                "{name} must contain a Boolean8 value"
            ))),
        }
    }

    pub(super) fn required_i16(sprm: &Sprm, name: &str) -> Result<i16> {
        sprm.operand_i16()
            .ok_or_else(|| PackageError::Corrupted(format!("{name} is missing its 16-bit operand")))
    }

    pub(super) fn xas(sprm: &Sprm, name: &str) -> Result<i16> {
        let value = Self::required_i16(sprm, name)?;
        if !(-31_680..=31_680).contains(&value) {
            return Err(PackageError::Corrupted(format!(
                "{name} value {value} is outside -31680..=31680"
            )));
        }
        Ok(value)
    }

    pub(super) fn unsigned_twips(sprm: &Sprm, name: &str) -> Result<u16> {
        let value = sprm.operand_word().ok_or_else(|| {
            PackageError::Corrupted(format!("{name} is missing its 16-bit operand"))
        })?;
        if value > 31_680 {
            return Err(PackageError::Corrupted(format!(
                "{name} value {value} exceeds 31680"
            )));
        }
        Ok(value)
    }

    pub(super) fn required_i32(sprm: &Sprm, name: &str) -> Result<i32> {
        let bytes: [u8; 4] = sprm.operand_bytes().try_into().map_err(|_| {
            PackageError::Corrupted(format!("{name} is missing its 32-bit operand"))
        })?;
        Ok(i32::from_le_bytes(bytes))
    }

    pub(super) fn line_hundredths(sprm: &Sprm, name: &str) -> Result<i16> {
        let value = Self::required_i16(sprm, name)?;
        if !(-20..=31_680).contains(&value) {
            return Err(PackageError::Corrupted(format!(
                "{name} value {value} is outside -20..=31680"
            )));
        }
        Ok(value)
    }

    pub(super) fn nonnegative_distance(sprm: &Sprm, name: &str) -> Result<i16> {
        let value = Self::required_i16(sprm, name)?;
        if !(0..=31_680).contains(&value) {
            return Err(PackageError::Corrupted(format!(
                "{name} value {value} is outside 0..=31680"
            )));
        }
        Ok(value)
    }

    pub(super) fn parse_numbering_revision(sprm: &Sprm) -> Result<NumberingRevisionProperties> {
        let operand = sprm.operand_bytes();
        if operand.len() != 128 {
            return Err(PackageError::Corrupted(
                "sprmPNumRM operand must contain exactly 128 bytes".to_string(),
            ));
        }
        let was_numbered = match operand[0] {
            0 => false,
            1 => true,
            _ => {
                return Err(PackageError::Corrupted(
                    "NumRM.fNumRM must be a Boolean8 value".to_string(),
                ));
            },
        };
        let author = i16::from_le_bytes([operand[2], operand[3]]);
        let author_index = u16::try_from(author)
            .map_err(|_| PackageError::Corrupted("NumRM author index is negative".to_string()))?;
        let timestamp = u32::from_le_bytes([operand[4], operand[5], operand[6], operand[7]]);
        let placeholder_positions: [u8; 9] = operand[8..17].try_into().expect("fixed NumRM slice");
        let number_formats: [u8; 9] = operand[17..26].try_into().expect("fixed NumRM slice");
        let numbers = std::array::from_fn(|index| {
            let offset = 28 + index * 4;
            u32::from_le_bytes(
                operand[offset..offset + 4]
                    .try_into()
                    .expect("fixed NumRM integer"),
            )
        });
        let string_length = usize::from(u16::from_le_bytes([operand[64], operand[65]]));
        if string_length > 31 {
            return Err(PackageError::Corrupted(
                "NumRM format string exceeds its 31-code-unit field".to_string(),
            ));
        }
        let units = (0..string_length)
            .map(|index| {
                let offset = 66 + index * 2;
                u16::from_le_bytes([operand[offset], operand[offset + 1]])
            })
            .collect::<Vec<_>>();
        let format_string = String::from_utf16(&units).map_err(|_| {
            PackageError::Corrupted("NumRM format string is invalid UTF-16".to_string())
        })?;
        if placeholder_positions
            .iter()
            .any(|position| usize::from(*position) > string_length)
        {
            return Err(PackageError::Corrupted(
                "NumRM placeholder position exceeds its format string".to_string(),
            ));
        }
        Ok(NumberingRevisionProperties {
            was_numbered,
            author_index,
            timestamp,
            placeholder_positions,
            number_formats,
            numbers,
            format_string,
        })
    }

    /// Handle tab stops (sprmPChgTabsPapx).
    ///
    /// Tab stops are stored as:
    /// - 1 byte: number of tabs to delete (delSize)
    /// - delSize * 2 bytes: positions to delete
    /// - 1 byte: number of tabs to add (addSize)
    /// - addSize * 2 bytes: positions to add
    /// - addSize bytes: tab descriptors (jc + tlc)
    pub(super) fn handle_tabs(
        pap: &mut ParagraphProperties,
        sprm: &Sprm,
        delete_close: bool,
    ) -> Result<()> {
        let bytes = sprm.operand_bytes();
        if bytes.len() < 2 {
            return Err(PackageError::Corrupted(
                "DOC tab-change operand must contain delete and add counts".to_string(),
            ));
        }
        let delete_count = usize::from(bytes[0]);
        if delete_count > 64 {
            return Err(PackageError::Corrupted(
                "DOC tab-change delete count exceeds 64".to_string(),
            ));
        }
        let delete_bytes = if delete_close { 4 } else { 2 };
        let add_count_offset = 1 + delete_count * delete_bytes;
        if add_count_offset >= bytes.len() {
            return Err(PackageError::Corrupted(
                "DOC tab-change delete arrays are truncated".to_string(),
            ));
        }
        let add_count = usize::from(bytes[add_count_offset]);
        if add_count > 64 {
            return Err(PackageError::Corrupted(
                "DOC tab-change add count exceeds 64".to_string(),
            ));
        }
        let expected = add_count_offset + 1 + add_count * 3;
        if bytes.len() != expected {
            return Err(PackageError::Corrupted(format!(
                "DOC tab-change operand has {} bytes; expected {expected}",
                bytes.len()
            )));
        }

        let mut tab_map: std::collections::BTreeMap<i32, TabStop> =
            pap.tab_stops.iter().map(|t| (t.position, *t)).collect();
        let mut previous_delete = None;
        for index in 0..delete_count {
            let position = read_i16_le(bytes, 1 + index * 2).map_err(|error| {
                PackageError::Corrupted(format!("invalid tab deletion: {error}"))
            })?;
            if position > 31_680 || previous_delete.is_some_and(|previous| position <= previous) {
                return Err(PackageError::Corrupted(
                    "DOC tab deletion positions must be ascending and at most 31680".to_string(),
                ));
            }
            previous_delete = Some(position);
            let radius = if delete_close {
                let close_offset = 1 + delete_count * 2 + index * 2;
                let stored = read_i16_le(bytes, close_offset).map_err(|error| {
                    PackageError::Corrupted(format!("invalid close-tab distance: {error}"))
                })?;
                if !(-31_678..=31_682).contains(&stored) {
                    return Err(PackageError::Corrupted(format!(
                        "DOC close-tab distance {stored} is outside the XAS_plusOne range"
                    )));
                }
                i32::from(stored).saturating_sub(1).max(25)
            } else {
                25
            };
            let center = i32::from(position);
            tab_map.retain(|tab, _| (*tab - center).abs() > radius);
        }
        let positions_start = add_count_offset + 1;
        let descriptors_start = positions_start + add_count * 2;
        let mut previous_add = None;
        for index in 0..add_count {
            let position = read_i16_le(bytes, positions_start + index * 2).map_err(|error| {
                PackageError::Corrupted(format!("invalid tab addition: {error}"))
            })?;
            if !(-31_680..=31_680).contains(&position)
                || previous_add.is_some_and(|previous| position <= previous)
            {
                return Err(PackageError::Corrupted(
                    "DOC tab addition positions must be ascending XAS values".to_string(),
                ));
            }
            previous_add = Some(position);
            let descriptor = bytes[descriptors_start + index];
            let alignment = match descriptor & 0x07 {
                0 => TabAlignment::Left,
                1 => TabAlignment::Center,
                2 => TabAlignment::Right,
                3 => TabAlignment::Decimal,
                4 => TabAlignment::Bar,
                6 => TabAlignment::List,
                invalid => {
                    return Err(PackageError::Corrupted(format!(
                        "DOC tab has invalid alignment {invalid}"
                    )));
                },
            };
            let leader = if alignment == TabAlignment::Bar {
                // The leader field is explicitly ignored for bar tabs.
                TabLeader::None
            } else {
                match (descriptor >> 3) & 0x07 {
                    0 => TabLeader::None,
                    1 => TabLeader::Dots,
                    2 => TabLeader::Hyphens,
                    3 => TabLeader::Underline,
                    4 => TabLeader::Heavy,
                    5 => TabLeader::MiddleDot,
                    7 => TabLeader::DefaultLeader,
                    invalid => {
                        return Err(PackageError::Corrupted(format!(
                            "DOC tab has invalid leader {invalid}"
                        )));
                    },
                }
            };
            tab_map.insert(
                i32::from(position),
                TabStop {
                    position: i32::from(position),
                    alignment,
                    leader,
                },
            );
        }
        pap.tab_stops = tab_map.into_values().collect();
        Ok(())
    }

    /// Parse a Word 6/7 `BRC10` paragraph border and normalize it to `Brc80` units.
    pub(super) fn parse_border10(sprm: &Sprm) -> Result<Option<Border>> {
        let raw = sprm.operand_word().ok_or_else(|| {
            PackageError::Corrupted("DOC paragraph BRC10 must contain exactly 2 bytes".to_string())
        })?;
        let width_code = (raw & 0x07) as u8;
        let type_code = ((raw >> 3) & 0x03) as u8;
        if type_code != 0 && width_code == 0 {
            return Err(PackageError::Corrupted(
                "DOC paragraph BRC10 has a border type with zero line width".to_string(),
            ));
        }
        let (style, width) = match width_code {
            6 => (BorderStyle::Dotted, 6),
            7 => (BorderStyle::Dashed, 6),
            width => {
                let style = match type_code {
                    0 => return Ok(None),
                    1 => BorderStyle::Single,
                    2 => BorderStyle::Thick,
                    3 => BorderStyle::Double,
                    _ => unreachable!(),
                };
                (style, width * 6)
            },
        };
        let color_index = ((raw >> 6) & 0x1F) as u8;
        let color = match color_index {
            0 => None,
            index @ 1..=16 => Some(Self::get_ico_color(index)),
            invalid => {
                return Err(PackageError::Corrupted(format!(
                    "DOC paragraph BRC10 has invalid color index {invalid}"
                )));
            },
        };
        Ok(Some(Border {
            style,
            width,
            color,
            spacing: ((raw >> 11) & 0x1F) as u8,
            shadow: raw & 0x20 != 0,
            frame: false,
        }))
    }

    /// Parse a Word 97 `Brc80` paragraph border.
    pub(super) fn parse_border80(sprm: &Sprm) -> Result<Option<Border>> {
        let data = sprm.operand_bytes();
        if data.len() != 4 {
            return Err(PackageError::Corrupted(
                "DOC paragraph Brc80 must contain exactly 4 bytes".to_string(),
            ));
        }
        let Some(style) = Self::parse_border_style(data[1], false)? else {
            return Ok(None);
        };
        let color = match data[2] {
            0 => None,
            index @ 1..=16 => Some(Self::get_ico_color(index)),
            invalid => {
                return Err(PackageError::Corrupted(format!(
                    "DOC paragraph Brc80 has invalid color index {invalid}"
                )));
            },
        };
        Ok(Some(Border {
            style,
            width: data[0],
            color,
            spacing: data[3] & 0x1F,
            shadow: data[3] & 0x20 != 0,
            frame: data[3] & 0x40 != 0,
        }))
    }

    /// Parse a current 8-byte `Brc` wrapped by a `BrcOperand`.
    pub(super) fn parse_current_border(sprm: &Sprm) -> Result<Option<Border>> {
        let data = sprm.operand_bytes();
        if data.len() != 8 {
            return Err(PackageError::Corrupted(
                "DOC paragraph BrcOperand must contain exactly 8 bytes".to_string(),
            ));
        }
        let Some(style) = Self::parse_border_style(data[5], true)? else {
            return Ok(None);
        };
        let color = match data[3] {
            0 => Some((data[0], data[1], data[2])),
            0xFF => None,
            invalid => {
                return Err(PackageError::Corrupted(format!(
                    "DOC paragraph Brc has invalid automatic-color flag {invalid:#04x}"
                )));
            },
        };
        Ok(Some(Border {
            style,
            width: data[4],
            color,
            spacing: data[6] & 0x1F,
            shadow: data[6] & 0x20 != 0,
            frame: data[6] & 0x40 != 0,
        }))
    }

    pub(super) fn parse_border_style(code: u8, current: bool) -> Result<Option<BorderStyle>> {
        Ok(Some(match code {
            0 => return Ok(None),
            1 => BorderStyle::Single,
            3 => BorderStyle::Double,
            5 => BorderStyle::Thick,
            6 => BorderStyle::Dotted,
            7 => BorderStyle::Dashed,
            8 => BorderStyle::DotDash,
            9 => BorderStyle::DotDotDash,
            10 => BorderStyle::Triple,
            11 => BorderStyle::ThinThickSmallGap,
            12 => BorderStyle::ThickThinSmallGap,
            13 => BorderStyle::ThinThickThinSmallGap,
            14 => BorderStyle::ThinThickMediumGap,
            15 => BorderStyle::ThickThinMediumGap,
            16 => BorderStyle::ThinThickThinMediumGap,
            17 => BorderStyle::ThinThickLargeGap,
            18 => BorderStyle::ThickThinLargeGap,
            19 => BorderStyle::ThinThickThinLargeGap,
            20 => BorderStyle::Wave,
            21 => BorderStyle::DoubleWave,
            22 => BorderStyle::DashSmallGap,
            23 => BorderStyle::DashDotStroked,
            24 => BorderStyle::ThreeDEmboss,
            25 => BorderStyle::ThreeDEngrave,
            26 if current => BorderStyle::Outset,
            27 if current => BorderStyle::Inset,
            invalid => {
                return Err(PackageError::Corrupted(format!(
                    "DOC paragraph border has invalid type {invalid:#04x}"
                )));
            },
        }))
    }

    /// Parse shading from Shd80 (2 bytes).
    pub(super) fn parse_shd80(shd: u16) -> Result<Option<Shading>> {
        if shd == u16::MAX {
            return Ok(None);
        }
        let ico_fore = (shd & 0x1F) as u8;
        let ico_back = ((shd >> 5) & 0x1F) as u8;
        let ipat = ((shd >> 10) & 0x3F) as u8;
        let pattern = ShadingPattern::from_u8(ipat).ok_or_else(|| {
            PackageError::Corrupted(format!("sprmPShd80 has invalid pattern {ipat:#04x}"))
        })?;
        if pattern == ShadingPattern::Auto {
            return Ok(None);
        }
        let palette_color = |index| match index {
            0 => Ok(None),
            value @ 1..=16 => Ok(Some(Self::get_ico_color(value))),
            invalid => Err(PackageError::Corrupted(format!(
                "sprmPShd80 has invalid color index {invalid}"
            ))),
        };
        Ok(Some(Shading {
            foreground_color: palette_color(ico_fore)?,
            background_color: palette_color(ico_back)?,
            pattern,
        }))
    }

    /// Parse shading from ShadingDescriptor (10 bytes).
    pub(super) fn parse_shading_descriptor(sprm: &Sprm) -> Result<Option<Shading>> {
        let data = sprm.operand_bytes();
        if data.len() != 10 {
            return Err(PackageError::Corrupted(
                "sprmPShd SHDOperand must contain exactly 10 bytes".to_string(),
            ));
        }
        let pattern_code = read_u16_le(data, 8).map_err(|error| {
            PackageError::Corrupted(format!("sprmPShd has invalid pattern: {error}"))
        })?;
        let pattern = u8::try_from(pattern_code)
            .ok()
            .and_then(ShadingPattern::from_u8)
            .ok_or_else(|| {
                PackageError::Corrupted(format!("sprmPShd has invalid pattern {pattern_code:#06x}"))
            })?;
        if pattern == ShadingPattern::Auto {
            return Ok(None);
        }
        let colorref = |bytes: &[u8]| match bytes[3] {
            0 => Ok(Some((bytes[0], bytes[1], bytes[2]))),
            0xFF => Ok(None),
            invalid => Err(PackageError::Corrupted(format!(
                "sprmPShd has invalid automatic-color flag {invalid:#04x}"
            ))),
        };
        Ok(Some(Shading {
            foreground_color: colorref(&data[..4])?,
            background_color: colorref(&data[4..8])?,
            pattern,
        }))
    }

    /// Get color from ico index.
    pub(super) fn get_ico_color(ico: u8) -> (u8, u8, u8) {
        match ico {
            0 => (0, 0, 0),        // Auto/Black
            1 => (0, 0, 0),        // Black
            2 => (0, 0, 255),      // Blue
            3 => (0, 255, 255),    // Cyan
            4 => (0, 255, 0),      // Green
            5 => (255, 0, 255),    // Magenta
            6 => (255, 0, 0),      // Red
            7 => (255, 255, 0),    // Yellow
            8 => (255, 255, 255),  // White
            9 => (0, 0, 128),      // Dark Blue
            10 => (0, 128, 128),   // Dark Cyan
            11 => (0, 128, 0),     // Dark Green
            12 => (128, 0, 128),   // Dark Magenta
            13 => (128, 0, 0),     // Dark Red
            14 => (128, 128, 0),   // Dark Yellow
            15 => (128, 128, 128), // Dark Gray
            16 => (192, 192, 192), // Light Gray
            _ => (0, 0, 0),
        }
    }
}
