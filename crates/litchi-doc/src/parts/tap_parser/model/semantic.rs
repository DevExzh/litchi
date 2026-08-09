//! Semantic table-property state transitions.

use super::prelude::*;

impl TapParser<'_> {
    /// Apply a single SPRM to table properties.
    ///
    /// Based on Apache POI's unCompressTAPOperation method.

    pub(in crate::parts::tap_parser) fn apply_sprm_to_tap(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        grpprl: &[u8],
        inside_conditional: bool,
        stylesheet: Option<&StyleSheet>,
    ) -> Result<()> {
        // Use shared SPRM operation extraction
        let operation = get_sprm_operation(sprm.opcode);

        match operation {
            // sprmTJc90 (0x5400) - Physical table justification
            0x00 => {
                tap.legacy_physical_justification =
                    Some(Self::parse_justification(sprm, "sprmTJc90")?);
                Self::resolve_justification(tap);
            },
            // sprmTDxaLeft (0x9601) - Table indent from left
            0x01 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 2 {
                    return Err(PackageError::Corrupted(
                        "sprmTDxaLeft operand must contain 2 bytes".to_string(),
                    ));
                }
                let indent = binary_to_doc_result(read_i16_le(operand, 0))?;
                if !(-31_680..=31_680).contains(&indent) {
                    return Err(PackageError::Corrupted(
                        "sprmTDxaLeft is outside the XAS range".to_string(),
                    ));
                }
                let current_origin = i32::from(tap.cell_boundaries.first().copied().unwrap_or(0))
                    + i32::from(tap.gap_half);
                let adjust = i32::from(indent) - current_origin;
                let boundaries = tap
                    .cell_boundaries
                    .iter()
                    .map(|boundary| i32::from(*boundary) + adjust)
                    .map(|boundary| {
                        i16::try_from(boundary).map_err(|_| {
                            PackageError::Corrupted(
                                "sprmTDxaLeft overflows table coordinates".to_string(),
                            )
                        })
                    })
                    .collect::<Result<Vec<_>>>()?;
                if boundaries
                    .iter()
                    .any(|boundary| !(-31_680..=31_680).contains(boundary))
                {
                    return Err(PackageError::Corrupted(
                        "sprmTDxaLeft produces a coordinate outside the XAS range".to_string(),
                    ));
                }
                tap.cell_boundaries = boundaries;
                tap.indent_left = indent;
            },
            // sprmTDxaGapHalf (0x9602) - Half the width of spacing between cells
            0x02 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 2 {
                    return Err(PackageError::Corrupted(
                        "sprmTDxaGapHalf operand must contain 2 bytes".to_string(),
                    ));
                }
                let gap = binary_to_doc_result(read_i16_le(operand, 0))?;
                if !(0..=31_680).contains(&gap) {
                    return Err(PackageError::Corrupted(
                        "sprmTDxaGapHalf is outside its allowed XAS range".to_string(),
                    ));
                }
                if let Some(first) = tap.cell_boundaries.first().copied() {
                    let adjusted = i32::from(first) + i32::from(tap.gap_half) - i32::from(gap);
                    if !(-31_680..=31_680).contains(&adjusted)
                        || tap
                            .cell_boundaries
                            .get(1)
                            .is_some_and(|second| adjusted > i32::from(*second))
                    {
                        return Err(PackageError::Corrupted(
                            "sprmTDxaGapHalf produces an invalid first cell boundary".to_string(),
                        ));
                    }
                    tap.cell_boundaries[0] = adjusted as i16;
                }
                tap.gap_half = gap;
            },
            // sprmTFCantSplit (0x3403) - Row can't be split across pages
            0x03 => {
                tap.allow_row_break = !Self::parse_bool8(sprm, "sprmTFCantSplit90")?;
            },
            // sprmTTableHeader (0x3404) - Row is header row
            0x04 => {
                tap.is_header_row = Self::parse_bool8(sprm, "sprmTTableHeader")?;
            },
            // sprmTTableBorders (0xD605) - Table borders
            0x05 => {
                self.parse_table_borders(tap, sprm, grpprl)?;
            },
            // 0x06 - obsolete (Word 1.x)
            0x06 => {},
            // sprmTDyaRowHeight (0x9407) - Row height
            0x07 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 2 {
                    return Err(PackageError::Corrupted(
                        "sprmTDyaRowHeight operand must contain 2 bytes".to_string(),
                    ));
                }
                let height = binary_to_doc_result(read_i16_le(operand, 0))?;
                if !(-31_680..=31_680).contains(&height) {
                    return Err(PackageError::Corrupted(
                        "sprmTDyaRowHeight is outside the YAS range".to_string(),
                    ));
                }
                tap.row_height = (height != 0).then_some(height);
            },
            // sprmTDefTable (0xD608) - Table definition
            0x08 => {
                self.parse_table_definition(tap, sprm, grpprl)?;
            },
            // sprmTDefTableShd80 (0xD609) - Table shading
            0x09 => {
                // Parse cell shading information
                // Format: variable length array of ShadingDescriptor80 (2 bytes each)
                // Based on Apache POI's handling of sprmTDefTableShd
                self.parse_cell_shading(tap, sprm, grpprl)?;
            },
            // Full-color default shading chunks for cells 45..63, 1..22, and 23..44.
            0x0C => self.parse_full_cell_shading(tap, sprm, 44, false)?,
            // sprmTTlp (0x740A) - Table auto-format look specifier
            0x0A => {
                let operand = sprm.operand_bytes();
                if operand.len() != 4 {
                    return Err(PackageError::Corrupted(
                        "sprmTTlp operand must contain 4 bytes".to_string(),
                    ));
                }
                let bits = binary_to_doc_result(read_u16_le(operand, 2))?;
                let flags = TableLookFlags::from_bits(bits).ok_or_else(|| {
                    PackageError::Corrupted("sprmTTlp Fatl padding bits are nonzero".to_string())
                })?;
                tap.table_look = Some(TableLook {
                    autoformat_index: binary_to_doc_result(read_i16_le(operand, 0))?,
                    flags,
                });
            },
            0x0B => {
                tap.legacy_right_to_left = Self::parse_bool16(sprm, "sprmTFBiDi")?;
                tap.right_to_left = tap.legacy_right_to_left || tap.modern_right_to_left;
                Self::resolve_justification(tap);
            },
            0x0D => {
                let operand = sprm.operand_bytes();
                if operand.len() != 1 || operand[0] & 0x0F != 0 {
                    return Err(PackageError::Corrupted(
                        "sprmTPc contains invalid PositionCode padding".to_string(),
                    ));
                }
                tap.positioning = Some(TablePositioning {
                    vertical_anchor: match (operand[0] >> 4) & 0x03 {
                        0 => TableVerticalAnchor::Margin,
                        1 => TableVerticalAnchor::Page,
                        2 => TableVerticalAnchor::Paragraph,
                        3 => TableVerticalAnchor::None,
                        _ => unreachable!(),
                    },
                    horizontal_anchor: match operand[0] >> 6 {
                        0 => TableHorizontalAnchor::Column,
                        1 => TableHorizontalAnchor::Margin,
                        2 => TableHorizontalAnchor::Page,
                        3 => TableHorizontalAnchor::None,
                        _ => unreachable!(),
                    },
                });
            },
            0x0E => tap.horizontal_position = Self::parse_horizontal_position(sprm)?,
            0x0F => tap.vertical_position = Self::parse_vertical_position(sprm)?,
            0x10 => {
                tap.distance_from_text_left = Self::parse_wrap_distance(sprm, "sprmTDxaFromText")?;
            },
            0x11 => {
                tap.distance_from_text_top = Self::parse_wrap_distance(sprm, "sprmTDyaFromText")?;
            },
            0x12 => self.parse_full_cell_shading(tap, sprm, 0, false)?,
            // Full-color row border defaults.
            0x13 => self.parse_full_table_borders(tap, sprm)?,
            0x14 => tap.preferred_width = Self::parse_fts_width(sprm, WidthUsage::Table)?,
            0x15 => tap.auto_fit = Self::parse_bool8(sprm, "sprmTFAutofit")?,
            0x16 => self.parse_full_cell_shading(tap, sprm, 22, false)?,
            0x17 => tap.width_before = Self::parse_fts_width(sprm, WidthUsage::TablePart)?,
            0x18 => tap.width_after = Self::parse_fts_width(sprm, WidthUsage::TablePart)?,
            0x19 => tap.keep_with_next = Self::parse_bool8(sprm, "sprmTFKeepFollow")?,
            // Per-cell colors for top, left, bottom, and right borders.
            0x1A..=0x1D => self.parse_cell_border_colors(tap, sprm, operation)?,
            0x1E => {
                tap.distance_from_text_right =
                    Self::parse_wrap_distance(sprm, "sprmTDxaFromTextRight")?;
            },
            0x1F => {
                tap.distance_from_text_bottom =
                    Self::parse_wrap_distance(sprm, "sprmTDyaFromTextBottom")?;
            },
            // Legacy range border override.
            0x20 => self.parse_cell_border_range(tap, sprm, false)?,
            // sprmTInsert (0x7621) - Insert cells
            0x21 => {
                self.handle_insert_cells(tap, sprm)?;
            },
            0x22 => self.handle_delete_cells(tap, sprm)?,
            0x23 => self.handle_column_width(tap, sprm)?,
            0x24 => self.handle_horizontal_merge(tap, sprm, true)?,
            0x25 => self.handle_horizontal_merge(tap, sprm, false)?,
            0x29 => self.parse_cell_text_flow(tap, sprm)?,
            0x2B => self.parse_vertical_merge(tap, sprm)?,
            0x2C => self.parse_vertical_alignment(tap, sprm)?,
            // Full-color shading over every or every other cell in a range.
            0x2D => self.parse_full_cell_shading_range(tap, sprm, false)?,
            0x2E => self.parse_full_cell_shading_range(tap, sprm, true)?,
            // Full-color range border override.
            0x2F => self.parse_cell_border_range(tap, sprm, true)?,
            // Full-color shading applied to every cell in the table row.
            0x60 => self.parse_full_table_shading(tap, sprm)?,
            0x61 => tap.preferred_indent = Self::parse_fts_width(sprm, WidthUsage::Indent)?,
            0x62 => self.parse_cell_border_types(tap, sprm)?,
            // sprmTCellPadding / sprmTCellPaddingDefault
            0x32 | 0x34 => {
                self.parse_cell_padding(tap, sprm, grpprl)?;
            },
            0x33 => self.parse_cell_spacing(tap, sprm)?,
            0x35 => self.parse_cell_width(tap, sprm)?,
            0x36 => self.parse_cell_range_bool(tap, sprm, CellBoolProperty::FitText)?,
            0x39 => self.parse_cell_range_bool(tap, sprm, CellBoolProperty::NoWrap)?,
            0x3A => {
                let operand = sprm.operand_bytes();
                if operand.len() != 2 {
                    return Err(PackageError::Corrupted(
                        "sprmTIstd operand must contain 2 bytes".to_string(),
                    ));
                }
                let requested = binary_to_doc_result(read_u16_le(operand, 0))?;
                if let Some(stylesheet) = stylesheet {
                    let (effective, style) = stylesheet.resolve_table_properties(requested)?;
                    Self::apply_table_style(tap, style, effective);
                } else {
                    tap.table_style_index = Some(requested);
                }
            },
            0x3E => self.parse_style_cell_padding(tap, sprm)?,
            0x42 => self.parse_cell_range_bool(tap, sprm, CellBoolProperty::HideMark)?,
            0x64 => {
                tap.modern_right_to_left = Self::parse_bool16(sprm, "sprmTFBiDi90")?;
                tap.right_to_left = tap.legacy_right_to_left || tap.modern_right_to_left;
                Self::resolve_justification(tap);
            },
            0x65 => {
                tap.allow_overlap = !Self::parse_bool8(sprm, "sprmTFNoAllowOverlap")?;
            },
            // Modern row can't-split property supersedes sprmTFCantSplit90.
            0x66 => {
                tap.allow_row_break = !Self::parse_bool8(sprm, "sprmTFCantSplit")?;
            },
            // sprmTPropRMark (0xD667) - Row property revision mark
            0x67 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 7 {
                    return Err(PackageError::Corrupted(
                        "sprmTPropRMark operand must contain exactly 7 bytes".to_string(),
                    ));
                }
                tap.has_formatting_revision = Some(match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(PackageError::Corrupted(
                            "sprmTPropRMark must begin with a Boolean8 value".to_string(),
                        ));
                    },
                });
                let author = i16::from_le_bytes([operand[1], operand[2]]);
                tap.formatting_revision_author_index =
                    Some(u16::try_from(author).map_err(|_| {
                        PackageError::Corrupted(
                            "sprmTPropRMark author index is negative".to_string(),
                        )
                    })?);
                let timestamp =
                    u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
                crate::revision::decode_dttm(timestamp)?;
                tap.formatting_revision_timestamp = Some(timestamp);
            },
            // Raw defaults preserve ShdNil as table-style inheritance.
            0x70 => self.parse_full_cell_shading(tap, sprm, 0, true)?,
            0x71 => self.parse_full_cell_shading(tap, sprm, 22, true)?,
            0x72 => self.parse_full_cell_shading(tap, sprm, 44, true)?,
            // sprmTWall (0x3668) - Preserve properties before tracked changes
            0x68 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 1 {
                    return Err(PackageError::Corrupted(
                        "sprmTWall operand must contain exactly 1 byte".to_string(),
                    ));
                }
                let enabled = match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(PackageError::Corrupted(
                            "sprmTWall must contain a Boolean8 value".to_string(),
                        ));
                    },
                };
                tap.preserved_properties_for_revision = if enabled {
                    let mut previous = tap.clone();
                    previous.properties_preserved_for_revision = false;
                    previous.preserved_properties_for_revision = None;
                    Some(Box::new(previous))
                } else {
                    None
                };
                tap.properties_preserved_for_revision = enabled;
            },
            0x6A => {
                if inside_conditional {
                    return Err(PackageError::Corrupted(
                        "sprmTCnf cannot be nested inside another sprmTCnf".to_string(),
                    ));
                }
                self.parse_conditional_formatting(tap, sprm)?;
            },
            0x69 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 4 {
                    return Err(PackageError::Corrupted(
                        "sprmTIpgp operand must contain exactly 4 bytes".to_string(),
                    ));
                }
                let identifier = binary_to_doc_result(read_u32_le(operand, 0))?;
                if identifier == 0 {
                    return Err(PackageError::Corrupted(
                        "sprmTIpgp cannot reference PGPInfo identifier zero".to_string(),
                    ));
                }
                tap.paragraph_group_id = Some(identifier);
            },
            0x79 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 4 {
                    return Err(PackageError::Corrupted(
                        "sprmTRsid operand must contain exactly 4 bytes".to_string(),
                    ));
                }
                tap.revision_save_id = Some(binary_to_doc_result(read_u32_le(operand, 0))?);
            },
            0x7C => {
                tap.style_defaults.vertical_alignment =
                    Some(match Self::parse_byte(sprm, "sprmTCellVertAlignStyle")? {
                        0 => VerticalAlignment::Top,
                        1 => VerticalAlignment::Center,
                        2 => VerticalAlignment::Bottom,
                        _ => {
                            return Err(PackageError::Corrupted(
                                "sprmTCellVertAlignStyle contains an invalid alignment".to_string(),
                            ));
                        },
                    });
            },
            0x7D => {
                tap.style_defaults.no_wrap = Some(Self::parse_bool8(sprm, "sprmTCellNoWrapStyle")?);
            },
            0x7F..=0x86 => {
                if !inside_conditional {
                    return Err(PackageError::Corrupted(
                        "DOC table-style border SPRMs are only valid inside sprmTCnf".to_string(),
                    ));
                }
                self.parse_style_border(tap, sprm, operation)?;
            },
            0x87 => self.parse_style_shading(tap, sprm)?,
            0x88 => {
                tap.style_defaults.horizontal_band_size =
                    Some(Self::parse_band_size(sprm, "sprmTCHorzBands")?);
            },
            0x89 => {
                tap.style_defaults.vertical_band_size =
                    Some(Self::parse_band_size(sprm, "sprmTCVertBands")?);
            },
            0x8A => {
                tap.modern_logical_justification =
                    Some(Self::parse_justification(sprm, "sprmTJc")?);
                Self::resolve_justification(tap);
            },
            // Other table SPRMs.
            _ => {
                // Unknown or unhandled SPRM - skip
            },
        }

        Ok(())
    }

    /// Apply the properties controlled by an `UpxTapx` while retaining the
    /// structural, positioning, sizing, and revision state that MS-DOC says a
    /// `sprmTIstd` must preserve.

    pub(in crate::parts::tap_parser) fn apply_table_style(
        tap: &mut TableProperties,
        style: TableProperties,
        effective: u16,
    ) {
        tap.legacy_physical_justification = style.legacy_physical_justification;
        tap.modern_logical_justification = style.modern_logical_justification;
        tap.style_defaults = style.style_defaults;
        tap.conditional_formats = style.conditional_formats;
        tap.width_before = style.width_before;
        tap.preferred_indent = style.preferred_indent;
        tap.cell_spacing = style.cell_spacing;
        tap.is_header_row = style.is_header_row;
        tap.allow_row_break = style.allow_row_break;
        tap.border_top = style.border_top;
        tap.border_left = style.border_left;
        tap.border_bottom = style.border_bottom;
        tap.border_right = style.border_right;
        tap.border_horizontal = style.border_horizontal;
        tap.border_vertical = style.border_vertical;
        tap.table_style_index = Some(effective);
        Self::resolve_justification(tap);
    }

    pub(in crate::parts::tap_parser) fn resolve_justification(tap: &mut TableProperties) {
        tap.justification = tap.modern_logical_justification.unwrap_or_else(|| {
            let Some(physical) = tap.legacy_physical_justification else {
                return TableJustification::Left;
            };
            if tap.right_to_left {
                match physical {
                    TableJustification::Left => TableJustification::Right,
                    TableJustification::Center => TableJustification::Center,
                    TableJustification::Right => TableJustification::Left,
                }
            } else {
                physical
            }
        });
    }
}
