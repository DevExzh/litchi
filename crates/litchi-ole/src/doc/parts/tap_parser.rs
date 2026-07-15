/// TAP (Table Properties) parser with arena allocator support.
///
/// This module provides comprehensive TAP parsing based on Apache POI's
/// TableSprmUncompressor implementation. Uses arena allocators for efficient
/// memory management of temporary parsing structures.
///
/// Reference: Apache POI's org.apache.poi.hwpf.sprm.TableSprmUncompressor
use super::super::package::{DocError, Result};

/// Helper function to read a single byte.
#[inline]
fn read_byte(data: &[u8], offset: usize) -> BinaryResult<u8> {
    if offset >= data.len() {
        return Err(litchi_core::binary::BinaryError::InsufficientData {
            expected: offset + 1,
            available: data.len(),
        });
    }
    Ok(data[offset])
}

/// Convert BinaryResult to DocError Result.
#[inline]
fn binary_to_doc_result<T>(result: BinaryResult<T>) -> Result<T> {
    result.map_err(|e| DocError::InvalidFormat(format!("Binary read error: {}", e)))
}
use super::tap::{
    BorderStyle, BorderType, CellMergeStatus, CellProperties, CellShading, CellSpacing,
    CellSpacingSource, ShadingPattern, TableHorizontalAnchor, TableHorizontalPosition,
    TableJustification, TableLook, TableLookFlags, TablePositioning, TableProperties,
    TableVerticalAnchor, TableVerticalPosition, TableWidth, TextDirection, VerticalAlignment,
    VerticalMergeStatus, WidthType,
};
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::get_sprm_operation;
use bumpalo::Bump;
use litchi_core::binary::{BinaryResult, read_i16_le, read_u16_le};

/// TAP parser with arena allocation for temporary structures.
///
/// Uses bumpalo arena allocator for zero-cost temporary allocations
/// during TAP parsing. The arena is automatically cleaned up when
/// the parser is dropped.
pub struct TapParser<'arena> {
    /// Arena allocator for temporary parsing data (reserved for future use)
    #[allow(dead_code)]
    arena: &'arena Bump,
}

#[derive(Debug, Clone, Copy)]
enum CellBoolProperty {
    FitText,
    NoWrap,
    HideMark,
}

#[derive(Debug, Clone, Copy)]
enum WidthUsage {
    Table,
    TablePart,
    Indent,
}

impl<'arena> TapParser<'arena> {
    /// Create a new TAP parser with an arena allocator.
    ///
    /// # Arguments
    ///
    /// * `arena` - Bump allocator arena for temporary allocations
    pub fn new(arena: &'arena Bump) -> Self {
        Self { arena }
    }

    /// Parse table properties from SPRM list.
    ///
    /// Based on Apache POI's uncompressTAP method.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (Single Property Modifiers)
    ///
    /// # Returns
    ///
    /// Parsed TableProperties structure
    pub fn parse_tap(&self, grpprl: &[u8]) -> Result<TableProperties> {
        // Parse all SPRMs using arena for temporary storage
        let sprms = parse_sprms(grpprl);
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != grpprl.len() {
            return Err(DocError::Corrupted(
                "TAP grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }

        // Find sprmTDefTable (0xD608 / operation 0x08) to initialize TAP
        let mut tap = self.find_and_init_tap(&sprms)?;

        // Apply each TAP-type SPRM to the table properties
        for sprm in sprms {
            if Self::is_tap_sprm(sprm.opcode) {
                self.apply_sprm_to_tap(&mut tap, &sprm, grpprl)?;
            }
        }

        Self::validate_preferred_indent(&tap)?;

        Ok(tap)
    }

    /// Find sprmTDefTable and initialize TAP structure.
    ///
    /// This SPRM defines the basic table structure including cell count
    /// and cell boundaries.
    fn find_and_init_tap(&self, sprms: &[Sprm]) -> Result<TableProperties> {
        for sprm in sprms {
            if sprm.opcode == 0xD608 {
                // Found sprmTDefTable
                // The shared decoder removes the long-SPRM size field, so the
                // first operand byte is itcMac.
                if let Some(cell_count) = sprm.operand_byte() {
                    let cell_count = cell_count as usize;
                    if cell_count > 63 {
                        return Err(DocError::Corrupted(
                            "sprmTDefTable contains more than 63 columns".to_string(),
                        ));
                    }
                    return Ok(TableProperties::with_cell_count(cell_count));
                }
            }
        }

        // No table definition found - use default with 1 cell
        Ok(TableProperties::with_cell_count(1))
    }

    /// Check if a SPRM is a TAP (table) SPRM.
    ///
    /// TAP SPRMs have type 5 (bits 10-12 of opcode).
    fn is_tap_sprm(opcode: u16) -> bool {
        ((opcode >> 10) & 0x07) == 5
    }

    fn parse_bool8(sprm: &Sprm, name: &str) -> Result<bool> {
        let operand = sprm.operand_bytes();
        if operand.len() != 1 || !matches!(operand[0], 0 | 1) {
            return Err(DocError::Corrupted(format!(
                "{name} must contain one Boolean8 value"
            )));
        }
        Ok(operand[0] != 0)
    }

    fn parse_bool16(sprm: &Sprm, name: &str) -> Result<bool> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(DocError::Corrupted(format!(
                "{name} must contain one Bool16 value"
            )));
        }
        match binary_to_doc_result(read_u16_le(operand, 0))? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(DocError::Corrupted(format!(
                "{name} contains an invalid Bool16 value"
            ))),
        }
    }

    fn parse_horizontal_position(sprm: &Sprm) -> Result<TableHorizontalPosition> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(DocError::Corrupted(
                "sprmTDxaAbs operand must contain 2 bytes".to_string(),
            ));
        }
        let stored = binary_to_doc_result(read_i16_le(operand, 0))?;
        Ok(match stored {
            0 => TableHorizontalPosition::Left,
            -4 => TableHorizontalPosition::Center,
            -8 => TableHorizontalPosition::Right,
            -12 => TableHorizontalPosition::Inside,
            -16 => TableHorizontalPosition::Outside,
            _ => {
                let offset = i32::from(stored) - 1;
                if !(-31_679..=31_681).contains(&offset) {
                    return Err(DocError::Corrupted(
                        "sprmTDxaAbs is outside the XAS_plusOne range".to_string(),
                    ));
                }
                TableHorizontalPosition::Offset(offset as i16)
            },
        })
    }

    fn parse_vertical_position(sprm: &Sprm) -> Result<TableVerticalPosition> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(DocError::Corrupted(
                "sprmTDyaAbs operand must contain 2 bytes".to_string(),
            ));
        }
        let stored = binary_to_doc_result(read_i16_le(operand, 0))?;
        Ok(match stored {
            0 => TableVerticalPosition::Inline,
            -4 => TableVerticalPosition::Top,
            -8 => TableVerticalPosition::Center,
            -12 => TableVerticalPosition::Bottom,
            -16 => TableVerticalPosition::Inside,
            -20 => TableVerticalPosition::Outside,
            _ => {
                let offset = i32::from(stored) - 1;
                if !(-31_679..=31_681).contains(&offset) {
                    return Err(DocError::Corrupted(
                        "sprmTDyaAbs is outside the YAS_plusOne range".to_string(),
                    ));
                }
                TableVerticalPosition::Offset(offset as i16)
            },
        })
    }

    fn parse_wrap_distance(sprm: &Sprm, name: &str) -> Result<u16> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(DocError::Corrupted(format!(
                "{name} operand must contain 2 bytes"
            )));
        }
        let value = binary_to_doc_result(read_u16_le(operand, 0))?;
        if value > 31_680 {
            return Err(DocError::Corrupted(format!(
                "{name} is outside its nonnegative distance range"
            )));
        }
        Ok(value)
    }

    fn parse_fts_width(sprm: &Sprm, usage: WidthUsage) -> Result<Option<TableWidth>> {
        let operand = sprm.operand_bytes();
        if operand.len() != 3 {
            return Err(DocError::Corrupted(
                "DOC preferred-width operand must contain 3 bytes".to_string(),
            ));
        }
        let raw_value = binary_to_doc_result(read_u16_le(operand, 1))?;
        let signed_value = i16::from_le_bytes([operand[1], operand[2]]);
        let invalid = || {
            DocError::Corrupted(
                "DOC preferred-width operand has invalid units or value".to_string(),
            )
        };
        Ok(match (usage, operand[0]) {
            (WidthUsage::Table | WidthUsage::Indent, 0) if raw_value == 0 => None,
            (WidthUsage::TablePart, 0) => None,
            (_, 1) if raw_value == 0 => Some(TableWidth {
                value: 0,
                width_type: WidthType::Auto,
            }),
            (WidthUsage::Table, 2) if raw_value <= 30_000 => Some(TableWidth {
                value: raw_value as i16,
                width_type: WidthType::Percentage,
            }),
            (WidthUsage::TablePart, 2) if raw_value <= 5_000 => Some(TableWidth {
                value: raw_value as i16,
                width_type: WidthType::Percentage,
            }),
            (WidthUsage::Table | WidthUsage::TablePart, 3) if raw_value <= 31_680 => {
                Some(TableWidth {
                    value: raw_value as i16,
                    width_type: WidthType::Twips,
                })
            },
            (WidthUsage::Indent, 3) if (-31_560..=31_680).contains(&signed_value) => {
                Some(TableWidth {
                    value: signed_value,
                    width_type: WidthType::Twips,
                })
            },
            _ => return Err(invalid()),
        })
    }

    fn validate_preferred_indent(tap: &TableProperties) -> Result<()> {
        let Some(TableWidth {
            value: indent,
            width_type: WidthType::Twips,
        }) = tap.preferred_indent
        else {
            return Ok(());
        };
        let table_width = match tap.preferred_width {
            Some(TableWidth {
                value,
                width_type: WidthType::Twips,
            }) => i32::from(value),
            _ => {
                i32::from(tap.cell_boundaries.last().copied().unwrap_or(0))
                    - i32::from(tap.cell_boundaries.first().copied().unwrap_or(0))
            },
        };
        if i32::from(indent) + table_width > 31_680 {
            return Err(DocError::Corrupted(
                "DOC preferred table indent places the right edge beyond 31680 twips".to_string(),
            ));
        }
        Ok(())
    }

    /// Apply a single SPRM to table properties.
    ///
    /// Based on Apache POI's unCompressTAPOperation method.
    fn apply_sprm_to_tap(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        grpprl: &[u8],
    ) -> Result<()> {
        // Use shared SPRM operation extraction
        let operation = get_sprm_operation(sprm.opcode);

        match operation {
            // sprmTJc90 (0x5400) - Physical table justification
            0x00 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 2 {
                    return Err(DocError::Corrupted(
                        "sprmTJc90 operand must contain 2 bytes".to_string(),
                    ));
                }
                tap.justification = match binary_to_doc_result(read_u16_le(operand, 0))? {
                    0 => TableJustification::Left,
                    1 => TableJustification::Center,
                    2 => TableJustification::Right,
                    _ => {
                        return Err(DocError::Corrupted(
                            "sprmTJc90 contains an invalid justification".to_string(),
                        ));
                    },
                };
            },
            // sprmTDxaLeft (0x9601) - Table indent from left
            0x01 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 2 {
                    return Err(DocError::Corrupted(
                        "sprmTDxaLeft operand must contain 2 bytes".to_string(),
                    ));
                }
                let indent = binary_to_doc_result(read_i16_le(operand, 0))?;
                if !(-31_680..=31_680).contains(&indent) {
                    return Err(DocError::Corrupted(
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
                            DocError::Corrupted(
                                "sprmTDxaLeft overflows table coordinates".to_string(),
                            )
                        })
                    })
                    .collect::<Result<Vec<_>>>()?;
                if boundaries
                    .iter()
                    .any(|boundary| !(-31_680..=31_680).contains(boundary))
                {
                    return Err(DocError::Corrupted(
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
                    return Err(DocError::Corrupted(
                        "sprmTDxaGapHalf operand must contain 2 bytes".to_string(),
                    ));
                }
                let gap = binary_to_doc_result(read_i16_le(operand, 0))?;
                if !(0..=31_680).contains(&gap) {
                    return Err(DocError::Corrupted(
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
                        return Err(DocError::Corrupted(
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
                    return Err(DocError::Corrupted(
                        "sprmTDyaRowHeight operand must contain 2 bytes".to_string(),
                    ));
                }
                let height = binary_to_doc_result(read_i16_le(operand, 0))?;
                if !(-31_680..=31_680).contains(&height) {
                    return Err(DocError::Corrupted(
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
                    return Err(DocError::Corrupted(
                        "sprmTTlp operand must contain 4 bytes".to_string(),
                    ));
                }
                let bits = binary_to_doc_result(read_u16_le(operand, 2))?;
                let flags = TableLookFlags::from_bits(bits).ok_or_else(|| {
                    DocError::Corrupted("sprmTTlp Fatl padding bits are nonzero".to_string())
                })?;
                tap.table_look = Some(TableLook {
                    autoformat_index: binary_to_doc_result(read_i16_le(operand, 0))?,
                    flags,
                });
            },
            0x0B => {
                tap.legacy_right_to_left = Self::parse_bool16(sprm, "sprmTFBiDi")?;
                tap.right_to_left = tap.legacy_right_to_left || tap.modern_right_to_left;
            },
            0x0D => {
                let operand = sprm.operand_bytes();
                if operand.len() != 1 || operand[0] & 0x0F != 0 {
                    return Err(DocError::Corrupted(
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
                tap.distance_from_text_left = Self::parse_wrap_distance(sprm, "sprmTDxaFromText")?
            },
            0x11 => {
                tap.distance_from_text_top = Self::parse_wrap_distance(sprm, "sprmTDyaFromText")?
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
                    Self::parse_wrap_distance(sprm, "sprmTDxaFromTextRight")?
            },
            0x1F => {
                tap.distance_from_text_bottom =
                    Self::parse_wrap_distance(sprm, "sprmTDyaFromTextBottom")?
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
                    return Err(DocError::Corrupted(
                        "sprmTIstd operand must contain 2 bytes".to_string(),
                    ));
                }
                tap.table_style_index = Some(binary_to_doc_result(read_u16_le(operand, 0))?);
            },
            0x42 => self.parse_cell_range_bool(tap, sprm, CellBoolProperty::HideMark)?,
            0x64 => {
                tap.modern_right_to_left = Self::parse_bool16(sprm, "sprmTFBiDi90")?;
                tap.right_to_left = tap.legacy_right_to_left || tap.modern_right_to_left;
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
                    return Err(DocError::Corrupted(
                        "sprmTPropRMark operand must contain exactly 7 bytes".to_string(),
                    ));
                }
                tap.has_formatting_revision = Some(match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(DocError::Corrupted(
                            "sprmTPropRMark must begin with a Boolean8 value".to_string(),
                        ));
                    },
                });
                let author = i16::from_le_bytes([operand[1], operand[2]]);
                tap.formatting_revision_author_index =
                    Some(u16::try_from(author).map_err(|_| {
                        DocError::Corrupted("sprmTPropRMark author index is negative".to_string())
                    })?);
                tap.formatting_revision_timestamp = Some(u32::from_le_bytes([
                    operand[3], operand[4], operand[5], operand[6],
                ]));
            },
            // Raw defaults preserve ShdNil as table-style inheritance.
            0x70 => self.parse_full_cell_shading(tap, sprm, 0, true)?,
            0x71 => self.parse_full_cell_shading(tap, sprm, 22, true)?,
            0x72 => self.parse_full_cell_shading(tap, sprm, 44, true)?,
            // sprmTWall (0x3668) - Preserve properties before tracked changes
            0x68 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 1 {
                    return Err(DocError::Corrupted(
                        "sprmTWall operand must contain exactly 1 byte".to_string(),
                    ));
                }
                tap.properties_preserved_for_revision = match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(DocError::Corrupted(
                            "sprmTWall must contain a Boolean8 value".to_string(),
                        ));
                    },
                };
            },
            // Other table SPRMs.
            _ => {
                // Unknown or unhandled SPRM - skip
            },
        }

        Ok(())
    }

    /// Parse table definition (sprmTDefTable - 0xD608).
    ///
    /// Format:
    /// - 1 byte: itcMac (cell count)
    /// - (itcMac+1) * 2 bytes: rgdxaCenter (cell boundaries)
    /// - itcMac * 20 bytes: rgtc (cell descriptors) [optional]
    fn parse_table_definition(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
    ) -> Result<()> {
        let data = sprm.operand_bytes();
        if data.is_empty() {
            return Err(DocError::Corrupted(
                "sprmTDefTable does not contain a column count".to_string(),
            ));
        }

        // Read cell count
        let itc_mac = binary_to_doc_result(read_byte(data, 0))? as usize;
        if itc_mac > 63 {
            return Err(DocError::Corrupted(
                "sprmTDefTable contains more than 63 columns".to_string(),
            ));
        }
        let start_of_tcs = 1 + ((itc_mac + 1) * 2);
        if data.len() < start_of_tcs || (data.len() - start_of_tcs) % 20 != 0 {
            return Err(DocError::Corrupted(
                "sprmTDefTable contains incomplete boundaries or cell descriptors".to_string(),
            ));
        }
        tap.cell_count = itc_mac;

        // Read cell boundaries (rgdxaCenter)
        let mut boundaries = Vec::with_capacity(itc_mac + 1);
        for i in 0..=itc_mac {
            let boundary_offset = 1 + (i * 2);
            if boundary_offset + 1 < data.len() {
                boundaries.push(binary_to_doc_result(read_i16_le(data, boundary_offset))?);
            }
        }
        if boundaries
            .iter()
            .any(|boundary| !(-31_680..=31_680).contains(boundary))
        {
            return Err(DocError::Corrupted(
                "sprmTDefTable cell boundary is outside the XAS range".to_string(),
            ));
        }
        if boundaries.windows(2).any(|pair| pair[0] > pair[1]) {
            return Err(DocError::Corrupted(
                "sprmTDefTable cell boundaries are not non-decreasing".to_string(),
            ));
        }
        tap.cell_boundaries = boundaries;

        // Calculate where cell descriptors start
        let has_tcs = start_of_tcs < data.len();

        // Read cell descriptors (TableCellDescriptor - TC)
        if has_tcs {
            let mut cell_props = Vec::with_capacity(itc_mac);
            for i in 0..itc_mac {
                let tc_offset = start_of_tcs + (i * 20); // Each TC is 20 bytes
                if tc_offset + 20 <= data.len() {
                    cell_props.push(self.parse_table_cell_descriptor(data, tc_offset)?);
                } else {
                    cell_props.push(CellProperties::default());
                }
            }
            tap.cell_properties = cell_props;
        } else {
            // No TC data - use defaults
            tap.cell_properties = vec![CellProperties::default(); itc_mac];
        }

        Ok(())
    }

    /// Parse a TableCellDescriptor (TC) structure.
    ///
    /// TC structure (20 bytes total):
    /// - bytes 0-1: flags (fVertical, fBackward, fRotateFont, fVertMerge, fVertRestart, etc.)
    /// - bytes 2-3: wWidth (preferred cell width)
    /// - bytes 4-7: brcTop (top border, 4 bytes)
    /// - bytes 8-11: brcLeft (left border, 4 bytes)
    /// - bytes 12-15: brcBottom (bottom border, 4 bytes)
    /// - bytes 16-19: brcRight (right border, 4 bytes)
    fn parse_table_cell_descriptor(&self, data: &[u8], offset: usize) -> Result<CellProperties> {
        let mut props = CellProperties::default();

        if offset + 20 > data.len() {
            return Ok(props);
        }

        // Read flags (bytes 0-1)
        let flags = binary_to_doc_result(read_u16_le(data, offset))?;

        // Bits 0-1: horizontal merge state.
        props.merge_status = match flags & 0x0003 {
            0 => CellMergeStatus::None,
            1 => CellMergeStatus::Merged,
            2 | 3 => CellMergeStatus::First,
            _ => unreachable!(),
        };

        // Bits 2-4: text flow.
        props.text_direction = match (flags >> 2) & 0x07 {
            0 => TextDirection::LrTb,
            1 => TextDirection::TbRl,
            3 => TextDirection::BtLr,
            4 => TextDirection::LrBt,
            5 => TextDirection::TbLr,
            value => {
                return Err(DocError::Corrupted(format!(
                    "TC80 contains invalid textFlow value {value}"
                )));
            },
        };

        // Bits 5-6: vertical merge state.
        props.vertical_merge_status = match (flags >> 5) & 0x03 {
            0 => VerticalMergeStatus::None,
            1 => VerticalMergeStatus::Merged,
            3 => VerticalMergeStatus::First,
            value => {
                return Err(DocError::Corrupted(format!(
                    "TC80 contains invalid vertMerge value {value}"
                )));
            },
        };

        // Bits 7-8: vertical alignment.
        let vert_align = (flags >> 7) & 0x03;
        props.vertical_alignment = match vert_align {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            value => {
                return Err(DocError::Corrupted(format!(
                    "TC80 contains invalid vertAlign value {value}"
                )));
            },
        };

        props.fit_text = flags & 0x1000 != 0;
        props.no_wrap = flags & 0x2000 != 0;
        props.hide_mark = flags & 0x4000 != 0;

        // Read preferred width (bytes 2-3)
        let w_width = binary_to_doc_result(read_i16_le(data, offset + 2))?;
        if w_width < 0 {
            return Err(DocError::Corrupted(
                "TC80 preferred width is negative".to_string(),
            ));
        }
        let fts_width = (flags >> 9) & 0x07;
        props.preferred_width = match fts_width {
            0 => None,
            1 => Some(TableWidth {
                value: w_width,
                width_type: WidthType::Auto,
            }),
            2 => Some(TableWidth {
                value: w_width,
                width_type: WidthType::Percentage,
            }),
            3 => Some(TableWidth {
                value: w_width,
                width_type: WidthType::Twips,
            }),
            value => {
                return Err(DocError::Corrupted(format!(
                    "TC80 contains invalid ftsWidth value {value}"
                )));
            },
        };

        // Read borders (4 bytes each)
        props.borders.top = Self::parse_border_code(data, offset + 4)?;
        props.borders.left = Self::parse_border_code(data, offset + 8)?;
        props.borders.bottom = Self::parse_border_code(data, offset + 12)?;
        props.borders.right = Self::parse_border_code(data, offset + 16)?;

        Ok(props)
    }

    /// Parse a BorderCode structure (4 bytes).
    ///
    /// BorderCode format:
    /// - byte 0: dptLineWidth (line width in 1/8 points)
    /// - byte 1: brcType (border type)
    /// - byte 2: ico (color index)
    /// - byte 3: spacing and effect flags
    fn parse_border_code(data: &[u8], offset: usize) -> Result<Option<BorderStyle>> {
        if offset + 4 > data.len() {
            return Ok(None);
        }

        let width = binary_to_doc_result(read_byte(data, offset))?;
        let border_type = binary_to_doc_result(read_byte(data, offset + 1))?;
        if data[offset..offset + 4] == [0xFF; 4] {
            return Ok(None);
        }
        let ico = binary_to_doc_result(read_byte(data, offset + 2))?;
        let effects = binary_to_doc_result(read_byte(data, offset + 3))?;

        // If width is 0 and type is 0, no border
        if width == 0 && border_type == 0 {
            return Ok(None);
        }

        let btype = Self::parse_border_type(border_type, false)?;

        let color = if ico == 0 {
            None
        } else if ico <= 16 {
            Self::ico_to_rgb(ico)
        } else {
            return Err(DocError::Corrupted(format!(
                "Brc80 contains invalid color index {ico}"
            )));
        };

        Ok(Some(BorderStyle {
            width,
            color,
            border_type: btype,
            spacing: effects & 0x1F,
            shadow: effects & 0x20 != 0,
            frame: effects & 0x40 != 0,
        }))
    }

    fn parse_border_type(border_type: u8, full: bool) -> Result<BorderType> {
        Ok(match border_type {
            0 => BorderType::None,
            1 => BorderType::Single,
            3 => BorderType::Double,
            5 => BorderType::Thick,
            6 => BorderType::Dotted,
            7 => BorderType::Dashed,
            8 => BorderType::DotDash,
            9 => BorderType::DotDotDash,
            10 => BorderType::Triple,
            11 => BorderType::ThinThickSmall,
            12 => BorderType::ThickThinSmall,
            13 => BorderType::ThinThickThinSmall,
            14 => BorderType::ThinThickMedium,
            15 => BorderType::ThickThinMedium,
            16 => BorderType::ThinThickThinMedium,
            17 => BorderType::ThinThickLarge,
            18 => BorderType::ThickThinLarge,
            19 => BorderType::ThinThickThinLarge,
            20 => BorderType::Wave,
            21 => BorderType::DoubleWave,
            22 => BorderType::DashSmall,
            23 => BorderType::DashDotStroked,
            24 => BorderType::Emboss,
            25 => BorderType::Engrave,
            26 if full => BorderType::Outset,
            27 if full => BorderType::Inset,
            value => {
                return Err(DocError::Corrupted(format!(
                    "DOC border contains invalid border type {value:#04x}"
                )));
            },
        })
    }

    /// Parse table borders (sprmTTableBorders - 0xD605).
    ///
    /// Contains 6 BorderCode structures (4 bytes each):
    /// - Top, Left, Bottom, Right, Horizontal, Vertical
    fn parse_table_borders(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 24 {
            return Err(DocError::Corrupted(
                "DOC TableBordersOperand80 must contain 24 bytes".to_string(),
            ));
        }

        // Parse 6 border codes (each 4 bytes)
        tap.border_top = Self::parse_border_code(operand, 0)?;
        tap.border_left = Self::parse_border_code(operand, 4)?;
        tap.border_bottom = Self::parse_border_code(operand, 8)?;
        tap.border_right = Self::parse_border_code(operand, 12)?;
        tap.border_horizontal = Self::parse_border_code(operand, 16)?;
        tap.border_vertical = Self::parse_border_code(operand, 20)?;

        Ok(())
    }

    fn parse_full_table_borders(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 48 {
            return Err(DocError::Corrupted(
                "DOC TableBordersOperand must contain 48 bytes".to_string(),
            ));
        }
        tap.border_top = Self::parse_full_border(&operand[0..8])?;
        tap.border_left = Self::parse_full_border(&operand[8..16])?;
        tap.border_bottom = Self::parse_full_border(&operand[16..24])?;
        tap.border_right = Self::parse_full_border(&operand[24..32])?;
        tap.border_horizontal = Self::parse_full_border(&operand[32..40])?;
        tap.border_vertical = Self::parse_full_border(&operand[40..48])?;
        Ok(())
    }

    fn parse_full_border(bytes: &[u8]) -> Result<Option<BorderStyle>> {
        if bytes.len() != 8 {
            return Err(DocError::Corrupted(
                "DOC Brc must contain exactly 8 bytes".to_string(),
            ));
        }
        if bytes[4..] == [0xFF; 4] {
            return Ok(None);
        }
        let width = bytes[4];
        let border_type = bytes[5];
        if width == 0 && border_type == 0 {
            return Ok(None);
        }
        let effects = bytes[6];
        Ok(Some(BorderStyle {
            width,
            color: Self::parse_colorref(&bytes[..4])?,
            border_type: Self::parse_border_type(border_type, true)?,
            spacing: effects & 0x1F,
            shadow: effects & 0x20 != 0,
            frame: effects & 0x40 != 0,
        }))
    }

    fn parse_cell_border_colors(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        operation: u16,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != tap.cell_properties.len() * 4 {
            return Err(DocError::Corrupted(
                "DOC cell border color array does not match the row".to_string(),
            ));
        }
        for (cell, colorref) in tap.cell_properties.iter_mut().zip(operand.chunks_exact(4)) {
            let border = match operation {
                0x1A => &mut cell.borders.top,
                0x1B => &mut cell.borders.left,
                0x1C => &mut cell.borders.bottom,
                0x1D => &mut cell.borders.right,
                _ => unreachable!(),
            };
            if colorref == [0xFF; 4] {
                *border = None;
            } else {
                let color = Self::parse_colorref(colorref)?;
                if let Some(border) = border {
                    border.color = color;
                }
            }
        }
        Ok(())
    }

    fn parse_cell_border_range(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        full: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        let expected_len = if full { 11 } else { 7 };
        if operand.len() != expected_len {
            return Err(DocError::Corrupted(format!(
                "DOC cell border range operand must contain {expected_len} bytes"
            )));
        }
        let first = operand[0] as usize;
        let limit = operand[1] as usize;
        if first >= tap.cell_properties.len() || limit < first || limit > tap.cell_properties.len()
        {
            return Err(DocError::Corrupted(
                "DOC cell border range exceeds the row".to_string(),
            ));
        }
        let sides = operand[2];
        let allowed_sides = if full { 0x3F } else { 0x0F };
        if sides & !allowed_sides != 0 {
            return Err(DocError::Corrupted(
                "DOC cell border range contains invalid side flags".to_string(),
            ));
        }
        let border = if full {
            Self::parse_full_border(&operand[3..])?
        } else {
            Self::parse_border_code(&operand[3..], 0)?
        };
        for cell in &mut tap.cell_properties[first..limit] {
            if sides & 0x01 != 0 {
                cell.borders.top = border;
            }
            if sides & 0x02 != 0 {
                cell.borders.left = border;
            }
            if sides & 0x04 != 0 {
                cell.borders.bottom = border;
            }
            if sides & 0x08 != 0 {
                cell.borders.right = border;
            }
            if sides & 0x10 != 0 {
                cell.borders.diagonal_down = border;
            }
            if sides & 0x20 != 0 {
                cell.borders.diagonal_up = border;
            }
        }
        Ok(())
    }

    fn cell_range(tap: &TableProperties, operand: &[u8]) -> Result<std::ops::Range<usize>> {
        if operand.len() < 2 {
            return Err(DocError::Corrupted(
                "DOC cell range operand is truncated".to_string(),
            ));
        }
        let first = operand[0] as usize;
        let limit = operand[1] as usize;
        if first >= tap.cell_properties.len() || limit < first || limit > tap.cell_properties.len()
        {
            return Err(DocError::Corrupted(
                "DOC cell property range exceeds the row".to_string(),
            ));
        }
        Ok(first..limit)
    }

    fn parse_cell_text_flow(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 4 {
            return Err(DocError::Corrupted(
                "DOC cell text-flow operand must contain 4 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let value = binary_to_doc_result(read_u16_le(operand, 2))?;
        let direction = match value {
            0 => TextDirection::LrTb,
            1 => TextDirection::TbRl,
            3 => TextDirection::BtLr,
            4 => TextDirection::LrBt,
            5 => TextDirection::TbLr,
            _ => {
                return Err(DocError::Corrupted(
                    "DOC cell text-flow value is invalid".to_string(),
                ));
            },
        };
        for cell in &mut tap.cell_properties[range] {
            cell.text_direction = direction;
        }
        Ok(())
    }

    fn parse_vertical_merge(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 || operand[0] as usize >= tap.cell_properties.len() {
            return Err(DocError::Corrupted(
                "DOC vertical-merge operand is invalid for the row".to_string(),
            ));
        }
        tap.cell_properties[operand[0] as usize].vertical_merge_status = match operand[1] {
            0 => VerticalMergeStatus::None,
            1 => VerticalMergeStatus::Merged,
            3 => VerticalMergeStatus::First,
            _ => {
                return Err(DocError::Corrupted(
                    "DOC vertical-merge flag is invalid".to_string(),
                ));
            },
        };
        Ok(())
    }

    fn parse_vertical_alignment(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 3 {
            return Err(DocError::Corrupted(
                "DOC vertical-alignment operand must contain 3 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let alignment = match operand[2] {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            _ => {
                return Err(DocError::Corrupted(
                    "DOC cell vertical-alignment value is invalid".to_string(),
                ));
            },
        };
        for cell in &mut tap.cell_properties[range] {
            cell.vertical_alignment = alignment;
        }
        Ok(())
    }

    fn parse_cell_width(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 5 {
            return Err(DocError::Corrupted(
                "DOC cell-width operand must contain 5 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let value = binary_to_doc_result(read_i16_le(operand, 3))?;
        let width = match operand[2] {
            0 => None,
            1 if value == 0 => Some(TableWidth {
                value,
                width_type: WidthType::Auto,
            }),
            2 if (0..=5_000).contains(&value) => Some(TableWidth {
                value,
                width_type: WidthType::Percentage,
            }),
            3 if (0..=31_680).contains(&value) => Some(TableWidth {
                value,
                width_type: WidthType::Twips,
            }),
            _ => {
                return Err(DocError::Corrupted(
                    "DOC preferred cell width has invalid units or value".to_string(),
                ));
            },
        };
        for cell in &mut tap.cell_properties[range] {
            cell.preferred_width = width;
        }
        Ok(())
    }

    fn parse_cell_range_bool(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        property: CellBoolProperty,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 3 || !matches!(operand[2], 0 | 1) {
            return Err(DocError::Corrupted(
                "DOC Boolean cell-range operand is invalid".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let value = operand[2] != 0;
        for cell in &mut tap.cell_properties[range] {
            match property {
                CellBoolProperty::FitText => cell.fit_text = value,
                CellBoolProperty::NoWrap => cell.no_wrap = value,
                CellBoolProperty::HideMark => cell.hide_mark = value,
            }
        }
        Ok(())
    }

    /// Handle cell insertion (sprmTInsert - 0x7621).
    ///
    /// Operand format (4 bytes):
    /// - byte 0: index (where to insert)
    /// - byte 1: count (how many cells to insert)
    /// - bytes 2-3: width (width of new cells in twips)
    fn handle_insert_cells(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 4 {
            return Err(DocError::Corrupted(
                "DOC cell-insertion operand must contain 4 bytes".to_string(),
            ));
        }
        let index = operand[0] as usize;
        let count = operand[1] as usize;
        let width = binary_to_doc_result(read_u16_le(operand, 2))?;
        if index > tap.cell_count || count == 0 || tap.cell_count + count > 63 || width > 31_680 {
            return Err(DocError::Corrupted(
                "DOC cell insertion has an invalid index, count, or width".to_string(),
            ));
        }
        let added_width = i32::from(width) * count as i32;
        let first_boundary = i32::from(*tap.cell_boundaries.first().unwrap_or(&0));
        let last_boundary = i32::from(*tap.cell_boundaries.last().unwrap_or(&0));
        if last_boundary - first_boundary + added_width > 31_680
            || last_boundary + added_width > 31_680
        {
            return Err(DocError::Corrupted(
                "DOC cell insertion makes the table wider than 31680 twips".to_string(),
            ));
        }

        let insertion_x = i32::from(tap.cell_boundaries[index]);
        let mut boundaries = Vec::with_capacity(tap.cell_count + count + 1);
        boundaries.extend(tap.cell_boundaries[..=index].iter().copied().map(i32::from));
        for inserted in 1..=count {
            boundaries.push(insertion_x + i32::from(width) * inserted as i32);
        }
        boundaries.extend(
            tap.cell_boundaries[index + 1..]
                .iter()
                .map(|boundary| i32::from(*boundary) + added_width),
        );
        tap.cell_boundaries = boundaries
            .into_iter()
            .map(|boundary| {
                i16::try_from(boundary).map_err(|_| {
                    DocError::Corrupted("DOC cell insertion overflows coordinates".to_string())
                })
            })
            .collect::<Result<Vec<_>>>()?;
        tap.cell_properties.splice(
            index..index,
            std::iter::repeat_n(CellProperties::default(), count),
        );
        tap.cell_count += count;

        Ok(())
    }

    fn handle_delete_cells(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(DocError::Corrupted(
                "DOC cell-deletion operand must contain 2 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        if range.len() >= tap.cell_count {
            return Err(DocError::Corrupted(
                "DOC cell deletion must leave at least one cell".to_string(),
            ));
        }
        if range.is_empty() {
            return Ok(());
        }
        let first = range.start;
        let limit = range.end;
        tap.cell_properties.drain(range);
        tap.cell_boundaries.drain(first..limit);
        tap.cell_count = tap.cell_properties.len();
        Ok(())
    }

    fn handle_column_width(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 4 {
            return Err(DocError::Corrupted(
                "DOC column-width operand must contain 4 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let width = binary_to_doc_result(read_u16_le(operand, 2))?;
        if width > 31_680 {
            return Err(DocError::Corrupted(
                "DOC column width is outside its allowed range".to_string(),
            ));
        }
        let mut boundaries: Vec<i32> = tap.cell_boundaries.iter().copied().map(i32::from).collect();
        for index in range {
            let current = boundaries[index + 1] - boundaries[index];
            let delta = i32::from(width) - current;
            for boundary in &mut boundaries[index + 1..] {
                *boundary += delta;
            }
        }
        if boundaries
            .iter()
            .any(|boundary| !(-31_680..=31_680).contains(boundary))
        {
            return Err(DocError::Corrupted(
                "DOC column widths produce a coordinate outside the XAS range".to_string(),
            ));
        }
        tap.cell_boundaries = boundaries
            .into_iter()
            .map(i16::try_from)
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(|_| {
                DocError::Corrupted("DOC column widths overflow coordinates".to_string())
            })?;
        Ok(())
    }

    fn handle_horizontal_merge(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        merge: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(DocError::Corrupted(
                "DOC horizontal merge operand must contain 2 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        for (offset, cell) in tap.cell_properties[range].iter_mut().enumerate() {
            cell.merge_status = if merge {
                if offset == 0 {
                    CellMergeStatus::First
                } else {
                    CellMergeStatus::Merged
                }
            } else {
                CellMergeStatus::None
            };
        }
        Ok(())
    }

    /// Parse cell padding (sprmTCellPaddingDefault - 0xD634).
    ///
    /// Format:
    /// - byte 0: itcFirst (first cell index)
    /// - byte 1: itcLim (limit cell index, exclusive)
    /// - byte 2: grfbrc (flags indicating which borders to apply padding to)
    /// - byte 3: ftsWidth (width type)
    /// - bytes 4-5: wWidth (padding width)
    fn parse_cell_padding(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 6 {
            return Err(DocError::Corrupted(
                "DOC cell padding operand must contain exactly 6 bytes".to_string(),
            ));
        }

        let itc_first = operand[0] as usize;
        let itc_lim = operand[1] as usize;
        let grf_brc = operand[2];
        let fts_width = operand[3];
        let w_width = binary_to_doc_result(read_u16_le(operand, 4))?;
        if grf_brc & !0x0F != 0 {
            return Err(DocError::Corrupted(
                "DOC cell padding side mask contains reserved bits".to_string(),
            ));
        }
        if !matches!(fts_width, 0x00 | 0x03) {
            return Err(DocError::Corrupted(
                "DOC cell padding width type must be ftsNil or ftsDxa".to_string(),
            ));
        }
        if (fts_width == 0x00 && w_width != 0) || w_width > 31_680 {
            return Err(DocError::Corrupted(
                "DOC cell padding width is outside its allowed range".to_string(),
            ));
        }

        let is_default = sprm.opcode == 0xD634;
        let range = if is_default {
            if itc_first != 0 || itc_lim != 1 {
                return Err(DocError::Corrupted(
                    "DOC default cell padding range must be 0..1".to_string(),
                ));
            }
            0..tap.cell_properties.len()
        } else {
            if itc_first >= tap.cell_properties.len()
                || itc_lim < itc_first
                || itc_lim > tap.cell_properties.len()
            {
                return Err(DocError::Corrupted(
                    "DOC cell padding range exceeds the table row".to_string(),
                ));
            }
            itc_first..itc_lim
        };
        let padding = (fts_width == 0x03).then_some(w_width as i16);

        // Apply padding to specified cells
        for c in range {
            let cell = &mut tap.cell_properties[c];

            // Apply padding based on grfbrc flags
            if (grf_brc & 0x01) != 0 {
                cell.padding_top = padding;
            }
            if (grf_brc & 0x02) != 0 {
                cell.padding_left = padding;
            }
            if (grf_brc & 0x04) != 0 {
                cell.padding_bottom = padding;
            }
            if (grf_brc & 0x08) != 0 {
                cell.padding_right = padding;
            }
        }

        Ok(())
    }

    fn parse_cell_spacing(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 6 || operand[0] != 0 || operand[1] != 1 || operand[2] != 0x0F {
            return Err(DocError::Corrupted(
                "DOC default cell spacing must target range 0..1 and all sides".to_string(),
            ));
        }
        let width = binary_to_doc_result(read_u16_le(operand, 4))?;
        if width > 15_840 {
            return Err(DocError::Corrupted(
                "DOC default cell spacing exceeds 15840 twips".to_string(),
            ));
        }
        tap.cell_spacing = match operand[3] {
            0 if width == 0 => None,
            3 => Some(CellSpacing {
                width,
                source: CellSpacingSource::Explicit,
            }),
            0x13 => Some(CellSpacing {
                width,
                source: CellSpacingSource::TableBorder,
            }),
            _ => {
                return Err(DocError::Corrupted(
                    "DOC default cell spacing has invalid units or value".to_string(),
                ));
            },
        };
        Ok(())
    }

    /// Convert ico (color index) to RGB.
    ///
    /// Based on POI's color index mapping.
    fn ico_to_rgb(ico: u8) -> Option<(u8, u8, u8)> {
        Some(match ico {
            0 => return None,
            1 => (0, 0, 0),
            2 => (0, 0, 255),
            3 => (0, 255, 255),
            4 => (0, 255, 0),
            5 => (255, 0, 255),
            6 => (255, 0, 0),
            7 => (255, 255, 0),
            8 => (255, 255, 255),
            9 => (0, 0, 128),
            10 => (0, 128, 128),
            11 => (0, 128, 0),
            12 => (128, 0, 128),
            13 => (128, 0, 0),
            14 => (128, 128, 0),
            15 => (128, 128, 128),
            16 => (192, 192, 192),
            _ => return None,
        })
    }

    /// Parse cell shading (`sprmTDefTableShd80`).
    ///
    /// This SPRM contains an array of ShadingDescriptor80 structures (2 bytes each),
    /// one for each cell in the table.
    ///
    /// Based on Apache POI's table shading handling.
    fn parse_cell_shading(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
    ) -> Result<()> {
        let bytes = sprm.operand_bytes();
        if bytes.len() % 2 != 0 || bytes.len() / 2 > tap.cell_properties.len() {
            return Err(DocError::Corrupted(
                "DOC Shd80 array has an invalid size for the table row".to_string(),
            ));
        }

        // Parse shading descriptors for each cell
        for (i, descriptor) in bytes.chunks_exact(2).enumerate() {
            let shd = binary_to_doc_result(read_u16_le(descriptor, 0))?;
            if shd == u16::MAX {
                tap.cell_properties[i].shading = None;
                tap.cell_properties[i].shading_inherits_from_style = false;
                tap.cell_properties[i].background_color = None;
                continue;
            }
            let ico_fore = (shd & 0x1F) as u8;
            let ico_back = ((shd >> 5) & 0x1F) as u8;
            let ipat = ((shd >> 10) & 0x3F) as u8;
            if ico_fore > 16 || ico_back > 16 {
                return Err(DocError::Corrupted(
                    "DOC Shd80 contains an invalid palette index".to_string(),
                ));
            }
            let pattern = ShadingPattern::from_u8(ipat).ok_or_else(|| {
                DocError::Corrupted("DOC Shd80 contains an invalid pattern index".to_string())
            })?;
            let shading = CellShading {
                foreground_color: Self::ico_to_rgb(ico_fore),
                background_color: Self::ico_to_rgb(ico_back),
                pattern,
            };
            tap.cell_properties[i].background_color = shading.background_color;
            tap.cell_properties[i].shading = Some(shading);
            tap.cell_properties[i].shading_inherits_from_style = false;
        }

        Ok(())
    }

    fn parse_full_cell_shading(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        first_cell: usize,
        raw: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() % 10 != 0 || operand.len() > 220 {
            return Err(DocError::Corrupted(
                "DOC table Shd array has an invalid byte count".to_string(),
            ));
        }
        let count = operand.len() / 10;
        let chunk_limit = if first_cell == 44 { 19 } else { 22 };
        if count > chunk_limit || first_cell.saturating_add(count) > tap.cell_properties.len() {
            return Err(DocError::Corrupted(
                "DOC table Shd array exceeds its cell chunk".to_string(),
            ));
        }
        for (offset, bytes) in operand.chunks_exact(10).enumerate() {
            Self::apply_full_shading(&mut tap.cell_properties[first_cell + offset], bytes, raw)?;
        }
        Ok(())
    }

    fn parse_full_cell_shading_range(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        odd_only: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 12 {
            return Err(DocError::Corrupted(
                "DOC table range shading operand must contain exactly 12 bytes".to_string(),
            ));
        }
        let first = operand[0] as usize;
        let limit = operand[1] as usize;
        if first >= tap.cell_properties.len() || limit < first || limit > tap.cell_properties.len()
        {
            return Err(DocError::Corrupted(
                "DOC table range shading exceeds the row".to_string(),
            ));
        }
        let step = if odd_only { 2 } else { 1 };
        for index in (first..limit).step_by(step) {
            Self::apply_full_shading(&mut tap.cell_properties[index], &operand[2..], false)?;
        }
        Ok(())
    }

    fn parse_full_table_shading(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 10 {
            return Err(DocError::Corrupted(
                "DOC whole-table shading operand must contain exactly 10 bytes".to_string(),
            ));
        }
        for cell in &mut tap.cell_properties {
            Self::apply_full_shading(cell, operand, false)?;
        }
        Ok(())
    }

    fn apply_full_shading(cell: &mut CellProperties, bytes: &[u8], raw: bool) -> Result<()> {
        if bytes.len() != 10 {
            return Err(DocError::Corrupted(
                "DOC Shd must contain exactly 10 bytes".to_string(),
            ));
        }
        let is_nil = bytes[..8].iter().all(|byte| *byte == 0xFF) && bytes[8..] == [0, 0];
        if is_nil {
            cell.shading = None;
            cell.background_color = None;
            cell.shading_inherits_from_style = raw;
            return Ok(());
        }
        let is_auto =
            bytes[..4] == [0, 0, 0, 0xFF] && bytes[4..8] == [0, 0, 0, 0xFF] && bytes[8..] == [0, 0];
        if is_auto {
            cell.shading = None;
            cell.background_color = None;
            cell.shading_inherits_from_style = false;
            return Ok(());
        }
        let foreground_color = Self::parse_colorref(&bytes[..4])?;
        let background_color = Self::parse_colorref(&bytes[4..8])?;
        let pattern_value = binary_to_doc_result(read_u16_le(bytes, 8))?;
        let pattern = u8::try_from(pattern_value)
            .ok()
            .and_then(ShadingPattern::from_u8)
            .ok_or_else(|| {
                DocError::Corrupted("DOC Shd contains an invalid pattern index".to_string())
            })?;
        cell.shading = Some(CellShading {
            foreground_color,
            background_color,
            pattern,
        });
        cell.background_color = background_color;
        cell.shading_inherits_from_style = false;
        Ok(())
    }

    fn parse_colorref(bytes: &[u8]) -> Result<Option<(u8, u8, u8)>> {
        if bytes.len() != 4 || !matches!(bytes[3], 0x00 | 0xFF) {
            return Err(DocError::Corrupted(
                "DOC COLORREF has an invalid automatic-color flag".to_string(),
            ));
        }
        Ok((bytes[3] == 0).then_some((bytes[0], bytes[1], bytes[2])))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
        assert_eq!(tap.cell_properties[1].background_color, Some((255, 0, 0)));
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
    fn rejects_malformed_core_row_geometry_and_pagination() {
        let arena = Bump::new();
        let parser = TapParser::new(&arena);
        let parse_with = |opcode, operand: &[u8]| {
            let mut grpprl = table_definition_grpprl(&[1, 0, 0, 100, 0]);
            append_fixed_sprm(&mut grpprl, opcode, operand);
            parser.parse_tap(&grpprl)
        };

        assert!(parse_with(0x5400, &[3, 0]).is_err());
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
        let tap = parser.parse_tap(&grpprl).unwrap();
        assert_eq!(tap.has_formatting_revision, Some(true));
        assert_eq!(tap.formatting_revision_author_index, Some(1));
        assert_eq!(tap.formatting_revision_timestamp, Some(timestamp));
        assert!(tap.properties_preserved_for_revision);

        let invalid_wall = [0x68, 0x36, 2];
        assert!(parser.parse_tap(&invalid_wall).is_err());
    }
}
