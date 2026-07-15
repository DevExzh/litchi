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
    BorderStyle, BorderType, CellMergeStatus, CellProperties, TableJustification, TableProperties,
    TableWidth, TextDirection, VerticalAlignment, VerticalMergeStatus, WidthType,
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
            // sprmTJc (0x5400) - Table justification
            0x00 => {
                if let Some(jc) = sprm.operand_byte() {
                    tap.justification = match jc {
                        0 => TableJustification::Left,
                        1 => TableJustification::Center,
                        2 => TableJustification::Right,
                        _ => TableJustification::Left,
                    };
                }
            },
            // sprmTDxaLeft (0x9601) - Table indent from left
            0x01 => {
                if let Some(offset) = sprm.operand_word() {
                    let adjust = offset as i16
                        - (tap.cell_boundaries.first().copied().unwrap_or(0) + tap.gap_half);
                    for boundary in &mut tap.cell_boundaries {
                        *boundary += adjust;
                    }
                }
            },
            // sprmTDxaGapHalf (0x9602) - Half the width of spacing between cells
            0x02 => {
                if let Some(gap) = sprm.operand_word() {
                    if !tap.cell_boundaries.is_empty() {
                        let adjust = tap.gap_half - gap as i16;
                        tap.cell_boundaries[0] += adjust;
                    }
                    tap.gap_half = gap as i16;
                }
            },
            // sprmTFCantSplit (0x3403) - Row can't be split across pages
            0x03 => {
                if let Some(flag) = sprm.operand_byte() {
                    tap.allow_row_break = flag == 0;
                }
            },
            // sprmTTableHeader (0x3404) - Row is header row
            0x04 => {
                if let Some(flag) = sprm.operand_byte() {
                    tap.is_header_row = flag != 0;
                }
            },
            // sprmTTableBorders (0xD605) - Table borders
            0x05 => {
                self.parse_table_borders(tap, sprm, grpprl)?;
            },
            // 0x06 - obsolete (Word 1.x)
            0x06 => {},
            // sprmTDyaRowHeight (0x9407) - Row height
            0x07 => {
                if let Some(height) = sprm.operand_word() {
                    tap.row_height = Some(height as i16);
                }
            },
            // sprmTDefTable (0xD608) - Table definition
            0x08 => {
                self.parse_table_definition(tap, sprm, grpprl)?;
            },
            // sprmTDefTableShd (0xD609 / 0xD60A) - Table shading
            0x09 | 0x0A => {
                // Parse cell shading information
                // Format: variable length array of ShadingDescriptor80 (2 bytes each)
                // Based on Apache POI's handling of sprmTDefTableShd
                self.parse_cell_shading(tap, sprm, grpprl)?;
            },
            // sprmTTlp (0x740B) - Table look specifier
            0x0B => {
                // Table look specifier for table styles
                // This is a complex structure that defines table style properties
                // For basic parsing, we can skip this as it's mainly for styling
                if let Some(tlp) = sprm.operand_dword() {
                    // TLP structure contains bit flags for various table style options
                    // Not critical for basic text extraction
                    let _ = tlp;
                }
            },
            // sprmTInsert (0x7621) - Insert cells
            0x21 => {
                self.handle_insert_cells(tap, sprm)?;
            },
            // sprmTCellPaddingDefault (0xD634) - Default cell padding
            0x34 => {
                self.parse_cell_padding(tap, sprm, grpprl)?;
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
            // Other table SPRMs (0x22-0x2C, etc.)
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
    /// - byte 2-3: ico (color index) or RGB color
    fn parse_border_code(data: &[u8], offset: usize) -> Result<Option<BorderStyle>> {
        if offset + 4 > data.len() {
            return Ok(None);
        }

        let width = binary_to_doc_result(read_byte(data, offset))?;
        let border_type = binary_to_doc_result(read_byte(data, offset + 1))?;
        let color_word = binary_to_doc_result(read_u16_le(data, offset + 2))?;

        // If width is 0 and type is 0, no border
        if width == 0 && border_type == 0 {
            return Ok(None);
        }

        let btype = match border_type {
            0 => BorderType::None,
            1 => BorderType::Single,
            2 => BorderType::Thick,
            3 => BorderType::Double,
            5 => BorderType::Dotted,
            6 => BorderType::Dashed,
            7 => BorderType::DotDash,
            8 => BorderType::DotDotDash,
            9 => BorderType::Triple,
            10 => BorderType::ThinThickSmall,
            11 => BorderType::ThickThinSmall,
            12 => BorderType::ThinThickThinSmall,
            _ => BorderType::Single,
        };

        // Color extraction from Word color format (based on POI)
        let color = if color_word == 0 || color_word == 0xFFFF {
            None
        } else {
            Some((
                (color_word & 0x1F) as u8 * 8,         // Red (5 bits)
                ((color_word >> 5) & 0x1F) as u8 * 8,  // Green (5 bits)
                ((color_word >> 10) & 0x1F) as u8 * 8, // Blue (5 bits)
            ))
        };

        Ok(Some(BorderStyle {
            width,
            color,
            border_type: btype,
        }))
    }

    /// Parse table borders (sprmTTableBorders - 0xD605).
    ///
    /// Contains 6 BorderCode structures (4 bytes each):
    /// - Top, Left, Bottom, Right, Horizontal, Vertical
    fn parse_table_borders(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        grpprl: &[u8],
    ) -> Result<()> {
        let offset = sprm.offset + 3; // Skip SPRM header
        if offset + 24 > grpprl.len() {
            return Ok(());
        }

        // Parse 6 border codes (each 4 bytes)
        tap.border_top = Self::parse_border_code(grpprl, offset)?;
        tap.border_left = Self::parse_border_code(grpprl, offset + 4)?;
        tap.border_bottom = Self::parse_border_code(grpprl, offset + 8)?;
        tap.border_right = Self::parse_border_code(grpprl, offset + 12)?;
        tap.border_horizontal = Self::parse_border_code(grpprl, offset + 16)?;
        tap.border_vertical = Self::parse_border_code(grpprl, offset + 20)?;

        Ok(())
    }

    /// Handle cell insertion (sprmTInsert - 0x7621).
    ///
    /// Operand format (4 bytes):
    /// - byte 0: index (where to insert)
    /// - byte 1: count (how many cells to insert)
    /// - bytes 2-3: width (width of new cells in twips)
    fn handle_insert_cells(&self, tap: &mut TableProperties, sprm: &Sprm) -> Result<()> {
        if let Some(operand) = sprm.operand_dword() {
            let index = ((operand >> 24) & 0xFF) as usize;
            let count = ((operand >> 16) & 0xFF) as usize;
            let width = (operand & 0xFFFF) as i16;

            let itc_mac = tap.cell_count;
            let insert_at = index.min(itc_mac);

            // Create new arrays with space for inserted cells
            let mut new_boundaries = Vec::with_capacity(itc_mac + count + 1);
            let mut new_cells = Vec::with_capacity(itc_mac + count);

            // Copy boundaries before insertion point
            new_boundaries.extend_from_slice(&tap.cell_boundaries[..=insert_at]);

            // Copy cells before insertion point
            if insert_at < tap.cell_properties.len() {
                new_cells.extend_from_slice(&tap.cell_properties[..insert_at]);
            }

            // Insert new cells
            for _i in 0..count {
                let prev_boundary = new_boundaries.last().copied().unwrap_or(0);
                new_boundaries.push(prev_boundary + width);
                new_cells.push(CellProperties::default());
            }

            // Copy remaining boundaries and cells
            if insert_at < tap.cell_boundaries.len() {
                new_boundaries.extend_from_slice(&tap.cell_boundaries[insert_at..]);
            }
            if insert_at < tap.cell_properties.len() {
                new_cells.extend_from_slice(&tap.cell_properties[insert_at..]);
            }

            tap.cell_boundaries = new_boundaries;
            tap.cell_properties = new_cells;
            tap.cell_count = itc_mac + count;
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
        grpprl: &[u8],
    ) -> Result<()> {
        let offset = sprm.offset + 3; // Skip SPRM header
        if offset + 6 > grpprl.len() {
            return Ok(());
        }

        let itc_first = binary_to_doc_result(read_byte(grpprl, offset))? as usize;
        let itc_lim = binary_to_doc_result(read_byte(grpprl, offset + 1))? as usize;
        let grf_brc = binary_to_doc_result(read_byte(grpprl, offset + 2))?;
        let _fts_width = binary_to_doc_result(read_byte(grpprl, offset + 3))?;
        let w_width = binary_to_doc_result(read_u16_le(grpprl, offset + 4))? as i16;

        // Apply padding to specified cells
        for c in itc_first..itc_lim {
            if c >= tap.cell_properties.len() {
                break;
            }

            let cell = &mut tap.cell_properties[c];

            // Apply padding based on grfbrc flags
            if (grf_brc & 0x01) != 0 {
                cell.padding_top = Some(w_width);
            }
            if (grf_brc & 0x02) != 0 {
                cell.padding_left = Some(w_width);
            }
            if (grf_brc & 0x04) != 0 {
                cell.padding_bottom = Some(w_width);
            }
            if (grf_brc & 0x08) != 0 {
                cell.padding_right = Some(w_width);
            }
        }

        Ok(())
    }

    /// Convert ico (color index) to RGB.
    ///
    /// Based on POI's color index mapping.
    fn ico_to_rgb(ico: u8) -> (u8, u8, u8) {
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
            _ => (0, 0, 0),        // Default to black
        }
    }

    /// Parse cell shading (sprmTDefTableShd).
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
        let shd_size = 2; // ShadingDescriptor80 is 2 bytes

        // Parse shading descriptors for each cell
        let num_shd = bytes.len() / shd_size;
        for i in 0..num_shd.min(tap.cell_count) {
            let offset = i * shd_size;
            if offset + 1 < bytes.len()
                && let Ok(shd) = read_u16_le(bytes, offset)
            {
                // Parse Shd80 structure
                let ico_fore = (shd & 0x1F) as u8;
                let ico_back = ((shd >> 5) & 0x1F) as u8;
                let ipat = ((shd >> 10) & 0x3F) as u8;

                // Apply shading to cell if pattern is not Clear
                if ipat != 0 && i < tap.cell_properties.len() {
                    let _fg_color = Self::ico_to_rgb(ico_fore);
                    let _bg_color = Self::ico_to_rgb(ico_back);
                    // Store shading info in cell properties if needed
                    // For now, we mainly extract structure without applying
                    let _ = ipat;
                }
            }
        }

        Ok(())
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
    fn test_border_code_parsing() {
        let data = vec![
            0x08, // width = 8 (1 point)
            0x01, // type = single
            0x00, 0x00, // color = black
        ];

        let border = TapParser::parse_border_code(&data, 0).unwrap();
        assert!(border.is_some());
        let border = border.unwrap();
        assert_eq!(border.width, 8);
        assert_eq!(border.border_type, BorderType::Single);
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
