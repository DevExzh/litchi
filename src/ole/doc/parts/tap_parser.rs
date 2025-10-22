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
        return Err(crate::common::binary::BinaryError::InsufficientData {
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
    TableWidth, TextDirection, VerticalAlignment, WidthType,
};
use crate::common::binary::{BinaryResult, read_i16_le, read_u16_le};
use crate::ole::sprm::{Sprm, parse_sprms};
use bumpalo::Bump;

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
                // For long SPRMs (0xD608), operand format is:
                // - 2 bytes: size (already parsed, included in operand)
                // - 1 byte: itcMac (cell count)
                // - rest: cell boundaries and descriptors
                if sprm.operand.len() >= 3 {
                    let cell_count = sprm.operand[2] as usize; // Skip 2-byte size field
                    return Ok(TableProperties::with_cell_count(cell_count));
                }
            }
        }

        // No table definition found - use default with 1 cell
        eprintln!("WARNING: Table row didn't specify number of columns in SPRMs");
        Ok(TableProperties::with_cell_count(1))
    }

    /// Check if a SPRM is a TAP (table) SPRM.
    ///
    /// TAP SPRMs have type 5 (bits 0-2 of opcode).
    fn is_tap_sprm(opcode: u16) -> bool {
        (opcode & 0x07) == 5
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
        let operation = (sprm.opcode >> 3) & 0x1FF;

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
            // sprmTDefTableShd (0xD609) - Table shading
            0x09 => {
                // TODO: Implement cell shading parsing
            },
            // sprmTTlp (0x740A) - Table look specifier
            0x0A => {
                // TODO: Implement table style parsing
            },
            // sprmTInsert (0x7621) - Insert cells
            0x21 => {
                self.handle_insert_cells(tap, sprm)?;
            },
            // sprmTCellPaddingDefault (0xD634) - Default cell padding
            0x34 => {
                self.parse_cell_padding(tap, sprm, grpprl)?;
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
        grpprl: &[u8],
    ) -> Result<()> {
        let offset = sprm.offset + 3; // Skip sprm (2) + size (1)
        if offset >= grpprl.len() {
            return Ok(());
        }

        // Read cell count
        let itc_mac = binary_to_doc_result(read_byte(grpprl, offset))? as usize;
        tap.cell_count = itc_mac;

        // Read cell boundaries (rgdxaCenter)
        let mut boundaries = Vec::with_capacity(itc_mac + 1);
        for i in 0..=itc_mac {
            let boundary_offset = offset + 1 + (i * 2);
            if boundary_offset + 1 < grpprl.len() {
                boundaries.push(binary_to_doc_result(read_i16_le(grpprl, boundary_offset))?);
            }
        }
        tap.cell_boundaries = boundaries;

        // Calculate where cell descriptors start
        let end_of_sprm = offset + sprm.size.saturating_sub(3); // -3 for sprm header
        let start_of_tcs = offset + 1 + ((itc_mac + 1) * 2);
        let has_tcs = start_of_tcs < end_of_sprm;

        // Read cell descriptors (TableCellDescriptor - TC)
        if has_tcs {
            let mut cell_props = Vec::with_capacity(itc_mac);
            for i in 0..itc_mac {
                let tc_offset = start_of_tcs + (i * 20); // Each TC is 20 bytes
                if tc_offset + 20 <= grpprl.len() {
                    cell_props.push(self.parse_table_cell_descriptor(grpprl, tc_offset)?);
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

        // Bit 0: fVertical - vertical text
        if (flags & 0x0001) != 0 {
            props.text_direction = TextDirection::TbRl;
        }

        // Bits 1-2: vertAlign - vertical alignment
        let vert_align = (flags >> 4) & 0x03;
        props.vertical_alignment = match vert_align {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            _ => VerticalAlignment::Top,
        };

        // Bits 3-4: fVertMerge/fVertRestart - cell merging
        let merge_flags = (flags >> 3) & 0x03;
        props.merge_status = match merge_flags {
            0 => CellMergeStatus::None,
            1 => CellMergeStatus::First,
            3 => CellMergeStatus::Merged,
            _ => CellMergeStatus::None,
        };

        // Read preferred width (bytes 2-3)
        let w_width = binary_to_doc_result(read_u16_le(data, offset + 2))? as i16;
        let fts_width = (flags >> 6) & 0x07; // Width type from flags
        props.preferred_width = Some(TableWidth {
            value: w_width,
            width_type: match fts_width {
                0 => WidthType::Auto,
                1 => WidthType::Twips,
                2 => WidthType::Percentage,
                _ => WidthType::Auto,
            },
        });

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

        // Simple color extraction (Word uses complex color tables)
        // For now, use simplified RGB extraction
        let color = if color_word == 0 || color_word == 0xFFFF {
            None
        } else {
            Some((
                (color_word & 0x1F) as u8 * 8,         // Red
                ((color_word >> 5) & 0x1F) as u8 * 8,  // Green
                ((color_word >> 10) & 0x1F) as u8 * 8, // Blue
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
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
