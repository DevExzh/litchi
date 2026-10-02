//! Legacy and full-color table border codecs.

use super::prelude::*;

impl TapParser<'_> {
    /// Parse a `BorderCode` structure (4 bytes).
    ///
    /// `BorderCode` format:
    /// - byte 0: dptLineWidth (line width in 1/8 points)
    /// - byte 1: brcType (border type)
    /// - byte 2: ico (color index)
    /// - byte 3: spacing and effect flags
    ///
    /// An undefined `brcType` reads as [`BorderType::None`] and an `ico`
    /// above 16 as automatic color, each recorded against `opcode`.
    pub(in crate::parts::tap_parser) fn parse_border_code(
        data: &[u8],
        offset: usize,
        opcode: u16,
        tolerance: &mut ToleranceReport,
    ) -> Result<Option<BorderStyle>> {
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

        let btype = Self::parse_border_type(border_type, false, opcode, tolerance);

        // An Ico outside 0..=16 names no color; like LibreOffice, read it as
        // automatic (the same value as ico 0) and record the repair.
        if ico > 16 {
            tolerance.record(SprmDefect::BorderColor, opcode);
        }
        let color = Self::ico_to_rgb(ico);

        Ok(Some(BorderStyle {
            width,
            color,
            border_type: btype,
            spacing: effects & 0x1F,
            shadow: effects & 0x20 != 0,
            frame: effects & 0x40 != 0,
        }))
    }

    /// Decode a `brcType`. MS-DOC 2.9.17's restriction on 0x1A/0x1B in a
    /// `Brc80` is a writer constraint; an undefined value maps to
    /// [`BorderType::None`] — LibreOffice's `ConvertBorderStyleFromWord`
    /// default — and is recorded against `opcode`.
    pub(in crate::parts::tap_parser) fn parse_border_type(
        border_type: u8,
        full: bool,
        opcode: u16,
        tolerance: &mut ToleranceReport,
    ) -> BorderType {
        match border_type {
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
            _ => {
                tolerance.record(SprmDefect::BorderType, opcode);
                BorderType::None
            },
        }
    }

    /// Parse table borders (sprmTTableBorders - 0xD605).
    ///
    /// Contains 6 `BorderCode` structures (4 bytes each):
    /// - Top, Left, Bottom, Right, Horizontal, Vertical

    pub(in crate::parts::tap_parser) fn parse_table_borders(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
        tolerance: &mut ToleranceReport,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 24 {
            return Err(PackageError::Corrupted(
                "DOC TableBordersOperand80 must contain 24 bytes".to_string(),
            ));
        }

        // Parse 6 border codes (each 4 bytes)
        let opcode = sprm.opcode;
        tap.border_top = Self::parse_border_code(operand, 0, opcode, tolerance)?;
        tap.border_left = Self::parse_border_code(operand, 4, opcode, tolerance)?;
        tap.border_bottom = Self::parse_border_code(operand, 8, opcode, tolerance)?;
        tap.border_right = Self::parse_border_code(operand, 12, opcode, tolerance)?;
        tap.border_horizontal = Self::parse_border_code(operand, 16, opcode, tolerance)?;
        tap.border_vertical = Self::parse_border_code(operand, 20, opcode, tolerance)?;

        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_full_table_borders(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        tolerance: &mut ToleranceReport,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 48 {
            return Err(PackageError::Corrupted(
                "DOC TableBordersOperand must contain 48 bytes".to_string(),
            ));
        }
        let opcode = sprm.opcode;
        tap.border_top = Self::parse_full_border(&operand[0..8], opcode, tolerance)?;
        tap.border_left = Self::parse_full_border(&operand[8..16], opcode, tolerance)?;
        tap.border_bottom = Self::parse_full_border(&operand[16..24], opcode, tolerance)?;
        tap.border_right = Self::parse_full_border(&operand[24..32], opcode, tolerance)?;
        tap.border_horizontal = Self::parse_full_border(&operand[32..40], opcode, tolerance)?;
        tap.border_vertical = Self::parse_full_border(&operand[40..48], opcode, tolerance)?;
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_full_border(
        bytes: &[u8],
        opcode: u16,
        tolerance: &mut ToleranceReport,
    ) -> Result<Option<BorderStyle>> {
        if bytes.len() != 8 {
            return Err(PackageError::Corrupted(
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
            border_type: Self::parse_border_type(border_type, true, opcode, tolerance),
            spacing: effects & 0x1F,
            shadow: effects & 0x20 != 0,
            frame: effects & 0x40 != 0,
        }))
    }

    pub(in crate::parts::tap_parser) fn parse_cell_border_colors(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        operation: u16,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != tap.cell_properties.len() * 4 {
            return Err(PackageError::Corrupted(
                "DOC cell border color array does not match the row".to_string(),
            ));
        }
        for (cell, colorref) in tap
            .cell_properties
            .iter_mut()
            .zip(operand.as_chunks::<4>().0.iter())
        {
            let (border, direct) = match operation {
                0x1A => (&mut cell.borders.top, &mut cell.direct_style.border_top),
                0x1B => (&mut cell.borders.left, &mut cell.direct_style.border_left),
                0x1C => (
                    &mut cell.borders.bottom,
                    &mut cell.direct_style.border_bottom,
                ),
                0x1D => (&mut cell.borders.right, &mut cell.direct_style.border_right),
                _ => unreachable!(),
            };
            *direct = true;
            if *colorref == [0xFF; 4] {
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

    pub(in crate::parts::tap_parser) fn parse_cell_border_range(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        full: bool,
        tolerance: &mut ToleranceReport,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        let expected_len = if full { 11 } else { 7 };
        if operand.len() != expected_len {
            return Err(PackageError::Corrupted(format!(
                "DOC cell border range operand must contain {expected_len} bytes"
            )));
        }
        let first = operand[0] as usize;
        let limit = operand[1] as usize;
        if first >= tap.cell_properties.len() || limit < first || limit > tap.cell_properties.len()
        {
            return Err(PackageError::Corrupted(
                "DOC cell border range exceeds the row".to_string(),
            ));
        }
        let mut sides = operand[2];
        let allowed_sides = if full { 0x3F } else { 0x0F };
        if sides & !allowed_sides != 0 {
            // MS-DOC 2.9.304 defines only the low four side bits on the legacy
            // operand (0x3F on the full one, MS-DOC 2.9.305), but producers
            // set the diagonal and stray bits anyway. Word and LibreOffice
            // test the defined bits and ignore the rest, so the read masks
            // rather than rejects and records the repair.
            tolerance.record(SprmDefect::CellBorderSideFlags, sprm.opcode);
            sides &= allowed_sides;
        }
        let border = if full {
            Self::parse_full_border(&operand[3..], sprm.opcode, tolerance)?
        } else {
            Self::parse_border_code(&operand[3..], 0, sprm.opcode, tolerance)?
        };
        for cell in &mut tap.cell_properties[first..limit] {
            if sides & 0x01 != 0 {
                cell.borders.top = border;
                cell.direct_style.border_top = true;
            }
            if sides & 0x02 != 0 {
                cell.borders.left = border;
                cell.direct_style.border_left = true;
            }
            if sides & 0x04 != 0 {
                cell.borders.bottom = border;
                cell.direct_style.border_bottom = true;
            }
            if sides & 0x08 != 0 {
                cell.borders.right = border;
                cell.direct_style.border_right = true;
            }
            if sides & 0x10 != 0 {
                cell.borders.diagonal_down = border;
                cell.direct_style.border_diagonal_down = true;
            }
            if sides & 0x20 != 0 {
                cell.borders.diagonal_up = border;
                cell.direct_style.border_diagonal_up = true;
            }
        }
        Ok(())
    }
}
