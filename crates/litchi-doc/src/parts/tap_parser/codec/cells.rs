//! Cell-range, width, padding, and style-cell codecs.

use super::prelude::*;

impl TapParser<'_> {
    pub(in crate::parts::tap_parser) fn cell_range(
        tap: &TableProperties,
        operand: &[u8],
    ) -> Result<std::ops::Range<usize>> {
        if operand.len() < 2 {
            return Err(PackageError::Corrupted(
                "DOC cell range operand is truncated".to_string(),
            ));
        }
        let first = operand[0] as usize;
        let limit = operand[1] as usize;
        if first >= tap.cell_properties.len() || limit < first || limit > tap.cell_properties.len()
        {
            return Err(PackageError::Corrupted(
                "DOC cell property range exceeds the row".to_string(),
            ));
        }
        Ok(first..limit)
    }

    pub(in crate::parts::tap_parser) fn parse_cell_width(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 5 {
            return Err(PackageError::Corrupted(
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
                return Err(PackageError::Corrupted(
                    "DOC preferred cell width has invalid units or value".to_string(),
                ));
            },
        };
        for cell in &mut tap.cell_properties[range] {
            cell.preferred_width = width;
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_cell_padding(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 6 {
            return Err(PackageError::Corrupted(
                "DOC cell padding operand must contain exactly 6 bytes".to_string(),
            ));
        }

        let itc_first = operand[0] as usize;
        let itc_lim = operand[1] as usize;
        let grf_brc = operand[2];
        let fts_width = operand[3];
        let w_width = binary_to_doc_result(read_u16_le(operand, 4))?;
        if grf_brc & !0x0F != 0 {
            return Err(PackageError::Corrupted(
                "DOC cell padding side mask contains reserved bits".to_string(),
            ));
        }
        if !matches!(fts_width, 0x00 | 0x03) {
            return Err(PackageError::Corrupted(
                "DOC cell padding width type must be ftsNil or ftsDxa".to_string(),
            ));
        }
        if (fts_width == 0x00 && w_width != 0) || w_width > 31_680 {
            return Err(PackageError::Corrupted(
                "DOC cell padding width is outside its allowed range".to_string(),
            ));
        }

        let is_default = sprm.opcode == 0xD634;
        let range = if is_default {
            if itc_first != 0 || itc_lim != 1 {
                return Err(PackageError::Corrupted(
                    "DOC default cell padding range must be 0..1".to_string(),
                ));
            }
            0..tap.cell_properties.len()
        } else {
            if itc_first >= tap.cell_properties.len()
                || itc_lim < itc_first
                || itc_lim > tap.cell_properties.len()
            {
                return Err(PackageError::Corrupted(
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
                cell.direct_style.padding_top = true;
            }
            if (grf_brc & 0x02) != 0 {
                cell.padding_left = padding;
                cell.direct_style.padding_left = true;
            }
            if (grf_brc & 0x04) != 0 {
                cell.padding_bottom = padding;
                cell.direct_style.padding_bottom = true;
            }
            if (grf_brc & 0x08) != 0 {
                cell.padding_right = padding;
                cell.direct_style.padding_right = true;
            }
        }

        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_style_cell_padding(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 6
            || operand[0] != 0
            || operand[1] != 1
            || operand[2] == 0
            || operand[2] & !0x0F != 0
            || operand[3] != 0x03
        {
            return Err(PackageError::Corrupted(
                "sprmTCellPaddingStyle contains an invalid CSSA".to_string(),
            ));
        }
        let width = binary_to_doc_result(read_u16_le(operand, 4))?;
        if width > 31_680 {
            return Err(PackageError::Corrupted(
                "sprmTCellPaddingStyle exceeds 31680 twips".to_string(),
            ));
        }
        for (mask, padding) in [
            (0x01, &mut tap.style_defaults.padding_top),
            (0x02, &mut tap.style_defaults.padding_left),
            (0x04, &mut tap.style_defaults.padding_bottom),
            (0x08, &mut tap.style_defaults.padding_right),
        ] {
            if operand[2] & mask != 0 {
                *padding = Some(width);
            }
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_style_border(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        operation: u16,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 8 || operand[4..] == [0xFF; 4] {
            return Err(PackageError::Corrupted(
                "DOC table-style border must contain a non-nil 8-byte Brc".to_string(),
            ));
        }
        let border = match Self::parse_full_border(operand)? {
            Some(border) => TableStyleBorder::Border(border),
            None => TableStyleBorder::NoBorder,
        };
        let target = match operation {
            0x7F => &mut tap.style_defaults.border_top,
            0x80 => &mut tap.style_defaults.border_bottom,
            0x81 => &mut tap.style_defaults.border_left,
            0x82 => &mut tap.style_defaults.border_right,
            0x83 => &mut tap.style_defaults.border_inside_horizontal,
            0x84 => &mut tap.style_defaults.border_inside_vertical,
            0x85 => &mut tap.style_defaults.border_diagonal_down,
            0x86 => &mut tap.style_defaults.border_diagonal_up,
            _ => unreachable!(),
        };
        *target = Some(border);
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_style_shading(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 10 {
            return Err(PackageError::Corrupted(
                "DOC table-style shading must contain a 10-byte Shd".to_string(),
            ));
        }
        let is_nil = operand[..8].iter().all(|byte| *byte == 0xFF) && operand[8..] == [0, 0];
        if is_nil {
            return Ok(());
        }
        let is_auto = operand[..4] == [0, 0, 0, 0xFF]
            && operand[4..8] == [0, 0, 0, 0xFF]
            && operand[8..] == [0, 0];
        if is_auto {
            tap.style_defaults.shading = Some(TableStyleShading::NoShading);
            return Ok(());
        }
        let mut cell = CellProperties::default();
        Self::apply_full_shading(&mut cell, operand, false)?;
        let shading = cell.shading.ok_or_else(|| {
            PackageError::Corrupted(
                "DOC table-style shading did not contain a Shd value".to_string(),
            )
        })?;
        tap.style_defaults.shading = Some(TableStyleShading::Shading(shading));
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_cell_spacing(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 6 || operand[0] != 0 || operand[1] != 1 || operand[2] != 0x0F {
            return Err(PackageError::Corrupted(
                "DOC default cell spacing must target range 0..1 and all sides".to_string(),
            ));
        }
        let width = binary_to_doc_result(read_u16_le(operand, 4))?;
        if width > 15_840 {
            return Err(PackageError::Corrupted(
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
                return Err(PackageError::Corrupted(
                    "DOC default cell spacing has invalid units or value".to_string(),
                ));
            },
        };
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_cell_border_types(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if !operand.len().is_multiple_of(4) || operand.len() / 4 > tap.cell_properties.len() {
            return Err(PackageError::Corrupted(
                "DOC cell border-type array has an invalid size for the row".to_string(),
            ));
        }
        for (cell, types) in tap.cell_properties.iter_mut().zip(operand.chunks_exact(4)) {
            let top = Self::parse_border_type(types[0], true)?;
            let left = Self::parse_border_type(types[1], true)?;
            let bottom = Self::parse_border_type(types[2], true)?;
            let right = Self::parse_border_type(types[3], true)?;
            let overrides = CellBorderTypes {
                top: Some(top),
                left: Some(left),
                bottom: Some(bottom),
                right: Some(right),
            };
            cell.border_type_overrides = overrides;
            for (border, border_type) in [
                (&mut cell.borders.top, top),
                (&mut cell.borders.left, left),
                (&mut cell.borders.bottom, bottom),
                (&mut cell.borders.right, right),
            ] {
                if let Some(border) = border {
                    border.border_type = border_type;
                }
            }
        }
        Ok(())
    }
}
