//! Legacy and full-color table shading codecs.

use super::prelude::*;

impl<'arena> TapParser<'arena> {
    pub(in crate::parts::tap_parser) fn parse_cell_shading(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
    ) -> Result<()> {
        let bytes = sprm.operand_bytes();
        if !bytes.len().is_multiple_of(2) || bytes.len() / 2 > tap.cell_properties.len() {
            return Err(PackageError::Corrupted(
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
                tap.cell_properties[i].direct_style.shading = true;
                continue;
            }
            let ico_fore = (shd & 0x1F) as u8;
            let ico_back = ((shd >> 5) & 0x1F) as u8;
            let ipat = ((shd >> 10) & 0x3F) as u8;
            if ico_fore > 16 || ico_back > 16 {
                return Err(PackageError::Corrupted(
                    "DOC Shd80 contains an invalid palette index".to_string(),
                ));
            }
            let pattern = ShadingPattern::from_u8(ipat).ok_or_else(|| {
                PackageError::Corrupted("DOC Shd80 contains an invalid pattern index".to_string())
            })?;
            let shading = CellShading {
                foreground_color: Self::ico_to_rgb(ico_fore),
                background_color: Self::ico_to_rgb(ico_back),
                pattern,
            };
            tap.cell_properties[i].background_color = shading.background_color;
            tap.cell_properties[i].shading = Some(shading);
            tap.cell_properties[i].shading_inherits_from_style = false;
            tap.cell_properties[i].direct_style.shading = true;
        }

        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_full_cell_shading(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        first_cell: usize,
        raw: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if !operand.len().is_multiple_of(10) || operand.len() > 220 {
            return Err(PackageError::Corrupted(
                "DOC table Shd array has an invalid byte count".to_string(),
            ));
        }
        let count = operand.len() / 10;
        let chunk_limit = if first_cell == 44 { 19 } else { 22 };
        if count > chunk_limit || first_cell.saturating_add(count) > tap.cell_properties.len() {
            return Err(PackageError::Corrupted(
                "DOC table Shd array exceeds its cell chunk".to_string(),
            ));
        }
        for (offset, bytes) in operand.chunks_exact(10).enumerate() {
            Self::apply_full_shading(&mut tap.cell_properties[first_cell + offset], bytes, raw)?;
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_full_cell_shading_range(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        odd_only: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 12 {
            return Err(PackageError::Corrupted(
                "DOC table range shading operand must contain exactly 12 bytes".to_string(),
            ));
        }
        let first = operand[0] as usize;
        let limit = operand[1] as usize;
        if first >= tap.cell_properties.len() || limit < first || limit > tap.cell_properties.len()
        {
            return Err(PackageError::Corrupted(
                "DOC table range shading exceeds the row".to_string(),
            ));
        }
        let step = if odd_only { 2 } else { 1 };
        for index in (first..limit).step_by(step) {
            Self::apply_full_shading(&mut tap.cell_properties[index], &operand[2..], false)?;
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_full_table_shading(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 10 {
            return Err(PackageError::Corrupted(
                "DOC whole-table shading operand must contain exactly 10 bytes".to_string(),
            ));
        }
        for cell in &mut tap.cell_properties {
            Self::apply_full_shading(cell, operand, false)?;
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn apply_full_shading(
        cell: &mut CellProperties,
        bytes: &[u8],
        raw: bool,
    ) -> Result<()> {
        if bytes.len() != 10 {
            return Err(PackageError::Corrupted(
                "DOC Shd must contain exactly 10 bytes".to_string(),
            ));
        }
        let is_nil = bytes[..8].iter().all(|byte| *byte == 0xFF) && bytes[8..] == [0, 0];
        if is_nil {
            cell.shading = None;
            cell.background_color = None;
            cell.shading_inherits_from_style = raw;
            cell.direct_style.shading = !raw;
            return Ok(());
        }
        let is_auto =
            bytes[..4] == [0, 0, 0, 0xFF] && bytes[4..8] == [0, 0, 0, 0xFF] && bytes[8..] == [0, 0];
        if is_auto {
            cell.shading = None;
            cell.background_color = None;
            cell.shading_inherits_from_style = false;
            cell.direct_style.shading = true;
            return Ok(());
        }
        let foreground_color = Self::parse_colorref(&bytes[..4])?;
        let background_color = Self::parse_colorref(&bytes[4..8])?;
        let pattern_value = binary_to_doc_result(read_u16_le(bytes, 8))?;
        let pattern = u8::try_from(pattern_value)
            .ok()
            .and_then(ShadingPattern::from_u8)
            .ok_or_else(|| {
                PackageError::Corrupted("DOC Shd contains an invalid pattern index".to_string())
            })?;
        cell.shading = Some(CellShading {
            foreground_color,
            background_color,
            pattern,
        });
        cell.background_color = background_color;
        cell.shading_inherits_from_style = false;
        cell.direct_style.shading = true;
        Ok(())
    }
}
