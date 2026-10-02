//! Table and cell-definition codecs.

use super::prelude::*;

impl TapParser<'_> {
    pub(in crate::parts::tap_parser) fn parse_table_definition(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        _grpprl: &[u8],
        tolerance: &mut ToleranceReport,
    ) -> Result<()> {
        let data = sprm.operand_bytes();
        if data.is_empty() {
            return Err(PackageError::Corrupted(
                "sprmTDefTable does not contain a column count".to_string(),
            ));
        }

        // Read cell count
        let itc_mac = binary_to_doc_result(read_byte(data, 0))? as usize;
        if itc_mac > 63 {
            return Err(PackageError::Corrupted(
                "sprmTDefTable contains more than 63 columns".to_string(),
            ));
        }
        let start_of_tcs = 1 + ((itc_mac + 1) * 2);
        if data.len() < start_of_tcs || !(data.len() - start_of_tcs).is_multiple_of(20) {
            return Err(PackageError::Corrupted(
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
        // A boundary outside the XAS range cannot be laid out; Word clamps
        // it to the page coordinate range. Clamping preserves the ordering
        // checked below, so the repair never invents an inverted cell.
        for (index, boundary) in boundaries.iter_mut().enumerate() {
            if !(-31_680..=31_680).contains(boundary) {
                tolerance.record(
                    TableDefinitionDefect::BoundaryRange,
                    u16::try_from(index).unwrap_or(u16::MAX),
                );
                *boundary = (*boundary).clamp(-31_680, 31_680);
            }
        }
        if boundaries.windows(2).any(|pair| pair[0] > pair[1]) {
            return Err(PackageError::Corrupted(
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
                    cell_props.push(self.parse_table_cell_descriptor(
                        data,
                        tc_offset,
                        u16::try_from(i).unwrap_or(u16::MAX),
                        sprm.opcode,
                        tolerance,
                    )?);
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

    /// Parse a `TableCellDescriptor` (TC) structure.
    ///
    /// TC structure (20 bytes total):
    /// - bytes 0-1: flags (fVertical, fBackward, fRotateFont, fVertMerge, fVertRestart, etc.)
    /// - bytes 2-3: wWidth (preferred cell width)
    /// - bytes 4-7: brcTop (top border, 4 bytes)
    /// - bytes 8-11: brcLeft (left border, 4 bytes)
    /// - bytes 12-15: brcBottom (bottom border, 4 bytes)
    /// - bytes 16-19: brcRight (right border, 4 bytes)
    ///
    /// MS-DOC 2.9.313 defines the used values of each flag field but places
    /// no reader-side rejection on the unused encodings. Like Word and
    /// LibreOffice, an unused encoding reads as the field's neutral member
    /// and is recorded as a [`TableDefinitionDefect`] against `cell_index`.

    pub(in crate::parts::tap_parser) fn parse_table_cell_descriptor(
        &self,
        data: &[u8],
        offset: usize,
        cell_index: u16,
        opcode: u16,
        tolerance: &mut ToleranceReport,
    ) -> Result<CellProperties> {
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
            _ => {
                tolerance.record(TableDefinitionDefect::TextFlow, cell_index);
                TextDirection::LrTb
            },
        };

        // Bits 5-6: vertical merge state.
        props.vertical_merge_status = match (flags >> 5) & 0x03 {
            0 => VerticalMergeStatus::None,
            1 => VerticalMergeStatus::Merged,
            3 => VerticalMergeStatus::First,
            _ => {
                tolerance.record(TableDefinitionDefect::VerticalMerge, cell_index);
                VerticalMergeStatus::None
            },
        };

        // Bits 7-8: vertical alignment.
        let vert_align = (flags >> 7) & 0x03;
        props.vertical_alignment = match vert_align {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            _ => {
                tolerance.record(TableDefinitionDefect::VerticalAlignment, cell_index);
                VerticalAlignment::Top
            },
        };
        props.direct_style.vertical_alignment = true;

        props.fit_text = flags & 0x1000 != 0;
        props.no_wrap = flags & 0x2000 != 0;
        props.direct_style.no_wrap = true;
        props.hide_mark = flags & 0x4000 != 0;

        // Read preferred width (bytes 2-3)
        let w_width = binary_to_doc_result(read_i16_le(data, offset + 2))?;
        // A negative width cannot be laid out; the cell falls back to no
        // preferred width, as if ftsWidth were ftsNil.
        if w_width < 0 {
            tolerance.record(TableDefinitionDefect::NegativeWidth, cell_index);
        }
        let fts_width = (flags >> 9) & 0x07;
        props.preferred_width = match fts_width {
            0 => None,
            1..=3 if w_width < 0 => None,
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
            _ => {
                tolerance.record(TableDefinitionDefect::WidthUnit, cell_index);
                None
            },
        };

        // Read borders (4 bytes each)
        props.borders.top = Self::parse_border_code(data, offset + 4, opcode, tolerance)?;
        props.borders.left = Self::parse_border_code(data, offset + 8, opcode, tolerance)?;
        props.borders.bottom = Self::parse_border_code(data, offset + 12, opcode, tolerance)?;
        props.borders.right = Self::parse_border_code(data, offset + 16, opcode, tolerance)?;
        props.direct_style.border_top = data[offset + 4..offset + 8] != [0xFF; 4];
        props.direct_style.border_left = data[offset + 8..offset + 12] != [0xFF; 4];
        props.direct_style.border_bottom = data[offset + 12..offset + 16] != [0xFF; 4];
        props.direct_style.border_right = data[offset + 16..offset + 20] != [0xFF; 4];

        Ok(props)
    }
}
