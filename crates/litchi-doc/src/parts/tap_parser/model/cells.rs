//! Semantic cell-range and merge state transitions.

use super::prelude::*;

impl TapParser<'_> {
    pub(in crate::parts::tap_parser) fn parse_cell_text_flow(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 4 {
            return Err(PackageError::Corrupted(
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
                return Err(PackageError::Corrupted(
                    "DOC cell text-flow value is invalid".to_string(),
                ));
            },
        };
        for cell in &mut tap.cell_properties[range] {
            cell.text_direction = direction;
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_vertical_merge(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 || operand[0] as usize >= tap.cell_properties.len() {
            return Err(PackageError::Corrupted(
                "DOC vertical-merge operand is invalid for the row".to_string(),
            ));
        }
        tap.cell_properties[operand[0] as usize].vertical_merge_status = match operand[1] {
            0 => VerticalMergeStatus::None,
            1 => VerticalMergeStatus::Merged,
            3 => VerticalMergeStatus::First,
            _ => {
                return Err(PackageError::Corrupted(
                    "DOC vertical-merge flag is invalid".to_string(),
                ));
            },
        };
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_vertical_alignment(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 3 {
            return Err(PackageError::Corrupted(
                "DOC vertical-alignment operand must contain 3 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let alignment = match operand[2] {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            _ => {
                return Err(PackageError::Corrupted(
                    "DOC cell vertical-alignment value is invalid".to_string(),
                ));
            },
        };
        for cell in &mut tap.cell_properties[range] {
            cell.vertical_alignment = alignment;
            cell.direct_style.vertical_alignment = true;
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn parse_cell_range_bool(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        property: CellBoolProperty,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 3 || !matches!(operand[2], 0 | 1) {
            return Err(PackageError::Corrupted(
                "DOC Boolean cell-range operand is invalid".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let value = operand[2] != 0;
        for cell in &mut tap.cell_properties[range] {
            match property {
                CellBoolProperty::FitText => cell.fit_text = value,
                CellBoolProperty::NoWrap => {
                    cell.no_wrap = value;
                    cell.direct_style.no_wrap = true;
                },
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

    pub(in crate::parts::tap_parser) fn handle_insert_cells(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 4 {
            return Err(PackageError::Corrupted(
                "DOC cell-insertion operand must contain 4 bytes".to_string(),
            ));
        }
        let index = operand[0] as usize;
        let count = operand[1] as usize;
        let width = binary_to_doc_result(read_u16_le(operand, 2))?;
        if index > tap.cell_count || count == 0 || tap.cell_count + count > 63 || width > 31_680 {
            return Err(PackageError::Corrupted(
                "DOC cell insertion has an invalid index, count, or width".to_string(),
            ));
        }
        let added_width = i32::from(width) * count as i32;
        let first_boundary = i32::from(*tap.cell_boundaries.first().unwrap_or(&0));
        let last_boundary = i32::from(*tap.cell_boundaries.last().unwrap_or(&0));
        if last_boundary - first_boundary + added_width > 31_680
            || last_boundary + added_width > 31_680
        {
            return Err(PackageError::Corrupted(
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
                    PackageError::Corrupted("DOC cell insertion overflows coordinates".to_string())
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

    pub(in crate::parts::tap_parser) fn handle_delete_cells(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(
                "DOC cell-deletion operand must contain 2 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        if range.len() >= tap.cell_count {
            return Err(PackageError::Corrupted(
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

    pub(in crate::parts::tap_parser) fn handle_column_width(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 4 {
            return Err(PackageError::Corrupted(
                "DOC column-width operand must contain 4 bytes".to_string(),
            ));
        }
        let range = Self::cell_range(tap, operand)?;
        let width = binary_to_doc_result(read_u16_le(operand, 2))?;
        if width > 31_680 {
            return Err(PackageError::Corrupted(
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
            return Err(PackageError::Corrupted(
                "DOC column widths produce a coordinate outside the XAS range".to_string(),
            ));
        }
        tap.cell_boundaries = boundaries
            .into_iter()
            .map(i16::try_from)
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(|_| {
                PackageError::Corrupted("DOC column widths overflow coordinates".to_string())
            })?;
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn handle_horizontal_merge(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
        merge: bool,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(
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
}
