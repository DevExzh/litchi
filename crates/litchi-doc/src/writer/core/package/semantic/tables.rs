use crate::writer::core::{
    codec,
    model::{RevisionWriterData, WritableParagraph, WriteError, Writer, utf16_code_unit_len},
};
use crate::writer::font_table::FontTableBuilder;
use crate::writer::piece_table::Piece;
impl Writer {
    pub(in crate::writer::core::package) fn append_tables_to_main_story(
        &self,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        field_char_cps: &mut Vec<(u32, u16)>,
        font_builder: &mut FontTableBuilder,
        revision_data: Option<&RevisionWriterData>,
    ) -> Result<(), WriteError> {
        for table in &self.tables {
            let mut encountered_body_row = false;
            let mut vertical_merges = table
                .rows
                .first()
                .map(|row| vec![false; row.cells.len()])
                .unwrap_or_default();
            for row in &table.rows {
                let column_count = row.cells.len();
                if !(1..=63).contains(&column_count) {
                    return Err(WriteError::InvalidData(
                        "DOC table rows must contain between 1 and 63 cells".to_string(),
                    ));
                }
                if row.formatting.cells.len() != column_count {
                    return Err(WriteError::InvalidData(
                        "DOC table row formatting must define every cell".to_string(),
                    ));
                }
                if row.formatting.is_header && encountered_body_row {
                    return Err(WriteError::InvalidData(
                        "DOC header rows must form a contiguous prefix of the table".to_string(),
                    ));
                }
                encountered_body_row |= !row.formatting.is_header;
                for (index, cell) in row.formatting.cells.iter().enumerate() {
                    match cell.vertical_merge {
                        crate::parts::tap::VerticalMergeStatus::None => {
                            vertical_merges[index] = false;
                        },
                        crate::parts::tap::VerticalMergeStatus::First => {
                            vertical_merges[index] = true;
                        },
                        crate::parts::tap::VerticalMergeStatus::Merged => {
                            if !vertical_merges[index] {
                                return Err(WriteError::InvalidData(format!(
                                    "DOC cell {index} continues a vertical merge that was not started"
                                )));
                            }
                        },
                    }
                }
                for cell in &row.cells {
                    if cell.paragraphs.is_empty() {
                        return Err(WriteError::InvalidData(
                            "DOC table cells must contain at least one paragraph".to_string(),
                        ));
                    }
                    let last_paragraph = cell.paragraphs.len() - 1;
                    for (index, paragraph) in cell.paragraphs.iter().enumerate() {
                        let terminator = if index == last_paragraph {
                            0x0007
                        } else {
                            0x000D
                        };
                        Self::append_table_paragraph(
                            paragraph,
                            terminator,
                            text_fc_start,
                            text_stream,
                            current_cp,
                            pieces,
                            chpx_entries,
                            papx_entries,
                            field_char_cps,
                            font_builder,
                            revision_data,
                        )?;
                    }
                }

                let fc_start = text_fc_start
                    .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                        WriteError::InvalidData(
                            "DOC text stream exceeds 32-bit FC space".to_string(),
                        )
                    })?)
                    .ok_or_else(|| {
                        WriteError::InvalidData("DOC table row FC overflows".to_string())
                    })?;
                text_stream.extend_from_slice(&0x0007u16.to_le_bytes());
                let fc_end = fc_start.checked_add(2).ok_or_else(|| {
                    WriteError::InvalidData("DOC table row FC overflows".to_string())
                })?;
                chpx_entries.push((fc_start, fc_end, Vec::new()));
                papx_entries.push((
                    fc_start,
                    fc_end,
                    codec::build_table_row_papx_grpprl(&row.formatting)?,
                ));
                let cp_end = current_cp.checked_add(1).ok_or_else(|| {
                    WriteError::InvalidData("DOC table CP range overflows".to_string())
                })?;
                pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
                *current_cp = cp_end;
            }

            // The main document must end in U+000D. A non-table paragraph also
            // separates adjacent writer table objects into distinct tables.
            Self::append_empty_main_paragraph(
                text_fc_start,
                text_stream,
                current_cp,
                pieces,
                chpx_entries,
                papx_entries,
            )?;
        }
        Ok(())
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "parameters map one-to-one to a fixed DOC record or semantic construction"
    )]
    pub(in crate::writer::core::package) fn append_table_paragraph(
        paragraph: &WritableParagraph,
        terminator: u16,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        field_char_cps: &mut Vec<(u32, u16)>,
        font_builder: &mut FontTableBuilder,
        revision_data: Option<&RevisionWriterData>,
    ) -> Result<(), WriteError> {
        let fc_start = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                WriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                WriteError::InvalidData("DOC table paragraph FC overflows".to_string())
            })?;
        let mut paragraph_cps = 0u32;
        let mut last_chpx = None;
        for run in &paragraph.runs {
            let run_fc_start = text_fc_start
                .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                    WriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
                })?)
                .ok_or_else(|| WriteError::InvalidData("DOC table run FC overflows".to_string()))?;
            let run_cps = utf16_code_unit_len(&run.text)?;
            let mut offset = 0u32;
            for ch in run.text.chars() {
                let cp = current_cp
                    .checked_add(paragraph_cps)
                    .and_then(|value| value.checked_add(offset))
                    .ok_or_else(|| {
                        WriteError::InvalidData(
                            "DOC table field character CP overflows".to_string(),
                        )
                    })?;
                if matches!(ch as u32, 0x0013..=0x0015) {
                    field_char_cps.push((cp, ch as u16));
                }
                offset = offset.checked_add(ch.len_utf16() as u32).ok_or_else(|| {
                    WriteError::InvalidData("DOC table run CP range overflows".to_string())
                })?;
            }
            for unit in run.text.encode_utf16() {
                text_stream.extend_from_slice(&unit.to_le_bytes());
            }
            let run_fc_end = run_fc_start
                .checked_add(run_cps.checked_mul(2).ok_or_else(|| {
                    WriteError::InvalidData("DOC table run FC overflows".to_string())
                })?)
                .ok_or_else(|| WriteError::InvalidData("DOC table run FC overflows".to_string()))?;
            chpx_entries.push((
                run_fc_start,
                run_fc_end,
                codec::build_revision_chpx_grpprl(&run.formatting, font_builder, revision_data)?,
            ));
            last_chpx = Some(chpx_entries.len() - 1);
            paragraph_cps = paragraph_cps.checked_add(run_cps).ok_or_else(|| {
                WriteError::InvalidData("DOC table paragraph CP range overflows".to_string())
            })?;
        }
        text_stream.extend_from_slice(&terminator.to_le_bytes());
        let fc_end = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                WriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                WriteError::InvalidData("DOC table paragraph FC overflows".to_string())
            })?;
        if let Some(index) = last_chpx {
            chpx_entries[index].1 = fc_end;
        } else {
            chpx_entries.push((fc_start, fc_end, Vec::new()));
        }
        let mut papx = codec::build_revision_papx_grpprl(&paragraph.formatting, revision_data)?;
        codec::append_table_depth_sprms(&mut papx);
        papx_entries.push((fc_start, fc_end, papx));
        let cp_end = current_cp
            .checked_add(paragraph_cps)
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| {
                WriteError::InvalidData("DOC table paragraph CP range overflows".to_string())
            })?;
        pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
        *current_cp = cp_end;
        Ok(())
    }

    pub(in crate::writer::core::package) fn append_empty_main_paragraph(
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
    ) -> Result<(), WriteError> {
        let fc_start = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                WriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                WriteError::InvalidData("DOC final paragraph FC overflows".to_string())
            })?;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_end = fc_start.checked_add(2).ok_or_else(|| {
            WriteError::InvalidData("DOC final paragraph FC overflows".to_string())
        })?;
        chpx_entries.push((fc_start, fc_end, Vec::new()));
        papx_entries.push((fc_start, fc_end, Vec::new()));
        let cp_end = current_cp.checked_add(1).ok_or_else(|| {
            WriteError::InvalidData("DOC final paragraph CP overflows".to_string())
        })?;
        pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
        *current_cp = cp_end;
        Ok(())
    }
}
