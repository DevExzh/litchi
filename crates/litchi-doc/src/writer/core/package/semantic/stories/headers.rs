use crate::sprm_operations::*;
use crate::writer::core::{codec, model::*};
use crate::writer::font_table::FontTableBuilder;
use crate::writer::piece_table::Piece;
impl Writer {
    /// Build header/footer story text and PlcfHdd
    ///
    /// Appends header/footer text to `text_stream`, extends CHPX/PAPX entries and pieces.
    /// Returns (plcfhdd_bytes, header_cp_length). If no header/footer set, returns None.
    #[allow(clippy::too_many_arguments)] // TODO: Refactor to reduce arguments
    pub(in crate::writer::core::package) fn build_header_story(
        &self,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
        header_pic_offsets: &[u32],
    ) -> Result<Option<HeaderStoryData>, WriteError> {
        // Short-circuit if nothing set
        if self.header_even.is_none()
            && self.header_odd.is_none()
            && self.header_first.is_none()
            && self.footer_even.is_none()
            && self.footer_odd.is_none()
            && self.footer_first.is_none()
        {
            return Ok(None);
        }
        let story_text_start = text_stream.len();

        // Build index->paragraph mapping for 12 slots per MS-DOC PlcfHdd / Apache POI:
        //   Slots 0-5:  footnote/endnote separator/continuation stories
        //   Slot 6:     even page header (section 0)
        //   Slot 7:     odd page header (section 0) — "default" when no facing pages
        //   Slot 8:     even page footer (section 0)
        //   Slot 9:     odd page footer (section 0) — "default" when no facing pages
        //   Slot 10:    first page header (section 0)
        //   Slot 11:    first page footer (section 0)
        // PlcfHdd has 14 CPs (12 slot starts + story end + ignored final CP).
        // Verified against LibreOffice DOC writer output.
        let mut idx_paragraphs: [Option<&[HeaderFooterParagraph]>; 12] = [None; 12];
        if let Some(ref paragraphs) = self.header_even {
            idx_paragraphs[6] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.header_odd {
            idx_paragraphs[7] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.header_first {
            idx_paragraphs[10] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.footer_even {
            idx_paragraphs[8] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.footer_odd {
            idx_paragraphs[9] = Some(paragraphs);
        }
        if let Some(ref paragraphs) = self.footer_first {
            idx_paragraphs[11] = Some(paragraphs);
        }

        // Local CP within header story (counts only header subdocument)
        // Empty slots consume no CPs. Non-empty header/footer stories contain a content paragraph
        // mark and a separate guard paragraph mark.
        let mut header_cp: u32 = 0;
        let mut cp_starts: [u32; 12] = [0; 12];
        let mut field_char_cps = Vec::new();
        let mut shape_anchor_cps: Vec<(u32, FloatingAnchorKind)> = Vec::new();

        for i in 0..12 {
            cp_starts[i] = header_cp;
            if let Some(paragraphs) = idx_paragraphs[i] {
                let mut field_state = HeaderFieldState::default();
                let fc_story_start = checked_text_fc(text_fc_start, text_stream.len())?;
                let mut story_chars = 0u32;

                for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
                    // Paragraphs appended by insert_header_text_box /
                    // insert_header_picture hold 0x0008 anchors; record their
                    // story CPs and the anchored item kind.
                    let anchor_kind = self
                        .header_anchors
                        .iter()
                        .find(|anchor| {
                            anchor.slot == i && anchor.paragraph_index == paragraph_index
                        })
                        .map(|anchor| anchor.kind);
                    if let Some(kind) = anchor_kind {
                        shape_anchor_cps.push((header_cp + story_chars, kind));
                    }
                    let fc_para_start = checked_text_fc(text_fc_start, text_stream.len())?;
                    let mut paragraph_chars = 0u32;
                    let mut last_chpx = None;

                    for (text, formatting) in &paragraph.runs {
                        let run_chars = utf16_code_unit_len(text)?;
                        let mut marker_cp = header_cp
                            .checked_add(story_chars)
                            .and_then(|value| value.checked_add(paragraph_chars))
                            .ok_or_else(|| {
                                WriteError::InvalidData(
                                    "DOC header/footer field CP range overflows".to_string(),
                                )
                            })?;
                        for character in text.chars() {
                            if field_state.observe(character, formatting)? {
                                field_char_cps.push((marker_cp, character as u16));
                            }
                            marker_cp = marker_cp
                                .checked_add(character.len_utf16() as u32)
                                .ok_or_else(|| {
                                    WriteError::InvalidData(
                                        "DOC header/footer field CP range overflows".to_string(),
                                    )
                                })?;
                        }
                        if run_chars == 0 {
                            continue;
                        }
                        let run_fc_start = checked_text_fc(text_fc_start, text_stream.len())?;
                        for unit in text.encode_utf16() {
                            text_stream.extend_from_slice(&unit.to_le_bytes());
                        }
                        let run_fc_end = checked_text_fc(text_fc_start, text_stream.len())?;
                        // Header picture anchors also carry sprmCPicLocation
                        // pointing at the picture's Data-stream block.
                        let mut grpprl = codec::build_chpx_grpprl(formatting, font_builder);
                        if let Some(FloatingAnchorKind::Picture(pic_index)) = anchor_kind {
                            let pic_offset =
                                header_pic_offsets.get(pic_index as usize).ok_or_else(|| {
                                    WriteError::InvalidData(format!(
                                        "DOC header picture index {pic_index} is out of range"
                                    ))
                                })?;
                            grpprl.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
                            grpprl.extend_from_slice(&pic_offset.to_le_bytes());
                        }
                        chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                        last_chpx = Some(chpx_entries.len() - 1);
                        paragraph_chars =
                            paragraph_chars.checked_add(run_chars).ok_or_else(|| {
                                WriteError::InvalidData(
                                    "DOC header/footer paragraph CP range overflows".to_string(),
                                )
                            })?;
                    }

                    text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                    let fc_para_end = checked_text_fc(text_fc_start, text_stream.len())?;
                    if let Some(index) = last_chpx {
                        chpx_entries[index].1 = fc_para_end;
                    } else {
                        chpx_entries.push((fc_para_start, fc_para_end, Vec::new()));
                    }
                    papx_entries.push((
                        fc_para_start,
                        fc_para_end,
                        codec::build_papx_grpprl(&paragraph.formatting),
                    ));
                    story_chars = story_chars
                        .checked_add(paragraph_chars)
                        .and_then(|value| value.checked_add(1))
                        .ok_or_else(|| {
                            WriteError::InvalidData(
                                "DOC header/footer story CP range overflows".to_string(),
                            )
                        })?;
                }

                // Guard paragraph mark required between stories.
                let fc_guard_start = checked_text_fc(text_fc_start, text_stream.len())?;
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                let fc_guard_end = checked_text_fc(text_fc_start, text_stream.len())?;
                chpx_entries.push((fc_guard_start, fc_guard_end, Vec::new()));
                papx_entries.push((
                    fc_guard_start,
                    fc_guard_end,
                    codec::build_papx_grpprl(&ParagraphFormatting::default()),
                ));
                story_chars = story_chars.checked_add(1).ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC header/footer story CP range overflows".to_string(),
                    )
                })?;

                let cp_story_end = current_cp_total.checked_add(story_chars).ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC header/footer total CP range overflows".to_string(),
                    )
                })?;
                pieces.push(Piece::new(
                    *current_cp_total,
                    cp_story_end,
                    fc_story_start,
                    true,
                ));
                *current_cp_total = cp_story_end;
                header_cp = header_cp.checked_add(story_chars).ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC header/footer subdocument CP range overflows".to_string(),
                    )
                })?;
                field_state.finish()?;
            }
        }

        // The header subdocument ends with an extra paragraph mark. The second-to-last PlcfHdd
        // CP terminates the final story at ccpHdd - 1; the last CP is ignored.
        let stories_end = header_cp;
        let fc_trailing = text_fc_start + text_stream.len() as u32;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_trailing_end = fc_trailing + 2;
        chpx_entries.push((fc_trailing, fc_trailing_end, Vec::new()));
        papx_entries.push((
            fc_trailing,
            fc_trailing_end,
            codec::build_papx_grpprl(&ParagraphFormatting::default()),
        ));
        pieces.push(Piece::new(
            *current_cp_total,
            *current_cp_total + 1,
            fc_trailing,
            true,
        ));
        *current_cp_total += 1;
        header_cp += 1;

        let mut plcfhdd = Vec::with_capacity((12 + 2) * 4);
        for cp_start in &cp_starts {
            plcfhdd.extend_from_slice(&cp_start.to_le_bytes());
        }
        plcfhdd.extend_from_slice(&stories_end.to_le_bytes());
        plcfhdd.extend_from_slice(&header_cp.to_le_bytes());

        let fields = if field_char_cps.is_empty() {
            Vec::new()
        } else {
            crate::writer::fields::build_plcffld(
                &field_char_cps,
                header_cp,
                &text_stream[story_text_start..],
            )?
        };
        Ok(Some(HeaderStoryData {
            plcfhdd,
            fields,
            char_count: header_cp,
            shape_anchor_cps,
        }))
    }
}
