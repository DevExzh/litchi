use crate::writer::core::{codec, model::*};
use crate::writer::font_table::FontTableBuilder;
use crate::writer::footnotes::FootnoteEntry;
use crate::writer::piece_table::Piece;
impl Writer {
    /// Build footnote or endnote subdocument text and PLCFs.
    ///
    /// Per MS-DOC spec:
    /// - Each note text MUST begin with U+0002 (auto-numbered reference mark) with fSpec=1
    /// - PlcffndRef final CP MUST equal `ccp_text` (main document character count)
    /// - PlcffndTxt CPs are relative to the note subdocument start
    ///
    /// `actual_ref_cps`: actual CPs in main doc where U+0002 refs were injected (entry order).
    /// `ccp_text`: FibRgLw97.ccpText — needed for the mandatory final CP in PlcffndRef.
    #[allow(clippy::too_many_arguments)]
    pub(in crate::writer::core::package) fn build_note_story(
        entries: &[FootnoteEntry],
        actual_ref_cps: &[u32],
        ccp_text: u32,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
    ) -> Result<Option<NoteStoryData>, WriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() != actual_ref_cps.len() {
            return Err(WriteError::InvalidData(
                "every DOC note must have a reference in the main document".to_string(),
            ));
        }

        let mut ordered = entries
            .iter()
            .zip(actual_ref_cps.iter().copied())
            .collect::<Vec<_>>();
        ordered.sort_by_key(|(_, cp)| *cp);
        if ordered.windows(2).any(|pair| pair[0].1 == pair[1].1) {
            return Err(WriteError::InvalidData(
                "DOC note references must have unique character positions".to_string(),
            ));
        }
        if ordered.iter().any(|(_, cp)| *cp >= ccp_text) {
            return Err(WriteError::InvalidData(
                "DOC note reference lies outside the main document".to_string(),
            ));
        }

        let mut note_cp: u32 = 0;
        // PlcffndTxt: n story starts, one story terminator, and one ignored final CP.
        let mut txt_cps: Vec<u32> = vec![0];

        for (entry, _) in &ordered {
            let fc_para_start = text_fc_start + text_stream.len() as u32;

            // 1) Auto-numbered reference mark U+0002 with fSpec=1 CHPX
            //    This is what Word displays as the footnote number in the note area.
            let fc_ref = fc_para_start;
            text_stream.extend_from_slice(&0x0002u16.to_le_bytes());
            let fc_ref_end = fc_ref + 2;
            let ref_grpprl = codec::build_chpx_grpprl(
                &CharacterFormatting {
                    special: Some(true),
                    ..Default::default()
                },
                font_builder,
            );
            chpx_entries.push((fc_ref, fc_ref_end, ref_grpprl));

            // 2) Note body text
            let text = &entry.text;
            let text_chars = utf16_code_unit_len(text)?;
            let fc_text_start = text_fc_start + text_stream.len() as u32;
            for u in text.encode_utf16() {
                text_stream.extend_from_slice(&u.to_le_bytes());
            }
            let fc_text_end = fc_text_start + text_chars * 2;
            let body_grpprl =
                codec::build_chpx_grpprl(&CharacterFormatting::default(), font_builder);
            chpx_entries.push((fc_text_start, fc_text_end, body_grpprl));

            // 3) Paragraph mark (chEop 0x0D) — extends last CHPX
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            if let Some(last) = chpx_entries.last_mut() {
                last.1 += 2;
            }
            let fc_para_end = text_fc_start + text_stream.len() as u32;

            // PAPX for this note paragraph
            papx_entries.push((
                fc_para_start,
                fc_para_end,
                codec::build_papx_grpprl(&ParagraphFormatting::default()),
            ));

            // Piece: 1 (auto-ref) + text_chars + 1 (para mark)
            let total_chars = 1 + text_chars + 1;
            pieces.push(Piece::new(
                *current_cp_total,
                *current_cp_total + total_chars,
                fc_para_start,
                true,
            ));
            *current_cp_total += total_chars;
            note_cp += total_chars;

            txt_cps.push(note_cp);
        }

        // Trailing guard paragraph mark — mandatory per MS-DOC spec:
        // "The entire footnote subdocument MUST end with a paragraph mark."
        // This is an EXTRA paragraph mark beyond the last footnote's own \r.
        // LibreOffice and POI both write this guard.
        {
            let fc_guard = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_guard_end = fc_guard + 2;
            chpx_entries.push((fc_guard, fc_guard_end, Vec::new()));
            papx_entries.push((
                fc_guard,
                fc_guard_end,
                codec::build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(
                *current_cp_total,
                *current_cp_total + 1,
                fc_guard,
                true,
            ));
            *current_cp_total += 1;
            note_cp += 1;
            txt_cps.push(note_cp);
        }

        // PlcffndRef: actual reference CPs + mandatory final CP = ccpText
        let mut ref_cps = ordered.iter().map(|(_, cp)| *cp).collect::<Vec<_>>();
        ref_cps.push(ccp_text);

        // Serialize PlcffndRef: (n+1) CPs then n FRDs (2 bytes each)
        let mut plcf_ref = Vec::with_capacity(ref_cps.len() * 4 + entries.len() * 2);
        for cp in &ref_cps {
            plcf_ref.extend_from_slice(&cp.to_le_bytes());
        }
        // FRD nAuto is nonzero for an automatically numbered note.
        for (entry, _) in &ordered {
            plcf_ref.extend_from_slice(&entry.number.max(1).to_le_bytes());
        }

        // Serialize PlcffndTxt: (n+2) CPs for n footnotes (n stories + 1 guard + 1 final)
        let mut plcf_txt = Vec::with_capacity(txt_cps.len() * 4);
        for cp in &txt_cps {
            plcf_txt.extend_from_slice(&cp.to_le_bytes());
        }

        Ok(Some((plcf_ref, plcf_txt, note_cp)))
    }
}
