//! Semantic DOC stories and table assembly.

use crate::sprm_operations::*;
use crate::writer::bookmarks::BookmarkEntry;
use crate::writer::comments::CommentEntry;
use crate::writer::fib::FibBuilder;
use crate::writer::font_table::FontTableBuilder;
use crate::writer::footnotes::FootnoteEntry;
use crate::writer::piece_table::Piece;
use crate::writer::smart_tags::SmartTagTableData;
use std::collections::HashMap;

use super::super::{codec, model::*};

impl DocWriter {
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
    pub(super) fn build_note_story(
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
    ) -> Result<Option<NoteStoryData>, DocWriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() != actual_ref_cps.len() {
            return Err(DocWriteError::InvalidData(
                "every DOC note must have a reference in the main document".to_string(),
            ));
        }

        let mut ordered = entries
            .iter()
            .zip(actual_ref_cps.iter().copied())
            .collect::<Vec<_>>();
        ordered.sort_by_key(|(_, cp)| *cp);
        if ordered.windows(2).any(|pair| pair[0].1 == pair[1].1) {
            return Err(DocWriteError::InvalidData(
                "DOC note references must have unique character positions".to_string(),
            ));
        }
        if ordered.iter().any(|(_, cp)| *cp >= ccp_text) {
            return Err(DocWriteError::InvalidData(
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

    /// Append the comment subdocument and build its owner, reference, and text tables.
    #[allow(clippy::too_many_arguments)]
    pub(super) fn build_comment_story(
        entries: &[CommentEntry],
        actual_ref_cps: &[u32],
        ccp_text: u32,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
    ) -> Result<Option<CommentStoryData>, DocWriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() != actual_ref_cps.len() {
            return Err(DocWriteError::InvalidData(
                "every DOC comment must have a reference in the main document".to_string(),
            ));
        }

        let mut ordered = entries
            .iter()
            .zip(actual_ref_cps.iter().copied())
            .collect::<Vec<_>>();
        ordered.sort_by_key(|(_, cp)| *cp);
        if ordered.windows(2).any(|pair| pair[0].1 == pair[1].1) {
            return Err(DocWriteError::InvalidData(
                "DOC comment references must have unique character positions".to_string(),
            ));
        }
        if ordered.iter().any(|(_, cp)| *cp >= ccp_text) {
            return Err(DocWriteError::InvalidData(
                "DOC comment reference lies outside the main document".to_string(),
            ));
        }

        let mut owners = Vec::<String>::new();
        let mut owner_indexes = Vec::with_capacity(ordered.len());
        for (entry, _) in &ordered {
            let author_len = entry.author.encode_utf16().count();
            if author_len >= 56 {
                return Err(DocWriteError::InvalidData(
                    "DOC comment author names must contain fewer than 56 UTF-16 code units"
                        .to_string(),
                ));
            }
            let initials_len = entry.initials.encode_utf16().count();
            if initials_len > 9 {
                return Err(DocWriteError::InvalidData(
                    "DOC comment initials must contain at most nine UTF-16 code units".to_string(),
                ));
            }
            let index = if let Some(index) = owners.iter().position(|owner| owner == &entry.author)
            {
                index
            } else {
                if owners.len() >= 0x7FFF {
                    return Err(DocWriteError::InvalidData(
                        "DOC comment owner array exceeds 0x7FFF entries".to_string(),
                    ));
                }
                owners.push(entry.author.clone());
                owners.len() - 1
            };
            owner_indexes.push(index as u16);
        }

        let mut owner_bytes = Vec::new();
        for owner in &owners {
            let units = owner.encode_utf16().collect::<Vec<_>>();
            owner_bytes.extend_from_slice(&(units.len() as u16).to_le_bytes());
            owner_bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }

        let ranged_count = ordered
            .iter()
            .filter(|(entry, _)| entry.range.is_some())
            .count();
        if ranged_count > 0x3FFC {
            return Err(DocWriteError::InvalidData(
                "DOC annotation bookmark table exceeds 0x3FFC entries".to_string(),
            ));
        }
        let bookmark_sentinel = ccp_text.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC annotation bookmark sentinel overflows".to_string())
        })?;
        let mut bookmark_tags = vec![None; ordered.len()];
        let mut ranges = Vec::<(u32, u32, u32)>::with_capacity(ranged_count);
        for (index, (entry, _)) in ordered.iter().enumerate() {
            let Some((start, end)) = entry.range else {
                continue;
            };
            if start > end || end > ccp_text {
                return Err(DocWriteError::InvalidData(
                    "DOC comment range must be ordered and inside the main document".to_string(),
                ));
            }
            let tag = i32::try_from(index).map_err(|_| {
                DocWriteError::InvalidData("DOC comment bookmark tag overflows".to_string())
            })? as u32;
            bookmark_tags[index] = Some(tag);
            ranges.push((tag, start, end));
        }

        let mut bookmark_names = Vec::new();
        let mut bookmark_starts = Vec::new();
        let mut bookmark_ends = Vec::new();
        if !ranges.is_empty() {
            let mut start_order = ranges.clone();
            start_order.sort_by_key(|&(tag, start, _)| (start, tag));
            let mut end_order = ranges.clone();
            end_order.sort_by_key(|&(tag, _, end)| (end, tag));
            let end_indexes = end_order
                .iter()
                .enumerate()
                .map(|(index, &(tag, _, _))| (tag, index as u16))
                .collect::<HashMap<_, _>>();

            bookmark_names.extend_from_slice(&0xFFFFu16.to_le_bytes());
            bookmark_names.extend_from_slice(&(ranges.len() as u16).to_le_bytes());
            bookmark_names.extend_from_slice(&10u16.to_le_bytes());
            for &(tag, _, _) in &start_order {
                bookmark_names.extend_from_slice(&0u16.to_le_bytes());
                bookmark_names.extend_from_slice(&0x0100u16.to_le_bytes());
                bookmark_names.extend_from_slice(&tag.to_le_bytes());
                bookmark_names.extend_from_slice(&(-1i32).to_le_bytes());
            }

            for &(_, start, _) in &start_order {
                bookmark_starts.extend_from_slice(&start.to_le_bytes());
            }
            bookmark_starts.extend_from_slice(&bookmark_sentinel.to_le_bytes());
            for &(tag, _, _) in &start_order {
                bookmark_starts.extend_from_slice(&end_indexes[&tag].to_le_bytes());
                bookmark_starts.extend_from_slice(&0u16.to_le_bytes());
            }

            for &(_, _, end) in &end_order {
                bookmark_ends.extend_from_slice(&end.to_le_bytes());
            }
            bookmark_ends.extend_from_slice(&bookmark_sentinel.to_le_bytes());
        }

        let mut extended_metadata = Vec::with_capacity(ordered.len() * 18);
        let mut active_ancestors = Vec::<usize>::new();
        for (index, (entry, _)) in ordered.iter().enumerate() {
            let metadata = entry
                .extended_metadata
                .unwrap_or(crate::CommentExtendedMetadata {
                    modified_at: None,
                    depth: 0,
                    parent_index: None,
                    is_ink: false,
                });
            let depth = usize::try_from(metadata.depth).map_err(|_| {
                DocWriteError::InvalidData("DOC comment reply depth is too large".to_string())
            })?;
            if depth > active_ancestors.len() {
                return Err(DocWriteError::InvalidData(
                    "DOC comment reply tree must be in pre-order".to_string(),
                ));
            }
            active_ancestors.truncate(depth);
            let parent_delta = match (depth, metadata.parent_index) {
                (0, None) => 0,
                (0, Some(_)) | (_, None) => {
                    return Err(DocWriteError::InvalidData(
                        "DOC comment parent and reply depth are inconsistent".to_string(),
                    ));
                },
                (_, Some(parent)) => {
                    let expected = active_ancestors.get(depth - 1).copied().ok_or_else(|| {
                        DocWriteError::InvalidData(
                            "DOC comment reply tree is malformed".to_string(),
                        )
                    })?;
                    if parent != expected {
                        return Err(DocWriteError::InvalidData(
                            "DOC comment parent does not match pre-order reply depth".to_string(),
                        ));
                    }
                    i32::try_from(parent as i64 - index as i64).map_err(|_| {
                        DocWriteError::InvalidData(
                            "DOC comment parent offset exceeds the binary format".to_string(),
                        )
                    })?
                },
            };
            extended_metadata.extend_from_slice(&pack_dttm(metadata.modified_at)?.to_le_bytes());
            extended_metadata.extend_from_slice(&0u16.to_le_bytes());
            extended_metadata.extend_from_slice(&metadata.depth.to_le_bytes());
            extended_metadata.extend_from_slice(&parent_delta.to_le_bytes());
            extended_metadata.extend_from_slice(&(u32::from(metadata.is_ink) << 1).to_le_bytes());
            active_ancestors.push(index);
        }

        let mut comment_cp = 0u32;
        let mut text_cps = vec![0u32];
        for (entry, _) in &ordered {
            let fc_story_start = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x0005u16.to_le_bytes());
            let fc_marker_end = fc_story_start + 2;
            let marker_grpprl = codec::build_chpx_grpprl(
                &CharacterFormatting {
                    special: Some(true),
                    ..Default::default()
                },
                font_builder,
            );
            chpx_entries.push((fc_story_start, fc_marker_end, marker_grpprl));

            let body_chars = utf16_code_unit_len(&entry.text)?;
            let fc_body_start = text_fc_start + text_stream.len() as u32;
            text_stream.extend(entry.text.encode_utf16().flat_map(u16::to_le_bytes));
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_story_end = text_fc_start + text_stream.len() as u32;
            chpx_entries.push((
                fc_body_start,
                fc_story_end,
                codec::build_chpx_grpprl(&CharacterFormatting::default(), font_builder),
            ));
            papx_entries.push((
                fc_story_start,
                fc_story_end,
                codec::build_papx_grpprl(&ParagraphFormatting::default()),
            ));

            let story_chars = body_chars.checked_add(2).ok_or_else(|| {
                DocWriteError::InvalidData("DOC comment story CP overflows".to_string())
            })?;
            let story_end = current_cp_total.checked_add(story_chars).ok_or_else(|| {
                DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
            })?;
            pieces.push(Piece::new(
                *current_cp_total,
                story_end,
                fc_story_start,
                true,
            ));
            *current_cp_total = story_end;
            comment_cp = comment_cp.checked_add(story_chars).ok_or_else(|| {
                DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
            })?;
            text_cps.push(comment_cp);
        }

        let fc_guard = text_fc_start + text_stream.len() as u32;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_guard_end = fc_guard + 2;
        chpx_entries.push((fc_guard, fc_guard_end, Vec::new()));
        papx_entries.push((
            fc_guard,
            fc_guard_end,
            codec::build_papx_grpprl(&ParagraphFormatting::default()),
        ));
        let guard_end = current_cp_total.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
        })?;
        pieces.push(Piece::new(*current_cp_total, guard_end, fc_guard, true));
        *current_cp_total = guard_end;
        comment_cp = comment_cp.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
        })?;
        text_cps.push(comment_cp);

        let mut references = Vec::with_capacity((ordered.len() + 1) * 4 + ordered.len() * 30);
        for (_, cp) in &ordered {
            references.extend_from_slice(&cp.to_le_bytes());
        }
        references.extend_from_slice(&ccp_text.to_le_bytes());
        for (index, ((entry, _), author_index)) in ordered.iter().zip(owner_indexes).enumerate() {
            let initials = entry.initials.encode_utf16().collect::<Vec<_>>();
            references.extend_from_slice(&(initials.len() as u16).to_le_bytes());
            for index in 0..9 {
                references
                    .extend_from_slice(&initials.get(index).copied().unwrap_or(0).to_le_bytes());
            }
            references.extend_from_slice(&author_index.to_le_bytes());
            references.extend_from_slice(&0u16.to_le_bytes());
            references.extend_from_slice(&0u16.to_le_bytes());
            let tag = bookmark_tags[index].map_or(-1, |tag| tag as i32);
            references.extend_from_slice(&tag.to_le_bytes());
        }

        let mut text_positions = Vec::with_capacity(text_cps.len() * 4);
        for cp in text_cps {
            text_positions.extend_from_slice(&cp.to_le_bytes());
        }

        Ok(Some(CommentStoryData {
            owners: owner_bytes,
            references,
            text_positions,
            bookmark_names,
            bookmark_starts,
            bookmark_ends,
            extended_metadata,
            char_count: comment_cp,
        }))
    }

    pub(super) fn append_comment_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        comment: &CommentStoryData,
    ) {
        let mut offset = table_stream.len() as u32;
        fib.set_grp_xst_atn_owners(offset, comment.owners.len() as u32);
        table_stream.extend_from_slice(&comment.owners);

        offset = table_stream.len() as u32;
        fib.set_plcfand_ref(offset, comment.references.len() as u32);
        table_stream.extend_from_slice(&comment.references);

        offset = table_stream.len() as u32;
        fib.set_plcfand_txt(offset, comment.text_positions.len() as u32);
        table_stream.extend_from_slice(&comment.text_positions);

        if !comment.bookmark_names.is_empty() {
            offset = table_stream.len() as u32;
            fib.set_sttbf_atn_bkmk(offset, comment.bookmark_names.len() as u32);
            table_stream.extend_from_slice(&comment.bookmark_names);

            offset = table_stream.len() as u32;
            fib.set_plcf_atn_bkf(offset, comment.bookmark_starts.len() as u32);
            table_stream.extend_from_slice(&comment.bookmark_starts);

            offset = table_stream.len() as u32;
            fib.set_plcf_atn_bkl(offset, comment.bookmark_ends.len() as u32);
            table_stream.extend_from_slice(&comment.bookmark_ends);
        }

        offset = table_stream.len() as u32;
        fib.set_atrd_extra(offset, comment.extended_metadata.len() as u32);
        table_stream.extend_from_slice(&comment.extended_metadata);
    }

    pub(super) fn build_bookmark_tables(
        entries: &[BookmarkEntry],
        document_end: u32,
    ) -> Result<Option<BookmarkTableData>, DocWriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() > 0x3FFB {
            return Err(DocWriteError::InvalidData(
                "DOC standard bookmark table exceeds 0x3FFB entries".to_string(),
            ));
        }
        let mut unique = std::collections::HashSet::with_capacity(entries.len());
        let mut records = Vec::with_capacity(entries.len());
        for (index, entry) in entries.iter().enumerate() {
            let units = entry.name.encode_utf16().collect::<Vec<_>>();
            if units.is_empty() || units.len() >= 40 || !unique.insert(entry.name.clone()) {
                return Err(DocWriteError::InvalidData(
                    "DOC bookmark names must be unique and contain 1 through 39 UTF-16 code units"
                        .to_string(),
                ));
            }
            if entry.start > entry.end || entry.end > document_end {
                return Err(DocWriteError::InvalidData(
                    "DOC bookmark range must be ordered and inside the document parts".to_string(),
                ));
            }
            let mut bkc = u16::from(entry.is_native) << 14;
            if let Some((first, limit)) = entry.column_range {
                if first >= limit || first > 0x7F || limit > 0x3F {
                    return Err(DocWriteError::InvalidData(
                        "DOC bookmark column range exceeds BKC limits".to_string(),
                    ));
                }
                bkc |= 0x8000 | u16::from(first) | (u16::from(limit) << 8);
            }
            records.push((index, entry, units, bkc));
        }

        let sentinel = document_end.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC bookmark sentinel CP overflows".to_string())
        })?;
        let mut start_order = records.iter().collect::<Vec<_>>();
        start_order.sort_by_key(|record| (record.1.start, record.0));
        let mut end_order = records.iter().collect::<Vec<_>>();
        end_order.sort_by_key(|record| (record.1.end, record.0));
        let end_indexes = end_order
            .iter()
            .enumerate()
            .map(|(end_index, record)| (record.0, end_index as u16))
            .collect::<HashMap<_, _>>();

        let mut names = Vec::new();
        names.extend_from_slice(&0xFFFFu16.to_le_bytes());
        names.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        names.extend_from_slice(&0u16.to_le_bytes());
        for record in &start_order {
            names.extend_from_slice(&(record.2.len() as u16).to_le_bytes());
            names.extend(record.2.iter().copied().flat_map(u16::to_le_bytes));
        }

        let mut starts = Vec::with_capacity((entries.len() + 1) * 4 + entries.len() * 4);
        for record in &start_order {
            starts.extend_from_slice(&record.1.start.to_le_bytes());
        }
        starts.extend_from_slice(&sentinel.to_le_bytes());
        for record in &start_order {
            starts.extend_from_slice(&end_indexes[&record.0].to_le_bytes());
            starts.extend_from_slice(&record.3.to_le_bytes());
        }

        let mut ends = Vec::with_capacity((entries.len() + 1) * 4);
        for record in &end_order {
            ends.extend_from_slice(&record.1.end.to_le_bytes());
        }
        ends.extend_from_slice(&sentinel.to_le_bytes());
        Ok(Some(BookmarkTableData {
            names,
            starts,
            ends,
        }))
    }

    pub(super) fn build_revision_writer_data(
        &self,
    ) -> Result<Option<RevisionWriterData>, DocWriteError> {
        let mut authors = vec!["Unknown".to_string()];
        let mut indexes = HashMap::from([("Unknown".to_string(), 0u16)]);
        let mut has_revisions = false;
        let mut index_author = |author: &str| -> Result<(), DocWriteError> {
            has_revisions = true;
            if !indexes.contains_key(author) {
                if authors.len() >= 0x8000 {
                    return Err(DocWriteError::InvalidData(
                        "DOC revision author table exceeds the signed author-index range"
                            .to_string(),
                    ));
                }
                let index = authors.len() as u16;
                authors.push(author.to_string());
                indexes.insert(author.to_string(), index);
            }
            Ok(())
        };
        if let Some(revision) = &self.section_formatting_revision {
            index_author(&revision.author)?;
        }
        for style in &self.styles {
            if let Some(revision) = &style.revision {
                index_author(&revision.author)?;
            }
        }
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            let mut formatting = Some(&paragraph.formatting);
            while let Some(current) = formatting {
                if let Some(revision) = &current.formatting_revision {
                    index_author(&revision.author)?;
                }
                if let Some(revision) = &current.numbering_revision {
                    index_author(&revision.author)?;
                }
                formatting = current.preserved_properties_for_revision.as_deref();
            }
            for run in &paragraph.runs {
                let mut formatting = Some(&run.formatting);
                while let Some(current) = formatting {
                    if let Some(revision) = &current.insertion_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.deletion_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.formatting_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.display_field_revision {
                        index_author(&revision.author)?;
                    }
                    formatting = current.preserved_properties_for_revision.as_deref();
                }
            }
        }
        if !has_revisions {
            return Ok(None);
        }

        let mut table = Vec::new();
        table.extend_from_slice(&0xFFFFu16.to_le_bytes());
        table.extend_from_slice(&(authors.len() as u16).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        for author in authors {
            let units = author.encode_utf16().collect::<Vec<_>>();
            let length = u16::try_from(units.len()).map_err(|_| {
                DocWriteError::InvalidData(
                    "DOC revision author exceeds the STTB string-length limit".to_string(),
                )
            })?;
            table.extend_from_slice(&length.to_le_bytes());
            table.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        Ok(Some(RevisionWriterData { indexes, table }))
    }

    pub(super) fn validate_style_reference(
        &self,
        index: u16,
        expected_kind: crate::StyleKind,
        context: &str,
    ) -> Result<(), DocWriteError> {
        let actual_kind = match index {
            0 => Some(crate::StyleKind::Paragraph),
            10 => Some(crate::StyleKind::Character),
            15..=0x0FFC => self
                .styles
                .get(usize::from(index - 15))
                .map(|style| style.kind),
            _ => None,
        };
        let Some(actual_kind) = actual_kind else {
            return Err(DocWriteError::InvalidData(format!(
                "{context} references undefined DOC style index {index}"
            )));
        };
        if actual_kind != expected_kind {
            return Err(DocWriteError::InvalidData(format!(
                "{context} references {actual_kind:?} DOC style {index}, expected {expected_kind:?}"
            )));
        }
        Ok(())
    }

    pub(super) fn validate_character_style_references(
        &self,
        formatting: &CharacterFormatting,
        context: &str,
    ) -> Result<(), DocWriteError> {
        if let Some(index) = formatting.style_index {
            self.validate_style_reference(index, crate::StyleKind::Character, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_character_style_references(previous, context)?;
        }
        Ok(())
    }

    pub(super) fn validate_paragraph_style_references(
        &self,
        formatting: &ParagraphFormatting,
        context: &str,
    ) -> Result<(), DocWriteError> {
        if let Some(index) = formatting.style_index {
            self.validate_style_reference(index, crate::StyleKind::Paragraph, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_paragraph_style_references(previous, context)?;
        }
        Ok(())
    }

    pub(super) fn validate_table_style_references(
        &self,
        formatting: &crate::writer::tap::TableRow,
        context: &str,
    ) -> Result<(), DocWriteError> {
        if let Some(index) = formatting.table_style_index {
            self.validate_style_reference(index, crate::StyleKind::Table, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_table_style_references(previous, context)?;
        }
        Ok(())
    }

    pub(super) fn validate_style_references(&self) -> Result<(), DocWriteError> {
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            self.validate_paragraph_style_references(
                &paragraph.formatting,
                "DOC paragraph formatting",
            )?;
            for run in &paragraph.runs {
                self.validate_character_style_references(
                    &run.formatting,
                    "DOC character formatting",
                )?;
            }
        }
        for table in &self.tables {
            for row in &table.rows {
                self.validate_table_style_references(&row.formatting, "DOC table row formatting")?;
            }
        }
        Ok(())
    }

    pub(super) fn append_revision_author_table(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        revisions: &RevisionWriterData,
    ) {
        let offset = table_stream.len() as u32;
        fib.set_sttbf_rmark(offset, revisions.table.len() as u32);
        table_stream.extend_from_slice(&revisions.table);
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn append_tables_to_main_story(
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
    ) -> Result<(), DocWriteError> {
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
                    return Err(DocWriteError::InvalidData(
                        "DOC table rows must contain between 1 and 63 cells".to_string(),
                    ));
                }
                if row.formatting.cells.len() != column_count {
                    return Err(DocWriteError::InvalidData(
                        "DOC table row formatting must define every cell".to_string(),
                    ));
                }
                if row.formatting.is_header && encountered_body_row {
                    return Err(DocWriteError::InvalidData(
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
                                return Err(DocWriteError::InvalidData(format!(
                                    "DOC cell {index} continues a vertical merge that was not started"
                                )));
                            }
                        },
                    }
                }
                for cell in &row.cells {
                    if cell.paragraphs.is_empty() {
                        return Err(DocWriteError::InvalidData(
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
                        DocWriteError::InvalidData(
                            "DOC text stream exceeds 32-bit FC space".to_string(),
                        )
                    })?)
                    .ok_or_else(|| {
                        DocWriteError::InvalidData("DOC table row FC overflows".to_string())
                    })?;
                text_stream.extend_from_slice(&0x0007u16.to_le_bytes());
                let fc_end = fc_start.checked_add(2).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table row FC overflows".to_string())
                })?;
                chpx_entries.push((fc_start, fc_end, Vec::new()));
                papx_entries.push((
                    fc_start,
                    fc_end,
                    codec::build_table_row_papx_grpprl(&row.formatting)?,
                ));
                let cp_end = current_cp.checked_add(1).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table CP range overflows".to_string())
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

    #[allow(clippy::too_many_arguments)]
    pub(super) fn append_table_paragraph(
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
    ) -> Result<(), DocWriteError> {
        let fc_start = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph FC overflows".to_string())
            })?;
        let mut paragraph_cps = 0u32;
        let mut last_chpx = None;
        for run in &paragraph.runs {
            let run_fc_start = text_fc_start
                .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                    DocWriteError::InvalidData(
                        "DOC text stream exceeds 32-bit FC space".to_string(),
                    )
                })?)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run FC overflows".to_string())
                })?;
            let run_cps = utf16_code_unit_len(&run.text)?;
            let mut offset = 0u32;
            for ch in run.text.chars() {
                let cp = current_cp
                    .checked_add(paragraph_cps)
                    .and_then(|value| value.checked_add(offset))
                    .ok_or_else(|| {
                        DocWriteError::InvalidData(
                            "DOC table field character CP overflows".to_string(),
                        )
                    })?;
                if matches!(ch as u32, 0x0013..=0x0015) {
                    field_char_cps.push((cp, ch as u16));
                }
                offset = offset.checked_add(ch.len_utf16() as u32).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run CP range overflows".to_string())
                })?;
            }
            for unit in run.text.encode_utf16() {
                text_stream.extend_from_slice(&unit.to_le_bytes());
            }
            let run_fc_end = run_fc_start
                .checked_add(run_cps.checked_mul(2).ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run FC overflows".to_string())
                })?)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("DOC table run FC overflows".to_string())
                })?;
            chpx_entries.push((
                run_fc_start,
                run_fc_end,
                codec::build_revision_chpx_grpprl(&run.formatting, font_builder, revision_data)?,
            ));
            last_chpx = Some(chpx_entries.len() - 1);
            paragraph_cps = paragraph_cps.checked_add(run_cps).ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph CP range overflows".to_string())
            })?;
        }
        text_stream.extend_from_slice(&terminator.to_le_bytes());
        let fc_end = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC table paragraph FC overflows".to_string())
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
                DocWriteError::InvalidData("DOC table paragraph CP range overflows".to_string())
            })?;
        pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
        *current_cp = cp_end;
        Ok(())
    }

    pub(super) fn append_empty_main_paragraph(
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        current_cp: &mut u32,
        pieces: &mut Vec<Piece>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
    ) -> Result<(), DocWriteError> {
        let fc_start = text_fc_start
            .checked_add(u32::try_from(text_stream.len()).map_err(|_| {
                DocWriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
            })?)
            .ok_or_else(|| {
                DocWriteError::InvalidData("DOC final paragraph FC overflows".to_string())
            })?;
        text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
        let fc_end = fc_start.checked_add(2).ok_or_else(|| {
            DocWriteError::InvalidData("DOC final paragraph FC overflows".to_string())
        })?;
        chpx_entries.push((fc_start, fc_end, Vec::new()));
        papx_entries.push((fc_start, fc_end, Vec::new()));
        let cp_end = current_cp.checked_add(1).ok_or_else(|| {
            DocWriteError::InvalidData("DOC final paragraph CP overflows".to_string())
        })?;
        pieces.push(Piece::new(*current_cp, cp_end, fc_start, true));
        *current_cp = cp_end;
        Ok(())
    }

    pub(super) fn append_bookmark_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        bookmarks: &BookmarkTableData,
    ) {
        let mut offset = table_stream.len() as u32;
        fib.set_sttbf_bkmk(offset, bookmarks.names.len() as u32);
        table_stream.extend_from_slice(&bookmarks.names);
        offset = table_stream.len() as u32;
        fib.set_plcf_bkf(offset, bookmarks.starts.len() as u32);
        table_stream.extend_from_slice(&bookmarks.starts);
        offset = table_stream.len() as u32;
        fib.set_plcf_bkl(offset, bookmarks.ends.len() as u32);
        table_stream.extend_from_slice(&bookmarks.ends);
    }

    pub(super) fn append_smart_tag_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        smart_tags: &SmartTagTableData,
    ) {
        if let Some(data) = &smart_tags.infos {
            let offset = table_stream.len() as u32;
            fib.set_sttbf_bkmk_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.starts {
            let offset = table_stream.len() as u32;
            fib.set_plcf_bkf_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.ends {
            let offset = table_stream.len() as u32;
            fib.set_plcf_bkl_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.factoid_data {
            let offset = table_stream.len() as u32;
            fib.set_factoid_data(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.recognizer_ranges {
            let offset = table_stream.len() as u32;
            fib.set_plcf_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
    }

    /// Build header/footer story text and PlcfHdd
    ///
    /// Appends header/footer text to `text_stream`, extends CHPX/PAPX entries and pieces.
    /// Returns (plcfhdd_bytes, header_cp_length). If no header/footer set, returns None.
    #[allow(clippy::too_many_arguments)] // TODO: Refactor to reduce arguments
    pub(super) fn build_header_story(
        &self,
        text_fc_start: u32,
        text_stream: &mut Vec<u8>,
        chpx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        papx_entries: &mut Vec<(u32, u32, Vec<u8>)>,
        pieces: &mut Vec<Piece>,
        current_cp_total: &mut u32,
        font_builder: &mut FontTableBuilder,
        header_pic_offsets: &[u32],
    ) -> Result<Option<HeaderStoryData>, DocWriteError> {
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
                                DocWriteError::InvalidData(
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
                                    DocWriteError::InvalidData(
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
                                    DocWriteError::InvalidData(format!(
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
                                DocWriteError::InvalidData(
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
                            DocWriteError::InvalidData(
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
                    DocWriteError::InvalidData(
                        "DOC header/footer story CP range overflows".to_string(),
                    )
                })?;

                let cp_story_end = current_cp_total.checked_add(story_chars).ok_or_else(|| {
                    DocWriteError::InvalidData(
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
                    DocWriteError::InvalidData(
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
