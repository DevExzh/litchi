use crate::writer::comments::CommentEntry;
use crate::writer::core::{codec, model::*};
use crate::writer::fib::FibBuilder;
use crate::writer::font_table::FontTableBuilder;
use crate::writer::piece_table::Piece;
use std::collections::HashMap;
impl Writer {
    /// Append the comment subdocument and build its owner, reference, and text tables.
    #[allow(clippy::too_many_arguments)]
    pub(in crate::writer::core::package) fn build_comment_story(
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
    ) -> Result<Option<CommentStoryData>, WriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() != actual_ref_cps.len() {
            return Err(WriteError::InvalidData(
                "every DOC comment must have a reference in the main document".to_string(),
            ));
        }

        let mut ordered = entries
            .iter()
            .zip(actual_ref_cps.iter().copied())
            .collect::<Vec<_>>();
        ordered.sort_by_key(|(_, cp)| *cp);
        if ordered.windows(2).any(|pair| pair[0].1 == pair[1].1) {
            return Err(WriteError::InvalidData(
                "DOC comment references must have unique character positions".to_string(),
            ));
        }
        if ordered.iter().any(|(_, cp)| *cp >= ccp_text) {
            return Err(WriteError::InvalidData(
                "DOC comment reference lies outside the main document".to_string(),
            ));
        }

        let mut owners = Vec::<String>::new();
        let mut owner_indexes = Vec::with_capacity(ordered.len());
        for (entry, _) in &ordered {
            let author_len = entry.author.encode_utf16().count();
            if author_len >= 56 {
                return Err(WriteError::InvalidData(
                    "DOC comment author names must contain fewer than 56 UTF-16 code units"
                        .to_string(),
                ));
            }
            let initials_len = entry.initials.encode_utf16().count();
            if initials_len > 9 {
                return Err(WriteError::InvalidData(
                    "DOC comment initials must contain at most nine UTF-16 code units".to_string(),
                ));
            }
            let index = if let Some(index) = owners.iter().position(|owner| owner == &entry.author)
            {
                index
            } else {
                if owners.len() >= 0x7FFF {
                    return Err(WriteError::InvalidData(
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
            return Err(WriteError::InvalidData(
                "DOC annotation bookmark table exceeds 0x3FFC entries".to_string(),
            ));
        }
        let bookmark_sentinel = ccp_text.checked_add(1).ok_or_else(|| {
            WriteError::InvalidData("DOC annotation bookmark sentinel overflows".to_string())
        })?;
        let mut bookmark_tags = vec![None; ordered.len()];
        let mut ranges = Vec::<(u32, u32, u32)>::with_capacity(ranged_count);
        for (index, (entry, _)) in ordered.iter().enumerate() {
            let Some((start, end)) = entry.range else {
                continue;
            };
            if start > end || end > ccp_text {
                return Err(WriteError::InvalidData(
                    "DOC comment range must be ordered and inside the main document".to_string(),
                ));
            }
            let tag = i32::try_from(index).map_err(|_| {
                WriteError::InvalidData("DOC comment bookmark tag overflows".to_string())
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
                WriteError::InvalidData("DOC comment reply depth is too large".to_string())
            })?;
            if depth > active_ancestors.len() {
                return Err(WriteError::InvalidData(
                    "DOC comment reply tree must be in pre-order".to_string(),
                ));
            }
            active_ancestors.truncate(depth);
            let parent_delta = match (depth, metadata.parent_index) {
                (0, None) => 0,
                (0, Some(_)) | (_, None) => {
                    return Err(WriteError::InvalidData(
                        "DOC comment parent and reply depth are inconsistent".to_string(),
                    ));
                },
                (_, Some(parent)) => {
                    let expected = active_ancestors.get(depth - 1).copied().ok_or_else(|| {
                        WriteError::InvalidData("DOC comment reply tree is malformed".to_string())
                    })?;
                    if parent != expected {
                        return Err(WriteError::InvalidData(
                            "DOC comment parent does not match pre-order reply depth".to_string(),
                        ));
                    }
                    i32::try_from(parent as i64 - index as i64).map_err(|_| {
                        WriteError::InvalidData(
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
                WriteError::InvalidData("DOC comment story CP overflows".to_string())
            })?;
            let story_end = current_cp_total.checked_add(story_chars).ok_or_else(|| {
                WriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
            })?;
            pieces.push(Piece::new(
                *current_cp_total,
                story_end,
                fc_story_start,
                true,
            ));
            *current_cp_total = story_end;
            comment_cp = comment_cp.checked_add(story_chars).ok_or_else(|| {
                WriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
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
            WriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
        })?;
        pieces.push(Piece::new(*current_cp_total, guard_end, fc_guard, true));
        *current_cp_total = guard_end;
        comment_cp = comment_cp.checked_add(1).ok_or_else(|| {
            WriteError::InvalidData("DOC comment subdocument CP overflows".to_string())
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

    pub(in crate::writer::core::package) fn append_comment_tables(
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
}
