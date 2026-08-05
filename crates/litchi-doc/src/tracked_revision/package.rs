//! Package-layer transactional editor for DOC tracked revisions.

use super::Limits;
use super::codec::*;
use super::model::*;
use crate::package::{Error as PackageError, Result};
use crate::sprm_operations::*;
use crate::writer::ChpxFkpBuilder;
use litchi_ole_common::object::{Editor as ObjectEditor, Targets};

#[derive(Clone)]
pub struct DocTrackedRevisionEditor {
    package: ObjectEditor,
    word_path: Vec<String>,
    table_path: Vec<String>,
    word: Vec<u8>,
    table: Vec<u8>,
    pieces: Vec<RawPiece>,
    chpx: Vec<FcRun>,
    papx: Vec<PapxRun>,
    authors: Vec<String>,
    cp_tables: Vec<CpTable>,
    main_ccp: u32,
    changed: bool,
}

impl DocTrackedRevisionEditor {
    pub fn open(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let package =
            ObjectEditor::open(bytes, Targets::default(), limits).map_err(PackageError::from)?;
        let word_path = vec!["WordDocument".to_string()];
        let word = package
            .stream(&word_path)
            .ok_or_else(|| corrupted("WordDocument stream is missing"))?
            .to_vec();
        if word.len() < FIB_FC_LCB + (STTBFRMARK + 1) * 8 || u16_at(&word, 0)? != 0xA5EC {
            return Err(corrupted(
                "tracked-revision editing requires Word 97+ FIB data",
            ));
        }
        let flags = u16_at(&word, 10)?;
        if flags & 0x0100 != 0 || u32_at(&word, 14)? != 0 {
            return Err(corrupted("encrypted DOC cannot be edited"));
        }
        let table_path = vec![
            if flags & 0x0200 != 0 {
                "1Table"
            } else {
                "0Table"
            }
            .to_string(),
        ];
        let table = package
            .stream(&table_path)
            .ok_or_else(|| corrupted("selected Table stream is missing"))?
            .to_vec();
        reject_protection(&word, &table)?;
        let main_ccp = u32_at(&word, FIB_CCP_TEXT)?;
        let pieces = parse_clx(&word, &table)?;
        if pieces.last().is_none_or(|piece| piece.end < main_ccp) {
            return Err(corrupted("piece table does not cover the main story"));
        }
        let chpx = parse_chpx(&word, &table)?;
        let papx = parse_papx(&word, &table)?;
        let authors = parse_authors(&word, &table)?;
        let cp_tables = [
            (PLCFANDREF, 30),
            (PLCFFLD_MOM, 2),
            (PLCFBKF, 4),
            (PLCFBKL, 0),
            (PLCFATNBKF, 4),
            (PLCFATNBKL, 0),
        ]
        .into_iter()
        .filter_map(|(index, size)| parse_cp_table(&word, &table, index, size).transpose())
        .collect::<Result<Vec<_>>>()?;
        let editor = Self {
            package,
            word_path,
            table_path,
            word,
            table,
            pieces,
            chpx,
            papx,
            authors,
            cp_tables,
            main_ccp,
            changed: false,
        };
        if editor.revisions()?.len() > MAX_REVISIONS {
            return Err(corrupted("revision count exceeds resource limit"));
        }
        Ok(editor)
    }

    pub fn is_changed(&self) -> bool {
        self.changed
    }

    pub fn authors(&self) -> &[String] {
        &self.authors
    }

    /// Lists character and PAPX property revisions, merging adjacent runs with
    /// identical metadata even when the range crosses piece boundaries.
    pub fn revisions(&self) -> Result<Vec<DocTrackedRevision>> {
        let mut output = Vec::new();
        for run in &self.chpx {
            let sprms = strict_sprms(&run.grpprl)?;
            for (kind, flag, author_op, time_op, reason_op, rsid_op) in [
                (
                    DocTrackedRevisionKind::Insertion,
                    SPRM_C_F_RMARK,
                    SPRM_C_IBST_RMARK,
                    SPRM_C_DTTM_RMARK,
                    SPRM_C_IDSL_RMARK,
                    SPRM_C_RSID_TEXT,
                ),
                (
                    DocTrackedRevisionKind::Deletion,
                    SPRM_C_F_RMARK_DEL,
                    SPRM_C_IBST_RMARK_DEL,
                    SPRM_C_DTTM_RMARK_DEL,
                    SPRM_C_IDSL_RMARK_DEL,
                    SPRM_C_RSID_RM_DEL,
                ),
            ] {
                if sprms
                    .iter()
                    .any(|s| s.opcode == flag && s.operand_byte() == Some(1))
                {
                    let metadata = metadata_from_sprms(
                        &sprms,
                        author_op,
                        time_op,
                        reason_op,
                        rsid_op,
                        &self.authors,
                    )?;
                    self.push_fc_revision(&mut output, run.start, run.end, kind, metadata)?;
                }
            }
            for opcode in [SPRM_C_PROP_RMARK90, SPRM_C_PROP_RMARK_CURRENT] {
                if let Some(mark) = sprms.iter().rev().find(|s| s.opcode == opcode) {
                    if mark.operand_bytes().first() == Some(&1) {
                        let metadata = property_metadata(
                            mark.operand_bytes(),
                            &sprms,
                            SPRM_C_RSID_PROP,
                            &self.authors,
                        )?;
                        self.push_fc_revision(
                            &mut output,
                            run.start,
                            run.end,
                            DocTrackedRevisionKind::CharacterFormatting,
                            metadata,
                        )?;
                    }
                    break;
                }
            }
        }
        for run in &self.papx {
            let body = run
                .grpprl
                .get(2..)
                .ok_or_else(|| corrupted("PAPX has no style index"))?;
            let sprms = strict_sprms(body)?;
            for (kind, op, rsid) in [
                (
                    DocTrackedRevisionKind::ParagraphFormatting,
                    [
                        SPRM_P_PROP_RMARK,
                        SPRM_P_PROP_RMARK90,
                        SPRM_P_PROP_RMARK_CURRENT,
                    ]
                    .as_slice(),
                    None,
                ),
                (
                    DocTrackedRevisionKind::TableRowFormatting,
                    [SPRM_T_PROP_RMARK].as_slice(),
                    Some(SPRM_T_RSID),
                ),
            ] {
                if let Some(mark) = sprms.iter().rev().find(|s| op.contains(&s.opcode))
                    && mark.operand_bytes().first() == Some(&1)
                {
                    let metadata = property_metadata(
                        mark.operand_bytes(),
                        &sprms,
                        rsid.unwrap_or(0),
                        &self.authors,
                    )?;
                    self.push_fc_revision(&mut output, run.start, run.end, kind, metadata)?;
                }
            }
        }
        output.sort_by_key(|r| (r.start_cp, r.end_cp, kind_order(r.kind)));
        merge_adjacent(&mut output);
        infer_moves(&mut output);
        Ok(output)
    }

    /// Inserts inert plain text and marks it as an insertion or move destination.
    /// Field delimiters, object markers, paragraph marks, and macro characters
    /// are rejected rather than interpreted.
    pub fn add_text(
        &mut self,
        cp: u32,
        text: &str,
        kind: DocTrackedRevisionKind,
        metadata: DocTrackedRevisionMetadata,
    ) -> Result<DocTrackedRevision> {
        if !matches!(
            kind,
            DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::MoveTo
        ) {
            return Err(corrupted(
                "add_text requires an insertion or move-to revision",
            ));
        }
        let units = text.encode_utf16().collect::<Vec<_>>();
        if units.is_empty()
            || units.len() > MAX_TEXT_UNITS
            || units.iter().any(|u| matches!(*u, 0..=8 | 11..=31 | 0xFFFC))
        {
            return Err(corrupted(
                "tracked text is empty, oversized, or contains an active control character",
            ));
        }
        if cp > self.main_ccp {
            return Err(corrupted("tracked insertion CP exceeds main story"));
        }
        validate_metadata(kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        for unit in &units {
            candidate.word.extend_from_slice(&unit.to_le_bytes());
        }
        let length =
            u32::try_from(units.len()).map_err(|_| corrupted("tracked text length exceeds u32"))?;
        insert_piece(
            &mut candidate.pieces,
            cp,
            length,
            u32::try_from(fc).map_err(|_| corrupted("FC exceeds u32"))?,
        )?;
        candidate.shift_cp_tables(cp, 0, length)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_add(length)
            .ok_or_else(|| corrupted("main story CP overflow"))?;
        let grpprl = encode_revision(kind, author, &metadata)?;
        let fc_end =
            u32::try_from(candidate.word.len()).map_err(|_| corrupted("FC exceeds u32"))?;
        candidate.chpx.push(FcRun {
            start: fc as u32,
            end: fc_end,
            grpprl,
        });
        candidate.chpx.sort_by_key(|run| run.start);
        candidate.rewrite_chpx()?;
        candidate.enable_tracking()?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        candidate.commit()?;
        *self = candidate;
        self.find_exact(cp, cp + length, kind)
    }

    /// Adds a revision mark to an existing main-story range.
    pub fn add(
        &mut self,
        start_cp: u32,
        end_cp: u32,
        kind: DocTrackedRevisionKind,
        metadata: DocTrackedRevisionMetadata,
    ) -> Result<DocTrackedRevision> {
        validate_range(start_cp, end_cp, self.main_ccp)?;
        validate_metadata(kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        candidate.mutate_mark(start_cp, end_cp, kind, Some((author, &metadata)))?;
        candidate.enable_tracking()?;
        candidate.commit()?;
        *self = candidate;
        self.find_exact(start_cp, end_cp, kind)
    }

    /// Replaces revision metadata without touching unrelated formatting SPRMs.
    pub fn update(
        &mut self,
        index: usize,
        metadata: DocTrackedRevisionMetadata,
    ) -> Result<DocTrackedRevision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        validate_metadata(revision.kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        candidate.mutate_mark(
            revision.start_cp,
            revision.end_cp,
            revision.kind,
            Some((author, &metadata)),
        )?;
        candidate.commit()?;
        *self = candidate;
        self.find_exact(revision.start_cp, revision.end_cp, revision.kind)
    }

    /// Removes a mark while retaining its text/current formatting.
    pub fn remove(&mut self, index: usize) -> Result<DocTrackedRevision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        let mut candidate = self.clone();
        candidate.mutate_mark(revision.start_cp, revision.end_cp, revision.kind, None)?;
        candidate.commit()?;
        *self = candidate;
        Ok(revision)
    }

    /// Accepts a revision using Word redline semantics.
    pub fn accept(&mut self, index: usize) -> Result<DocTrackedRevision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(
            revision.kind,
            DocTrackedRevisionKind::Deletion | DocTrackedRevisionKind::MoveFrom
        ) {
            self.delete_revision_text(&revision)?;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    /// Rejects a revision using Word redline semantics.
    pub fn reject(&mut self, index: usize) -> Result<DocTrackedRevision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(
            revision.kind,
            DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::MoveTo
        ) {
            self.delete_revision_text(&revision)?;
        } else if matches!(
            revision.kind,
            DocTrackedRevisionKind::CharacterFormatting
                | DocTrackedRevisionKind::ParagraphFormatting
                | DocTrackedRevisionKind::TableRowFormatting
        ) {
            let mut candidate = self.clone();
            candidate.reject_formatting_revision(&revision)?;
            candidate.commit()?;
            *self = candidate;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    fn delete_revision_text(&mut self, revision: &DocTrackedRevision) -> Result<()> {
        self.reject_destructive_interactions(revision.start_cp, revision.end_cp)?;
        let mut candidate = self.clone();
        delete_piece_range(&mut candidate.pieces, revision.start_cp, revision.end_cp)?;
        let removed = revision.end_cp - revision.start_cp;
        candidate.shift_cp_tables(revision.start_cp, removed, 0)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(removed)
            .ok_or_else(|| corrupted("main story CP underflow"))?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    fn reject_destructive_interactions(&self, start: u32, end: u32) -> Result<()> {
        let text = read_units(&self.word, &self.pieces, start, end)?;
        if text.iter().any(|u| matches!(*u, 0x13..=0x15)) {
            return Err(corrupted("accept/reject would delete a field boundary"));
        }
        for table in &self.cp_tables {
            for cp in &table.cps {
                if *cp > start && *cp < end {
                    return Err(corrupted(
                        "accept/reject would split a field, bookmark, or comment range",
                    ));
                }
            }
        }
        Ok(())
    }

    fn mutate_mark(
        &mut self,
        start: u32,
        end: u32,
        kind: DocTrackedRevisionKind,
        replacement: Option<(u16, &DocTrackedRevisionMetadata)>,
    ) -> Result<()> {
        let intervals = self.fc_intervals(start, end)?;
        match kind {
            DocTrackedRevisionKind::Insertion
            | DocTrackedRevisionKind::Deletion
            | DocTrackedRevisionKind::MoveFrom
            | DocTrackedRevisionKind::MoveTo
            | DocTrackedRevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    replace_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_chpx()?;
            },
            DocTrackedRevisionKind::ParagraphFormatting
            | DocTrackedRevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    replace_papx_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_papx()?;
            },
        }
        if let Some((_, metadata)) = replacement
            && !self.authors.iter().any(|a| a == &metadata.author)
        {
            return Err(corrupted("revision author indexing failed"));
        }
        Ok(())
    }

    fn reject_formatting_revision(&mut self, revision: &DocTrackedRevision) -> Result<()> {
        let intervals = self.fc_intervals(revision.start_cp, revision.end_cp)?;
        match revision.kind {
            DocTrackedRevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    restore_before_wall(
                        grp,
                        SPRM_C_WALL,
                        &revision_opcodes(DocTrackedRevisionKind::CharacterFormatting, true),
                    )
                })?;
                self.rewrite_chpx()
            },
            DocTrackedRevisionKind::ParagraphFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp
                        .get(..2)
                        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(
                        &grp[2..],
                        SPRM_P_WALL,
                        &revision_opcodes(DocTrackedRevisionKind::ParagraphFormatting, true),
                    )?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            },
            DocTrackedRevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp
                        .get(..2)
                        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(
                        &grp[2..],
                        SPRM_T_WALL,
                        &revision_opcodes(DocTrackedRevisionKind::TableRowFormatting, true),
                    )?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            },
            _ => Err(corrupted("revision is not a formatting revision")),
        }
    }

    fn fc_intervals(&self, start: u32, end: u32) -> Result<Vec<(u32, u32)>> {
        let mut output = Vec::new();
        for piece in &self.pieces {
            let left = start.max(piece.start);
            let right = end.min(piece.end);
            if left >= right {
                continue;
            }
            let scale = if piece.unicode { 2 } else { 1 };
            let fc_start = piece
                .fc
                .checked_add((left - piece.start) * scale)
                .ok_or_else(|| corrupted("FC overflow"))?;
            let fc_end = piece
                .fc
                .checked_add((right - piece.start) * scale)
                .ok_or_else(|| corrupted("FC overflow"))?;
            output.push((fc_start, fc_end));
        }
        if output.is_empty() {
            return Err(corrupted("revision range has no text pieces"));
        }
        Ok(output)
    }

    fn push_fc_revision(
        &self,
        output: &mut Vec<DocTrackedRevision>,
        fc_start: u32,
        fc_end: u32,
        kind: DocTrackedRevisionKind,
        metadata: ParsedMetadata,
    ) -> Result<()> {
        for piece in &self.pieces {
            let width = if piece.unicode { 2 } else { 1 };
            let piece_fc_end = piece
                .fc
                .checked_add((piece.end - piece.start) * width)
                .ok_or_else(|| corrupted("piece FC overflow"))?;
            let left = fc_start.max(piece.fc);
            let right = fc_end.min(piece_fc_end);
            if left >= right {
                continue;
            }
            if (left - piece.fc) % width != 0 || (right - piece.fc) % width != 0 {
                return Err(corrupted("CHPX boundary splits a text character"));
            }
            let start_cp = piece.start + (left - piece.fc) / width;
            let end_cp = piece.start + (right - piece.fc) / width;
            if start_cp < self.main_ccp {
                output.push(metadata.to_revision(kind, start_cp, end_cp.min(self.main_ccp)));
            }
        }
        Ok(())
    }

    fn find_exact(
        &self,
        start: u32,
        end: u32,
        kind: DocTrackedRevisionKind,
    ) -> Result<DocTrackedRevision> {
        self.revisions()?
            .into_iter()
            .find(|r| {
                r.start_cp == start
                    && r.end_cp == end
                    && (r.kind == kind
                        || matches!(
                            (r.kind, kind),
                            (
                                DocTrackedRevisionKind::Insertion,
                                DocTrackedRevisionKind::MoveTo
                            ) | (
                                DocTrackedRevisionKind::Deletion,
                                DocTrackedRevisionKind::MoveFrom
                            )
                        ))
            })
            .ok_or_else(|| corrupted("authored revision was not discoverable"))
    }

    fn author_index(&mut self, author: &str) -> Result<u16> {
        if author.is_empty() || author.encode_utf16().count() > u16::MAX as usize {
            return Err(corrupted("revision author is empty or too long"));
        }
        if self.authors.is_empty() {
            self.authors.push("Unknown".to_string());
        }
        if let Some(index) = self.authors.iter().position(|value| value == author) {
            return u16::try_from(index)
                .map_err(|_| corrupted("revision author index exceeds u16"));
        }
        if self.authors.len() >= MAX_AUTHORS {
            return Err(corrupted("revision author limit exceeded"));
        }
        self.authors.push(author.to_string());
        let bytes = serialize_authors(&self.authors)?;
        append_table_block(&mut self.word, &mut self.table, STTBFRMARK, &bytes)?;
        Ok((self.authors.len() - 1) as u16)
    }

    fn rewrite_chpx(&mut self) -> Result<()> {
        if self.chpx.is_empty() {
            return Err(corrupted("CHPX table has no runs"));
        }
        self.chpx.sort_by_key(|r| r.start);
        if self
            .chpx
            .iter()
            .any(|r| r.start >= r.end || r.grpprl.len() > 255)
        {
            return Err(corrupted("CHPX run is invalid or exceeds FKP limits"));
        }
        let mut builder = ChpxFkpBuilder::new();
        for run in &self.chpx {
            builder.add_entry(run.start, run.end, run.grpprl.clone());
        }
        let pages = builder.generate_pages().map_err(PackageError::from)?;
        let base = align512(self.word.len())?;
        self.word.resize(base, 0);
        let mut pns = Vec::new();
        for page in &pages.pages {
            pns.push(
                u32::try_from(self.word.len() / 512)
                    .map_err(|_| corrupted("CHPX page number exceeds u32"))?,
            );
            self.word.extend_from_slice(page);
        }
        let mut plc = Vec::new();
        for (start, _) in &pages.ranges {
            plc.extend_from_slice(&start.to_le_bytes());
        }
        plc.extend_from_slice(
            &pages
                .ranges
                .last()
                .ok_or_else(|| corrupted("CHPX page list is empty"))?
                .1
                .to_le_bytes(),
        );
        for pn in pns {
            plc.extend_from_slice(&pn.to_le_bytes());
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_CHPX, &plc)
    }

    fn rewrite_papx(&mut self) -> Result<()> {
        self.papx.sort_by_key(|r| r.start);
        let pages = build_papx_pages(&self.papx)?;
        let base = align512(self.word.len())?;
        self.word.resize(base, 0);
        let mut plc = Vec::new();
        for page in &pages {
            plc.extend_from_slice(&page.start.to_le_bytes());
        }
        plc.extend_from_slice(
            &pages
                .last()
                .ok_or_else(|| corrupted("PAPX page list is empty"))?
                .end
                .to_le_bytes(),
        );
        for page in pages {
            let pn = u32::try_from(self.word.len() / 512)
                .map_err(|_| corrupted("PAPX page number exceeds u32"))?;
            plc.extend_from_slice(&pn.to_le_bytes());
            self.word.extend_from_slice(&page.bytes);
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_PAPX, &plc)
    }

    fn shift_cp_tables(&mut self, start: u32, removed: u32, added: u32) -> Result<()> {
        let end = start
            .checked_add(removed)
            .ok_or_else(|| corrupted("CP range overflow"))?;
        for table in &mut self.cp_tables {
            for cp in &mut table.cps {
                *cp = if removed == 0 {
                    if *cp >= start {
                        cp.checked_add(added)
                            .ok_or_else(|| corrupted("PLCF CP overflow"))?
                    } else {
                        *cp
                    }
                } else if *cp <= start {
                    *cp
                } else if *cp >= end {
                    cp.checked_sub(removed)
                        .and_then(|v| v.checked_add(added))
                        .ok_or_else(|| corrupted("PLCF CP shift overflow"))?
                } else {
                    start
                        .checked_add(added)
                        .ok_or_else(|| corrupted("PLCF CP overflow"))?
                };
            }
            if table.cps.windows(2).any(|v| v[0] > v[1]) {
                return Err(corrupted("PLCF CPs became non-monotonic"));
            }
        }
        Ok(())
    }

    fn append_clx_and_cp_tables(&mut self) -> Result<()> {
        let clx = serialize_clx(&self.pieces)?;
        append_table_block(&mut self.word, &mut self.table, CLX, &clx)?;
        for table in &self.cp_tables {
            let mut bytes = Vec::new();
            for cp in &table.cps {
                bytes.extend_from_slice(&cp.to_le_bytes());
            }
            bytes.extend_from_slice(&table.records);
            append_table_block(&mut self.word, &mut self.table, table.index, &bytes)?;
        }
        Ok(())
    }

    fn enable_tracking(&mut self) -> Result<()> {
        let (offset, length) = fib_pair(&self.word, DOP)?;
        if length < 84 {
            return Err(corrupted("DOP is too short to enable revision tracking"));
        }
        let mut dop = slice(&self.table, offset, length, "DOP")?.to_vec();
        dop[5] |= 0x80;
        append_table_block(&mut self.word, &mut self.table, DOP, &dop)
    }

    fn patch_sizes(&mut self) -> Result<()> {
        let word_len =
            u32::try_from(self.word.len()).map_err(|_| corrupted("WordDocument exceeds u32"))?;
        put_u32(&mut self.word, FIB_CCP_TEXT, self.main_ccp)?;
        put_u32(&mut self.word, 28, word_len)?;
        put_u32(&mut self.word, 64, word_len)
    }

    fn commit(&mut self) -> Result<()> {
        self.package
            .put_stream(&self.word_path, self.word.clone())
            .map_err(PackageError::from)?;
        self.package
            .put_stream(&self.table_path, self.table.clone())
            .map_err(PackageError::from)?;
        self.changed = true;
        Ok(())
    }
}
