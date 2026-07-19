//! Transactional tracked-revision editing for Word 97+ binary documents.
//!
//! The editor deliberately uses append-only replacement tables and FKPs. It
//! never evaluates fields or macros and retains unknown SPRMs and unrelated
//! compound-file streams byte-for-byte.

use super::package::{DocError, Result};
use super::CommentDateTime;
use crate::doc::parts::fkp::{ChpxFkp, PapxFkp, ParagraphHeight};
use crate::doc::writer::ChpxFkpBuilder;
use crate::sprm::parse_sprms;
use crate::sprm_operations::*;
use crate::{LegacyOfficeObjectEditor, LegacyOfficeObjectFormat, LegacyOfficeObjectLimits};
use std::collections::{BTreeSet, HashMap};

const FIB_CCP_TEXT: usize = 76;
const FIB_FC_LCB: usize = 154;
const PLCFANDREF: usize = 4;
const PLCFBTE_CHPX: usize = 12;
const PLCFBTE_PAPX: usize = 13;
const PLCFFLD_MOM: usize = 16;
const PLCFBKF: usize = 22;
const PLCFBKL: usize = 23;
const DOP: usize = 31;
const CLX: usize = 33;
const PLCFATNBKF: usize = 42;
const PLCFATNBKL: usize = 43;
const STTBFRMARK: usize = 51;
const MAX_PIECES: usize = 65_536;
const MAX_REVISIONS: usize = 65_536;
const MAX_AUTHORS: usize = i16::MAX as usize;
const MAX_TEXT_UNITS: usize = 16 * 1024 * 1024;

/// A revision representation supported by binary DOC.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum DocTrackedRevisionKind {
    Insertion,
    Deletion,
    /// The deletion half of a move, paired by `revision_save_id`.
    MoveFrom,
    /// The insertion half of a move, paired by `revision_save_id`.
    MoveTo,
    CharacterFormatting,
    ParagraphFormatting,
    TableRowFormatting,
}

/// Metadata used to author or replace a revision mark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocTrackedRevisionMetadata {
    pub author: String,
    pub timestamp: Option<CommentDateTime>,
    pub reason: Option<u16>,
    pub revision_save_id: Option<u32>,
}

impl DocTrackedRevisionMetadata {
    pub fn new(author: impl Into<String>) -> Self {
        Self { author: author.into(), timestamp: None, reason: None, revision_save_id: None }
    }

    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }

    pub fn with_reason(mut self, reason: u16) -> Self {
        self.reason = Some(reason);
        self
    }

    pub fn with_revision_save_id(mut self, revision_save_id: u32) -> Self {
        self.revision_save_id = Some(revision_save_id);
        self
    }
}

/// A tracked range in main-story CP coordinates.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocTrackedRevision {
    pub kind: DocTrackedRevisionKind,
    pub start_cp: u32,
    pub end_cp: u32,
    pub author_index: u16,
    pub author: String,
    pub timestamp: Option<CommentDateTime>,
    pub reason: Option<u16>,
    pub revision_save_id: Option<u32>,
    /// Move pair identity when binary insertion/deletion marks share an RSID.
    pub move_pair_id: Option<u32>,
}

#[derive(Clone, Debug)]
struct RawPiece {
    start: u32,
    end: u32,
    fc: u32,
    unicode: bool,
    prefix: [u8; 2],
    prm: [u8; 2],
}

#[derive(Clone, Debug)]
struct FcRun {
    start: u32,
    end: u32,
    grpprl: Vec<u8>,
}

#[derive(Clone, Debug)]
struct PapxRun {
    start: u32,
    end: u32,
    grpprl: Vec<u8>,
    phe: ParagraphHeight,
}

#[derive(Clone, Debug)]
struct CpTable {
    index: usize,
    cps: Vec<u32>,
    records: Vec<u8>,
}

/// Atomic editor for tracked revisions in an existing binary DOC.
#[derive(Clone)]
pub struct DocTrackedRevisionEditor {
    package: LegacyOfficeObjectEditor,
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
    pub fn open(bytes: Vec<u8>, limits: LegacyOfficeObjectLimits) -> Result<Self> {
        let package = LegacyOfficeObjectEditor::open(&bytes, LegacyOfficeObjectFormat::Doc, limits)
            .map_err(DocError::from)?;
        let word_path = vec!["WordDocument".to_string()];
        let word = package.package_stream(&word_path)
            .ok_or_else(|| corrupted("WordDocument stream is missing"))?.to_vec();
        if word.len() < FIB_FC_LCB + (STTBFRMARK + 1) * 8 || u16_at(&word, 0)? != 0xA5EC {
            return Err(corrupted("tracked-revision editing requires Word 97+ FIB data"));
        }
        let flags = u16_at(&word, 10)?;
        if flags & 0x0100 != 0 || u32_at(&word, 14)? != 0 {
            return Err(corrupted("encrypted DOC cannot be edited"));
        }
        let table_path = vec![if flags & 0x0200 != 0 { "1Table" } else { "0Table" }.to_string()];
        let table = package.package_stream(&table_path)
            .ok_or_else(|| corrupted("selected Table stream is missing"))?.to_vec();
        reject_protection(&word, &table)?;
        let main_ccp = u32_at(&word, FIB_CCP_TEXT)?;
        let pieces = parse_clx(&word, &table)?;
        if pieces.last().is_none_or(|piece| piece.end < main_ccp) {
            return Err(corrupted("piece table does not cover the main story"));
        }
        let chpx = parse_chpx(&word, &table)?;
        let papx = parse_papx(&word, &table)?;
        let authors = parse_authors(&word, &table)?;
        let cp_tables = [(PLCFANDREF, 30), (PLCFFLD_MOM, 2), (PLCFBKF, 4),
            (PLCFBKL, 0), (PLCFATNBKF, 4), (PLCFATNBKL, 0)]
            .into_iter().filter_map(|(index, size)| parse_cp_table(&word, &table, index, size).transpose())
            .collect::<Result<Vec<_>>>()?;
        let editor = Self { package, word_path, table_path, word, table, pieces, chpx,
            papx, authors, cp_tables, main_ccp, changed: false };
        if editor.revisions()?.len() > MAX_REVISIONS {
            return Err(corrupted("revision count exceeds resource limit"));
        }
        Ok(editor)
    }

    pub fn is_changed(&self) -> bool { self.changed }

    pub fn authors(&self) -> &[String] { &self.authors }

    /// Lists character and PAPX property revisions, merging adjacent runs with
    /// identical metadata even when the range crosses piece boundaries.
    pub fn revisions(&self) -> Result<Vec<DocTrackedRevision>> {
        let mut output = Vec::new();
        for run in &self.chpx {
            let sprms = strict_sprms(&run.grpprl)?;
            for (kind, flag, author_op, time_op, reason_op, rsid_op) in [
                (DocTrackedRevisionKind::Insertion, SPRM_C_F_RMARK, SPRM_C_IBST_RMARK,
                    SPRM_C_DTTM_RMARK, SPRM_C_IDSL_RMARK, SPRM_C_RSID_TEXT),
                (DocTrackedRevisionKind::Deletion, SPRM_C_F_RMARK_DEL, SPRM_C_IBST_RMARK_DEL,
                    SPRM_C_DTTM_RMARK_DEL, SPRM_C_IDSL_RMARK_DEL, SPRM_C_RSID_RM_DEL),
            ] {
                if sprms.iter().any(|s| s.opcode == flag && s.operand_byte() == Some(1)) {
                    let metadata = metadata_from_sprms(&sprms, author_op, time_op, reason_op, rsid_op, &self.authors)?;
                    self.push_fc_revision(&mut output, run.start, run.end, kind, metadata)?;
                }
            }
            for opcode in [SPRM_C_PROP_RMARK90, SPRM_C_PROP_RMARK_CURRENT] {
                if let Some(mark) = sprms.iter().rev().find(|s| s.opcode == opcode) {
                    if mark.operand_bytes().first() == Some(&1) {
                        let metadata = property_metadata(mark.operand_bytes(), &sprms, SPRM_C_RSID_PROP, &self.authors)?;
                        self.push_fc_revision(&mut output, run.start, run.end,
                            DocTrackedRevisionKind::CharacterFormatting, metadata)?;
                    }
                    break;
                }
            }
        }
        for run in &self.papx {
            let body = run.grpprl.get(2..).ok_or_else(|| corrupted("PAPX has no style index"))?;
            let sprms = strict_sprms(body)?;
            for (kind, op, rsid) in [
                (DocTrackedRevisionKind::ParagraphFormatting,
                    [SPRM_P_PROP_RMARK, SPRM_P_PROP_RMARK90, SPRM_P_PROP_RMARK_CURRENT].as_slice(), None),
                (DocTrackedRevisionKind::TableRowFormatting, [SPRM_T_PROP_RMARK].as_slice(), Some(SPRM_T_RSID)),
            ] {
                if let Some(mark) = sprms.iter().rev().find(|s| op.contains(&s.opcode)) {
                    if mark.operand_bytes().first() == Some(&1) {
                        let metadata = property_metadata(mark.operand_bytes(), &sprms, rsid.unwrap_or(0), &self.authors)?;
                        self.push_fc_revision(&mut output, run.start, run.end, kind, metadata)?;
                    }
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
    pub fn add_text(&mut self, cp: u32, text: &str, kind: DocTrackedRevisionKind,
        metadata: DocTrackedRevisionMetadata) -> Result<DocTrackedRevision> {
        if !matches!(kind, DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::MoveTo) {
            return Err(corrupted("add_text requires an insertion or move-to revision"));
        }
        let units = text.encode_utf16().collect::<Vec<_>>();
        if units.is_empty() || units.len() > MAX_TEXT_UNITS || units.iter().any(|u| matches!(*u, 0..=8 | 11..=31 | 0xFFFC)) {
            return Err(corrupted("tracked text is empty, oversized, or contains an active control character"));
        }
        if cp > self.main_ccp { return Err(corrupted("tracked insertion CP exceeds main story")); }
        validate_metadata(kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        for unit in &units { candidate.word.extend_from_slice(&unit.to_le_bytes()); }
        let length = u32::try_from(units.len()).map_err(|_| corrupted("tracked text length exceeds u32"))?;
        insert_piece(&mut candidate.pieces, cp, length, u32::try_from(fc).map_err(|_| corrupted("FC exceeds u32"))?)?;
        candidate.shift_cp_tables(cp, 0, length)?;
        candidate.main_ccp = candidate.main_ccp.checked_add(length).ok_or_else(|| corrupted("main story CP overflow"))?;
        let grpprl = encode_revision(kind, author, &metadata)?;
        let fc_end = u32::try_from(candidate.word.len()).map_err(|_| corrupted("FC exceeds u32"))?;
        candidate.chpx.push(FcRun { start: fc as u32, end: fc_end, grpprl });
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
    pub fn add(&mut self, start_cp: u32, end_cp: u32, kind: DocTrackedRevisionKind,
        metadata: DocTrackedRevisionMetadata) -> Result<DocTrackedRevision> {
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
    pub fn update(&mut self, index: usize, metadata: DocTrackedRevisionMetadata) -> Result<DocTrackedRevision> {
        let revision = self.revisions()?.get(index).cloned().ok_or_else(|| corrupted("revision index is out of range"))?;
        validate_metadata(revision.kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        candidate.mutate_mark(revision.start_cp, revision.end_cp, revision.kind, Some((author, &metadata)))?;
        candidate.commit()?;
        *self = candidate;
        self.find_exact(revision.start_cp, revision.end_cp, revision.kind)
    }

    /// Removes a mark while retaining its text/current formatting.
    pub fn remove(&mut self, index: usize) -> Result<DocTrackedRevision> {
        let revision = self.revisions()?.get(index).cloned().ok_or_else(|| corrupted("revision index is out of range"))?;
        let mut candidate = self.clone();
        candidate.mutate_mark(revision.start_cp, revision.end_cp, revision.kind, None)?;
        candidate.commit()?;
        *self = candidate;
        Ok(revision)
    }

    /// Accepts a revision using Word redline semantics.
    pub fn accept(&mut self, index: usize) -> Result<DocTrackedRevision> {
        let revision = self.revisions()?.get(index).cloned().ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(revision.kind, DocTrackedRevisionKind::Deletion | DocTrackedRevisionKind::MoveFrom) {
            self.delete_revision_text(&revision)?;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    /// Rejects a revision using Word redline semantics.
    pub fn reject(&mut self, index: usize) -> Result<DocTrackedRevision> {
        let revision = self.revisions()?.get(index).cloned().ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(revision.kind, DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::MoveTo) {
            self.delete_revision_text(&revision)?;
        } else if matches!(revision.kind, DocTrackedRevisionKind::CharacterFormatting |
            DocTrackedRevisionKind::ParagraphFormatting | DocTrackedRevisionKind::TableRowFormatting) {
            let mut candidate = self.clone();
            candidate.reject_formatting_revision(&revision)?;
            candidate.commit()?;
            *self = candidate;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    pub fn finish(self) -> Result<Vec<u8>> { self.package.finish().map_err(DocError::from) }

    fn delete_revision_text(&mut self, revision: &DocTrackedRevision) -> Result<()> {
        self.reject_destructive_interactions(revision.start_cp, revision.end_cp)?;
        let mut candidate = self.clone();
        delete_piece_range(&mut candidate.pieces, revision.start_cp, revision.end_cp)?;
        let removed = revision.end_cp - revision.start_cp;
        candidate.shift_cp_tables(revision.start_cp, removed, 0)?;
        candidate.main_ccp = candidate.main_ccp.checked_sub(removed).ok_or_else(|| corrupted("main story CP underflow"))?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    fn reject_destructive_interactions(&self, start: u32, end: u32) -> Result<()> {
        let text = read_units(&self.word, &self.pieces, start, end)?;
        if text.iter().any(|u| matches!(*u, 0x13 | 0x14 | 0x15)) {
            return Err(corrupted("accept/reject would delete a field boundary"));
        }
        for table in &self.cp_tables {
            for cp in &table.cps {
                if *cp > start && *cp < end {
                    return Err(corrupted("accept/reject would split a field, bookmark, or comment range"));
                }
            }
        }
        Ok(())
    }

    fn mutate_mark(&mut self, start: u32, end: u32, kind: DocTrackedRevisionKind,
        replacement: Option<(u16, &DocTrackedRevisionMetadata)>) -> Result<()> {
        let intervals = self.fc_intervals(start, end)?;
        match kind {
            DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::Deletion |
            DocTrackedRevisionKind::MoveFrom | DocTrackedRevisionKind::MoveTo |
            DocTrackedRevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    replace_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_chpx()?;
            }
            DocTrackedRevisionKind::ParagraphFormatting | DocTrackedRevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    replace_papx_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_papx()?;
            }
        }
        if let Some((_, metadata)) = replacement {
            if !self.authors.iter().any(|a| a == &metadata.author) {
                return Err(corrupted("revision author indexing failed"));
            }
        }
        Ok(())
    }

    fn reject_formatting_revision(&mut self, revision: &DocTrackedRevision) -> Result<()> {
        let intervals = self.fc_intervals(revision.start_cp, revision.end_cp)?;
        match revision.kind {
            DocTrackedRevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    restore_before_wall(grp, SPRM_C_WALL, &revision_opcodes(
                        DocTrackedRevisionKind::CharacterFormatting, true))
                })?;
                self.rewrite_chpx()
            }
            DocTrackedRevisionKind::ParagraphFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp.get(..2).ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(&grp[2..], SPRM_P_WALL,
                        &revision_opcodes(DocTrackedRevisionKind::ParagraphFormatting, true))?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            }
            DocTrackedRevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp.get(..2).ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(&grp[2..], SPRM_T_WALL,
                        &revision_opcodes(DocTrackedRevisionKind::TableRowFormatting, true))?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            }
            _ => Err(corrupted("revision is not a formatting revision")),
        }
    }

    fn fc_intervals(&self, start: u32, end: u32) -> Result<Vec<(u32, u32)>> {
        let mut output = Vec::new();
        for piece in &self.pieces {
            let left = start.max(piece.start);
            let right = end.min(piece.end);
            if left >= right { continue; }
            let scale = if piece.unicode { 2 } else { 1 };
            let fc_start = piece.fc.checked_add((left - piece.start) * scale).ok_or_else(|| corrupted("FC overflow"))?;
            let fc_end = piece.fc.checked_add((right - piece.start) * scale).ok_or_else(|| corrupted("FC overflow"))?;
            output.push((fc_start, fc_end));
        }
        if output.is_empty() { return Err(corrupted("revision range has no text pieces")); }
        Ok(output)
    }

    fn push_fc_revision(&self, output: &mut Vec<DocTrackedRevision>, fc_start: u32, fc_end: u32,
        kind: DocTrackedRevisionKind, metadata: ParsedMetadata) -> Result<()> {
        for piece in &self.pieces {
            let width = if piece.unicode { 2 } else { 1 };
            let piece_fc_end = piece.fc.checked_add((piece.end - piece.start) * width).ok_or_else(|| corrupted("piece FC overflow"))?;
            let left = fc_start.max(piece.fc);
            let right = fc_end.min(piece_fc_end);
            if left >= right { continue; }
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

    fn find_exact(&self, start: u32, end: u32, kind: DocTrackedRevisionKind) -> Result<DocTrackedRevision> {
        self.revisions()?.into_iter().find(|r| r.start_cp == start && r.end_cp == end &&
            (r.kind == kind || matches!((r.kind, kind),
                (DocTrackedRevisionKind::Insertion, DocTrackedRevisionKind::MoveTo) |
                (DocTrackedRevisionKind::Deletion, DocTrackedRevisionKind::MoveFrom))))
            .ok_or_else(|| corrupted("authored revision was not discoverable"))
    }

    fn author_index(&mut self, author: &str) -> Result<u16> {
        if author.is_empty() || author.encode_utf16().count() > u16::MAX as usize {
            return Err(corrupted("revision author is empty or too long"));
        }
        if self.authors.is_empty() { self.authors.push("Unknown".to_string()); }
        if let Some(index) = self.authors.iter().position(|value| value == author) {
            return u16::try_from(index).map_err(|_| corrupted("revision author index exceeds u16"));
        }
        if self.authors.len() >= MAX_AUTHORS { return Err(corrupted("revision author limit exceeded")); }
        self.authors.push(author.to_string());
        let bytes = serialize_authors(&self.authors)?;
        append_table_block(&mut self.word, &mut self.table, STTBFRMARK, &bytes)?;
        Ok((self.authors.len() - 1) as u16)
    }

    fn rewrite_chpx(&mut self) -> Result<()> {
        if self.chpx.is_empty() { return Err(corrupted("CHPX table has no runs")); }
        self.chpx.sort_by_key(|r| r.start);
        if self.chpx.iter().any(|r| r.start >= r.end || r.grpprl.len() > 255) {
            return Err(corrupted("CHPX run is invalid or exceeds FKP limits"));
        }
        let mut builder = ChpxFkpBuilder::new();
        for run in &self.chpx { builder.add_entry(run.start, run.end, run.grpprl.clone()); }
        let pages = builder.generate_pages().map_err(DocError::from)?;
        let base = align512(self.word.len())?;
        self.word.resize(base, 0);
        let mut pns = Vec::new();
        for page in &pages.pages {
            pns.push(u32::try_from(self.word.len() / 512).map_err(|_| corrupted("CHPX page number exceeds u32"))?);
            self.word.extend_from_slice(page);
        }
        let mut plc = Vec::new();
        for (start, _) in &pages.ranges { plc.extend_from_slice(&start.to_le_bytes()); }
        plc.extend_from_slice(&pages.ranges.last().ok_or_else(|| corrupted("CHPX page list is empty"))?.1.to_le_bytes());
        for pn in pns { plc.extend_from_slice(&pn.to_le_bytes()); }
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
        plc.extend_from_slice(&pages.last().ok_or_else(|| corrupted("PAPX page list is empty"))?.end.to_le_bytes());
        for page in pages {
            let pn = u32::try_from(self.word.len() / 512).map_err(|_| corrupted("PAPX page number exceeds u32"))?;
            plc.extend_from_slice(&pn.to_le_bytes());
            self.word.extend_from_slice(&page.bytes);
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_PAPX, &plc)
    }

    fn shift_cp_tables(&mut self, start: u32, removed: u32, added: u32) -> Result<()> {
        let end = start.checked_add(removed).ok_or_else(|| corrupted("CP range overflow"))?;
        for table in &mut self.cp_tables {
            for cp in &mut table.cps {
                *cp = if removed == 0 {
                    if *cp >= start { cp.checked_add(added).ok_or_else(|| corrupted("PLCF CP overflow"))? } else { *cp }
                } else if *cp <= start { *cp } else if *cp >= end {
                    cp.checked_sub(removed).and_then(|v| v.checked_add(added)).ok_or_else(|| corrupted("PLCF CP shift overflow"))?
                } else { start.checked_add(added).ok_or_else(|| corrupted("PLCF CP overflow"))? };
            }
            if table.cps.windows(2).any(|v| v[0] > v[1]) { return Err(corrupted("PLCF CPs became non-monotonic")); }
        }
        Ok(())
    }

    fn append_clx_and_cp_tables(&mut self) -> Result<()> {
        let clx = serialize_clx(&self.pieces)?;
        append_table_block(&mut self.word, &mut self.table, CLX, &clx)?;
        for table in &self.cp_tables {
            let mut bytes = Vec::new();
            for cp in &table.cps { bytes.extend_from_slice(&cp.to_le_bytes()); }
            bytes.extend_from_slice(&table.records);
            append_table_block(&mut self.word, &mut self.table, table.index, &bytes)?;
        }
        Ok(())
    }

    fn enable_tracking(&mut self) -> Result<()> {
        let (offset, length) = fib_pair(&self.word, DOP)?;
        if length < 84 { return Err(corrupted("DOP is too short to enable revision tracking")); }
        let mut dop = slice(&self.table, offset, length, "DOP")?.to_vec();
        dop[5] |= 0x80;
        append_table_block(&mut self.word, &mut self.table, DOP, &dop)
    }

    fn patch_sizes(&mut self) -> Result<()> {
        let word_len = u32::try_from(self.word.len()).map_err(|_| corrupted("WordDocument exceeds u32"))?;
        put_u32(&mut self.word, FIB_CCP_TEXT, self.main_ccp)?;
        put_u32(&mut self.word, 28, word_len)?;
        put_u32(&mut self.word, 64, word_len)
    }

    fn commit(&mut self) -> Result<()> {
        self.package.replace_package_stream(&self.word_path, self.word.clone()).map_err(DocError::from)?;
        self.package.replace_package_stream(&self.table_path, self.table.clone()).map_err(DocError::from)?;
        self.changed = true;
        Ok(())
    }
}

#[derive(Clone)]
struct ParsedMetadata { author_index: u16, author: String, timestamp: Option<CommentDateTime>, reason: Option<u16>, rsid: Option<u32> }

impl ParsedMetadata {
    fn to_revision(&self, kind: DocTrackedRevisionKind, start_cp: u32, end_cp: u32) -> DocTrackedRevision {
        DocTrackedRevision { kind, start_cp, end_cp, author_index: self.author_index,
            author: self.author.clone(), timestamp: self.timestamp, reason: self.reason,
            revision_save_id: self.rsid, move_pair_id: None }
    }
}

fn metadata_from_sprms(sprms: &[crate::sprm::Sprm], author_op: u16, time_op: u16,
    reason_op: u16, rsid_op: u16, authors: &[String]) -> Result<ParsedMetadata> {
    let author_index = sprms.iter().rev().find(|s| s.opcode == author_op).and_then(|s| s.operand_word()).unwrap_or(0);
    let author = authors.get(author_index as usize).ok_or_else(|| corrupted("revision author index exceeds SttbfRMark"))?.clone();
    let timestamp = sprms.iter().rev().find(|s| s.opcode == time_op).and_then(|s| s.operand_dword()).map(decode_dttm).transpose()?.flatten();
    let reason = sprms.iter().rev().find(|s| s.opcode == reason_op).and_then(|s| s.operand_word());
    let rsid = sprms.iter().rev().find(|s| s.opcode == rsid_op).and_then(|s| s.operand_dword());
    Ok(ParsedMetadata { author_index, author, timestamp, reason, rsid })
}

fn property_metadata(operand: &[u8], sprms: &[crate::sprm::Sprm], rsid_op: u16,
    authors: &[String]) -> Result<ParsedMetadata> {
    if operand.len() != 7 { return Err(corrupted("property revision operand must be seven bytes")); }
    let author_index = u16::from_le_bytes([operand[1], operand[2]]);
    let author = authors.get(author_index as usize).ok_or_else(|| corrupted("property revision author exceeds SttbfRMark"))?.clone();
    let raw = u32::from_le_bytes(operand[3..7].try_into().unwrap());
    let timestamp = decode_dttm(raw)?;
    let rsid = (rsid_op != 0).then(|| sprms.iter().rev().find(|s| s.opcode == rsid_op).and_then(|s| s.operand_dword())).flatten();
    Ok(ParsedMetadata { author_index, author, timestamp, reason: None, rsid })
}

fn validate_metadata(kind: DocTrackedRevisionKind, metadata: &DocTrackedRevisionMetadata) -> Result<()> {
    if metadata.author.is_empty() { return Err(corrupted("revision author must not be empty")); }
    if metadata.reason.is_some_and(|v| v > 0x2B) { return Err(corrupted("revision reason is undefined")); }
    if matches!(kind, DocTrackedRevisionKind::MoveFrom | DocTrackedRevisionKind::MoveTo) && metadata.revision_save_id.is_none() {
        return Err(corrupted("move revisions require a shared revision_save_id"));
    }
    if let Some(value) = metadata.timestamp { let _ = pack_dttm(Some(value))?; }
    Ok(())
}

fn encode_revision(kind: DocTrackedRevisionKind, author: u16, metadata: &DocTrackedRevisionMetadata) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    match kind {
        DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::MoveTo => {
            push_byte(&mut output, SPRM_C_F_RMARK, 1); push_word(&mut output, SPRM_C_IBST_RMARK, author);
            if metadata.timestamp.is_some() { push_dword(&mut output, SPRM_C_DTTM_RMARK, pack_dttm(metadata.timestamp)?); }
            if let Some(v) = metadata.reason { push_word(&mut output, SPRM_C_IDSL_RMARK, v); }
            if let Some(v) = metadata.revision_save_id { push_dword(&mut output, SPRM_C_RSID_TEXT, v); }
        }
        DocTrackedRevisionKind::Deletion | DocTrackedRevisionKind::MoveFrom => {
            push_byte(&mut output, SPRM_C_F_RMARK_DEL, 1); push_word(&mut output, SPRM_C_IBST_RMARK_DEL, author);
            if metadata.timestamp.is_some() { push_dword(&mut output, SPRM_C_DTTM_RMARK_DEL, pack_dttm(metadata.timestamp)?); }
            if let Some(v) = metadata.reason { push_word(&mut output, SPRM_C_IDSL_RMARK_DEL, v); }
            if let Some(v) = metadata.revision_save_id { push_dword(&mut output, SPRM_C_RSID_RM_DEL, v); }
        }
        DocTrackedRevisionKind::CharacterFormatting => {
            output.extend_from_slice(&SPRM_C_PROP_RMARK_CURRENT.to_le_bytes()); output.push(7); output.push(1);
            output.extend_from_slice(&author.to_le_bytes()); output.extend_from_slice(&pack_dttm(metadata.timestamp)?.to_le_bytes());
            if let Some(v) = metadata.reason { push_word(&mut output, SPRM_C_IDSL_RMARK, v); }
            if let Some(v) = metadata.revision_save_id { push_dword(&mut output, SPRM_C_RSID_PROP, v); }
        }
        DocTrackedRevisionKind::ParagraphFormatting => {
            output.extend_from_slice(&SPRM_P_PROP_RMARK_CURRENT.to_le_bytes()); output.push(7); output.push(1);
            output.extend_from_slice(&author.to_le_bytes()); output.extend_from_slice(&pack_dttm(metadata.timestamp)?.to_le_bytes());
        }
        DocTrackedRevisionKind::TableRowFormatting => {
            output.extend_from_slice(&SPRM_T_PROP_RMARK.to_le_bytes()); output.push(7); output.push(1);
            output.extend_from_slice(&author.to_le_bytes()); output.extend_from_slice(&pack_dttm(metadata.timestamp)?.to_le_bytes());
            if let Some(v) = metadata.revision_save_id { push_dword(&mut output, SPRM_T_RSID, v); }
        }
    }
    Ok(output)
}

fn replace_revision_sprms(grp: &[u8], kind: DocTrackedRevisionKind,
    replacement: Option<(u16, &DocTrackedRevisionMetadata)>) -> Result<Vec<u8>> {
    let remove = revision_opcodes(kind, replacement.is_none());
    let mut output = retain_sprms(grp, &remove)?;
    if let Some((author, metadata)) = replacement { output.extend_from_slice(&encode_revision(kind, author, metadata)?); }
    if output.len() > 255 { return Err(corrupted("edited CHPX exceeds one-byte FKP limit")); }
    Ok(output)
}

fn replace_papx_revision_sprms(grp: &[u8], kind: DocTrackedRevisionKind,
    replacement: Option<(u16, &DocTrackedRevisionMetadata)>) -> Result<Vec<u8>> {
    let style = grp.get(..2).ok_or_else(|| corrupted("PAPX style index is truncated"))?;
    let body = grp.get(2..).unwrap();
    let mut output = style.to_vec();
    output.extend_from_slice(&retain_sprms(body, &revision_opcodes(kind, replacement.is_none()))?);
    if let Some((author, metadata)) = replacement { output.extend_from_slice(&encode_revision(kind, author, metadata)?); }
    if output.len() > 510 { return Err(corrupted("edited PAPX exceeds FKP limit")); }
    Ok(output)
}

fn revision_opcodes(kind: DocTrackedRevisionKind, remove_wall: bool) -> Vec<u16> {
    match kind {
        DocTrackedRevisionKind::Insertion | DocTrackedRevisionKind::MoveTo => vec![SPRM_C_F_RMARK, SPRM_C_IBST_RMARK, SPRM_C_DTTM_RMARK, SPRM_C_IDSL_RMARK, SPRM_C_RSID_TEXT],
        DocTrackedRevisionKind::Deletion | DocTrackedRevisionKind::MoveFrom => vec![SPRM_C_F_RMARK_DEL, SPRM_C_IBST_RMARK_DEL, SPRM_C_DTTM_RMARK_DEL, SPRM_C_IDSL_RMARK_DEL, SPRM_C_RSID_RM_DEL],
        DocTrackedRevisionKind::CharacterFormatting => {
            let mut value = vec![SPRM_C_PROP_RMARK90, SPRM_C_PROP_RMARK_CURRENT, SPRM_C_RSID_PROP];
            if remove_wall { value.push(SPRM_C_WALL); }
            value
        }
        DocTrackedRevisionKind::ParagraphFormatting => {
            let mut value = vec![SPRM_P_PROP_RMARK, SPRM_P_PROP_RMARK90, SPRM_P_PROP_RMARK_CURRENT];
            if remove_wall { value.push(SPRM_P_WALL); }
            value
        }
        DocTrackedRevisionKind::TableRowFormatting => {
            let mut value = vec![SPRM_T_PROP_RMARK, SPRM_T_RSID];
            if remove_wall { value.push(SPRM_T_WALL); }
            value
        }
    }
}

fn restore_before_wall(grp: &[u8], wall: u16, revision_marks: &[u16]) -> Result<Vec<u8>> {
    let sprms = strict_sprms(grp)?;
    if let Some(marker) = sprms.iter().rfind(|sprm| sprm.opcode == wall && sprm.operand_byte() == Some(1)) {
        Ok(grp[..marker.offset].to_vec())
    } else {
        // Some Word 97-era producers emitted PropRMark without a wall. There
        // is no recoverable previous state, so rejection safely clears only
        // the revision metadata and retains the current properties.
        retain_sprms(grp, revision_marks)
    }
}

fn retain_sprms(grp: &[u8], remove: &[u16]) -> Result<Vec<u8>> {
    let sprms = strict_sprms(grp)?;
    let mut output = Vec::with_capacity(grp.len());
    for sprm in sprms {
        if !remove.contains(&sprm.opcode) { output.extend_from_slice(&grp[sprm.offset..sprm.offset + sprm.size]); }
    }
    Ok(output)
}

fn strict_sprms(grp: &[u8]) -> Result<Vec<crate::sprm::Sprm>> {
    if grp.is_empty() { return Ok(Vec::new()); }
    let sprms = parse_sprms(grp);
    if sprms.last().map(|s| s.offset + s.size) != Some(grp.len()) {
        return Err(corrupted("malformed or overlapping SPRM sequence"));
    }
    Ok(sprms)
}

fn split_transform_chpx(runs: &mut Vec<FcRun>, intervals: &[(u32, u32)],
    transform: impl Fn(&[u8]) -> Result<Vec<u8>>) -> Result<()> {
    let mut output = Vec::new();
    for run in runs.iter() {
        let mut cuts = BTreeSet::from([run.start, run.end]);
        for (a, b) in intervals { if *a > run.start && *a < run.end { cuts.insert(*a); } if *b > run.start && *b < run.end { cuts.insert(*b); } }
        let cuts = cuts.into_iter().collect::<Vec<_>>();
        for pair in cuts.windows(2) {
            let covered = intervals.iter().any(|(a, b)| pair[0] >= *a && pair[1] <= *b);
            output.push(FcRun { start: pair[0], end: pair[1], grpprl: if covered { transform(&run.grpprl)? } else { run.grpprl.clone() } });
        }
    }
    if !intervals.iter().all(|(a, b)| {
        let mut cursor = *a;
        for run in output.iter().filter(|run| run.end > *a && run.start < *b) {
            if run.start > cursor { return false; }
            cursor = cursor.max(run.end);
            if cursor >= *b { return true; }
        }
        false
    }) {
        return Err(corrupted("tracked range is outside CHPX coverage"));
    }
    merge_fc_runs(&mut output);
    *runs = output;
    Ok(())
}

fn split_transform_papx(runs: &mut Vec<PapxRun>, intervals: &[(u32, u32)],
    transform: impl Fn(&[u8]) -> Result<Vec<u8>>) -> Result<()> {
    let mut hit = false;
    for run in runs.iter_mut() {
        if intervals.iter().any(|(a, b)| *a < run.end && *b > run.start) {
            run.grpprl = transform(&run.grpprl)?;
            hit = true;
        }
    }
    if !hit { return Err(corrupted("tracked property range is outside PAPX coverage")); }
    Ok(())
}

fn parse_chpx(word: &[u8], table: &[u8]) -> Result<Vec<FcRun>> {
    let (fcs, pages) = parse_bte(word, table, PLCFBTE_CHPX)?;
    let mut output = Vec::new();
    for (index, pn) in pages.iter().enumerate() {
        let offset = (*pn as usize).checked_mul(512).ok_or_else(|| corrupted("CHPX page offset overflow"))?;
        let page = word.get(offset..offset + 512).ok_or_else(|| corrupted("CHPX page exceeds WordDocument"))?;
        let fkp = ChpxFkp::parse(page, word).ok_or_else(|| corrupted("malformed CHPX FKP"))?;
        for entry_index in 0..fkp.count() {
            let entry = fkp.entry(entry_index).ok_or_else(|| corrupted("PAPX entry is missing"))?;
            if entry.fc < fcs[index] || entry.end_fc > fcs[index + 1] { return Err(corrupted("CHPX FKP overlaps its BTE range")); }
            strict_sprms(&entry.grpprl)?;
            output.push(FcRun { start: entry.fc, end: entry.end_fc, grpprl: entry.grpprl.clone() });
        }
    }
    output.sort_by_key(|r| r.start);
    if output.windows(2).any(|v| v[0].end > v[1].start) { return Err(corrupted("CHPX runs overlap")); }
    Ok(output)
}

fn parse_papx(word: &[u8], table: &[u8]) -> Result<Vec<PapxRun>> {
    let (fcs, pages) = parse_bte(word, table, PLCFBTE_PAPX)?;
    let mut output = Vec::new();
    for (index, pn) in pages.iter().enumerate() {
        let offset = (*pn as usize).checked_mul(512).ok_or_else(|| corrupted("PAPX page offset overflow"))?;
        let page = word.get(offset..offset + 512).ok_or_else(|| corrupted("PAPX page exceeds WordDocument"))?;
        let fkp = PapxFkp::parse(page, word).ok_or_else(|| corrupted("malformed PAPX FKP"))?;
        for entry_index in 0..fkp.count() {
            let entry = fkp.entry(entry_index).ok_or_else(|| corrupted("PAPX entry is missing"))?;
            if entry.fc < fcs[index] || entry.end_fc > fcs[index + 1] { return Err(corrupted("PAPX FKP overlaps its BTE range")); }
            let grpprl = entry.grpprl.clone();
            if grpprl.len() < 2 { return Err(corrupted("PAPX style index is missing")); }
            strict_sprms(&grpprl[2..])?;
            output.push(PapxRun { start: entry.fc, end: entry.end_fc, grpprl,
                phe: entry.paragraph_height.ok_or_else(|| corrupted("PAPX PHE is missing"))? });
        }
    }
    output.sort_by_key(|r| r.start);
    if output.windows(2).any(|v| v[0].end > v[1].start) { return Err(corrupted("PAPX runs overlap")); }
    Ok(output)
}

struct BuiltPapxPage { start: u32, end: u32, bytes: Vec<u8> }

fn build_papx_pages(runs: &[PapxRun]) -> Result<Vec<BuiltPapxPage>> {
    let mut pages = Vec::new();
    let mut start = 0;
    while start < runs.len() {
        let mut count = 0;
        let mut used = 1usize;
        for run in &runs[start..] {
            let data = papx_storage_size(run.grpprl.len())?;
            let front = (count + 2) * 4 + (count + 1) * 13;
            if front + used + data > 512 { break; }
            count += 1; used += data;
        }
        if count == 0 { return Err(corrupted("one PAPX run cannot fit in an FKP")); }
        let subset = &runs[start..start + count];
        pages.push(BuiltPapxPage { start: subset[0].start, end: subset.last().unwrap().end,
            bytes: build_papx_page(subset)? });
        start += count;
    }
    Ok(pages)
}

fn papx_storage_size(len: usize) -> Result<usize> {
    if len == 0 || len > 510 { return Err(corrupted("PAPX length is invalid")); }
    let prefix = if len % 2 == 0 { 2 } else { 1 };
    let raw = prefix + len;
    Ok(raw + raw % 2)
}

fn build_papx_page(runs: &[PapxRun]) -> Result<Vec<u8>> {
    let mut page = vec![0u8; 512];
    let n = runs.len();
    for (i, run) in runs.iter().enumerate() { page[i * 4..i * 4 + 4].copy_from_slice(&run.start.to_le_bytes()); }
    page[n * 4..n * 4 + 4].copy_from_slice(&runs[n - 1].end.to_le_bytes());
    page[511] = u8::try_from(n).map_err(|_| corrupted("too many PAPX entries"))?;
    let bx = (n + 1) * 4;
    let mut cursor = 511usize;
    for (i, run) in runs.iter().enumerate().rev() {
        let prefix = if run.grpprl.len() % 2 == 0 { 2 } else { 1 };
        cursor = cursor.checked_sub(prefix + run.grpprl.len()).ok_or_else(|| corrupted("PAPX page overflow"))?;
        cursor &= !1;
        page[bx + i * 13] = u8::try_from(cursor / 2).map_err(|_| corrupted("PAPX offset exceeds byte"))?;
        write_phe(&mut page[bx + i * 13 + 1..bx + i * 13 + 13], run.phe);
        if prefix == 1 { page[cursor] = ((run.grpprl.len() + 1) / 2) as u8; }
        else { page[cursor] = 0; page[cursor + 1] = (run.grpprl.len() / 2) as u8; }
        page[cursor + prefix..cursor + prefix + run.grpprl.len()].copy_from_slice(&run.grpprl);
    }
    Ok(page)
}

fn write_phe(slot: &mut [u8], value: ParagraphHeight) {
    slot[0..2].copy_from_slice(&value.info_field.to_le_bytes()); slot[2..4].copy_from_slice(&value.reserved.to_le_bytes());
    slot[4..8].copy_from_slice(&value.dxa_col.to_le_bytes()); slot[8..12].copy_from_slice(&value.dym_line_or_height.to_le_bytes());
}

fn parse_bte(word: &[u8], table: &[u8], index: usize) -> Result<(Vec<u32>, Vec<u32>)> {
    let (offset, length) = fib_pair(word, index)?;
    let data = slice(table, offset, length, "bin table")?;
    if data.len() < 12 || (data.len() - 4) % 8 != 0 { return Err(corrupted("bin table shape is invalid")); }
    let n = (data.len() - 4) / 8;
    let mut fcs = Vec::with_capacity(n + 1); let mut pages = Vec::with_capacity(n);
    for i in 0..=n { fcs.push(u32_at(data, i * 4)?); }
    for i in 0..n { pages.push(u32_at(data, (n + 1) * 4 + i * 4)? & 0x003F_FFFF); }
    if fcs.windows(2).any(|v| v[0] >= v[1]) || pages.iter().any(|v| *v == 0) { return Err(corrupted("bin table ranges are invalid")); }
    Ok((fcs, pages))
}

fn parse_clx(word: &[u8], table: &[u8]) -> Result<Vec<RawPiece>> {
    let (offset, length) = fib_pair(word, CLX)?; let data = slice(table, offset, length, "CLX")?;
    if data.first() != Some(&2) { return Err(corrupted("fast-save RgPrc CLX is unsupported")); }
    let size = u32_at(data, 1)? as usize;
    if size + 5 != data.len() || size < 4 || (size - 4) % 12 != 0 { return Err(corrupted("CLX PlcPcd shape is invalid")); }
    let n = (size - 4) / 12;
    if n == 0 || n > MAX_PIECES { return Err(corrupted("piece count exceeds limits")); }
    let cp_base = 5; let pcd_base = cp_base + (n + 1) * 4; let mut output = Vec::with_capacity(n);
    for i in 0..n {
        let start = u32_at(data, cp_base + i * 4)?; let end = u32_at(data, cp_base + (i + 1) * 4)?;
        if start >= end || output.last().is_some_and(|p: &RawPiece| p.end != start) { return Err(corrupted("piece CPs overlap or contain gaps")); }
        let p = &data[pcd_base + i * 8..pcd_base + i * 8 + 8]; let raw = u32_at(p, 2)?;
        if raw & 0x8000_0000 != 0 { return Err(corrupted("piece FC reserved bit is set")); }
        let unicode = raw & 0x4000_0000 == 0; let fc = if unicode { raw & 0x3FFF_FFFF } else { (raw & 0x3FFF_FFFF) / 2 };
        let byte_len = (end - start).checked_mul(if unicode { 2 } else { 1 }).ok_or_else(|| corrupted("piece byte length overflow"))?;
        if fc.checked_add(byte_len).is_none_or(|v| v as usize > word.len()) { return Err(corrupted("piece text exceeds WordDocument")); }
        if p[6] != 0 || p[7] != 0 { return Err(corrupted("piece PRMs are unsupported for safe mutation")); }
        output.push(RawPiece { start, end, fc, unicode, prefix: [p[0], p[1]], prm: [p[6], p[7]] });
    }
    Ok(output)
}

fn serialize_clx(pieces: &[RawPiece]) -> Result<Vec<u8>> {
    if pieces.is_empty() || pieces.len() > MAX_PIECES { return Err(corrupted("piece count is invalid")); }
    let plc = pieces.len().checked_mul(12).and_then(|v| v.checked_add(4)).ok_or_else(|| corrupted("PlcPcd size overflow"))?;
    let mut out = vec![2]; out.extend_from_slice(&(plc as u32).to_le_bytes());
    for p in pieces { out.extend_from_slice(&p.start.to_le_bytes()); } out.extend_from_slice(&pieces.last().unwrap().end.to_le_bytes());
    for p in pieces { out.extend_from_slice(&p.prefix); let raw = if p.unicode { p.fc } else { p.fc.checked_mul(2).ok_or_else(|| corrupted("compressed FC overflow"))? | 0x4000_0000 }; out.extend_from_slice(&raw.to_le_bytes()); out.extend_from_slice(&p.prm); }
    Ok(out)
}

fn insert_piece(pieces: &mut Vec<RawPiece>, cp: u32, length: u32, fc: u32) -> Result<()> {
    let index = pieces.iter().position(|p| cp >= p.start && cp <= p.end).ok_or_else(|| corrupted("insertion CP has no piece"))?;
    for piece in pieces.iter_mut().skip(index + 1) {
        piece.start = piece.start.checked_add(length).ok_or_else(|| corrupted("piece CP overflow"))?;
        piece.end = piece.end.checked_add(length).ok_or_else(|| corrupted("piece CP overflow"))?;
    }
    let original = pieces[index].clone(); let mut replacement = Vec::new();
    if cp > original.start { let mut left = original.clone(); left.end = cp; replacement.push(left); }
    let inserted_end = cp.checked_add(length).ok_or_else(|| corrupted("piece CP overflow"))?;
    replacement.push(RawPiece { start: cp, end: inserted_end, fc, unicode: true, prefix: [0, 0], prm: [0, 0] });
    if cp < original.end { let mut right = original; let width = if right.unicode { 2 } else { 1 }; right.fc += (cp - right.start) * width; right.start = cp + length; right.end += length; replacement.push(right); }
    pieces.splice(index..=index, replacement);
    normalize_piece_cps(pieces)
}

fn delete_piece_range(pieces: &mut Vec<RawPiece>, start: u32, end: u32) -> Result<()> {
    let mut output = Vec::new(); let removed = end - start;
    for piece in pieces.iter() {
        if piece.end <= start { output.push(piece.clone()); continue; }
        if piece.start >= end { let mut p = piece.clone(); p.start -= removed; p.end -= removed; output.push(p); continue; }
        if piece.start < start { let mut left = piece.clone(); left.end = start; output.push(left); }
        if piece.end > end { let mut right = piece.clone(); let width = if right.unicode { 2 } else { 1 }; right.fc += (end - right.start) * width; right.start = start; right.end = piece.end - removed; output.push(right); }
    }
    if output.is_empty() { return Err(corrupted("cannot delete every document piece")); }
    *pieces = output; normalize_piece_cps(pieces)
}

fn normalize_piece_cps(pieces: &[RawPiece]) -> Result<()> {
    if pieces.windows(2).any(|v| v[0].end != v[1].start) || pieces.iter().any(|p| p.start >= p.end) { return Err(corrupted("piece rewrite produced gaps or overlaps")); }
    Ok(())
}

fn parse_authors(word: &[u8], table: &[u8]) -> Result<Vec<String>> {
    let (offset, length) = fib_pair(word, STTBFRMARK)?;
    if length == 0 { return Ok(vec!["Unknown".to_string()]); }
    let data = slice(table, offset, length, "SttbfRMark")?;
    if data.len() < 6 || u16_at(data, 0)? != 0xFFFF || u16_at(data, 4)? != 0 { return Err(corrupted("SttbfRMark header is invalid")); }
    let count = u16_at(data, 2)? as usize; let mut cursor = 6; let mut output = Vec::with_capacity(count);
    for _ in 0..count { let n = u16_at(data, cursor)? as usize; cursor += 2; let bytes = data.get(cursor..cursor + n * 2).ok_or_else(|| corrupted("revision author is truncated"))?; let units = bytes.chunks_exact(2).map(|v| u16::from_le_bytes([v[0],v[1]])).collect::<Vec<_>>(); output.push(String::from_utf16(&units).map_err(|_| corrupted("revision author is invalid UTF-16"))?); cursor += n * 2; }
    if cursor != data.len() || output.first().map(String::as_str) != Some("Unknown") { return Err(corrupted("SttbfRMark must begin with Unknown and have no trailing bytes")); }
    Ok(output)
}

fn serialize_authors(authors: &[String]) -> Result<Vec<u8>> {
    if authors.first().map(String::as_str) != Some("Unknown") || authors.len() > MAX_AUTHORS { return Err(corrupted("revision author table is invalid")); }
    let mut out = Vec::new(); out.extend_from_slice(&0xFFFFu16.to_le_bytes()); out.extend_from_slice(&(authors.len() as u16).to_le_bytes()); out.extend_from_slice(&0u16.to_le_bytes());
    for author in authors { let units = author.encode_utf16().collect::<Vec<_>>(); let n = u16::try_from(units.len()).map_err(|_| corrupted("revision author is too long"))?; out.extend_from_slice(&n.to_le_bytes()); for unit in units { out.extend_from_slice(&unit.to_le_bytes()); } }
    Ok(out)
}

fn parse_cp_table(word: &[u8], table: &[u8], index: usize, record_size: usize) -> Result<Option<CpTable>> {
    let (offset, length) = fib_pair(word, index)?; if length == 0 { return Ok(None); }
    let data = slice(table, offset, length, "CP-indexed PLCF")?;
    if data.len() < 4 || (data.len() - 4) % (4 + record_size) != 0 { return Err(corrupted("CP-indexed PLCF shape is invalid")); }
    let n = (data.len() - 4) / (4 + record_size); let mut cps = Vec::with_capacity(n + 1);
    for i in 0..=n { cps.push(u32_at(data, i * 4)?); }
    if cps.windows(2).any(|v| v[0] > v[1]) { return Err(corrupted("CP-indexed PLCF overlaps")); }
    Ok(Some(CpTable { index, cps, records: data[(n + 1) * 4..].to_vec() }))
}

fn reject_protection(word: &[u8], table: &[u8]) -> Result<()> {
    let (offset, length) = fib_pair(word, DOP)?; if length == 0 { return Ok(()); }
    let dop = slice(table, offset, length, "DOP")?;
    if dop.len() < 84 { return Err(corrupted("DOP is truncated")); }
    let protected = dop[6] & 0x10 != 0 || dop[7] & (0x02 | 0x20 | 0x40) != 0 || i32::from_le_bytes(dop[78..82].try_into().unwrap()) != 0;
    if protected { return Err(corrupted("protected DOC cannot be edited")); }
    Ok(())
}

fn read_units(word: &[u8], pieces: &[RawPiece], start: u32, end: u32) -> Result<Vec<u16>> {
    let mut out = Vec::new();
    for piece in pieces { let a = start.max(piece.start); let b = end.min(piece.end); if a >= b { continue; } let relative = a - piece.start; let count = b - a; if piece.unicode { let offset = (piece.fc + relative * 2) as usize; let bytes = word.get(offset..offset + count as usize * 2).ok_or_else(|| corrupted("text range exceeds WordDocument"))?; out.extend(bytes.chunks_exact(2).map(|v| u16::from_le_bytes([v[0],v[1]]))); } else { let offset = (piece.fc + relative) as usize; let bytes = word.get(offset..offset + count as usize).ok_or_else(|| corrupted("text range exceeds WordDocument"))?; out.extend(bytes.iter().map(|v| *v as u16)); } }
    Ok(out)
}

fn infer_moves(revisions: &mut [DocTrackedRevision]) {
    let mut groups: HashMap<u32, (Vec<usize>, Vec<usize>)> = HashMap::new();
    for (i, revision) in revisions.iter().enumerate() { if let Some(rsid) = revision.revision_save_id { let group = groups.entry(rsid).or_default(); match revision.kind { DocTrackedRevisionKind::Insertion => group.1.push(i), DocTrackedRevisionKind::Deletion => group.0.push(i), _ => {} } } }
    for (rsid, (from, to)) in groups { if from.len() == 1 && to.len() == 1 { revisions[from[0]].kind = DocTrackedRevisionKind::MoveFrom; revisions[to[0]].kind = DocTrackedRevisionKind::MoveTo; revisions[from[0]].move_pair_id = Some(rsid); revisions[to[0]].move_pair_id = Some(rsid); } }
}

fn merge_adjacent(output: &mut Vec<DocTrackedRevision>) {
    output.sort_by_key(|item| (kind_order(item.kind), item.author_index, item.reason,
        item.revision_save_id, item.start_cp, item.end_cp));
    let mut merged: Vec<DocTrackedRevision> = Vec::new();
    for item in output.drain(..) { if let Some(last) = merged.last_mut() { if last.end_cp == item.start_cp && last.kind == item.kind && last.author_index == item.author_index && last.timestamp == item.timestamp && last.reason == item.reason && last.revision_save_id == item.revision_save_id { last.end_cp = item.end_cp; continue; } } merged.push(item); }
    merged.sort_by_key(|item| (item.start_cp, item.end_cp, kind_order(item.kind)));
    *output = merged;
}

fn merge_fc_runs(runs: &mut Vec<FcRun>) { let mut out: Vec<FcRun> = Vec::new(); for run in runs.drain(..) { if let Some(last) = out.last_mut() { if last.end == run.start && last.grpprl == run.grpprl { last.end = run.end; continue; } } out.push(run); } *runs = out; }
fn kind_order(kind: DocTrackedRevisionKind) -> u8 { match kind { DocTrackedRevisionKind::Insertion|DocTrackedRevisionKind::MoveTo=>0, DocTrackedRevisionKind::Deletion|DocTrackedRevisionKind::MoveFrom=>1, DocTrackedRevisionKind::CharacterFormatting=>2, DocTrackedRevisionKind::ParagraphFormatting=>3, DocTrackedRevisionKind::TableRowFormatting=>4 } }
fn validate_range(start:u32,end:u32,limit:u32)->Result<()> { if start>=end || end>limit { Err(corrupted("tracked revision range is empty or exceeds the main story")) } else { Ok(()) } }
fn pack_dttm(value: Option<CommentDateTime>) -> Result<u32> { let Some(v)=value else{return Ok(0)}; if !(1900..=2411).contains(&v.year)||!(1..=12).contains(&v.month)||!(1..=31).contains(&v.day)||v.hour>23||v.minute>59||v.weekday>6{return Err(corrupted("revision timestamp is outside DTTM limits"));} Ok(u32::from(v.minute)|u32::from(v.hour)<<6|u32::from(v.day)<<11|u32::from(v.month)<<16|u32::from(v.year-1900)<<20|u32::from(v.weekday)<<29) }
fn decode_dttm(raw:u32)->Result<Option<CommentDateTime>> { if raw==0{return Ok(None)}; let value=CommentDateTime{minute:(raw&0x3f)as u8,hour:((raw>>6)&0x1f)as u8,day:((raw>>11)&0x1f)as u8,month:((raw>>16)&0xf)as u8,year:((raw>>20)&0x1ff)as u16+1900,weekday:((raw>>29)&7)as u8}; let _=pack_dttm(Some(value))?; Ok(Some(value)) }
fn push_byte(out:&mut Vec<u8>,op:u16,v:u8){out.extend_from_slice(&op.to_le_bytes());out.push(v)}
fn push_word(out:&mut Vec<u8>,op:u16,v:u16){out.extend_from_slice(&op.to_le_bytes());out.extend_from_slice(&v.to_le_bytes())}
fn push_dword(out:&mut Vec<u8>,op:u16,v:u32){out.extend_from_slice(&op.to_le_bytes());out.extend_from_slice(&v.to_le_bytes())}
fn append_table_block(word:&mut[u8],table:&mut Vec<u8>,index:usize,data:&[u8])->Result<()> { let offset=u32::try_from(table.len()).map_err(|_|corrupted("Table stream exceeds u32"))?; table.extend_from_slice(data); put_fib_pair(word,index,offset,u32::try_from(data.len()).map_err(|_|corrupted("table block exceeds u32"))?) }
fn fib_pair(word:&[u8],index:usize)->Result<(u32,u32)>{Ok((u32_at(word,FIB_FC_LCB+index*8)?,u32_at(word,FIB_FC_LCB+index*8+4)?))}
fn put_fib_pair(word:&mut[u8],index:usize,fc:u32,lcb:u32)->Result<()>{put_u32(word,FIB_FC_LCB+index*8,fc)?;put_u32(word,FIB_FC_LCB+index*8+4,lcb)}
fn slice<'a>(data:&'a[u8],offset:u32,length:u32,name:&str)->Result<&'a[u8]>{let start=offset as usize;let end=start.checked_add(length as usize).ok_or_else(||corrupted(format!("{name} range overflow")))?;data.get(start..end).ok_or_else(||corrupted(format!("{name} exceeds stream")))}
fn u16_at(data:&[u8],offset:usize)->Result<u16>{data.get(offset..offset+2).map(|v|u16::from_le_bytes(v.try_into().unwrap())).ok_or_else(||corrupted("truncated u16"))}
fn u32_at(data:&[u8],offset:usize)->Result<u32>{data.get(offset..offset+4).map(|v|u32::from_le_bytes(v.try_into().unwrap())).ok_or_else(||corrupted("truncated u32"))}
fn put_u32(data:&mut[u8],offset:usize,value:u32)->Result<()>{data.get_mut(offset..offset+4).ok_or_else(||corrupted("truncated FIB field"))?.copy_from_slice(&value.to_le_bytes());Ok(())}
fn align2(v:usize)->Result<usize>{v.checked_add(1).map(|n|n&!1).ok_or_else(||corrupted("alignment overflow"))}
fn align512(v:usize)->Result<usize>{v.checked_add(511).map(|n|n&!511).ok_or_else(||corrupted("alignment overflow"))}
fn corrupted(message:impl Into<String>)->DocError{DocError::Corrupted(message.into())}
