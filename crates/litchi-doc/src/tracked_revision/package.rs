//! Package-layer transactional editor for DOC tracked revisions.

use super::Limits;
use super::codec::{
    CLX, DOP, FIB_CCP_TEXT, FIB_FC_LCB, MAX_AUTHORS, MAX_REVISIONS, MAX_TEXT_UNITS, PLCFANDREF,
    PLCFATNBKF, PLCFATNBKL, PLCFBKF, PLCFBKL, PLCFBTE_CHPX, PLCFBTE_PAPX, PLCFFLD_MOM,
    ParsedMetadata, STTBFRMARK, align2, align512, append_table_block, build_papx_pages, corrupted,
    delete_piece_range, encode_revision, fib_pair, infer_moves, insert_piece, kind_order,
    merge_adjacent, metadata_from_sprms, parse_authors, parse_chpx, parse_clx, parse_cp_table,
    parse_papx, property_metadata, put_u32, read_units, reject_protection,
    replace_papx_revision_sprms, replace_revision_sprms, restore_before_wall, retain_sprms,
    revision_opcodes, serialize_authors, serialize_clx, slice, split_transform_chpx,
    split_transform_papx, strict_sprms, u16_at, u32_at, validate_metadata, validate_range,
};
use super::model::{CpTable, FcRun, PapxRun, RawPiece, Revision, RevisionKind, RevisionMetadata};
use crate::package::{Error as PackageError, Result};
use crate::sprm_operations::{
    SPRM_C_DTTM_RMARK, SPRM_C_DTTM_RMARK_DEL, SPRM_C_F_BOLD, SPRM_C_F_RMARK, SPRM_C_F_RMARK_DEL,
    SPRM_C_IBST_RMARK, SPRM_C_IBST_RMARK_DEL, SPRM_C_IDSL_RMARK, SPRM_C_IDSL_RMARK_DEL,
    SPRM_C_PROP_RMARK_CURRENT, SPRM_C_PROP_RMARK90, SPRM_C_RSID_PROP, SPRM_C_RSID_RM_DEL,
    SPRM_C_RSID_TEXT, SPRM_C_WALL, SPRM_P_F_IN_TABLE, SPRM_P_PROP_RMARK, SPRM_P_PROP_RMARK_CURRENT,
    SPRM_P_PROP_RMARK90, SPRM_P_WALL, SPRM_T_PROP_RMARK, SPRM_T_RSID, SPRM_T_WALL,
};
use crate::writer::ChpxFkpBuilder;
use litchi_ole_common::object::{Editor as ObjectEditor, Targets};

#[derive(Clone)]
pub struct RevisionEditor {
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
    unmodeled_cp_tables: Vec<usize>,
    main_ccp: u32,
    changed: bool,
}

impl RevisionEditor {
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
        // Each modeled entry is a PLCF in the main/all-story CP coordinate
        // space. Fixed-size records stay opaque while their CP array moves.
        let modeled_cp_tables = [
            (2, 2),           // PlcffndRef / FRD
            (PLCFANDREF, 30), // PlcfandRef / ATRD
            (6, 12),          // PlcfSed / SED
            (PLCFFLD_MOM, 2), // PlcffldMom / FLD
            (PLCFBKF, 4),     // PlcfBkf / FBKF
            (PLCFBKL, 0),     // PlcfBkl
            (40, 26),         // PlcfSpaMom / SPA
            (PLCFATNBKF, 4),  // PlcfAtnBkf / FBKF
            (PLCFATNBKL, 0),  // PlcfAtnBkl
            (46, 2),          // PlcfendRef / FRD
            (54, 12),         // PlcfWKB / WKB
            (55, 2),          // PlcfSpl / SPLS
            (89, 4),          // PlcfAsumy / ASUMY
            (90, 2),          // PlcfGram / SPLS
            (93, 4),          // PlcfTch / TCH
            (98, 2),          // PlcfLvc / LSPD
            (115, 6),         // PlcfBkfFactoid / FBKFD
            (117, 4),         // PlcfBklFactoid / FBKLD
            (121, 6),         // PlcfBkfFcc / FBKFD
            (122, 4),         // PlcfBklFcc / FBKLD
            (124, 4),         // PlcfBkfBPRepairs / FBKF
            (125, 0),         // PlcfBklBPRepairs
            (132, 2),         // Plcffactoid / FactoidSpls
            (138, 6),         // PlcfBkfSdt / FBKFD
            (139, 4),         // PlcfBklSdt / FBKLD
            (142, 6),         // PlcfBkfProt / BKF
            (143, 0),         // PlcfBklProt
        ];
        let pair_count = usize::from(u16_at(&word, FIB_FC_LCB - 2)?);
        let cp_tables = modeled_cp_tables
            .into_iter()
            .filter(|(index, _size)| *index < pair_count)
            .filter_map(|(index, size)| parse_cp_table(&word, &table, index, size).transpose())
            .collect::<Result<Vec<_>>>()?;
        // Known CP-indexed tables whose coupled records are not owned here.
        // Length-changing edits refuse these instead of silently leaving stale
        // positions. Equal-length text and formatting edits remain safe.
        // PlcfSea has producer-private records. The cookie and UIM records
        // carry coupled character lengths that a CP-only splice cannot repair.
        let unmodeled_cp_tables = [14, 101, 110, 116]
            .into_iter()
            .filter(|index| *index < pair_count)
            .filter_map(|index| {
                fib_pair(&word, index)
                    .map(|(_offset, length)| (length != 0).then_some(index))
                    .transpose()
            })
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
            unmodeled_cp_tables,
            main_ccp,
            changed: false,
        };
        if editor.revisions()?.len() > MAX_REVISIONS {
            return Err(corrupted("revision count exceeds resource limit"));
        }
        Ok(editor)
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub fn authors(&self) -> &[String] {
        &self.authors
    }

    /// Returns the exact main-story text after strict piece-table decoding.
    ///
    /// This deliberately differs from the permissive reader projection: a
    /// transaction must not treat malformed UTF-16 or a discontinuous piece
    /// table as editable text.
    pub(crate) fn main_story_text(&self) -> Result<String> {
        let mut output = String::new();
        let mut covered = 0u32;
        for piece in &self.pieces {
            if piece.start >= self.main_ccp {
                break;
            }
            let end = piece.end.min(self.main_ccp);
            if end <= piece.start {
                continue;
            }
            let count = end - piece.start;
            if piece.unicode {
                let offset = usize::try_from(piece.fc)
                    .map_err(|_| corrupted("Unicode piece offset exceeds usize"))?;
                let byte_count = usize::try_from(count)
                    .ok()
                    .and_then(|value| value.checked_mul(2))
                    .ok_or_else(|| corrupted("Unicode piece byte count overflow"))?;
                let bytes = self
                    .word
                    .get(offset..offset + byte_count)
                    .ok_or_else(|| corrupted("Unicode piece exceeds WordDocument"))?;
                let units = bytes
                    .chunks_exact(2)
                    .map(|value| u16::from_le_bytes([value[0], value[1]]))
                    .collect::<Vec<_>>();
                output.push_str(
                    &String::from_utf16(&units)
                        .map_err(|_| corrupted("main-story text contains invalid UTF-16"))?,
                );
            } else {
                let offset = usize::try_from(piece.fc)
                    .map_err(|_| corrupted("compressed piece offset exceeds usize"))?;
                let byte_count = usize::try_from(count)
                    .map_err(|_| corrupted("compressed piece byte count exceeds usize"))?;
                let bytes = self
                    .word
                    .get(offset..offset + byte_count)
                    .ok_or_else(|| corrupted("compressed piece exceeds WordDocument"))?;
                let (decoded, _, had_errors) = encoding_rs::WINDOWS_1252.decode(bytes);
                if had_errors || decoded.encode_utf16().count() != bytes.len() {
                    return Err(corrupted("compressed piece cannot be decoded losslessly"));
                }
                output.push_str(&decoded);
            }
            covered = covered
                .checked_add(count)
                .ok_or_else(|| corrupted("main-story CP count overflow"))?;
        }
        if covered != self.main_ccp || output.encode_utf16().count() != self.main_ccp as usize {
            return Err(corrupted(
                "piece table does not exactly cover the main story",
            ));
        }
        Ok(output)
    }

    /// Known CP-indexed FIB tables outside the length-changing splice model.
    #[must_use]
    pub(crate) fn unmodeled_length_dependencies(&self) -> &[usize] {
        &self.unmodeled_cp_tables
    }

    /// Whether all character runs in a non-empty range have byte-identical
    /// direct formatting. Length-changing replacement can preserve only this
    /// unambiguous dependency closure.
    pub(crate) fn has_uniform_character_format(&self, start: u32, end: u32) -> Result<bool> {
        let groups = self.character_groups(start, end)?;
        Ok(groups.windows(2).all(|pair| pair[0] == pair[1]))
    }

    /// Returns a uniform direct bold override, or `None` when the selected
    /// runs disagree or use a non-literal toggle value.
    pub(crate) fn uniform_bold_override(
        &self,
        start: u32,
        end: u32,
    ) -> Result<Option<Option<bool>>> {
        let groups = self.character_groups(start, end)?;
        let mut uniform = None;
        for group in groups {
            let value = strict_sprms(group)?
                .iter()
                .rev()
                .find(|sprm| sprm.opcode == SPRM_C_F_BOLD)
                .and_then(super::super::sprm::Sprm::operand_byte);
            let value = match value {
                Some(0) => Some(false),
                Some(1) => Some(true),
                Some(_) => return Ok(None),
                None => None,
            };
            match uniform {
                Some(previous) if previous != value => return Ok(None),
                Some(_) => {},
                None => uniform = Some(value),
            }
        }
        Ok(uniform)
    }

    /// Replaces a non-empty main-story range by appending one Unicode piece,
    /// shifting modeled CP tables, rebuilding CHPX FKPs, and publishing a new
    /// CLX. Callers must first prove a uniform character-format closure.
    pub(crate) fn replace_plain_text(
        &mut self,
        start: u32,
        end: u32,
        replacement: &str,
    ) -> Result<()> {
        validate_range(start, end, self.main_ccp)?;
        self.reject_destructive_interactions(start, end)?;
        if !self.has_uniform_character_format(start, end)? {
            return Err(corrupted(
                "length-changing body replacement crosses character formatting runs",
            ));
        }
        let groups = self.character_groups(start, end)?;
        let formatting = groups
            .first()
            .ok_or_else(|| corrupted("body replacement has no character formatting"))?
            .to_vec();
        let units = replacement.encode_utf16().collect::<Vec<_>>();
        if units.len() > MAX_TEXT_UNITS {
            return Err(corrupted("body replacement exceeds text resource limit"));
        }
        if units.len() != (end - start) as usize && !self.unmodeled_cp_tables.is_empty() {
            return Err(corrupted(
                "length-changing body replacement has unmodeled CP-indexed dependencies",
            ));
        }

        let mut candidate = self.clone();
        let removed = end - start;
        delete_piece_range(&mut candidate.pieces, start, end)?;
        let added = u32::try_from(units.len())
            .map_err(|_error| corrupted("body replacement length exceeds u32"))?;
        if added != 0 {
            let fc = align2(candidate.word.len())?;
            candidate.word.resize(fc, 0);
            for unit in units {
                candidate.word.extend_from_slice(&unit.to_le_bytes());
            }
            insert_piece(
                &mut candidate.pieces,
                start,
                added,
                u32::try_from(fc).map_err(|_error| corrupted("body replacement FC exceeds u32"))?,
            )?;
            candidate.chpx.push(FcRun {
                start: u32::try_from(fc)
                    .map_err(|_error| corrupted("body replacement FC exceeds u32"))?,
                end: u32::try_from(candidate.word.len())
                    .map_err(|_error| corrupted("body replacement FC exceeds u32"))?,
                grpprl: formatting,
            });
        }
        candidate.shift_cp_tables(start, removed, added)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(removed)
            .and_then(|value| value.checked_add(added))
            .ok_or_else(|| corrupted("main story CP replacement overflow"))?;
        candidate.rewrite_chpx()?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Sets or clears one direct bold override while retaining every other
    /// character SPRM and rebuilding the affected CHPX FKPs.
    pub(crate) fn set_character_bold_override(
        &mut self,
        start: u32,
        end: u32,
        value: Option<bool>,
    ) -> Result<()> {
        validate_range(start, end, self.main_ccp)?;
        let intervals = self.fc_intervals(start, end)?;
        let mut candidate = self.clone();
        split_transform_chpx(&mut candidate.chpx, &intervals, |group| {
            let mut output = retain_sprms(group, &[SPRM_C_F_BOLD])?;
            if let Some(enabled) = value {
                output.extend_from_slice(&SPRM_C_F_BOLD.to_le_bytes());
                output.push(u8::from(enabled));
            }
            if output.len() > 255 {
                return Err(corrupted("edited CHPX exceeds one-byte FKP limit"));
            }
            Ok(output)
        })?;
        candidate.rewrite_chpx()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Whether the paragraph ending at `cp` has the MS-DOC in-table flag.
    pub(crate) fn is_in_table_at_cp(&self, cp: u32) -> Result<bool> {
        let piece = self
            .pieces
            .iter()
            .find(|piece| piece.start <= cp && cp < piece.end)
            .ok_or_else(|| corrupted("paragraph terminator has no text piece"))?;
        let width = if piece.unicode { 2 } else { 1 };
        let fc = piece
            .fc
            .checked_add(
                cp.checked_sub(piece.start)
                    .ok_or_else(|| corrupted("paragraph CP underflow"))?
                    .checked_mul(width)
                    .ok_or_else(|| corrupted("paragraph FC overflow"))?,
            )
            .ok_or_else(|| corrupted("paragraph FC overflow"))?;
        let run = self
            .papx
            .iter()
            .find(|run| run.start <= fc && fc < run.end)
            .ok_or_else(|| corrupted("paragraph terminator has no PAPX run"))?;
        let body = run
            .grpprl
            .get(2..)
            .ok_or_else(|| corrupted("PAPX has no style index"))?;
        Ok(strict_sprms(body)?
            .iter()
            .rev()
            .find(|sprm| sprm.opcode == SPRM_P_F_IN_TABLE)
            .is_some_and(|sprm| sprm.operand_byte() == Some(1)))
    }

    /// Main-story length in MS-DOC CP (UTF-16 code-unit) coordinates.
    #[must_use]
    pub(crate) const fn main_story_cp_len(&self) -> u32 {
        self.main_ccp
    }

    fn character_groups(&self, start: u32, end: u32) -> Result<Vec<&[u8]>> {
        let intervals = self.fc_intervals(start, end)?;
        let mut output = Vec::new();
        for (interval_start, interval_end) in intervals {
            let mut cursor = interval_start;
            for run in &self.chpx {
                let left = interval_start.max(run.start);
                let right = interval_end.min(run.end);
                if left >= right {
                    continue;
                }
                if left > cursor {
                    return Err(corrupted("CHPX formatting has a physical FC gap"));
                }
                output.push(run.grpprl.as_slice());
                cursor = cursor.max(right);
            }
            if cursor < interval_end {
                return Err(corrupted("CHPX formatting does not cover body text"));
            }
        }
        if output.is_empty() {
            return Err(corrupted("body text has no CHPX formatting"));
        }
        Ok(output)
    }

    /// Lists character and PAPX property revisions, merging adjacent runs with
    /// identical metadata even when the range crosses piece boundaries.
    pub fn revisions(&self) -> Result<Vec<Revision>> {
        let mut output = Vec::new();
        for run in &self.chpx {
            let sprms = strict_sprms(&run.grpprl)?;
            for (kind, flag, author_op, time_op, reason_op, rsid_op) in [
                (
                    RevisionKind::Insertion,
                    SPRM_C_F_RMARK,
                    SPRM_C_IBST_RMARK,
                    SPRM_C_DTTM_RMARK,
                    SPRM_C_IDSL_RMARK,
                    SPRM_C_RSID_TEXT,
                ),
                (
                    RevisionKind::Deletion,
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
                            RevisionKind::CharacterFormatting,
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
                    RevisionKind::ParagraphFormatting,
                    [
                        SPRM_P_PROP_RMARK,
                        SPRM_P_PROP_RMARK90,
                        SPRM_P_PROP_RMARK_CURRENT,
                    ]
                    .as_slice(),
                    None,
                ),
                (
                    RevisionKind::TableRowFormatting,
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
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        if !matches!(kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
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
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
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
    pub fn update(&mut self, index: usize, metadata: RevisionMetadata) -> Result<Revision> {
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
    pub fn remove(&mut self, index: usize) -> Result<Revision> {
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
    pub fn accept(&mut self, index: usize) -> Result<Revision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(
            revision.kind,
            RevisionKind::Deletion | RevisionKind::MoveFrom
        ) {
            self.delete_revision_text(&revision)?;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    /// Rejects a revision using Word redline semantics.
    pub fn reject(&mut self, index: usize) -> Result<Revision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(
            revision.kind,
            RevisionKind::Insertion | RevisionKind::MoveTo
        ) {
            self.delete_revision_text(&revision)?;
        } else if matches!(
            revision.kind,
            RevisionKind::CharacterFormatting
                | RevisionKind::ParagraphFormatting
                | RevisionKind::TableRowFormatting
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

    fn delete_revision_text(&mut self, revision: &Revision) -> Result<()> {
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
        kind: RevisionKind,
        replacement: Option<(u16, &RevisionMetadata)>,
    ) -> Result<()> {
        let intervals = self.fc_intervals(start, end)?;
        match kind {
            RevisionKind::Insertion
            | RevisionKind::Deletion
            | RevisionKind::MoveFrom
            | RevisionKind::MoveTo
            | RevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    replace_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_chpx()?;
            },
            RevisionKind::ParagraphFormatting | RevisionKind::TableRowFormatting => {
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

    fn reject_formatting_revision(&mut self, revision: &Revision) -> Result<()> {
        let intervals = self.fc_intervals(revision.start_cp, revision.end_cp)?;
        match revision.kind {
            RevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    restore_before_wall(
                        grp,
                        SPRM_C_WALL,
                        &revision_opcodes(RevisionKind::CharacterFormatting, true),
                    )
                })?;
                self.rewrite_chpx()
            },
            RevisionKind::ParagraphFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp
                        .get(..2)
                        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(
                        &grp[2..],
                        SPRM_P_WALL,
                        &revision_opcodes(RevisionKind::ParagraphFormatting, true),
                    )?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            },
            RevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp
                        .get(..2)
                        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(
                        &grp[2..],
                        SPRM_T_WALL,
                        &revision_opcodes(RevisionKind::TableRowFormatting, true),
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
        output: &mut Vec<Revision>,
        fc_start: u32,
        fc_end: u32,
        kind: RevisionKind,
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

    fn find_exact(&self, start: u32, end: u32, kind: RevisionKind) -> Result<Revision> {
        self.revisions()?
            .into_iter()
            .find(|r| {
                r.start_cp == start
                    && r.end_cp == end
                    && (r.kind == kind
                        || matches!(
                            (r.kind, kind),
                            (RevisionKind::Insertion, RevisionKind::MoveTo)
                                | (RevisionKind::Deletion, RevisionKind::MoveFrom)
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
