//! MS-DOC codecs and bounded binary helpers for tracked revisions.

use super::model::{CpTable, FcRun, PapxRun, RawPiece, Revision, RevisionKind, RevisionMetadata};
use crate::DateTime;
use crate::package::{Error as PackageError, Result};
use crate::parts::fkp::{ChpxFkp, PapxFkp, ParagraphHeight};
use crate::sprm::parse_sprms;
use crate::sprm_operations::{
    SPRM_C_DTTM_RMARK, SPRM_C_DTTM_RMARK_DEL, SPRM_C_F_RMARK, SPRM_C_F_RMARK_DEL,
    SPRM_C_IBST_RMARK, SPRM_C_IBST_RMARK_DEL, SPRM_C_IDSL_RMARK, SPRM_C_IDSL_RMARK_DEL,
    SPRM_C_PROP_RMARK_CURRENT, SPRM_C_PROP_RMARK90, SPRM_C_RSID_PROP, SPRM_C_RSID_RM_DEL,
    SPRM_C_RSID_TEXT, SPRM_C_WALL, SPRM_P_F_IN_TABLE, SPRM_P_PROP_RMARK, SPRM_P_PROP_RMARK_CURRENT,
    SPRM_P_PROP_RMARK90, SPRM_P_WALL, SPRM_T_PROP_RMARK, SPRM_T_RSID, SPRM_T_WALL,
};
use std::collections::{BTreeSet, HashMap};

pub(super) const FIB_CCP_TEXT: usize = 76;
pub(super) const FIB_FC_LCB: usize = 154;
pub(super) const PLCFANDREF: usize = 4;
pub(super) const PLCFBTE_CHPX: usize = 12;
pub(super) const PLCFBTE_PAPX: usize = 13;
pub(super) const PLCFFLD_MOM: usize = 16;
pub(super) const PLCFBKF: usize = 22;
pub(super) const PLCFBKL: usize = 23;
pub(super) const DOP: usize = 31;
pub(super) const CLX: usize = 33;
pub(super) const PLCFATNBKF: usize = 42;
pub(super) const PLCFATNBKL: usize = 43;
pub(super) const STTBFRMARK: usize = 51;
pub(super) const MAX_PIECES: usize = 65_536;
pub(super) const MAX_REVISIONS: usize = 65_536;
pub(super) const MAX_AUTHORS: usize = i16::MAX as usize;
pub(super) const MAX_TEXT_UNITS: usize = 16 * 1024 * 1024;

#[derive(Clone)]
pub(super) struct ParsedMetadata {
    author_index: u16,
    author: String,
    timestamp: Option<DateTime>,
    reason: Option<u16>,
    rsid: Option<u32>,
}

impl ParsedMetadata {
    pub(super) fn to_revision(&self, kind: RevisionKind, start_cp: u32, end_cp: u32) -> Revision {
        Revision {
            kind,
            start_cp,
            end_cp,
            author_index: self.author_index,
            author: self.author.clone(),
            timestamp: self.timestamp,
            reason: self.reason,
            revision_save_id: self.rsid,
            move_pair_id: None,
        }
    }
}

pub(super) fn metadata_from_sprms(
    sprms: &[crate::sprm::Sprm],
    author_op: u16,
    time_op: u16,
    reason_op: u16,
    rsid_op: u16,
    authors: &[String],
) -> Result<ParsedMetadata> {
    let author_index = sprms
        .iter()
        .rev()
        .find(|s| s.opcode == author_op)
        .and_then(super::super::sprm::Sprm::operand_word)
        .unwrap_or(0);
    let author = authors
        .get(author_index as usize)
        .ok_or_else(|| corrupted("revision author index exceeds SttbfRMark"))?
        .clone();
    let timestamp = sprms
        .iter()
        .rev()
        .find(|s| s.opcode == time_op)
        .and_then(super::super::sprm::Sprm::operand_dword)
        .map(decode_dttm)
        .transpose()?
        .flatten();
    let reason = sprms
        .iter()
        .rev()
        .find(|s| s.opcode == reason_op)
        .and_then(super::super::sprm::Sprm::operand_word);
    let rsid = sprms
        .iter()
        .rev()
        .find(|s| s.opcode == rsid_op)
        .and_then(super::super::sprm::Sprm::operand_dword);
    Ok(ParsedMetadata {
        author_index,
        author,
        timestamp,
        reason,
        rsid,
    })
}

pub(super) fn property_metadata(
    operand: &[u8],
    sprms: &[crate::sprm::Sprm],
    rsid_op: u16,
    authors: &[String],
) -> Result<ParsedMetadata> {
    if operand.len() != 7 {
        return Err(corrupted("property revision operand must be seven bytes"));
    }
    let author_index = u16::from_le_bytes([operand[1], operand[2]]);
    let author = authors
        .get(author_index as usize)
        .ok_or_else(|| corrupted("property revision author exceeds SttbfRMark"))?
        .clone();
    let raw = u32::from_le_bytes(array_at(operand, 3, "property revision timestamp")?);
    let timestamp = decode_dttm(raw)?;
    let rsid = (rsid_op != 0)
        .then(|| {
            sprms
                .iter()
                .rev()
                .find(|s| s.opcode == rsid_op)
                .and_then(super::super::sprm::Sprm::operand_dword)
        })
        .flatten();
    Ok(ParsedMetadata {
        author_index,
        author,
        timestamp,
        reason: None,
        rsid,
    })
}

pub(super) fn validate_metadata(kind: RevisionKind, metadata: &RevisionMetadata) -> Result<()> {
    if metadata.author.is_empty() {
        return Err(corrupted("revision author must not be empty"));
    }
    if metadata.reason.is_some_and(|v| v > 0x2B) {
        return Err(corrupted("revision reason is undefined"));
    }
    if matches!(kind, RevisionKind::MoveFrom | RevisionKind::MoveTo)
        && metadata.revision_save_id.is_none()
    {
        return Err(corrupted(
            "move revisions require a shared revision_save_id",
        ));
    }
    if let Some(value) = metadata.timestamp {
        pack_dttm(Some(value))?;
    }
    Ok(())
}

pub(super) fn encode_revision(
    kind: RevisionKind,
    author: u16,
    metadata: &RevisionMetadata,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    match kind {
        RevisionKind::Insertion | RevisionKind::MoveTo => {
            push_byte(&mut output, SPRM_C_F_RMARK, 1);
            push_word(&mut output, SPRM_C_IBST_RMARK, author);
            if metadata.timestamp.is_some() {
                push_dword(
                    &mut output,
                    SPRM_C_DTTM_RMARK,
                    pack_dttm(metadata.timestamp)?,
                );
            }
            if let Some(v) = metadata.reason {
                push_word(&mut output, SPRM_C_IDSL_RMARK, v);
            }
            if let Some(v) = metadata.revision_save_id {
                push_dword(&mut output, SPRM_C_RSID_TEXT, v);
            }
        },
        RevisionKind::Deletion | RevisionKind::MoveFrom => {
            push_byte(&mut output, SPRM_C_F_RMARK_DEL, 1);
            push_word(&mut output, SPRM_C_IBST_RMARK_DEL, author);
            if metadata.timestamp.is_some() {
                push_dword(
                    &mut output,
                    SPRM_C_DTTM_RMARK_DEL,
                    pack_dttm(metadata.timestamp)?,
                );
            }
            if let Some(v) = metadata.reason {
                push_word(&mut output, SPRM_C_IDSL_RMARK_DEL, v);
            }
            if let Some(v) = metadata.revision_save_id {
                push_dword(&mut output, SPRM_C_RSID_RM_DEL, v);
            }
        },
        RevisionKind::CharacterFormatting => {
            output.extend_from_slice(&SPRM_C_PROP_RMARK_CURRENT.to_le_bytes());
            output.push(7);
            output.push(1);
            output.extend_from_slice(&author.to_le_bytes());
            output.extend_from_slice(&pack_dttm(metadata.timestamp)?.to_le_bytes());
            if let Some(v) = metadata.reason {
                push_word(&mut output, SPRM_C_IDSL_RMARK, v);
            }
            if let Some(v) = metadata.revision_save_id {
                push_dword(&mut output, SPRM_C_RSID_PROP, v);
            }
        },
        RevisionKind::ParagraphFormatting => {
            output.extend_from_slice(&SPRM_P_PROP_RMARK_CURRENT.to_le_bytes());
            output.push(7);
            output.push(1);
            output.extend_from_slice(&author.to_le_bytes());
            output.extend_from_slice(&pack_dttm(metadata.timestamp)?.to_le_bytes());
        },
        RevisionKind::TableRowFormatting => {
            output.extend_from_slice(&SPRM_T_PROP_RMARK.to_le_bytes());
            output.push(7);
            output.push(1);
            output.extend_from_slice(&author.to_le_bytes());
            output.extend_from_slice(&pack_dttm(metadata.timestamp)?.to_le_bytes());
            if let Some(v) = metadata.revision_save_id {
                push_dword(&mut output, SPRM_T_RSID, v);
            }
        },
    }
    Ok(output)
}

pub(super) fn replace_revision_sprms(
    grp: &[u8],
    kind: RevisionKind,
    replacement: Option<(u16, &RevisionMetadata)>,
) -> Result<Vec<u8>> {
    let remove = revision_opcodes(kind, replacement.is_none());
    let mut output = retain_sprms(grp, &remove)?;
    if let Some((author, metadata)) = replacement {
        output.extend_from_slice(&encode_revision(kind, author, metadata)?);
    }
    if output.len() > 255 {
        return Err(corrupted("edited CHPX exceeds one-byte FKP limit"));
    }
    Ok(output)
}

pub(super) fn replace_papx_revision_sprms(
    grp: &[u8],
    kind: RevisionKind,
    replacement: Option<(u16, &RevisionMetadata)>,
) -> Result<Vec<u8>> {
    let style = grp
        .get(..2)
        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
    let body = grp
        .get(2..)
        .ok_or_else(|| corrupted("PAPX body is truncated"))?;
    let mut output = style.to_vec();
    output.extend_from_slice(&retain_sprms(
        body,
        &revision_opcodes(kind, replacement.is_none()),
    )?);
    if let Some((author, metadata)) = replacement {
        output.extend_from_slice(&encode_revision(kind, author, metadata)?);
    }
    if output.len() > 510 {
        return Err(corrupted("edited PAPX exceeds FKP limit"));
    }
    Ok(output)
}

pub(super) fn revision_opcodes(kind: RevisionKind, remove_wall: bool) -> Vec<u16> {
    match kind {
        RevisionKind::Insertion | RevisionKind::MoveTo => vec![
            SPRM_C_F_RMARK,
            SPRM_C_IBST_RMARK,
            SPRM_C_DTTM_RMARK,
            SPRM_C_IDSL_RMARK,
            SPRM_C_RSID_TEXT,
        ],
        RevisionKind::Deletion | RevisionKind::MoveFrom => vec![
            SPRM_C_F_RMARK_DEL,
            SPRM_C_IBST_RMARK_DEL,
            SPRM_C_DTTM_RMARK_DEL,
            SPRM_C_IDSL_RMARK_DEL,
            SPRM_C_RSID_RM_DEL,
        ],
        RevisionKind::CharacterFormatting => {
            let mut value = vec![
                SPRM_C_PROP_RMARK90,
                SPRM_C_PROP_RMARK_CURRENT,
                SPRM_C_RSID_PROP,
            ];
            if remove_wall {
                value.push(SPRM_C_WALL);
            }
            value
        },
        RevisionKind::ParagraphFormatting => {
            let mut value = vec![
                SPRM_P_PROP_RMARK,
                SPRM_P_PROP_RMARK90,
                SPRM_P_PROP_RMARK_CURRENT,
            ];
            if remove_wall {
                value.push(SPRM_P_WALL);
            }
            value
        },
        RevisionKind::TableRowFormatting => {
            let mut value = vec![SPRM_T_PROP_RMARK, SPRM_T_RSID];
            if remove_wall {
                value.push(SPRM_T_WALL);
            }
            value
        },
    }
}

pub(super) fn restore_before_wall(
    grp: &[u8],
    wall: u16,
    revision_marks: &[u16],
) -> Result<Vec<u8>> {
    let sprms = strict_sprms(grp)?;
    if let Some(marker) = sprms
        .iter()
        .rfind(|sprm| sprm.opcode == wall && sprm.operand_byte() == Some(1))
    {
        Ok(grp[..marker.offset].to_vec())
    } else {
        // Some Word 97-era producers emitted PropRMark without a wall. There
        // is no recoverable previous state, so rejection safely clears only
        // the revision metadata and retains the current properties.
        retain_sprms(grp, revision_marks)
    }
}

pub(super) fn retain_sprms(grp: &[u8], remove: &[u16]) -> Result<Vec<u8>> {
    let sprms = strict_sprms(grp)?;
    let mut output = Vec::with_capacity(grp.len());
    for sprm in sprms {
        if !remove.contains(&sprm.opcode) {
            output.extend_from_slice(&grp[sprm.offset..sprm.offset + sprm.size]);
        }
    }
    Ok(output)
}

pub(super) fn strict_sprms(grp: &[u8]) -> Result<Vec<crate::sprm::Sprm>> {
    parse_sprms(grp)
        .map_err(|error| corrupted(format!("malformed or overlapping SPRM sequence: {error}")))
}

/// Parse one PAPX group and retain only the final table-membership state.
///
/// PAPX carries a two-byte style index before its SPRM body. The reverse
/// lookup deliberately mirrors the legacy reader semantics: duplicate
/// `sprmPFInTable` records are resolved by the last record, and only the
/// literal byte value `1` means that the paragraph is in a table.
pub(super) fn papx_in_table_state(grpprl: &[u8]) -> Result<bool> {
    let body = grpprl
        .get(2..)
        .ok_or_else(|| corrupted("PAPX has no style index"))?;
    Ok(strict_sprms(body)?
        .iter()
        .rev()
        .find(|sprm| sprm.opcode == SPRM_P_F_IN_TABLE)
        .is_some_and(|sprm| sprm.operand_byte() == Some(1)))
}

pub(super) fn split_transform_chpx(
    runs: &mut Vec<FcRun>,
    intervals: &[(u32, u32)],
    transform: impl Fn(&[u8]) -> Result<Vec<u8>>,
) -> Result<()> {
    let mut output = Vec::new();
    for run in runs.iter() {
        let mut cuts = BTreeSet::from([run.start, run.end]);
        for (a, b) in intervals {
            if *a > run.start && *a < run.end {
                cuts.insert(*a);
            }
            if *b > run.start && *b < run.end {
                cuts.insert(*b);
            }
        }
        let cuts = cuts.into_iter().collect::<Vec<_>>();
        for pair in cuts.windows(2) {
            let covered = intervals
                .iter()
                .any(|(a, b)| pair[0] >= *a && pair[1] <= *b);
            output.push(FcRun {
                start: pair[0],
                end: pair[1],
                grpprl: if covered {
                    transform(&run.grpprl)?
                } else {
                    run.grpprl.clone()
                },
            });
        }
    }
    if !intervals.iter().all(|(a, b)| {
        let mut cursor = *a;
        for run in output.iter().filter(|run| run.end > *a && run.start < *b) {
            if run.start > cursor {
                return false;
            }
            cursor = cursor.max(run.end);
            if cursor >= *b {
                return true;
            }
        }
        false
    }) {
        return Err(corrupted("tracked range is outside CHPX coverage"));
    }
    merge_fc_runs(&mut output);
    *runs = output;
    Ok(())
}

pub(super) fn split_transform_papx(
    runs: &mut [PapxRun],
    intervals: &[(u32, u32)],
    transform: impl Fn(&[u8]) -> Result<Vec<u8>>,
) -> Result<()> {
    let mut replacements = Vec::new();
    for (index, run) in runs.iter().enumerate() {
        if intervals
            .iter()
            .any(|(a, b)| *a < run.end && *b > run.start)
        {
            let grpprl = transform(&run.grpprl)?;
            let in_table = papx_in_table_state(&grpprl)?;
            replacements.push((index, grpprl, in_table));
        }
    }
    if replacements.is_empty() {
        return Err(corrupted("tracked property range is outside PAPX coverage"));
    }
    for (index, grpprl, in_table) in replacements {
        let run = &mut runs[index];
        run.replace_grpprl(grpprl, in_table);
    }
    Ok(())
}

pub(super) fn parse_chpx(word: &[u8], table: &[u8]) -> Result<Vec<FcRun>> {
    let (fcs, pages) = parse_bte(word, table, PLCFBTE_CHPX)?;
    let mut output = Vec::new();
    for (index, pn) in pages.iter().enumerate() {
        let offset = (*pn as usize)
            .checked_mul(512)
            .ok_or_else(|| corrupted("CHPX page offset overflow"))?;
        let page = word
            .get(offset..offset + 512)
            .ok_or_else(|| corrupted("CHPX page exceeds WordDocument"))?;
        let fkp = ChpxFkp::parse(page, word).ok_or_else(|| corrupted("malformed CHPX FKP"))?;
        for entry_index in 0..fkp.count() {
            let entry = fkp
                .entry(entry_index)
                .ok_or_else(|| corrupted("PAPX entry is missing"))?;
            if entry.fc < fcs[index] || entry.end_fc > fcs[index + 1] {
                return Err(corrupted("CHPX FKP overlaps its BTE range"));
            }
            strict_sprms(&entry.grpprl)?;
            output.push(FcRun {
                start: entry.fc,
                end: entry.end_fc,
                grpprl: entry.grpprl.clone(),
            });
        }
    }
    output.sort_by_key(|r| r.start);
    if output.windows(2).any(|v| v[0].end > v[1].start) {
        return Err(corrupted("CHPX runs overlap"));
    }
    Ok(output)
}

pub(super) fn parse_papx(word: &[u8], table: &[u8]) -> Result<Vec<PapxRun>> {
    let (fcs, pages) = parse_bte(word, table, PLCFBTE_PAPX)?;
    let mut output = Vec::new();
    for (index, pn) in pages.iter().enumerate() {
        let offset = (*pn as usize)
            .checked_mul(512)
            .ok_or_else(|| corrupted("PAPX page offset overflow"))?;
        let page = word
            .get(offset..offset + 512)
            .ok_or_else(|| corrupted("PAPX page exceeds WordDocument"))?;
        let fkp = PapxFkp::parse(page, word).ok_or_else(|| corrupted("malformed PAPX FKP"))?;
        for entry_index in 0..fkp.count() {
            let entry = fkp
                .entry(entry_index)
                .ok_or_else(|| corrupted("PAPX entry is missing"))?;
            if entry.fc < fcs[index] || entry.end_fc > fcs[index + 1] {
                return Err(corrupted("PAPX FKP overlaps its BTE range"));
            }
            let grpprl = entry.grpprl.clone();
            if grpprl.len() < 2 {
                return Err(corrupted("PAPX style index is missing"));
            }
            let in_table = papx_in_table_state(&grpprl)?;
            output.push(PapxRun::new(
                entry.fc,
                entry.end_fc,
                grpprl,
                in_table,
                entry
                    .paragraph_height
                    .ok_or_else(|| corrupted("PAPX PHE is missing"))?,
            ));
        }
    }
    output.sort_by_key(|r| r.start);
    if output.windows(2).any(|v| v[0].end > v[1].start) {
        return Err(corrupted("PAPX runs overlap"));
    }
    Ok(output)
}

pub(super) struct BuiltPapxPage {
    pub(super) start: u32,
    pub(super) end: u32,
    pub(super) bytes: Vec<u8>,
}

pub(super) fn build_papx_pages(runs: &[PapxRun]) -> Result<Vec<BuiltPapxPage>> {
    let mut pages = Vec::new();
    let mut start = 0;
    while start < runs.len() {
        let mut count = 0;
        let mut used = 1usize;
        for run in &runs[start..] {
            let data = papx_storage_size(run.grpprl.len())?;
            let front = (count + 2) * 4 + (count + 1) * 13;
            if front + used + data > 512 {
                break;
            }
            count += 1;
            used += data;
        }
        if count == 0 {
            return Err(corrupted("one PAPX run cannot fit in an FKP"));
        }
        let subset = &runs[start..start + count];
        let first = subset
            .first()
            .ok_or_else(|| corrupted("PAPX page has no runs"))?;
        let last = subset
            .last()
            .ok_or_else(|| corrupted("PAPX page has no runs"))?;
        pages.push(BuiltPapxPage {
            start: first.start,
            end: last.end,
            bytes: build_papx_page(subset)?,
        });
        start += count;
    }
    Ok(pages)
}

pub(super) fn papx_storage_size(len: usize) -> Result<usize> {
    if len == 0 || len > 510 {
        return Err(corrupted("PAPX length is invalid"));
    }
    let prefix = if len.is_multiple_of(2) { 2 } else { 1 };
    let raw = prefix + len;
    Ok(raw + raw % 2)
}

pub(super) fn build_papx_page(runs: &[PapxRun]) -> Result<Vec<u8>> {
    let mut page = vec![0u8; 512];
    let n = runs.len();
    for (i, run) in runs.iter().enumerate() {
        page[i * 4..i * 4 + 4].copy_from_slice(&run.start.to_le_bytes());
    }
    page[n * 4..n * 4 + 4].copy_from_slice(&runs[n - 1].end.to_le_bytes());
    page[511] = u8::try_from(n).map_err(|_| corrupted("too many PAPX entries"))?;
    let bx = (n + 1) * 4;
    let mut cursor = 511usize;
    for (i, run) in runs.iter().enumerate().rev() {
        let prefix = if run.grpprl.len() % 2 == 0 { 2 } else { 1 };
        cursor = cursor
            .checked_sub(prefix + run.grpprl.len())
            .ok_or_else(|| corrupted("PAPX page overflow"))?;
        cursor &= !1;
        page[bx + i * 13] =
            u8::try_from(cursor / 2).map_err(|_| corrupted("PAPX offset exceeds byte"))?;
        write_phe(&mut page[bx + i * 13 + 1..bx + i * 13 + 13], run.phe);
        if prefix == 1 {
            page[cursor] = (run.grpprl.len().div_ceil(2)) as u8;
        } else {
            page[cursor] = 0;
            page[cursor + 1] = (run.grpprl.len() / 2) as u8;
        }
        page[cursor + prefix..cursor + prefix + run.grpprl.len()].copy_from_slice(&run.grpprl);
    }
    Ok(page)
}

pub(super) fn write_phe(slot: &mut [u8], value: ParagraphHeight) {
    slot[0..2].copy_from_slice(&value.info_field.to_le_bytes());
    slot[2..4].copy_from_slice(&value.reserved.to_le_bytes());
    slot[4..8].copy_from_slice(&value.dxa_col.to_le_bytes());
    slot[8..12].copy_from_slice(&value.dym_line_or_height.to_le_bytes());
}

pub(super) fn parse_bte(word: &[u8], table: &[u8], index: usize) -> Result<(Vec<u32>, Vec<u32>)> {
    let (offset, length) = fib_pair(word, index)?;
    let data = slice(table, offset, length, "bin table")?;
    if data.len() < 12 || (data.len() - 4) % 8 != 0 {
        return Err(corrupted("bin table shape is invalid"));
    }
    let n = (data.len() - 4) / 8;
    let mut fcs = Vec::with_capacity(n + 1);
    let mut pages = Vec::with_capacity(n);
    for i in 0..=n {
        fcs.push(u32_at(data, i * 4)?);
    }
    for i in 0..n {
        pages.push(u32_at(data, (n + 1) * 4 + i * 4)? & 0x003F_FFFF);
    }
    if fcs.windows(2).any(|v| v[0] >= v[1]) || pages.contains(&0) {
        return Err(corrupted("bin table ranges are invalid"));
    }
    Ok((fcs, pages))
}

pub(super) fn parse_clx(word: &[u8], table: &[u8]) -> Result<Vec<RawPiece>> {
    let (offset, length) = fib_pair(word, CLX)?;
    let data = slice(table, offset, length, "CLX")?;
    if data.first() != Some(&2) {
        return Err(corrupted("fast-save RgPrc CLX is unsupported"));
    }
    let size = u32_at(data, 1)? as usize;
    if size + 5 != data.len() || size < 4 || !(size - 4).is_multiple_of(12) {
        return Err(corrupted("CLX PlcPcd shape is invalid"));
    }
    let n = (size - 4) / 12;
    if n == 0 || n > MAX_PIECES {
        return Err(corrupted("piece count exceeds limits"));
    }
    let cp_base = 5;
    let pcd_base = cp_base + (n + 1) * 4;
    let mut output = Vec::with_capacity(n);
    for i in 0..n {
        let start = u32_at(data, cp_base + i * 4)?;
        let end = u32_at(data, cp_base + (i + 1) * 4)?;
        if start >= end || output.last().is_some_and(|p: &RawPiece| p.end != start) {
            return Err(corrupted("piece CPs overlap or contain gaps"));
        }
        let p = &data[pcd_base + i * 8..pcd_base + i * 8 + 8];
        let raw = u32_at(p, 2)?;
        if raw & 0x8000_0000 != 0 {
            return Err(corrupted("piece FC reserved bit is set"));
        }
        let unicode = raw & 0x4000_0000 == 0;
        let fc = if unicode {
            raw & 0x3FFF_FFFF
        } else {
            (raw & 0x3FFF_FFFF) / 2
        };
        let byte_len = (end - start)
            .checked_mul(if unicode { 2 } else { 1 })
            .ok_or_else(|| corrupted("piece byte length overflow"))?;
        if fc
            .checked_add(byte_len)
            .is_none_or(|v| v as usize > word.len())
        {
            return Err(corrupted("piece text exceeds WordDocument"));
        }
        if p[6] != 0 || p[7] != 0 {
            return Err(corrupted("piece PRMs are unsupported for safe mutation"));
        }
        output.push(RawPiece {
            start,
            end,
            fc,
            unicode,
            prefix: [p[0], p[1]],
            prm: [p[6], p[7]],
        });
    }
    Ok(output)
}

pub(super) fn serialize_clx(pieces: &[RawPiece]) -> Result<Vec<u8>> {
    if pieces.is_empty() || pieces.len() > MAX_PIECES {
        return Err(corrupted("piece count is invalid"));
    }
    let plc = pieces
        .len()
        .checked_mul(12)
        .and_then(|v| v.checked_add(4))
        .ok_or_else(|| corrupted("PlcPcd size overflow"))?;
    let mut out = vec![2];
    out.extend_from_slice(&(plc as u32).to_le_bytes());
    for p in pieces {
        out.extend_from_slice(&p.start.to_le_bytes());
    }
    let end = pieces
        .last()
        .ok_or_else(|| corrupted("piece table is empty"))?
        .end;
    out.extend_from_slice(&end.to_le_bytes());
    for p in pieces {
        out.extend_from_slice(&p.prefix);
        let raw = if p.unicode {
            p.fc
        } else {
            p.fc.checked_mul(2)
                .ok_or_else(|| corrupted("compressed FC overflow"))?
                | 0x4000_0000
        };
        out.extend_from_slice(&raw.to_le_bytes());
        out.extend_from_slice(&p.prm);
    }
    Ok(out)
}

pub(super) fn insert_piece(
    pieces: &mut Vec<RawPiece>,
    cp: u32,
    length: u32,
    fc: u32,
) -> Result<()> {
    let index = pieces
        .iter()
        .position(|p| cp >= p.start && cp <= p.end)
        .ok_or_else(|| corrupted("insertion CP has no piece"))?;
    for piece in pieces.iter_mut().skip(index + 1) {
        piece.start = piece
            .start
            .checked_add(length)
            .ok_or_else(|| corrupted("piece CP overflow"))?;
        piece.end = piece
            .end
            .checked_add(length)
            .ok_or_else(|| corrupted("piece CP overflow"))?;
    }
    let original = pieces[index].clone();
    let mut replacement = Vec::new();
    if cp > original.start {
        let mut left = original.clone();
        left.end = cp;
        replacement.push(left);
    }
    let inserted_end = cp
        .checked_add(length)
        .ok_or_else(|| corrupted("piece CP overflow"))?;
    replacement.push(RawPiece {
        start: cp,
        end: inserted_end,
        fc,
        unicode: true,
        prefix: [0, 0],
        prm: [0, 0],
    });
    if cp < original.end {
        let mut right = original;
        let width = if right.unicode { 2 } else { 1 };
        right.fc += (cp - right.start) * width;
        right.start = cp + length;
        right.end += length;
        replacement.push(right);
    }
    pieces.splice(index..=index, replacement);
    normalize_piece_cps(pieces)
}

pub(super) fn delete_piece_range(pieces: &mut Vec<RawPiece>, start: u32, end: u32) -> Result<()> {
    let mut output = Vec::new();
    let removed = end - start;
    for piece in pieces.iter() {
        if piece.end <= start {
            output.push(piece.clone());
            continue;
        }
        if piece.start >= end {
            let mut p = piece.clone();
            p.start -= removed;
            p.end -= removed;
            output.push(p);
            continue;
        }
        if piece.start < start {
            let mut left = piece.clone();
            left.end = start;
            output.push(left);
        }
        if piece.end > end {
            let mut right = piece.clone();
            let width = if right.unicode { 2 } else { 1 };
            right.fc += (end - right.start) * width;
            right.start = start;
            right.end = piece.end - removed;
            output.push(right);
        }
    }
    if output.is_empty() {
        return Err(corrupted("cannot delete every document piece"));
    }
    *pieces = output;
    normalize_piece_cps(pieces)
}

pub(super) fn normalize_piece_cps(pieces: &[RawPiece]) -> Result<()> {
    if pieces.windows(2).any(|v| v[0].end != v[1].start) || pieces.iter().any(|p| p.start >= p.end)
    {
        return Err(corrupted("piece rewrite produced gaps or overlaps"));
    }
    Ok(())
}

pub(super) fn parse_authors(word: &[u8], table: &[u8]) -> Result<Vec<String>> {
    let (offset, length) = fib_pair(word, STTBFRMARK)?;
    if length == 0 {
        return Ok(vec!["Unknown".to_string()]);
    }
    let data = slice(table, offset, length, "SttbfRMark")?;
    if data.len() < 6 || u16_at(data, 0)? != 0xFFFF || u16_at(data, 4)? != 0 {
        return Err(corrupted("SttbfRMark header is invalid"));
    }
    let count = u16_at(data, 2)? as usize;
    let mut cursor = 6;
    let mut output = Vec::with_capacity(count);
    for _ in 0..count {
        let n = u16_at(data, cursor)? as usize;
        cursor += 2;
        let bytes = data
            .get(cursor..cursor + n * 2)
            .ok_or_else(|| corrupted("revision author is truncated"))?;
        let units = bytes
            .chunks_exact(2)
            .map(|v| u16::from_le_bytes([v[0], v[1]]))
            .collect::<Vec<_>>();
        output.push(
            String::from_utf16(&units)
                .map_err(|_| corrupted("revision author is invalid UTF-16"))?,
        );
        cursor += n * 2;
    }
    if cursor != data.len() || output.first().map(String::as_str) != Some("Unknown") {
        return Err(corrupted(
            "SttbfRMark must begin with Unknown and have no trailing bytes",
        ));
    }
    Ok(output)
}

pub(super) fn serialize_authors(authors: &[String]) -> Result<Vec<u8>> {
    if authors.first().map(String::as_str) != Some("Unknown") || authors.len() > MAX_AUTHORS {
        return Err(corrupted("revision author table is invalid"));
    }
    let mut out = Vec::new();
    out.extend_from_slice(&0xFFFFu16.to_le_bytes());
    out.extend_from_slice(&(authors.len() as u16).to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    for author in authors {
        let units = author.encode_utf16().collect::<Vec<_>>();
        let n = u16::try_from(units.len()).map_err(|_| corrupted("revision author is too long"))?;
        out.extend_from_slice(&n.to_le_bytes());
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(out)
}

pub(super) fn parse_cp_table(
    word: &[u8],
    table: &[u8],
    index: usize,
    record_size: usize,
) -> Result<Option<CpTable>> {
    let (offset, length) = fib_pair(word, index)?;
    if length == 0 {
        return Ok(None);
    }
    let data = slice(table, offset, length, "CP-indexed PLCF")?;
    if data.len() < 4 || (data.len() - 4) % (4 + record_size) != 0 {
        return Err(corrupted("CP-indexed PLCF shape is invalid"));
    }
    let n = (data.len() - 4) / (4 + record_size);
    let mut cps = Vec::with_capacity(n + 1);
    for i in 0..=n {
        cps.push(u32_at(data, i * 4)?);
    }
    if cps.windows(2).any(|v| v[0] > v[1]) {
        return Err(corrupted("CP-indexed PLCF overlaps"));
    }
    Ok(Some(CpTable {
        index,
        cps,
        records: data[(n + 1) * 4..].to_vec(),
    }))
}

pub(super) fn reject_protection(word: &[u8], table: &[u8]) -> Result<()> {
    let (offset, length) = fib_pair(word, DOP)?;
    if length == 0 {
        return Ok(());
    }
    let dop = slice(table, offset, length, "DOP")?;
    if dop.len() < 84 {
        return Err(corrupted("DOP is truncated"));
    }
    let protected = dop[6] & 0x10 != 0
        || dop[7] & (0x02 | 0x20 | 0x40) != 0
        || i32::from_le_bytes(array_at(dop, 78, "DOP protection key")?) != 0;
    if protected {
        return Err(corrupted("protected DOC cannot be edited"));
    }
    Ok(())
}

pub(super) fn read_units(
    word: &[u8],
    pieces: &[RawPiece],
    start: u32,
    end: u32,
) -> Result<Vec<u16>> {
    let mut out = Vec::new();
    for piece in pieces {
        let a = start.max(piece.start);
        let b = end.min(piece.end);
        if a >= b {
            continue;
        }
        let relative = a - piece.start;
        let count = b - a;
        if piece.unicode {
            let offset = (piece.fc + relative * 2) as usize;
            let bytes = word
                .get(offset..offset + count as usize * 2)
                .ok_or_else(|| corrupted("text range exceeds WordDocument"))?;
            out.extend(
                bytes
                    .chunks_exact(2)
                    .map(|v| u16::from_le_bytes([v[0], v[1]])),
            );
        } else {
            let offset = (piece.fc + relative) as usize;
            let bytes = word
                .get(offset..offset + count as usize)
                .ok_or_else(|| corrupted("text range exceeds WordDocument"))?;
            out.extend(bytes.iter().map(|v| u16::from(*v)));
        }
    }
    Ok(out)
}

pub(super) fn infer_moves(revisions: &mut [Revision]) {
    let mut groups: HashMap<u32, (Vec<usize>, Vec<usize>)> = HashMap::new();
    for (i, revision) in revisions.iter().enumerate() {
        if let Some(rsid) = revision.revision_save_id {
            let group = groups.entry(rsid).or_default();
            match revision.kind {
                RevisionKind::Insertion => group.1.push(i),
                RevisionKind::Deletion => group.0.push(i),
                _ => {},
            }
        }
    }
    for (rsid, (from, to)) in groups {
        if from.len() == 1 && to.len() == 1 {
            revisions[from[0]].kind = RevisionKind::MoveFrom;
            revisions[to[0]].kind = RevisionKind::MoveTo;
            revisions[from[0]].move_pair_id = Some(rsid);
            revisions[to[0]].move_pair_id = Some(rsid);
        }
    }
}

pub(super) fn merge_adjacent(output: &mut Vec<Revision>) {
    output.sort_by_key(|item| {
        (
            kind_order(item.kind),
            item.author_index,
            item.reason,
            item.revision_save_id,
            item.start_cp,
            item.end_cp,
        )
    });
    let mut merged: Vec<Revision> = Vec::new();
    for item in output.drain(..) {
        if let Some(last) = merged.last_mut()
            && last.end_cp == item.start_cp
            && last.kind == item.kind
            && last.author_index == item.author_index
            && last.timestamp == item.timestamp
            && last.reason == item.reason
            && last.revision_save_id == item.revision_save_id
        {
            last.end_cp = item.end_cp;
            continue;
        }
        merged.push(item);
    }
    merged.sort_by_key(|item| (item.start_cp, item.end_cp, kind_order(item.kind)));
    *output = merged;
}

pub(super) fn merge_fc_runs(runs: &mut Vec<FcRun>) {
    let mut out: Vec<FcRun> = Vec::new();
    for run in runs.drain(..) {
        if let Some(last) = out.last_mut()
            && last.end == run.start
            && last.grpprl == run.grpprl
        {
            last.end = run.end;
            continue;
        }
        out.push(run);
    }
    *runs = out;
}
pub(super) fn kind_order(kind: RevisionKind) -> u8 {
    match kind {
        RevisionKind::Insertion | RevisionKind::MoveTo => 0,
        RevisionKind::Deletion | RevisionKind::MoveFrom => 1,
        RevisionKind::CharacterFormatting => 2,
        RevisionKind::ParagraphFormatting => 3,
        RevisionKind::TableRowFormatting => 4,
    }
}

#[cfg(test)]
mod papx_cache_tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    fn table_sprm(value: u8) -> [u8; 3] {
        let opcode = SPRM_P_F_IN_TABLE.to_le_bytes();
        [opcode[0], opcode[1], value]
    }

    fn papx_grpprl(sprms: &[u8]) -> Vec<u8> {
        let mut grpprl = vec![0x34, 0x12];
        grpprl.extend_from_slice(sprms);
        grpprl
    }

    fn papx_run(start: u32, end: u32, grpprl: Vec<u8>) -> PapxRun {
        let in_table = papx_in_table_state(&grpprl).expect("test PAPX should parse");
        PapxRun::new(
            start,
            end,
            grpprl,
            in_table,
            ParagraphHeight {
                info_field: 0,
                reserved: 0,
                dxa_col: 0,
                dym_line_or_height: 0,
            },
        )
    }

    #[test]
    fn papx_table_state_is_absent_single_and_last_duplicate() {
        assert!(!papx_in_table_state(&papx_grpprl(&[])).unwrap());
        assert!(papx_in_table_state(&papx_grpprl(&table_sprm(1))).unwrap());
        assert!(!papx_in_table_state(&papx_grpprl(&table_sprm(0))).unwrap());
        assert!(
            papx_in_table_state(&papx_grpprl(&[table_sprm(0), table_sprm(1),].concat())).unwrap()
        );
        assert!(
            !papx_in_table_state(&papx_grpprl(&[table_sprm(1), table_sprm(0),].concat())).unwrap()
        );
        assert!(!papx_in_table_state(&papx_grpprl(&table_sprm(2))).unwrap());
    }

    #[test]
    fn papx_table_state_rejects_malformed_grpprl_before_caching() {
        let malformed = papx_grpprl(&SPRM_P_F_IN_TABLE.to_le_bytes());
        assert!(papx_in_table_state(&malformed).is_err());
        let mut valid_prefix = papx_grpprl(&table_sprm(1));
        valid_prefix.extend_from_slice(&SPRM_P_F_IN_TABLE.to_le_bytes());
        assert!(papx_in_table_state(&valid_prefix).is_err());
        assert!(papx_in_table_state(&[]).is_err());
        assert!(papx_in_table_state(&[0]).is_err());
    }

    #[test]
    fn split_transform_papx_updates_cache_only_after_strict_success() {
        let original = papx_grpprl(&table_sprm(1));
        let mut runs = vec![papx_run(10, 20, original.clone())];

        split_transform_papx(&mut runs, &[(10, 11)], |grpprl| {
            let mut transformed = grpprl.to_vec();
            transformed.extend_from_slice(&table_sprm(0));
            Ok(transformed)
        })
        .unwrap();
        assert!(!runs[0].in_table());
        assert_eq!(runs[0].grpprl, [original, table_sprm(0).to_vec()].concat());

        let before = runs[0].clone();
        assert!(
            split_transform_papx(&mut runs, &[(20, 21)], |_| {
                Ok(papx_grpprl(&SPRM_P_F_IN_TABLE.to_le_bytes()))
            })
            .is_err()
        );
        assert_eq!(runs[0].grpprl, before.grpprl);
        assert_eq!(runs[0].in_table(), before.in_table());

        split_transform_papx(&mut runs, &[(19, 20)], |grpprl| Ok(grpprl.to_vec())).unwrap();
        assert!(!runs[0].in_table());
    }

    #[test]
    fn restoring_before_wall_recomputes_table_state_and_keeps_wire_prefix() {
        let mut body = Vec::from(table_sprm(1));
        body.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
        body.push(1);
        body.extend_from_slice(&table_sprm(0));
        let original = papx_grpprl(&body);
        let mut runs = vec![papx_run(0, 4, original)];

        split_transform_papx(&mut runs, &[(1, 2)], |grpprl| {
            let mut restored = grpprl[..2].to_vec();
            restored.extend_from_slice(&restore_before_wall(
                &grpprl[2..],
                SPRM_P_WALL,
                &[SPRM_P_WALL],
            )?);
            Ok(restored)
        })
        .unwrap();
        assert!(runs[0].in_table());
        assert_eq!(runs[0].grpprl, papx_grpprl(&table_sprm(1)));
    }

    #[test]
    fn split_transform_papx_is_atomic_when_a_later_run_is_malformed() {
        let first = papx_grpprl(&table_sprm(1));
        let second = papx_grpprl(&table_sprm(0));
        let mut runs = vec![
            papx_run(0, 10, first.clone()),
            papx_run(10, 20, second.clone()),
        ];
        let before = runs.clone();
        assert!(
            split_transform_papx(&mut runs, &[(0, 20)], |grpprl| {
                if grpprl == first.as_slice() {
                    Ok(papx_grpprl(&table_sprm(0)))
                } else {
                    Ok(papx_grpprl(&SPRM_P_F_IN_TABLE.to_le_bytes()))
                }
            })
            .is_err()
        );
        for (actual, expected) in runs.iter().zip(before) {
            assert_eq!(actual.grpprl, expected.grpprl);
            assert_eq!(actual.in_table(), expected.in_table());
        }
    }

    #[test]
    fn papx_cache_stays_paired_through_sort_and_serialization() {
        let first = papx_grpprl(&table_sprm(1));
        let second = papx_grpprl(&table_sprm(0));
        let mut runs = vec![
            papx_run(10, 20, second.clone()),
            papx_run(0, 10, first.clone()),
        ];
        runs.sort_by_key(|run| run.start);
        assert!(runs[0].in_table());
        assert!(!runs[1].in_table());

        let pages = build_papx_pages(&runs).unwrap();
        let parsed = PapxFkp::parse(&pages[0].bytes, &[]).unwrap();
        assert_eq!(parsed.entry(0).unwrap().grpprl, first);
        assert_eq!(parsed.entry(1).unwrap().grpprl, second);
    }

    #[test]
    fn cached_state_matches_strict_reverse_lookup_for_duplicate_sequences() {
        let values = [0, 1, 2, 255];
        for first in values {
            for last in values {
                let body = [table_sprm(first), table_sprm(last)].concat();
                let grpprl = papx_grpprl(&body);
                let expected = strict_sprms(&grpprl[2..])
                    .unwrap()
                    .iter()
                    .rev()
                    .find(|sprm| sprm.opcode == SPRM_P_F_IN_TABLE)
                    .is_some_and(|sprm| sprm.operand_byte() == Some(1));
                assert_eq!(papx_in_table_state(&grpprl).unwrap(), expected);
                assert_eq!(grpprl, papx_grpprl(&body));
            }
        }
    }

    #[test]
    fn papx_run_cache_is_send_sync() {
        assert_send_sync::<PapxRun>();
    }
}
pub(super) fn validate_range(start: u32, end: u32, limit: u32) -> Result<()> {
    if start >= end || end > limit {
        Err(corrupted(
            "tracked revision range is empty or exceeds the main story",
        ))
    } else {
        Ok(())
    }
}
pub(super) fn pack_dttm(value: Option<DateTime>) -> Result<u32> {
    let Some(v) = value else { return Ok(0) };
    if !(1900..=2411).contains(&v.year)
        || !(1..=12).contains(&v.month)
        || !(1..=31).contains(&v.day)
        || v.hour > 23
        || v.minute > 59
        || v.weekday > 6
    {
        return Err(corrupted("revision timestamp is outside DTTM limits"));
    }
    Ok(u32::from(v.minute)
        | u32::from(v.hour) << 6
        | u32::from(v.day) << 11
        | u32::from(v.month) << 16
        | u32::from(v.year - 1900) << 20
        | u32::from(v.weekday) << 29)
}
pub(super) fn decode_dttm(raw: u32) -> Result<Option<DateTime>> {
    if raw == 0 {
        return Ok(None);
    }
    let value = DateTime {
        minute: (raw & 0x3f) as u8,
        hour: ((raw >> 6) & 0x1f) as u8,
        day: ((raw >> 11) & 0x1f) as u8,
        month: ((raw >> 16) & 0xf) as u8,
        year: ((raw >> 20) & 0x1ff) as u16 + 1900,
        weekday: ((raw >> 29) & 7) as u8,
    };
    pack_dttm(Some(value))?;
    Ok(Some(value))
}
pub(super) fn push_byte(out: &mut Vec<u8>, op: u16, v: u8) {
    out.extend_from_slice(&op.to_le_bytes());
    out.push(v);
}
pub(super) fn push_word(out: &mut Vec<u8>, op: u16, v: u16) {
    out.extend_from_slice(&op.to_le_bytes());
    out.extend_from_slice(&v.to_le_bytes());
}
pub(super) fn push_dword(out: &mut Vec<u8>, op: u16, v: u32) {
    out.extend_from_slice(&op.to_le_bytes());
    out.extend_from_slice(&v.to_le_bytes());
}
pub(super) fn append_table_block(
    word: &mut [u8],
    table: &mut Vec<u8>,
    index: usize,
    data: &[u8],
) -> Result<()> {
    let offset = u32::try_from(table.len()).map_err(|_| corrupted("Table stream exceeds u32"))?;
    table.extend_from_slice(data);
    put_fib_pair(
        word,
        index,
        offset,
        u32::try_from(data.len()).map_err(|_| corrupted("table block exceeds u32"))?,
    )
}
pub(super) fn fib_pair(word: &[u8], index: usize) -> Result<(u32, u32)> {
    Ok((
        u32_at(word, FIB_FC_LCB + index * 8)?,
        u32_at(word, FIB_FC_LCB + index * 8 + 4)?,
    ))
}
pub(super) fn put_fib_pair(word: &mut [u8], index: usize, fc: u32, lcb: u32) -> Result<()> {
    put_u32(word, FIB_FC_LCB + index * 8, fc)?;
    put_u32(word, FIB_FC_LCB + index * 8 + 4, lcb)
}
pub(super) fn slice<'a>(data: &'a [u8], offset: u32, length: u32, name: &str) -> Result<&'a [u8]> {
    let start = offset as usize;
    let end = start
        .checked_add(length as usize)
        .ok_or_else(|| corrupted(format!("{name} range overflow")))?;
    data.get(start..end)
        .ok_or_else(|| corrupted(format!("{name} exceeds stream")))
}
pub(super) fn u16_at(data: &[u8], offset: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(array_at(data, offset, "u16")?))
}
pub(super) fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    Ok(u32::from_le_bytes(array_at(data, offset, "u32")?))
}
pub(super) fn put_u32(data: &mut [u8], offset: usize, value: u32) -> Result<()> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| corrupted("FIB field offset overflow"))?;
    data.get_mut(offset..end)
        .ok_or_else(|| corrupted("truncated FIB field"))?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}
pub(super) fn array_at<const N: usize>(data: &[u8], offset: usize, name: &str) -> Result<[u8; N]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| corrupted(format!("{name} offset overflow")))?;
    data.get(offset..end)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or_else(|| corrupted(format!("truncated {name}")))
}
pub(super) fn align2(v: usize) -> Result<usize> {
    v.checked_add(1)
        .map(|n| n & !1)
        .ok_or_else(|| corrupted("alignment overflow"))
}
pub(super) fn align512(v: usize) -> Result<usize> {
    v.checked_add(511)
        .map(|n| n & !511)
        .ok_or_else(|| corrupted("alignment overflow"))
}
pub(super) fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
