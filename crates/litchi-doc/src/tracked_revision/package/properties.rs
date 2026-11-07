//! Paragraph-property, field, bookmark and section primitives of the
//! revision editor (record 0324).
//!
//! Each primitive clones the editor, rewrites only the structures it owns
//! (a paragraph's PAPX, the main-story text and its `PlcffldMom` markers,
//! the standard bookmark tables, or one section's SEPX), appends the
//! rewritten tables to their streams and publishes the candidate once. Every
//! byte outside those structures is kept.

use super::{
    CpTable, FcRun, MAX_TEXT_UNITS, PLCFBKF, PLCFBKL, PLCFFLD_MOM, PLCFSED, PapxRun, Result,
    RevisionEditor, align2, append_table_block, containing_piece, corrupted, delete_piece_range,
    fib_pair, insert_piece, papx_sprms, read_units, slice, u16_at, u32_at,
};
use crate::parts::fib::FileInformationBlock;
use crate::parts::styles::StyleSheet;
use crate::tracked_revision::codec::papx_table_state;

/// `FibRgFcLcb97` index of the standard bookmark names, `SttbfBkmk`.
const STTBF_BKMK: usize = 21;
/// `sprmPHugePapx`: a PAPX stored in the Data stream, never edited here.
const SPRM_P_HUGE_PAPX: u16 = 0x6646;

/// One standard bookmark as the editor's tables store it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct BookmarkRange {
    pub(crate) name: String,
    pub(crate) start: u32,
    pub(crate) end: u32,
}

impl RevisionEditor {
    /// The document's stylesheet, parsed from the current table stream.
    pub(crate) fn stylesheet(&self) -> Result<StyleSheet> {
        let fib = FileInformationBlock::parse(&self.word)?;
        StyleSheet::parse(&fib, &self.table)
    }

    /// The PAPX (style index and SPRMs) of the paragraph whose mark is at
    /// `mark`.
    pub(crate) fn paragraph_grpprl(&self, mark: u32) -> Result<Vec<u8>> {
        Ok(self.paragraph_papx_at(mark)?.grpprl.clone())
    }

    /// Replace the PAPX of the paragraph `start..=mark` (its mark at
    /// `mark`), splitting the PAPX run so no other paragraph changes, and
    /// rebuild the PAPX FKPs.
    pub(crate) fn set_paragraph_grpprl(
        &mut self,
        start: u32,
        mark: u32,
        grpprl: Vec<u8>,
    ) -> Result<()> {
        if start > mark || mark >= self.main_ccp {
            return Err(corrupted("paragraph range is outside the main story"));
        }
        if grpprl.len() < 2 {
            return Err(corrupted("PAPX has no style index"));
        }
        let current = self.paragraph_papx_at(mark)?;
        if papx_sprms(&current.grpprl[2..])?
            .iter()
            .any(|sprm| sprm.opcode == SPRM_P_HUGE_PAPX)
        {
            return Err(corrupted("the paragraph stores a huge PAPX"));
        }
        let body = &grpprl[2..];
        let _ = papx_sprms(body)?;
        let table_state = papx_table_state(&grpprl)?;
        let mark_fc = self.fc_at_cp(mark)?;
        let width = containing_piece(&self.pieces, mark)
            .map_or(2, |piece| if piece.unicode { 2 } else { 1 });
        let mark_end = mark_fc
            .checked_add(width)
            .ok_or_else(|| corrupted("paragraph mark FC overflow"))?;
        // The end of the previous paragraph's mark, when it is physically
        // inside the same PAPX run.
        let previous_end = match start.checked_sub(1) {
            Some(previous) => {
                let fc = self.fc_at_cp(previous)?;
                let width = containing_piece(&self.pieces, previous)
                    .map_or(2, |piece| if piece.unicode { 2 } else { 1 });
                fc.checked_add(width)
            },
            None => None,
        };
        let mut candidate = self.clone();
        split_papx(&mut candidate.papx, mark_end);
        if let Some(previous_end) = previous_end
            && previous_end <= mark_fc
        {
            split_papx(&mut candidate.papx, previous_end);
        }
        let run = candidate
            .papx
            .iter_mut()
            .find(|run| run.start <= mark_fc && mark_fc < run.end)
            .ok_or_else(|| corrupted("paragraph mark has no PAPX run"))?;
        run.replace_grpprl(grpprl, table_state);
        candidate.rewrite_papx()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Insert `runs` (text units with the CHPX of each) at main-story CP
    /// `at` and the field markers `fields` (CP, `FLD` bytes), shifting every
    /// modeled CP table, rebuilding the CHPX FKPs and the CLX.
    pub(crate) fn insert_runs_with_fields(
        &mut self,
        at: u32,
        runs: &[(Vec<u16>, Vec<u8>)],
        fields: &[(u32, [u8; 2])],
    ) -> Result<()> {
        if at >= self.main_ccp {
            return Err(corrupted("insertion is outside the main story"));
        }
        if !self.unmodeled_cp_tables.is_empty() {
            return Err(corrupted("insertion has unmodeled CP-indexed dependencies"));
        }
        let added = runs.iter().map(|(units, _)| units.len()).sum::<usize>();
        if added == 0 || added > MAX_TEXT_UNITS {
            return Err(corrupted(
                "inserted text is empty or exceeds the text limit",
            ));
        }
        let added =
            u32::try_from(added).map_err(|_error| corrupted("inserted text exceeds u32"))?;
        let saved_selection = self.prepared_saved_selection_splice(at, at, added)?;
        let mut candidate = self.clone();
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        let fc = u32::try_from(fc).map_err(|_error| corrupted("inserted text FC exceeds u32"))?;
        let mut cursor = fc;
        for (units, grpprl) in runs {
            if grpprl.len() > 255 {
                return Err(corrupted("inserted CHPX exceeds one-byte FKP limit"));
            }
            for unit in units {
                candidate.word.extend_from_slice(&unit.to_le_bytes());
            }
            let length = u32::try_from(units.len() * 2)
                .map_err(|_error| corrupted("inserted run exceeds u32"))?;
            if length == 0 {
                continue;
            }
            candidate.chpx.push(FcRun {
                start: cursor,
                end: cursor + length,
                grpprl: grpprl.clone(),
            });
            cursor += length;
        }
        insert_piece(&mut candidate.pieces, at, added, fc)?;
        candidate.shift_cp_tables(at, 0, added)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_add(added)
            .ok_or_else(|| corrupted("main story CP insertion overflow"))?;
        if !fields.is_empty() {
            let terminal = candidate.main_ccp;
            let table = candidate.cp_table_mut(PLCFFLD_MOM, 2, terminal);
            for (cp, record) in fields {
                insert_entry(table, *cp, record)?;
            }
        }
        candidate.rewrite_chpx()?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Remove the main-story text `start..end` and the `PlcffldMom` markers
    /// inside it, shifting every other modeled CP table. Any other modeled
    /// position strictly inside the range is refused.
    pub(crate) fn delete_range_with_fields(&mut self, start: u32, end: u32) -> Result<()> {
        if start >= end || end > self.main_ccp {
            return Err(corrupted("deletion is outside the main story"));
        }
        if !self.unmodeled_cp_tables.is_empty() {
            return Err(corrupted("deletion has unmodeled CP-indexed dependencies"));
        }
        let units = read_units(&self.word, &self.pieces, start, end)?;
        if units.iter().any(|unit| matches!(*unit, 0x0D | 0x07)) {
            return Err(corrupted("deletion would remove a paragraph or cell mark"));
        }
        if self
            .cp_tables
            .iter()
            .filter(|table| table.index != PLCFFLD_MOM)
            .any(|table| table.cps.iter().any(|cp| *cp > start && *cp < end))
        {
            return Err(corrupted("deletion would collapse an anchored range"));
        }
        let removed = end - start;
        let saved_selection = self.prepared_saved_selection_splice(start, end, 0)?;
        let mut candidate = self.clone();
        if let Some(table) = candidate
            .cp_tables
            .iter_mut()
            .find(|table| table.index == PLCFFLD_MOM)
        {
            remove_entries(table, start, end, 2)?;
        }
        delete_piece_range(&mut candidate.pieces, start, end)?;
        candidate.shift_cp_tables(start, removed, 0)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(removed)
            .ok_or_else(|| corrupted("main story CP deletion underflow"))?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// The `PlcffldMom` markers, as (CP, `FLD` bytes) in CP order.
    pub(crate) fn main_field_markers(&self) -> Vec<(u32, [u8; 2])> {
        self.cp_tables
            .iter()
            .find(|table| table.index == PLCFFLD_MOM)
            .map(|table| {
                table
                    .cps
                    .iter()
                    .zip(table.records.chunks_exact(2))
                    .map(|(cp, record)| (*cp, [record[0], record[1]]))
                    .collect()
            })
            .unwrap_or_default()
    }

    /// The standard bookmarks, in start order.
    pub(crate) fn bookmark_ranges(&self) -> Result<Vec<BookmarkRange>> {
        let names = self.bookmark_names()?;
        let starts = self.cp_tables.iter().find(|table| table.index == PLCFBKF);
        let ends = self.cp_tables.iter().find(|table| table.index == PLCFBKL);
        let (Some(starts), Some(ends)) = (starts, ends) else {
            if names.is_empty() {
                return Ok(Vec::new());
            }
            return Err(corrupted("bookmark names have no position tables"));
        };
        let count = starts.cps.len().saturating_sub(1);
        if names.len() != count || starts.records.len() != count * 4 {
            return Err(corrupted("bookmark tables disagree in length"));
        }
        let mut output = Vec::with_capacity(count);
        for (index, name) in names.into_iter().enumerate() {
            let ibkl = usize::from(u16_at(&starts.records, index * 4)?);
            let end = *ends
                .cps
                .get(ibkl)
                .filter(|_| ibkl < ends.cps.len().saturating_sub(1))
                .ok_or_else(|| corrupted("bookmark end index is out of range"))?;
            output.push(BookmarkRange {
                name,
                start: starts.cps[index],
                end,
            });
        }
        Ok(output)
    }

    /// `SttbfBkmk.cbExtra`, 0 when the table is absent.
    fn bookmark_extra_bytes(&self) -> Result<u16> {
        let (offset, length) = fib_pair(&self.word, STTBF_BKMK)?;
        if length == 0 {
            return Ok(0);
        }
        u16_at(slice(&self.table, offset, length, "SttbfBkmk")?, 4)
    }

    fn bookmark_names(&self) -> Result<Vec<String>> {
        let (offset, length) = fib_pair(&self.word, STTBF_BKMK)?;
        if length == 0 {
            return Ok(Vec::new());
        }
        let data = slice(&self.table, offset, length, "SttbfBkmk")?;
        if u16_at(data, 0)? != 0xFFFF {
            return Err(corrupted("SttbfBkmk is not an extended string table"));
        }
        let count = usize::from(u16_at(data, 2)?);
        let extra = usize::from(u16_at(data, 4)?);
        let mut cursor = 6usize;
        let mut names = Vec::with_capacity(count);
        for _ in 0..count {
            let length = usize::from(u16_at(data, cursor)?);
            cursor += 2;
            let bytes = data
                .get(cursor..cursor + length * 2)
                .ok_or_else(|| corrupted("SttbfBkmk entry exceeds its table"))?;
            let units = bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect::<Vec<_>>();
            names.push(
                String::from_utf16(&units)
                    .map_err(|_error| corrupted("bookmark name is invalid"))?,
            );
            cursor += length * 2 + extra;
        }
        Ok(names)
    }

    /// Replace the standard bookmarks with `bookmarks`, rewriting
    /// `SttbfBkmk`, `PlcfBkf` and `PlcfBkl`. Every bookmark keeps its
    /// `FBKF.bkc` when its name survives; a new one is a plain text range.
    pub(crate) fn set_bookmarks(&mut self, bookmarks: &[BookmarkRange]) -> Result<()> {
        if self.bookmark_extra_bytes()? != 0 {
            return Err(corrupted(
                "SttbfBkmk carries per-name extra data the bookmark editor does not model",
            ));
        }
        if bookmarks.len() > 0x3FFB {
            return Err(corrupted("bookmark table exceeds 0x3FFB entries"));
        }
        let previous = self.bookmark_ranges()?;
        let starts = self.cp_tables.iter().find(|table| table.index == PLCFBKF);
        let bkc_of = |name: &str| -> Result<u16> {
            let Some(starts) = starts else {
                return Ok(0);
            };
            match previous.iter().position(|bookmark| bookmark.name == name) {
                Some(index) => u16_at(&starts.records, index * 4 + 2),
                None => Ok(0),
            }
        };
        let mut start_order = (0..bookmarks.len()).collect::<Vec<_>>();
        start_order.sort_by_key(|index| (bookmarks[*index].start, *index));
        let mut end_order = (0..bookmarks.len()).collect::<Vec<_>>();
        end_order.sort_by_key(|index| (bookmarks[*index].end, *index));
        let mut end_index = vec![0u16; bookmarks.len()];
        for (position, index) in end_order.iter().enumerate() {
            end_index[*index] =
                u16::try_from(position).map_err(|_error| corrupted("bookmark index overflow"))?;
        }
        let terminal = |table: Option<&CpTable>| {
            table
                .and_then(|table| table.cps.last().copied())
                .unwrap_or(self.main_ccp + 1)
        };
        let start_terminal = terminal(starts);
        let end_terminal = terminal(self.cp_tables.iter().find(|table| table.index == PLCFBKL));
        let mut names = Vec::new();
        names.extend_from_slice(&0xFFFFu16.to_le_bytes());
        names.extend_from_slice(
            &u16::try_from(bookmarks.len())
                .map_err(|_error| corrupted("bookmark count overflow"))?
                .to_le_bytes(),
        );
        names.extend_from_slice(&0u16.to_le_bytes());
        let mut start_cps = Vec::with_capacity(bookmarks.len() + 1);
        let mut records = Vec::with_capacity(bookmarks.len() * 4);
        for index in &start_order {
            let bookmark = &bookmarks[*index];
            let units = bookmark.name.encode_utf16().collect::<Vec<_>>();
            if units.is_empty() || units.len() >= 40 || bookmark.start > bookmark.end {
                return Err(corrupted("bookmark name or range is invalid"));
            }
            names.extend_from_slice(
                &u16::try_from(units.len())
                    .map_err(|_error| corrupted("bookmark name overflow"))?
                    .to_le_bytes(),
            );
            names.extend(units.iter().flat_map(|unit| unit.to_le_bytes()));
            start_cps.push(bookmark.start);
            records.extend_from_slice(&end_index[*index].to_le_bytes());
            records.extend_from_slice(&bkc_of(&bookmark.name)?.to_le_bytes());
        }
        start_cps.push(start_terminal.max(start_cps.last().copied().unwrap_or(0)));
        let mut end_cps = end_order
            .iter()
            .map(|index| bookmarks[*index].end)
            .collect::<Vec<_>>();
        end_cps.push(end_terminal.max(end_cps.last().copied().unwrap_or(0)));
        let mut candidate = self.clone();
        candidate
            .cp_tables
            .retain(|table| table.index != PLCFBKF && table.index != PLCFBKL);
        if bookmarks.is_empty() {
            super::put_fib_pair(&mut candidate.word, STTBF_BKMK, 0, 0)?;
            super::put_fib_pair(&mut candidate.word, PLCFBKF, 0, 0)?;
            super::put_fib_pair(&mut candidate.word, PLCFBKL, 0, 0)?;
        } else {
            append_table_block(
                &mut candidate.word,
                &mut candidate.table,
                STTBF_BKMK,
                &names,
            )?;
            candidate.cp_tables.push(CpTable {
                index: PLCFBKF,
                cps: start_cps,
                records,
            });
            candidate.cp_tables.push(CpTable {
                index: PLCFBKL,
                cps: end_cps,
                records: Vec::new(),
            });
        }
        candidate.append_clx_and_cp_tables()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// The sections, as (start CP, end CP, SEPX SPRMs), in order.
    pub(crate) fn sections(&self) -> Result<Vec<(u32, u32, Vec<u8>)>> {
        let Some(table) = self.cp_tables.iter().find(|table| table.index == PLCFSED) else {
            return Ok(Vec::new());
        };
        let count = table.cps.len().saturating_sub(1);
        let mut output = Vec::with_capacity(count);
        for index in 0..count {
            let fc = u32_at(&table.records, index * 12 + 2)?;
            let grpprl = if fc == u32::MAX {
                Vec::new()
            } else {
                let length = u32::from(u16_at(&self.word, fc as usize)?);
                slice(&self.word, fc + 2, length, "SEPX")?.to_vec()
            };
            output.push((table.cps[index], table.cps[index + 1], grpprl));
        }
        Ok(output)
    }

    /// Point section `index` at a new SEPX holding `grpprl`, appended to the
    /// `WordDocument` stream; the old SEPX bytes stay unreferenced.
    pub(crate) fn set_section_grpprl(&mut self, index: usize, grpprl: &[u8]) -> Result<()> {
        let _ = papx_sprms(grpprl)?;
        let length = u16::try_from(grpprl.len()).map_err(|_error| corrupted("SEPX too long"))?;
        let mut candidate = self.clone();
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        candidate.word.extend_from_slice(&length.to_le_bytes());
        candidate.word.extend_from_slice(grpprl);
        let fc = u32::try_from(fc).map_err(|_error| corrupted("SEPX FC exceeds u32"))?;
        let table = candidate
            .cp_tables
            .iter_mut()
            .find(|table| table.index == PLCFSED)
            .ok_or_else(|| corrupted("the document has no section table"))?;
        let record = table
            .records
            .get_mut(index * 12 + 2..index * 12 + 6)
            .ok_or_else(|| corrupted("section index is out of range"))?;
        record.copy_from_slice(&fc.to_le_bytes());
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    fn cp_table_mut(&mut self, index: usize, record_size: usize, terminal: u32) -> &mut CpTable {
        let _ = record_size;
        if let Some(position) = self.cp_tables.iter().position(|table| table.index == index) {
            return &mut self.cp_tables[position];
        }
        self.cp_tables.push(CpTable {
            index,
            cps: vec![terminal],
            records: Vec::new(),
        });
        let last = self.cp_tables.len() - 1;
        &mut self.cp_tables[last]
    }
}

/// Split the PAPX run containing `fc` (strictly inside it) in two.
fn split_papx(runs: &mut Vec<PapxRun>, fc: u32) {
    if let Some(index) = runs.iter().position(|run| run.start < fc && fc < run.end) {
        let mut right = runs[index].clone();
        right.start = fc;
        runs[index].end = fc;
        runs.insert(index + 1, right);
    }
}

/// Insert one marker into a PLCF, before the first entry at or after `cp`;
/// the terminal CP stays last.
fn insert_entry(table: &mut CpTable, cp: u32, record: &[u8]) -> Result<()> {
    let entries = table.cps.len().saturating_sub(1);
    let index = table.cps[..entries].partition_point(|existing| *existing < cp);
    if table.cps.last().is_some_and(|terminal| *terminal < cp) {
        return Err(corrupted("PLCF marker is past the table's terminal CP"));
    }
    table.cps.insert(index, cp);
    let offset = index * record.len();
    table.records.splice(offset..offset, record.iter().copied());
    Ok(())
}

/// Remove every entry of a PLCF whose CP lies in `start..end`.
fn remove_entries(table: &mut CpTable, start: u32, end: u32, record_size: usize) -> Result<()> {
    let entries = table.cps.len().saturating_sub(1);
    if table.records.len() != entries * record_size {
        return Err(corrupted("PLCF records disagree with its CPs"));
    }
    let mut index = entries;
    while index > 0 {
        index -= 1;
        if table.cps[index] >= start && table.cps[index] < end {
            table.cps.remove(index);
            let offset = index * record_size;
            table.records.drain(offset..offset + record_size);
        }
    }
    Ok(())
}
