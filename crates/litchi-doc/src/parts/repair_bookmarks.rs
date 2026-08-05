//! Repair-bookmark tables (`SttbfBkmkBPRepairs`, `PlcfbkfBPRepairs`, and
//! `PlcfbklBPRepairs`).
//!
//! When Word repairs the bookmark pairs of a document while loading it, it
//! records each repair as a bookmark whose description lives in
//! `SttbfBkmkBPRepairs` (MS-DOC 2.9.280) and whose range lives in the parallel
//! `Plcfbkf`/`Plcfbkl` pair (MS-DOC 2.8.10 and 2.8.12). The tables are parsed
//! as inert metadata: descriptions are never evaluated and no repair is ever
//! applied or reverted.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use std::collections::HashSet;

/// Table-pointer index of `fcSttbfbkmkBPRepairs`/`lcbSttbfbkmkBPRepairs`
/// (MS-DOC 2.5.8 FibRgFcLcb2002).
const STTBF_BKMK_BP_REPAIRS: usize = 123;
/// Table-pointer index of `fcPlcfbkfBPRepairs`/`lcbPlcfbkfBPRepairs`
/// (MS-DOC 2.5.8 FibRgFcLcb2002).
const PLCF_BKF_BP_REPAIRS: usize = 124;
/// Table-pointer index of `fcPlcfbklBPRepairs`/`lcbPlcfbklBPRepairs`
/// (MS-DOC 2.5.8 FibRgFcLcb2002).
const PLCF_BKL_BP_REPAIRS: usize = 125;

/// `fExtend` value of an extended STTB (MS-DOC 2.9.271).
const STTB_F_EXTEND: u16 = 0xFFFF;
/// Maximum number of repair bookmarks in a document (MS-DOC 2.9.280).
const MAX_REPAIR_BOOKMARKS: u16 = 0x7FF0;
/// Size of one `FBKF` data element in a `Plcfbkf` (MS-DOC 2.9.70).
const FBKF_SIZE: usize = 4;
/// Size cap mirroring the other STTB readers; far beyond any conforming table.
const MAX_TABLE_BYTES: usize = 16 * 1024 * 1024;

/// `BKC` flag bit `fPub`, which must be zero (MS-DOC 2.9.8).
const BKC_F_PUB: u16 = 0x0080;
/// `BKC` flag bit `fNative` (MS-DOC 2.9.8).
const BKC_F_NATIVE: u16 = 0x4000;
/// `BKC` flag bit `fCol` (MS-DOC 2.9.8).
const BKC_F_COL: u16 = 0x8000;
/// `BKC` mask of `itcFirst` (MS-DOC 2.9.8).
const BKC_ITC_FIRST_MASK: u16 = 0x007F;
/// `BKC` shift and mask of `itcLim` (MS-DOC 2.9.8).
const BKC_ITC_LIM_SHIFT: u16 = 8;
const BKC_ITC_LIM_MASK: u16 = 0x003F;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn decode_utf16(data: &[u8], context: &str) -> Result<String> {
    char::decode_utf16(
        data.chunks_exact(2)
            .map(|unit| u16::from_le_bytes([unit[0], unit[1]])),
    )
    .collect::<std::result::Result<String, _>>()
    .map_err(|error| corrupted(format!("invalid {context}: {error}")))
}

/// One repair bookmark: a description string and the repaired range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RepairBookmark {
    /// Description of the repair, stored parallel to the bookmark in
    /// `SttbfBkmkBPRepairs` (MS-DOC 2.9.280).
    pub description: String,
    /// Start CP of the repaired range.
    pub start: u32,
    /// CP of the first character following the repaired range.
    pub end: u32,
    /// Whether the bookmark is expected to survive a save as RTF, HTML, or
    /// XML (`BKC.fNative`, MS-DOC 2.9.8).
    pub is_native: bool,
    /// Table-column range, when the bookmark is confined to columns
    /// (`BKC.fCol`, MS-DOC 2.9.8).
    pub column_range: Option<(u8, u8)>,
}

/// The repair bookmarks of a document, in start-CP order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentRepairBookmarks {
    bookmarks: Vec<RepairBookmark>,
    /// Trailing CP of the start PLC, verbatim. It is ignored when reading per
    /// MS-DOC 2.8.10 and preserved only so serialization stays faithful.
    final_start_cp: u32,
    /// Trailing CP of the end PLC, preserved for the same reason.
    final_end_cp: u32,
}

impl DocumentRepairBookmarks {
    /// Create a repair-bookmark set, validating count, ordering, and ranges.
    ///
    /// `final_start_cp` and `final_end_cp` are the ignored trailing CPs of the
    /// two PLCs (MS-DOC 2.8.10).
    pub fn try_new(
        bookmarks: Vec<RepairBookmark>,
        final_start_cp: u32,
        final_end_cp: u32,
    ) -> Result<Self> {
        validate_bookmarks(&bookmarks)?;
        Ok(Self {
            bookmarks,
            final_start_cp,
            final_end_cp,
        })
    }

    /// Parse the three parallel repair-bookmark tables addressed by the FIB,
    /// or `None` when the document carries none of them.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentRepairBookmarks>> {
        let lengths = [
            STTBF_BKMK_BP_REPAIRS,
            PLCF_BKF_BP_REPAIRS,
            PLCF_BKL_BP_REPAIRS,
        ]
        .map(|index| fib.get_table_pointer(index).map_or(0, |(_, length)| length));
        if lengths.iter().all(|&length| length == 0) {
            return Ok(None);
        }
        if lengths.contains(&0) {
            return Err(corrupted(
                "the three parallel repair-bookmark tables must be present together",
            ));
        }

        let descriptions = parse_descriptions(required_slice(
            fib,
            table_stream,
            STTBF_BKMK_BP_REPAIRS,
            "SttbfBkmkBPRepairs",
        )?)?;
        let (starts, final_start_cp) = parse_start_plcf(required_slice(
            fib,
            table_stream,
            PLCF_BKF_BP_REPAIRS,
            "PlcfbkfBPRepairs",
        )?)?;
        let ends = parse_end_plcf(required_slice(
            fib,
            table_stream,
            PLCF_BKL_BP_REPAIRS,
            "PlcfbklBPRepairs",
        )?)?;
        let Some((&final_end_cp, bookmark_ends)) = ends.split_last() else {
            return Err(corrupted("PlcfbklBPRepairs is missing its trailing CP"));
        };
        // MS-DOC 2.5.8: the Plcfbkf and Plcfbkl MUST contain the same number
        // of data elements, and the Plcfbkf is parallel to the STTB.
        if descriptions.len() != starts.len() || starts.len() != bookmark_ends.len() {
            return Err(corrupted(
                "repair-bookmark description, start, and end table counts do not match",
            ));
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        validate_cps(
            starts.iter().map(|(cp, _)| *cp),
            document_end,
            "PlcfbkfBPRepairs",
        )?;
        validate_cps(
            bookmark_ends.iter().copied(),
            document_end,
            "PlcfbklBPRepairs",
        )?;

        let mut used_end_indexes = HashSet::with_capacity(starts.len());
        let mut bookmarks = Vec::with_capacity(starts.len());
        for ((start, fbkf), description) in starts.iter().zip(descriptions) {
            let end_index = usize::from(fbkf.end_index);
            if end_index >= bookmark_ends.len() || !used_end_indexes.insert(end_index) {
                return Err(corrupted(
                    "repair-bookmark ibkl values must be unique and in range",
                ));
            }
            let end = bookmark_ends[end_index];
            if *start > end {
                return Err(corrupted("repair-bookmark start CP exceeds its end CP"));
            }
            bookmarks.push(RepairBookmark {
                description,
                start: *start,
                end,
                is_native: fbkf.is_native,
                column_range: fbkf.column_range,
            });
        }

        Ok(Some(Self {
            bookmarks,
            final_start_cp,
            final_end_cp,
        }))
    }

    /// All repair bookmarks in start-CP order.
    pub fn bookmarks(&self) -> &[RepairBookmark] {
        &self.bookmarks
    }

    /// Serialize the three parallel tables deterministically, in
    /// `SttbfBkmkBPRepairs`, `PlcfbkfBPRepairs`, `PlcfbklBPRepairs` order.
    ///
    /// End CPs are written in ascending order, with each start's `ibkl`
    /// selecting its end, matching the layout Word produces.
    pub fn to_bytes(&self) -> Result<(Vec<u8>, Vec<u8>, Vec<u8>)> {
        validate_bookmarks(&self.bookmarks)?;

        let mut descriptions = Vec::new();
        descriptions.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        descriptions.extend_from_slice(&(self.bookmarks.len() as u16).to_le_bytes());
        descriptions.extend_from_slice(&0u16.to_le_bytes());
        for bookmark in &self.bookmarks {
            let units: Vec<u16> = bookmark.description.encode_utf16().collect();
            descriptions.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for unit in units {
                descriptions.extend_from_slice(&unit.to_le_bytes());
            }
        }

        // Sort the end CPs ascending and give every start the index of its
        // end in that order. The sort is stable, so equal ends keep the
        // one-to-one correspondence.
        let mut end_order: Vec<usize> = (0..self.bookmarks.len()).collect();
        end_order.sort_by_key(|&index| self.bookmarks[index].end);
        let mut end_index_of = vec![0u16; self.bookmarks.len()];
        for (position, &index) in end_order.iter().enumerate() {
            end_index_of[index] = position as u16;
        }

        let mut starts = Vec::new();
        for bookmark in &self.bookmarks {
            starts.extend_from_slice(&bookmark.start.to_le_bytes());
        }
        starts.extend_from_slice(&self.final_start_cp.to_le_bytes());
        for (index, bookmark) in self.bookmarks.iter().enumerate() {
            starts.extend_from_slice(&end_index_of[index].to_le_bytes());
            let column_bits = match bookmark.column_range {
                Some((first, limit)) => {
                    BKC_F_COL | u16::from(first) | (u16::from(limit) << BKC_ITC_LIM_SHIFT)
                },
                None => 0,
            };
            let native_bit = if bookmark.is_native { BKC_F_NATIVE } else { 0 };
            starts.extend_from_slice(&(column_bits | native_bit).to_le_bytes());
        }

        let mut ends = Vec::new();
        for index in end_order {
            ends.extend_from_slice(&self.bookmarks[index].end.to_le_bytes());
        }
        ends.extend_from_slice(&self.final_end_cp.to_le_bytes());

        Ok((descriptions, starts, ends))
    }
}

/// Per-entry data decoded from one `FBKF` (MS-DOC 2.9.70).
#[derive(Debug)]
struct StartData {
    /// Zero-based index of the matching entry in the end PLC (`ibkl`).
    end_index: u16,
    /// Whether the bookmark is expected to survive a save as RTF/HTML/XML.
    is_native: bool,
    /// Table-column range, when the bookmark is confined to columns.
    column_range: Option<(u8, u8)>,
}

/// Validate the count cap, start-CP ordering, and per-bookmark ranges shared
/// by the parser and the serializers.
fn validate_bookmarks(bookmarks: &[RepairBookmark]) -> Result<()> {
    if bookmarks.len() > usize::from(MAX_REPAIR_BOOKMARKS) {
        return Err(corrupted("repair-bookmark count exceeds 0x7FF0"));
    }
    if bookmarks
        .iter()
        .any(|bookmark| bookmark.start > bookmark.end)
        || bookmarks
            .windows(2)
            .any(|pair| pair[0].start > pair[1].start)
    {
        return Err(corrupted(
            "repair-bookmark start CPs must be monotonic and not exceed their ends",
        ));
    }
    Ok(())
}

/// Parse `SttbfBkmkBPRepairs` (MS-DOC 2.9.280): an extended STTB of
/// description strings without extra data.
fn parse_descriptions(data: &[u8]) -> Result<Vec<String>> {
    if data.len() > MAX_TABLE_BYTES {
        return Err(corrupted("SttbfBkmkBPRepairs exceeds the table size cap"));
    }
    if data.len() < 6
        || read_u16(data, 0, "SttbfBkmkBPRepairs fExtend")? != STTB_F_EXTEND
        || read_u16(data, 4, "SttbfBkmkBPRepairs cbExtra")? != 0
    {
        return Err(corrupted("SttbfBkmkBPRepairs has an invalid header"));
    }
    let count = read_u16(data, 2, "SttbfBkmkBPRepairs cData")?;
    if count > MAX_REPAIR_BOOKMARKS {
        return Err(corrupted("SttbfBkmkBPRepairs contains too many entries"));
    }
    let mut descriptions = Vec::with_capacity(usize::from(count));
    let mut offset = 6usize;
    for index in 0..usize::from(count) {
        let units = usize::from(read_u16(
            data,
            offset,
            &format!("SttbfBkmkBPRepairs string {index} length"),
        )?);
        let start = offset
            .checked_add(2)
            .ok_or_else(|| corrupted("SttbfBkmkBPRepairs string offset overflows"))?;
        let end = start
            .checked_add(
                units
                    .checked_mul(2)
                    .ok_or_else(|| corrupted("SttbfBkmkBPRepairs string size overflows"))?,
            )
            .ok_or_else(|| corrupted("SttbfBkmkBPRepairs string range overflows"))?;
        let bytes = data
            .get(start..end)
            .ok_or_else(|| corrupted(format!("SttbfBkmkBPRepairs string {index} is truncated")))?;
        descriptions.push(decode_utf16(
            bytes,
            &format!("SttbfBkmkBPRepairs string {index}"),
        )?);
        offset = end;
    }
    if offset != data.len() {
        return Err(corrupted("SttbfBkmkBPRepairs has trailing bytes"));
    }
    Ok(descriptions)
}

/// Parse `PlcfbkfBPRepairs`: a `Plcfbkf` whose 4-byte data elements are
/// `FBKF` structures (MS-DOC 2.8.10 and 2.9.70). Returns one entry per
/// bookmark plus the ignored trailing CP.
fn parse_start_plcf(data: &[u8]) -> Result<(Vec<(u32, StartData)>, u32)> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4 + FBKF_SIZE) {
        return Err(corrupted("PlcfbkfBPRepairs has an invalid byte length"));
    }
    let count = (data.len() - 4) / (4 + FBKF_SIZE);
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfbkfBPRepairs position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * FBKF_SIZE;
        let bkc = read_u16(data, property + 2, "repair-bookmark BKC")?;
        if bkc & BKC_F_PUB != 0 {
            return Err(corrupted("repair-bookmark BKC fPub must be zero"));
        }
        let column_range = if bkc & BKC_F_COL != 0 {
            let first = (bkc & BKC_ITC_FIRST_MASK) as u8;
            let limit = ((bkc >> BKC_ITC_LIM_SHIFT) & BKC_ITC_LIM_MASK) as u8;
            if first >= limit {
                return Err(corrupted(
                    "repair-bookmark BKC column range is empty or reversed",
                ));
            }
            Some((first, limit))
        } else {
            None
        };
        values.push((
            read_u32(data, index * 4, "repair-bookmark start CP")?,
            StartData {
                end_index: read_u16(data, property, "repair-bookmark ibkl")?,
                is_native: bkc & BKC_F_NATIVE != 0,
                column_range,
            },
        ));
    }
    let final_cp = read_u32(data, count * 4, "repair-bookmark trailing CP")?;
    Ok((values, final_cp))
}

/// Parse `PlcfbklBPRepairs`: a `Plcfbkl`, a PLC of end CPs without data
/// elements (MS-DOC 2.8.12). The final entry is the ignored trailing CP.
fn parse_end_plcf(data: &[u8]) -> Result<Vec<u32>> {
    if data.len() < 4 || !data.len().is_multiple_of(4) {
        return Err(corrupted("PlcfbklBPRepairs has an invalid byte length"));
    }
    let mut ends = Vec::with_capacity(data.len() / 4);
    for offset in (0..data.len()).step_by(4) {
        ends.push(read_u32(data, offset, "repair-bookmark end CP")?);
    }
    Ok(ends)
}

/// Every CP except the ignored trailing one must lie within the document
/// parts and be monotonic (MS-DOC 2.8.10).
fn validate_cps(cps: impl Iterator<Item = u32>, document_end: u32, name: &str) -> Result<()> {
    let mut previous = None;
    for cp in cps {
        if cp > document_end || previous.is_some_and(|value| value > cp) {
            return Err(corrupted(format!(
                "{name} contains out-of-range or non-monotonic CPs"
            )));
        }
        previous = Some(cp);
    }
    Ok(())
}

fn required_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let (offset, length) = fib
        .get_table_pointer(index)
        .filter(|(_, length)| *length != 0)
        .ok_or_else(|| corrupted(format!("{name} is missing")))?;
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIB_POINTERS: usize = 145;

    fn fib_with_pointers(pairs: &[(usize, u32, u32)]) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + FIB_POINTERS * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        // ccpText: ten characters in the main document part.
        data[76..80].copy_from_slice(&10u32.to_le_bytes());
        data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());
        for (index, offset, length) in pairs {
            let pointer = 154 + index * 8;
            data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        }
        FileInformationBlock::parse(&data).unwrap()
    }

    fn description_sttb(descriptions: &[&str]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(descriptions.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for description in descriptions {
            let units: Vec<u16> = description.encode_utf16().collect();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for unit in units {
                data.extend_from_slice(&unit.to_le_bytes());
            }
        }
        data
    }

    /// Start PLC for two overlapping bookmarks: starts at CP 1 and 2,
    /// trailing CP 11. The first bookmark is native and ends at end index 1;
    /// the second spans columns 1..3 and ends at end index 0.
    fn start_plcf() -> Vec<u8> {
        let mut data = Vec::new();
        for cp in [1u32, 2, 11] {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&1u16.to_le_bytes()); // ibkl
        data.extend_from_slice(&BKC_F_NATIVE.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes()); // ibkl
        data.extend_from_slice(&(BKC_F_COL | 0x0301).to_le_bytes());
        data
    }

    /// End PLC for two bookmarks: ends at CP 5 and 8, trailing CP 11.
    fn end_plcf() -> Vec<u8> {
        let mut data = Vec::new();
        for cp in [5u32, 8, 11] {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data
    }

    fn sample_bookmarks() -> Vec<RepairBookmark> {
        vec![
            RepairBookmark {
                description: "Repaired bookmark order".to_string(),
                start: 1,
                end: 8,
                is_native: true,
                column_range: None,
            },
            RepairBookmark {
                description: "Repaired table bookmark".to_string(),
                start: 2,
                end: 5,
                is_native: false,
                column_range: Some((1, 3)),
            },
        ]
    }

    /// Build a FIB plus table stream holding two repair bookmarks, returning
    /// the offset of the start PLC inside the stream.
    fn fixture() -> (FileInformationBlock, Vec<u8>, usize) {
        let descriptions =
            description_sttb(&["Repaired bookmark order", "Repaired table bookmark"]);
        let starts = start_plcf();
        let ends = end_plcf();

        let mut table_stream = Vec::new();
        let descriptions_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&descriptions);
        let starts_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&starts);
        let ends_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&ends);

        let fib = fib_with_pointers(&[
            (
                STTBF_BKMK_BP_REPAIRS,
                descriptions_offset,
                descriptions.len() as u32,
            ),
            (PLCF_BKF_BP_REPAIRS, starts_offset, starts.len() as u32),
            (PLCF_BKL_BP_REPAIRS, ends_offset, ends.len() as u32),
        ]);
        (fib, table_stream, descriptions.len())
    }

    #[test]
    fn parses_repair_bookmarks_and_round_trips() {
        let (fib, table_stream, starts_offset) = fixture();
        let parsed = DocumentRepairBookmarks::parse(&fib, &table_stream)
            .unwrap()
            .expect("repair bookmarks present");
        assert_eq!(parsed.bookmarks(), sample_bookmarks().as_slice());

        let (descriptions, starts, ends) = parsed.to_bytes().unwrap();
        let ends_offset = starts_offset + start_plcf().len();
        assert_eq!(descriptions, &table_stream[..starts_offset]);
        assert_eq!(starts, &table_stream[starts_offset..ends_offset]);
        assert_eq!(ends, &table_stream[ends_offset..]);
    }

    #[test]
    fn constructor_round_trips_through_parser() {
        let tables = DocumentRepairBookmarks::try_new(sample_bookmarks(), 11, 11).unwrap();
        let (descriptions, starts, ends) = tables.to_bytes().unwrap();
        let mut table_stream = Vec::new();
        table_stream.extend_from_slice(&descriptions);
        let starts_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&starts);
        let ends_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&ends);
        let fib = fib_with_pointers(&[
            (STTBF_BKMK_BP_REPAIRS, 0, descriptions.len() as u32),
            (PLCF_BKF_BP_REPAIRS, starts_offset, starts.len() as u32),
            (PLCF_BKL_BP_REPAIRS, ends_offset, ends.len() as u32),
        ]);
        let parsed = DocumentRepairBookmarks::parse(&fib, &table_stream)
            .unwrap()
            .expect("repair bookmarks present");
        assert_eq!(parsed, tables);
    }

    #[test]
    fn absent_tables_yield_none() {
        let fib = fib_with_pointers(&[]);
        assert!(DocumentRepairBookmarks::parse(&fib, &[]).unwrap().is_none());
    }

    #[test]
    fn rejects_partial_table_sets() {
        let (fib, table_stream, _) = fixture();
        let mut fib_data = fib.raw_data().to_vec();
        let pointer = 154 + PLCF_BKL_BP_REPAIRS * 8;
        fib_data[pointer + 4..pointer + 8].copy_from_slice(&0u32.to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentRepairBookmarks::parse(&fib, &table_stream).is_err());
    }

    #[test]
    fn rejects_mismatched_table_counts() {
        let starts = start_plcf();

        // One description against two PLC entries.
        let descriptions = description_sttb(&["Repaired bookmark order"]);
        let ends = end_plcf();
        let mut table_stream = descriptions.clone();
        let starts_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&starts);
        let ends_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&ends);
        let fib = fib_with_pointers(&[
            (STTBF_BKMK_BP_REPAIRS, 0, descriptions.len() as u32),
            (PLCF_BKF_BP_REPAIRS, starts_offset, starts.len() as u32),
            (PLCF_BKL_BP_REPAIRS, ends_offset, ends.len() as u32),
        ]);
        assert!(DocumentRepairBookmarks::parse(&fib, &table_stream).is_err());

        // One end against two starts.
        let descriptions =
            description_sttb(&["Repaired bookmark order", "Repaired table bookmark"]);
        let short_ends = end_plcf()[..8].to_vec();
        let mut table_stream = descriptions.clone();
        let starts_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&starts);
        let ends_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&short_ends);
        let fib = fib_with_pointers(&[
            (STTBF_BKMK_BP_REPAIRS, 0, descriptions.len() as u32),
            (PLCF_BKF_BP_REPAIRS, starts_offset, starts.len() as u32),
            (PLCF_BKL_BP_REPAIRS, ends_offset, short_ends.len() as u32),
        ]);
        assert!(DocumentRepairBookmarks::parse(&fib, &table_stream).is_err());
    }

    #[test]
    fn rejects_malformed_description_table() {
        assert!(parse_descriptions(&[]).is_err());
        // Non-extended STTB header.
        let mut bytes = description_sttb(&["repair"]);
        bytes[0] = 0;
        bytes[1] = 0;
        assert!(parse_descriptions(&bytes).is_err());
        // Nonzero cbExtra.
        let mut bytes = description_sttb(&["repair"]);
        bytes[4] = 2;
        assert!(parse_descriptions(&bytes).is_err());
        // Declared count beyond the payload.
        let mut bytes = description_sttb(&["repair"]);
        bytes[2] = 2;
        assert!(parse_descriptions(&bytes).is_err());
        // Truncated string.
        let mut bytes = description_sttb(&["repair"]);
        bytes.truncate(bytes.len() - 2);
        assert!(parse_descriptions(&bytes).is_err());
        // Trailing bytes.
        let mut bytes = description_sttb(&["repair"]);
        bytes.push(0);
        assert!(parse_descriptions(&bytes).is_err());
        // Unpaired surrogate.
        let mut bytes = description_sttb(&["repair"]);
        bytes[8] = 0x00;
        bytes[9] = 0xD8;
        assert!(parse_descriptions(&bytes).is_err());
    }

    #[test]
    fn rejects_malformed_plcs() {
        // Invalid byte lengths.
        assert!(parse_start_plcf(&[0u8; 7]).is_err());
        assert!(parse_end_plcf(&[0u8; 6]).is_err());

        let (fib, table_stream, starts_offset) = fixture();

        // fPub set in the first FBKF.
        let mut public = table_stream.clone();
        public[starts_offset + 14] |= 0x80;
        assert!(DocumentRepairBookmarks::parse(&fib, &public).is_err());

        // Reversed column range in the second FBKF.
        let mut reversed = table_stream.clone();
        reversed[starts_offset + 18..starts_offset + 20].copy_from_slice(&0x8102u16.to_le_bytes());
        assert!(DocumentRepairBookmarks::parse(&fib, &reversed).is_err());

        // Duplicate ibkl values: both starts reference end index 1.
        let mut duplicate = table_stream.clone();
        duplicate[starts_offset + 16..starts_offset + 18].copy_from_slice(&1u16.to_le_bytes());
        assert!(DocumentRepairBookmarks::parse(&fib, &duplicate).is_err());

        // Out-of-range ibkl.
        let mut out_of_range = table_stream.clone();
        out_of_range[starts_offset + 12..starts_offset + 14].copy_from_slice(&7u16.to_le_bytes());
        assert!(DocumentRepairBookmarks::parse(&fib, &out_of_range).is_err());

        // Start CP beyond the document parts.
        let mut beyond = table_stream.clone();
        beyond[starts_offset..starts_offset + 4].copy_from_slice(&11u32.to_le_bytes());
        assert!(DocumentRepairBookmarks::parse(&fib, &beyond).is_err());
    }

    #[test]
    fn rejects_start_beyond_end() {
        // Monotonic starts and ends, but the first start crosses its end.
        let descriptions = description_sttb(&["first", "second"]);
        let mut starts = Vec::new();
        for cp in [5u32, 6, 11] {
            starts.extend_from_slice(&cp.to_le_bytes());
        }
        starts.extend_from_slice(&0u16.to_le_bytes());
        starts.extend_from_slice(&0u16.to_le_bytes());
        starts.extend_from_slice(&1u16.to_le_bytes());
        starts.extend_from_slice(&0u16.to_le_bytes());
        let mut ends = Vec::new();
        for cp in [4u32, 9, 11] {
            ends.extend_from_slice(&cp.to_le_bytes());
        }
        let mut table_stream = descriptions.clone();
        let starts_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&starts);
        let ends_offset = table_stream.len() as u32;
        table_stream.extend_from_slice(&ends);
        let fib = fib_with_pointers(&[
            (STTBF_BKMK_BP_REPAIRS, 0, descriptions.len() as u32),
            (PLCF_BKF_BP_REPAIRS, starts_offset, starts.len() as u32),
            (PLCF_BKL_BP_REPAIRS, ends_offset, ends.len() as u32),
        ]);
        assert!(DocumentRepairBookmarks::parse(&fib, &table_stream).is_err());
    }

    #[test]
    fn rejects_constructor_violations() {
        let bookmark = |start, end| RepairBookmark {
            description: String::new(),
            start,
            end,
            is_native: false,
            column_range: None,
        };
        // Count cap.
        let many = vec![bookmark(0, 0); usize::from(MAX_REPAIR_BOOKMARKS) + 1];
        assert!(DocumentRepairBookmarks::try_new(many, 0, 0).is_err());
        // Start beyond end.
        assert!(DocumentRepairBookmarks::try_new(vec![bookmark(5, 4)], 0, 0).is_err());
        // Non-monotonic starts.
        assert!(
            DocumentRepairBookmarks::try_new(vec![bookmark(5, 6), bookmark(4, 6)], 0, 0).is_err()
        );
    }
}
