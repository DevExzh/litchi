//! Word format consistency-checker bookmark tables.
//!
//! These are the `SttbfBkmkFcc`, `PlcfBkfFcc`, and `PlcfBklFcc` tables
//! (MS-DOC 2.9.282, 2.8.10, and 2.9.64) that record the regions of text the
//! format consistency checker flagged, together with the kind of inconsistency
//! found in each region.
//!
//! All structures are parsed as inert metadata: no formatting is analyzed,
//! compared, or modified.

use super::fib::FileInformationBlock;
use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

/// Table-pointer index of `fcSttbfBkmkFcc`/`lcbSttbfBkmkFcc`.
const STTBF_BKMK_FCC: usize = 120;
/// Table-pointer index of `fcPlcfBkfFcc`/`lcbPlcfBkfFcc`.
const PLCF_BKF_FCC: usize = 121;
/// Table-pointer index of `fcPlcfBklFcc`/`lcbPlcfBklFcc`.
const PLCF_BKL_FCC: usize = 122;

/// `fExtend` value of an extended STTB.
const STTB_F_EXTEND: u16 = 0xFFFF;
/// Maximum number of format consistency-checker bookmarks in a document
/// (`SttbfBkmkFcc` `cData` limit, MS-DOC 2.9.282).
const MAX_MARKS: u16 = 0x7FF0;
/// `cchData` of every `SttbfBkmkFcc` string: one `DPCID` is ten UTF-16 code
/// units (MS-DOC 2.9.282).
const DPCID_CHARS: u16 = 0x000A;
/// Size of one `DPCID` structure in bytes (MS-DOC 2.9.64).
const DPCID_SIZE: usize = 20;

/// `DPCID` flag bit `fSquiggle`.
const FLAG_SQUIGGLE: u32 = 0x0000_0001;
/// `DPCID` flag bit `fIgnored`.
const FLAG_IGNORED: u32 = 0x0000_0002;
/// `DPCID` flag bit `fSquiggleChanged`.
const FLAG_SQUIGGLE_CHANGED: u32 = 0x0000_0004;

/// `FCCT` bit `fcctChp`.
const FCCT_CHP: u8 = 0x01;
/// `FCCT` bit `fcctPap`, which must be zero (MS-DOC 2.9.74).
const FCCT_PAP: u8 = 0x02;
/// `FCCT` bit `fcctTap`.
const FCCT_TAP: u8 = 0x04;
/// `FCCT` bit `fcctSep`.
const FCCT_SEP: u8 = 0x08;

/// `BKC` flag bit `fPub`, which must be zero.
const BKC_F_PUB: u16 = 0x0080;
/// `BKC` flag bit `fNative`.
const BKC_F_NATIVE: u16 = 0x4000;
/// `BKC` flag bit `fCol`.
const BKC_F_COL: u16 = 0x8000;
/// `BKC` mask of `itcFirst`.
const BKC_ITC_FIRST_MASK: u16 = 0x007F;
/// `BKC` shift and mask of `itcLim`.
const BKC_ITC_LIM_SHIFT: u16 = 8;
const BKC_ITC_LIM_MASK: u16 = 0x003F;

/// The kind of formatting the format consistency checker flagged in a region
/// of text (`IDPCI`, MS-DOC 2.9.120).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormatConsistencyKind {
    /// Character formatting is inconsistent with the rest of the document.
    CharacterFormatting,
    /// Character style is identical to a character style used elsewhere.
    CharacterStyle,
    /// Paragraph formatting is inconsistent with the rest of the document.
    ParagraphFormatting,
    /// Paragraph style is identical to a paragraph style used elsewhere.
    ParagraphStyle,
    /// Numbered or bulleted list item formatting is inconsistent.
    ListFormatting,
    /// List style is identical to a list style used elsewhere.
    ListStyle,
    /// Table style is identical to a table style used elsewhere.
    TableStyle,
    /// Characters were changed while revision marking was on.
    RevisedCharacters,
    /// Paragraphs were changed while revision marking was on.
    RevisedParagraphs,
    /// Tables were changed while revision marking was on.
    RevisedTables,
    /// Sections were changed while revision marking was on.
    RevisedSections,
    /// An inline picture was combined with an identical picture elsewhere.
    CombinedImage,
}

/// Property categories flagged for one mark (`FCCT`, MS-DOC 2.9.74).
///
/// `fcctPap` must be zero in a valid document and is rejected while parsing,
/// so it has no representation here.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct FormatConsistencyProperties {
    /// Character properties were flagged as inconsistent.
    pub character: bool,
    /// Table properties were flagged as inconsistent.
    pub table: bool,
    /// Line-separation properties were flagged as inconsistent.
    pub line_separation: bool,
}

/// The `DPCID` stored parallel to one format consistency-checker bookmark
/// (MS-DOC 2.9.64).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormatConsistencyInfo {
    /// Unique identifier of the mark within the document.
    pub id: u32,
    /// The kind of formatting that was flagged.
    pub kind: FormatConsistencyKind,
    /// Property categories that were flagged.
    pub properties: FormatConsistencyProperties,
    /// Whether an application is expected to display a squiggle under the
    /// region.
    pub squiggle: bool,
    /// Whether the user asked that this flag be ignored.
    pub ignored: bool,
    /// Whether the squiggle under the region has recently changed.
    pub squiggle_changed: bool,
}

/// A validated format consistency-checker mark.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormatConsistencyMark {
    /// Start CP of the flagged region.
    pub start: u32,
    /// End CP of the flagged region.
    pub end: u32,
    /// Overlap depth recorded at the start of the region.
    pub start_depth: u16,
    /// Overlap depth recorded at the end of the region.
    pub end_depth: u16,
    /// Whether the bookmark was created by Word rather than a producer.
    pub is_native: bool,
    /// Table-column range, when the mark is confined to columns.
    pub column_range: Option<(u8, u8)>,
    /// The flagged-inconsistency data stored parallel to the bookmark.
    pub info: FormatConsistencyInfo,
}

/// The format consistency-checker marks of a document, in start-CP order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentFormatConsistencyMarks {
    marks: Vec<FormatConsistencyMark>,
}

impl DocumentFormatConsistencyMarks {
    /// Parse the three parallel format consistency-checker tables addressed by
    /// the FIB, or `None` when the document carries none of them.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentFormatConsistencyMarks>> {
        let lengths = [STTBF_BKMK_FCC, PLCF_BKF_FCC, PLCF_BKL_FCC]
            .map(|index| fib.get_table_pointer(index).map_or(0, |(_, length)| length));
        if lengths.iter().all(|&length| length == 0) {
            return Ok(None);
        }
        if lengths.contains(&0) {
            return Err(corrupted(
                "the three parallel format consistency-checker tables must be present together",
            ));
        }

        let infos = parse_infos(required_slice(
            fib,
            table_stream,
            STTBF_BKMK_FCC,
            "SttbfBkmkFcc",
        )?)?;
        let starts = parse_start_plcf(required_slice(
            fib,
            table_stream,
            PLCF_BKF_FCC,
            "PlcfBkfFcc",
        )?)?;
        let ends = parse_end_plcf(required_slice(
            fib,
            table_stream,
            PLCF_BKL_FCC,
            "PlcfBklFcc",
        )?)?;
        if infos.len() != starts.len() || starts.len() != ends.len() {
            return Err(corrupted(
                "format consistency-checker info, start, and end table counts do not match",
            ));
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        validate_positions(&starts, document_end, "PlcfBkfFcc")?;
        validate_positions(&ends, document_end, "PlcfBklFcc")?;

        let mut used_end_indexes = HashSet::with_capacity(starts.len());
        let mut marks = Vec::with_capacity(starts.len());
        for (start_index, ((start, start_data), info)) in starts.iter().zip(&infos).enumerate() {
            let end_index = usize::from(start_data.end_index);
            if end_index >= ends.len() || !used_end_indexes.insert(end_index) {
                return Err(corrupted(
                    "format consistency-checker start end-index values must be unique and in range",
                ));
            }
            let (end, end_data) = &ends[end_index];
            if usize::from(end_data.start_index) != start_index {
                return Err(corrupted(
                    "format consistency-checker start and end bookmark indexes are not reciprocal",
                ));
            }
            if start > end {
                return Err(corrupted(
                    "format consistency-checker start CP exceeds its end CP",
                ));
            }
            marks.push(FormatConsistencyMark {
                start: *start,
                end: *end,
                start_depth: start_data.depth,
                end_depth: end_data.depth,
                is_native: start_data.is_native,
                column_range: start_data.column_range,
                info: *info,
            });
        }

        Ok(Some(Self { marks }))
    }

    /// All marks in start-CP order.
    #[must_use]
    pub fn marks(&self) -> &[FormatConsistencyMark] {
        &self.marks
    }
}

#[derive(Debug)]
struct StartData {
    end_index: u16,
    depth: u16,
    is_native: bool,
    column_range: Option<(u8, u8)>,
}

#[derive(Debug)]
struct EndData {
    start_index: u16,
    depth: u16,
}

/// Parse `SttbfBkmkFcc` (MS-DOC 2.9.282): an extended STTB whose strings are
/// `DPCID` structures (MS-DOC 2.9.64).
fn parse_infos(data: &[u8]) -> Result<Vec<FormatConsistencyInfo>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbfBkmkFcc fExtend")? != STTB_F_EXTEND
        || read_u16(data, 4, "SttbfBkmkFcc cbExtra")? != 0
    {
        return Err(corrupted("SttbfBkmkFcc has an invalid header"));
    }
    let count = read_u16(data, 2, "SttbfBkmkFcc cData")?;
    if count > MAX_MARKS {
        return Err(corrupted("SttbfBkmkFcc contains too many entries"));
    }
    let count = usize::from(count);
    let entry_size = 2usize + DPCID_SIZE;
    let expected = 6usize
        .checked_add(
            count
                .checked_mul(entry_size)
                .ok_or_else(|| corrupted("SttbfBkmkFcc size overflows"))?,
        )
        .ok_or_else(|| corrupted("SttbfBkmkFcc size overflows"))?;
    if data.len() != expected {
        return Err(corrupted(
            "SttbfBkmkFcc byte length does not match its count",
        ));
    }
    let mut infos = Vec::with_capacity(count);
    let mut ids = HashSet::with_capacity(count);
    let mut offset = 6usize;
    for _ in 0..count {
        if read_u16(data, offset, "SttbfBkmkFcc cchData")? != DPCID_CHARS {
            return Err(corrupted(
                "SttbfBkmkFcc entries must contain ten UTF-16 code units",
            ));
        }
        let dpcid = offset + 2;
        let flags = read_u32(data, dpcid + 2, "DPCID flags")?;
        let kind = match read_u32(data, dpcid + 6, "DPCID idpci")? {
            0 => FormatConsistencyKind::CharacterFormatting,
            1 => FormatConsistencyKind::CharacterStyle,
            2 => FormatConsistencyKind::ParagraphFormatting,
            3 => FormatConsistencyKind::ParagraphStyle,
            4 => FormatConsistencyKind::ListFormatting,
            5 => FormatConsistencyKind::ListStyle,
            6 => FormatConsistencyKind::TableStyle,
            7 => FormatConsistencyKind::RevisedCharacters,
            8 => FormatConsistencyKind::RevisedParagraphs,
            9 => FormatConsistencyKind::RevisedTables,
            10 => FormatConsistencyKind::RevisedSections,
            11 => FormatConsistencyKind::CombinedImage,
            _ => return Err(corrupted("DPCID contains an invalid idpci")),
        };
        // DPCID `idata` is undefined and MUST be ignored.
        let fcct = read_u8(data, dpcid + 14, "DPCID fcct")?;
        if fcct & FCCT_PAP != 0 {
            return Err(corrupted("DPCID fcctPap must be zero"));
        }
        let id = read_u32(data, dpcid + 15, "DPCID id")?;
        if !ids.insert(id) {
            return Err(corrupted("DPCID ids must be unique"));
        }
        infos.push(FormatConsistencyInfo {
            id,
            kind,
            properties: FormatConsistencyProperties {
                character: fcct & FCCT_CHP != 0,
                table: fcct & FCCT_TAP != 0,
                line_separation: fcct & FCCT_SEP != 0,
            },
            squiggle: flags & FLAG_SQUIGGLE != 0,
            ignored: flags & FLAG_IGNORED != 0,
            squiggle_changed: flags & FLAG_SQUIGGLE_CHANGED != 0,
        });
        offset += entry_size;
    }
    Ok(infos)
}

/// Parse `PlcfBkfFcc`: a `Plcfbkfd` whose 6-byte data elements are `FBKFD`
/// structures (MS-DOC 2.9.71).
fn parse_start_plcf(data: &[u8]) -> Result<Vec<(u32, StartData)>> {
    let count = plcf_count(data, 6, "PlcfBkfFcc")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfFcc position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 6;
        let bkc = read_u16(data, property + 2, "format consistency-checker BKC")?;
        if bkc & BKC_F_PUB != 0 {
            return Err(corrupted(
                "format consistency-checker BKC fPub must be zero",
            ));
        }
        let column_range = if bkc & BKC_F_COL != 0 {
            let first = (bkc & BKC_ITC_FIRST_MASK) as u8;
            let limit = ((bkc >> BKC_ITC_LIM_SHIFT) & BKC_ITC_LIM_MASK) as u8;
            if first >= limit {
                return Err(corrupted(
                    "format consistency-checker BKC column range is empty or reversed",
                ));
            }
            Some((first, limit))
        } else {
            None
        };
        values.push((
            read_u32(data, index * 4, "format consistency-checker start CP")?,
            StartData {
                end_index: read_u16(data, property, "format consistency-checker ibkl")?,
                depth: read_u16(data, property + 4, "format consistency-checker start depth")?,
                is_native: bkc & BKC_F_NATIVE != 0,
                column_range,
            },
        ));
    }
    Ok(values)
}

/// Parse `PlcfBklFcc`: a `Plcfbkld` whose 4-byte data elements are `FBKLD`
/// structures (MS-DOC 2.9.72).
fn parse_end_plcf(data: &[u8]) -> Result<Vec<(u32, EndData)>> {
    let count = plcf_count(data, 4, "PlcfBklFcc")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBklFcc position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 4;
        values.push((
            read_u32(data, index * 4, "format consistency-checker end CP")?,
            EndData {
                start_index: read_u16(data, property, "format consistency-checker ibkf")?,
                depth: read_u16(data, property + 2, "format consistency-checker end depth")?,
            },
        ));
    }
    Ok(values)
}

fn plcf_count(data: &[u8], property_size: usize, name: &str) -> Result<usize> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4 + property_size) {
        return Err(corrupted(format!("{name} has an invalid byte length")));
    }
    Ok((data.len() - 4) / (4 + property_size))
}

fn validate_positions<T>(values: &[(u32, T)], document_end: u32, name: &str) -> Result<()> {
    if values.iter().any(|(cp, _)| *cp > document_end)
        || values.windows(2).any(|pair| pair[0].0 > pair[1].0)
    {
        return Err(corrupted(format!(
            "{name} contains out-of-range or non-monotonic CPs"
        )));
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
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u8(data: &[u8], offset: usize, field: &str) -> Result<u8> {
    data.get(offset)
        .copied()
        .ok_or_else(|| corrupted(format!("invalid {field}: truncated data")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIB_POINTERS: usize = 145;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn dpcid(id: u32, kind: u32, flags: u32, fcct: u8) -> [u8; 22] {
        let mut entry = [0u8; 22];
        entry[0..2].copy_from_slice(&DPCID_CHARS.to_le_bytes());
        entry[4..8].copy_from_slice(&flags.to_le_bytes());
        entry[8..12].copy_from_slice(&kind.to_le_bytes());
        entry[16] = fcct;
        entry[17..21].copy_from_slice(&id.to_le_bytes());
        entry
    }

    fn info_table(entries: &[[u8; 22]]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for entry in entries {
            data.extend_from_slice(entry);
        }
        data
    }

    /// Build a FIB plus table stream holding two marks: a plain mark at
    /// CP 1..5 and a column-scoped mark at CP 2..8.
    fn fixture() -> (FileInformationBlock, Vec<u8>, usize) {
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());

        let mut table = Vec::new();
        let infos = info_table(&[
            dpcid(7, 0, FLAG_SQUIGGLE | FLAG_IGNORED, FCCT_CHP),
            dpcid(9, 4, 0, FCCT_CHP | FCCT_SEP),
        ]);
        let infos_offset = table.len() as u32;
        table.extend_from_slice(&infos);
        set_fib_pointer(
            &mut fib_data,
            STTBF_BKMK_FCC,
            infos_offset,
            infos.len() as u32,
        );

        let starts_offset = table.len();
        for cp in [1u32, 2, 11] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&1u16.to_le_bytes()); // ibkl
        table.extend_from_slice(&0x4000u16.to_le_bytes()); // bkc: fNative
        table.extend_from_slice(&0u16.to_le_bytes()); // cDepth
        table.extend_from_slice(&0u16.to_le_bytes()); // ibkl
        table.extend_from_slice(&0x8301u16.to_le_bytes()); // bkc: fCol, columns 1..3
        table.extend_from_slice(&1u16.to_le_bytes()); // cDepth
        set_fib_pointer(&mut fib_data, PLCF_BKF_FCC, starts_offset as u32, 24);

        let ends_offset = table.len() as u32;
        for cp in [5u32, 8, 11] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&1u16.to_le_bytes()); // ibkf
        table.extend_from_slice(&1u16.to_le_bytes()); // cDepth
        table.extend_from_slice(&0u16.to_le_bytes()); // ibkf
        table.extend_from_slice(&0u16.to_le_bytes()); // cDepth
        set_fib_pointer(&mut fib_data, PLCF_BKL_FCC, ends_offset, 20);

        (
            FileInformationBlock::parse(&fib_data).unwrap(),
            table,
            starts_offset,
        )
    }

    #[test]
    fn parses_format_consistency_marks() {
        let (fib, table, _) = fixture();
        let parsed = DocumentFormatConsistencyMarks::parse(&fib, &table)
            .unwrap()
            .expect("marks present");
        assert_eq!(
            parsed.marks(),
            [
                FormatConsistencyMark {
                    start: 1,
                    end: 8,
                    start_depth: 0,
                    end_depth: 0,
                    is_native: true,
                    column_range: None,
                    info: FormatConsistencyInfo {
                        id: 7,
                        kind: FormatConsistencyKind::CharacterFormatting,
                        properties: FormatConsistencyProperties {
                            character: true,
                            table: false,
                            line_separation: false,
                        },
                        squiggle: true,
                        ignored: true,
                        squiggle_changed: false,
                    },
                },
                FormatConsistencyMark {
                    start: 2,
                    end: 5,
                    start_depth: 1,
                    end_depth: 1,
                    is_native: false,
                    column_range: Some((1, 3)),
                    info: FormatConsistencyInfo {
                        id: 9,
                        kind: FormatConsistencyKind::ListFormatting,
                        properties: FormatConsistencyProperties {
                            character: true,
                            table: false,
                            line_separation: true,
                        },
                        squiggle: false,
                        ignored: false,
                        squiggle_changed: false,
                    },
                },
            ]
        );
    }

    #[test]
    fn absent_tables_yield_none() {
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(
            DocumentFormatConsistencyMarks::parse(&fib, &[])
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn rejects_partial_table_sets() {
        let (mut fib, table, _) = fixture();
        let mut fib_data = fib.raw_data().to_vec();
        set_fib_pointer(&mut fib_data, PLCF_BKL_FCC, 0, 0);
        fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentFormatConsistencyMarks::parse(&fib, &table).is_err());
    }

    #[test]
    fn rejects_malformed_tables() {
        let (fib, table, starts_offset) = fixture();

        // Non-monotonic start CPs.
        let mut bad_starts = table.clone();
        bad_starts[starts_offset..starts_offset + 4].copy_from_slice(&11u32.to_le_bytes());
        assert!(DocumentFormatConsistencyMarks::parse(&fib, &bad_starts).is_err());

        // Duplicate ibkl values: both starts reference end index 1.
        let mut duplicate_ibkl = table.clone();
        duplicate_ibkl[starts_offset + 18..starts_offset + 20].copy_from_slice(&1u16.to_le_bytes());
        assert!(DocumentFormatConsistencyMarks::parse(&fib, &duplicate_ibkl).is_err());

        // Broken reciprocal indexes: the second end no longer points back at
        // the second start.
        let mut bad_ibkf = table.clone();
        let second_ibkf = table.len() - 4;
        bad_ibkf[second_ibkf..second_ibkf + 2].copy_from_slice(&1u16.to_le_bytes());
        assert!(DocumentFormatConsistencyMarks::parse(&fib, &bad_ibkf).is_err());

        // fcctPap set.
        assert!(parse_infos(&info_table(&[dpcid(7, 0, 0, FCCT_PAP)])).is_err());

        // Invalid idpci value.
        assert!(parse_infos(&info_table(&[dpcid(7, 12, 0, 0)])).is_err());

        // Duplicate ids.
        assert!(parse_infos(&info_table(&[dpcid(7, 0, 0, 0), dpcid(7, 2, 0, 0)])).is_err());

        // Wrong cchData.
        let mut wrong_chars = info_table(&[dpcid(7, 0, 0, 0)]);
        wrong_chars[6..8].copy_from_slice(&9u16.to_le_bytes());
        assert!(parse_infos(&wrong_chars).is_err());

        // Truncated table.
        assert!(parse_infos(&info_table(&[dpcid(7, 0, 0, 0)])[..10]).is_err());
    }
}
