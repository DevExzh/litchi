//! MS-DOC codecs for `SttbfBkmkProt`, `PlcfBkfProt`, `PlcfBklProt`, and
//! `SttbProtUser`.

use super::model::{Mode, Range, Ranges, Reserved, Role, Selector, User};
use crate::package::{DocError, Result};
use crate::parts::fib::FileInformationBlock;
use std::collections::HashSet;

/// Table-pointer index of `fcSttbfBkmkProt`/`lcbSttbfBkmkProt`.
pub(super) const STTBF_BKMK_PROT: usize = 141;
/// Table-pointer index of `fcPlcfBkfProt`/`lcbPlcfBkfProt`.
pub(super) const PLCF_BKF_PROT: usize = 142;
/// Table-pointer index of `fcPlcfBklProt`/`lcbPlcfBklProt`.
pub(super) const PLCF_BKL_PROT: usize = 143;
/// Table-pointer index of `fcSttbProtUser`/`lcbSttbProtUser`.
pub(super) const STTB_PROT_USER: usize = 144;

/// Maximum number of range-level protection bookmarks (`SttbfBkmkProt.cData`,
/// MS-DOC 2.9.283).
const MAX_RANGES: u32 = 0x7FF0;
/// Maximum username length in UTF-16 code units (`SttbProtUser.cchData`,
/// MS-DOC 2.9.293).
const MAX_USER_NAME_CHARS: u16 = 0x00FF;

pub(super) const STTB_F_EXTEND: u16 = 0xFFFF;
pub(super) const PRTI_SIZE: usize = 8;
pub(super) const USER_ROLE_SIZE: u16 = 2;

pub(super) const BKC_ITC_FIRST_MASK: u16 = 0x007F;
pub(super) const BKC_ITC_LIM_SHIFT: u16 = 8;
pub(super) const BKC_ITC_LIM_MASK: u16 = 0x003F;
pub(super) const BKC_F_NATIVE: u16 = 0x4000;
pub(super) const BKC_F_COL: u16 = 0x8000;

pub(super) struct Assignment {
    editor: Selector,
    mode: Mode,
    prti_i: u16,
    prti_use_me: u16,
    bookmark_data: Box<[u8]>,
}

pub(super) struct Start {
    cp: u32,
    end_index: u32,
    bkc: u16,
    is_native: bool,
    column: Option<u8>,
}

impl Ranges {
    /// Parse the four range-level protection tables addressed by the FIB, or
    /// return `None` when the document carries none of them.
    ///
    /// This is the package integration boundary. The semantic model above
    /// does not know about streams or FIB pointer indexes; this codec owns
    /// those details and returns one validated immutable snapshot.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let bookmark_lengths = [STTBF_BKMK_PROT, PLCF_BKF_PROT, PLCF_BKL_PROT]
            .map(|index| fib.get_table_pointer(index).map_or(0, |(_, length)| length));
        let user_data = optional_slice(fib, table_stream, STTB_PROT_USER, "SttbProtUser")?;
        if bookmark_lengths.iter().all(|&length| length == 0) && user_data.is_none() {
            return Ok(None);
        }
        if !bookmark_lengths.iter().all(|&length| length != 0)
            && !bookmark_lengths.iter().all(|&length| length == 0)
        {
            return Err(corrupted(
                "the three parallel range-level protection bookmark tables must be present together",
            ));
        }

        let (assignments, starts, ends) = if bookmark_lengths.iter().all(|&length| length != 0) {
            (
                parse_assignments(required_slice(
                    fib,
                    table_stream,
                    STTBF_BKMK_PROT,
                    "SttbfBkmkProt",
                )?)?,
                parse_starts(required_slice(
                    fib,
                    table_stream,
                    PLCF_BKF_PROT,
                    "PlcfBkfProt",
                )?)?,
                parse_ends(required_slice(
                    fib,
                    table_stream,
                    PLCF_BKL_PROT,
                    "PlcfBklProt",
                )?)?,
            )
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        };

        if assignments.len() != starts.len() || starts.len() != ends.len() {
            return Err(corrupted(
                "range-level protection info, start, and end table counts do not match",
            ));
        }

        let users = user_data.map(parse_users).transpose()?.unwrap_or_default();
        if starts.is_empty() {
            return Ok(Some(Self::from_parts(users, Vec::new())));
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        validate_starts(&starts, document_end)?;
        validate_ends(&ends, document_end)?;

        let mut used_end_indexes = HashSet::with_capacity(starts.len());
        let mut ranges = Vec::with_capacity(starts.len());
        for (start, assignment) in starts.iter().zip(assignments) {
            let end_index = usize::try_from(start.end_index)
                .map_err(|_| corrupted("range-level protection end index exceeds usize"))?;
            if end_index >= ends.len() || !used_end_indexes.insert(end_index) {
                return Err(corrupted(
                    "range-level protection end indexes must be unique and in range",
                ));
            }
            let end = ends[end_index];
            if start.cp > end {
                return Err(corrupted(
                    "range-level protection start CP exceeds its end CP",
                ));
            }
            if let Selector::User(index) = assignment.editor
                && usize::from(index) > users.len()
            {
                return Err(corrupted(
                    "PRTI uidSel user index exceeds the SttbProtUser entry count",
                ));
            }
            ranges.push(Range::from_parts(
                start.cp,
                end,
                start.is_native,
                start.column,
                assignment.editor,
                assignment.mode,
                Reserved::new(
                    start.bkc,
                    assignment.prti_i,
                    assignment.prti_use_me,
                    assignment.bookmark_data,
                ),
            ));
        }

        Ok(Some(Self::from_parts(users, ranges)))
    }
}

/// Parse `SttbfBkmkProt` (MS-DOC 2.9.283). The normally empty STTB string
/// slot is retained as bounded opaque data when a producer populates it.
pub(super) fn parse_assignments(data: &[u8]) -> Result<Vec<Assignment>> {
    if data.len() < 8
        || read_u16(data, 0, "SttbfBkmkProt fExtend")? != STTB_F_EXTEND
        || read_u16(data, 6, "SttbfBkmkProt cbExtra")? != PRTI_SIZE as u16
    {
        return Err(corrupted("SttbfBkmkProt has an invalid header"));
    }
    let count = read_u32(data, 2, "SttbfBkmkProt cData")?;
    if count > MAX_RANGES {
        return Err(corrupted("SttbfBkmkProt contains too many entries"));
    }
    let count =
        usize::try_from(count).map_err(|_| corrupted("SttbfBkmkProt count exceeds usize"))?;
    let payload = data.len() - 8;
    if count > payload / (2 + PRTI_SIZE) {
        return Err(corrupted("SttbfBkmkProt count exceeds the table length"));
    }

    let mut assignments = Vec::with_capacity(count);
    let mut offset = 8usize;
    for _ in 0..count {
        let chars = usize::from(read_u16(data, offset, "SttbfBkmkProt cchData")?);
        let data_start = offset
            .checked_add(2)
            .ok_or_else(|| corrupted("SttbfBkmkProt string offset overflows"))?;
        let data_len = chars
            .checked_mul(2)
            .ok_or_else(|| corrupted("SttbfBkmkProt string length overflows"))?;
        let data_end = data_start
            .checked_add(data_len)
            .ok_or_else(|| corrupted("SttbfBkmkProt string range overflows"))?;
        let bookmark_data = data
            .get(data_start..data_end)
            .ok_or_else(|| corrupted("SttbfBkmkProt string is truncated"))?
            .to_vec()
            .into_boxed_slice();
        let prti = data_end;
        let prti_end = prti
            .checked_add(PRTI_SIZE)
            .ok_or_else(|| corrupted("SttbfBkmkProt PRTI range overflows"))?;
        if data.get(prti..prti_end).is_none() {
            return Err(corrupted("SttbfBkmkProt PRTI is truncated"));
        }
        assignments.push(Assignment {
            editor: Selector::from_raw(read_u16(data, prti, "PRTI uidSel")?),
            mode: Mode::from_raw(read_u16(data, prti + 2, "PRTI iProt")?),
            prti_i: read_u16(data, prti + 4, "PRTI i")?,
            prti_use_me: read_u16(data, prti + 6, "PRTI fUseMe")?,
            bookmark_data,
        });
        offset = prti_end;
    }
    if offset != data.len() {
        return Err(corrupted("SttbfBkmkProt contains trailing bytes"));
    }
    Ok(assignments)
}

/// Parse `PlcfBkfProt`, whose 6-byte elements are `BKF` structures.
pub(super) fn parse_starts(data: &[u8]) -> Result<Vec<Start>> {
    let count = plcf_count(data, 6, "PlcfBkfProt")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfProt position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties
            .checked_add(
                index
                    .checked_mul(6)
                    .ok_or_else(|| corrupted("PlcfBkfProt property offset overflows"))?,
            )
            .ok_or_else(|| corrupted("PlcfBkfProt property offset overflows"))?;
        let bkc = read_u16(data, property + 4, "range-level protection BKC")?;
        let column = if bkc & BKC_F_COL != 0 {
            let first = bkc & BKC_ITC_FIRST_MASK;
            let limit = (bkc >> BKC_ITC_LIM_SHIFT) & BKC_ITC_LIM_MASK;
            if limit != first + 1 {
                return Err(corrupted(
                    "range-level protection BKC column range must span exactly one column",
                ));
            }
            Some(first as u8)
        } else {
            None
        };
        values.push(Start {
            cp: read_u32(data, index * 4, "range-level protection start CP")?,
            end_index: read_u32(data, property, "range-level protection ibkl")?,
            bkc,
            is_native: bkc & BKC_F_NATIVE != 0,
            column,
        });
    }
    Ok(values)
}

/// Parse `PlcfBklProt`, which contains only CPs and a terminal CP.
pub(super) fn parse_ends(data: &[u8]) -> Result<Vec<u32>> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4) {
        return Err(corrupted("PlcfBklProt has an invalid byte length"));
    }
    let count = data.len() / 4 - 1;
    let mut ends = Vec::with_capacity(count);
    for index in 0..count {
        ends.push(read_u32(data, index * 4, "range-level protection end CP")?);
    }
    Ok(ends)
}

/// Parse `SttbProtUser` (MS-DOC 2.9.293).
pub(super) fn parse_users(data: &[u8]) -> Result<Vec<User>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbProtUser fExtend")? != STTB_F_EXTEND
        || read_u16(data, 4, "SttbProtUser cbExtra")? != USER_ROLE_SIZE
    {
        return Err(corrupted("SttbProtUser has an invalid header"));
    }
    let count = usize::from(read_u16(data, 2, "SttbProtUser cData")?);
    let payload = data.len() - 6;
    if count > payload / 4 {
        return Err(corrupted("SttbProtUser count exceeds the table length"));
    }

    let mut users = Vec::with_capacity(count);
    let mut unique = HashSet::with_capacity(count);
    let mut offset = 6usize;
    for _ in 0..count {
        let length = read_u16(data, offset, "SttbProtUser cchData")?;
        if length > MAX_USER_NAME_CHARS {
            return Err(corrupted(
                "SttbProtUser usernames must not exceed 255 characters",
            ));
        }
        let name_start = offset
            .checked_add(2)
            .ok_or_else(|| corrupted("SttbProtUser username offset overflows"))?;
        let byte_length = usize::from(length)
            .checked_mul(2)
            .ok_or_else(|| corrupted("SttbProtUser username range overflows"))?;
        let end = name_start
            .checked_add(byte_length)
            .ok_or_else(|| corrupted("SttbProtUser username range overflows"))?;
        let bytes = data
            .get(name_start..end)
            .ok_or_else(|| corrupted("SttbProtUser username is truncated"))?;
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let name = String::from_utf16(&units)
            .map_err(|_| corrupted("SttbProtUser username is invalid UTF-16"))?;
        if !unique.insert(name.clone()) {
            return Err(corrupted("SttbProtUser usernames must be unique"));
        }
        let role = Role::from_raw(read_u16(data, end, "SttbProtUser role")?);
        users.push(User { name, role });
        offset = end
            .checked_add(2)
            .ok_or_else(|| corrupted("SttbProtUser role offset overflows"))?;
    }
    if offset != data.len() {
        return Err(corrupted("SttbProtUser contains trailing bytes"));
    }
    Ok(users)
}

fn plcf_count(data: &[u8], property_size: usize, name: &str) -> Result<usize> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4 + property_size) {
        return Err(corrupted(format!("{name} has an invalid byte length")));
    }
    Ok((data.len() - 4) / (4 + property_size))
}

fn validate_starts(values: &[Start], document_end: u32) -> Result<()> {
    if values.iter().any(|start| start.cp > document_end)
        || values.windows(2).any(|pair| pair[0].cp > pair[1].cp)
    {
        return Err(corrupted(
            "PlcfBkfProt contains out-of-range or non-monotonic CPs",
        ));
    }
    Ok(())
}

fn validate_ends(ends: &[u32], document_end: u32) -> Result<()> {
    if ends.iter().any(|&cp| cp > document_end) || ends.windows(2).any(|pair| pair[0] > pair[1]) {
        return Err(corrupted(
            "PlcfBklProt contains out-of-range or non-monotonic CPs",
        ));
    }
    Ok(())
}

fn required_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<&'a [u8]> {
    optional_slice(fib, table_stream, index, name)?
        .ok_or_else(|| corrupted(format!("{name} is missing")))
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, name: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{name} offset overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, name: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| corrupted(format!("{name} offset overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}
