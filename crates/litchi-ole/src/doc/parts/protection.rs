//! Word 2003 range-level protection ("editable ranges") metadata.
//!
//! These are the bookmark-delimited editable ranges of password-protected
//! forms (`SttbfBkmkProt`, `PlcfBkfProt`, `PlcfBklProt`, and `SttbProtUser`,
//! MS-DOC 2.8.1, 2.9.220, 2.9.283, 2.9.293, and 2.9.334).
//!
//! All structures are parsed as inert metadata: usernames are stored, never
//! authenticated, no protection policy is enforced, and documents are never
//! unlocked or modified.

use super::fib::FileInformationBlock;
use crate::doc::package::{DocError, Result};
use std::collections::HashSet;

/// Table-pointer index of `fcSttbfBkmkProt`/`lcbSttbfBkmkProt`.
const STTBF_BKMK_PROT: usize = 141;
/// Table-pointer index of `fcPlcfBkfProt`/`lcbPlcfBkfProt`.
const PLCF_BKF_PROT: usize = 142;
/// Table-pointer index of `fcPlcfBklProt`/`lcbPlcfBklProt`.
const PLCF_BKL_PROT: usize = 143;
/// Table-pointer index of `fcSttbProtUser`/`lcbSttbProtUser`.
const STTB_PROT_USER: usize = 144;

/// Maximum number of range-level protection bookmarks in a document
/// (`SttbfBkmkProt` `cData` limit, MS-DOC 2.9.283).
const MAX_PROTECTED_RANGES: u32 = 0x7FF0;
/// Maximum length of a `SttbProtUser` username in UTF-16 code units
/// (`cchData` limit, MS-DOC 2.9.293).
const MAX_USER_NAME_CHARS: u16 = 0x00FF;

/// `fExtend` marker shared by extended STTB structures.
const STTB_F_EXTEND: u16 = 0xFFFF;
/// `cbExtra` of `SttbfBkmkProt`: one `PRTI` per entry (MS-DOC 2.9.283).
const PRTI_SIZE: usize = 8;
/// `cbExtra` of `SttbProtUser`: one role value per entry (MS-DOC 2.9.293).
const USER_ROLE_SIZE: u16 = 2;

/// `ProtectionType::iProtReadWrite`, the only value a `PRTI` may carry.
const IPROT_READ_WRITE: u16 = 0x0001;

/// `UID::uidEditors` (MS-DOC 2.9.333).
const UID_EDITORS: u16 = 0xFFFB;
/// `UID::uidOwners` (MS-DOC 2.9.333).
const UID_OWNERS: u16 = 0xFFFC;
/// `UID::uidEveryone` (MS-DOC 2.9.333).
const UID_EVERYONE: u16 = 0xFFFF;
/// Largest value a `UidSel` may take while still being a 1-based user index:
/// well-known `UID` values occupy the negative range of the signed integer.
const MAX_USER_INDEX: u16 = 0x7FFF;

/// `BKC` bit layout (MS-DOC 2.9.8).
const BKC_ITC_FIRST_MASK: u16 = 0x007F;
const BKC_F_PUB: u16 = 0x0080;
const BKC_ITC_LIM_SHIFT: u16 = 8;
const BKC_ITC_LIM_MASK: u16 = 0x003F;
const BKC_F_NATIVE: u16 = 0x4000;
const BKC_F_COL: u16 = 0x8000;

/// The permitted editors of one protected text range (`UidSel`, MS-DOC
/// 2.9.334).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UidSel {
    /// `uidEveryone`: all users.
    Everyone,
    /// `uidEditors`: editors of the document.
    Editors,
    /// `uidOwners`: owners of the document.
    Owners,
    /// A 1-based index into the `SttbProtUser` username table.
    User(u16),
}

impl UidSel {
    fn parse(raw: u16) -> Result<Self> {
        match raw {
            UID_EVERYONE => Ok(Self::Everyone),
            UID_EDITORS => Ok(Self::Editors),
            UID_OWNERS => Ok(Self::Owners),
            1..=MAX_USER_INDEX => Ok(Self::User(raw)),
            _ => Err(corrupted(
                "PRTI uidSel is neither a user index nor an allowed UID",
            )),
        }
    }
}

/// The role recorded for one username in `SttbProtUser` (MS-DOC 2.9.293).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionUserRole {
    /// No role is specified for the username.
    Unspecified,
    /// The username specifies an owner.
    Owner,
    /// The username specifies an editor.
    Editor,
}

/// One username from `SttbProtUser`, either a "DOMAIN\NAME" account or an
/// e-mail address. Stored verbatim; never authenticated.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionUser {
    pub name: String,
    pub role: ProtectionUserRole,
}

/// One validated range-level protection bookmark and its editor assignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ProtectedRange {
    /// Character position where the protected range begins.
    pub start: u32,
    /// Character position of the first character beyond the range.
    pub end: u32,
    /// Whether the bookmark is expected to survive saving as RTF/HTML/XML.
    pub is_native: bool,
    /// The single table column the range spans, when the bookmark is a
    /// table-column bookmark (`BKC.fCol`); for range-level protection
    /// bookmarks `itcLim` is always exactly one past `itcFirst`.
    pub column: Option<u8>,
    /// The users permitted to edit this range.
    pub editors: UidSel,
}

/// Complete Word 2003 range-level protection metadata: the usernames from
/// `SttbProtUser` plus the editable ranges from the parallel bookmark tables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentProtectedRanges {
    users: Vec<ProtectionUser>,
    ranges: Vec<ProtectedRange>,
}

impl DocumentProtectedRanges {
    /// Parse the four range-level protection tables addressed by the FIB, or
    /// `None` when the document carries none of them.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentProtectedRanges>> {
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
                parse_editor_assignments(required_slice(
                    fib,
                    table_stream,
                    STTBF_BKMK_PROT,
                    "SttbfBkmkProt",
                )?)?,
                parse_start_plcf(required_slice(fib, table_stream, PLCF_BKF_PROT, "PlcfBkfProt")?)?,
                parse_end_plcf(required_slice(fib, table_stream, PLCF_BKL_PROT, "PlcfBklProt")?)?,
            )
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        };
        if assignments.len() != starts.len() || starts.len() != ends.len() {
            return Err(corrupted(
                "range-level protection info, start, and end table counts do not match",
            ));
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        validate_positions(&starts, document_end, "PlcfBkfProt")?;
        validate_end_positions(&ends, document_end)?;

        let users = user_data
            .map(parse_users)
            .transpose()?
            .unwrap_or_default();

        let mut used_end_indexes = HashSet::with_capacity(starts.len());
        let mut ranges = Vec::with_capacity(starts.len());
        for ((start, start_data), editors) in starts.iter().zip(&assignments) {
            let end_index = usize::try_from(start_data.end_index)
                .map_err(|_| corrupted("range-level protection end index exceeds usize"))?;
            if end_index >= ends.len() || !used_end_indexes.insert(end_index) {
                return Err(corrupted(
                    "range-level protection end indexes must be unique and in range",
                ));
            }
            let end = ends[end_index];
            if start > &end {
                return Err(corrupted(
                    "range-level protection start CP exceeds its end CP",
                ));
            }
            if let UidSel::User(index) = editors {
                if usize::from(*index) > users.len() {
                    return Err(corrupted(
                        "PRTI uidSel user index exceeds the SttbProtUser entry count",
                    ));
                }
            }
            ranges.push(ProtectedRange {
                start: *start,
                end,
                is_native: start_data.is_native,
                column: start_data.column,
                editors: *editors,
            });
        }

        Ok(Some(Self { users, ranges }))
    }

    /// Usernames from `SttbProtUser`, in table order (1-based `UidSel::User`
    /// indexes refer to this order).
    pub fn users(&self) -> &[ProtectionUser] {
        &self.users
    }

    /// The editable ranges in start-CP order.
    pub fn ranges(&self) -> &[ProtectedRange] {
        &self.ranges
    }

    /// Resolve a 1-based `UidSel::User` index to its username entry.
    pub fn user(&self, index: u16) -> Option<&ProtectionUser> {
        usize::from(index).checked_sub(1).and_then(|zero| self.users.get(zero))
    }

    /// The username permitted to edit `range`, when its editor assignment is
    /// an indexed user rather than a well-known group.
    pub fn editors_of(&self, range: &ProtectedRange) -> Option<&ProtectionUser> {
        match range.editors {
            UidSel::User(index) => self.user(index),
            _ => None,
        }
    }
}

#[derive(Debug)]
struct StartData {
    end_index: u32,
    is_native: bool,
    column: Option<u8>,
}

/// Parse `SttbfBkmkProt` (MS-DOC 2.9.283): an extended STTB whose strings are
/// all empty and whose extra data are `PRTI` structures (MS-DOC 2.9.220).
fn parse_editor_assignments(data: &[u8]) -> Result<Vec<UidSel>> {
    if data.len() < 8
        || read_u16(data, 0, "SttbfBkmkProt fExtend")? != STTB_F_EXTEND
        || read_u16(data, 6, "SttbfBkmkProt cbExtra")? != PRTI_SIZE as u16
    {
        return Err(corrupted("SttbfBkmkProt has an invalid header"));
    }
    let count = read_u32(data, 2, "SttbfBkmkProt cData")?;
    if count > MAX_PROTECTED_RANGES {
        return Err(corrupted("SttbfBkmkProt contains too many entries"));
    }
    let count = usize::try_from(count).map_err(|_| corrupted("SttbfBkmkProt count exceeds usize"))?;
    let entry_size = 2usize + PRTI_SIZE;
    let expected = 8usize
        .checked_add(
            count
                .checked_mul(entry_size)
                .ok_or_else(|| corrupted("SttbfBkmkProt size overflows"))?,
        )
        .ok_or_else(|| corrupted("SttbfBkmkProt size overflows"))?;
    if data.len() != expected {
        return Err(corrupted(
            "SttbfBkmkProt byte length does not match its count",
        ));
    }
    let mut assignments = Vec::with_capacity(count);
    let mut offset = 8usize;
    for _ in 0..count {
        if read_u16(data, offset, "SttbfBkmkProt cchData")? != 0 {
            return Err(corrupted("SttbfBkmkProt strings must be empty"));
        }
        let prti = offset + 2;
        let editors = UidSel::parse(read_u16(data, prti, "PRTI uidSel")?)?;
        if read_u16(data, prti + 2, "PRTI iProt")? != IPROT_READ_WRITE {
            return Err(corrupted("PRTI iProt must be iProtReadWrite"));
        }
        // PRTI `i` and `fUseMe` are undefined and MUST be ignored.
        assignments.push(editors);
        offset += entry_size;
    }
    Ok(assignments)
}

/// Parse `PlcfBkfProt`: a `Plcbkf` whose 6-byte data elements are `BKF`
/// structures (`ibkl` plus `bkc`, MS-DOC 2.8.1 and 2.9.9).
fn parse_start_plcf(data: &[u8]) -> Result<Vec<(u32, StartData)>> {
    let count = plcf_count(data, 6, "PlcfBkfProt")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfProt position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 6;
        let bkc = read_u16(data, property + 4, "range-level protection BKC")?;
        if bkc & BKC_F_PUB != 0 {
            return Err(corrupted("range-level protection BKC fPub must be zero"));
        }
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
        values.push((
            read_u32(data, index * 4, "range-level protection start CP")?,
            StartData {
                end_index: read_u32(data, property, "range-level protection ibkl")?,
                is_native: bkc & BKC_F_NATIVE != 0,
                column,
            },
        ));
    }
    Ok(values)
}

/// Parse `PlcfBkfProt`'s companion `Plcbkl`, which contains only CPs and no
/// additional data (MS-DOC 2.8.1); element `ibkl` holds the end CP.
fn parse_end_plcf(data: &[u8]) -> Result<Vec<u32>> {
    if data.len() < 8 || data.len() % 4 != 0 {
        return Err(corrupted("PlcfBklProt has an invalid byte length"));
    }
    let count = data.len() / 4 - 1;
    let mut ends = Vec::with_capacity(count);
    for index in 0..count {
        ends.push(read_u32(data, index * 4, "range-level protection end CP")?);
    }
    Ok(ends)
}

/// Parse `SttbProtUser` (MS-DOC 2.9.293): an extended STTB of unique
/// usernames, each followed by a 2-byte role value.
fn parse_users(data: &[u8]) -> Result<Vec<ProtectionUser>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbProtUser fExtend")? != STTB_F_EXTEND
        || read_u16(data, 4, "SttbProtUser cbExtra")? != USER_ROLE_SIZE
    {
        return Err(corrupted("SttbProtUser has an invalid header"));
    }
    let count = usize::from(read_u16(data, 2, "SttbProtUser cData")?);
    let mut users = Vec::with_capacity(count);
    let mut unique = HashSet::with_capacity(count);
    let mut offset = 6usize;
    for _ in 0..count {
        let length = read_u16(data, offset, "SttbProtUser cchData")?;
        if length > MAX_USER_NAME_CHARS {
            return Err(corrupted("SttbProtUser usernames must not exceed 255 characters"));
        }
        offset += 2;
        let byte_length = usize::from(length) * 2;
        let end = offset
            .checked_add(byte_length)
            .ok_or_else(|| corrupted("SttbProtUser username range overflows"))?;
        let bytes = data
            .get(offset..end)
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
        let role = match read_u16(data, end, "SttbProtUser role")? {
            0x0000 => ProtectionUserRole::Unspecified,
            UID_OWNERS => ProtectionUserRole::Owner,
            UID_EDITORS => ProtectionUserRole::Editor,
            _ => return Err(corrupted("SttbProtUser contains an invalid role")),
        };
        users.push(ProtectionUser { name, role });
        offset = end + 2;
    }
    if offset != data.len() {
        return Err(corrupted("SttbProtUser contains trailing bytes"));
    }
    Ok(users)
}

fn plcf_count(data: &[u8], property_size: usize, name: &str) -> Result<usize> {
    if data.len() < 4 || (data.len() - 4) % (4 + property_size) != 0 {
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

fn validate_end_positions(ends: &[u32], document_end: u32) -> Result<()> {
    if ends.iter().any(|&cp| cp > document_end)
        || ends.windows(2).any(|pair| pair[0] > pair[1])
    {
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
    let bytes = data
        .get(offset..offset.saturating_add(2))
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, name: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset.saturating_add(4))
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal FIB whose table-pointer array covers indexes 0..145,
    /// with a main-document length of `document_end` characters.
    fn fib_bytes(document_end: u32) -> Vec<u8> {
        let pointer_count = 145usize;
        let mut bytes = vec![0u8; 154 + pointer_count * 8];
        bytes[..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        bytes[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        bytes[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
        bytes[76..80].copy_from_slice(&document_end.to_le_bytes());
        bytes[152..154].copy_from_slice(&(pointer_count as u16).to_le_bytes());
        bytes
    }

    fn set_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let base = 154 + index * 8;
        fib[base..base + 4].copy_from_slice(&offset.to_le_bytes());
        fib[base + 4..base + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn utf16(text: &str) -> Vec<u8> {
        text.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn sttb_prot_user(users: &[(&str, u16)]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(users.len() as u16).to_le_bytes());
        data.extend_from_slice(&USER_ROLE_SIZE.to_le_bytes());
        for (name, role) in users {
            let encoded = utf16(name);
            data.extend_from_slice(&((encoded.len() / 2) as u16).to_le_bytes());
            data.extend_from_slice(&encoded);
            data.extend_from_slice(&role.to_le_bytes());
        }
        data
    }

    fn sttbf_bkmk_prot(editors: &[UidSel]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(editors.len() as u32).to_le_bytes());
        data.extend_from_slice(&(PRTI_SIZE as u16).to_le_bytes());
        for editor in editors {
            let raw = match editor {
                UidSel::Everyone => UID_EVERYONE,
                UidSel::Editors => UID_EDITORS,
                UidSel::Owners => UID_OWNERS,
                UidSel::User(index) => *index,
            };
            data.extend_from_slice(&0u16.to_le_bytes()); // cchData
            data.extend_from_slice(&raw.to_le_bytes()); // uidSel
            data.extend_from_slice(&IPROT_READ_WRITE.to_le_bytes()); // iProt
            data.extend_from_slice(&0u16.to_le_bytes()); // i (ignored)
            data.extend_from_slice(&0u16.to_le_bytes()); // fUseMe (ignored)
        }
        data
    }

    /// (start CP, ibkl, bkc) entries.
    fn plcf_bkf_prot(entries: &[(u32, u32, u16)], terminal_cp: u32) -> Vec<u8> {
        let mut data = Vec::new();
        for (cp, _, _) in entries {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal_cp.to_le_bytes());
        for (_, ibkl, bkc) in entries {
            data.extend_from_slice(&ibkl.to_le_bytes());
            data.extend_from_slice(&bkc.to_le_bytes());
        }
        data
    }

    fn plcf_bkl_prot(end_cps: &[u32], terminal_cp: u32) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in end_cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal_cp.to_le_bytes());
        data
    }

    struct Tables {
        users: Vec<u8>,
        infos: Vec<u8>,
        starts: Vec<u8>,
        ends: Vec<u8>,
    }

    impl Tables {
        fn typical() -> Self {
            Self {
                users: sttb_prot_user(&[
                    ("CONTOSO\\alice", UID_EDITORS),
                    ("bob@example.com", UID_OWNERS),
                ]),
                infos: sttbf_bkmk_prot(&[UidSel::User(1), UidSel::Everyone]),
                starts: plcf_bkf_prot(&[(2, 0, BKC_F_NATIVE), (4, 1, 0)], 12),
                ends: plcf_bkl_prot(&[7, 9], 12),
            }
        }

        fn assemble(&self) -> (Vec<u8>, Vec<u8>) {
            let mut fib = fib_bytes(10);
            let mut table = Vec::new();
            for (index, data) in [
                (STTBF_BKMK_PROT, &self.infos),
                (PLCF_BKF_PROT, &self.starts),
                (PLCF_BKL_PROT, &self.ends),
                (STTB_PROT_USER, &self.users),
            ] {
                if !data.is_empty() {
                    set_pointer(&mut fib, index, table.len() as u32, data.len() as u32);
                    table.extend_from_slice(data);
                }
            }
            (fib, table)
        }

        fn parse(&self) -> Result<Option<DocumentProtectedRanges>> {
            let (fib, table) = self.assemble();
            let fib = FileInformationBlock::parse(&fib).unwrap();
            DocumentProtectedRanges::parse(&fib, &table)
        }
    }

    #[test]
    fn parses_range_level_protection_tables() {
        let parsed = Tables::typical().parse().unwrap().unwrap();
        assert_eq!(
            parsed.users(),
            &[
                ProtectionUser {
                    name: "CONTOSO\\alice".to_string(),
                    role: ProtectionUserRole::Editor,
                },
                ProtectionUser {
                    name: "bob@example.com".to_string(),
                    role: ProtectionUserRole::Owner,
                },
            ]
        );
        // Start CP 2 links through ibkl 0 to end CP 7; start CP 4 to end CP 9.
        assert_eq!(
            parsed.ranges(),
            &[
                ProtectedRange {
                    start: 2,
                    end: 7,
                    is_native: true,
                    column: None,
                    editors: UidSel::User(1),
                },
                ProtectedRange {
                    start: 4,
                    end: 9,
                    is_native: false,
                    column: None,
                    editors: UidSel::Everyone,
                },
            ]
        );
        let range = &parsed.ranges()[0];
        assert_eq!(
            parsed.editors_of(range),
            Some(&ProtectionUser {
                name: "CONTOSO\\alice".to_string(),
                role: ProtectionUserRole::Editor,
            })
        );
        assert!(parsed.editors_of(&parsed.ranges()[1]).is_none());
        assert!(parsed.user(3).is_none());
    }

    #[test]
    fn reports_absent_tables_as_none() {
        let fib = fib_bytes(10);
        let fib = FileInformationBlock::parse(&fib).unwrap();
        assert!(DocumentProtectedRanges::parse(&fib, &[]).unwrap().is_none());
    }

    #[test]
    fn parses_username_table_without_bookmark_tables() {
        let tables = Tables {
            infos: Vec::new(),
            starts: Vec::new(),
            ends: Vec::new(),
            ..Tables::typical()
        };
        let parsed = tables.parse().unwrap().unwrap();
        assert_eq!(parsed.users().len(), 2);
        assert!(parsed.ranges().is_empty());
    }

    #[test]
    fn rejects_partially_present_bookmark_tables() {
        let tables = Tables {
            ends: Vec::new(),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn rejects_mismatched_parallel_counts() {
        let tables = Tables {
            infos: sttbf_bkmk_prot(&[UidSel::Everyone]),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn rejects_out_of_bounds_and_reserved_uid_selectors() {
        let tables = Tables {
            infos: sttbf_bkmk_prot(&[UidSel::User(3), UidSel::Everyone]),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
        assert!(UidSel::parse(0).is_err());
        assert!(UidSel::parse(0xFFFA).is_err());
        assert!(UidSel::parse(0xFFFD).is_err());
    }

    #[test]
    fn rejects_duplicate_or_dangling_end_indexes() {
        let tables = Tables {
            starts: plcf_bkf_prot(&[(2, 0, 0), (4, 0, 0)], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
        let tables = Tables {
            starts: plcf_bkf_prot(&[(2, 5, 0), (4, 0, 0)], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn rejects_reversed_and_out_of_range_cps() {
        let tables = Tables {
            starts: plcf_bkf_prot(&[(9, 1, 0), (4, 0, 0)], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
        let tables = Tables {
            ends: plcf_bkl_prot(&[11, 7], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn validates_single_column_constraint() {
        // itcFirst 2 with itcLim 3 is the only allowed non-zero column range.
        let valid = BKC_F_COL | 2 | (3 << BKC_ITC_LIM_SHIFT);
        let tables = Tables {
            starts: plcf_bkf_prot(&[(2, 1, valid), (4, 0, 0)], 12),
            ..Tables::typical()
        };
        let parsed = tables.parse().unwrap().unwrap();
        assert_eq!(parsed.ranges()[0].column, Some(2));
        let invalid = BKC_F_COL | 2 | (4 << BKC_ITC_LIM_SHIFT);
        let tables = Tables {
            starts: plcf_bkf_prot(&[(2, 1, invalid), (4, 0, 0)], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
        // fPub must be zero.
        let tables = Tables {
            starts: plcf_bkf_prot(&[(2, 1, BKC_F_PUB), (4, 0, 0)], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn rejects_invalid_sttb_framing() {
        // Wrong iProt kind.
        let mut infos = sttbf_bkmk_prot(&[UidSel::Everyone]);
        infos[12..14].copy_from_slice(&0x0004u16.to_le_bytes());
        assert!(parse_editor_assignments(&infos).is_err());
        // Nonempty strings are not allowed.
        let mut infos = sttbf_bkmk_prot(&[UidSel::Everyone]);
        infos[8..10].copy_from_slice(&1u16.to_le_bytes());
        assert!(parse_editor_assignments(&infos).is_err());
        // Wrong cbExtra.
        let mut infos = sttbf_bkmk_prot(&[UidSel::Everyone]);
        infos[6..8].copy_from_slice(&4u16.to_le_bytes());
        assert!(parse_editor_assignments(&infos).is_err());
        // Count disagreeing with the byte length.
        let mut infos = sttbf_bkmk_prot(&[UidSel::Everyone]);
        infos[2..6].copy_from_slice(&2u32.to_le_bytes());
        assert!(parse_editor_assignments(&infos).is_err());
        // Trailing bytes in the username table.
        let mut users = sttb_prot_user(&[("a", 0x0000)]);
        users.extend_from_slice(&[0, 0]);
        assert!(parse_users(&users).is_err());
        // Duplicate usernames.
        assert!(parse_users(&sttb_prot_user(&[("a", 0x0000), ("a", 0x0000)])).is_err());
        // Invalid role values.
        assert!(parse_users(&sttb_prot_user(&[("a", 0x0001)])).is_err());
    }

    #[test]
    fn rejects_truncated_tables() {
        assert!(parse_editor_assignments(&sttbf_bkmk_prot(&[UidSel::Everyone])[..12]).is_err());
        assert!(parse_users(&sttb_prot_user(&[("alice", 0x0000)])[..8]).is_err());
        assert!(parse_start_plcf(&[0u8; 9]).is_err());
        assert!(parse_end_plcf(&[0u8; 6]).is_err());
    }
}
