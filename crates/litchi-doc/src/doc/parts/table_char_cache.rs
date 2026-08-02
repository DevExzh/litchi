//! Deprecated table-character cache (`PlcfTch`, MS-DOC 2.8.29).
//!
//! The `PlcfTch` referenced by `fcPlcfTch` is a producer cache of table
//! characters that Word itself is instructed to ignore (MS-DOC 2.5.7
//! `fcPlcfTch`). Its data elements are `Tch` structures (MS-DOC 2.9.320);
//! the cache is read strictly for metadata but never acted upon.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

/// Table-pointer index of `fcPlcfTch`/`lcbPlcfTch` (MS-DOC 2.5.7 FibRgFcLcb2000).
const TABLE_CHAR_FIB_INDEX: usize = 93;
const MAX_TCH_ENTRIES: usize = 1_000_000;
/// CPs are signed 31-bit positions (MS-DOC 2.2.1).
const MAX_CP: u32 = i32::MAX as u32;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Table-character information for one CP range (`Tch`, MS-DOC 2.9.320; 4 bytes).
///
/// The 31 `unused` bits SHOULD be zero and SHOULD be ignored; they are masked
/// out when reading and written as zero.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCharInfo {
    is_unknown: bool,
}

impl TableCharInfo {
    /// Serialized size of one `Tch` (MS-DOC 2.9.320).
    pub const SIZE: usize = 4;
    const FUNK_MASK: u32 = 0x1;

    /// Create table-character information for a CP range.
    pub const fn new(is_unknown: bool) -> Self {
        Self { is_unknown }
    }

    /// Decode one 4-byte `Tch`. The `unused` bits are ignored.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("Tch must be exactly 4 bytes"));
        }
        let raw = read_u32(data, 0, "Tch")?;
        Ok(Self {
            is_unknown: raw & Self::FUNK_MASK != 0,
        })
    }

    /// Serialize with zeroed `unused` bits.
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        u32::from(self.is_unknown).to_le_bytes()
    }

    /// Whether the table-character cache for the CP range is unknown (`fUnk`).
    pub fn is_unknown(&self) -> bool {
        self.is_unknown
    }
}

/// One cache entry applying to text starting at `start_cp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCharEntry {
    start_cp: u32,
    info: TableCharInfo,
}

impl TableCharEntry {
    pub const fn new(start_cp: u32, info: TableCharInfo) -> Self {
        Self { start_cp, info }
    }

    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    pub fn info(&self) -> TableCharInfo {
        self.info
    }
}

/// A typed `PlcfTch` (MS-DOC 2.8.29).
///
/// CPs are nondecreasing and describe ranges of the main document. The last
/// CP only terminates the final range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableCharacterCache {
    entries: Vec<TableCharEntry>,
    terminal_cp: u32,
}

impl TableCharacterCache {
    /// Serialized stride per entry: one CP plus one `Tch`.
    const STRIDE: usize = 4 + TableCharInfo::SIZE;

    pub fn try_new(entries: Vec<TableCharEntry>, terminal_cp: u32) -> Result<Self> {
        validate_entries(&entries, terminal_cp, None)?;
        Ok(Self {
            entries,
            terminal_cp,
        })
    }

    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_max_cp(data, None)
    }

    fn parse_bytes_with_max_cp(data: &[u8], maximum_cp: Option<u32>) -> Result<Self> {
        if data.len() < 4 || !(data.len() - 4).is_multiple_of(Self::STRIDE) {
            return Err(corrupted("PlcfTch length must have form 8n + 4"));
        }
        let count = (data.len() - 4) / Self::STRIDE;
        if count > MAX_TCH_ENTRIES {
            return Err(corrupted("PlcfTch exceeds one-million-entry cap"));
        }
        let cp_bytes = count
            .checked_add(1)
            .and_then(|value| value.checked_mul(4))
            .ok_or_else(|| corrupted("PlcfTch CP array size overflows"))?;
        let mut positions = Vec::with_capacity(count + 1);
        for index in 0..=count {
            positions.push(read_u32(data, index * 4, "PlcfTch CP")?);
        }
        let terminal_cp = positions[count];
        let mut entries = Vec::with_capacity(count);
        for (index, &start_cp) in positions[..count].iter().enumerate() {
            let element_start = cp_bytes + index * TableCharInfo::SIZE;
            let info = TableCharInfo::from_bytes(
                &data[element_start..element_start + TableCharInfo::SIZE],
            )?;
            entries.push(TableCharEntry::new(start_cp, info));
        }
        validate_entries(&entries, terminal_cp, maximum_cp)?;
        Ok(Self {
            entries,
            terminal_cp,
        })
    }

    pub fn entries(&self) -> &[TableCharEntry] {
        &self.entries
    }
    /// Final PLC CP; it only terminates the last range.
    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }
    pub fn len(&self) -> usize {
        self.entries.len()
    }
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Serialize the complete PLC deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_entries(&self.entries, self.terminal_cp, None)?;
        let size = self
            .entries
            .len()
            .checked_mul(Self::STRIDE)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| corrupted("PlcfTch serialized size overflows"))?;
        let mut data = Vec::with_capacity(size);
        for entry in &self.entries {
            data.extend_from_slice(&entry.start_cp.to_le_bytes());
        }
        data.extend_from_slice(&self.terminal_cp.to_le_bytes());
        for entry in &self.entries {
            data.extend_from_slice(&entry.info.to_bytes());
        }
        Ok(data)
    }

    /// Parse the `PlcfTch` from the Table Stream, when present.
    ///
    /// CPs describe ranges of the main document; they are bounded by
    /// `FibRgLw97.ccpText` plus the two terminating mark positions allowed by
    /// the format's canonical deprecated form (MS-DOC 2.8.29).
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(TABLE_CHAR_FIB_INDEX) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let maximum_cp = fib
            .get_main_doc_range()
            .1
            .checked_add(2)
            .ok_or_else(|| corrupted("main-document PlcfTch CP ceiling overflows"))?;
        let start =
            usize::try_from(offset).map_err(|_| corrupted("PlcfTch offset is too large"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("PlcfTch length is too large"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("PlcfTch range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("PlcfTch extends beyond the table stream"))?;
        Self::parse_bytes_with_max_cp(data, Some(maximum_cp)).map(Some)
    }
}

fn validate_entries(
    entries: &[TableCharEntry],
    terminal_cp: u32,
    maximum_cp: Option<u32>,
) -> Result<()> {
    if entries.len() > MAX_TCH_ENTRIES {
        return Err(corrupted("PlcfTch exceeds one-million-entry cap"));
    }
    let mut previous = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.start_cp > MAX_CP {
            return Err(corrupted(format!(
                "PlcfTch CP {index} exceeds signed CP range"
            )));
        }
        if previous.is_some_and(|value| entry.start_cp < value) {
            return Err(corrupted("PlcfTch CPs are not nondecreasing"));
        }
        previous = Some(entry.start_cp);
    }
    if terminal_cp > MAX_CP {
        return Err(corrupted("PlcfTch terminal CP exceeds signed CP range"));
    }
    if previous.is_some_and(|value| terminal_cp < value) {
        return Err(corrupted("PlcfTch terminal CP precedes the final entry"));
    }
    if maximum_cp.is_some_and(|maximum| terminal_cp > maximum) {
        return Err(corrupted("PlcfTch CP exceeds the main document"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plc_bytes(cps: &[u32], terminal: u32, infos: &[TableCharInfo]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal.to_le_bytes());
        for info in infos {
            data.extend_from_slice(&info.to_bytes());
        }
        data
    }

    #[test]
    fn tch_round_trips_and_ignores_unused_bits() {
        let info = TableCharInfo::new(true);
        let bytes = info.to_bytes();
        assert_eq!(bytes, [1, 0, 0, 0]);
        assert_eq!(TableCharInfo::from_bytes(&bytes).unwrap(), info);
        // The 31 unused bits SHOULD be ignored on read.
        let messy = 0xFFFF_FFFFu32.to_le_bytes();
        assert_eq!(TableCharInfo::from_bytes(&messy).unwrap(), info);
        assert_eq!(
            TableCharInfo::from_bytes(&[0, 0, 0, 0]).unwrap(),
            TableCharInfo::new(false)
        );
        assert!(TableCharInfo::from_bytes(&[0; 3]).is_err());
    }

    #[test]
    fn plcftch_parses_and_round_trips() {
        let infos = [TableCharInfo::new(true), TableCharInfo::new(false)];
        let bytes = plc_bytes(&[0, 100], 102, &infos);
        let cache = TableCharacterCache::parse_bytes(&bytes).unwrap();
        assert_eq!(cache.len(), 2);
        assert_eq!(cache.terminal_cp(), 102);
        assert!(cache.entries()[0].info().is_unknown());
        assert!(!cache.entries()[1].info().is_unknown());
        assert_eq!(cache.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_plc_shapes_and_positions() {
        assert!(TableCharacterCache::parse_bytes(&[]).is_err());
        assert!(TableCharacterCache::parse_bytes(&[0; 13]).is_err());
        // Decreasing CPs.
        let bytes = plc_bytes(&[10, 5], 20, &[TableCharInfo::new(false); 2]);
        assert!(TableCharacterCache::parse_bytes(&bytes).is_err());
        // Terminal CP before the final entry.
        let bytes = plc_bytes(&[10], 5, &[TableCharInfo::new(false)]);
        assert!(TableCharacterCache::parse_bytes(&bytes).is_err());
        // CPs beyond the signed range.
        let bytes = plc_bytes(&[0x8000_0000], 0x8000_0001, &[TableCharInfo::new(false)]);
        assert!(TableCharacterCache::parse_bytes(&bytes).is_err());
    }

    fn fib_with_tch_pointer(ccp_text: u32, offset: u32, length: u32) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + 117 * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        data[152..154].copy_from_slice(&117u16.to_le_bytes());
        // FibRgLw97.ccpText at offset 0x4C bounds the cache CPs.
        data[0x4C..0x50].copy_from_slice(&ccp_text.to_le_bytes());
        let pointer = 154 + TABLE_CHAR_FIB_INDEX * 8;
        data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
        data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        FileInformationBlock::parse(&data).unwrap()
    }

    #[test]
    fn parses_canonical_deprecated_cache_through_fib() {
        // The format's canonical undefined cache: CPs 0, ccpText, ccpText + 2.
        let infos = [TableCharInfo::new(true), TableCharInfo::new(true)];
        let bytes = plc_bytes(&[0, 100], 102, &infos);
        let fib = fib_with_tch_pointer(100, 8, bytes.len() as u32);
        let mut table_stream = vec![0u8; 8];
        table_stream.extend_from_slice(&bytes);
        let cache = TableCharacterCache::parse(&fib, &table_stream)
            .unwrap()
            .unwrap();
        assert_eq!(cache.len(), 2);
        assert_eq!(cache.terminal_cp(), 102);
    }

    #[test]
    fn rejects_cache_cp_beyond_main_document() {
        let bytes = plc_bytes(&[0, 100], 200, &[TableCharInfo::new(false); 2]);
        let fib = fib_with_tch_pointer(100, 0, bytes.len() as u32);
        assert!(TableCharacterCache::parse(&fib, &bytes).is_err());
        // A zero length means the offset is undefined and ignored.
        let fib = fib_with_tch_pointer(100, 0xDEAD, 0);
        assert!(TableCharacterCache::parse(&fib, &[]).unwrap().is_none());
    }
}
