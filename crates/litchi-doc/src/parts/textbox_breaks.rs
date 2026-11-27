//! Textbox break tables (`PlcfTxbxBkd` and `PlcfTxbxHdrBkd`).
//!
//! These PLCs (MS-DOC 2.8.30 and 2.8.31) associate ranges of the textboxes
//! and header-textboxes subdocuments with `FTXBXS` objects from the
//! `PlcftxbxTxt`/`PlcfHdrtxbxTxt` tables. Their data elements are `Tbkd`
//! structures (MS-DOC 2.9.312). The version-specific flag bits are producer
//! caches that the format instructs readers to ignore; they are never
//! interpreted here.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use super::textbox::{FIB_INDEX_PLCF_HDR_TXBX_TXT, FIB_INDEX_PLCF_TXBX_TXT, FTXBXS_LEN};

/// Table-pointer index of `fcPlcfTxbxBkd`/`lcbPlcfTxbxBkd` (MS-DOC 2.5.6 `FibRgFcLcb97`).
const TXBX_BKD_FIB_INDEX: usize = 75;
/// Table-pointer index of `fcPlcfTxbxHdrBkd`/`lcbPlcfTxbxHdrBkd` (MS-DOC 2.5.6 `FibRgFcLcb97`).
const TXBX_HDR_BKD_FIB_INDEX: usize = 76;
const MAX_TBKD_ENTRIES: usize = 1_000_000;
/// CPs are signed 31-bit positions (MS-DOC 2.2.1).
const MAX_CP: u32 = i32::MAX as u32;
/// Serialized stride per PLC entry: one CP plus one `Tbkd`.
const ENTRY_STRIDE: usize = 4 + TextBoxBreak::SIZE;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> Result<i16> {
    litchi_core::binary::read_i16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Which textbox story a break table describes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextBoxBreakKind {
    /// `PlcfTxbxBkd` (MS-DOC 2.8.30): ranges of the textboxes subdocument.
    Main,
    /// `PlcfTxbxHdrBkd` (MS-DOC 2.8.31): ranges of the header-textboxes subdocument.
    Header,
}

impl TextBoxBreakKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Main => "PlcfTxbxBkd",
            Self::Header => "PlcfTxbxHdrBkd",
        }
    }

    const fn fib_index(self) -> usize {
        match self {
            Self::Main => TXBX_BKD_FIB_INDEX,
            Self::Header => TXBX_HDR_BKD_FIB_INDEX,
        }
    }

    const fn text_fib_index(self) -> usize {
        match self {
            Self::Main => FIB_INDEX_PLCF_TXBX_TXT,
            Self::Header => FIB_INDEX_PLCF_HDR_TXBX_TXT,
        }
    }

    /// Story-relative end CP bounding the break table's CPs.
    fn story_end(self, fib: &FileInformationBlock) -> u32 {
        let range = match self {
            Self::Main => fib.get_textbox_range(),
            Self::Header => fib.get_header_textbox_range(),
        };
        range.map_or(0, |(start, end)| end.saturating_sub(start))
    }
}

/// One textbox association (`Tbkd`, MS-DOC 2.9.312; 6 bytes).
///
/// The `reserved1`, `fMarkDelete`, and `reserved2` bits MUST be zero and MUST
/// be ignored, and `dcpDepend`, `fUnk`, and `fTextOverflow` are deprecated
/// version-specific caches that SHOULD be ignored. All of them are masked out
/// when reading and written as zero.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextBoxBreak {
    itxbxs: i16,
}

impl TextBoxBreak {
    /// Serialized size of one `Tbkd` (MS-DOC 2.9.312).
    pub const SIZE: usize = 6;

    /// Create an association with the `FTXBXS` object at `itxbxs`.
    #[must_use]
    pub const fn new(itxbxs: i16) -> Self {
        Self { itxbxs }
    }

    /// Decode one 6-byte `Tbkd`, ignoring the deprecated fields and flag bits.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("Tbkd must be exactly 6 bytes"));
        }
        Ok(Self {
            itxbxs: read_i16(data, 0, "Tbkd itxbxs")?,
        })
    }

    /// Serialize with a zeroed `dcpDepend` and zeroed flag bits.
    #[must_use]
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        let mut data = [0u8; Self::SIZE];
        data[0..2].copy_from_slice(&self.itxbxs.to_le_bytes());
        data
    }

    /// Index of the associated `FTXBXS` object within the corresponding
    /// `PlcftxbxTxt`/`PlcfHdrtxbxTxt`. Meaningless on the final entry.
    #[must_use]
    pub fn itxbxs(&self) -> i16 {
        self.itxbxs
    }
}

/// One break-table entry applying to text starting at `start_cp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextBoxBreakEntry {
    start_cp: u32,
    break_info: TextBoxBreak,
}

impl TextBoxBreakEntry {
    #[must_use]
    pub const fn new(start_cp: u32, break_info: TextBoxBreak) -> Self {
        Self {
            start_cp,
            break_info,
        }
    }

    #[must_use]
    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    #[must_use]
    pub fn break_info(&self) -> TextBoxBreak {
        self.break_info
    }
}

/// A typed `PlcfTxbxBkd` or `PlcfTxbxHdrBkd`.
///
/// CPs are strictly increasing; duplicate CPs are forbidden (MS-DOC 2.8.30,
/// 2.8.31). The last CP only terminates the final range, and the final
/// `Tbkd` is not associated with any `FTXBXS` object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextBoxBreakTable {
    kind: TextBoxBreakKind,
    entries: Vec<TextBoxBreakEntry>,
    terminal_cp: u32,
}

impl TextBoxBreakTable {
    pub fn try_new(
        kind: TextBoxBreakKind,
        entries: Vec<TextBoxBreakEntry>,
        terminal_cp: u32,
    ) -> Result<Self> {
        validate_entries(kind, &entries, terminal_cp, None, None)?;
        Ok(Self {
            kind,
            entries,
            terminal_cp,
        })
    }

    pub fn parse_bytes(kind: TextBoxBreakKind, data: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(kind, data, None, None)
    }

    fn parse_bytes_with_limits(
        kind: TextBoxBreakKind,
        data: &[u8],
        maximum_cp: Option<u32>,
        textbox_count: Option<u32>,
    ) -> Result<Self> {
        let name = kind.name();
        if data.len() < 4 || !(data.len() - 4).is_multiple_of(ENTRY_STRIDE) {
            return Err(corrupted(format!(
                "{name} length must have form {ENTRY_STRIDE}n + 4"
            )));
        }
        let count = (data.len() - 4) / ENTRY_STRIDE;
        if count > MAX_TBKD_ENTRIES {
            return Err(corrupted(format!("{name} exceeds one-million-entry cap")));
        }
        let cp_bytes = count
            .checked_add(1)
            .and_then(|value| value.checked_mul(4))
            .ok_or_else(|| corrupted(format!("{name} CP array size overflows")))?;
        let mut positions = Vec::with_capacity(count + 1);
        for index in 0..=count {
            positions.push(read_u32(data, index * 4, "Tbkd PLC CP")?);
        }
        let terminal_cp = positions[count];
        let mut entries = Vec::with_capacity(count);
        for (index, &start_cp) in positions[..count].iter().enumerate() {
            let element_start = cp_bytes + index * TextBoxBreak::SIZE;
            let break_info =
                TextBoxBreak::from_bytes(&data[element_start..element_start + TextBoxBreak::SIZE])?;
            entries.push(TextBoxBreakEntry::new(start_cp, break_info));
        }
        validate_entries(kind, &entries, terminal_cp, maximum_cp, textbox_count)?;
        Ok(Self {
            kind,
            entries,
            terminal_cp,
        })
    }

    #[must_use]
    pub fn kind(&self) -> TextBoxBreakKind {
        self.kind
    }
    #[must_use]
    pub fn entries(&self) -> &[TextBoxBreakEntry] {
        &self.entries
    }
    /// Final PLC CP; it only terminates the last range.
    #[must_use]
    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Serialize the complete PLC deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_entries(self.kind, &self.entries, self.terminal_cp, None, None)?;
        let size = self
            .entries
            .len()
            .checked_mul(ENTRY_STRIDE)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| corrupted("Tbkd PLC serialized size overflows"))?;
        let mut data = Vec::with_capacity(size);
        for entry in &self.entries {
            data.extend_from_slice(&entry.start_cp.to_le_bytes());
        }
        data.extend_from_slice(&self.terminal_cp.to_le_bytes());
        for entry in &self.entries {
            data.extend_from_slice(&entry.break_info.to_bytes());
        }
        Ok(data)
    }
}

fn validate_entries(
    kind: TextBoxBreakKind,
    entries: &[TextBoxBreakEntry],
    terminal_cp: u32,
    maximum_cp: Option<u32>,
    textbox_count: Option<u32>,
) -> Result<()> {
    let name = kind.name();
    if entries.len() > MAX_TBKD_ENTRIES {
        return Err(corrupted(format!("{name} exceeds one-million-entry cap")));
    }
    let mut previous = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.start_cp > MAX_CP {
            return Err(corrupted(format!(
                "{name} CP {index} exceeds signed CP range"
            )));
        }
        if let Some(value) = previous
            && entry.start_cp <= value
        {
            return Err(corrupted(format!("{name} CPs are not strictly increasing")));
        }
        previous = Some(entry.start_cp);
    }
    if terminal_cp > MAX_CP {
        return Err(corrupted(format!(
            "{name} terminal CP exceeds signed CP range"
        )));
    }
    if previous.is_some_and(|value| terminal_cp < value) {
        return Err(corrupted(format!(
            "{name} terminal CP precedes the final entry"
        )));
    }
    if maximum_cp.is_some_and(|maximum| terminal_cp > maximum) {
        return Err(corrupted(format!(
            "{name} CP exceeds the textbox subdocument"
        )));
    }
    // All but the final Tbkd must reference a valid FTXBXS index; the final
    // entry is unassociated and its itxbxs is ignored (MS-DOC 2.9.312).
    let associated = entries.len().saturating_sub(1);
    for entry in &entries[..associated] {
        let itxbxs = entry.break_info.itxbxs();
        if itxbxs < 0 {
            return Err(corrupted(format!("{name} has a negative FTXBXS index")));
        }
        if textbox_count.is_some_and(|count| itxbxs as u32 >= count) {
            return Err(corrupted(format!(
                "{name} FTXBXS index exceeds the textbox text table"
            )));
        }
    }
    Ok(())
}

/// Optional main and header textbox break tables for a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextBoxBreakTables {
    main: Option<TextBoxBreakTable>,
    header: Option<TextBoxBreakTable>,
}

impl TextBoxBreakTables {
    /// Parse both break PLCFs from the Table Stream.
    ///
    /// CPs are bounded by the corresponding subdocument's character count.
    /// When the matching `FTXBXS` text table is well formed, every associated
    /// `itxbxs` must index one of its objects.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        Ok(Self {
            main: parse_fib_table(fib, table_stream, TextBoxBreakKind::Main)?,
            header: parse_fib_table(fib, table_stream, TextBoxBreakKind::Header)?,
        })
    }

    /// Main textbox break table (`PlcfTxbxBkd`, MS-DOC 2.8.30).
    #[must_use]
    pub fn main(&self) -> Option<&TextBoxBreakTable> {
        self.main.as_ref()
    }
    /// Header textbox break table (`PlcfTxbxHdrBkd`, MS-DOC 2.8.31).
    #[must_use]
    pub fn header(&self) -> Option<&TextBoxBreakTable> {
        self.header.as_ref()
    }
}

fn parse_fib_table(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    kind: TextBoxBreakKind,
) -> Result<Option<TextBoxBreakTable>> {
    let Some((offset, length)) = fib.get_table_pointer(kind.fib_index()) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let name = kind.name();
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))?;
    let textbox_count = fib
        .get_table_pointer(kind.text_fib_index())
        .map(|(_, length)| length as usize)
        .filter(|length| *length >= 4 && (*length - 4) % (4 + FTXBXS_LEN) == 0)
        .map(|length| ((length - 4) / (4 + FTXBXS_LEN)) as u32);
    TextBoxBreakTable::parse_bytes_with_limits(kind, data, Some(kind.story_end(fib)), textbox_count)
        .map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plc_bytes(cps: &[u32], terminal: u32, breaks: &[TextBoxBreak]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal.to_le_bytes());
        for break_info in breaks {
            data.extend_from_slice(&break_info.to_bytes());
        }
        data
    }

    #[test]
    fn tbkd_round_trips_and_ignores_deprecated_fields() {
        let break_info = TextBoxBreak::new(3);
        let bytes = break_info.to_bytes();
        assert_eq!(bytes, [3, 0, 0, 0, 0, 0]);
        assert_eq!(TextBoxBreak::from_bytes(&bytes).unwrap(), break_info);
        // dcpDepend and the flag bits SHOULD be ignored on read.
        let messy = [3, 0, 0x7F, 0x80, 0xFF, 0xFF];
        assert_eq!(TextBoxBreak::from_bytes(&messy).unwrap(), break_info);
        assert!(TextBoxBreak::from_bytes(&bytes[..5]).is_err());
    }

    #[test]
    fn break_table_parses_and_round_trips() {
        let breaks = [
            TextBoxBreak::new(0),
            TextBoxBreak::new(1),
            TextBoxBreak::new(-1),
        ];
        let bytes = plc_bytes(&[0, 5, 9], 12, &breaks);
        let table = TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Main, &bytes).unwrap();
        assert_eq!(table.len(), 3);
        assert_eq!(table.terminal_cp(), 12);
        assert_eq!(table.entries()[1].break_info().itxbxs(), 1);
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_duplicate_cps_and_bad_shapes() {
        let breaks = [TextBoxBreak::new(0); 2];
        assert!(TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Main, &[]).is_err());
        assert!(TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Header, &[0; 6]).is_err());
        // Duplicate CPs are forbidden.
        let bytes = plc_bytes(&[5, 5], 9, &breaks);
        assert!(TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Main, &bytes).is_err());
        // Decreasing CPs.
        let bytes = plc_bytes(&[5, 4], 9, &breaks);
        assert!(TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Main, &bytes).is_err());
        // Terminal CP before the final entry.
        let bytes = plc_bytes(&[5], 4, &[TextBoxBreak::new(0)]);
        assert!(TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Main, &bytes).is_err());
        // Negative itxbxs on an associated (non-final) entry.
        let bytes = plc_bytes(&[0, 5], 9, &[TextBoxBreak::new(-1), TextBoxBreak::new(0)]);
        assert!(TextBoxBreakTable::parse_bytes(TextBoxBreakKind::Main, &bytes).is_err());
    }

    fn fib_with_pointers(ccp_txbx: u32, pairs: &[(usize, u32, u32)]) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + 117 * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        data[152..154].copy_from_slice(&117u16.to_le_bytes());
        // FibRgLw97.ccpTxbx at offset 0x64 bounds the main break-table CPs.
        data[0x64..0x68].copy_from_slice(&ccp_txbx.to_le_bytes());
        for (index, offset, length) in pairs {
            let pointer = 154 + index * 8;
            data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        }
        FileInformationBlock::parse(&data).unwrap()
    }

    /// Well-formed PlcftxbxTxt length for `count` FTXBXS objects.
    const fn txt_plc_len(count: u32) -> u32 {
        4 + count * (4 + FTXBXS_LEN as u32)
    }

    #[test]
    fn parses_both_tables_through_fib_with_ftxbxs_bounds() {
        let main = plc_bytes(&[0, 5], 9, &[TextBoxBreak::new(1), TextBoxBreak::new(0)]);
        let mut table_stream = vec![0u8; 4];
        table_stream.extend_from_slice(&main);
        let fib = fib_with_pointers(
            9,
            &[
                (FIB_INDEX_PLCF_TXBX_TXT, 0x800, txt_plc_len(2)),
                (TXBX_BKD_FIB_INDEX, 4, main.len() as u32),
            ],
        );
        let tables = TextBoxBreakTables::parse(&fib, &table_stream).unwrap();
        assert_eq!(tables.main().unwrap().len(), 2);
        assert!(tables.header().is_none());
    }

    #[test]
    fn rejects_itxbxs_beyond_ftxbxs_table() {
        let main = plc_bytes(&[0, 5], 9, &[TextBoxBreak::new(2), TextBoxBreak::new(0)]);
        let fib = fib_with_pointers(
            9,
            &[
                (FIB_INDEX_PLCF_TXBX_TXT, 0x800, txt_plc_len(2)),
                (TXBX_BKD_FIB_INDEX, 0, main.len() as u32),
            ],
        );
        assert!(TextBoxBreakTables::parse(&fib, &main).is_err());
        // A malformed text table disables the bound instead of failing.
        let fib = fib_with_pointers(
            9,
            &[
                (FIB_INDEX_PLCF_TXBX_TXT, 0x800, 7),
                (TXBX_BKD_FIB_INDEX, 0, main.len() as u32),
            ],
        );
        assert!(TextBoxBreakTables::parse(&fib, &main).is_ok());
    }

    #[test]
    fn rejects_break_cp_beyond_textbox_subdocument() {
        let main = plc_bytes(&[0, 5], 20, &[TextBoxBreak::new(0), TextBoxBreak::new(0)]);
        let fib = fib_with_pointers(9, &[(TXBX_BKD_FIB_INDEX, 0, main.len() as u32)]);
        assert!(TextBoxBreakTables::parse(&fib, &main).is_err());
    }
}
