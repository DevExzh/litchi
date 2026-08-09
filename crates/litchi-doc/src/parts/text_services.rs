//! Windows Text Services Framework metadata (`Plcfuim` and `PlfguidUim`).
//!
//! The `Plcfuim` (MS-DOC 2.8.33) records data provided by text input services
//! such as handwriting recognizers; its data elements are `UIM` structures
//! (MS-DOC 2.9.335) whose service GUIDs are listed in the `PlfguidUim`
//! (MS-DOC 2.9.198). The service-provided payloads remain opaque blobs in the
//! Table Stream and are never interpreted.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// Table-pointer index of `fcPlcfuim`/`lcbPlcfuim` (MS-DOC 2.5.8 `FibRgFcLcb2002`).
const UIM_FIB_INDEX: usize = 110;
/// Table-pointer index of `fcPlfguidUim`/`lcbPlfguidUim` (MS-DOC 2.5.8 `FibRgFcLcb2002`).
const UIM_GUID_FIB_INDEX: usize = 111;
const MAX_UIM_ENTRIES: usize = 1_000_000;
/// CPs are signed 31-bit positions in the set of all document parts (MS-DOC 2.2.1).
const MAX_CP: u32 = i32::MAX as u32;
/// Serialized size of one GUID in `PlfguidUim.rgguidUim`.
const GUID_SIZE: usize = 16;
/// Serialized stride per PLC entry: one CP plus one `UIM`.
const ENTRY_STRIDE: usize = 4 + Uim::SIZE;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> Result<i16> {
    litchi_core::binary::read_i16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    litchi_core::binary::read_i32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// A typed `PlfguidUim` (MS-DOC 2.9.198): the GUID table referenced by `UIM`s.
///
/// GUIDs are stored as opaque 16-byte values; they identify the text service
/// category or CLSID of the service that provided a `UIM`'s data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UimGuidTable {
    guids: Vec<[u8; GUID_SIZE]>,
}

impl UimGuidTable {
    /// Create a GUID table, enforcing the entry cap.
    pub fn try_new(guids: Vec<[u8; GUID_SIZE]>) -> Result<Self> {
        if guids.len() > MAX_UIM_ENTRIES {
            return Err(corrupted("PlfguidUim exceeds one-million-entry cap"));
        }
        Ok(Self { guids })
    }

    /// Parse a `PlfguidUim`: a 4-byte `iMac` followed by `iMac` GUIDs.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() < 4 {
            return Err(corrupted("PlfguidUim is missing its iMac count"));
        }
        let count = read_u32(data, 0, "PlfguidUim iMac")?;
        let count =
            usize::try_from(count).map_err(|_| corrupted("PlfguidUim iMac is too large"))?;
        if count > MAX_UIM_ENTRIES {
            return Err(corrupted("PlfguidUim exceeds one-million-entry cap"));
        }
        let expected = count
            .checked_mul(GUID_SIZE)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| corrupted("PlfguidUim size overflows"))?;
        if data.len() != expected {
            return Err(corrupted("PlfguidUim length does not match its iMac count"));
        }
        let mut guids = Vec::with_capacity(count);
        for index in 0..count {
            let start = 4 + index * GUID_SIZE;
            let mut guid = [0u8; GUID_SIZE];
            guid.copy_from_slice(&data[start..start + GUID_SIZE]);
            guids.push(guid);
        }
        Ok(Self { guids })
    }

    /// Serialize the complete table deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.guids.len() > MAX_UIM_ENTRIES {
            return Err(corrupted("PlfguidUim exceeds one-million-entry cap"));
        }
        let count = u32::try_from(self.guids.len())
            .map_err(|_| corrupted("PlfguidUim iMac is too large"))?;
        let mut data = Vec::with_capacity(4 + self.guids.len() * GUID_SIZE);
        data.extend_from_slice(&count.to_le_bytes());
        for guid in &self.guids {
            data.extend_from_slice(guid);
        }
        Ok(data)
    }

    /// The GUIDs in table order (`rgguidUim`).
    #[must_use]
    pub fn guids(&self) -> &[[u8; GUID_SIZE]] {
        &self.guids
    }
    #[must_use]
    pub fn len(&self) -> usize {
        self.guids.len()
    }
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.guids.is_empty()
    }
}

/// One text-services data record (`UIM`, MS-DOC 2.9.335; 20 bytes).
///
/// `guid_type_index` and `clsid_tip_index` index the `PlfguidUim` GUID table.
/// The `data_len` bytes at `data_offset` in the Table Stream are the opaque
/// service-provided payload; they are never interpreted.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Uim {
    guid_type_index: u16,
    clsid_tip_index: u16,
    data_offset: u32,
    text_len: u32,
    data_len: u32,
    private_data: u32,
}

impl Uim {
    /// Serialized size of one `UIM` (MS-DOC 2.9.335).
    pub const SIZE: usize = 20;

    /// Create a record, validating the signed field ranges.
    pub fn try_new(
        guid_type_index: u16,
        clsid_tip_index: u16,
        data_offset: u32,
        text_len: u32,
        data_len: u32,
        private_data: u32,
    ) -> Result<Self> {
        if data_offset > MAX_CP {
            return Err(corrupted("UIM fc exceeds the signed offset range"));
        }
        if text_len > MAX_CP {
            return Err(corrupted("UIM cch exceeds the signed character range"));
        }
        Ok(Self {
            guid_type_index,
            clsid_tip_index,
            data_offset,
            text_len,
            data_len,
            private_data,
        })
    }

    /// Decode one 20-byte `UIM`. GUID-table bounds are checked by the PLC.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("UIM must be exactly 20 bytes"));
        }
        let guid_type_index = read_i16(data, 0, "UIM iguidType")?;
        let clsid_tip_index = read_i16(data, 2, "UIM iclsidTip")?;
        let data_offset = read_i32(data, 4, "UIM fc")?;
        let text_len = read_i32(data, 8, "UIM cch")?;
        if guid_type_index < 0 {
            return Err(corrupted("UIM iguidType must be nonnegative"));
        }
        if clsid_tip_index < 0 {
            return Err(corrupted("UIM iclsidTip must be nonnegative"));
        }
        if data_offset < 0 {
            return Err(corrupted("UIM fc must be nonnegative"));
        }
        if text_len < 0 {
            return Err(corrupted("UIM cch must be nonnegative"));
        }
        Ok(Self {
            guid_type_index: guid_type_index as u16,
            clsid_tip_index: clsid_tip_index as u16,
            data_offset: data_offset as u32,
            text_len: text_len as u32,
            data_len: read_u32(data, 12, "UIM cb")?,
            private_data: read_u32(data, 16, "UIM dwPrivate")?,
        })
    }

    /// Serialize exactly as decoded.
    #[must_use]
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        let mut data = [0u8; Self::SIZE];
        data[0..2].copy_from_slice(&self.guid_type_index.to_le_bytes());
        data[2..4].copy_from_slice(&self.clsid_tip_index.to_le_bytes());
        data[4..8].copy_from_slice(&self.data_offset.to_le_bytes());
        data[8..12].copy_from_slice(&self.text_len.to_le_bytes());
        data[12..16].copy_from_slice(&self.data_len.to_le_bytes());
        data[16..20].copy_from_slice(&self.private_data.to_le_bytes());
        data
    }

    /// Index of the service-category GUID within `PlfguidUim.rgguidUim`.
    #[must_use]
    pub fn guid_type_index(&self) -> u16 {
        self.guid_type_index
    }
    /// Index of the service CLSID GUID within `PlfguidUim.rgguidUim`.
    #[must_use]
    pub fn clsid_tip_index(&self) -> u16 {
        self.clsid_tip_index
    }
    /// Offset of the opaque service payload within the Table Stream (`fc`).
    #[must_use]
    pub fn data_offset(&self) -> u32 {
        self.data_offset
    }
    /// Characters of main-document text the record describes (`cch`).
    #[must_use]
    pub fn text_len(&self) -> u32 {
        self.text_len
    }
    /// Size in bytes of the payload at `data_offset` (`cb`).
    #[must_use]
    pub fn data_len(&self) -> u32 {
        self.data_len
    }
    /// Opaque service-generated private data (`dwPrivate`).
    #[must_use]
    pub fn private_data(&self) -> u32 {
        self.private_data
    }
}

/// One text-services record applying to text starting at `start_cp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UimEntry {
    start_cp: u32,
    uim: Uim,
}

impl UimEntry {
    #[must_use]
    pub const fn new(start_cp: u32, uim: Uim) -> Self {
        Self { start_cp, uim }
    }

    #[must_use]
    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    #[must_use]
    pub fn uim(&self) -> &Uim {
        &self.uim
    }
}

/// A typed `Plcfuim` (MS-DOC 2.8.33).
///
/// Unlike an ordinary PLC the elements are not sorted by CP, and duplicate
/// CPs are valid. The last CP is undefined and MUST be ignored.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UimTable {
    entries: Vec<UimEntry>,
    terminal_cp: u32,
}

impl UimTable {
    pub fn try_new(entries: Vec<UimEntry>, terminal_cp: u32) -> Result<Self> {
        validate_entries(&entries, terminal_cp, None, None, None)?;
        Ok(Self {
            entries,
            terminal_cp,
        })
    }

    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(data, None, None, None)
    }

    fn parse_bytes_with_limits(
        data: &[u8],
        maximum_cp: Option<u32>,
        guid_count: Option<u32>,
        table_stream_len: Option<u32>,
    ) -> Result<Self> {
        if data.len() < 4 || !(data.len() - 4).is_multiple_of(ENTRY_STRIDE) {
            return Err(corrupted(format!(
                "Plcfuim length must have form {ENTRY_STRIDE}n + 4"
            )));
        }
        let count = (data.len() - 4) / ENTRY_STRIDE;
        if count > MAX_UIM_ENTRIES {
            return Err(corrupted("Plcfuim exceeds one-million-entry cap"));
        }
        let cp_bytes = count
            .checked_add(1)
            .and_then(|value| value.checked_mul(4))
            .ok_or_else(|| corrupted("Plcfuim CP array size overflows"))?;
        let mut positions = Vec::with_capacity(count + 1);
        for index in 0..=count {
            positions.push(read_u32(data, index * 4, "Plcfuim CP")?);
        }
        let terminal_cp = positions[count];
        let mut entries = Vec::with_capacity(count);
        for (index, &start_cp) in positions[..count].iter().enumerate() {
            let element_start = cp_bytes + index * Uim::SIZE;
            let uim = Uim::from_bytes(&data[element_start..element_start + Uim::SIZE])?;
            entries.push(UimEntry::new(start_cp, uim));
        }
        validate_entries(
            &entries,
            terminal_cp,
            maximum_cp,
            guid_count,
            table_stream_len,
        )?;
        Ok(Self {
            entries,
            terminal_cp,
        })
    }

    #[must_use]
    pub fn entries(&self) -> &[UimEntry] {
        &self.entries
    }
    /// Final PLC CP. The format leaves this undefined; it MUST be ignored.
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

    /// Serialize the complete PLC deterministically, preserving entry order.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_entries(&self.entries, self.terminal_cp, None, None, None)?;
        let size = self
            .entries
            .len()
            .checked_mul(ENTRY_STRIDE)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| corrupted("Plcfuim serialized size overflows"))?;
        let mut data = Vec::with_capacity(size);
        for entry in &self.entries {
            data.extend_from_slice(&entry.start_cp.to_le_bytes());
        }
        data.extend_from_slice(&self.terminal_cp.to_le_bytes());
        for entry in &self.entries {
            data.extend_from_slice(&entry.uim.to_bytes());
        }
        Ok(data)
    }
}

fn validate_entries(
    entries: &[UimEntry],
    _terminal_cp: u32,
    maximum_cp: Option<u32>,
    guid_count: Option<u32>,
    table_stream_len: Option<u32>,
) -> Result<()> {
    if entries.len() > MAX_UIM_ENTRIES {
        return Err(corrupted("Plcfuim exceeds one-million-entry cap"));
    }
    for (index, entry) in entries.iter().enumerate() {
        if entry.start_cp > MAX_CP {
            return Err(corrupted(format!(
                "Plcfuim CP {index} exceeds signed CP range"
            )));
        }
        if maximum_cp.is_some_and(|maximum| entry.start_cp > maximum) {
            return Err(corrupted("Plcfuim CP exceeds the document parts"));
        }
        let uim = &entry.uim;
        if let Some(count) = guid_count {
            if u32::from(uim.guid_type_index) >= count {
                return Err(corrupted("UIM iguidType exceeds the PlfguidUim GUIDs"));
            }
            if u32::from(uim.clsid_tip_index) >= count {
                return Err(corrupted("UIM iclsidTip exceeds the PlfguidUim GUIDs"));
            }
        }
        if let Some(stream_len) = table_stream_len {
            let payload_end = u64::from(uim.data_offset) + u64::from(uim.data_len);
            if payload_end > u64::from(stream_len) {
                return Err(corrupted("UIM payload extends beyond the table stream"));
            }
        }
    }
    // The last CP is undefined and MUST be ignored (MS-DOC 2.8.33), so the
    // terminal CP is stored verbatim without validation.
    Ok(())
}

/// Optional text-services metadata for a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextServicesTables {
    records: Option<UimTable>,
    guids: Option<UimGuidTable>,
}

impl TextServicesTables {
    /// Parse the `Plcfuim` and `PlfguidUim` from the Table Stream.
    ///
    /// When the GUID table is present, every `UIM` index must fall inside it.
    /// Every payload range (`fc`/`cb`) must fit inside the Table Stream.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let guids = parse_guid_table(fib, table_stream)?;
        let records = parse_uim_table(fib, table_stream, guids.as_ref())?;
        Ok(Self { records, guids })
    }

    /// Text-services records (`Plcfuim`, MS-DOC 2.8.33).
    #[must_use]
    pub fn records(&self) -> Option<&UimTable> {
        self.records.as_ref()
    }
    /// Referenced service GUIDs (`PlfguidUim`, MS-DOC 2.9.198).
    #[must_use]
    pub fn guids(&self) -> Option<&UimGuidTable> {
        self.guids.as_ref()
    }
}

fn table_slice<'a>(
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
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))?;
    Ok(Some(data))
}

fn parse_guid_table(
    fib: &FileInformationBlock,
    table_stream: &[u8],
) -> Result<Option<UimGuidTable>> {
    let Some(data) = table_slice(fib, table_stream, UIM_GUID_FIB_INDEX, "PlfguidUim")? else {
        return Ok(None);
    };
    UimGuidTable::parse_bytes(data).map(Some)
}

fn parse_uim_table(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    guids: Option<&UimGuidTable>,
) -> Result<Option<UimTable>> {
    let Some(data) = table_slice(fib, table_stream, UIM_FIB_INDEX, "Plcfuim")? else {
        return Ok(None);
    };
    let maximum_cp = fib
        .get_document_parts_end()
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| corrupted("document-parts Plcfuim CP ceiling overflows"))?;
    let guid_count = guids.map(|table| table.len() as u32);
    let stream_len = u32::try_from(table_stream.len())
        .map_err(|_| corrupted("table stream is too large for UIM bounds"))?;
    UimTable::parse_bytes_with_limits(data, Some(maximum_cp), guid_count, Some(stream_len))
        .map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn uim(guid_type: u16, clsid: u16, offset: u32, len: u32) -> Uim {
        Uim::try_new(guid_type, clsid, offset, 3, len, 0xABCD).unwrap()
    }

    fn plc_bytes(cps: &[u32], terminal: u32, uims: &[Uim]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal.to_le_bytes());
        for uim in uims {
            data.extend_from_slice(&uim.to_bytes());
        }
        data
    }

    #[test]
    fn uim_round_trips_exactly() {
        let record = uim(1, 0, 48, 16);
        let bytes = record.to_bytes();
        assert_eq!(bytes.len(), Uim::SIZE);
        assert_eq!(Uim::from_bytes(&bytes).unwrap(), record);
        assert!(Uim::from_bytes(&bytes[..19]).is_err());
        // Signed fields reject negatives.
        let mut negative = bytes;
        negative[1] = 0x80;
        assert!(Uim::from_bytes(&negative).is_err());
        let mut negative = bytes;
        negative[7] = 0x80;
        assert!(Uim::from_bytes(&negative).is_err());
        assert!(Uim::try_new(0, 0, 0x8000_0000, 0, 0, 0).is_err());
        assert!(Uim::try_new(0, 0, 0, 0x8000_0000, 0, 0).is_err());
    }

    #[test]
    fn guid_table_parses_and_round_trips() {
        let guids = UimGuidTable::try_new(vec![[0x11; 16], [0x22; 16]]).unwrap();
        let bytes = guids.to_bytes().unwrap();
        assert_eq!(bytes.len(), 4 + 2 * 16);
        let parsed = UimGuidTable::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed, guids);
        assert_eq!(parsed.guids()[1], [0x22; 16]);
        // Mismatched iMac/length and truncated headers are rejected.
        assert!(UimGuidTable::parse_bytes(&bytes[..20]).is_err());
        assert!(UimGuidTable::parse_bytes(&[1, 0]).is_err());
        let mut overlong = bytes.clone();
        overlong.push(0);
        assert!(UimGuidTable::parse_bytes(&overlong).is_err());
    }

    #[test]
    fn plcfuim_preserves_unsorted_and_duplicate_cps() {
        let records = [uim(0, 0, 0, 8), uim(1, 1, 8, 8), uim(0, 1, 16, 8)];
        // Elements are not sorted by CP; duplicates are valid.
        let bytes = plc_bytes(&[20, 4, 4], 0xFFFF_FFFE, &records);
        let table = UimTable::parse_bytes(&bytes).unwrap();
        assert_eq!(table.len(), 3);
        assert_eq!(table.terminal_cp(), 0xFFFF_FFFE);
        assert_eq!(table.entries()[1].start_cp(), 4);
        assert_eq!(table.entries()[1].uim().guid_type_index(), 1);
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_plc_shapes() {
        assert!(UimTable::parse_bytes(&[]).is_err());
        assert!(UimTable::parse_bytes(&[0; 9]).is_err());
        let mut bytes = plc_bytes(&[4], 8, &[uim(0, 0, 0, 0)]);
        bytes.pop();
        assert!(UimTable::parse_bytes(&bytes).is_err());
        // CPs beyond the signed range.
        let bytes = plc_bytes(&[0x8000_0000], 0, &[uim(0, 0, 0, 0)]);
        assert!(UimTable::parse_bytes(&bytes).is_err());
    }

    fn fib_with_pointers(ccp_text: u32, pairs: &[(usize, u32, u32)]) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + 117 * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        data[152..154].copy_from_slice(&117u16.to_le_bytes());
        // FibRgLw97.ccpText at offset 0x4C feeds the document-parts ceiling.
        data[0x4C..0x50].copy_from_slice(&ccp_text.to_le_bytes());
        for (index, offset, length) in pairs {
            let pointer = 154 + index * 8;
            data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        }
        FileInformationBlock::parse(&data).unwrap()
    }

    #[test]
    fn parses_both_tables_through_fib_with_guid_and_stream_bounds() {
        let guids = UimGuidTable::try_new(vec![[0x11; 16], [0x22; 16]])
            .unwrap()
            .to_bytes()
            .unwrap();
        let records = plc_bytes(&[10], 20, &[uim(1, 0, 60, 4)]);
        let mut table_stream = vec![0u8; 4];
        table_stream.extend_from_slice(&guids);
        table_stream.extend_from_slice(&records);
        table_stream.resize(128, 0);
        let fib = fib_with_pointers(
            100,
            &[
                (UIM_GUID_FIB_INDEX, 4, guids.len() as u32),
                (
                    UIM_FIB_INDEX,
                    (4 + guids.len()) as u32,
                    records.len() as u32,
                ),
            ],
        );
        let tables = TextServicesTables::parse(&fib, &table_stream).unwrap();
        assert_eq!(tables.guids().unwrap().len(), 2);
        assert_eq!(tables.records().unwrap().len(), 1);
    }

    #[test]
    fn rejects_guid_index_outside_guid_table() {
        let guids = UimGuidTable::try_new(vec![[0x11; 16]])
            .unwrap()
            .to_bytes()
            .unwrap();
        let records = plc_bytes(&[10], 20, &[uim(1, 0, 60, 4)]);
        let mut table_stream = guids.clone();
        table_stream.extend_from_slice(&records);
        table_stream.resize(64, 0);
        let fib = fib_with_pointers(
            100,
            &[
                (UIM_GUID_FIB_INDEX, 0, guids.len() as u32),
                (UIM_FIB_INDEX, guids.len() as u32, records.len() as u32),
            ],
        );
        assert!(TextServicesTables::parse(&fib, &table_stream).is_err());
        // Without a GUID table the index bounds check does not apply.
        let fib = fib_with_pointers(100, &[(UIM_FIB_INDEX, 0, records.len() as u32)]);
        let mut records_only = records.clone();
        records_only.resize(64, 0);
        assert!(TextServicesTables::parse(&fib, &records_only).is_ok());
    }

    #[test]
    fn rejects_payload_outside_table_stream_and_cps_beyond_parts() {
        let records = plc_bytes(&[10], 20, &[uim(0, 0, 60, 16)]);
        let fib = fib_with_pointers(100, &[(UIM_FIB_INDEX, 0, records.len() as u32)]);
        // Payload [60, 76) overruns the 64-byte table stream.
        let mut table_stream = records.clone();
        table_stream.resize(64, 0);
        assert!(TextServicesTables::parse(&fib, &table_stream).is_err());
        // CP beyond the document parts.
        let records = plc_bytes(&[500], 600, &[uim(0, 0, 0, 0)]);
        let fib = fib_with_pointers(100, &[(UIM_FIB_INDEX, 0, records.len() as u32)]);
        assert!(TextServicesTables::parse(&fib, &records).is_err());
    }
}
