//! Word 97/2000 `SttbSavedBy` save-history parsing.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

const SAVED_BY_FIB_INDEX: usize = 71;
const MAX_SAVED_BY_STRINGS: usize = 20;
/// Exact maximum for the header plus 20 maximum-length UTF-16 STTB strings.
const MAX_SAVED_BY_BYTES: usize = 6 + MAX_SAVED_BY_STRINGS * (2 + u16::MAX as usize * 2);

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// One author/location pair in a document's save history.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SavedByEntry {
    author: String,
    location: String,
}

impl SavedByEntry {
    pub fn new(author: impl Into<String>, location: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            location: location.into(),
        }
    }

    pub fn author(&self) -> &str {
        &self.author
    }

    /// Saved path as inert metadata; this library never resolves or opens it.
    pub fn location(&self) -> &str {
        &self.location
    }
}

/// Ordered `SttbSavedBy` entries, earliest save first.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SavedByTable {
    entries: Vec<SavedByEntry>,
}

impl SavedByTable {
    /// Construct a save-history table after applying the on-disk limits.
    pub fn try_new(entries: Vec<SavedByEntry>) -> Result<Self> {
        validate_entries(&entries)?;
        Ok(Self { entries })
    }

    /// Parse the optional table selected by FIB `fcSttbSavedBy/lcbSttbSavedBy`.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let Some((offset, length)) = fib
            .get_table_pointer(SAVED_BY_FIB_INDEX)
            .filter(|(_, length)| *length != 0)
        else {
            return Ok(Self::default());
        };
        let start =
            usize::try_from(offset).map_err(|_| corrupted("SttbSavedBy offset is too large"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("SttbSavedBy length is too large"))?;
        if length > MAX_SAVED_BY_BYTES {
            return Err(corrupted(
                "SttbSavedBy exceeds its specification-derived size cap",
            ));
        }
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("SttbSavedBy range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("SttbSavedBy extends beyond the table stream"))?;
        Self::parse_bytes(data)
    }

    /// Parse one complete `SttbSavedBy` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_SAVED_BY_BYTES {
            return Err(corrupted(
                "SttbSavedBy exceeds its specification-derived size cap",
            ));
        }
        if data.len() < 6
            || read_u16(data, 0, "SttbSavedBy fExtend")? != 0xFFFF
            || read_u16(data, 4, "SttbSavedBy cbExtra")? != 0
        {
            return Err(corrupted("SttbSavedBy has an invalid header"));
        }
        let string_count = usize::from(read_u16(data, 2, "SttbSavedBy cData")?);
        if string_count > MAX_SAVED_BY_STRINGS || string_count % 2 != 0 {
            return Err(corrupted(
                "SttbSavedBy cData must be even and no greater than 20",
            ));
        }

        let mut strings = Vec::with_capacity(string_count);
        let mut offset = 6usize;
        for index in 0..string_count {
            let unit_count = usize::from(read_u16(
                data,
                offset,
                &format!("SttbSavedBy string {index} length"),
            )?);
            offset = offset
                .checked_add(2)
                .ok_or_else(|| corrupted("SttbSavedBy string offset overflows"))?;
            let byte_count = unit_count
                .checked_mul(2)
                .ok_or_else(|| corrupted("SttbSavedBy string length overflows"))?;
            let end = offset
                .checked_add(byte_count)
                .ok_or_else(|| corrupted("SttbSavedBy string range overflows"))?;
            let bytes = data
                .get(offset..end)
                .ok_or_else(|| corrupted(format!("SttbSavedBy string {index} is truncated")))?;
            let units = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            strings.push(String::from_utf16(&units).map_err(|_| {
                corrupted(format!(
                    "SttbSavedBy string {index} contains invalid UTF-16"
                ))
            })?);
            offset = end;
        }
        if offset != data.len() {
            return Err(corrupted("SttbSavedBy has trailing bytes"));
        }

        let entries = strings
            .chunks_exact(2)
            .map(|pair| SavedByEntry::new(pair[0].clone(), pair[1].clone()))
            .collect();
        Ok(Self { entries })
    }

    pub fn entries(&self) -> &[SavedByEntry] {
        &self.entries
    }

    pub fn latest(&self) -> Option<&SavedByEntry> {
        self.entries.last()
    }

    /// Serialize deterministically as a complete extended-character STTB.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let size = validate_entries(&self.entries)?;
        let mut data = Vec::with_capacity(size);
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        let string_count = u16::try_from(self.entries.len() * 2)
            .map_err(|_| corrupted("SttbSavedBy string count overflows"))?;
        data.extend_from_slice(&string_count.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for entry in &self.entries {
            write_string(&mut data, &entry.author)?;
            write_string(&mut data, &entry.location)?;
        }
        Ok(data)
    }
}

fn validate_entries(entries: &[SavedByEntry]) -> Result<usize> {
    if entries.len() > MAX_SAVED_BY_STRINGS / 2 {
        return Err(corrupted(
            "SttbSavedBy contains more than 10 save-history entries",
        ));
    }
    let mut size = 6usize;
    for entry in entries {
        for value in [&entry.author, &entry.location] {
            let units = value.encode_utf16().count();
            if units > usize::from(u16::MAX) {
                return Err(corrupted(
                    "SttbSavedBy string exceeds 65535 UTF-16 code units",
                ));
            }
            size = size
                .checked_add(2)
                .and_then(|size| size.checked_add(units * 2))
                .ok_or_else(|| corrupted("SttbSavedBy serialized size overflows"))?;
        }
    }
    if size > MAX_SAVED_BY_BYTES {
        return Err(corrupted(
            "SttbSavedBy exceeds its specification-derived size cap",
        ));
    }
    Ok(size)
}

fn write_string(data: &mut Vec<u8>, value: &str) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let count = u16::try_from(units.len())
        .map_err(|_| corrupted("SttbSavedBy string exceeds 65535 UTF-16 code units"))?;
    data.extend_from_slice(&count.to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(strings: &[&str]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&(strings.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for string in strings {
            let units = string.encode_utf16().collect::<Vec<_>>();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        data
    }

    fn fib_with_pointer(offset: u32, length: u32) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + 93 * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        data[152..154].copy_from_slice(&93u16.to_le_bytes());
        let pointer = 154 + SAVED_BY_FIB_INDEX * 8;
        data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
        data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        FileInformationBlock::parse(&data).unwrap()
    }

    #[test]
    fn unicode_save_history_round_trips_exactly() {
        let bytes = table(&["Alice 😀", "C:\\资料\\draft.doc", "Bob", "D:\\final.doc"]);
        let parsed = SavedByTable::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed.entries().len(), 2);
        assert_eq!(parsed.entries()[0].author(), "Alice 😀");
        assert_eq!(parsed.latest().unwrap().location(), "D:\\final.doc");
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_headers_counts_strings_and_ranges() {
        assert!(SavedByTable::parse_bytes(&[]).is_err());

        let mut wrong_extend = table(&[]);
        wrong_extend[0..2].copy_from_slice(&0u16.to_le_bytes());
        assert!(SavedByTable::parse_bytes(&wrong_extend).is_err());

        let mut extra_data = table(&[]);
        extra_data[4..6].copy_from_slice(&1u16.to_le_bytes());
        assert!(SavedByTable::parse_bytes(&extra_data).is_err());
        assert!(SavedByTable::parse_bytes(&table(&["odd"])).is_err());
        assert!(SavedByTable::parse_bytes(&table(&[""; 22])).is_err());

        let mut truncated = table(&["author", "path"]);
        truncated.pop();
        assert!(SavedByTable::parse_bytes(&truncated).is_err());

        let mut invalid_utf16 = table(&["", ""]);
        invalid_utf16[6..8].copy_from_slice(&1u16.to_le_bytes());
        invalid_utf16.insert(8, 0x00);
        invalid_utf16.insert(9, 0xD8);
        assert!(SavedByTable::parse_bytes(&invalid_utf16).is_err());

        let bytes = table(&["author", "path"]);
        let fib = fib_with_pointer(1, bytes.len() as u32);
        assert!(SavedByTable::parse(&fib, &bytes).is_err());
    }
}
