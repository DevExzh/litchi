//! Standard bookmark tables for Word 97+ binary documents.

use super::super::bookmark::Bookmark;
use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use crate::plcf::Plcf;
use std::collections::HashSet;

/// Parsed standard bookmarks in start-CP order.
#[derive(Debug, Clone, Default)]
pub struct BookmarksTable {
    bookmarks: Vec<Bookmark>,
}

impl BookmarksTable {
    /// Parse `SttbfBkmk`, `PlcfBkf`, and `PlcfBkl`.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let lengths = [21usize, 22, 23].map(|index| {
            fib.get_table_pointer(index)
                .map(|(_, length)| length)
                .unwrap_or(0)
        });
        if lengths.iter().all(|&length| length == 0) {
            return Ok(Self::default());
        }
        if lengths.contains(&0) {
            return Err(PackageError::Corrupted(
                "standard bookmark tables must be present together".to_string(),
            ));
        }

        let names = parse_names(required_slice(fib, table_stream, 21, "SttbfBkmk")?)?;
        let starts_data = required_slice(fib, table_stream, 22, "PlcfBkf")?;
        if starts_data.len() < 4 || (starts_data.len() - 4) % 8 != 0 {
            return Err(PackageError::Corrupted(
                "PlcfBkf has an invalid byte length".to_string(),
            ));
        }
        let starts = Plcf::parse(starts_data, 4)
            .ok_or_else(|| PackageError::Corrupted("PlcfBkf is malformed".to_string()))?;
        if starts.count() != names.len() {
            return Err(PackageError::Corrupted(
                "SttbfBkmk and PlcfBkf counts do not match".to_string(),
            ));
        }

        let document_end = fib.get_document_parts_end().ok_or_else(|| {
            PackageError::Corrupted("document-part character counts overflow".to_string())
        })?;
        validate_cps(&starts, document_end, "PlcfBkf")?;

        let ends_data = required_slice(fib, table_stream, 23, "PlcfBkl")?;
        if ends_data.len() != (names.len() + 1) * 4 {
            return Err(PackageError::Corrupted(
                "PlcfBkl count does not match PlcfBkf".to_string(),
            ));
        }
        let mut ends = Vec::with_capacity(names.len() + 1);
        for offset in (0..ends_data.len()).step_by(4) {
            ends.push(
                litchi_core::binary::read_u32_le(ends_data, offset).map_err(|error| {
                    PackageError::Corrupted(format!("invalid PlcfBkl CP: {error}"))
                })?,
            );
        }
        // The final CP of a bookmark PLC is ignored per [MS-DOC] 2.8.10;
        // writers disagree on whether it counts the paragraph mark that
        // separates the document parts, so no constraint is placed on it.
        if ends[..ends.len() - 1].iter().any(|&cp| cp > document_end)
            || ends[..ends.len() - 1]
                .windows(2)
                .any(|pair| pair[0] > pair[1])
        {
            return Err(PackageError::Corrupted(
                "PlcfBkl has invalid or non-monotonic CPs".to_string(),
            ));
        }

        let mut used_end_indexes = HashSet::with_capacity(names.len());
        let mut bookmarks = Vec::with_capacity(names.len());
        for (index, name) in names.into_iter().enumerate() {
            let property = starts
                .property(index)
                .ok_or_else(|| PackageError::Corrupted("PlcfBkf is missing an FBKF".to_string()))?;
            let end_index = usize::from(read_u16(property, 0, "bookmark ibkl")?);
            if end_index >= ends.len() - 1 || !used_end_indexes.insert(end_index) {
                return Err(PackageError::Corrupted(
                    "bookmark ibkl values must be unique and in range".to_string(),
                ));
            }
            let bkc = read_u16(property, 2, "bookmark BKC")?;
            if bkc & 0x0080 != 0 {
                return Err(PackageError::Corrupted(
                    "bookmark BKC fPub must be zero".to_string(),
                ));
            }
            let start = starts.position(index).ok_or_else(|| {
                PackageError::Corrupted("PlcfBkf is missing a start CP".to_string())
            })?;
            let end = ends[end_index];
            if start > end {
                return Err(PackageError::Corrupted(
                    "bookmark start CP exceeds its end CP".to_string(),
                ));
            }
            let column_range = if bkc & 0x8000 != 0 {
                let first = (bkc & 0x007F) as u8;
                let limit = ((bkc >> 8) & 0x003F) as u8;
                if first >= limit {
                    return Err(PackageError::Corrupted(
                        "bookmark BKC column range is empty or reversed".to_string(),
                    ));
                }
                Some((first, limit))
            } else {
                None
            };
            bookmarks.push(Bookmark {
                name,
                start,
                end,
                is_native: bkc & 0x4000 != 0,
                column_range,
            });
        }

        Ok(Self { bookmarks })
    }

    /// All standard bookmarks in start-CP order.
    pub fn bookmarks(&self) -> &[Bookmark] {
        &self.bookmarks
    }
}

fn parse_names(data: &[u8]) -> Result<Vec<String>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbfBkmk fExtend")? != 0xFFFF
        || read_u16(data, 4, "SttbfBkmk cbExtra")? != 0
    {
        return Err(PackageError::Corrupted(
            "SttbfBkmk has an invalid header".to_string(),
        ));
    }
    let count = usize::from(read_u16(data, 2, "SttbfBkmk count")?);
    if count > 0x3FFB {
        return Err(PackageError::Corrupted(
            "SttbfBkmk contains too many names".to_string(),
        ));
    }
    let mut offset = 6usize;
    let mut names = Vec::with_capacity(count);
    let mut unique = HashSet::with_capacity(count);
    for _ in 0..count {
        let length = usize::from(read_u16(data, offset, "bookmark name length")?);
        if length == 0 || length >= 40 {
            return Err(PackageError::Corrupted(
                "bookmark names must contain 1 through 39 UTF-16 characters".to_string(),
            ));
        }
        offset = offset
            .checked_add(2)
            .ok_or_else(|| PackageError::Corrupted("bookmark name offset overflows".to_string()))?;
        let byte_length = length
            .checked_mul(2)
            .ok_or_else(|| PackageError::Corrupted("bookmark name length overflows".to_string()))?;
        let end = offset
            .checked_add(byte_length)
            .ok_or_else(|| PackageError::Corrupted("bookmark name range overflows".to_string()))?;
        let bytes = data
            .get(offset..end)
            .ok_or_else(|| PackageError::Corrupted("bookmark name is truncated".to_string()))?;
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let name = String::from_utf16(&units)
            .map_err(|_| PackageError::Corrupted("bookmark name is invalid UTF-16".to_string()))?;
        if !unique.insert(name.clone()) {
            return Err(PackageError::Corrupted(
                "bookmark names must be unique".to_string(),
            ));
        }
        names.push(name);
        offset = end;
    }
    if offset != data.len() {
        return Err(PackageError::Corrupted(
            "SttbfBkmk contains trailing bytes".to_string(),
        ));
    }
    Ok(names)
}

fn validate_cps(plcf: &Plcf<'_>, document_end: u32, name: &str) -> Result<()> {
    // Every CP except the last must lie within the document parts and be
    // monotonic. The final CP of a bookmark PLC is ignored per [MS-DOC]
    // 2.8.10, so no constraint is placed on it.
    let mut previous = None;
    for index in 0..plcf.count() {
        let cp = plcf
            .position(index)
            .ok_or_else(|| PackageError::Corrupted(format!("{name} is missing a CP")))?;
        if cp > document_end || previous.is_some_and(|value| value > cp) {
            return Err(PackageError::Corrupted(format!(
                "{name} has out-of-range or non-monotonic CPs"
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
        .ok_or_else(|| PackageError::Corrupted(format!("{name} is missing")))?;
    let start = usize::try_from(offset)
        .map_err(|_| PackageError::Corrupted(format!("{name} offset is too large")))?;
    let length = usize::try_from(length)
        .map_err(|_| PackageError::Corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| PackageError::Corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| PackageError::Corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| PackageError::Corrupted(format!("invalid {field}: {error}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn bookmark_names(names: &[&str]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&(names.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for name in names {
            let units = name.encode_utf16().collect::<Vec<_>>();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        data
    }

    fn fixture() -> (FileInformationBlock, Vec<u8>, usize) {
        let mut fib_data = vec![0; 154 + 93 * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes());
        fib_data[152..154].copy_from_slice(&93u16.to_le_bytes());

        let mut table = Vec::new();
        let names = bookmark_names(&["Outer", "_Cell"]);
        let names_offset = table.len() as u32;
        table.extend_from_slice(&names);
        set_fib_pointer(&mut fib_data, 21, names_offset, names.len() as u32);

        let starts_offset = table.len();
        for cp in [1u32, 2, 11] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&1u16.to_le_bytes());
        table.extend_from_slice(&0x4000u16.to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        table.extend_from_slice(&0x8301u16.to_le_bytes());
        set_fib_pointer(&mut fib_data, 22, starts_offset as u32, 20);

        let ends_offset = table.len() as u32;
        for cp in [5u32, 8, 11] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        set_fib_pointer(&mut fib_data, 23, ends_offset, 12);

        (
            FileInformationBlock::parse(&fib_data).unwrap(),
            table,
            starts_offset,
        )
    }

    #[test]
    fn parses_overlapping_and_column_bookmarks() {
        let (fib, table, _) = fixture();
        let parsed = BookmarksTable::parse(&fib, &table).unwrap();
        assert_eq!(
            parsed.bookmarks(),
            [
                Bookmark {
                    name: "Outer".to_string(),
                    start: 1,
                    end: 8,
                    is_native: true,
                    column_range: None,
                },
                Bookmark {
                    name: "_Cell".to_string(),
                    start: 2,
                    end: 5,
                    is_native: false,
                    column_range: Some((1, 3)),
                },
            ]
        );
        assert!(parsed.bookmarks()[1].is_hidden());
    }

    #[test]
    fn rejects_malformed_standard_bookmark_tables() {
        let (fib, table, starts_offset) = fixture();

        // The final CP is ignored per [MS-DOC] 2.8.10, so an unexpected
        // value there must not reject the document.
        let mut ignored_final_cp = table.clone();
        ignored_final_cp[starts_offset + 8..starts_offset + 12]
            .copy_from_slice(&10u32.to_le_bytes());
        assert!(BookmarksTable::parse(&fib, &ignored_final_cp).is_ok());

        let mut out_of_range_start = table.clone();
        out_of_range_start[starts_offset..starts_offset + 4].copy_from_slice(&11u32.to_le_bytes());
        assert!(BookmarksTable::parse(&fib, &out_of_range_start).is_err());

        let mut duplicate_ibkl = table.clone();
        duplicate_ibkl[starts_offset + 16..starts_offset + 18].copy_from_slice(&1u16.to_le_bytes());
        assert!(BookmarksTable::parse(&fib, &duplicate_ibkl).is_err());

        let mut public = table.clone();
        public[starts_offset + 14] |= 0x80;
        assert!(BookmarksTable::parse(&fib, &public).is_err());

        let mut reversed_columns = table.clone();
        reversed_columns[starts_offset + 18..starts_offset + 20]
            .copy_from_slice(&0x8102u16.to_le_bytes());
        assert!(BookmarksTable::parse(&fib, &reversed_columns).is_err());

        let duplicate_names = bookmark_names(&["Same", "Same"]);
        assert!(parse_names(&duplicate_names).is_err());
    }
}
