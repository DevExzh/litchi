//! Word `SttbListNames` table for named `LISTNUM` list definitions.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use std::collections::HashSet;

const FIB_INDEX: usize = 91;
const MAX_NAME_UNITS: usize = 255;
const MAX_NAME_COUNT: usize = u16::MAX as usize;
const MAX_TABLE_BYTES: usize = 6 + MAX_NAME_COUNT * (2 + MAX_NAME_UNITS * 2);

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Ordered names parallel to the document's `PlfLst.rgLstf` array.
///
/// Empty strings represent unnamed list definitions. Non-empty names are unique.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ListNamesTable {
    names: Vec<String>,
}

impl ListNamesTable {
    pub fn try_new(names: Vec<String>) -> Result<Self> {
        validate_names(&names)?;
        Ok(Self { names })
    }

    /// Parse the optional FIB index-91 table range.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start =
            usize::try_from(offset).map_err(|_| corrupted("SttbListNames offset is too large"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("SttbListNames length is too large"))?;
        if length > MAX_TABLE_BYTES {
            return Err(corrupted(
                "SttbListNames exceeds its specification-derived size cap",
            ));
        }
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("SttbListNames range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("SttbListNames extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `SttbListNames` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_TABLE_BYTES {
            return Err(corrupted(
                "SttbListNames exceeds its specification-derived size cap",
            ));
        }
        if data.len() < 6
            || read_u16(data, 0, "SttbListNames fExtend")? != 0xFFFF
            || read_u16(data, 4, "SttbListNames cbExtra")? != 0
        {
            return Err(corrupted("SttbListNames has an invalid header"));
        }

        let count = usize::from(read_u16(data, 2, "SttbListNames cData")?);
        let mut names = Vec::with_capacity(count);
        let mut offset = 6usize;
        for index in 0..count {
            let unit_count = usize::from(read_u16(
                data,
                offset,
                &format!("SttbListNames string {index} length"),
            )?);
            if unit_count > MAX_NAME_UNITS {
                return Err(corrupted(format!(
                    "SttbListNames string {index} exceeds 255 UTF-16 code units"
                )));
            }
            offset = offset
                .checked_add(2)
                .ok_or_else(|| corrupted("SttbListNames string offset overflows"))?;
            let byte_count = unit_count
                .checked_mul(2)
                .ok_or_else(|| corrupted("SttbListNames string length overflows"))?;
            let end = offset
                .checked_add(byte_count)
                .ok_or_else(|| corrupted("SttbListNames string range overflows"))?;
            let bytes = data
                .get(offset..end)
                .ok_or_else(|| corrupted(format!("SttbListNames string {index} is truncated")))?;
            let units = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            names.push(String::from_utf16(&units).map_err(|_| {
                corrupted(format!(
                    "SttbListNames string {index} contains invalid UTF-16"
                ))
            })?);
            offset = end;
        }
        if offset != data.len() {
            return Err(corrupted("SttbListNames has trailing bytes"));
        }
        Self::try_new(names)
    }

    pub fn len(&self) -> usize {
        self.names.len()
    }
    pub fn is_empty(&self) -> bool {
        self.names.is_empty()
    }
    pub fn entries(&self) -> &[String] {
        &self.names
    }
    pub fn get(&self, index: usize) -> Option<&str> {
        self.names.get(index).map(String::as_str)
    }

    /// Return a non-empty name at an index; empty slots are treated as unnamed.
    pub fn name(&self, index: usize) -> Option<&str> {
        self.get(index).filter(|name| !name.is_empty())
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = &str> {
        self.names.iter().map(String::as_str)
    }

    /// Serialize the complete STTB deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let size = validate_names(&self.names)?;
        let count = u16::try_from(self.names.len())
            .map_err(|_| corrupted("SttbListNames contains more than 65535 strings"))?;
        let mut data = Vec::with_capacity(size);
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&count.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for name in &self.names {
            let units = name.encode_utf16().collect::<Vec<_>>();
            let unit_count = u16::try_from(units.len())
                .map_err(|_| corrupted("SttbListNames string length exceeds u16"))?;
            data.extend_from_slice(&unit_count.to_le_bytes());
            data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        Ok(data)
    }
}

fn validate_names(names: &[String]) -> Result<usize> {
    if names.len() > MAX_NAME_COUNT {
        return Err(corrupted("SttbListNames contains more than 65535 strings"));
    }
    let mut unique = HashSet::with_capacity(names.len());
    let mut size = 6usize;
    for (index, name) in names.iter().enumerate() {
        let units = name.encode_utf16().count();
        if units > MAX_NAME_UNITS {
            return Err(corrupted(format!(
                "SttbListNames string {index} exceeds 255 UTF-16 code units"
            )));
        }
        if !name.is_empty() && !unique.insert(name.as_str()) {
            return Err(corrupted(format!(
                "SttbListNames contains duplicate non-empty name {name:?}"
            )));
        }
        size = size
            .checked_add(2)
            .and_then(|value| value.checked_add(units.checked_mul(2)?))
            .ok_or_else(|| corrupted("SttbListNames serialized size overflows"))?;
    }
    if size > MAX_TABLE_BYTES {
        return Err(corrupted(
            "SttbListNames exceeds its specification-derived size cap",
        ));
    }
    Ok(size)
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_REFERENCE_HEX: &str = concat!(
        "ffff04000000",
        "07005700570034004e0075006d003100",
        "07005700570031004e0075006d003900",
        "08005700570038004e0075006d0031003100",
        "0000",
    );

    fn decode_hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| {
                let digit = |byte: u8| match byte {
                    b'0'..=b'9' => byte - b'0',
                    b'a'..=b'f' => byte - b'a' + 10,
                    _ => panic!("invalid test hex"),
                };
                digit(pair[0]) << 4 | digit(pair[1])
            })
            .collect()
    }

    #[test]
    fn parses_and_round_trips_poi_reference_table() {
        let bytes = decode_hex(POI_REFERENCE_HEX);
        let table = ListNamesTable::parse_bytes(&bytes).unwrap();
        assert_eq!(table.entries(), &["WW4Num1", "WW1Num9", "WW8Num11", ""]);
        assert_eq!(table.name(2), Some("WW8Num11"));
        assert_eq!(table.name(3), None);
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn preserves_empty_parallel_slots_and_utf16() {
        let table = ListNamesTable::try_new(vec![
            String::new(),
            "List alpha".to_string(),
            String::new(),
            "\u{1f4cb}".to_string(),
        ])
        .unwrap();
        assert_eq!(
            ListNamesTable::parse_bytes(&table.to_bytes().unwrap()).unwrap(),
            table
        );
    }

    #[test]
    fn rejects_malformed_headers_lengths_and_encoding() {
        let reference = decode_hex(POI_REFERENCE_HEX);
        let mut bytes = reference.clone();
        bytes[0] = 0;
        assert!(ListNamesTable::parse_bytes(&bytes).is_err());
        let mut bytes = reference.clone();
        bytes[4] = 1;
        assert!(ListNamesTable::parse_bytes(&bytes).is_err());
        let mut bytes = reference.clone();
        bytes[6..8].copy_from_slice(&256u16.to_le_bytes());
        assert!(ListNamesTable::parse_bytes(&bytes).is_err());
        assert!(ListNamesTable::parse_bytes(&reference[..reference.len() - 1]).is_err());
        let mut bytes = reference;
        bytes.push(0);
        assert!(ListNamesTable::parse_bytes(&bytes).is_err());
        assert!(ListNamesTable::parse_bytes(&[0xff, 0xff, 1, 0, 0, 0, 1, 0, 0, 0xd8]).is_err());
    }

    #[test]
    fn rejects_duplicate_and_oversized_names() {
        assert!(ListNamesTable::try_new(vec!["same".into(), "same".into()]).is_err());
        assert!(ListNamesTable::try_new(vec!["x".repeat(256)]).is_err());
        assert!(ListNamesTable::try_new(vec![String::new(), String::new()]).is_ok());
    }
}
