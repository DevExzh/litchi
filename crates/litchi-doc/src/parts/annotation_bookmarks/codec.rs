//! Binary codec for `SttbfAtnBkmk` and fixed-size `ATNBE` records.

use super::model::{Tag, TagId, Tags};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;

/// Parse the optional FIB-addressed annotation-bookmark table.
pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Tags>> {
    let Some((offset, length)) = fib
        .get_table_pointer(validation::FIB_INDEX)
        .filter(|(_, length)| *length != 0)
    else {
        return Ok(None);
    };
    let data = validation::table_range(table_stream, offset, length)?;
    parse_bytes(data).map(Some)
}

/// Parse one complete `SttbfAtnBkmk` payload.
pub fn parse_bytes(data: &[u8]) -> Result<Tags> {
    if data.len() > validation::MAX_TABLE_BYTES {
        return Err(corrupted(
            "SttbfAtnBkmk exceeds its specification-derived size cap",
        ));
    }
    if data.len() < 6 {
        return Err(corrupted("SttbfAtnBkmk header is truncated"));
    }
    if read_u16(data, 0, "SttbfAtnBkmk fExtend")? != 0xFFFF {
        return Err(corrupted("SttbfAtnBkmk fExtend must be 0xFFFF"));
    }
    let count = usize::from(read_u16(data, 2, "SttbfAtnBkmk cData")?);
    if count > validation::MAX_ENTRIES {
        return Err(corrupted("SttbfAtnBkmk cData exceeds 0x3FFC entries"));
    }
    if read_u16(data, 4, "SttbfAtnBkmk cbExtra")? != validation::ATNBE_SIZE as u16 {
        return Err(corrupted("SttbfAtnBkmk cbExtra must be 0x000A"));
    }

    let mut offset = 6usize;
    let mut entries = Vec::with_capacity(count);
    for index in 0..count {
        if read_u16(data, offset, "SttbfAtnBkmk cchData")? != 0 {
            return Err(corrupted(format!(
                "SttbfAtnBkmk entry {index} cchData must be zero"
            )));
        }
        offset = offset
            .checked_add(2)
            .ok_or_else(|| corrupted("SttbfAtnBkmk entry offset overflows"))?;
        let end = offset
            .checked_add(validation::ATNBE_SIZE)
            .ok_or_else(|| corrupted("ATNBE range overflows"))?;
        let record = data
            .get(offset..end)
            .ok_or_else(|| corrupted(format!("ATNBE entry {index} is truncated")))?;
        if read_u16(record, 0, "ATNBE bmc")? != validation::BMC_ANNOTATION {
            return Err(corrupted(format!("ATNBE entry {index} bmc must be 0x0100")));
        }
        let tag = read_u32(record, 2, "ATNBE lTag")?;
        if read_i32(record, 6, "ATNBE lTagOld")? != -1 {
            return Err(corrupted(format!("ATNBE entry {index} lTagOld must be -1")));
        }
        entries.push(Tag::new(TagId::new(tag)));
        offset = end;
    }
    if offset != data.len() {
        return Err(corrupted("SttbfAtnBkmk contains trailing bytes"));
    }
    Tags::try_new(entries)
}

/// Serialize one complete `SttbfAtnBkmk` payload.
pub fn to_bytes(value: &Tags) -> Result<Vec<u8>> {
    validation::tags(value)?;
    let capacity = 6usize
        .checked_add(
            value
                .len()
                .checked_mul(2 + validation::ATNBE_SIZE)
                .ok_or_else(|| corrupted("SttbfAtnBkmk serialized size overflows"))?,
        )
        .ok_or_else(|| corrupted("SttbfAtnBkmk serialized size overflows"))?;
    let mut data = Vec::with_capacity(capacity);
    data.extend_from_slice(&0xFFFFu16.to_le_bytes());
    data.extend_from_slice(
        &u16::try_from(value.len())
            .map_err(|_| corrupted("SttbfAtnBkmk cData exceeds u16::MAX"))?
            .to_le_bytes(),
    );
    data.extend_from_slice(&(validation::ATNBE_SIZE as u16).to_le_bytes());
    for entry in value.entries() {
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&validation::BMC_ANNOTATION.to_le_bytes());
        data.extend_from_slice(&entry.id().raw().to_le_bytes());
        data.extend_from_slice(&(-1i32).to_le_bytes());
    }
    Ok(data)
}

impl Tags {
    /// Parse the optional FIB-addressed table.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        parse(fib, table_stream)
    }

    /// Parse one complete table payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        parse_bytes(data)
    }

    /// Serialize one complete table payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        to_bytes(self)
    }
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
    Ok(u32::from_le_bytes(
        bytes.try_into().expect("validated u32 width"),
    ))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    Ok(read_u32(data, offset, field)? as i32)
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
