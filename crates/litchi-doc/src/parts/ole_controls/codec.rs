//! Binary codec for `OcxInfo` and `RgxOcxInfo`.

use super::model::{Flags, OCX_INFO_SIZE, OcxInfo, RgxOcxInfo, Story};
use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

const ARRAY_HEADER_SIZE: usize = 4;

/// Decode one complete `RgxOcxInfo` payload.
pub fn parse_bytes(data: &[u8]) -> Result<RgxOcxInfo> {
    let count = usize::try_from(read_u32(data, 0, "RgxOcxInfo cOcxInfo")?)
        .map_err(|_| corrupted("RgxOcxInfo count exceeds usize"))?;
    if data.len() < ARRAY_HEADER_SIZE {
        return Err(corrupted("RgxOcxInfo header is truncated"));
    }
    let expected = count
        .checked_mul(OCX_INFO_SIZE)
        .and_then(|size| size.checked_add(ARRAY_HEADER_SIZE))
        .ok_or_else(|| corrupted("RgxOcxInfo size overflows"))?;
    if expected != data.len() {
        return Err(corrupted(format!(
            "RgxOcxInfo requires {expected} bytes for {count} records, got {}",
            data.len()
        )));
    }

    let mut infos = Vec::with_capacity(count);
    let mut cookies = HashSet::with_capacity(count);
    for index in 0..count {
        let offset = index
            .checked_mul(OCX_INFO_SIZE)
            .and_then(|value| value.checked_add(ARRAY_HEADER_SIZE))
            .ok_or_else(|| corrupted("OcxInfo offset overflows"))?;
        let info = decode_info(data, offset)?;
        if !cookies.insert(info.cookie()) {
            return Err(corrupted("OcxInfo dwCookie values must be unique"));
        }
        infos.push(info);
    }
    Ok(RgxOcxInfo::from_infos(infos))
}

/// Encode one complete `RgxOcxInfo` payload.
pub fn to_bytes(table: &RgxOcxInfo) -> Result<Vec<u8>> {
    table.validate()?;
    let count =
        u32::try_from(table.len()).map_err(|_| corrupted("RgxOcxInfo count exceeds u32::MAX"))?;
    let capacity = table
        .len()
        .checked_mul(OCX_INFO_SIZE)
        .and_then(|size| size.checked_add(ARRAY_HEADER_SIZE))
        .ok_or_else(|| corrupted("RgxOcxInfo size overflows"))?;
    let mut data = Vec::with_capacity(capacity);
    data.extend_from_slice(&count.to_le_bytes());
    for info in table.infos() {
        encode_info(&mut data, *info);
    }
    Ok(data)
}

fn decode_info(data: &[u8], offset: usize) -> Result<OcxInfo> {
    let record = data
        .get(
            offset
                ..offset
                    .checked_add(OCX_INFO_SIZE)
                    .ok_or_else(|| corrupted("OcxInfo record range overflows"))?,
        )
        .ok_or_else(|| corrupted("OcxInfo record is truncated"))?;
    let flags = Flags::from_raw(read_u16(record, 14, "OcxInfo flags")?)?;
    let story = Story::from_raw(read_u16(record, 16, "OcxInfo idoc")?)?;
    Ok(OcxInfo::new(
        read_u32(record, 0, "OcxInfo dwCookie")?,
        read_u32(record, 4, "OcxInfo ifld")?,
        read_u32(record, 8, "OcxInfo hAccel")?,
        read_u16(record, 12, "OcxInfo cAccel")?,
        flags,
        story,
        read_u16(record, 18, "OcxInfo reserved2")?,
    ))
}

fn encode_info(data: &mut Vec<u8>, info: OcxInfo) {
    data.extend_from_slice(&info.cookie().to_le_bytes());
    data.extend_from_slice(&info.field_index().to_le_bytes());
    data.extend_from_slice(&info.accelerator_handle().to_le_bytes());
    data.extend_from_slice(&info.accelerator_count().to_le_bytes());
    data.extend_from_slice(&info.flags().raw().to_le_bytes());
    data.extend_from_slice(&info.story().raw().to_le_bytes());
    data.extend_from_slice(&info.reserved().to_le_bytes());
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
