//! Primitive `PowerPoint` record framing helpers.

use crate::embedded::object::editor::Result;
use crate::package::Error;

pub(crate) fn slice(data: &[u8], offset: usize) -> Result<&[u8]> {
    let len = u32_at(data, offset + 4)? as usize;
    let end = offset
        .checked_add(8)
        .and_then(|value| value.checked_add(len))
        .ok_or_else(|| Error::Corrupted("record length overflow".into()))?;
    data.get(offset..end)
        .ok_or_else(|| Error::Corrupted("truncated record".into()))
}

pub(crate) fn type_of(record: &[u8]) -> Result<u16> {
    record
        .get(2..4)
        .map(|value| u16::from_le_bytes([value[0], value[1]]))
        .ok_or_else(|| Error::Corrupted("truncated header".into()))
}

pub(crate) fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    let bytes: &[u8; 4] = data
        .get(offset..offset + 4)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| Error::Corrupted("truncated u32".into()))?;
    Ok(u32::from_le_bytes(*bytes))
}
