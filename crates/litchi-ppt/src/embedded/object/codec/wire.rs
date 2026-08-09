//! Common `PowerPoint` record framing and bounded primitive decoding.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

pub(crate) fn parse_bool(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => corrupted(format!("ExOleEmbedAtom {field} is not a bool1")),
    }
}

pub(crate) fn u32_at(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

pub(crate) fn require_atom(
    record: &Record,
    version: u16,
    instance: u16,
    kind: RecordType,
    length: usize,
    context: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.record_type_raw != kind.as_u16()
        || record.data.len() != length
        || usize::try_from(record.data_length).ok() != Some(length)
    {
        return corrupted(format!("{context} has an invalid header or size"));
    }
    Ok(())
}

pub(crate) fn record_bytes(
    version: u16,
    instance: u16,
    kind: RecordType,
    data: &[u8],
) -> Result<Vec<u8>> {
    record_bytes_raw(version, instance, kind.as_u16(), data)
}

pub(crate) fn record_bytes_raw(
    version: u16,
    instance: u16,
    kind: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    if version > 0x000f || instance > 0x0fff {
        return corrupted("PowerPoint record header exceeds its encoded domain");
    }
    let length = u32::try_from(data.len())
        .map_err(|_err| Error::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

pub(crate) fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
