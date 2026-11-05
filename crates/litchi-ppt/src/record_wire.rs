//! Byte-level `[MS-PPT]` record helpers shared by the element editors
//! (record 0323).
//!
//! Element edits rewrite one record path of an exact persisted record and
//! copy every other record byte for byte. These helpers frame records
//! without decoding unrelated payloads.

pub(crate) use crate::text_edit::{
    children, drawing_payload as payload, rebuild, record_instance, record_type, record_version,
    shape_id_of,
};

use crate::package::{Error, Result};

/// The eight-byte `RecordHeader` length.
pub(crate) const HEADER_LEN: usize = 8;

/// Encode one record: `version` (4 bits), `instance` (12 bits), type and
/// payload.
pub(crate) fn encode(
    record_type: u16,
    version: u16,
    instance: u16,
    payload: &[u8],
) -> Result<Vec<u8>> {
    let length = u32::try_from(payload.len())
        .map_err(|_error| Error::ResourceLimit("PPT record exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(HEADER_LEN.saturating_add(payload.len()));
    bytes.extend_from_slice(&((version & 0x0F) | (instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(payload);
    Ok(bytes)
}

/// Encode one container (version `0xF`) from its encoded children.
pub(crate) fn container(record_type: u16, instance: u16, children: &[Vec<u8>]) -> Result<Vec<u8>> {
    encode(record_type, 0x0F, instance, &children.concat())
}

/// UTF-16LE code units of `text`, without a terminator.
pub(crate) fn utf16(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// Collect the direct child records of `record`.
pub(crate) fn child_list(record: &[u8]) -> Result<Vec<&[u8]>> {
    children(record)?.collect()
}

/// Collect the records of a raw record sequence (a payload that is not
/// itself framed as a container, such as `BinaryTagData`).
pub(crate) fn sequence(bytes: &[u8]) -> Result<Vec<&[u8]>> {
    let mut records = Vec::new();
    let mut offset = 0;
    while offset < bytes.len() {
        let tail = &bytes[offset..];
        let length = payload(tail)?.len().saturating_add(HEADER_LEN);
        records.push(&tail[..length]);
        offset += length;
    }
    Ok(records)
}
