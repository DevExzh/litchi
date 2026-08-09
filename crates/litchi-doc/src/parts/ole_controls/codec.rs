//! Exact little-endian codecs for `OcxInfo`, `RgxOcxInfo`, and `ObjectPool` `ODT`.

use super::model::{Format, Metadata, OcxInfo, Persist1, Persist2, RgxOcxInfo, Story};
use super::validation;
use crate::package::{Error as PackageError, Result};

/// The fixed serialized size of one `OcxInfo` record.
pub(super) const OCX_INFO_SIZE: usize = 20;
const ARRAY_HEADER_SIZE: usize = 4;
const ODT_MIN_SIZE: usize = 4;
const ODT_MAX_SIZE: usize = 6;

/// Decode one complete `RgxOcxInfo` payload.
pub fn parse_bytes(data: &[u8]) -> Result<RgxOcxInfo> {
    let count = usize::try_from(read_u32(data, 0, "RgxOcxInfo cOcxInfo")?)
        .map_err(|_| corrupted("RgxOcxInfo count exceeds usize"))?;
    validation::table_size(count, data.len(), OCX_INFO_SIZE)?;

    let mut infos = Vec::with_capacity(count);
    for index in 0..count {
        let offset = index
            .checked_mul(OCX_INFO_SIZE)
            .and_then(|value| value.checked_add(ARRAY_HEADER_SIZE))
            .ok_or_else(|| corrupted("OcxInfo offset overflows"))?;
        if let Some(info) = decode_info(data, offset)? {
            infos.push(info);
        }
    }
    validation::infos(&infos)?;
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

/// Decode one complete `ObjectPool` `ODT`/`ObjInfo` stream.
pub fn parse_metadata(data: &[u8]) -> Result<Metadata> {
    if data.len() != ODT_MIN_SIZE && data.len() != ODT_MAX_SIZE {
        return Err(corrupted("ObjectPool ObjInfo ODT must be 4 or 6 bytes"));
    }
    let persist1 = Persist1::from_raw(read_u16(data, 0, "ODTPersist1")?)?;
    let format = Format::from_raw(read_u16(data, 2, "ODT cf")?)?;
    let persist2 = if data.len() == ODT_MAX_SIZE {
        Some(Persist2::from_raw(read_u16(data, 4, "ODTPersist2")?)?)
    } else {
        None
    };
    Metadata::try_new(persist1, format, persist2)
}

/// Encode one complete `ObjectPool` `ODT`/`ObjInfo` stream.
pub fn to_metadata_bytes(metadata: &Metadata) -> Result<Vec<u8>> {
    metadata.validate()?;
    let mut data = Vec::with_capacity(if metadata.persist2().is_some() { 6 } else { 4 });
    data.extend_from_slice(&metadata.persist1().raw().to_le_bytes());
    data.extend_from_slice(&metadata.format().raw().to_le_bytes());
    if let Some(persist2) = metadata.persist2() {
        data.extend_from_slice(&persist2.raw().to_le_bytes());
    }
    Ok(data)
}

impl RgxOcxInfo {
    /// Decode one complete `RgxOcxInfo` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        parse_bytes(data)
    }

    /// Encode one complete `RgxOcxInfo` payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        to_bytes(self)
    }
}

impl Metadata {
    /// Decode one complete `ObjectPool` `ODT`/`ObjInfo` stream.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        parse_metadata(data)
    }

    /// Encode one complete `ObjectPool` `ODT`/`ObjInfo` stream.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        to_metadata_bytes(self)
    }
}

/// Whether one serialized record is an unwritten array slot: some producers
/// pad the `RgxOcxInfo` array with a marker whose `dwCookie`/`ifld` are all
/// ones and every remaining byte is zero. The slot references no control and
/// cannot satisfy MS-DOC 2.9.161 (`fifld` MUST be 1, `idoc` MUST be a story),
/// so readers skip it the way Word and other implementations do.
fn is_unwritten_slot(record: &[u8]) -> bool {
    record[..8].iter().all(|byte| *byte == 0xFF) && record[8..].iter().all(|byte| *byte == 0)
}

fn decode_info(data: &[u8], offset: usize) -> Result<Option<OcxInfo>> {
    let record = data
        .get(
            offset
                ..offset
                    .checked_add(OCX_INFO_SIZE)
                    .ok_or_else(|| corrupted("OcxInfo record range overflows"))?,
        )
        .ok_or_else(|| corrupted("OcxInfo record is truncated"))?;
    if is_unwritten_slot(record) {
        return Ok(None);
    }
    let flags = super::model::Flags::from_raw(read_u16(record, 14, "OcxInfo flags")?)?;
    let story = Story::from_raw(read_u16(record, 16, "OcxInfo idoc")?)?;
    Ok(Some(OcxInfo::new(
        read_u32(record, 0, "OcxInfo dwCookie")?,
        read_u32(record, 4, "OcxInfo ifld")?,
        read_u32(record, 8, "OcxInfo hAccel")?,
        read_u16(record, 12, "OcxInfo cAccel")?,
        flags,
        story,
        read_u16(record, 18, "OcxInfo reserved2")?,
    )))
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
