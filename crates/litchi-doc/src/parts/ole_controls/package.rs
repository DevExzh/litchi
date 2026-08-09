//! FIB/table-stream seam for `RgxOcxInfo`.

use super::codec::{parse_bytes, parse_metadata};
use super::model::{
    EPRINT_STREAM, Entry, Metadata, OBJ_INFO_STREAM, OBJECT_POOL_STORAGE, OCX_DATA_STREAM,
    ObjectPool, PRINT_STREAM, RgxOcxInfo, StorageName,
};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::{Object, Objects};

/// `FibRgFcLcb97` index of `fcPlcocx`/`lcbPlcocx` (MS-DOC 2.5).
pub const FIB_INDEX_PLC_OCX: usize = 85;

/// Parse the optional OLE-control metadata addressed by the FIB.
///
/// This seam reads only the table-stream range. It does not inspect the
/// referenced fields or resolve any OLE storage.
pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<RgxOcxInfo>> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_PLC_OCX) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }

    let start = usize::try_from(offset).map_err(|_| corrupted("fcPlcocx offset exceeds usize"))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted("lcbPlcocx length exceeds usize"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("fcPlcocx range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("RgxOcxInfo extends beyond the table stream"))?;
    parse_bytes(data).map(Some)
}

/// Parse inert metadata from selected `ObjectPool` storages.
///
/// The common OLE layer supplies only captured storage and stream bytes. This
/// DOC owner interprets the `\x03ObjInfo` and optional `\x03OCXDATA` names
/// without opening, activating, rendering, or executing the object payload.
pub fn parse_pool(objects: &Objects) -> Result<ObjectPool> {
    let mut entries = Vec::with_capacity(objects.len());
    for object in objects {
        entries.push(parse_object(object)?);
    }
    ObjectPool::try_new(entries)
}

/// Parse one selected `ObjectPool` storage into passive metadata.
pub fn parse_object(object: &Object) -> Result<Entry> {
    let path = object.path();
    if path.len() != 2 || path.first().map(String::as_str) != Some(OBJECT_POOL_STORAGE) {
        return Err(corrupted(
            "selected OLE object is not a direct ObjectPool storage",
        ));
    }
    let name = StorageName::try_new(path[1].clone())?;
    let metadata = parse_metadata(
        object
            .stream(&[OBJ_INFO_STREAM])
            .ok_or_else(|| corrupted("ObjectPool storage is missing its ObjInfo stream"))?,
    )?;
    let control_data_present = object.stream(&[OCX_DATA_STREAM]).is_some();
    let print_present = object.stream(&[PRINT_STREAM]).is_some();
    let enhanced_print_present = object.stream(&[EPRINT_STREAM]).is_some();
    Entry::try_with_streams(
        name,
        object.storage().clsid().map(str::to_owned),
        Some(metadata),
        control_data_present,
        print_present,
        enhanced_print_present,
    )
}

/// Parse one `ObjInfo` stream without requiring a complete CFB object.
pub fn parse_object_metadata(data: &[u8]) -> Result<Metadata> {
    parse_metadata(data)
}

impl RgxOcxInfo {
    /// Parse the optional `fcPlcocx`/`lcbPlcocx` table range.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        parse(fib, table_stream)
    }
}

impl ObjectPool {
    /// Inspect selected `ObjectPool` storages without opening their payloads.
    pub fn parse(objects: &Objects) -> Result<Self> {
        parse_pool(objects)
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
