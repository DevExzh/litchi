//! FIB/table-stream seam for `RgxOcxInfo`.

use super::codec::parse_bytes;
use super::model::RgxOcxInfo;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;

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
        if offset != 0 {
            return Err(corrupted("fcPlcocx must be zero when lcbPlcocx is zero"));
        }
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

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
