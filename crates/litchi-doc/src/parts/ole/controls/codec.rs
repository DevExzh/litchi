//! MS-DOC `RgxOcxInfo` table codec.

use super::model::{Control, Controls};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use std::collections::HashSet;

/// Table-pointer index of `fcPlcocx`/`lcbPlcocx` (MS-DOC 2.9.229).
pub(super) const PLCOCX: usize = 85;

/// Size of the `dwCookie` prefix in one `OcxInfo` structure (MS-DOC 2.9.161).
const OCX_INFO_SIZE: usize = 4;

impl Controls {
    /// Parse the `RgxOcxInfo` addressed by the FIB, or `None` when the
    /// document contains no OLE controls.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Controls>> {
        let Some((offset, length)) = fib.get_table_pointer(PLCOCX) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }

        let start =
            usize::try_from(offset).map_err(|_| corrupted("RgxOcxInfo offset exceeds usize"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("RgxOcxInfo length exceeds usize"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("RgxOcxInfo range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("RgxOcxInfo extends beyond the table stream"))?;
        Ok(Some(Self::parse_bytes(data)?))
    }

    /// Parse one `RgxOcxInfo` structure.
    ///
    /// MS-DOC defines `OcxInfo` as a 4-byte `dwCookie` and states that the
    /// data SHOULD be ignored; Word 2003+ emits larger entries for backward
    /// compatibility (MS-DOC product-behavior note <228>). Parsing therefore
    /// accepts any evenly padded per-entry stride and reads `dwCookie` at its
    /// start. All trailing bytes remain intentionally inert.
    pub fn parse_bytes(data: &[u8]) -> Result<Controls> {
        let count = usize::try_from(read_u32(data, 0, "RgxOcxInfo cOcxInfo")?)
            .map_err(|_| corrupted("RgxOcxInfo count exceeds usize"))?;
        let payload = data
            .get(OCX_INFO_SIZE..)
            .ok_or_else(|| corrupted("RgxOcxInfo header is truncated"))?;

        if count == 0 {
            if !payload.is_empty() {
                return Err(corrupted("RgxOcxInfo byte length does not match its count"));
            }
            return Ok(Controls::from_controls(Vec::new()));
        }

        // Check the minimum required bytes before reserving. This keeps the
        // allocation bounded by the supplied table even for hostile counts.
        if count > payload.len() / OCX_INFO_SIZE || !payload.len().is_multiple_of(count) {
            return Err(corrupted("RgxOcxInfo byte length does not match its count"));
        }
        let stride = payload.len() / count;
        if stride < OCX_INFO_SIZE {
            return Err(corrupted("RgxOcxInfo byte length does not match its count"));
        }

        let mut controls = Vec::with_capacity(count);
        let mut cookies = HashSet::with_capacity(count);
        for index in 0..count {
            let entry_offset = OCX_INFO_SIZE
                .checked_add(
                    index
                        .checked_mul(stride)
                        .ok_or_else(|| corrupted("OcxInfo entry offset overflows"))?,
                )
                .ok_or_else(|| corrupted("OcxInfo entry offset overflows"))?;
            let cookie = read_u32(data, entry_offset, "OcxInfo dwCookie")?;
            if !cookies.insert(cookie) {
                return Err(corrupted("OcxInfo dwCookie values must be unique"));
            }
            controls.push(Control { cookie });
        }
        Ok(Controls::from_controls(controls))
    }
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
