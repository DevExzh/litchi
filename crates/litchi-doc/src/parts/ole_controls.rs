//! Legacy Word OLE control information (`RgxOcxInfo`, MS-DOC 2.9.229).
//!
//! The table lists the OLE controls (checkboxes, option buttons, and so on)
//! that a document contains. The data carried by each entry is explicitly
//! inert: MS-DOC specifies that consumers should ignore it, and this reader
//! only exposes the unique control cookies.
//!
//! No control is instantiated, no OLE object is activated, and no control
//! code is executed.

use super::fib::FileInformationBlock;
use crate::package::{DocError, Result};
use std::collections::HashSet;

/// Table-pointer index of `fcPlcocx`/`lcbPlcocx`.
const PLCOCX: usize = 85;

/// Size of one `OcxInfo` structure in bytes (MS-DOC 2.9.161).
const OCX_INFO_SIZE: usize = 4;

/// The OLE controls recorded in a document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentOleControls {
    controls: Vec<OleControlInfo>,
}

/// One `OcxInfo` entry (MS-DOC 2.9.161).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OleControlInfo {
    /// Unique index of this control within the document's `RgxOcxInfo`.
    pub cookie: u32,
}

impl DocumentOleControls {
    /// Parse the `RgxOcxInfo` addressed by the FIB, or `None` when the
    /// document contains no OLE controls.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentOleControls>> {
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
    /// compatibility (MS-DOC product-behavior note \<228\>). Parsing therefore
    /// accepts any per-entry stride and reads `dwCookie` at its start.
    pub fn parse_bytes(data: &[u8]) -> Result<DocumentOleControls> {
        let count = usize::try_from(read_u32(data, 0, "RgxOcxInfo cOcxInfo")?)
            .map_err(|_| corrupted("RgxOcxInfo count exceeds usize"))?;
        if count == 0 {
            if data.len() != 4 {
                return Err(corrupted("RgxOcxInfo byte length does not match its count"));
            }
            return Ok(Self {
                controls: Vec::new(),
            });
        }
        let stride = (data.len() - 4) / count;
        if !(data.len() - 4).is_multiple_of(count) || stride < OCX_INFO_SIZE {
            return Err(corrupted("RgxOcxInfo byte length does not match its count"));
        }
        let mut controls = Vec::with_capacity(count);
        let mut cookies = HashSet::with_capacity(count);
        for index in 0..count {
            let cookie = read_u32(data, 4 + index * stride, "OcxInfo dwCookie")?;
            if !cookies.insert(cookie) {
                return Err(corrupted("OcxInfo dwCookie values must be unique"));
            }
            controls.push(OleControlInfo { cookie });
        }
        Ok(Self { controls })
    }

    /// All recorded OLE controls, in table order.
    pub fn controls(&self) -> &[OleControlInfo] {
        &self.controls
    }
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIB_POINTERS: usize = 93;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn rgx_ocx_info(cookies: &[u32]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&(cookies.len() as u32).to_le_bytes());
        for cookie in cookies {
            data.extend_from_slice(&cookie.to_le_bytes());
        }
        data
    }

    fn fixture(cookies: &[u32]) -> (FileInformationBlock, Vec<u8>) {
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());

        let table = rgx_ocx_info(cookies);
        set_fib_pointer(&mut fib_data, PLCOCX, 0, table.len() as u32);
        (FileInformationBlock::parse(&fib_data).unwrap(), table)
    }

    #[test]
    fn parses_ole_controls() {
        let (fib, table) = fixture(&[3, 0, 17]);
        let parsed = DocumentOleControls::parse(&fib, &table)
            .unwrap()
            .expect("controls present");
        assert_eq!(
            parsed.controls(),
            [
                OleControlInfo { cookie: 3 },
                OleControlInfo { cookie: 0 },
                OleControlInfo { cookie: 17 },
            ]
        );
    }

    #[test]
    fn absent_or_empty_table_yields_none() {
        // No `fcPlcocx` pointer at all.
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentOleControls::parse(&fib, &[]).unwrap().is_none());

        // A zero-length pointer is ignored per MS-DOC 2.5 (fcPlcocx).
        let (fib, table) = fixture(&[]);
        let mut fib_data = fib.raw_data().to_vec();
        set_fib_pointer(&mut fib_data, PLCOCX, 0, 0);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentOleControls::parse(&fib, &table).unwrap().is_none());
    }

    #[test]
    fn parses_empty_control_list() {
        let (fib, table) = fixture(&[]);
        let parsed = DocumentOleControls::parse(&fib, &table)
            .unwrap()
            .expect("empty table present");
        assert!(parsed.controls().is_empty());
    }

    #[test]
    fn accepts_word_padded_entries() {
        // Word 2003+ pads each `OcxInfo` beyond `dwCookie`; here 20 bytes per
        // entry, as emitted for the travel-form fixture.
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
        data.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0xFF]);
        data.extend_from_slice(&[0; 12]);
        let parsed = DocumentOleControls::parse_bytes(&data).unwrap();
        assert_eq!(
            parsed.controls(),
            [OleControlInfo {
                cookie: 0xFFFF_FFFF
            }]
        );
    }

    #[test]
    fn rejects_malformed_tables() {
        // Duplicate cookies.
        assert!(DocumentOleControls::parse_bytes(&rgx_ocx_info(&[3, 3])).is_err());

        // Declared count disagrees with the byte length.
        assert!(DocumentOleControls::parse_bytes(&rgx_ocx_info(&[3])[..6]).is_err());
        let mut misaligned = rgx_ocx_info(&[3, 4]);
        misaligned.pop();
        assert!(DocumentOleControls::parse_bytes(&misaligned).is_err());

        // A nonzero count with no entries.
        assert!(DocumentOleControls::parse_bytes(&1u32.to_le_bytes()).is_err());

        // A zero count with trailing bytes.
        assert!(DocumentOleControls::parse_bytes(&[0; 8]).is_err());

        // Truncated header.
        assert!(DocumentOleControls::parse_bytes(&[0, 0]).is_err());
    }
}
