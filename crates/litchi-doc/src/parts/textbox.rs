//! Textbox story parsing (PlcftxbxTxt / FTXBXS).
//!
//! The PlcftxbxTxt ([MS-DOC] 2.8.32) associates ranges of the textbox story
//! (the subdocument counted by ccpTxbx) with shape objects. Its data elements
//! are FTXBXS structures ([MS-DOC] 2.9.106); the last one is a reusable spare
//! and carries no shape.

use super::super::package::{Error as PackageError, Result};

/// Size of one FTXBXS structure in bytes.
pub const FTXBXS_LEN: usize = 22;

/// FIB index (into `FileInformationBlock::get_table_pointer`) of the
/// `fcPlcftxbxTxt`/`lcbPlcftxbxTxt` pair.
pub(crate) const FIB_INDEX_PLCF_TXBX_TXT: usize = 56;
/// FIB index of the `fcPlcfHdrtxbxTxt`/`lcbPlcfHdrtxbxTxt` pair (header
/// textbox story).
pub(crate) const FIB_INDEX_PLCF_HDR_TXBX_TXT: usize = 58;

/// A text box: the shape it belongs to and its text range within the
/// textbox story (story-relative character positions).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextBoxEntry {
    /// Shape identifier (`lid`); matches the `spid` of the shape's
    /// OfficeArtFSP and the `lid` of the shape's Spa.
    pub shape_id: u32,
    /// Story-relative start CP of this text box's text.
    pub start_cp: u32,
    /// Story-relative end CP of this text box's text.
    pub end_cp: u32,
}

/// A text box in a Word document: its shape and plain-text content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocTextBox {
    /// Shape identifier; matches the `spid` of the shape's OfficeArtFSP and
    /// the `lid` of the shape's Spa.
    pub shape_id: u32,
    /// Plain text content. Paragraphs are separated by '\r'; the trailing
    /// paragraph mark that terminates the text in the story is stripped.
    pub text: String,
    /// Which header the text box is anchored in, for header-story text
    /// boxes; `None` for main-story text boxes.
    pub header_kind: Option<super::headers::HeaderFooterType>,
}

/// Parse a PlcftxbxTxt, returning the real (non-reusable) text box entries
/// in story order.
pub fn parse_plcf_txbx_txt(data: &[u8]) -> Result<Vec<TextBoxEntry>> {
    // Standard PLC: n+1 CPs, n FTXBXS structures of FTXBXS_LEN bytes.
    let stride = 4 + FTXBXS_LEN;
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(stride) {
        return Err(PackageError::InvalidFormat(
            "Invalid PlcftxbxTxt length".to_string(),
        ));
    }
    let count = (data.len() - 4) / stride;
    let cps_start = 0;
    let entries_start = (count + 1) * 4;

    let mut entries = Vec::new();
    for index in 0..count {
        let base = entries_start + index * FTXBXS_LEN;
        let f_reusable = u16::from_le_bytes(data[base + 8..base + 10].try_into().unwrap_or([0; 2]));
        let lid = i32::from_le_bytes(data[base + 14..base + 18].try_into().unwrap_or([0; 4]));
        // Skip the reusable spare structures; they carry no shape (lid = 0).
        if f_reusable & 1 != 0 || lid <= 0 {
            continue;
        }
        let cp_at = |i: usize| {
            u32::from_le_bytes(
                data[cps_start + i * 4..cps_start + i * 4 + 4]
                    .try_into()
                    .unwrap_or([0; 4]),
            )
        };
        entries.push(TextBoxEntry {
            shape_id: lid as u32,
            start_cp: cp_at(index),
            end_cp: cp_at(index + 1),
        });
    }
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_plcf() -> Vec<u8> {
        let mut plcf = Vec::new();
        // Two real text boxes + one spare: aCPs [0, 3, 9, 10].
        for cp in [0u32, 3, 9, 10] {
            plcf.extend_from_slice(&cp.to_le_bytes());
        }
        let mut real = |lid: i32| {
            plcf.extend_from_slice(&1i32.to_le_bytes()); // cTxbx
            plcf.extend_from_slice(&0i32.to_le_bytes()); // cTxbxEdit
            plcf.extend_from_slice(&0u16.to_le_bytes()); // fReusable
            plcf.extend_from_slice(&(-1i32).to_le_bytes()); // itxbxsDest
            plcf.extend_from_slice(&lid.to_le_bytes()); // lid
            plcf.extend_from_slice(&0i32.to_le_bytes()); // txidUndo
        };
        real(1027);
        real(1031);
        // Spare: reusable, lid 0.
        plcf.extend_from_slice(&(-1i32).to_le_bytes());
        plcf.extend_from_slice(&0i32.to_le_bytes());
        plcf.extend_from_slice(&1u16.to_le_bytes());
        plcf.extend_from_slice(&(-1i32).to_le_bytes());
        plcf.extend_from_slice(&0i32.to_le_bytes());
        plcf.extend_from_slice(&0i32.to_le_bytes());
        plcf
    }

    #[test]
    fn parse_plcf_txbx_txt_skips_spare_and_reads_ranges() {
        let entries = parse_plcf_txbx_txt(&sample_plcf()).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].shape_id, 1027);
        assert_eq!((entries[0].start_cp, entries[0].end_cp), (0, 3));
        assert_eq!(entries[1].shape_id, 1031);
        assert_eq!((entries[1].start_cp, entries[1].end_cp), (3, 9));
    }

    #[test]
    fn parse_plcf_txbx_txt_rejects_bad_length() {
        assert!(parse_plcf_txbx_txt(&[]).is_err());
        assert!(parse_plcf_txbx_txt(&[0u8; 13]).is_err());
    }
}
