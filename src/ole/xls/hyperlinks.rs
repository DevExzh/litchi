//! Hyperlink parsing for XLS BIFF8 files.
//!
//! Parses HLINK records (0x01B8) which associate a cell range with a hyperlink
//! target (URL, file path, internal document reference, or UNC path).
//!
//! # Record Format (MS-XLS 2.4.63, based on Apache POI `HyperlinkRecord`)
//!
//! ```text
//! Offset  Size  Field
//! 0       8     Ref8U (rwFirst, rwLast, colFirst, colLast) - u16 each
//! 8       16    GUID (always STD_MONIKER {79EAC9D0-BAF9-11CE-8C82-00AA004BA90B})
//! 24      4     streamVersion (must be 2)
//! 28      4     linkOpts - combination of HLINK_* flags
//! 32      var   Optional label, target frame, moniker, address, textMark
//! ```

use crate::common::binary;
use crate::ole::xls::error::{XlsError, XlsResult};

/// HLINK record type identifier.
pub const RECORD_TYPE: u16 = 0x01B8;

/// Link flag bits (from POI's `HyperlinkRecord`).
const HLINK_URL: u32 = 0x0001;
/// Absolute path flag (documented for completeness; not currently checked during parsing).
const _HLINK_ABS: u32 = 0x0002;
const HLINK_LABEL: u32 = 0x0014;
const HLINK_PLACE: u32 = 0x0008;
const HLINK_TARGET_FRAME: u32 = 0x0080;
const HLINK_UNC_PATH: u32 = 0x0100;

/// Well-known moniker GUIDs (little-endian on disk).
const URL_MONIKER: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
const FILE_MONIKER: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// The type/target of a hyperlink.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HyperlinkTarget {
    /// External URL (http, https, ftp, mailto, etc.)
    Url(String),
    /// Local file path
    File {
        /// Short (8.3) filename from the record
        short_filename: String,
        /// Long filename if present
        long_filename: Option<String>,
    },
    /// Internal document reference (e.g. `Sheet1!A1`)
    Document(String),
    /// UNC network path
    Unc(String),
}

/// A parsed hyperlink from an HLINK record.
#[derive(Debug, Clone)]
pub struct XlsHyperlink {
    /// First row of the cell range (0-based).
    pub first_row: u16,
    /// Last row of the cell range (0-based).
    pub last_row: u16,
    /// First column of the cell range (0-based).
    pub first_col: u16,
    /// Last column of the cell range (0-based).
    pub last_col: u16,
    /// Display label (if present).
    pub label: Option<String>,
    /// Target frame (e.g. `_blank`), rarely used.
    pub target_frame: Option<String>,
    /// The hyperlink target.
    pub target: HyperlinkTarget,
    /// Text mark / bookmark within target (e.g. `Sheet1!A1` suffix).
    pub text_mark: Option<String>,
}

impl XlsHyperlink {
    /// Return the effective display address for this hyperlink.
    pub fn address(&self) -> &str {
        match &self.target {
            HyperlinkTarget::Url(url) => url,
            HyperlinkTarget::File {
                long_filename,
                short_filename,
            } => long_filename.as_deref().unwrap_or(short_filename),
            HyperlinkTarget::Document(doc) => doc,
            HyperlinkTarget::Unc(path) => path,
        }
    }
}

/// Read a NUL-terminated UTF-16LE string of `n_chars` characters from `data` at `offset`.
///
/// Returns the decoded string (without trailing NUL) and number of bytes consumed.
fn read_utf16le_string(data: &[u8], offset: usize, n_chars: usize) -> XlsResult<(String, usize)> {
    let byte_len = n_chars
        .checked_mul(2)
        .ok_or_else(|| XlsError::InvalidData("UTF-16 string length overflow".to_string()))?;

    if offset + byte_len > data.len() {
        return Err(XlsError::InvalidLength {
            expected: offset + byte_len,
            found: data.len(),
        });
    }

    let slice = &data[offset..offset + byte_len];
    let words: Vec<u16> = slice
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();

    let s = String::from_utf16(&words)
        .map_err(|e| XlsError::InvalidData(format!("Invalid UTF-16 in hyperlink: {}", e)))?;

    // Strip trailing NUL if present
    let cleaned = s.trim_end_matches('\0').to_string();
    Ok((cleaned, byte_len))
}

/// Read a compressed (ISO-8859-1) string of `n_bytes` bytes.
fn read_compressed_string(
    data: &[u8],
    offset: usize,
    n_bytes: usize,
) -> XlsResult<(String, usize)> {
    if offset + n_bytes > data.len() {
        return Err(XlsError::InvalidLength {
            expected: offset + n_bytes,
            found: data.len(),
        });
    }

    let slice = &data[offset..offset + n_bytes];
    // ISO-8859-1 bytes map directly to first 256 Unicode code points
    let s: String = slice.iter().map(|&b| b as char).collect();
    let cleaned = s.trim_end_matches('\0').to_string();
    Ok((cleaned, n_bytes))
}

/// Parse a single HLINK record.
pub fn parse_hlink_record(data: &[u8]) -> XlsResult<XlsHyperlink> {
    // Minimum: 8 (ref) + 16 (GUID) + 4 (version) + 4 (flags) = 32
    if data.len() < 32 {
        return Err(XlsError::InvalidLength {
            expected: 32,
            found: data.len(),
        });
    }

    let first_row = binary::read_u16_le_at(data, 0)?;
    let last_row = binary::read_u16_le_at(data, 2)?;
    let first_col = binary::read_u16_le_at(data, 4)?;
    let last_col = binary::read_u16_le_at(data, 6)?;

    // Skip GUID (16 bytes at offset 8) and streamVersion (4 bytes at offset 24)
    let stream_version = binary::read_u32_le_at(data, 24)?;
    if stream_version != 2 {
        return Err(XlsError::InvalidRecord {
            record_type: RECORD_TYPE,
            message: format!("Expected streamVersion=2, got {}", stream_version),
        });
    }

    let link_opts = binary::read_u32_le_at(data, 28)?;
    let mut offset = 32;

    // Optional label
    let label = if (link_opts & HLINK_LABEL) != 0 {
        if offset + 4 > data.len() {
            return Err(XlsError::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            });
        }
        let label_len = binary::read_u32_le_at(data, offset)? as usize;
        offset += 4;
        let (s, consumed) = read_utf16le_string(data, offset, label_len)?;
        offset += consumed;
        Some(s)
    } else {
        None
    };

    // Optional target frame
    let target_frame = if (link_opts & HLINK_TARGET_FRAME) != 0 {
        if offset + 4 > data.len() {
            return Err(XlsError::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            });
        }
        let len = binary::read_u32_le_at(data, offset)? as usize;
        offset += 4;
        let (s, consumed) = read_utf16le_string(data, offset, len)?;
        offset += consumed;
        Some(s)
    } else {
        None
    };

    // Parse the link target
    let mut target = HyperlinkTarget::Document(String::new());

    // UNC path
    if (link_opts & HLINK_URL) != 0 && (link_opts & HLINK_UNC_PATH) != 0 {
        if offset + 4 > data.len() {
            return Err(XlsError::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            });
        }
        let n_chars = binary::read_u32_le_at(data, offset)? as usize;
        offset += 4;
        let (address, consumed) = read_utf16le_string(data, offset, n_chars)?;
        offset += consumed;
        target = HyperlinkTarget::Unc(address);
    }

    // URL or file moniker
    if (link_opts & HLINK_URL) != 0 && (link_opts & HLINK_UNC_PATH) == 0 {
        if offset + 16 > data.len() {
            return Err(XlsError::InvalidLength {
                expected: offset + 16,
                found: data.len(),
            });
        }
        let moniker = &data[offset..offset + 16];
        offset += 16;

        if moniker == URL_MONIKER {
            if offset + 4 > data.len() {
                return Err(XlsError::InvalidLength {
                    expected: offset + 4,
                    found: data.len(),
                });
            }
            let length = binary::read_u32_le_at(data, offset)? as usize;
            offset += 4;

            // Per POI: length may include a 24-byte tail or may be exact
            let remaining = data.len() - offset;
            let n_chars = if (link_opts & HLINK_PLACE) != 0 {
                // There's a text mark after; use length minus potential tail
                if length > 24 {
                    // Try to detect tail presence: length = address_bytes + 24
                    let addr_bytes = length.saturating_sub(24);
                    if addr_bytes > 0 && addr_bytes <= remaining {
                        addr_bytes / 2
                    } else {
                        length / 2
                    }
                } else {
                    length / 2
                }
            } else if length == remaining {
                length / 2
            } else if length > 24 {
                (length - 24) / 2
            } else {
                length / 2
            };

            let (address, consumed) = read_utf16le_string(data, offset, n_chars)?;
            offset += consumed;

            // Skip the 24-byte tail if present
            let tail_size = length.saturating_sub(n_chars * 2);
            if tail_size > 0 && offset + tail_size <= data.len() {
                offset += tail_size;
            }

            target = HyperlinkTarget::Url(address);
        } else if moniker == FILE_MONIKER {
            if offset + 2 > data.len() {
                return Err(XlsError::InvalidLength {
                    expected: offset + 2,
                    found: data.len(),
                });
            }
            let _file_opts = binary::read_u16_le_at(data, offset)?;
            offset += 2;

            if offset + 4 > data.len() {
                return Err(XlsError::InvalidLength {
                    expected: offset + 4,
                    found: data.len(),
                });
            }
            let short_len = binary::read_u32_le_at(data, offset)? as usize;
            offset += 4;

            // Short filename is compressed (single-byte) per POI
            let (short_filename, consumed) = read_compressed_string(data, offset, short_len)?;
            offset += consumed;

            // Skip 24-byte file tail
            if offset + 24 <= data.len() {
                offset += 24;
            }

            // Optional long filename
            let long_filename = if offset + 4 <= data.len() {
                let size = binary::read_u32_le_at(data, offset)? as usize;
                offset += 4;
                if size > 0 && offset + 4 <= data.len() {
                    let char_data_size = binary::read_u32_le_at(data, offset)? as usize;
                    offset += 4;
                    // Skip usKeyValue (2 bytes)
                    if offset + 2 <= data.len() {
                        offset += 2;
                    }
                    let n_chars = char_data_size / 2;
                    if n_chars > 0 && offset + n_chars * 2 <= data.len() {
                        let (long_name, consumed) = read_utf16le_string(data, offset, n_chars)?;
                        offset += consumed;
                        Some(long_name)
                    } else {
                        None
                    }
                } else {
                    None
                }
            } else {
                None
            };

            target = HyperlinkTarget::File {
                short_filename,
                long_filename,
            };
        }
        // STD_MONIKER and unknown monikers are silently skipped
    }

    // Optional text mark (internal document reference suffix)
    let text_mark = if (link_opts & HLINK_PLACE) != 0 {
        if offset + 4 <= data.len() {
            let len = binary::read_u32_le_at(data, offset)? as usize;
            offset += 4;
            let (s, _consumed) = read_utf16le_string(data, offset, len)?;
            Some(s)
        } else {
            None
        }
    } else {
        None
    };

    // If no URL/file/UNC target was set but we have a text mark, it's a document link
    if matches!(&target, HyperlinkTarget::Document(s) if s.is_empty())
        && let Some(ref tm) = text_mark
    {
        target = HyperlinkTarget::Document(tm.clone());
    }

    Ok(XlsHyperlink {
        first_row,
        last_row,
        first_col,
        last_col,
        label,
        target_frame,
        target,
        text_mark,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_document_link() {
        // Minimal document link: Sheet1!A1
        let mut data = Vec::new();
        // Ref8U: row 3, col 0
        data.extend_from_slice(&3u16.to_le_bytes()); // rwFirst
        data.extend_from_slice(&3u16.to_le_bytes()); // rwLast
        data.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        data.extend_from_slice(&0u16.to_le_bytes()); // colLast
        // STD_MONIKER GUID
        data.extend_from_slice(&[
            0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B,
            0xA9, 0x0B,
        ]);
        // streamVersion = 2
        data.extend_from_slice(&2u32.to_le_bytes());
        // flags: HLINK_LABEL | HLINK_PLACE = 0x14 | 0x08 = 0x1C
        data.extend_from_slice(&0x0000_001Cu32.to_le_bytes());
        // Label: "place" (6 chars including NUL)
        data.extend_from_slice(&6u32.to_le_bytes());
        for c in "place\0".encode_utf16() {
            data.extend_from_slice(&c.to_le_bytes());
        }
        // Text mark: "Sheet1!A1" (10 chars including NUL)
        data.extend_from_slice(&10u32.to_le_bytes());
        for c in "Sheet1!A1\0".encode_utf16() {
            data.extend_from_slice(&c.to_le_bytes());
        }

        let link = parse_hlink_record(&data).unwrap();
        assert_eq!(link.first_row, 3);
        assert_eq!(link.last_row, 3);
        assert_eq!(link.label.as_deref(), Some("place"));
        assert_eq!(link.address(), "Sheet1!A1");
    }
}
