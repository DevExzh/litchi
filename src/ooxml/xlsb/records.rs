//! XLSB record parsing for Excel 2007+ binary format
//!
//! XLSB (Excel Binary Workbook) uses a different record structure than
//! the older XLS BIFF format. Records are stored in a ZIP container
//! and use a binary record format with variable-length encoding.

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};
use bytes::Bytes;
use std::io::Read;

/// XLSB record header (variable length encoding)
#[derive(Debug, Clone)]
pub struct XlsbRecordHeader {
    pub record_type: u16,
    pub data_len: usize,
}

impl XlsbRecordHeader {
    /// Read record header with variable-length encoding
    pub fn read<R: Read>(reader: &mut R) -> XlsbResult<Self> {
        let mut b = [0u8; 1];
        reader.read_exact(&mut b)?;
        let mut record_type = (b[0] & 0x7F) as u16;

        if (b[0] & 0x80) != 0 {
            reader.read_exact(&mut b)?;
            record_type |= ((b[0] & 0x7F) as u16) << 7;

            if (b[0] & 0x80) != 0 {
                reader.read_exact(&mut b)?;
                record_type |= ((b[0] & 0x7F) as u16) << 14;
            }
        }

        // Read variable-length data size
        let mut data_len = 0usize;
        let mut shift = 0;

        loop {
            reader.read_exact(&mut b)?;
            data_len |= ((b[0] & 0x7F) as usize) << shift;
            shift += 7;

            if (b[0] & 0x80) == 0 {
                break;
            }

            if shift > 28 {
                return Err(XlsbError::InvalidLength {
                    expected: 0,
                    found: data_len,
                });
            }
        }

        Ok(XlsbRecordHeader {
            record_type,
            data_len,
        })
    }
}

/// XLSB record with header and data
#[derive(Debug, Clone)]
pub struct XlsbRecord {
    pub header: XlsbRecordHeader,
    pub data: Bytes,
}

impl XlsbRecord {
    /// Read a complete XLSB record
    pub fn read<R: Read>(reader: &mut R) -> XlsbResult<Self> {
        let header = XlsbRecordHeader::read(reader)?;

        let mut data_buf = vec![0u8; header.data_len];
        reader.read_exact(&mut data_buf)?;
        let data = Bytes::from(data_buf);

        Ok(XlsbRecord { header, data })
    }
}

/// Iterator over XLSB records in a stream
pub struct XlsbRecordIter<R> {
    reader: R,
}

impl<R: Read> XlsbRecordIter<R> {
    pub fn new(reader: R) -> Self {
        XlsbRecordIter { reader }
    }
}

impl<R: Read> Iterator for XlsbRecordIter<R> {
    type Item = XlsbResult<XlsbRecord>;

    fn next(&mut self) -> Option<Self::Item> {
        match XlsbRecord::read(&mut self.reader) {
            Ok(record) => Some(Ok(record)),
            Err(XlsbError::Io(e)) if e.kind() == std::io::ErrorKind::UnexpectedEof => None,
            Err(e) => Some(Err(e)),
        }
    }
}

/// XLSB record types (matching pyxlsb2 constants)
#[allow(dead_code)]
pub mod record_types {
    pub const WORKBOOK_PROP: u16 = 153; // 0x99
    pub const BUNDLE_SH: u16 = 156; // 0x9C
    pub const BEGIN_SHEET: u16 = 129; // 0x81
    pub const END_SHEET: u16 = 130; // 0x82
    pub const BEGIN_SST: u16 = 159; // 0x9F
    pub const END_SST: u16 = 160; // 0xA0
    pub const SST_ITEM: u16 = 19; // 0x13
    pub const ROW_HDR: u16 = 0; // 0x00
    pub const CELL_BLANK: u16 = 1; // 0x01
    pub const CELL_RK: u16 = 2; // 0x02
    pub const CELL_ERROR: u16 = 3; // 0x03
    pub const CELL_BOOL: u16 = 4; // 0x04
    pub const CELL_REAL: u16 = 5; // 0x05
    pub const CELL_ST: u16 = 6; // 0x06
    pub const CELL_ISST: u16 = 7; // 0x07
    pub const FMLA_STRING: u16 = 8; // 0x08
    pub const FMLA_NUM: u16 = 9; // 0x09
    pub const FMLA_BOOL: u16 = 10; // 0x0A
    pub const FMLA_ERROR: u16 = 11; // 0x0B
    pub const BEGIN_BUNDLE_SHS: u16 = 141; // 0x8D
    pub const END_BUNDLE_SHS: u16 = 144; // 0x90
    pub const BEGIN_SHEET_DATA: u16 = 145; // 0x91
    pub const END_SHEET_DATA: u16 = 146; // 0x92
}

/// Decode wide string (UTF-16LE) from XLSB format
pub fn wide_str(buf: &[u8], str_len: &mut usize) -> XlsbResult<String> {
    if buf.len() < 4 {
        return Err(XlsbError::InvalidLength {
            expected: 4,
            found: buf.len(),
        });
    }

    let len = binary::read_u32_le_at(buf, 0)? as usize;
    if buf.len() < 4 + len * 2 {
        return Err(XlsbError::WideStringLength {
            expected: 4 + len * 2,
            actual: buf.len(),
        });
    }

    *str_len = 4 + len * 2;
    let utf16_data = &buf[4..*str_len];

    // Convert UTF-16LE to UTF-8 using encoding_rs like calamine
    use encoding_rs::UTF_16LE;
    Ok(UTF_16LE.decode(utf16_data).0.into_owned())
}

/// Decode wide string (UTF-16LE) from XLSB format and return consumed bytes
pub fn wide_str_with_len(buf: &[u8]) -> XlsbResult<(String, usize)> {
    if buf.len() < 4 {
        return Err(XlsbError::InvalidLength {
            expected: 4,
            found: buf.len(),
        });
    }

    let len = binary::read_u32_le_at(buf, 0)? as usize;
    if buf.len() < 4 + len * 2 {
        return Err(XlsbError::WideStringLength {
            expected: 4 + len * 2,
            actual: buf.len(),
        });
    }

    let consumed = 4 + len * 2;
    let utf16_data = &buf[4..consumed];

    // Convert UTF-16LE to UTF-8 using encoding_rs like calamine
    use encoding_rs::UTF_16LE;
    Ok((UTF_16LE.decode(utf16_data).0.into_owned(), consumed))
}

/// Workbook properties record
#[derive(Debug, Clone)]
pub struct WorkbookPropRecord {
    pub is_date1904: bool,
}

impl WorkbookPropRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.is_empty() {
            return Ok(WorkbookPropRecord { is_date1904: false });
        }

        let flags = data[0];
        let is_date1904 = (flags & 0x01) != 0;

        Ok(WorkbookPropRecord { is_date1904 })
    }
}

/// Bundle sheet record (worksheet metadata)
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct BundleSheetRecord {
    pub id: u32,
    pub name: String,
    pub visible: u8,
    pub sheet_type: u8,
}

impl BundleSheetRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 12 {
            return Err(XlsbError::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let id = binary::read_u32_le_at(data, 0)?;
        let visible = data[4];
        let sheet_type = data[5];
        let rel_len = binary::read_u32_le_at(data, 8)? as usize;

        if rel_len != 0xFFFF_FFFF {
            let rel_len = rel_len * 2; // UTF-16 bytes
            if data.len() < 12 + rel_len {
                return Err(XlsbError::InvalidLength {
                    expected: 12 + rel_len,
                    found: data.len(),
                });
            }

            // Skip the relationship ID (rel_len bytes of UTF-16)
            let name_start = 12 + rel_len;
            if data.len() < name_start + 4 {
                return Err(XlsbError::InvalidLength {
                    expected: name_start + 4,
                    found: data.len(),
                });
            }

            // Read sheet name as wide string
            let (name, _) = wide_str_with_len(&data[name_start..])?;

            Ok(BundleSheetRecord {
                id,
                name: name.to_string(),
                visible,
                sheet_type,
            })
        } else {
            // Handle the case where rel_len is 0xFFFFFFFF
            // This might indicate a different format or no relationship
            Ok(BundleSheetRecord {
                id,
                name: format!("Sheet{}", id),
                visible,
                sheet_type,
            })
        }
    }
}

/// SST item record (shared string)
#[derive(Debug, Clone)]
pub struct SstItemRecord {
    pub string: String,
}

impl SstItemRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        let mut str_len = 0;
        let string = wide_str(&data[1..], &mut str_len)?;

        Ok(SstItemRecord { string })
    }
}

/// Row header record
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct RowHeaderRecord {
    pub row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

#[allow(dead_code)]
impl RowHeaderRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let row = binary::read_u32_le_at(data, 0)?;
        let first_col = binary::read_u16_le_at(data, 4)?;
        let last_col = binary::read_u16_le_at(data, 6)?;

        Ok(RowHeaderRecord {
            row,
            first_col,
            last_col,
        })
    }
}

/// Cell value types
#[derive(Debug, Clone)]
pub enum CellValue {
    Blank,
    Bool(bool),
    Error(u8),
    Real(f64),
    String(String),
    Isst(u32), // Index into shared string table
}

/// Cell record base
#[derive(Debug, Clone)]
pub struct CellRecord {
    pub row: u32,
    pub col: u16,
    pub value: CellValue,
}

impl CellRecord {
    pub fn parse(record_type: u16, data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 6 {
            return Err(XlsbError::InvalidLength {
                expected: 6,
                found: data.len(),
            });
        }

        let row = binary::read_u32_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 4)?;

        let value = match record_type {
            record_types::CELL_BLANK => CellValue::Blank,
            record_types::CELL_BOOL => {
                if data.len() < 7 {
                    return Err(XlsbError::InvalidLength {
                        expected: 7,
                        found: data.len(),
                    });
                }
                CellValue::Bool(data[6] != 0)
            },
            record_types::CELL_ERROR => {
                if data.len() < 7 {
                    return Err(XlsbError::InvalidLength {
                        expected: 7,
                        found: data.len(),
                    });
                }
                CellValue::Error(data[6])
            },
            record_types::CELL_REAL => {
                if data.len() < 14 {
                    return Err(XlsbError::InvalidLength {
                        expected: 14,
                        found: data.len(),
                    });
                }
                CellValue::Real(binary::read_f64_le_at(data, 6)?)
            },
            record_types::CELL_ST => {
                let mut str_len = 0;
                let string = wide_str(&data[6..], &mut str_len)?;
                CellValue::String(string.to_owned())
            },
            record_types::CELL_ISST => {
                if data.len() < 10 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                CellValue::Isst(binary::read_u32_le_at(data, 6)?)
            },
            record_types::CELL_RK => {
                if data.len() < 10 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                let rk_value = binary::read_u32_le_at(data, 6)?;
                let real_value = rk_to_f64(rk_value);
                CellValue::Real(real_value)
            },
            _ => return Err(XlsbError::InvalidRecordType(record_type)),
        };

        Ok(CellRecord { row, col, value })
    }
}

/// Convert RK value to f64 (same as XLS)
pub fn rk_to_f64(rk: u32) -> f64 {
    let d100 = (rk & 0x02) != 0;
    let is_int = (rk & 0x01) != 0;

    let value = if is_int {
        let int_val = (rk >> 2) as i32;
        if d100 {
            if int_val % 100 != 0 {
                int_val as f64 / 100.0
            } else {
                (int_val / 100) as f64
            }
        } else {
            int_val as f64
        }
    } else {
        // Float value - reconstruct from 30 bits
        let mut float_bits = [0u8; 8];
        float_bits[0..4].copy_from_slice(&(rk & 0xFFFFFFFC).to_le_bytes());
        f64::from_le_bytes(float_bits)
    };

    if d100 && !is_int {
        value / 100.0
    } else {
        value
    }
}

/// Record iterator for XLSB parsing
pub struct RecordIter<R> {
    reader: R,
    buffer: [u8; 1],
}

impl<R: Read> RecordIter<R> {
    pub fn new(reader: R) -> Self {
        RecordIter {
            reader,
            buffer: [0],
        }
    }

    pub fn from_cursor(cursor: std::io::Cursor<&[u8]>) -> RecordIter<std::io::Cursor<&[u8]>> {
        RecordIter::new(cursor)
    }

    fn read_u8(&mut self) -> Result<u8, std::io::Error> {
        self.reader.read_exact(&mut self.buffer)?;
        Ok(self.buffer[0])
    }

    /// Read next type, until we have no future record
    pub fn read_type(&mut self) -> Result<u16, std::io::Error> {
        let b = self.read_u8()?;
        let typ = if (b & 0x80) == 0x80 {
            (b & 0x7F) as u16 + (((self.read_u8()? & 0x7F) as u16) << 7)
        } else {
            b as u16
        };
        Ok(typ)
    }

    pub fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> Result<usize, std::io::Error> {
        let mut b = self.read_u8()?;
        let mut len = (b & 0x7F) as usize;
        for i in 1..4 {
            if (b & 0x80) == 0 {
                break;
            }
            b = self.read_u8()?;
            len += ((b & 0x7F) as usize) << (7 * i);
        }
        if buf.len() < len {
            *buf = vec![0; len];
        }

        self.reader.read_exact(&mut buf[..len])?;
        Ok(len)
    }

    /// Reads next type, and discard blocks between `start` and `end`
    pub fn next_skip_blocks(
        &mut self,
        record_type: u16,
        bounds: &[(u16, Option<u16>)],
        buf: &mut Vec<u8>,
    ) -> Result<usize, XlsbError> {
        loop {
            let typ = self.read_type().map_err(XlsbError::Io)?;
            let len = self.fill_buffer(buf).map_err(XlsbError::Io)?;
            if typ == record_type {
                return Ok(len);
            }
            if let Some(end) = bounds.iter().find(|b| b.0 == typ).and_then(|b| b.1) {
                while self.read_type().map_err(XlsbError::Io)? != end {
                    let _ = self.fill_buffer(buf).map_err(XlsbError::Io)?;
                }
                let _ = self.fill_buffer(buf).map_err(XlsbError::Io)?;
            }
        }
    }
}
