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
    #[inline]
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

/// XLSB record types (matching MS-XLSB specification)
/// Reference: [MS-XLSB] https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/
#[allow(dead_code)]
pub mod record_types {
    // Basic cell records
    pub const ROW_HDR: u16 = 0x0000;
    pub const CELL_BLANK: u16 = 0x0001;
    pub const CELL_RK: u16 = 0x0002;
    pub const CELL_ERROR: u16 = 0x0003;
    pub const CELL_BOOL: u16 = 0x0004;
    pub const CELL_REAL: u16 = 0x0005;
    pub const CELL_ST: u16 = 0x0006;
    pub const CELL_ISST: u16 = 0x0007;

    // Formula records
    pub const FMLA_STRING: u16 = 0x0008;
    pub const FMLA_NUM: u16 = 0x0009;
    pub const FMLA_BOOL: u16 = 0x000A;
    pub const FMLA_ERROR: u16 = 0x000B;

    // Shared string table
    pub const SST_ITEM: u16 = 0x0013;

    // Format and style records
    pub const FONT: u16 = 0x002B;
    pub const FMT: u16 = 0x002C;
    pub const FILL: u16 = 0x002D;
    pub const BORDER: u16 = 0x002E;
    pub const XF: u16 = 0x002F;
    pub const STYLE: u16 = 0x0030;
    pub const CELL_META: u16 = 0x0031;
    pub const VALUE_META: u16 = 0x0032;

    // Column and dimension records
    pub const COL_INFO: u16 = 0x003C;
    pub const CELL_R_STRING: u16 = 0x003E;

    // Workbook structure records
    pub const FILE_VERSION: u16 = 0x0080;
    pub const BEGIN_SHEET: u16 = 0x0081;
    pub const END_SHEET: u16 = 0x0082;
    pub const BEGIN_BOOK: u16 = 0x0083;
    pub const END_BOOK: u16 = 0x0084;
    pub const BEGIN_WS_VIEWS: u16 = 0x0085;
    pub const END_WS_VIEWS: u16 = 0x0086;
    pub const BEGIN_BOOK_VIEWS: u16 = 0x0087;
    pub const END_BOOK_VIEWS: u16 = 0x0088;
    pub const BEGIN_WS_VIEW: u16 = 0x0089;
    pub const END_WS_VIEW: u16 = 0x008A;
    pub const BEGIN_CS_VIEWS: u16 = 0x008B;
    pub const END_CS_VIEWS: u16 = 0x008C;
    pub const BEGIN_CS_VIEW: u16 = 0x008D;
    pub const END_CS_VIEW: u16 = 0x008E;
    pub const BEGIN_BUNDLE_SHS: u16 = 0x008F;
    pub const END_BUNDLE_SHS: u16 = 0x0090;
    pub const BEGIN_SHEET_DATA: u16 = 0x0091;
    pub const END_SHEET_DATA: u16 = 0x0092;
    pub const WS_PROP: u16 = 0x0093;
    pub const WS_DIM: u16 = 0x0094;
    pub const WORKBOOK_PROP: u16 = 0x0099;
    pub const BUNDLE_SH: u16 = 0x009C;
    pub const CALC_PROP: u16 = 0x009D;
    pub const BOOK_VIEW: u16 = 0x009E;
    pub const BEGIN_SST: u16 = 0x009F;
    pub const END_SST: u16 = 0x00A0;

    // Filter records
    pub const BEGIN_A_FILTER: u16 = 0x00A1;
    pub const END_A_FILTER: u16 = 0x00A2;
    pub const BEGIN_FILTER_COLUMN: u16 = 0x00A3;
    pub const END_FILTER_COLUMN: u16 = 0x00A4;
    pub const BEGIN_FILTERS: u16 = 0x00A5;
    pub const END_FILTERS: u16 = 0x00A6;
    pub const FILTER: u16 = 0x00A7;
    pub const COLOR_FILTER: u16 = 0x00A8;
    pub const ICON_FILTER: u16 = 0x00A9;
    pub const TOP10_FILTER: u16 = 0x00AA;
    pub const DYNAMIC_FILTER: u16 = 0x00AB;
    pub const BEGIN_CUSTOM_FILTERS: u16 = 0x00AC;
    pub const END_CUSTOM_FILTERS: u16 = 0x00AD;
    pub const CUSTOM_FILTER: u16 = 0x00AE;
    pub const A_FILTER_DATE_GROUP_ITEM: u16 = 0x00AF;

    // Merge cells
    pub const MERGE_CELL: u16 = 0x00B0;
    pub const BEGIN_MERGE_CELLS: u16 = 0x00B1;
    pub const END_MERGE_CELLS: u16 = 0x00B2;

    // Named ranges
    pub const NAME: u16 = 0x0027;

    // Formulas and tables
    pub const ARR_FMLA: u16 = 0x01AA;
    pub const SHR_FMLA: u16 = 0x01AB;
    pub const TABLE: u16 = 0x01AC;

    // External connections and links
    pub const BEGIN_EXTERNALS: u16 = 0x0161;
    pub const END_EXTERNALS: u16 = 0x0162;
    pub const SUP_BOOK_SRC: u16 = 0x0163;
    pub const SUP_SELF: u16 = 0x0165;
    pub const SUP_SAME: u16 = 0x0166;
    pub const SUP_TABS: u16 = 0x0167;
    pub const BEGIN_SUP_BOOK: u16 = 0x0168;
    pub const PLACEHOLDER_NAME: u16 = 0x0169;
    pub const EXTERN_SHEET: u16 = 0x016A;
    pub const EXTERN_TABLE_START: u16 = 0x016B;
    pub const EXTERN_TABLE_END: u16 = 0x016C;
    pub const EXTERN_ROW_HDR: u16 = 0x016E;
    pub const EXTERN_CELL_BLANK: u16 = 0x016F;
    pub const EXTERN_CELL_REAL: u16 = 0x0170;
    pub const EXTERN_CELL_BOOL: u16 = 0x0171;
    pub const EXTERN_CELL_ERROR: u16 = 0x0172;
    pub const EXTERN_CELL_STRING: u16 = 0x0173;
    pub const END_SUP_BOOK: u16 = 0x024C;

    // Style sheet records
    pub const BEGIN_STYLE_SHEET: u16 = 0x0116;
    pub const END_STYLE_SHEET: u16 = 0x0117;
    pub const BEGIN_FMTS: u16 = 0x0267;
    pub const END_FMTS: u16 = 0x0268;
    pub const BEGIN_FONTS: u16 = 0x0263;
    pub const END_FONTS: u16 = 0x0264;
    pub const BEGIN_FILLS: u16 = 0x025B;
    pub const END_FILLS: u16 = 0x025C;
    pub const BEGIN_BORDERS: u16 = 0x0265;
    pub const END_BORDERS: u16 = 0x0266;
    pub const BEGIN_CELL_XFS: u16 = 0x0269;
    pub const END_CELL_XFS: u16 = 0x026A;
    pub const BEGIN_STYLES: u16 = 0x026B;
    pub const END_STYLES: u16 = 0x026C;
    pub const BEGIN_CELL_STYLE_XFS: u16 = 0x0272;
    pub const END_CELL_STYLE_XFS: u16 = 0x0273;
    pub const BEGIN_DXFS: u16 = 0x01F9;
    pub const END_DXFS: u16 = 0x01FA;
    pub const DXF: u16 = 0x01FB;
    pub const BEGIN_TABLE_STYLES: u16 = 0x01FC;
    pub const END_TABLE_STYLES: u16 = 0x01FD;

    // Comments
    pub const BEGIN_COMMENTS: u16 = 0x0274;
    pub const END_COMMENTS: u16 = 0x0275;
    pub const BEGIN_COMMENT_AUTHORS: u16 = 0x0276;
    pub const END_COMMENT_AUTHORS: u16 = 0x0277;
    pub const COMMENT_AUTHOR: u16 = 0x0278;
    pub const BEGIN_COMMENT_LIST: u16 = 0x0279;
    pub const END_COMMENT_LIST: u16 = 0x027A;
    pub const BEGIN_COMMENT: u16 = 0x027B;
    pub const END_COMMENT: u16 = 0x027C;
    pub const COMMENT_TEXT: u16 = 0x027D;

    // Hyperlinks
    pub const H_LINK: u16 = 0x01EE;

    // Page setup
    pub const MARGINS: u16 = 0x01DC;
    pub const PRINT_OPTIONS: u16 = 0x01DD;
    pub const PAGE_SETUP: u16 = 0x01DE;
    pub const BEGIN_HEADER_FOOTER: u16 = 0x01DF;
    pub const END_HEADER_FOOTER: u16 = 0x01E0;

    // Column information
    pub const BEGIN_COL_INFOS: u16 = 0x0186;
    pub const END_COL_INFOS: u16 = 0x0187;

    // Drawing
    pub const DRAWING: u16 = 0x0226;
    pub const LEGACY_DRAWING: u16 = 0x0227;
    pub const LEGACY_DRAWING_HF: u16 = 0x0228;

    // Data validation
    pub const BEGIN_D_VALS: u16 = 0x023D;
    pub const END_D_VALS: u16 = 0x023E;
    pub const D_VAL: u16 = 0x0040;

    // Conditional formatting
    pub const BEGIN_CF_RULE: u16 = 0x01CF;
    pub const END_CF_RULE: u16 = 0x01D0;
    pub const BEGIN_ICON_SET: u16 = 0x01D1;
    pub const END_ICON_SET: u16 = 0x01D2;
    pub const BEGIN_DATABAR: u16 = 0x01D3;
    pub const END_DATABAR: u16 = 0x01D4;
    pub const BEGIN_COLOR_SCALE: u16 = 0x01D5;
    pub const END_COLOR_SCALE: u16 = 0x01D6;
    pub const CFVO: u16 = 0x01D7;

    // Protection
    pub const BOOK_PROTECTION: u16 = 0x0216;
    pub const SHEET_PROTECTION: u16 = 0x0217;
    pub const RANGE_PROTECTION: u16 = 0x0218;

    // Miscellaneous
    pub const WS_FMT_INFO: u16 = 0x01E5;
    pub const BIG_NAME: u16 = 0x0271;
    pub const FILE_SHARING: u16 = 0x0224;
    pub const OLE_SIZE: u16 = 0x0225;
    pub const WEB_OPT: u16 = 0x0229;
    pub const PHONETIC_INFO: u16 = 0x0219;

    // Excel 2013+ records
    pub const ABS_PATH15: u16 = 0x0817;
    pub const BEGIN_SPARKLINE_GROUPS: u16 = 0x0422;
    pub const END_SPARKLINE_GROUPS: u16 = 0x0423;
    pub const BEGIN_SPARKLINE_GROUP: u16 = 0x0411;
    pub const END_SPARKLINE_GROUP: u16 = 0x0412;
    pub const SPARKLINE: u16 = 0x0413;
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

    // Convert UTF-16LE to UTF-8 using encoding_rs
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

    // Convert UTF-16LE to UTF-8 using encoding_rs
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
    Formula {
        value: Box<CellValue>,
        formula: Option<Vec<u8>>, // Raw formula bytes
    },
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
        if data.len() < 4 {
            return Err(XlsbError::InvalidLength {
                expected: 4,
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
            // Formula records - parse formula bytes and cached value
            record_types::FMLA_STRING => {
                if data.len() < 10 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                // Skip style_id (4 bytes) and flags (1 byte) and formula length (4 bytes)
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();

                // Read cached string value after formula
                let mut str_len = 0;
                let string = wide_str(&data[10 + formula_len..], &mut str_len)?;
                CellValue::Formula {
                    value: Box::new(CellValue::String(string)),
                    formula: Some(formula_bytes),
                }
            },
            record_types::FMLA_NUM => {
                if data.len() < 18 {
                    return Err(XlsbError::InvalidLength {
                        expected: 18,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 8 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len + 8,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let num_value = binary::read_f64_le_at(data, 10 + formula_len)?;
                CellValue::Formula {
                    value: Box::new(CellValue::Real(num_value)),
                    formula: Some(formula_bytes),
                }
            },
            record_types::FMLA_BOOL => {
                if data.len() < 11 {
                    return Err(XlsbError::InvalidLength {
                        expected: 11,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 1 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len + 1,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let bool_value = data[10 + formula_len] != 0;
                CellValue::Formula {
                    value: Box::new(CellValue::Bool(bool_value)),
                    formula: Some(formula_bytes),
                }
            },
            record_types::FMLA_ERROR => {
                if data.len() < 11 {
                    return Err(XlsbError::InvalidLength {
                        expected: 11,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 1 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len + 1,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let error_code = data[10 + formula_len];
                CellValue::Formula {
                    value: Box::new(CellValue::Error(error_code)),
                    formula: Some(formula_bytes),
                }
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

/// Column information record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct ColInfoRecord {
    pub first_col: u32,
    pub last_col: u32,
    pub width: f64,
    pub style_xf: u32,
    pub custom_width: bool,
    pub hidden: bool,
    pub best_fit: bool,
}

impl ColInfoRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 12 {
            return Err(XlsbError::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let first_col = binary::read_u32_le_at(data, 0)?;
        let last_col = binary::read_u32_le_at(data, 4)?;
        // Width is stored as 256ths of a character
        let width_raw = binary::read_u32_le_at(data, 8)?;
        let width = width_raw as f64 / 256.0;

        let style_xf = if data.len() >= 16 {
            binary::read_u32_le_at(data, 12)?
        } else {
            0
        };

        let flags = if data.len() >= 18 {
            binary::read_u16_le_at(data, 16)?
        } else {
            0
        };

        let custom_width = (flags & 0x0002) != 0;
        let hidden = (flags & 0x0001) != 0;
        let best_fit = (flags & 0x0008) != 0;

        Ok(ColInfoRecord {
            first_col,
            last_col,
            width,
            style_xf,
            custom_width,
            hidden,
            best_fit,
        })
    }
}

/// Merged cell record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct MergeCellRecord {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl MergeCellRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        Ok(MergeCellRecord {
            row_first: binary::read_u32_le_at(data, 0)?,
            row_last: binary::read_u32_le_at(data, 4)?,
            col_first: binary::read_u32_le_at(data, 8)?,
            col_last: binary::read_u32_le_at(data, 12)?,
        })
    }
}

/// Hyperlink record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct HyperlinkRecord {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
    pub r_id: String,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

impl HyperlinkRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        let row_first = binary::read_u32_le_at(data, 0)?;
        let row_last = binary::read_u32_le_at(data, 4)?;
        let col_first = binary::read_u32_le_at(data, 8)?;
        let col_last = binary::read_u32_le_at(data, 12)?;

        let mut offset = 16;

        // Read relationship ID
        let (r_id, consumed) = wide_str_with_len(&data[offset..])?;
        offset += consumed;

        // Read location (optional)
        let (location, consumed) = if offset < data.len() {
            let (loc, c) = wide_str_with_len(&data[offset..])?;
            (if loc.is_empty() { None } else { Some(loc) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read tooltip (optional)
        let (tooltip, consumed) = if offset < data.len() {
            let (tt, c) = wide_str_with_len(&data[offset..])?;
            (if tt.is_empty() { None } else { Some(tt) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read display text (optional)
        let display = if offset < data.len() {
            let (disp, _) = wide_str_with_len(&data[offset..])?;
            if disp.is_empty() { None } else { Some(disp) }
        } else {
            None
        };

        Ok(HyperlinkRecord {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location,
            tooltip,
            display,
        })
    }
}

/// Named range record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct NameRecord {
    pub name: String,
    pub formula: Option<Vec<u8>>,
    pub sheet_id: Option<u32>,
    pub hidden: bool,
    pub function: bool,
}

impl NameRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let flags = binary::read_u32_le_at(data, 0)?;
        let hidden = (flags & 0x0001) != 0;
        let function = (flags & 0x0002) != 0;

        // Sheet ID (-1 for global scope, otherwise sheet-specific)
        let sheet_id_raw = binary::read_u32_le_at(data, 4)? as i32;
        let sheet_id = if sheet_id_raw == -1 {
            None
        } else {
            Some(sheet_id_raw as u32)
        };

        let mut offset = 8;

        // Read name
        let (name, consumed) = wide_str_with_len(&data[offset..])?;
        offset += consumed;

        // Read formula if present
        let formula = if offset < data.len() {
            let formula_len = binary::read_u32_le_at(data, offset)? as usize;
            offset += 4;
            if data.len() >= offset + formula_len {
                Some(data[offset..offset + formula_len].to_vec())
            } else {
                None
            }
        } else {
            None
        };

        Ok(NameRecord {
            name,
            formula,
            sheet_id,
            hidden,
            function,
        })
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
