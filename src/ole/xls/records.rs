//! BIFF record parsing for XLS files
//!
//! This module handles the parsing of BIFF (Binary Interchange File Format)
//! records used in Excel XLS files. BIFF records contain various types of
//! data including cell values, formatting, formulas, and metadata.

use std::io::{Read, Seek, SeekFrom};

use crate::ole::binary;
use crate::ole::xls::error::{XlsError, XlsResult};
use crate::ole::xls::utils;

/// BIFF record header (4 bytes: type + length)
#[derive(Debug, Clone)]
pub struct RecordHeader {
    pub record_type: u16,
    pub data_len: u16,
}

impl RecordHeader {
    /// Parse record header from stream
    pub fn read<R: Read>(reader: &mut R) -> XlsResult<Self> {
        let mut buf = [0u8; 4];
        reader.read_exact(&mut buf)?;
        let record_type = u16::from_le_bytes([buf[0], buf[1]]);
        let data_len = u16::from_le_bytes([buf[2], buf[3]]);

        Ok(RecordHeader {
            record_type,
            data_len,
        })
    }
}

/// Iterator over BIFF records in a stream
pub struct RecordIter<R> {
    reader: R,
    stream_len: u64,
    current_pos: u64,
}

impl<R: Read + Seek> RecordIter<R> {
    pub fn new(mut reader: R) -> XlsResult<Self> {
        let stream_len = reader.seek(SeekFrom::End(0))?;
        reader.seek(SeekFrom::Start(0))?;

        Ok(RecordIter {
            reader,
            stream_len,
            current_pos: 0,
        })
    }

    /// Seek to a specific position in the stream
    pub fn seek(&mut self, pos: u64) -> XlsResult<()> {
        self.reader.seek(SeekFrom::Start(pos))?;
        self.current_pos = pos;
        Ok(())
    }
}

impl<R: Read + Seek> Iterator for RecordIter<R> {
    type Item = XlsResult<Record>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_pos >= self.stream_len {
            return None;
        }

        match Record::read(&mut self.reader) {
            Ok(record) => {
                self.current_pos += 4 + record.header.data_len as u64;
                Some(Ok(record))
            }
            Err(e) => Some(Err(e)),
        }
    }
}

/// A BIFF record with header and data
#[derive(Debug, Clone)]
pub struct Record {
    pub header: RecordHeader,
    pub data: Vec<u8>,
}

impl Record {
    /// Read a complete record from the stream
    pub fn read<R: Read>(reader: &mut R) -> XlsResult<Self> {
        let header = RecordHeader::read(reader)?;

        let mut data = vec![0u8; header.data_len as usize];
        reader.read_exact(&mut data)?;

        Ok(Record { header, data })
    }
}

/// BIFF versions supported
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BiffVersion {
    Biff2 = 0x0200,
    Biff3 = 0x0300,
    Biff4 = 0x0400,
    Biff5 = 0x0500,
    Biff8 = 0x0600,
}

impl BiffVersion {
    pub fn from_bof_version(version: u16) -> Option<Self> {
        match version {
            0x0200 | 0x0002 | 0x0007 => Some(BiffVersion::Biff2),
            0x0300 => Some(BiffVersion::Biff3),
            0x0400 => Some(BiffVersion::Biff4),
            0x0500 => Some(BiffVersion::Biff5),
            0x0600 => Some(BiffVersion::Biff8),
            _ => None,
        }
    }

    pub fn supports_unicode(&self) -> bool {
        matches!(self, BiffVersion::Biff8)
    }
}

/// BOF (Beginning of File) record
#[derive(Debug, Clone)]
pub struct BofRecord {
    pub version: BiffVersion,
    pub is_1904_date_system: bool,
}

impl BofRecord {
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 4 {
            return Err(XlsError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let biff_version = binary::read_u16_le_at(data, 0)?;
        let dt = if data.len() >= 6 {
            binary::read_u16_le_at(data, 4)?
        } else {
            0
        };

        let version = BiffVersion::from_bof_version(biff_version)
            .ok_or_else(|| XlsError::UnsupportedBiffVersion(biff_version))?;

        let is_1904_date_system = dt == 1;

        Ok(BofRecord {
            version,
            is_1904_date_system,
        })
    }
}

/// Dimensions record (worksheet bounds)
#[derive(Debug, Clone)]
pub struct DimensionsRecord {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u32,
    pub last_col: u32,
}

impl DimensionsRecord {
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        match data.len() {
            10 => {
                // BIFF5-BIFF8
                Ok(DimensionsRecord {
                    first_row: binary::read_u16_le_at(data, 0)? as u32,
                    last_row: binary::read_u16_le_at(data, 2)? as u32,
                    first_col: binary::read_u16_le_at(data, 4)? as u32,
                    last_col: binary::read_u16_le_at(data, 6)? as u32,
                })
            }
            14 => {
                // BIFF8 with 32-bit row indices
                Ok(DimensionsRecord {
                    first_row: binary::read_u32_le_at(data, 0)?,
                    last_row: binary::read_u32_le_at(data, 4)?,
                    first_col: binary::read_u16_le_at(data, 8)? as u32,
                    last_col: binary::read_u16_le_at(data, 10)? as u32,
                })
            }
            _ => Err(XlsError::InvalidLength {
                expected: 10,
                found: data.len(),
            }),
        }
    }
}

/// Sheet visibility types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetVisible {
    Visible = 0x00,
    Hidden = 0x01,
    VeryHidden = 0x02,
}

impl SheetVisible {
    pub fn from_u8(value: u8) -> XlsResult<Self> {
        match value & 0x3 {
            0x00 => Ok(SheetVisible::Visible),
            0x01 => Ok(SheetVisible::Hidden),
            0x02 => Ok(SheetVisible::VeryHidden),
            v => Err(XlsError::InvalidRecord {
                record_type: 0x0085, // BoundSheet8
                message: format!("Invalid visibility value: {}", v),
            }),
        }
    }
}

/// Sheet types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetType {
    WorkSheet,
    MacroSheet,
    ChartSheet,
    VBModule,
}

impl SheetType {
    pub fn from_u8(value: u8) -> XlsResult<Self> {
        match value {
            0x00 => Ok(SheetType::WorkSheet),
            0x01 => Ok(SheetType::MacroSheet),
            0x02 => Ok(SheetType::ChartSheet),
            0x06 => Ok(SheetType::VBModule),
            v => Err(XlsError::InvalidRecord {
                record_type: 0x0085, // BoundSheet8
                message: format!("Invalid sheet type: {}", v),
            }),
        }
    }
}

/// BoundSheet8 record (worksheet metadata)
#[derive(Debug, Clone)]
pub struct BoundSheetRecord {
    pub position: u32,
    pub visible: SheetVisible,
    pub sheet_type: SheetType,
    pub name: String,
}

impl BoundSheetRecord {
    pub fn parse(data: &[u8], encoding: &XlsEncoding) -> XlsResult<Self> {
        if data.len() < 8 {
            return Err(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let position = binary::read_u32_le_at(data, 0)?;
        let visible = SheetVisible::from_u8(data[4])?;
        let sheet_type = SheetType::from_u8(data[5])?;

        // Skip 2 bytes and parse the name
        let name_data = &data[6..];
        let name = utils::parse_short_string(name_data, encoding)?;

        Ok(BoundSheetRecord {
            position,
            visible,
            sheet_type,
            name,
        })
    }
}

/// Codepage/encoding information
#[derive(Debug, Clone)]
pub enum XlsEncoding {
    /// Single-byte encoding with codepage
    Codepage(u16),
    /// UTF-16 little endian (BIFF8+)
    Utf16Le,
}

impl XlsEncoding {
    pub fn from_codepage(codepage: u16) -> XlsResult<Self> {
        // Use codepage crate for proper encoding support
        // For now, we'll handle common codepages
        match codepage {
            1200 => Ok(XlsEncoding::Utf16Le),
            cp => Ok(XlsEncoding::Codepage(cp)),
        }
    }

    pub fn decode(&self, data: &[u8]) -> XlsResult<String> {
        match self {
            XlsEncoding::Utf16Le => {
                // UTF-16 LE decoding
                if data.len() % 2 != 0 {
                    return Err(XlsError::Encoding("Invalid UTF-16 data length".to_string()));
                }
                let utf16_data: Vec<u16> = data.chunks(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect();
                String::from_utf16(&utf16_data)
                    .map_err(|e| XlsError::Encoding(format!("UTF-16 decoding error: {}", e)))
            }
            XlsEncoding::Codepage(cp) => {
                // For now, assume Latin-1 for common codepages
                // In production, use proper codepage conversion
                String::from_utf8(data.to_vec())
                    .or_else(|_| {
                        // Fallback to lossy conversion
                        Ok(String::from_utf8_lossy(data).into_owned())
                    })
                    .map_err(|e: std::string::FromUtf8Error| XlsError::Encoding(format!("Codepage {} decoding error: {}", cp, e)))
            }
        }
    }
}

/// SST (Shared String Table) record
#[derive(Debug, Clone)]
pub struct SharedStringTable {
    pub strings: Vec<String>,
}

impl SharedStringTable {
    /// Parse SST from potentially multiple records (SST + CONTINUE)
    pub fn parse_from_records(records: &[Record], encoding: &XlsEncoding) -> XlsResult<Self> {
        if records.is_empty() {
            return Ok(SharedStringTable { strings: Vec::new() });
        }

        // Combine all SST and CONTINUE record data
        let mut combined_data = Vec::new();
        let mut found_sst = false;

        for record in records {
            match record.header.record_type {
                0x00FC => { // SST
                    if found_sst {
                        // Multiple SST records? Shouldn't happen
                        break;
                    }
                    found_sst = true;
                    // Skip the SST record header (4 bytes record type + 4 bytes length = 8 bytes total)
                    // But we need to include the SST data header (cstTotal + cstUnique = 8 bytes)
                    combined_data.extend_from_slice(&record.data);
                }
                0x003C => { // CONTINUE
                    if found_sst {
                        combined_data.extend_from_slice(&record.data);
                    }
                }
                _ => {
                    if found_sst {
                        // Stop when we hit a non-CONTINUE record after SST
                        break;
                    }
                }
            }
        }

        if combined_data.is_empty() {
            return Ok(SharedStringTable { strings: Vec::new() });
        }

        Self::parse(&combined_data, encoding)
    }

    pub fn parse(data: &[u8], encoding: &XlsEncoding) -> XlsResult<Self> {
        if data.len() < 8 {
            return Err(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        // Read SST header: cstTotal (4 bytes) and cstUnique (4 bytes)
        let cst_total = binary::read_u32_le(data, 0)? as usize;
        let cst_unique = binary::read_u32_le(data, 4)? as usize;

        let mut strings = Vec::with_capacity(cst_unique.min(10000)); // Cap for safety
        let mut offset = 8;

        // Parse each string entry in SST format
        for _ in 0..cst_unique {
            if offset + 3 > data.len() {
                break;
            }

            // Parse SST string format: cch (2 bytes) + flags (1 byte) + optional data
            let cch = binary::read_u16_le(data, offset)? as usize;
            let flags = data[offset + 2];

            let mut consumed = 3; // cch + flags

            // Rich text formatting (optional)
            if (flags & 0x08) != 0 {
                if offset + consumed + 2 > data.len() {
                    break;
                }
                let c_run = binary::read_u16_le(data, offset + consumed)?;
                consumed += 2;
                // Skip the formatting runs (4 bytes each)
                consumed += c_run as usize * 4;
            }

            // Phonetic information (optional)
            if (flags & 0x04) != 0 {
                if offset + consumed + 4 > data.len() {
                    break;
                }
                let cb_phonetic = binary::read_u32_le(data, offset + consumed)?;
                consumed += 4;
                // Skip the phonetic data
                consumed += cb_phonetic as usize;
            }

            // String data
            let is_unicode = (flags & 0x01) != 0;
            let string_len = if is_unicode { cch * 2 } else { cch };

            if offset + consumed + string_len > data.len() {
                break;
            }

            let string_data = &data[offset + consumed..offset + consumed + string_len];

            let string = if is_unicode {
                // UTF-16 LE
                String::from_utf16(&string_data
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect::<Vec<_>>())
                    .unwrap_or_default()
            } else {
                // 8-bit characters
                encoding.decode(string_data).unwrap_or_default()
            };

            strings.push(string);
            offset += consumed + string_len;
        }

        Ok(SharedStringTable { strings })
    }

    /// Parse a single string entry from SST data
    fn parse_string_entry(data: &[u8], encoding: &XlsEncoding) -> XlsResult<(String, usize)> {
        if data.len() < 3 {
            return Err(XlsError::InvalidLength {
                expected: 3,
                found: data.len(),
            });
        }

        // Read string header: cch (2 bytes) and flags (1 byte)
        let cch = binary::read_u16_le(data, 0)? as usize;
        let flags = data[2];

        let mut offset = 3;
        let mut consumed = 3;

        // Rich text formatting (optional)
        if (flags & 0x08) != 0 {
            if offset + 2 > data.len() {
                return Err(XlsError::InvalidData("Incomplete rich text header".to_string()));
            }
            let _cRun = binary::read_u16_le(data, offset)?;
            offset += 2;
            consumed += 2;
        }

        // Phonetic information (optional)
        if (flags & 0x04) != 0 {
            if offset + 4 > data.len() {
                return Err(XlsError::InvalidData("Incomplete phonetic header".to_string()));
            }
            let _cchPhonetic = binary::read_u32_le(data, offset)?;
            offset += 4;
            consumed += 4;
        }

        // String data
        let is_unicode = (flags & 0x01) != 0;
        let string_data;
        let string_consumed;

        if is_unicode {
            // UTF-16 LE
            let expected_bytes = cch * 2;
            if offset + expected_bytes > data.len() {
                return Err(XlsError::InvalidData("Incomplete Unicode string".to_string()));
            }
            string_data = &data[offset..offset + expected_bytes];
            string_consumed = expected_bytes;

            // Convert UTF-16 LE to String
            let utf16_words: Vec<u16> = string_data
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();

            match String::from_utf16(&utf16_words) {
                Ok(s) => Ok((s, consumed + string_consumed)),
                Err(e) => Err(XlsError::InvalidData(format!("Invalid UTF-16: {}", e))),
            }
        } else {
            // Compressed (8-bit characters)
            if offset + cch > data.len() {
                return Err(XlsError::InvalidData("Incomplete compressed string".to_string()));
            }
            string_data = &data[offset..offset + cch];
            string_consumed = cch;

            // Convert using the specified encoding
            match encoding.decode(string_data) {
                Ok(s) => Ok((s, consumed + string_consumed)),
                Err(e) => Err(XlsError::InvalidData(format!("Encoding error: {}", e))),
            }
        }
    }
}

/// XF (Extended Format) record - cell formatting
#[derive(Debug, Clone)]
pub struct ExtendedFormat {
    pub font_index: u16,
    pub format_index: u16,
}

impl ExtendedFormat {
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 4 {
            return Err(XlsError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let font_index = binary::read_u16_le_at(data, 0)?;
        let format_index = binary::read_u16_le_at(data, 2)?;

        Ok(ExtendedFormat {
            font_index,
            format_index,
        })
    }
}

/// Cell records
#[derive(Debug, Clone)]
pub enum CellRecord {
    Blank {
        row: u16,
        col: u16,
        xf_index: u16,
    },
    Number {
        row: u16,
        col: u16,
        xf_index: u16,
        value: f64,
    },
    Label {
        row: u16,
        col: u16,
        xf_index: u16,
        value: String,
    },
    BoolErr {
        row: u16,
        col: u16,
        xf_index: u16,
        value: BoolErrValue,
    },
    Rk {
        row: u16,
        col: u16,
        xf_index: u16,
        value: f64,
    },
    LabelSst {
        row: u16,
        col: u16,
        xf_index: u16,
        sst_index: u32,
    },
    Formula {
        row: u16,
        col: u16,
        xf_index: u16,
        value: FormulaValue,
        formula: Vec<u8>,
    },
}

#[derive(Debug, Clone)]
pub enum BoolErrValue {
    Bool(bool),
    Error(u8),
}

#[derive(Debug, Clone)]
pub enum FormulaValue {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
    Empty,
}

impl CellRecord {
    pub fn row(&self) -> u16 {
        match self {
            CellRecord::Blank { row, .. } => *row,
            CellRecord::Number { row, .. } => *row,
            CellRecord::Label { row, .. } => *row,
            CellRecord::BoolErr { row, .. } => *row,
            CellRecord::Rk { row, .. } => *row,
            CellRecord::LabelSst { row, .. } => *row,
            CellRecord::Formula { row, .. } => *row,
        }
    }

    pub fn col(&self) -> u16 {
        match self {
            CellRecord::Blank { col, .. } => *col,
            CellRecord::Number { col, .. } => *col,
            CellRecord::Label { col, .. } => *col,
            CellRecord::BoolErr { col, .. } => *col,
            CellRecord::Rk { col, .. } => *col,
            CellRecord::LabelSst { col, .. } => *col,
            CellRecord::Formula { col, .. } => *col,
        }
    }

    pub fn parse(record_type: u16, data: &[u8], encoding: &XlsEncoding) -> XlsResult<Self> {
        match record_type {
            0x0201 => Self::parse_blank(data), // Blank
            0x0203 => Self::parse_number(data), // Number
            0x0204 => Self::parse_label(data, encoding), // Label
            0x0205 => Self::parse_bool_err(data), // BoolErr
            0x027E => Self::parse_rk(data), // RK
            0x00FD => Self::parse_label_sst(data), // LabelSst
            0x0006 => Self::parse_formula(data), // Formula
            _ => Err(XlsError::InvalidRecord {
                record_type,
                message: "Unknown cell record type".to_string(),
            }),
        }
    }

    fn parse_blank(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 6 {
            return Err(XlsError::InvalidLength {
                expected: 6,
                found: data.len(),
            });
        }

        Ok(CellRecord::Blank {
            row: binary::read_u16_le_at(data, 0)?,
            col: binary::read_u16_le_at(data, 2)?,
            xf_index: binary::read_u16_le_at(data, 4)?,
        })
    }

    fn parse_number(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 14 {
            return Err(XlsError::InvalidLength {
                expected: 14,
                found: data.len(),
            });
        }

        Ok(CellRecord::Number {
            row: binary::read_u16_le_at(data, 0)?,
            col: binary::read_u16_le_at(data, 2)?,
            xf_index: binary::read_u16_le_at(data, 4)?,
            value: binary::read_f64_le_at(data, 6)?,
        })
    }

    fn parse_label(data: &[u8], encoding: &XlsEncoding) -> XlsResult<Self> {
        if data.len() < 8 {
            return Err(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let row = binary::read_u16_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 2)?;
        let xf_index = binary::read_u16_le_at(data, 4)?;
        let value = utils::parse_string_record(&data[6..], encoding)?;

        Ok(CellRecord::Label {
            row,
            col,
            xf_index,
            value,
        })
    }

    fn parse_bool_err(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 8 {
            return Err(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let row = binary::read_u16_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 2)?;
        let xf_index = binary::read_u16_le_at(data, 4)?;
        let value = if data[7] == 0 {
            BoolErrValue::Bool(data[6] != 0)
        } else {
            BoolErrValue::Error(data[6])
        };

        Ok(CellRecord::BoolErr {
            row,
            col,
            xf_index,
            value,
        })
    }

    fn parse_rk(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 10 {
            return Err(XlsError::InvalidLength {
                expected: 10,
                found: data.len(),
            });
        }

        let row = binary::read_u16_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 2)?;
        let xf_index = binary::read_u16_le_at(data, 4)?;
        let rk_value = binary::read_u32_le_at(data, 6)?;
        let value = utils::rk_to_f64(rk_value);

        Ok(CellRecord::Rk {
            row,
            col,
            xf_index,
            value,
        })
    }

    fn parse_label_sst(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 10 {
            return Err(XlsError::InvalidLength {
                expected: 10,
                found: data.len(),
            });
        }

        Ok(CellRecord::LabelSst {
            row: binary::read_u16_le_at(data, 0)?,
            col: binary::read_u16_le_at(data, 2)?,
            xf_index: binary::read_u16_le_at(data, 4)?,
            sst_index: binary::read_u32_le_at(data, 6)?,
        })
    }

    fn parse_formula(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 20 {
            return Err(XlsError::InvalidLength {
                expected: 20,
                found: data.len(),
            });
        }

        let row = binary::read_u16_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 2)?;
        let xf_index = binary::read_u16_le_at(data, 4)?;
        let value = utils::parse_formula_value(&data[6..14])?;
        let formula = data[20..].to_vec();

        Ok(CellRecord::Formula {
            row,
            col,
            xf_index,
            value,
            formula,
        })
    }
}
