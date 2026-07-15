//! BIFF record parsing for XLS files
//!
//! This module handles the parsing of BIFF (Binary Interchange File Format)
//! records used in Excel XLS files. BIFF records contain various types of
//! data including cell values, formatting, formulas, and metadata.

use std::io::{Read, Seek, SeekFrom};
use zerocopy::{FromBytes, LE, U16};

use crate::xls::error::{XlsError, XlsResult};
use crate::xls::utils;
use litchi_core::binary;

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
        let record_type = U16::<LE>::read_from_bytes(&buf[0..2])
            .map(|v| v.get())
            .unwrap_or(0);
        let data_len = U16::<LE>::read_from_bytes(&buf[2..4])
            .map(|v| v.get())
            .unwrap_or(0);

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
            },
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

    #[allow(dead_code)]
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
            .ok_or(XlsError::UnsupportedBiffVersion(biff_version))?;

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
            },
            14 => {
                // BIFF8 with 32-bit row indices
                Ok(DimensionsRecord {
                    first_row: binary::read_u32_le_at(data, 0)?,
                    last_row: binary::read_u32_le_at(data, 4)?,
                    first_col: binary::read_u16_le_at(data, 8)? as u32,
                    last_col: binary::read_u16_le_at(data, 10)? as u32,
                })
            },
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
#[allow(dead_code)]
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
    /// Create encoding from codepage identifier
    ///
    /// # Arguments
    ///
    /// * `codepage` - Windows codepage identifier (e.g., 1252 for Western European, 1200 for UTF-16LE)
    pub fn from_codepage(codepage: u16) -> XlsResult<Self> {
        match codepage {
            1200 => Ok(XlsEncoding::Utf16Le),
            cp => Ok(XlsEncoding::Codepage(cp)),
        }
    }

    /// Decode byte data using this encoding
    ///
    /// This method uses the shared codepage module for efficient and correct decoding.
    ///
    /// # Performance
    ///
    /// Uses zero-copy operations where possible and leverages optimized encoding_rs
    /// for codepage conversion.
    pub fn decode(&self, data: &[u8]) -> XlsResult<String> {
        match self {
            XlsEncoding::Utf16Le => {
                // Use shared UTF-16 LE decoder
                Ok(litchi_core::encoding::decode_utf16le(data))
            },
            XlsEncoding::Codepage(cp) => {
                // Use shared codepage decoder
                litchi_core::encoding::decode_bytes(data, Some(*cp as u32))
                    .ok_or_else(|| XlsError::Encoding(format!("Unsupported codepage: {}", cp)))
            },
        }
    }
}

/// SST (Shared String Table) record
#[derive(Debug, Clone)]
pub struct SharedStringTable {
    /// Plain text for each shared string, indexed by `LabelSst.isst`.
    pub strings: Vec<String>,
    /// Optional rich-text or phonetic properties, parallel to [`Self::strings`].
    ///
    /// Boxed sparse entries keep the common plain-string case compact.
    pub properties: Vec<Option<Box<SharedStringProperties>>>,
    /// Total number of references to shared strings in the workbook.
    pub total_count: u32,
}

/// Optional BIFF8 properties attached to a shared string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SharedStringProperties {
    /// Font changes in strictly increasing UTF-16 character positions.
    pub formatting_runs: Vec<SharedStringFormatRun>,
    /// East Asian phonetic (ruby) text and mappings, when present.
    pub phonetic: Option<PhoneticString>,
}

/// A BIFF8 `FormatRun` entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SharedStringFormatRun {
    pub character_index: u16,
    pub font_index: u16,
}

/// The character repertoire used for BIFF8 phonetic text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhoneticType {
    NarrowKatakana,
    WideKatakana,
    Hiragana,
    Any,
}

/// Horizontal alignment of BIFF8 phonetic text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhoneticAlignment {
    General,
    Left,
    Center,
    Distributed,
}

/// East Asian phonetic text stored in an SST `ExtRst` structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PhoneticString {
    pub font_index: u16,
    pub phonetic_type: PhoneticType,
    pub alignment: PhoneticAlignment,
    pub text: String,
    pub runs: Vec<PhoneticRun>,
    /// Producer-specific trailing bytes covered by `cbExtRst`.
    pub extra_data: Vec<u8>,
}

/// A BIFF8 `PhRuns` mapping from phonetic text to the base string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PhoneticRun {
    pub phonetic_text_index: u16,
    pub base_text_index: u16,
    pub base_text_length: u16,
}

impl SharedStringTable {
    /// Parse SST from potentially multiple records (SST + CONTINUE)
    pub fn parse_from_records(records: &[Record], encoding: &XlsEncoding) -> XlsResult<Self> {
        if records.is_empty() {
            return Ok(SharedStringTable {
                strings: Vec::new(),
                properties: Vec::new(),
                total_count: 0,
            });
        }

        if records[0].header.record_type != 0x00FC {
            return Err(XlsError::UnexpectedRecordType {
                expected: 0x00FC,
                found: records[0].header.record_type,
            });
        }
        if let Some(record) = records
            .iter()
            .skip(1)
            .find(|record| record.header.record_type != 0x003C)
        {
            return Err(XlsError::UnexpectedRecordType {
                expected: 0x003C,
                found: record.header.record_type,
            });
        }

        let segments: Vec<&[u8]> = records
            .iter()
            .map(|record| record.data.as_slice())
            .collect();
        Self::parse_segments(&segments, encoding)
    }

    pub fn parse(data: &[u8], encoding: &XlsEncoding) -> XlsResult<Self> {
        Self::parse_segments(&[data], encoding)
    }

    fn parse_segments(segments: &[&[u8]], _encoding: &XlsEncoding) -> XlsResult<Self> {
        let mut cursor = SstCursor::new(segments);
        cursor.ensure_current(8, "SST header")?;
        let total_count = cursor.read_u32_continued("SST total count")?;
        let unique_count = cursor.read_u32_continued("SST unique count")?;
        if total_count > i32::MAX as u32 || unique_count > i32::MAX as u32 {
            return Err(XlsError::InvalidData(
                "SST counts must be non-negative signed integers".to_string(),
            ));
        }
        if total_count < unique_count {
            return Err(XlsError::InvalidData(
                "SST total count is smaller than its unique count".to_string(),
            ));
        }

        let unique_count = unique_count as usize;
        let available = segments.iter().map(|segment| segment.len()).sum::<usize>();
        if unique_count > available.saturating_sub(8) / 3 {
            return Err(XlsError::InvalidData(format!(
                "SST declares {unique_count} strings but its records are too short"
            )));
        }

        let mut strings = Vec::new();
        let mut properties = Vec::new();
        strings.try_reserve_exact(unique_count).map_err(|error| {
            XlsError::InvalidData(format!("cannot allocate SST string index: {error}"))
        })?;
        properties
            .try_reserve_exact(unique_count)
            .map_err(|error| {
                XlsError::InvalidData(format!("cannot allocate SST property index: {error}"))
            })?;

        for string_index in 0..unique_count {
            cursor.ensure_current(3, "shared string header")?;
            let character_count = cursor.read_u16_continued("shared string character count")?;
            let flags = cursor.read_u8_continued("shared string flags")?;

            let run_count = if flags & 0x08 != 0 {
                cursor.ensure_current(2, "shared string rich-text count")?;
                cursor.read_u16_continued("shared string rich-text count")?
            } else {
                0
            };
            let extension_length = if flags & 0x04 != 0 {
                cursor.ensure_current(4, "shared string extension length")?;
                let length = cursor.read_u32_continued("shared string extension length")?;
                if length > i32::MAX as u32 {
                    return Err(XlsError::InvalidData(format!(
                        "shared string {string_index} has a negative extension length"
                    )));
                }
                length as usize
            } else {
                0
            };

            let text = cursor.read_characters(character_count, flags & 0x01 != 0)?;
            let formatting_runs =
                cursor.read_formatting_runs(run_count, character_count, string_index)?;
            let phonetic = if flags & 0x04 != 0 {
                let extension = cursor.read_bytes(extension_length, "shared string ExtRst")?;
                Some(parse_phonetic_string(
                    &extension,
                    character_count,
                    string_index,
                )?)
            } else {
                None
            };
            let property = if formatting_runs.is_empty() && phonetic.is_none() {
                None
            } else {
                Some(Box::new(SharedStringProperties {
                    formatting_runs,
                    phonetic,
                }))
            };
            strings.push(text);
            properties.push(property);
        }

        Ok(SharedStringTable {
            strings,
            properties,
            total_count,
        })
    }
}

struct SstCursor<'a> {
    segments: &'a [&'a [u8]],
    segment_index: usize,
    offset: usize,
}

impl<'a> SstCursor<'a> {
    fn new(segments: &'a [&'a [u8]]) -> Self {
        Self {
            segments,
            segment_index: 0,
            offset: 0,
        }
    }

    fn current(&self) -> &'a [u8] {
        self.segments
            .get(self.segment_index)
            .copied()
            .unwrap_or_default()
    }

    fn remaining(&self) -> usize {
        self.current().len().saturating_sub(self.offset)
    }

    fn ensure_current(&mut self, required: usize, context: &str) -> XlsResult<()> {
        while self.remaining() == 0 && self.segment_index + 1 < self.segments.len() {
            self.advance_segment(context)?;
        }
        if self.remaining() < required {
            return Err(XlsError::UnexpectedEndOfStream(format!(
                "{context} must fit in one BIFF record"
            )));
        }
        Ok(())
    }

    fn remaining_total(&self) -> usize {
        self.remaining()
            + self
                .segments
                .iter()
                .skip(self.segment_index + 1)
                .map(|segment| segment.len())
                .sum::<usize>()
    }

    fn advance_segment(&mut self, context: &str) -> XlsResult<()> {
        self.segment_index += 1;
        self.offset = 0;
        if self.segment_index >= self.segments.len() {
            return Err(XlsError::UnexpectedEndOfStream(context.to_string()));
        }
        Ok(())
    }

    fn read_u8_continued(&mut self, context: &str) -> XlsResult<u8> {
        if self.remaining() == 0 {
            self.advance_segment(context)?;
        }
        let value = self.current()[self.offset];
        self.offset += 1;
        Ok(value)
    }

    fn read_u16_continued(&mut self, context: &str) -> XlsResult<u16> {
        let mut bytes = [0; 2];
        self.read_exact(&mut bytes, context)?;
        Ok(u16::from_le_bytes(bytes))
    }

    fn read_u32_continued(&mut self, context: &str) -> XlsResult<u32> {
        let mut bytes = [0; 4];
        self.read_exact(&mut bytes, context)?;
        Ok(u32::from_le_bytes(bytes))
    }

    fn read_exact(&mut self, output: &mut [u8], context: &str) -> XlsResult<()> {
        let mut written = 0;
        while written < output.len() {
            if self.remaining() == 0 {
                self.advance_segment(context)?;
            }
            let count = self.remaining().min(output.len() - written);
            output[written..written + count]
                .copy_from_slice(&self.current()[self.offset..self.offset + count]);
            self.offset += count;
            written += count;
        }
        Ok(())
    }

    fn read_bytes(&mut self, length: usize, context: &str) -> XlsResult<Vec<u8>> {
        if length > self.remaining_total() {
            return Err(XlsError::UnexpectedEndOfStream(context.to_string()));
        }
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(length).map_err(|error| {
            XlsError::InvalidData(format!("cannot allocate {context}: {error}"))
        })?;
        bytes.resize(length, 0);
        self.read_exact(&mut bytes, context)?;
        Ok(bytes)
    }

    fn read_characters(&mut self, count: u16, mut high_byte: bool) -> XlsResult<String> {
        let mut characters = Vec::with_capacity(count as usize);
        while characters.len() < count as usize {
            let bytes_per_character = if high_byte { 2 } else { 1 };
            let available_characters = self.remaining() / bytes_per_character;
            let wanted = count as usize - characters.len();
            let chunk_characters = available_characters.min(wanted);

            if high_byte && self.remaining() % 2 != 0 && chunk_characters < wanted {
                return Err(XlsError::InvalidData(
                    "a UTF-16 shared string is split inside a code unit".to_string(),
                ));
            }
            for _ in 0..chunk_characters {
                let character = if high_byte {
                    let low = self.current()[self.offset];
                    let high = self.current()[self.offset + 1];
                    self.offset += 2;
                    u16::from_le_bytes([low, high])
                } else {
                    let character = self.current()[self.offset] as u16;
                    self.offset += 1;
                    character
                };
                characters.push(character);
            }

            if characters.len() == count as usize {
                break;
            }
            if self.remaining() != 0 {
                return Err(XlsError::InvalidData(
                    "shared string character data does not end at a record boundary".to_string(),
                ));
            }
            self.advance_segment("continued shared string character data")?;
            let continuation_flags = self.read_u8_continued("shared string continuation flags")?;
            if continuation_flags > 1 {
                return Err(XlsError::InvalidData(format!(
                    "invalid shared string continuation flags 0x{continuation_flags:02X}"
                )));
            }
            high_byte = continuation_flags == 1;
        }

        String::from_utf16(&characters)
            .map_err(|error| XlsError::Encoding(format!("UTF-16 decoding error: {error}")))
    }

    fn read_formatting_runs(
        &mut self,
        count: u16,
        character_count: u16,
        string_index: usize,
    ) -> XlsResult<Vec<SharedStringFormatRun>> {
        let mut runs = Vec::with_capacity(count as usize);
        let mut previous = None;
        for _ in 0..count {
            let character_index = self.read_u16_continued("shared string formatting run")?;
            let font_index = self.read_u16_continued("shared string formatting run")?;
            if character_index > character_count {
                return Err(XlsError::InvalidData(format!(
                    "shared string {string_index} has a formatting run past its text"
                )));
            }
            if previous.is_some_and(|value| character_index <= value) {
                return Err(XlsError::InvalidData(format!(
                    "shared string {string_index} formatting runs are not strictly increasing"
                )));
            }
            previous = Some(character_index);
            if character_index < character_count {
                runs.push(SharedStringFormatRun {
                    character_index,
                    font_index,
                });
            }
        }
        Ok(runs)
    }
}

fn parse_phonetic_string(
    data: &[u8],
    base_character_count: u16,
    string_index: usize,
) -> XlsResult<PhoneticString> {
    if data.len() < 14 {
        return Err(XlsError::InvalidLength {
            expected: 14,
            found: data.len(),
        });
    }
    // Both the marker and inner byte count are producer-controlled reserved
    // compatibility fields. MS-XLS requires readers to ignore the marker, and
    // Excel/POI accept stale inner counts while honoring outer cbExtRst.
    let _reserved = binary::read_u16_le(data, 0)?;
    let _payload_length = binary::read_u16_le(data, 2)?;

    let font_index = binary::read_u16_le(data, 4)?;
    let options = binary::read_u16_le(data, 6)?;
    let phonetic_type = match options & 0x0003 {
        0 => PhoneticType::NarrowKatakana,
        1 => PhoneticType::WideKatakana,
        2 => PhoneticType::Hiragana,
        _ => PhoneticType::Any,
    };
    let alignment = match (options >> 2) & 0x0003 {
        0 => PhoneticAlignment::General,
        1 => PhoneticAlignment::Left,
        2 => PhoneticAlignment::Center,
        _ => PhoneticAlignment::Distributed,
    };

    let run_count = binary::read_u16_le(data, 8)?;
    let character_count = binary::read_u16_le(data, 10)?;
    let repeated_character_count = binary::read_u16_le(data, 12)?;
    if run_count > 32767 || character_count > 32767 || character_count != repeated_character_count {
        return Err(XlsError::InvalidData(format!(
            "shared string {string_index} has invalid ExtRst string counts"
        )));
    }
    let text_byte_length = usize::from(character_count)
        .checked_mul(2)
        .ok_or_else(|| XlsError::InvalidData("ExtRst text length overflow".to_string()))?;
    let run_byte_length = usize::from(run_count)
        .checked_mul(6)
        .ok_or_else(|| XlsError::InvalidData("ExtRst run length overflow".to_string()))?;
    let required = 14usize
        .checked_add(text_byte_length)
        .and_then(|length| length.checked_add(run_byte_length))
        .ok_or_else(|| XlsError::InvalidData("ExtRst length overflow".to_string()))?;
    if required > data.len() {
        return Err(XlsError::InvalidLength {
            expected: required,
            found: data.len(),
        });
    }

    let text_bytes = &data[14..14 + text_byte_length];
    let text_words: Vec<u16> = text_bytes
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect();
    let text = String::from_utf16(&text_words)
        .map_err(|error| XlsError::Encoding(format!("ExtRst UTF-16 decoding error: {error}")))?;

    let mut runs = Vec::with_capacity(run_count as usize);
    let mut offset = 14 + text_byte_length;
    let mut previous_phonetic = None;
    let mut previous_base = None;
    let mut total_base_length = 0usize;
    for _ in 0..run_count {
        let phonetic_text_index = binary::read_u16_le(data, offset)?;
        let base_text_index = binary::read_u16_le(data, offset + 2)?;
        let base_text_length = binary::read_u16_le(data, offset + 4)?;
        if phonetic_text_index > 32767
            || base_text_index > 32767
            || base_text_length > 32767
            || phonetic_text_index >= character_count
            || base_text_index >= base_character_count
            || previous_phonetic.is_some_and(|value| phonetic_text_index <= value)
            || previous_base.is_some_and(|value| base_text_index <= value)
        {
            return Err(XlsError::InvalidData(format!(
                "shared string {string_index} has an invalid ExtRst phonetic run"
            )));
        }
        total_base_length = total_base_length.saturating_add(base_text_length as usize);
        previous_phonetic = Some(phonetic_text_index);
        previous_base = Some(base_text_index);
        runs.push(PhoneticRun {
            phonetic_text_index,
            base_text_index,
            base_text_length,
        });
        offset += 6;
    }
    if total_base_length > base_character_count as usize {
        return Err(XlsError::InvalidData(format!(
            "shared string {string_index} ExtRst runs exceed the base string"
        )));
    }

    Ok(PhoneticString {
        font_index,
        phonetic_type,
        alignment,
        text,
        runs,
        extra_data: data[required..].to_vec(),
    })
}

#[cfg(test)]
mod shared_string_tests {
    use super::*;

    fn record(record_type: u16, data: Vec<u8>) -> Record {
        Record {
            header: RecordHeader {
                record_type,
                data_len: data.len() as u16,
            },
            data,
        }
    }

    fn sst_header(total: u32, unique: u32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&total.to_le_bytes());
        data.extend_from_slice(&unique.to_le_bytes());
        data
    }

    #[test]
    fn parses_plain_compressed_and_utf16_strings() {
        let mut data = sst_header(3, 2);
        data.extend_from_slice(&[2, 0, 0, b'A', 0xC0]);
        data.extend_from_slice(&[1, 0, 1, 0x22, 0x6F]);

        let table = SharedStringTable::parse(&data, &XlsEncoding::Codepage(1251)).unwrap();

        assert_eq!(table.total_count, 3);
        // BIFF8 compressed Unicode supplies an implicit zero high byte; it is
        // not encoded in the workbook CODEPAGE.
        assert_eq!(table.strings, ["AÀ", "漢"]);
        assert_eq!(table.properties, [None, None]);
    }

    #[test]
    fn ignores_reserved_shared_string_flags() {
        let mut data = sst_header(1, 1);
        data.extend_from_slice(&[1, 0, 0xF2, b'A']);

        let table = SharedStringTable::parse(&data, &XlsEncoding::Utf16Le).unwrap();

        assert_eq!(table.strings, ["A"]);
    }

    #[test]
    fn parses_rich_text_after_the_character_data() {
        let mut data = sst_header(1, 1);
        data.extend_from_slice(&5u16.to_le_bytes());
        data.push(0x08);
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(b"Hello");
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());

        let table = SharedStringTable::parse(&data, &XlsEncoding::Utf16Le).unwrap();
        let properties = table.properties[0].as_deref().unwrap();

        assert_eq!(table.strings[0], "Hello");
        assert_eq!(
            properties.formatting_runs,
            [
                SharedStringFormatRun {
                    character_index: 0,
                    font_index: 1,
                },
                SharedStringFormatRun {
                    character_index: 2,
                    font_index: 3,
                },
            ]
        );
    }

    #[test]
    fn parses_phonetic_text_and_mappings() {
        let mut extension = Vec::new();
        extension.extend_from_slice(&1u16.to_le_bytes());
        extension.extend_from_slice(&20u16.to_le_bytes());
        extension.extend_from_slice(&7u16.to_le_bytes());
        extension.extend_from_slice(&10u16.to_le_bytes()); // Hiragana + centered
        extension.extend_from_slice(&1u16.to_le_bytes());
        extension.extend_from_slice(&2u16.to_le_bytes());
        extension.extend_from_slice(&2u16.to_le_bytes());
        for character in "とう".encode_utf16() {
            extension.extend_from_slice(&character.to_le_bytes());
        }
        extension.extend_from_slice(&0u16.to_le_bytes());
        extension.extend_from_slice(&0u16.to_le_bytes());
        extension.extend_from_slice(&2u16.to_le_bytes());
        assert_eq!(extension.len(), 24);

        let mut data = sst_header(1, 1);
        data.extend_from_slice(&2u16.to_le_bytes());
        data.push(0x05);
        data.extend_from_slice(&(extension.len() as u32).to_le_bytes());
        for character in "東京".encode_utf16() {
            data.extend_from_slice(&character.to_le_bytes());
        }
        data.extend_from_slice(&extension);

        let table = SharedStringTable::parse(&data, &XlsEncoding::Utf16Le).unwrap();
        let phonetic = table.properties[0]
            .as_deref()
            .unwrap()
            .phonetic
            .as_ref()
            .unwrap();

        assert_eq!(phonetic.font_index, 7);
        assert_eq!(phonetic.phonetic_type, PhoneticType::Hiragana);
        assert_eq!(phonetic.alignment, PhoneticAlignment::Center);
        assert_eq!(phonetic.text, "とう");
        assert_eq!(
            phonetic.runs,
            [PhoneticRun {
                phonetic_text_index: 0,
                base_text_index: 0,
                base_text_length: 2,
            }]
        );
    }

    #[test]
    fn changes_character_width_at_continue_boundaries() {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&4u16.to_le_bytes());
        first.push(0);
        first.extend_from_slice(b"AB");
        let mut second = vec![1];
        for character in "漢字".encode_utf16() {
            second.extend_from_slice(&character.to_le_bytes());
        }

        let records = [record(0x00FC, first), record(0x003C, second)];
        let table = SharedStringTable::parse_from_records(&records, &XlsEncoding::Utf16Le).unwrap();

        assert_eq!(table.strings, ["AB漢字"]);
    }

    #[test]
    fn reads_formatting_runs_split_across_continue_records() {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&2u16.to_le_bytes());
        first.push(0x08);
        first.extend_from_slice(&1u16.to_le_bytes());
        first.extend_from_slice(b"AB");
        first.push(1); // first byte of character_index
        let second = vec![0, 9, 0];

        let records = [record(0x00FC, first), record(0x003C, second)];
        let table = SharedStringTable::parse_from_records(&records, &XlsEncoding::Utf16Le).unwrap();

        assert_eq!(
            table.properties[0].as_deref().unwrap().formatting_runs,
            [SharedStringFormatRun {
                character_index: 1,
                font_index: 9,
            }]
        );
    }

    #[test]
    fn rejects_invalid_continue_flags_and_truncated_extensions() {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&2u16.to_le_bytes());
        first.push(0);
        first.push(b'A');
        let bad_flags = [record(0x00FC, first), record(0x003C, vec![2, b'B'])];
        assert!(SharedStringTable::parse_from_records(&bad_flags, &XlsEncoding::Utf16Le).is_err());

        let mut truncated = sst_header(1, 1);
        truncated.extend_from_slice(&1u16.to_le_bytes());
        truncated.push(0x04);
        truncated.extend_from_slice(&100u32.to_le_bytes());
        truncated.push(b'A');
        assert!(SharedStringTable::parse(&truncated, &XlsEncoding::Utf16Le).is_err());
    }

    #[test]
    fn reads_writer_generated_multirecord_sst() {
        let expected = vec!["a".repeat(9000), "漢".repeat(5000)];
        let mut bytes = Vec::new();
        crate::xls::writer::biff::write_sst(&mut bytes, &expected, 2).unwrap();
        let records: Vec<Record> = RecordIter::new(std::io::Cursor::new(bytes))
            .unwrap()
            .collect::<XlsResult<_>>()
            .unwrap();

        let table = SharedStringTable::parse_from_records(&records, &XlsEncoding::Utf16Le).unwrap();

        assert_eq!(table.strings, expected);
    }
}

/// XF (Extended Format) record - cell formatting
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct ExtendedFormat {
    pub font_index: u16,
    pub format_index: u16,
}

#[allow(dead_code)]
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
            0x0201 => Self::parse_blank(data),           // Blank
            0x0203 => Self::parse_number(data),          // Number
            0x0204 => Self::parse_label(data, encoding), // Label
            0x0205 => Self::parse_bool_err(data),        // BoolErr
            0x027E => Self::parse_rk(data),              // RK
            0x00FD => Self::parse_label_sst(data),       // LabelSst
            0x0006 => Self::parse_formula(data),         // Formula
            _ => Err(XlsError::InvalidRecord {
                record_type,
                message: "Unknown cell record type".to_string(),
            }),
        }
    }

    pub(crate) fn parse_mul_rk(data: &[u8]) -> XlsResult<Vec<Self>> {
        let (row, first_col, count) = Self::packed_cell_range(data, 6, "MulRk")?;
        let mut cells = Vec::with_capacity(count);
        for index in 0..count {
            let offset = 4 + index * 6;
            cells.push(Self::Rk {
                row,
                col: first_col + index as u16,
                xf_index: binary::read_u16_le_at(data, offset)?,
                value: utils::rk_to_f64(binary::read_u32_le_at(data, offset + 2)?),
            });
        }
        Ok(cells)
    }

    pub(crate) fn parse_mul_blank(data: &[u8]) -> XlsResult<Vec<Self>> {
        let (row, first_col, count) = Self::packed_cell_range(data, 2, "MulBlank")?;
        let mut cells = Vec::with_capacity(count);
        for index in 0..count {
            cells.push(Self::Blank {
                row,
                col: first_col + index as u16,
                xf_index: binary::read_u16_le_at(data, 4 + index * 2)?,
            });
        }
        Ok(cells)
    }

    fn packed_cell_range(
        data: &[u8],
        item_size: usize,
        record_name: &str,
    ) -> XlsResult<(u16, u16, usize)> {
        let Some(items_size) = data.len().checked_sub(6) else {
            return Err(XlsError::InvalidLength {
                expected: 6 + item_size * 2,
                found: data.len(),
            });
        };
        if items_size % item_size != 0 {
            return Err(XlsError::InvalidData(format!(
                "{record_name} payload does not contain whole packed cells"
            )));
        }
        let count = items_size / item_size;
        if !(2..=256).contains(&count) {
            return Err(XlsError::InvalidData(format!(
                "{record_name} contains {count} cells; expected 2 through 256"
            )));
        }

        let row = binary::read_u16_le_at(data, 0)?;
        let first_col = binary::read_u16_le_at(data, 2)?;
        let last_col = binary::read_u16_le_at(data, data.len() - 2)?;
        let expected_last = first_col
            .checked_add((count - 1) as u16)
            .ok_or_else(|| XlsError::InvalidData(format!("{record_name} column overflow")))?;
        if first_col > 254 || last_col != expected_last || last_col > 255 {
            return Err(XlsError::InvalidData(format!(
                "{record_name} column range {first_col}..={last_col} does not match {count} cells"
            )));
        }
        Ok((row, first_col, count))
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

#[cfg(test)]
mod packed_cell_tests {
    use super::*;

    #[test]
    fn expands_mul_rk_into_individual_numeric_cells() {
        let mut data = Vec::new();
        data.extend_from_slice(&7u16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&((42u32 << 2) | 0x02).to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&((1234u32 << 2) | 0x03).to_le_bytes());
        data.extend_from_slice(&4u16.to_le_bytes());

        let cells = CellRecord::parse_mul_rk(&data).unwrap();

        assert!(matches!(
            cells[0],
            CellRecord::Rk {
                row: 7,
                col: 3,
                xf_index: 1,
                value: 42.0
            }
        ));
        assert!(matches!(
            cells[1],
            CellRecord::Rk {
                row: 7,
                col: 4,
                xf_index: 2,
                value
            } if value == 12.34
        ));
    }

    #[test]
    fn expands_mul_blank_and_rejects_inconsistent_ranges() {
        let mut data = Vec::new();
        data.extend_from_slice(&9u16.to_le_bytes());
        data.extend_from_slice(&5u16.to_le_bytes());
        data.extend_from_slice(&11u16.to_le_bytes());
        data.extend_from_slice(&12u16.to_le_bytes());
        data.extend_from_slice(&6u16.to_le_bytes());

        let cells = CellRecord::parse_mul_blank(&data).unwrap();
        assert!(matches!(
            cells.as_slice(),
            [
                CellRecord::Blank {
                    row: 9,
                    col: 5,
                    xf_index: 11
                },
                CellRecord::Blank {
                    row: 9,
                    col: 6,
                    xf_index: 12
                }
            ]
        ));

        let last_column_offset = data.len() - 2;
        data[last_column_offset..].copy_from_slice(&7u16.to_le_bytes());
        assert!(CellRecord::parse_mul_blank(&data).is_err());
    }
}
