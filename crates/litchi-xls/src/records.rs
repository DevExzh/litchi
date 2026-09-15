//! BIFF record parsing for XLS files
//!
//! This module handles the parsing of BIFF (Binary Interchange File Format)
//! records used in Excel XLS files. BIFF records contain various types of
//! data including cell values, formatting, formulas, and metadata.

use crate::error::{Error, Result};
use crate::utils;
use litchi_biff::RecordRef;
use litchi_codepage::{Mbcs, Page};
use litchi_core::binary;

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
    #[must_use]
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

    #[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
    #[must_use]
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 4 {
            return Err(Error::InvalidLength {
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
            .ok_or(Error::UnsupportedBiffVersion(biff_version))?;

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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        match data.len() {
            10 => {
                // BIFF5-BIFF8
                Ok(DimensionsRecord {
                    first_row: u32::from(binary::read_u16_le_at(data, 0)?),
                    last_row: u32::from(binary::read_u16_le_at(data, 2)?),
                    first_col: u32::from(binary::read_u16_le_at(data, 4)?),
                    last_col: u32::from(binary::read_u16_le_at(data, 6)?),
                })
            },
            14 => {
                // BIFF8 with 32-bit row indices
                Ok(DimensionsRecord {
                    first_row: binary::read_u32_le_at(data, 0)?,
                    last_row: binary::read_u32_le_at(data, 4)?,
                    first_col: u32::from(binary::read_u16_le_at(data, 8)?),
                    last_col: u32::from(binary::read_u16_le_at(data, 10)?),
                })
            },
            _ => Err(Error::InvalidLength {
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_u8(value: u8) -> Result<Self> {
        match value & 0x3 {
            0x00 => Ok(SheetVisible::Visible),
            0x01 => Ok(SheetVisible::Hidden),
            0x02 => Ok(SheetVisible::VeryHidden),
            v => Err(Error::InvalidRecord {
                record_type: 0x0085, // BoundSheet8
                message: format!("Invalid visibility value: {v}"),
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_u8(value: u8) -> Result<Self> {
        match value {
            0x00 => Ok(SheetType::WorkSheet),
            0x01 => Ok(SheetType::MacroSheet),
            0x02 => Ok(SheetType::ChartSheet),
            0x06 => Ok(SheetType::VBModule),
            v => Err(Error::InvalidRecord {
                record_type: 0x0085, // BoundSheet8
                message: format!("Invalid sheet type: {v}"),
            }),
        }
    }
}

/// `BoundSheet8` record (worksheet metadata)
#[derive(Debug, Clone)]
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub struct BoundSheetRecord {
    pub position: u32,
    pub visible: SheetVisible,
    pub sheet_type: SheetType,
    pub name: String,
}

impl BoundSheetRecord {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8], encoding: &Encoding) -> Result<Self> {
        if data.len() < 8 {
            return Err(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let character_count = usize::from(data[6]);
        let string_flags = data[7];
        if string_flags & 0xfe != 0 {
            return Err(Error::InvalidRecord {
                record_type: 0x0085,
                message: "BoundSheet8 name has reserved string option bits".to_string(),
            });
        }
        let character_width = if string_flags & 1 != 0 { 2 } else { 1 };
        if data.len() != 8 + character_count * character_width {
            return Err(Error::InvalidRecord {
                record_type: 0x0085,
                message: "BoundSheet8 name length does not match its payload".to_string(),
            });
        }

        let position = binary::read_u32_le_at(data, 0)?;
        let visible = SheetVisible::from_u8(data[4])?;
        let sheet_type = SheetType::from_u8(data[5])?;

        // Skip 2 bytes and parse the name
        let name_data = &data[6..];
        let name = utils::parse_short_string(name_data, encoding)?;
        let name_length = name.encode_utf16().count();
        let forbidden = |character| {
            matches!(
                character,
                '\0' | '\u{0003}' | ':' | '\\' | '*' | '?' | '/' | '[' | ']'
            )
        };
        if !(1..=31).contains(&name_length)
            || name.chars().any(forbidden)
            || name.starts_with('\'')
            || name.ends_with('\'')
        {
            return Err(Error::InvalidRecord {
                record_type: 0x0085,
                message: format!("Invalid BoundSheet8 sheet name: {name:?}"),
            });
        }

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
pub enum Encoding {
    /// NUL-terminated byte-stream encoding with a checked code page.
    Codepage(Mbcs),
    /// UTF-16 little endian (BIFF8+)
    Utf16Le,
}

impl Encoding {
    /// Create encoding from codepage identifier
    ///
    /// # Arguments
    ///
    /// * `codepage` - Windows codepage identifier (e.g., 1252 for Western European, 1200 for UTF-16LE)
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_codepage(codepage: u16) -> Result<Self> {
        match codepage {
            1200 => Ok(Encoding::Utf16Le),
            cp => Mbcs::require(u32::from(cp))
                .map(Encoding::Codepage)
                .map_err(|error| Error::Encoding(error.to_string())),
        }
    }

    /// Decode byte data using this encoding
    ///
    /// Record terminators are handled here while the shared codec performs
    /// strict conversion without guessing format-specific boundaries.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn decode(&self, data: &[u8]) -> Result<String> {
        match self {
            Encoding::Utf16Le => {
                if !data.len().is_multiple_of(2) {
                    return Err(Error::Encoding(
                        "UTF-16LE text has an odd byte length".to_string(),
                    ));
                }
                let end = data
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .position(|pair| *pair == [0, 0])
                    .map_or(data.len(), |units| units * 2);
                Page::UTF_16LE
                    .decode(&data[..end])
                    .map(std::borrow::Cow::into_owned)
                    .map_err(|error| Error::Encoding(error.to_string()))
            },
            Encoding::Codepage(page) => {
                let end = data
                    .iter()
                    .position(|byte| *byte == 0)
                    .unwrap_or(data.len());
                page.decode(&data[..end])
                    .map(std::borrow::Cow::into_owned)
                    .map_err(|error| Error::Encoding(error.to_string()))
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_from_records(records: &[RecordRef<'_>], encoding: &Encoding) -> Result<Self> {
        if records.is_empty() {
            return Ok(SharedStringTable {
                strings: Vec::new(),
                properties: Vec::new(),
                total_count: 0,
            });
        }

        if records[0].kind().get() != 0x00FC {
            return Err(Error::UnexpectedRecordType {
                expected: 0x00FC,
                found: records[0].kind().get(),
            });
        }
        if let Some(record) = records
            .iter()
            .skip(1)
            .find(|record| record.kind().get() != 0x003C)
        {
            return Err(Error::UnexpectedRecordType {
                expected: 0x003C,
                found: record.kind().get(),
            });
        }

        let segments: Vec<&[u8]> = records.iter().map(|record| record.payload()).collect();
        Self::parse_segments(&segments, encoding)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8], encoding: &Encoding) -> Result<Self> {
        Self::parse_segments(&[data], encoding)
    }

    fn parse_segments(segments: &[&[u8]], _encoding: &Encoding) -> Result<Self> {
        let mut cursor = SstCursor::new(segments);
        cursor.ensure_current(8, "SST header")?;
        let total_count = cursor.read_u32_continued("SST total count")?;
        let unique_count = cursor.read_u32_continued("SST unique count")?;
        if total_count > i32::MAX as u32 || unique_count > i32::MAX as u32 {
            return Err(Error::InvalidData(
                "SST counts must be non-negative signed integers".to_string(),
            ));
        }
        if total_count < unique_count {
            return Err(Error::InvalidData(
                "SST total count is smaller than its unique count".to_string(),
            ));
        }

        let unique_count = unique_count as usize;
        let available = segments.iter().map(|segment| segment.len()).sum::<usize>();
        if unique_count > available.saturating_sub(8) / 3 {
            return Err(Error::InvalidData(format!(
                "SST declares {unique_count} strings but its records are too short"
            )));
        }

        let mut strings = Vec::new();
        let mut properties = Vec::new();
        strings.try_reserve_exact(unique_count).map_err(|error| {
            Error::InvalidData(format!("cannot allocate SST string index: {error}"))
        })?;
        properties
            .try_reserve_exact(unique_count)
            .map_err(|error| {
                Error::InvalidData(format!("cannot allocate SST property index: {error}"))
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
                    return Err(Error::InvalidData(format!(
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

/// Receives shared-string code units as [`SstCursor::walk_characters`] frames
/// them.
///
/// The implementations differ only in what they do with the code units; where
/// the units are, and which framings are refused, is decided once by the walk.
/// Runs are delivered as whole slices so that neither implementation pays the
/// per-code-unit segment lookup the framing walk would otherwise repeat.
trait CodeUnitSink {
    /// Accepts a run of little-endian UTF-16 code units. `chunk.len()` is even.
    fn push_wide(&mut self, chunk: &[u8]);

    /// Accepts a run of BIFF8 compressed-Unicode bytes. Each byte is an
    /// implicit `U+0000..=U+00FF` code unit with a zero high byte, so no
    /// compressed code unit can be a surrogate.
    fn push_compressed(&mut self, chunk: &[u8]);
}

/// Materializes the code units so that they can be transcoded to UTF-8.
struct CollectedCodeUnits {
    units: Vec<u16>,
}

impl CollectedCodeUnits {
    fn with_capacity(count: u16) -> Result<Self> {
        let mut units = Vec::new();
        units.try_reserve_exact(count as usize).map_err(|error| {
            Error::InvalidData(format!("cannot allocate shared string characters: {error}"))
        })?;
        Ok(Self { units })
    }
}

impl CodeUnitSink for CollectedCodeUnits {
    fn push_wide(&mut self, chunk: &[u8]) {
        self.units.extend(
            chunk
                .as_chunks::<2>()
                .0
                .iter()
                .map(|pair| u16::from_le_bytes(*pair)),
        );
    }

    fn push_compressed(&mut self, chunk: &[u8]) {
        self.units.extend(chunk.iter().copied().map(u16::from));
    }
}

/// Decides UTF-16 well-formedness without materializing anything.
///
/// `String::from_utf16` succeeds exactly when every high surrogate is
/// immediately followed by a low surrogate and no low surrogate stands alone,
/// so carrying one pending high surrogate across chunks — and therefore across
/// `Continue` record boundaries, including a boundary that switches between
/// compressed and uncompressed encoding — decides the same predicate.
#[derive(Default)]
struct SurrogatePairing {
    pending_high: bool,
    malformed: bool,
}

impl SurrogatePairing {
    /// Reports whether `String::from_utf16` would have accepted the walked
    /// units. A high surrogate still pending at the end of the walk is the
    /// end-of-input case `char::decode_utf16` also refuses.
    fn is_well_formed(&self) -> bool {
        !self.malformed && !self.pending_high
    }
}

impl CodeUnitSink for SurrogatePairing {
    fn push_wide(&mut self, chunk: &[u8]) {
        let pairs = chunk.as_chunks::<2>().0;
        // A little-endian code unit is a surrogate exactly when its high byte
        // lies in `0xD8..=0xDF`, so a chunk containing no such byte cannot
        // change either flag. This test is a byte scan the compiler can widen.
        if !self.pending_high && !pairs.iter().any(|pair| pair[1] & 0xF8 == 0xD8) {
            return;
        }
        let mut pending_high = self.pending_high;
        let mut malformed = self.malformed;
        for pair in pairs {
            let unit = u16::from_le_bytes(*pair);
            if pending_high {
                pending_high = false;
                if (0xDC00..=0xDFFF).contains(&unit) {
                    continue;
                }
                // `char::decode_utf16` reports the unpaired high surrogate and
                // then re-examines this unit, so this walk re-examines it too.
                // The re-examination cannot change the answer once `malformed`
                // is set — it is kept so that the two walks have the same shape.
                malformed = true;
            }
            match unit {
                0xD800..=0xDBFF => pending_high = true,
                0xDC00..=0xDFFF => malformed = true,
                _ => {},
            }
        }
        self.pending_high = pending_high;
        self.malformed = malformed;
    }

    fn push_compressed(&mut self, chunk: &[u8]) {
        // No compressed code unit is a surrogate, so the only reachable state
        // change is that a pending high surrogate is now followed by something
        // that cannot complete it.
        if self.pending_high && !chunk.is_empty() {
            self.pending_high = false;
            self.malformed = true;
        }
    }
}

struct SstCursor<'a> {
    segments: &'a [&'a [u8]],
    segment_index: usize,
    offset: usize,
    /// Summed length of every segment before `segment_index`.
    ///
    /// [`SstCursor::logical_position`] used to recompute this sum on every
    /// call, which made indexing one shared string cost a walk over every
    /// segment behind it. It is maintained instead by
    /// [`SstCursor::advance_segment`] and restored by [`SstCursor::seek`], and
    /// a debug assertion in `logical_position` recomputes the old sum and
    /// compares, so every test run proves the two agree.
    logical_base: usize,
}

/// A cursor position that [`SstCursor::seek`] can restore exactly.
#[derive(Clone, Copy)]
struct SstPosition {
    segment_index: usize,
    offset: usize,
    logical_base: usize,
}

impl<'a> SstCursor<'a> {
    fn new(segments: &'a [&'a [u8]]) -> Self {
        Self {
            segments,
            segment_index: 0,
            offset: 0,
            logical_base: 0,
        }
    }

    fn position(&self) -> SstPosition {
        SstPosition {
            segment_index: self.segment_index,
            offset: self.offset,
            logical_base: self.logical_base,
        }
    }

    fn seek(&mut self, at: SstPosition) {
        self.segment_index = at.segment_index;
        self.offset = at.offset;
        self.logical_base = at.logical_base;
    }

    /// Takes `N` bytes when they are all resident in the current segment.
    ///
    /// The header fields of a shared string are two or four bytes and almost
    /// always lie inside one `SST` or `Continue` payload. Reading them through
    /// [`SstCursor::read_exact`] cost a `copy_from_slice` call per field; this
    /// is one fixed-width load, and the continuation-crossing case still falls
    /// through to `read_exact`, which owns the only implementation of it.
    #[inline]
    fn take_resident<const N: usize>(&mut self) -> Option<[u8; N]> {
        let end = self.offset.checked_add(N)?;
        let bytes = <[u8; N]>::try_from(self.current().get(self.offset..end)?).ok()?;
        self.offset = end;
        Some(bytes)
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

    fn ensure_current(&mut self, required: usize, context: &str) -> Result<()> {
        while self.remaining() == 0 && self.segment_index + 1 < self.segments.len() {
            self.advance_segment(context)?;
        }
        if self.remaining() < required {
            return Err(Error::UnexpectedEndOfStream(format!(
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

    fn advance_segment(&mut self, context: &str) -> Result<()> {
        // The segment being left is `segment_index`; past the last segment
        // `current()` is empty and adds nothing, which is what the summed form
        // did when `take` ran past the end.
        self.logical_base = self.logical_base.saturating_add(self.current().len());
        self.segment_index += 1;
        self.offset = 0;
        if self.segment_index >= self.segments.len() {
            return Err(Error::UnexpectedEndOfStream(context.to_string()));
        }
        Ok(())
    }

    fn read_u8_continued(&mut self, context: &str) -> Result<u8> {
        while self.remaining() == 0 {
            self.advance_segment(context)?;
        }
        let value = self.current()[self.offset];
        self.offset += 1;
        Ok(value)
    }

    fn read_u16_continued(&mut self, context: &str) -> Result<u16> {
        if let Some(bytes) = self.take_resident::<2>() {
            return Ok(u16::from_le_bytes(bytes));
        }
        let mut bytes = [0; 2];
        self.read_exact(&mut bytes, context)?;
        Ok(u16::from_le_bytes(bytes))
    }

    fn read_u32_continued(&mut self, context: &str) -> Result<u32> {
        if let Some(bytes) = self.take_resident::<4>() {
            return Ok(u32::from_le_bytes(bytes));
        }
        let mut bytes = [0; 4];
        self.read_exact(&mut bytes, context)?;
        Ok(u32::from_le_bytes(bytes))
    }

    fn read_exact(&mut self, output: &mut [u8], context: &str) -> Result<()> {
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

    fn read_bytes(&mut self, length: usize, context: &str) -> Result<Vec<u8>> {
        if length > self.remaining_total() {
            return Err(Error::UnexpectedEndOfStream(context.to_string()));
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|error| Error::InvalidData(format!("cannot allocate {context}: {error}")))?;
        bytes.resize(length, 0);
        self.read_exact(&mut bytes, context)?;
        Ok(bytes)
    }

    /// Walks `count` code units of shared-string character data, following
    /// `Continue` record boundaries and their continuation flags, and hands
    /// every code unit to `sink`.
    ///
    /// This is the only implementation of shared-string character framing.
    /// [`Self::read_characters`], which materializes the text, and
    /// [`Self::measure_characters`], which only advances over it, both walk
    /// through here, so the two can never disagree about where a string ends or
    /// about which malformed framings are refused, in which order.
    fn walk_characters<S: CodeUnitSink>(
        &mut self,
        count: u16,
        mut high_byte: bool,
        sink: &mut S,
    ) -> Result<()> {
        let count = count as usize;
        let mut consumed = 0usize;
        while consumed < count {
            let bytes_per_character = if high_byte { 2 } else { 1 };
            let remaining = self.remaining();
            let available_characters = remaining / bytes_per_character;
            let wanted = count - consumed;
            let chunk_characters = available_characters.min(wanted);

            if high_byte && !remaining.is_multiple_of(2) && chunk_characters < wanted {
                return Err(Error::InvalidData(
                    "a UTF-16 shared string is split inside a code unit".to_string(),
                ));
            }
            // `chunk_characters <= remaining / bytes_per_character`, so the
            // chunk is inside the current segment by construction.
            let chunk_bytes = chunk_characters * bytes_per_character;
            let chunk = &self.current()[self.offset..self.offset + chunk_bytes];
            if high_byte {
                sink.push_wide(chunk);
            } else {
                sink.push_compressed(chunk);
            }
            self.offset += chunk_bytes;
            consumed += chunk_characters;

            if consumed == count {
                break;
            }
            if self.remaining() != 0 {
                return Err(Error::InvalidData(
                    "shared string character data does not end at a record boundary".to_string(),
                ));
            }
            self.advance_segment("continued shared string character data")?;
            let continuation_flags = self.read_u8_continued("shared string continuation flags")?;
            if continuation_flags > 1 {
                return Err(Error::InvalidData(format!(
                    "invalid shared string continuation flags 0x{continuation_flags:02X}"
                )));
            }
            high_byte = continuation_flags == 1;
        }
        Ok(())
    }

    fn read_characters(&mut self, count: u16, high_byte: bool) -> Result<String> {
        let mut sink = CollectedCodeUnits::with_capacity(count)?;
        self.walk_characters(count, high_byte, &mut sink)?;
        String::from_utf16(&sink.units)
            .map_err(|error| Error::Encoding(format!("UTF-16 decoding error: {error}")))
    }

    /// Advances over exactly the character data [`Self::read_characters`] would
    /// consume, and refuses exactly the inputs it refuses, without allocating
    /// the `Vec<u16>`, allocating the `String`, or transcoding to UTF-8.
    ///
    /// UTF-16 well-formedness is still decided, streaming, as the walk runs. A
    /// malformed string takes a cold path that rewinds and materializes the
    /// same span, so that the refusal is produced by `String::from_utf16`
    /// itself and its message is byte-identical to the materializing path's.
    fn measure_characters(&mut self, count: u16, high_byte: bool) -> Result<()> {
        let restart = self.position();
        let mut sink = SurrogatePairing::default();
        self.walk_characters(count, high_byte, &mut sink)?;
        if sink.is_well_formed() {
            return Ok(());
        }
        // Cold path only: the walk above has already proven this string is
        // malformed, so the cost of walking it a second time does not matter.
        self.seek(restart);
        self.read_characters(count, high_byte).map(|_| ())
    }

    /// Retains every formatting run, for the eager `SharedStringTable`.
    fn read_formatting_runs(
        &mut self,
        count: u16,
        character_count: u16,
        string_index: usize,
    ) -> Result<Vec<SharedStringFormatRun>> {
        self.walk_formatting_runs(count, character_count, string_index)
    }

    /// Walks `count` formatting runs, validating each one, and hands the ones
    /// inside the text to `S`.
    ///
    /// This is the only implementation of formatting-run validation. The eager
    /// table instantiates it with `Vec<SharedStringFormatRun>`; the
    /// source-backed shared-string walk, which discards what it returns,
    /// instantiates it with [`MeasuredRuns`] and allocates nothing per string.
    ///
    /// The reservation is part of the sink, so the measure sink also drops the
    /// refusal that guarded it: see [`MeasuredRuns`].
    fn walk_formatting_runs<S: FormatRunSink>(
        &mut self,
        count: u16,
        character_count: u16,
        string_index: usize,
    ) -> Result<S> {
        let mut runs = S::with_capacity(count)?;
        let mut previous = None;
        for _ in 0..count {
            let character_index = self.read_u16_continued("shared string formatting run")?;
            let font_index = self.read_u16_continued("shared string formatting run")?;
            if character_index > character_count {
                return Err(Error::InvalidData(format!(
                    "shared string {string_index} has a formatting run past its text"
                )));
            }
            if previous.is_some_and(|value| character_index <= value) {
                return Err(Error::InvalidData(format!(
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
) -> Result<PhoneticString> {
    if data.len() < 14 {
        return Err(Error::InvalidLength {
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
        return Err(Error::InvalidData(format!(
            "shared string {string_index} has invalid ExtRst string counts"
        )));
    }
    let text_byte_length = usize::from(character_count)
        .checked_mul(2)
        .ok_or_else(|| Error::InvalidData("ExtRst text length overflow".to_string()))?;
    let run_byte_length = usize::from(run_count)
        .checked_mul(6)
        .ok_or_else(|| Error::InvalidData("ExtRst run length overflow".to_string()))?;
    let required = 14usize
        .checked_add(text_byte_length)
        .and_then(|length| length.checked_add(run_byte_length))
        .ok_or_else(|| Error::InvalidData("ExtRst length overflow".to_string()))?;
    if required > data.len() {
        return Err(Error::InvalidLength {
            expected: required,
            found: data.len(),
        });
    }

    let text_bytes = &data[14..14 + text_byte_length];
    let text_words: Vec<u16> = text_bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect();
    let text = String::from_utf16(&text_words)
        .map_err(|error| Error::Encoding(format!("ExtRst UTF-16 decoding error: {error}")))?;

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
            return Err(Error::InvalidData(format!(
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
        return Err(Error::InvalidData(format!(
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

#[derive(Debug, Clone, Copy)]
pub(crate) struct SharedStringSstSegment {
    pub(crate) source_offset: u64,
    pub(crate) logical_offset: usize,
    pub(crate) len: usize,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct SharedStringEntryLocation {
    pub(crate) start: usize,
    pub(crate) end: usize,
}

#[derive(Debug)]
pub(crate) struct SharedStringSstScan {
    pub(crate) segments: Vec<SharedStringSstSegment>,
    pub(crate) entries: Vec<SharedStringEntryLocation>,
}

impl SharedStringSstScan {
    pub(crate) fn empty() -> Self {
        Self {
            segments: Vec::new(),
            entries: Vec::new(),
        }
    }
}

#[derive(Debug)]
pub(crate) enum SharedStringScanError {
    Biff(Error),
    Invalid(String),
    Allocation {
        resource: &'static str,
        requested: usize,
    },
}

impl From<Error> for SharedStringScanError {
    fn from(error: Error) -> Self {
        Self::Biff(error)
    }
}

impl<'a> SstCursor<'a> {
    /// The cursor's offset in the concatenation of every segment, which is what
    /// a [`SharedStringEntryLocation`] records.
    ///
    /// The summed form this replaces walked every segment behind the cursor on
    /// every call, and the SST scan calls it twice per shared string, so
    /// indexing an SST spread over `s` `Continue` records cost O(strings × s).
    /// The debug assertion recomputes that sum, so every debug test run — the
    /// corpus differential included — proves the maintained base still equals
    /// it. It also means the O(1) form is an optimization of release builds
    /// only; a debug build still pays the walk, on purpose.
    fn logical_position(&self) -> usize {
        debug_assert_eq!(
            self.logical_base,
            self.segments
                .iter()
                .take(self.segment_index)
                .map(|segment| segment.len())
                .sum::<usize>(),
            "SST cursor logical base drifted from the segments behind it"
        );
        self.logical_base.saturating_add(self.offset)
    }
}

/// What a shared-string walk does with the character data.
///
/// Every other part of one shared string — the header, the flags, the rich-text
/// run count, the `ExtRst` phonetic block and the `Continue` framing between
/// them — is walked once, by [`walk_one_shared_string`], for both modes.
trait SharedStringText: Sized {
    fn consume(cursor: &mut SstCursor<'_>, count: u16, high_byte: bool) -> Result<Self>;
}

impl SharedStringText for String {
    fn consume(cursor: &mut SstCursor<'_>, count: u16, high_byte: bool) -> Result<Self> {
        cursor.read_characters(count, high_byte)
    }
}

/// The measure-only mode: the character data is validated and stepped over, and
/// nothing is retained.
struct MeasuredText;

impl SharedStringText for MeasuredText {
    fn consume(cursor: &mut SstCursor<'_>, count: u16, high_byte: bool) -> Result<Self> {
        cursor.measure_characters(count, high_byte).map(|()| Self)
    }
}

/// What a formatting-run walk does with the runs it validates.
///
/// Every run is read and checked the same way whichever sink is used; the sink
/// decides only whether the runs inside the text are retained.
trait FormatRunSink: Sized {
    /// Reserves room for at most `count` runs, or refuses.
    fn with_capacity(count: u16) -> Result<Self>;

    /// Accepts one validated run that falls inside the string's text.
    fn push(&mut self, run: SharedStringFormatRun);
}

impl FormatRunSink for Vec<SharedStringFormatRun> {
    fn with_capacity(count: u16) -> Result<Self> {
        let mut runs = Self::new();
        runs.try_reserve_exact(count as usize).map_err(|error| {
            Error::InvalidData(format!(
                "cannot allocate shared string formatting runs: {error}"
            ))
        })?;
        Ok(runs)
    }

    fn push(&mut self, run: SharedStringFormatRun) {
        // The inherent `Vec::push`, spelled so that it cannot be read as the
        // trait method being defined here.
        Vec::push(self, run);
    }
}

/// The measure-only mode: the runs are validated in the same order and nothing
/// is retained, so no allocation is attempted per shared string.
///
/// Dropping the allocation drops the refusal that guarded it. Under an
/// allocator that cannot hand out `count * 4` bytes (at most 256 KiB), the
/// `Vec` sink refuses with `cannot allocate shared string formatting runs`
/// *before* reading a run; this sink walks the runs instead, so such a string
/// either succeeds or is refused by whichever run check it actually fails. That
/// is the one place where the two sinks are not interchangeable, and it needs
/// an exhausted allocator to reach: nothing about the input decides it.
struct MeasuredRuns;

impl FormatRunSink for MeasuredRuns {
    fn with_capacity(_count: u16) -> Result<Self> {
        Ok(Self)
    }

    fn push(&mut self, _run: SharedStringFormatRun) {}
}

/// Parses one shared string, returning its text.
fn parse_one_shared_string(
    cursor: &mut SstCursor<'_>,
    string_index: usize,
) -> Result<String, SharedStringScanError> {
    walk_one_shared_string::<String>(cursor, string_index)
}

fn walk_one_shared_string<T: SharedStringText>(
    cursor: &mut SstCursor<'_>,
    string_index: usize,
) -> Result<T, SharedStringScanError> {
    cursor
        .ensure_current(3, "shared string header")
        .map_err(SharedStringScanError::Biff)?;
    let character_count = cursor
        .read_u16_continued("shared string character count")
        .map_err(SharedStringScanError::Biff)?;
    let flags = cursor
        .read_u8_continued("shared string flags")
        .map_err(SharedStringScanError::Biff)?;
    let run_count = if flags & 0x08 != 0 {
        cursor
            .ensure_current(2, "shared string rich-text count")
            .map_err(SharedStringScanError::Biff)?;
        cursor
            .read_u16_continued("shared string rich-text count")
            .map_err(SharedStringScanError::Biff)?
    } else {
        0
    };
    let extension_length = if flags & 0x04 != 0 {
        cursor
            .ensure_current(4, "shared string extension length")
            .map_err(SharedStringScanError::Biff)?;
        let length = cursor
            .read_u32_continued("shared string extension length")
            .map_err(SharedStringScanError::Biff)?;
        if length > i32::MAX as u32 {
            return Err(SharedStringScanError::Invalid(format!(
                "shared string {string_index} has a negative extension length"
            )));
        }
        length as usize
    } else {
        0
    };

    let value = T::consume(cursor, character_count, flags & 0x01 != 0)
        .map_err(SharedStringScanError::Biff)?;
    // Both instantiations discard the runs, so both walk them without the
    // per-string `Vec`: the checks, their order and their messages are the
    // eager table's. What differs is the retention and, with it, the
    // reservation's own refusal -- see `MeasuredRuns`.
    cursor
        .walk_formatting_runs::<MeasuredRuns>(run_count, character_count, string_index)
        .map_err(SharedStringScanError::Biff)?;
    if flags & 0x04 != 0 {
        let extension = cursor
            .read_bytes(extension_length, "shared string ExtRst")
            .map_err(SharedStringScanError::Biff)?;
        parse_phonetic_string(&extension, character_count, string_index)
            .map_err(SharedStringScanError::Biff)?;
    }
    Ok(value)
}

pub(crate) fn scan_shared_string_records(
    records: &[RecordRef<'_>],
) -> Result<SharedStringSstScan, SharedStringScanError> {
    scan_shared_string_records_as::<MeasuredText>(records)
}

/// The SST offset scan, parameterised by what it does with each string's
/// character data.
///
/// Production only ever instantiates it with [`MeasuredText`]. The tests also
/// instantiate it with `String`, which is what the scan did before the measure
/// path existed, so the differential harness compares two instantiations of one
/// framing walk rather than two hand-written walks that could drift apart.
fn scan_shared_string_records_as<T: SharedStringText>(
    records: &[RecordRef<'_>],
) -> Result<SharedStringSstScan, SharedStringScanError> {
    if records.is_empty() {
        return Ok(SharedStringSstScan::empty());
    }
    if records[0].kind().get() != 0x00FC {
        return Err(SharedStringScanError::Biff(Error::UnexpectedRecordType {
            expected: 0x00FC,
            found: records[0].kind().get(),
        }));
    }
    if let Some(record) = records
        .iter()
        .skip(1)
        .find(|record| record.kind().get() != 0x003C)
    {
        return Err(SharedStringScanError::Biff(Error::UnexpectedRecordType {
            expected: 0x003C,
            found: record.kind().get(),
        }));
    }

    let mut segments = Vec::new();
    segments
        .try_reserve_exact(records.len())
        .map_err(|_| SharedStringScanError::Allocation {
            resource: "SST segment locator",
            requested: records.len(),
        })?;
    let mut payloads = Vec::new();
    payloads
        .try_reserve_exact(records.len())
        .map_err(|_| SharedStringScanError::Allocation {
            resource: "SST parser segments",
            requested: records.len(),
        })?;
    let mut logical_offset = 0usize;
    for record in records {
        let payload = record.payload();
        let next = logical_offset.checked_add(payload.len()).ok_or_else(|| {
            SharedStringScanError::Invalid("SST payload length overflow".to_owned())
        })?;
        segments.push(SharedStringSstSegment {
            source_offset: (record.offset() as u64).saturating_add(4),
            logical_offset,
            len: payload.len(),
        });
        payloads.push(payload);
        logical_offset = next;
    }

    let mut cursor = SstCursor::new(&payloads);
    cursor
        .ensure_current(8, "SST header")
        .map_err(SharedStringScanError::Biff)?;
    let total_count = cursor
        .read_u32_continued("SST total count")
        .map_err(SharedStringScanError::Biff)?;
    let unique_count = cursor
        .read_u32_continued("SST unique count")
        .map_err(SharedStringScanError::Biff)?;
    if total_count > i32::MAX as u32 || unique_count > i32::MAX as u32 {
        return Err(SharedStringScanError::Invalid(
            "SST counts must be non-negative signed integers".to_owned(),
        ));
    }
    if total_count < unique_count {
        return Err(SharedStringScanError::Invalid(
            "SST total count is smaller than its unique count".to_owned(),
        ));
    }
    let unique_count = usize::try_from(unique_count).map_err(|_| {
        SharedStringScanError::Invalid("SST unique count does not fit in usize".to_owned())
    })?;
    if unique_count > logical_offset.saturating_sub(8) / 3 {
        return Err(SharedStringScanError::Invalid(format!(
            "SST declares {unique_count} strings but its records are too short"
        )));
    }
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(unique_count)
        .map_err(|_| SharedStringScanError::Allocation {
            resource: "SST entry locator",
            requested: unique_count,
        })?;
    for string_index in 0..unique_count {
        let start = cursor.logical_position();
        // The scan retains offsets, never text, so production walks the string
        // rather than decoding it. The walk refuses the same inputs at the same
        // point in the sequence, with the same typed errors.
        walk_one_shared_string::<T>(&mut cursor, string_index)?;
        let end = cursor.logical_position();
        entries.push(SharedStringEntryLocation { start, end });
    }

    Ok(SharedStringSstScan { segments, entries })
}

pub(crate) fn decode_shared_string_entry(
    segments: &[&[u8]],
) -> Result<String, SharedStringScanError> {
    let mut cursor = SstCursor::new(segments);
    parse_one_shared_string(&mut cursor, 0)
}

#[cfg(test)]
mod sst_measure_tests {
    use super::*;
    use litchi_biff::{Encoder, Kind, Record as Frame, RecordRef, Records};
    use litchi_cfb::SharedOleFile;
    use litchi_core::OwnedSource;
    use std::path::{Path, PathBuf};
    use std::sync::Arc;

    /// The message `String::from_utf16` produces for malformed UTF-16, which is
    /// the refusal the open-time scan has always reported and must keep
    /// reporting. `char::decode_utf16` would say `unpaired surrogate found: …`
    /// instead; that difference is the whole reason the measure-only walk takes
    /// a cold path through `String::from_utf16` rather than reporting itself.
    const LONE_SURROGATE_MESSAGE: &str =
        "UTF-16 decoding error: invalid utf-16: lone surrogate found";

    fn frame(kind: u16, payload: &[u8]) -> Frame {
        let mut encoder = Encoder::new();
        encoder
            .push(Kind::from_wire(kind), payload)
            .expect("test frame fits the BIFF wire limit");
        Frame::open(encoder.finish()).expect("test frame is complete")
    }

    /// Frames `payloads` as one `SST` record followed by `Continue` records.
    fn sst_frames(payloads: &[Vec<u8>]) -> Vec<Frame> {
        payloads
            .iter()
            .enumerate()
            .map(|(index, payload)| frame(if index == 0 { 0x00FC } else { 0x003C }, payload))
            .collect()
    }

    fn sst_header(total: u32, unique: u32) -> Vec<u8> {
        let mut header = Vec::new();
        header.extend_from_slice(&total.to_le_bytes());
        header.extend_from_slice(&unique.to_le_bytes());
        header
    }

    /// Renders a scan outcome as text so that two instantiations can be compared
    /// exactly, including every error message.
    fn describe(result: &Result<SharedStringSstScan, SharedStringScanError>) -> String {
        match result {
            Ok(scan) => {
                let entries: Vec<(usize, usize)> = scan
                    .entries
                    .iter()
                    .map(|entry| (entry.start, entry.end))
                    .collect();
                format!("ok {entries:?}")
            },
            Err(SharedStringScanError::Biff(error)) => format!("biff {error}"),
            Err(SharedStringScanError::Invalid(message)) => format!("invalid {message}"),
            Err(SharedStringScanError::Allocation {
                resource,
                requested,
            }) => format!("allocation {resource} {requested}"),
        }
    }

    /// Runs both instantiations of the scan over the same records and asserts
    /// they agree exactly. Returns that shared outcome.
    fn scan_both_ways(records: &[RecordRef<'_>]) -> String {
        let measured = describe(&scan_shared_string_records_as::<MeasuredText>(records));
        let materialized = describe(&scan_shared_string_records_as::<String>(records));
        assert_eq!(
            measured, materialized,
            "the measure-only walk and the materializing walk disagree"
        );
        measured
    }

    fn scan_payloads(payloads: &[Vec<u8>]) -> String {
        let frames = sst_frames(payloads);
        let records: Vec<RecordRef<'_>> = frames.iter().map(Frame::as_ref).collect();
        scan_both_ways(&records)
    }

    /// One uncompressed string of `units`, entirely inside the `SST` record.
    fn one_wide_string(units: &[u16]) -> Vec<Vec<u8>> {
        let mut payload = sst_header(1, 1);
        payload.extend_from_slice(&(units.len() as u16).to_le_bytes());
        payload.push(0x01);
        for unit in units {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        vec![payload]
    }

    /// The same string split after `split_after` code units, so the remainder
    /// arrives in a `Continue` record that re-declares the encoding.
    fn one_wide_string_split(units: &[u16], split_after: usize) -> Vec<Vec<u8>> {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&(units.len() as u16).to_le_bytes());
        first.push(0x01);
        for unit in &units[..split_after] {
            first.extend_from_slice(&unit.to_le_bytes());
        }
        let mut second = vec![0x01u8];
        for unit in &units[split_after..] {
            second.extend_from_slice(&unit.to_le_bytes());
        }
        vec![first, second]
    }

    /// One string whose code units are delivered in the given runs. The first
    /// run lives in the `SST` record and every later run opens a `Continue`
    /// record with its own encoding flag, so a string can switch between
    /// compressed and uncompressed more than once.
    fn one_string_in_runs(runs: &[(bool, Vec<u16>)]) -> Vec<Vec<u8>> {
        let total: usize = runs.iter().map(|(_, units)| units.len()).sum();
        let mut payloads = Vec::new();
        for (index, (wide, units)) in runs.iter().enumerate() {
            let mut payload = if index == 0 {
                let mut head = sst_header(1, 1);
                head.extend_from_slice(&(total as u16).to_le_bytes());
                head
            } else {
                Vec::new()
            };
            payload.push(u8::from(*wide));
            for unit in units {
                if *wide {
                    payload.extend_from_slice(&unit.to_le_bytes());
                } else {
                    payload.push(u8::try_from(*unit).expect("a compressed unit fits in one byte"));
                }
            }
            payloads.push(payload);
        }
        payloads
    }

    /// An uncompressed head continued by a compressed tail, which is the
    /// boundary that can strand a high surrogate against bytes that cannot
    /// complete it.
    fn wide_head_compressed_tail(head: &[u16], tail: &[u8]) -> Vec<Vec<u8>> {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&((head.len() + tail.len()) as u16).to_le_bytes());
        first.push(0x01);
        for unit in head {
            first.extend_from_slice(&unit.to_le_bytes());
        }
        let mut second = vec![0x00u8];
        second.extend_from_slice(tail);
        vec![first, second]
    }

    #[test]
    fn a_lone_high_surrogate_keeps_the_from_utf16_message() {
        let frames = sst_frames(&one_wide_string(&[0xD800]));
        let records: Vec<RecordRef<'_>> = frames.iter().map(Frame::as_ref).collect();

        let error = scan_shared_string_records(&records)
            .expect_err("a shared string ending in a high surrogate is refused");
        let SharedStringScanError::Biff(Error::Encoding(message)) = &error else {
            panic!("expected a typed encoding refusal, got {error:?}");
        };
        assert_eq!(message, LONE_SURROGATE_MESSAGE);
        assert!(
            !message.contains("unpaired surrogate"),
            "the refusal must not be produced by char::decode_utf16: {message}"
        );
        assert_eq!(
            scan_both_ways(&records),
            format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}")
        );
    }

    #[test]
    fn every_unpaired_surrogate_shape_keeps_the_from_utf16_message() {
        let shapes: [(&str, Vec<u16>); 6] = [
            ("trailing high surrogate", vec![0x0041, 0xD800]),
            ("leading low surrogate", vec![0xDC00, 0x0041]),
            ("high surrogate before a plain unit", vec![0xD800, 0x0041]),
            ("two high surrogates", vec![0xD800, 0xD800]),
            ("low surrogate after a pair", vec![0xD800, 0xDC00, 0xDC00]),
            ("high surrogate at the maximum", vec![0xDBFF, 0x0041]),
        ];
        for (name, units) in shapes {
            let frames = sst_frames(&one_wide_string(&units));
            let records: Vec<RecordRef<'_>> = frames.iter().map(Frame::as_ref).collect();
            assert_eq!(
                scan_both_ways(&records),
                format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}"),
                "{name}"
            );
        }
    }

    #[test]
    fn a_surrogate_pair_split_across_a_continue_record_is_accepted() {
        // The high surrogate ends the SST record and the low surrogate opens the
        // Continue record, so the pairing state has to survive the boundary.
        let payloads = one_wide_string_split(&[0xD800, 0xDC00], 1);
        assert_eq!(scan_payloads(&payloads), "ok [(8, 16)]");
    }

    #[test]
    fn a_high_surrogate_stranded_by_a_compressed_continuation_is_refused() {
        // No compressed code unit can be a low surrogate, so the pending high
        // surrogate from the previous record can never be completed.
        let payloads = wide_head_compressed_tail(&[0x0041, 0xD800], b"BC");
        assert_eq!(
            scan_payloads(&payloads),
            format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}")
        );
    }

    #[test]
    fn a_high_surrogate_ending_a_string_before_a_compressed_continuation_is_refused() {
        // The pending high surrogate is the last unit of the wide run and the
        // continuation carries no unit at all, which is the end-of-input case.
        let payloads = wide_head_compressed_tail(&[0xD800], b"");
        assert_eq!(
            scan_payloads(&payloads),
            format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}")
        );
    }

    #[test]
    fn a_pending_high_surrogate_cannot_be_completed_across_an_intervening_run() {
        // `String::from_utf16` refuses `[D800, 0041, DC00]`: the high surrogate
        // is stranded by the plain unit and the low surrogate is then stray.
        // Both shapes need three runs, because only then does a run that cannot
        // complete the pair sit between the two halves.
        for middle in [(false, vec![0x0041u16]), (true, vec![0x0041u16])] {
            let runs = [
                (true, vec![0xD800u16]),
                middle.clone(),
                (true, vec![0xDC00u16]),
            ];
            assert_eq!(
                scan_payloads(&one_string_in_runs(&runs)),
                format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}"),
                "a pair must not form across {middle:?}"
            );
        }
    }

    #[test]
    fn a_pair_still_forms_across_an_empty_intervening_run() {
        // An empty run carries no code unit, so it cannot strand anything and
        // the pair is still complete. This is the boundary case that stops the
        // rule above from being stated as "any intervening run".
        let runs = [
            (true, vec![0xD800u16]),
            (false, Vec::new()),
            (true, vec![0xDC00u16]),
        ];
        assert_eq!(scan_payloads(&one_string_in_runs(&runs)), "ok [(8, 17)]");
    }

    #[test]
    fn measuring_and_materializing_agree_over_every_three_run_delivery() {
        const UNITS: [u16; 8] = [
            0x0041, 0x00FF, 0xD7FF, 0xD800, 0xDBFF, 0xDC00, 0xDFFF, 0xE000,
        ];
        let mut compared = 0usize;
        for first in UNITS {
            for middle in UNITS {
                for last in UNITS {
                    let mut deliveries = vec![vec![
                        (true, vec![first]),
                        (true, vec![middle]),
                        (true, vec![last]),
                    ]];
                    if middle <= 0x00FF {
                        deliveries.push(vec![
                            (true, vec![first]),
                            (false, vec![middle]),
                            (true, vec![last]),
                        ]);
                    }
                    for runs in deliveries {
                        scan_payloads(&one_string_in_runs(&runs));
                        compared += 1;
                    }
                }
            }
        }
        assert_eq!(compared, 640);
    }

    #[test]
    fn a_pair_survives_a_long_plain_run_before_it() {
        // Long enough that the chunk-level surrogate pre-scan takes its fast
        // exit on the leading units and still reports the trailing defect.
        let mut units = vec![0x0041u16; 512];
        units.push(0xD800);
        let payloads = one_wide_string(&units);
        assert_eq!(
            scan_payloads(&payloads),
            format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}")
        );

        units.push(0xDC00);
        let payloads = one_wide_string(&units);
        assert_eq!(scan_payloads(&payloads), "ok [(8, 1039)]");
    }

    #[test]
    fn measuring_and_materializing_agree_over_every_short_code_unit_sequence() {
        // Both surrogate halves, both extremes of each half, a plain unit, a
        // unit whose high byte neighbours the surrogate block, and the
        // replacement character.
        const UNITS: [u16; 8] = [
            0x0041, 0x00FF, 0xD7FF, 0xD800, 0xDBFF, 0xDC00, 0xDFFF, 0xE000,
        ];
        let mut sequences: Vec<Vec<u16>> = UNITS.iter().map(|unit| vec![*unit]).collect();
        for first in UNITS {
            for second in UNITS {
                sequences.push(vec![first, second]);
                for third in UNITS {
                    sequences.push(vec![first, second, third]);
                }
            }
        }
        let mut compared = 0usize;
        for units in &sequences {
            scan_payloads(&one_wide_string(units));
            compared += 1;
            for split_after in 0..units.len() {
                scan_payloads(&one_wide_string_split(units, split_after));
                compared += 1;
            }
            scan_payloads(&wide_head_compressed_tail(units, b"z"));
            compared += 1;
        }
        assert_eq!(compared, 2_840);
    }

    #[test]
    fn a_framing_defect_after_a_malformed_unit_still_wins() {
        // The first record already holds an unpaired high surrogate; the
        // continuation then ends inside a code unit. The framing refusal is
        // reported, because `String::from_utf16` only ever ran after the walk.
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&3u16.to_le_bytes());
        first.push(0x01);
        first.extend_from_slice(&0xD800u16.to_le_bytes());
        first.extend_from_slice(&0x0041u16.to_le_bytes());
        let second = vec![0x01u8, 0x00];

        assert_eq!(
            scan_payloads(&[first, second]),
            "biff Invalid data: a UTF-16 shared string is split inside a code unit"
        );
    }

    #[test]
    fn invalid_continuation_flags_after_a_malformed_unit_still_win() {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&3u16.to_le_bytes());
        first.push(0x01);
        first.extend_from_slice(&0xD800u16.to_le_bytes());
        first.extend_from_slice(&0x0041u16.to_le_bytes());
        let second = vec![0x02u8, 0x41, 0x00];

        assert_eq!(
            scan_payloads(&[first, second]),
            "biff Invalid data: invalid shared string continuation flags 0x02"
        );
    }

    #[test]
    fn a_malformed_unit_is_reported_before_a_bad_formatting_run() {
        // `read_characters` ran before `read_formatting_runs` and still does, so
        // the cold path has to report from the same position in the sequence.
        let mut payload = sst_header(1, 1);
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.push(0x09); // rich text plus the high-byte flag
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0xD800u16.to_le_bytes());
        payload.extend_from_slice(&9u16.to_le_bytes()); // a run past the text
        payload.extend_from_slice(&0u16.to_le_bytes());

        assert_eq!(
            scan_payloads(&[payload]),
            format!("biff Encoding error: {LONE_SURROGATE_MESSAGE}")
        );
    }

    #[test]
    fn a_bad_formatting_run_is_still_reported_when_the_text_is_well_formed() {
        let mut payload = sst_header(1, 1);
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.push(0x09);
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0x0041u16.to_le_bytes());
        payload.extend_from_slice(&9u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());

        assert_eq!(
            scan_payloads(&[payload]),
            "biff Invalid data: shared string 0 has a formatting run past its text"
        );
    }

    #[test]
    fn a_phonetic_block_is_still_walked_and_still_validated() {
        // ExtRst is parsed on both paths, so its refusals keep their position
        // after the character data and after the formatting runs.
        let mut payload = sst_header(1, 1);
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.push(0x05); // ExtRst plus the high-byte flag
        payload.extend_from_slice(&13u32.to_le_bytes()); // shorter than the 14-byte head
        payload.extend_from_slice(&0x0041u16.to_le_bytes());
        payload.extend_from_slice(&[0; 13]);

        assert_eq!(
            scan_payloads(&[payload]),
            "biff Invalid length: expected 14, found 13"
        );
    }

    #[test]
    fn an_empty_string_and_a_compressed_string_still_measure_the_same_span() {
        let mut payload = sst_header(2, 2);
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.push(0x00);
        payload.extend_from_slice(&2u16.to_le_bytes());
        payload.push(0x00);
        payload.extend_from_slice(b"hi");

        assert_eq!(scan_payloads(&[payload]), "ok [(8, 11), (11, 16)]");
    }

    /// The fixed-width reads take a resident fast path when the field lies
    /// inside one record, and fall through to `read_exact` when it does not.
    /// A formatting run is the only shared-string field that can straddle a
    /// `Continue` boundary — `ensure_current` keeps the header, the rich-text
    /// count and the extension length inside one record — so it is the only
    /// place where the fall-through, and `advance_segment`'s maintenance of the
    /// cursor's logical base, can be exercised from the scan.
    ///
    /// The second string exists to prove the base is still right *after* the
    /// crossing: its `start` has to be the first string's `end`, and both have
    /// to be offsets into the concatenation of the two records.
    #[test]
    fn a_formatting_run_split_across_a_continue_record_still_indexes_exactly() {
        let mut first = sst_header(2, 2);
        first.extend_from_slice(&2u16.to_le_bytes());
        first.push(0x08); // rich text, compressed characters
        first.extend_from_slice(&1u16.to_le_bytes());
        first.extend_from_slice(b"AB");
        first.push(1); // first byte of the run's character_index
        assert_eq!(first.len(), 16);

        let mut second = vec![0, 9, 0]; // its second byte, then font_index 9
        second.extend_from_slice(&2u16.to_le_bytes());
        second.push(0x00);
        second.extend_from_slice(b"CD");
        assert_eq!(second.len(), 8);

        assert_eq!(scan_payloads(&[first, second]), "ok [(8, 19), (19, 24)]");
    }

    fn test_data_root() -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    fn collect_xls_fixtures(root: &Path, found: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(root) else {
            return;
        };
        let mut children: Vec<PathBuf> = entries
            .filter_map(|entry| entry.ok())
            .map(|entry| entry.path())
            .collect();
        children.sort();
        for child in children {
            if child.is_dir() {
                collect_xls_fixtures(&child, found);
            } else if child
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xls") || extension.eq_ignore_ascii_case("xlt")
                })
            {
                found.push(child);
            }
        }
    }

    /// Returns the byte range of the `SST` record and its `Continue` run inside
    /// a workbook globals substream, framed by hand so that the harness does not
    /// share code with the scan it is checking.
    fn locate_sst(stream: &[u8]) -> Option<(usize, usize)> {
        let mut offset = 0usize;
        let mut start = None;
        while offset + 4 <= stream.len() {
            let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
            let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
            let end = offset.checked_add(4)?.checked_add(length)?;
            if end > stream.len() {
                return None;
            }
            match (kind, start) {
                (0x00FC, None) => start = Some(offset),
                (0x003C, Some(_)) => {},
                (_, Some(begin)) => return Some((begin, offset)),
                // The globals substream ends at its first EOF; an SST never
                // appears in a worksheet substream.
                (0x000A, None) => return None,
                (_, None) => {},
            }
            offset = end;
        }
        start.map(|begin| (begin, stream.len()))
    }

    fn workbook_stream(bytes: Vec<u8>) -> Result<Vec<u8>, String> {
        let file = SharedOleFile::open(Arc::new(OwnedSource::new(bytes)))
            .map_err(|error| format!("container: {error}"))?;
        let mut last = String::from("no workbook stream name matched");
        for name in ["Workbook", "Book", "WORKBOOK", "BOOK"] {
            match file.open_stream(&[name]) {
                Ok(stream) => return Ok(stream),
                Err(error) => last = format!("{name}: {error}"),
            }
        }
        Err(last)
    }

    /// Change 0574's falsification criterion for this change: the measure-only
    /// walk has to reproduce byte-identical `entries` over every fixture that
    /// carries an SST.
    #[test]
    fn every_sst_fixture_indexes_identically_both_ways() {
        let root = test_data_root();
        let mut fixtures = Vec::new();
        collect_xls_fixtures(&root, &mut fixtures);
        assert!(
            fixtures.len() > 100,
            "the XLS corpus should be present, found {}",
            fixtures.len()
        );

        let mut with_sst = 0usize;
        let mut indexed = 0usize;
        let mut refused = 0usize;
        let mut entries_compared = 0usize;
        let mut no_container = Vec::new();
        let mut no_sst = Vec::new();
        let mut refusals = Vec::new();

        for fixture in &fixtures {
            let relative = fixture
                .strip_prefix(&root)
                .unwrap_or(fixture)
                .to_string_lossy()
                .into_owned();
            let Ok(bytes) = std::fs::read(fixture) else {
                no_container.push(format!("{relative}: unreadable"));
                continue;
            };
            let stream = match workbook_stream(bytes) {
                Ok(stream) => stream,
                Err(reason) => {
                    no_container.push(format!("{relative}: {reason}"));
                    continue;
                },
            };
            let Some((begin, end)) = locate_sst(&stream) else {
                no_sst.push(relative);
                continue;
            };
            with_sst += 1;
            let span = &stream[begin..end];
            let Ok(records) = Records::new(span).collect::<Result<Vec<RecordRef<'_>>, _>>() else {
                no_container.push(format!("{relative}: SST span does not reframe"));
                continue;
            };

            let measured = scan_shared_string_records_as::<MeasuredText>(&records);
            let materialized = scan_shared_string_records_as::<String>(&records);
            match (&measured, &materialized) {
                (Ok(left), Ok(right)) => {
                    let left: Vec<(usize, usize)> =
                        left.entries.iter().map(|e| (e.start, e.end)).collect();
                    let right: Vec<(usize, usize)> =
                        right.entries.iter().map(|e| (e.start, e.end)).collect();
                    assert_eq!(left, right, "{relative} indexes differently");
                    entries_compared += left.len();
                    indexed += 1;
                },
                _ => {
                    assert_eq!(
                        describe(&measured),
                        describe(&materialized),
                        "{relative} is refused differently"
                    );
                    refused += 1;
                    refusals.push(format!("{relative}: {}", describe(&measured)));
                },
            }
        }

        println!(
            "sst-differential: fixtures={} with_sst={with_sst} indexed={indexed} refused={refused} entries={entries_compared}",
            fixtures.len()
        );
        for line in &no_container {
            println!("sst-differential: skipped {line}");
        }
        println!("sst-differential: without an SST = {}", no_sst.len());
        for line in &refusals {
            println!("sst-differential: refused {line}");
        }
        assert!(
            indexed >= 100,
            "expected the corpus to index at least 100 SSTs, indexed {indexed}"
        );
    }

    /// FNV-1a over a byte string. Not cryptographic; it only has to notice a
    /// moved offset, and the per-fixture lines the test prints say which one.
    fn fnv1a(bytes: &[u8]) -> u64 {
        let mut hash = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
        }
        hash
    }

    /// The SST index over the whole corpus, pinned to the value it had before
    /// change 0595 rewrote the walk.
    ///
    /// `every_sst_fixture_indexes_identically_both_ways` proves the two
    /// instantiations of the walk agree with *each other*; it cannot see a
    /// change that moves both. This pins the index itself: every segment
    /// locator and every entry boundary of every fixture, digested per fixture
    /// and then over the corpus. A failure prints the per-fixture digests, so
    /// the fixture that moved is named.
    ///
    /// `CORPUS_DIGEST` is a property of `test-data`, so adding or removing an
    /// XLS fixture changes it legitimately; rerun with `--nocapture` and read
    /// the per-fixture lines before accepting a new value.
    #[test]
    fn the_sst_index_over_the_corpus_is_pinned() {
        const CORPUS_DIGEST: u64 = 0x9cb1_4f5d_aa02_eebc;

        let root = test_data_root();
        let mut fixtures = Vec::new();
        collect_xls_fixtures(&root, &mut fixtures);
        assert!(
            fixtures.len() > 100,
            "the XLS corpus should be present, found {}",
            fixtures.len()
        );

        let mut lines = Vec::new();
        for fixture in &fixtures {
            let relative = fixture
                .strip_prefix(&root)
                .unwrap_or(fixture)
                .to_string_lossy()
                .into_owned();
            let Ok(bytes) = std::fs::read(fixture) else {
                continue;
            };
            let Ok(stream) = workbook_stream(bytes) else {
                continue;
            };
            let Some((begin, end)) = locate_sst(&stream) else {
                continue;
            };
            let Ok(records) =
                Records::new(&stream[begin..end]).collect::<Result<Vec<RecordRef<'_>>, _>>()
            else {
                continue;
            };
            let result = scan_shared_string_records(&records);
            let detail = match &result {
                Ok(scan) => {
                    let mut detail = String::new();
                    for segment in &scan.segments {
                        detail.push_str(&format!(
                            "s {} {} {}\n",
                            segment.source_offset, segment.logical_offset, segment.len
                        ));
                    }
                    for entry in &scan.entries {
                        detail.push_str(&format!("e {} {}\n", entry.start, entry.end));
                    }
                    format!(
                        "segments={} entries={} body={:016x}",
                        scan.segments.len(),
                        scan.entries.len(),
                        fnv1a(detail.as_bytes())
                    )
                },
                Err(_) => format!("refused={}", describe(&result)),
            };
            lines.push(format!("{relative}: {detail}"));
        }

        let corpus = lines.join("\n");
        let digest = fnv1a(corpus.as_bytes());
        println!("sst-index: fixtures={} digest={digest:#018x}", lines.len());
        for line in &lines {
            println!("sst-index: {line}");
        }
        assert_eq!(
            digest, CORPUS_DIGEST,
            "the SST index over the corpus moved; rerun with --nocapture and compare the per-fixture lines"
        );
    }
}

#[cfg(test)]
mod shared_string_tests {
    use super::*;
    use litchi_biff::{Encoder, Kind, Record as Frame, RecordRef, Records};

    fn record(record_type: u16, data: Vec<u8>) -> Frame {
        let mut encoder = Encoder::new();
        encoder
            .push(Kind::from_wire(record_type), &data)
            .expect("test frame fits the BIFF wire limit");
        Frame::open(encoder.finish()).expect("test frame is complete")
    }

    fn record_refs(records: &[Frame]) -> Vec<RecordRef<'_>> {
        records.iter().map(|record| record.as_ref()).collect()
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

        let encoding = Encoding::from_codepage(1251).unwrap();
        let table = SharedStringTable::parse(&data, &encoding).unwrap();

        assert_eq!(table.total_count, 3);
        // BIFF8 compressed Unicode supplies an implicit zero high byte; it is
        // not encoded in the workbook CODEPAGE.
        assert_eq!(table.strings, ["AÀ", "漢"]);
        assert_eq!(table.properties, [None, None]);
    }

    #[test]
    fn codepage_construction_and_utf16_decoding_are_strict() {
        assert!(Encoding::from_codepage(437).is_err());
        assert!(Encoding::from_codepage(1201).is_err());
        assert!(Encoding::Utf16Le.decode(b"A").is_err());
    }

    #[test]
    fn ignores_reserved_shared_string_flags() {
        let mut data = sst_header(1, 1);
        data.extend_from_slice(&[1, 0, 0xF2, b'A']);

        let table = SharedStringTable::parse(&data, &Encoding::Utf16Le).unwrap();

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

        let table = SharedStringTable::parse(&data, &Encoding::Utf16Le).unwrap();
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

        let table = SharedStringTable::parse(&data, &Encoding::Utf16Le).unwrap();
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
        let refs = record_refs(&records);
        let table = SharedStringTable::parse_from_records(&refs, &Encoding::Utf16Le).unwrap();

        assert_eq!(table.strings, ["AB漢字"]);
    }

    #[test]
    fn chained_empty_continuations_return_an_error_instead_of_panicking() {
        let mut first = sst_header(1, 1);
        first.extend_from_slice(&2u16.to_le_bytes());
        first.push(0);
        first.push(b'A');
        let records = [
            record(0x00FC, first),
            record(0x003C, Vec::new()),
            record(0x003C, Vec::new()),
        ];
        let refs = record_refs(&records);

        assert!(matches!(
            SharedStringTable::parse_from_records(&refs, &Encoding::Utf16Le),
            Err(Error::UnexpectedEndOfStream(_))
        ));
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
        let refs = record_refs(&records);
        let table = SharedStringTable::parse_from_records(&refs, &Encoding::Utf16Le).unwrap();

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
        let bad_refs = record_refs(&bad_flags);
        assert!(SharedStringTable::parse_from_records(&bad_refs, &Encoding::Utf16Le).is_err());

        let mut truncated = sst_header(1, 1);
        truncated.extend_from_slice(&1u16.to_le_bytes());
        truncated.push(0x04);
        truncated.extend_from_slice(&100u32.to_le_bytes());
        truncated.push(b'A');
        assert!(SharedStringTable::parse(&truncated, &Encoding::Utf16Le).is_err());
    }

    #[test]
    fn reads_writer_generated_multirecord_sst() {
        let expected = vec!["a".repeat(9000), "漢".repeat(5000)];
        let mut bytes = Vec::new();
        crate::writer::biff::write_sst(&mut bytes, &expected, 2).unwrap();
        let records: Vec<RecordRef<'_>> = Records::new(&bytes).collect::<Result<_, _>>().unwrap();

        let table = SharedStringTable::parse_from_records(&records, &Encoding::Utf16Le).unwrap();

        assert_eq!(table.strings, expected);
    }

    #[test]
    fn maps_truncated_shared_frame_errors_at_the_xls_boundary() {
        let truncated = [0xFC, 0x00, 0x04, 0x00, 0x01, 0x02];
        let error = Records::new(&truncated)
            .next()
            .expect("the malformed input yields one framing error")
            .expect_err("the payload is shorter than its declared length");
        let error = Error::from(error);

        assert!(matches!(
            error,
            Error::InvalidRecord {
                record_type: 0x00FC,
                ..
            }
        ));
    }
}

/// XF (Extended Format) record - cell formatting
#[derive(Debug, Clone)]
#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
pub struct ExtendedFormat {
    pub font_index: u16,
    pub format_index: u16,
}

#[allow(dead_code, reason = "retained as a BIFF compatibility building block")]
impl ExtendedFormat {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 4 {
            return Err(Error::InvalidLength {
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
        metadata: crate::formula_metadata::Metadata,
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
    /// String value stored in the immediately following BIFF String record.
    StringPending,
    String(String),
    Bool(bool),
    Error(u8),
    Empty,
}

impl CellRecord {
    #[must_use]
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

    #[must_use]
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

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(record_type: u16, data: &[u8], encoding: &Encoding) -> Result<Self> {
        match record_type {
            0x0201 => Self::parse_blank(data),           // Blank
            0x0203 => Self::parse_number(data),          // Number
            0x0204 => Self::parse_label(data, encoding), // Label
            0x0205 => Self::parse_bool_err(data),        // BoolErr
            0x027E => Self::parse_rk(data),              // RK
            0x00FD => Self::parse_label_sst(data),       // LabelSst
            0x0006 => Self::parse_formula(data),         // Formula
            _ => Err(Error::InvalidRecord {
                record_type,
                message: "Unknown cell record type".to_string(),
            }),
        }
    }

    pub(crate) fn parse_mul_rk(data: &[u8]) -> Result<Vec<Self>> {
        let (_, _, count) = Self::packed_cell_range(data, 6, "MulRk")?;
        let mut cells = Vec::with_capacity(count);
        Self::visit_mul_rk(data, |cell| cells.push(cell))?;
        Ok(cells)
    }

    pub(crate) fn parse_mul_blank(data: &[u8]) -> Result<Vec<Self>> {
        let (_, _, count) = Self::packed_cell_range(data, 2, "MulBlank")?;
        let mut cells = Vec::with_capacity(count);
        Self::visit_mul_blank(data, |cell| cells.push(cell))?;
        Ok(cells)
    }

    pub(crate) fn visit_mul_rk(data: &[u8], mut visitor: impl FnMut(Self)) -> Result<()> {
        let (row, first_col, count) = Self::packed_cell_range(data, 6, "MulRk")?;
        for index in 0..count {
            let offset = 4 + index * 6;
            visitor(Self::Rk {
                row,
                col: first_col + utils::truncate_usize_to_u16(index),
                xf_index: binary::read_u16_le_at(data, offset)?,
                value: utils::rk_to_f64(binary::read_u32_le_at(data, offset + 2)?),
            });
        }
        Ok(())
    }

    pub(crate) fn visit_mul_blank(data: &[u8], mut visitor: impl FnMut(Self)) -> Result<()> {
        let (row, first_col, count) = Self::packed_cell_range(data, 2, "MulBlank")?;
        for index in 0..count {
            visitor(Self::Blank {
                row,
                col: first_col + utils::truncate_usize_to_u16(index),
                xf_index: binary::read_u16_le_at(data, 4 + index * 2)?,
            });
        }
        Ok(())
    }

    fn packed_cell_range(
        data: &[u8],
        item_size: usize,
        record_name: &str,
    ) -> Result<(u16, u16, usize)> {
        let Some(items_size) = data.len().checked_sub(6) else {
            return Err(Error::InvalidLength {
                expected: 6 + item_size * 2,
                found: data.len(),
            });
        };
        if items_size % item_size != 0 {
            return Err(Error::InvalidData(format!(
                "{record_name} payload does not contain whole packed cells"
            )));
        }
        let count = items_size / item_size;
        if !(2..=256).contains(&count) {
            return Err(Error::InvalidData(format!(
                "{record_name} contains {count} cells; expected 2 through 256"
            )));
        }

        let row = binary::read_u16_le_at(data, 0)?;
        let first_col = binary::read_u16_le_at(data, 2)?;
        let last_col = binary::read_u16_le_at(data, data.len() - 2)?;
        let expected_last = first_col
            .checked_add(utils::truncate_usize_to_u16(count - 1))
            .ok_or_else(|| Error::InvalidData(format!("{record_name} column overflow")))?;
        if first_col > 254 || last_col != expected_last || last_col > 255 {
            return Err(Error::InvalidData(format!(
                "{record_name} column range {first_col}..={last_col} does not match {count} cells"
            )));
        }
        Ok((row, first_col, count))
    }

    fn parse_blank(data: &[u8]) -> Result<Self> {
        if data.len() < 6 {
            return Err(Error::InvalidLength {
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

    fn parse_number(data: &[u8]) -> Result<Self> {
        if data.len() < 14 {
            return Err(Error::InvalidLength {
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

    fn parse_label(data: &[u8], encoding: &Encoding) -> Result<Self> {
        if data.len() < 8 {
            return Err(Error::InvalidLength {
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

    fn parse_bool_err(data: &[u8]) -> Result<Self> {
        if data.len() < 8 {
            return Err(Error::InvalidLength {
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

    fn parse_rk(data: &[u8]) -> Result<Self> {
        if data.len() < 10 {
            return Err(Error::InvalidLength {
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

    fn parse_label_sst(data: &[u8]) -> Result<Self> {
        if data.len() < 10 {
            return Err(Error::InvalidLength {
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

    fn parse_formula(data: &[u8]) -> Result<Self> {
        let parsed = crate::formula_metadata::parse_record(data)?;
        Ok(CellRecord::Formula {
            row: parsed.row,
            col: parsed.col,
            xf_index: parsed.xf_index,
            value: parsed.value,
            metadata: parsed.metadata,
            formula: parsed.formula,
        })
    }

    pub(crate) fn parse_formula_preserving_defect(
        data: &[u8],
    ) -> Result<(Self, Option<crate::formula_metadata::FlagDefect>)> {
        let (parsed, defect) = crate::formula_metadata::parse_record_preserving(data)?;
        Ok((
            CellRecord::Formula {
                row: parsed.row,
                col: parsed.col,
                xf_index: parsed.xf_index,
                value: parsed.value,
                metadata: parsed.metadata,
                formula: parsed.formula,
            },
            defect,
        ))
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

    #[test]
    fn formula_uses_declared_token_length() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&4.5f64.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&[0x1E, 0x2A, 0x00]);

        let formula = CellRecord::parse(0x0006, &data, &Encoding::Utf16Le).unwrap();

        assert!(matches!(
            formula,
            CellRecord::Formula {
                value: FormulaValue::Number(4.5),
                formula,
                ..
            } if formula == [0x1E, 0x2A, 0x00]
        ));
    }
}
