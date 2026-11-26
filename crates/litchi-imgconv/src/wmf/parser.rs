// WMF file parser
//
// Parses Windows Metafile records and extracts relevant information
//
// ## Performance Optimizations
//
// This parser is optimized for minimal memory allocations:
//
// 1. **Zero-copy data storage**: Uses `Bytes` with reference counting instead of `Vec<u8>`
//    - The input data is copied once into a `Bytes` buffer
//    - All record params are zero-copy slices of this buffer via `Bytes::slice()`
//    - Eliminates N allocations where N = number of records
//
// 2. **Pre-allocated records vector**: Estimates capacity based on file size
//    - Reduces reallocation overhead during parsing
//    - Typical WMF files have 20-50 bytes per record on average
//
// 3. **Manual byte parsing**: Avoids zerocopy alignment issues
//    - Direct byte access for little-endian values
//    - No intermediate allocations for header parsing
//
// These optimizations significantly reduce calls to `_platform_memmove`,
// `alloc::raw_vec::RawVec::grow_one`, and `szone_malloc_should_clear`.

use bytes::Bytes;
use litchi_core::error::{Error, Result};

/// WMF file type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WmfFileType {
    /// Memory metafile
    Memory = 1,
    /// Disk metafile
    Disk = 2,
}

/// WMF header (Placeable Metafile Header)
///
/// This is an optional header that may precede the standard WMF header
#[derive(Debug, Clone)]
pub struct WmfPlaceableHeader {
    /// Key (should be 0x9AC6CDD7)
    pub key: u32,
    /// Left coordinate
    pub left: i16,
    /// Top coordinate
    pub top: i16,
    /// Right coordinate
    pub right: i16,
    /// Bottom coordinate
    pub bottom: i16,
    /// Units per inch
    pub inch: u16,
    /// Checksum
    pub checksum: u16,
}

impl WmfPlaceableHeader {
    const PLACEABLE_KEY: u32 = 0x9AC6CDD7;

    /// Check if data starts with a placeable header
    pub fn is_placeable(data: &[u8]) -> bool {
        if data.len() < 4 {
            return false;
        }
        let key = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        key == Self::PLACEABLE_KEY
    }

    /// Parse placeable header from data
    pub fn parse(data: &[u8]) -> Result<Self> {
        let (header, integrity) = Self::parse_compatible(data)?;
        if let Some(field) = integrity {
            return Err(Error::ParseError(field.to_string()));
        }
        Ok(header)
    }

    /// Parses the placeable header while reporting, rather than rejecting, an
    /// inconsistent integrity field.
    ///
    /// The key and the fixed header length identify the header; the reserved
    /// field, the units-per-inch field and the checksum are consistency hints
    /// that playback does not depend on.
    pub fn parse_compatible(data: &[u8]) -> Result<(Self, Option<&'static str>)> {
        if data.len() < 22 {
            return Err(Error::ParseError("WMF placeable header too short".into()));
        }

        // Parse header manually to avoid zerocopy alignment issues
        let key = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);

        if key != Self::PLACEABLE_KEY {
            return Err(Error::ParseError(format!(
                "Invalid WMF placeable key: 0x{:08X}",
                key
            )));
        }

        let left = i16::from_le_bytes([data[6], data[7]]);
        let top = i16::from_le_bytes([data[8], data[9]]);
        let right = i16::from_le_bytes([data[10], data[11]]);
        let bottom = i16::from_le_bytes([data[12], data[13]]);
        let inch = u16::from_le_bytes([data[14], data[15]]);
        let checksum = u16::from_le_bytes([data[20], data[21]]);

        let reserved = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
        let calculated_checksum = data[..20]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        let integrity = if reserved != 0 {
            Some("WMF placeable reserved field is nonzero")
        } else if inch == 0 {
            Some("WMF placeable units-per-inch must be nonzero")
        } else if checksum != calculated_checksum {
            Some("Invalid WMF placeable checksum")
        } else {
            None
        };

        Ok((
            Self {
                key,
                left,
                top,
                right,
                bottom,
                inch,
                checksum,
            },
            integrity,
        ))
    }

    /// Get width
    pub fn width(&self) -> i16 {
        self.right.saturating_sub(self.left)
    }

    /// Get height
    pub fn height(&self) -> i16 {
        self.bottom.saturating_sub(self.top)
    }
}

/// WMF standard header
#[derive(Debug, Clone)]
pub struct WmfHeader {
    /// File type (1 = memory, 2 = disk)
    pub file_type: u16,
    /// Header size in words (always 9)
    pub header_size: u16,
    /// Windows version
    pub version: u16,
    /// Size of file in words
    pub file_size: u32,
    /// Number of objects
    pub num_objects: u16,
    /// Size of largest record in words
    pub max_record: u32,
    /// Not used (always 0)
    pub num_params: u16,
}

impl WmfHeader {
    /// Parse WMF standard header
    pub fn parse(data: &[u8]) -> Result<Self> {
        let (header, deviation) = Self::parse_compatible(data)?;
        if let Some(deviation) = deviation {
            return Err(Error::ParseError(deviation.message()));
        }
        Ok(header)
    }

    /// Parses the standard header while reporting, rather than rejecting, a
    /// declared value that playback does not depend on.
    ///
    /// `mtSize` locates the record stream end for recording tools and
    /// `mtMaxRecord` only sizes their buffers; playback follows the record
    /// lengths themselves. The fixed header fields that determine where the
    /// records start stay mandatory.
    pub fn parse_compatible(data: &[u8]) -> Result<(Self, Option<WmfHeaderDeviation>)> {
        if data.len() < 18 {
            return Err(Error::ParseError("WMF header too short".into()));
        }

        // Parse header manually to avoid zerocopy alignment issues
        let file_type = u16::from_le_bytes([data[0], data[1]]);
        let header_size = u16::from_le_bytes([data[2], data[3]]);
        let version = u16::from_le_bytes([data[4], data[5]]);
        let file_size = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
        let num_objects = u16::from_le_bytes([data[10], data[11]]);
        let max_record = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
        let num_params = u16::from_le_bytes([data[16], data[17]]);

        if !matches!(file_type, 1 | 2) {
            return Err(Error::ParseError(format!(
                "Invalid WMF file type: {}",
                file_type
            )));
        }
        if header_size != 9 {
            return Err(Error::ParseError(format!(
                "Invalid WMF header size: {} words",
                header_size
            )));
        }
        if !matches!(version, 0x0100 | 0x0300) {
            return Err(Error::ParseError(format!(
                "Unsupported WMF version: 0x{version:04X}"
            )));
        }

        let deviation = if file_size < 12 {
            Some(WmfHeaderDeviation::FileSize(file_size))
        } else if max_record < 3 || max_record > file_size.saturating_sub(u32::from(header_size)) {
            Some(WmfHeaderDeviation::MaxRecord(max_record))
        } else if u32::from(num_objects) > file_size.saturating_sub(u32::from(header_size)) / 3 {
            Some(WmfHeaderDeviation::ObjectCount(num_objects))
        } else {
            None
        };
        Ok((
            Self {
                file_type,
                header_size,
                version,
                file_size,
                num_objects,
                max_record,
                num_params,
            },
            deviation,
        ))
    }
}

/// A declared standard-header value that record playback does not depend on.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WmfHeaderDeviation {
    /// `mtSize` cannot hold the fixed header and a terminating record.
    FileSize(u32),
    /// `mtMaxRecord` is not a record length inside the declared file size.
    MaxRecord(u32),
    /// `mtNoObjects` cannot fit in the declared file size.
    ObjectCount(u16),
}

impl WmfHeaderDeviation {
    /// Human-readable description with the strict-validation wording.
    #[must_use]
    pub fn message(self) -> String {
        match self {
            Self::FileSize(words) => format!("Invalid WMF file size: {words} words"),
            Self::MaxRecord(words) => {
                format!("Invalid WMF maximum record size: {words} words")
            },
            Self::ObjectCount(objects) => {
                format!("WMF object count {objects} cannot fit in the declared file size")
            },
        }
    }
}

/// WMF record
#[derive(Debug, Clone)]
pub struct WmfRecord {
    /// Record size in words (including size and function)
    pub size: u32,
    /// Record function
    pub function: u16,
    /// Record parameters (zero-copy slice of the original data)
    pub params: Bytes,
}

impl WmfRecord {
    /// Parse a WMF record
    ///
    /// # Arguments
    /// * `data` - Zero-copy bytes buffer containing the WMF data
    /// * `offset` - Offset in the buffer to start parsing
    ///
    /// # Returns
    /// A tuple of (parsed record, bytes consumed)
    pub fn parse(data: &Bytes, offset: usize, limits: &crate::Limits) -> Result<(Self, usize)> {
        Self::check_input(data, offset, limits)?;
        match Self::scan(data, offset, limits)? {
            RecordScan::Complete { record, consumed } => Ok((record, consumed)),
            RecordScan::Incomplete(reason) => Err(Error::ParseError(reason.message(offset))),
        }
    }

    fn check_input(data: &[u8], offset: usize, limits: &crate::Limits) -> Result<()> {
        if limits.max_uncompressed_bytes == 0 || limits.max_records == 0 {
            return Err(Error::ParseError(
                "WMF record parser limits must be greater than zero".into(),
            ));
        }
        if data.len() > limits.max_uncompressed_bytes {
            return Err(Error::ParseError(format!(
                "WMF record input exceeds limit {} bytes",
                limits.max_uncompressed_bytes
            )));
        }
        if offset % 2 != 0 {
            return Err(Error::ParseError(format!(
                "Unaligned WMF record offset: {}",
                offset
            )));
        }
        Ok(())
    }

    /// Reads one record header and its payload.
    ///
    /// Structural problems are reported as [`RecordScan::Incomplete`] so that
    /// playback can stop there; only a caller limit is an error.
    fn scan(data: &Bytes, offset: usize, limits: &crate::Limits) -> Result<RecordScan> {
        let Some(header) = data.get(offset..offset.saturating_add(6)) else {
            return Ok(RecordScan::Incomplete(IncompleteRecord::Header));
        };
        // Parse record header manually to avoid zerocopy alignment issues
        let size = u32::from_le_bytes([header[0], header[1], header[2], header[3]]);
        let function = u16::from_le_bytes([header[4], header[5]]);

        // Size is in words (16-bit), convert to bytes
        let size_words = usize::try_from(size)
            .map_err(|_| Error::ParseError("WMF record size does not fit in memory".into()))?;
        let size_bytes = size_words
            .checked_mul(2)
            .ok_or_else(|| Error::ParseError("WMF record size overflow".into()))?;
        let end = offset
            .checked_add(size_bytes)
            .ok_or_else(|| Error::ParseError("WMF record range overflow".into()))?;
        if size < 3 {
            return Ok(RecordScan::Incomplete(IncompleteRecord::Size(size)));
        }
        if end > data.len() {
            return Ok(RecordScan::Incomplete(IncompleteRecord::Payload {
                size,
                available: data.len() - offset,
            }));
        }
        if size_bytes > limits.max_uncompressed_bytes {
            return Err(Error::ParseError(format!(
                "WMF record size {} exceeds limit {}",
                size_bytes, limits.max_uncompressed_bytes
            )));
        }

        // Parameters start after size and function.
        // Zero-copy slice: this creates a shallow copy with reference counting
        let params = data.slice(offset + 6..end);
        Ok(RecordScan::Complete {
            record: Self {
                size,
                function,
                params,
            },
            consumed: size_bytes,
        })
    }

    /// Check if this is an EOF record
    pub const fn is_eof(&self) -> bool {
        self.function == 0x0000
    }
}

/// Outcome of reading one record header and its payload.
enum RecordScan {
    /// A complete record and its size in bytes.
    Complete {
        /// The record that was read.
        record: WmfRecord,
        /// Size of the record in bytes.
        consumed: usize,
    },
    /// The bytes at the offset cannot form a complete record.
    Incomplete(IncompleteRecord),
}

/// Why the bytes at a record offset do not form a complete record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum IncompleteRecord {
    /// Fewer than six bytes remain for the record header.
    Header,
    /// The record length cannot cover the size and function words.
    Size(u32),
    /// The record length extends beyond the remaining bytes.
    Payload {
        /// Declared record length in words.
        size: u32,
        /// Bytes available for this record.
        available: usize,
    },
}

impl IncompleteRecord {
    /// Strict-parser message for this condition.
    fn message(self, offset: usize) -> String {
        match self {
            Self::Header => "Insufficient data for WMF record".to_owned(),
            Self::Size(size) => format!("Invalid WMF record size: {size} at offset {offset}"),
            Self::Payload { size, .. } => {
                format!("Invalid WMF record size: {size} at offset {offset}")
            },
        }
    }
}

/// A framing deviation that a compatible parser records instead of rejecting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WmfParserWarning {
    /// A placeable-header integrity field is inconsistent.
    PlaceableHeader {
        /// The field or rule that does not hold.
        field: &'static str,
    },
    /// A declared standard-header value that playback does not use.
    Header(WmfHeaderDeviation),
    /// `mtSize` disagrees with the bytes present.
    DeclaredSize {
        /// Stream end implied by `mtSize`.
        declared: usize,
        /// Bytes actually present.
        actual: usize,
    },
    /// `mtMaxRecord` disagrees with the largest record length.
    MaxRecord {
        /// Declared maximum record length in words.
        declared: u32,
        /// Largest record length in words.
        found: u32,
    },
    /// Bytes follow the `META_EOF` record.
    TrailingData {
        /// Number of bytes after the `META_EOF` record.
        bytes: usize,
    },
    /// Playback stops where the remaining bytes cannot form a complete record.
    TruncatedRecord {
        /// Offset of the bytes that do not form a record.
        offset: usize,
        /// Bytes remaining at that offset.
        available: usize,
    },
    /// The record stream ends without a `META_EOF` record.
    MissingEof {
        /// Whether the last record ends exactly where `mtSize` says the
        /// metafile ends, so nothing indicates that records were lost.
        at_declared_end: bool,
    },
}

/// WMF file parser
#[derive(Debug)]
pub struct WmfParser {
    /// Optional placeable header
    pub placeable: Option<WmfPlaceableHeader>,
    /// Standard WMF header
    pub header: WmfHeader,
    /// All records
    pub records: Vec<WmfRecord>,
    /// Raw WMF data (zero-copy with reference counting)
    data: Bytes,
    /// Caller-selected raw parser policy retained by compatibility converters.
    limits: crate::Limits,
    /// Framing deviations accepted instead of rejected.
    warnings: Vec<WmfParserWarning>,
}

impl WmfParser {
    const MAX_PREALLOCATED_RECORDS: usize = 16_384;

    /// Create a new WMF parser from raw data (borrowed)
    ///
    /// This uses zero-copy techniques with `Bytes` for optimal performance.
    /// The input data is converted to `Bytes` once, and all records share
    /// references to slices of this buffer without additional allocations.
    ///
    /// Note: This method copies the input data. Use [`Self::from_owned`] if you
    /// already own the data to avoid the copy.
    pub fn new(data: &[u8], limits: &crate::Limits) -> Result<Self> {
        Self::check_input(data, limits)?;
        // Convert to Bytes - requires copying since input is borrowed
        let data = Bytes::copy_from_slice(data);
        Self::parse_internal(data, *limits, false)
    }

    /// Creates a parser that accepts the framing deviations Windows playback
    /// tolerates and records each one as a [`WmfParserWarning`].
    pub fn new_compatible(data: &[u8], limits: &crate::Limits) -> Result<Self> {
        Self::check_input(data, limits)?;
        Self::parse_internal(Bytes::copy_from_slice(data), *limits, true)
    }

    /// Create a new WMF parser from owned data (zero-copy)
    ///
    /// This is more efficient than [`Self::new`] as it takes ownership of the data
    /// without copying.
    ///
    /// # Example
    /// ```ignore
    /// let data = std::fs::read("file.wmf")?;
    /// let parser = WmfParser::from_owned(data, &litchi_imgconv::Limits::default())?;
    /// ```
    pub fn from_owned(data: Vec<u8>, limits: &crate::Limits) -> Result<Self> {
        Self::check_input(&data, limits)?;
        // Convert Vec to Bytes without copying
        let data = Bytes::from(data);
        Self::parse_internal(data, *limits, false)
    }

    /// Creates a compatible parser from an owned buffer without copying it.
    pub fn from_owned_compatible(data: Vec<u8>, limits: &crate::Limits) -> Result<Self> {
        Self::check_input(&data, limits)?;
        Self::parse_internal(Bytes::from(data), *limits, true)
    }

    /// Returns the framing deviations accepted while parsing.
    #[must_use]
    pub fn warnings(&self) -> &[WmfParserWarning] {
        &self.warnings
    }

    /// Internal parsing implementation shared by both constructors
    fn parse_internal(data: Bytes, limits: crate::Limits, compatible: bool) -> Result<Self> {
        let mut warnings = Vec::new();
        let mut offset = 0usize;

        // Check for placeable header
        let placeable = if WmfPlaceableHeader::is_placeable(&data) {
            let (header, integrity) = WmfPlaceableHeader::parse_compatible(&data)?;
            if let Some(field) = integrity {
                if !compatible {
                    return Err(Error::ParseError(field.to_owned()));
                }
                warnings.push(WmfParserWarning::PlaceableHeader { field });
            }
            offset = 22; // Placeable header is 22 bytes
            Some(header)
        } else {
            None
        };

        // Parse standard header
        let header_end = offset
            .checked_add(18)
            .ok_or_else(|| Error::ParseError("WMF header offset overflow".into()))?;
        if header_end > data.len() {
            return Err(Error::ParseError("WMF data too short for header".into()));
        }

        let (header, deviation) = WmfHeader::parse_compatible(&data[offset..])?;
        if let Some(deviation) = deviation {
            if !compatible {
                return Err(Error::ParseError(deviation.message()));
            }
            // An unusable mtSize is reported once, together with the byte
            // count the record stream really occupies.
            if !matches!(deviation, WmfHeaderDeviation::FileSize(_)) {
                warnings.push(WmfParserWarning::Header(deviation));
            }
        }
        if placeable.is_some() && header.file_type == WmfFileType::Disk as u16 {
            // The placeable header repeats the metafile handle, which is only
            // meaningful for a memory metafile.
            let handle = u16::from_le_bytes([data[4], data[5]]);
            if handle != 0 {
                if !compatible {
                    return Err(Error::ParseError(
                        "WMF placeable metafile handle must be zero on disk".into(),
                    ));
                }
                warnings.push(WmfParserWarning::PlaceableHeader {
                    field: "WMF placeable metafile handle must be zero on disk",
                });
            }
        }
        let metafile_bytes = usize::try_from(header.file_size)
            .map_err(|_| Error::ParseError("WMF file size does not fit in memory".into()))?
            .checked_mul(2)
            .ok_or_else(|| Error::ParseError("WMF file size overflow".into()))?;
        let declared_end = offset
            .checked_add(metafile_bytes)
            .ok_or_else(|| Error::ParseError("WMF declared range overflow".into()))?;
        offset = header_end;
        // Playback walks record lengths, so the stream itself bounds the walk;
        // the declared size only supplies the tail check below.
        let end = data.len();
        if declared_end != end {
            if !compatible {
                return Err(Error::ParseError(format!(
                    "WMF declared byte size {} does not match input length {}",
                    declared_end, end
                )));
            }
            warnings.push(WmfParserWarning::DeclaredSize {
                declared: declared_end,
                actual: end,
            });
        }

        let record_byte_bound = end.saturating_sub(offset) / 6;
        let mut records = Vec::with_capacity(
            record_byte_bound
                .min(Self::MAX_PREALLOCATED_RECORDS)
                .min(limits.max_records),
        );
        let mut actual_max_record = 0u32;
        let mut saw_eof = false;
        let mut trailing = 0usize;
        let mut incomplete = None;

        // Parse records - all params will be zero-copy slices of the data buffer
        while offset < end {
            if records.len() >= limits.max_records {
                return Err(Error::ParseError(format!(
                    "WMF record count exceeds limit {}",
                    limits.max_records
                )));
            }
            let (record, consumed) = match WmfRecord::scan(&data, offset, &limits)? {
                RecordScan::Complete { record, consumed } => (record, consumed),
                RecordScan::Incomplete(reason) => {
                    incomplete = Some(reason);
                    break;
                },
            };
            if record.size > header.max_record && !compatible {
                return Err(Error::ParseError(format!(
                    "WMF record size {} exceeds declared maximum {}",
                    record.size, header.max_record
                )));
            }
            actual_max_record = actual_max_record.max(record.size);
            let is_eof = record.is_eof();
            if is_eof && (record.size != 3 || !record.params.is_empty()) && !compatible {
                return Err(Error::ParseError("Malformed WMF EOF record".into()));
            }
            offset = offset
                .checked_add(consumed)
                .ok_or_else(|| Error::ParseError("WMF record offset overflow".into()))?;
            records.push(record);

            if is_eof {
                saw_eof = true;
                trailing = end - offset;
                break;
            }
        }

        if let Some(reason) = incomplete {
            if !compatible {
                return Err(Error::ParseError(reason.message(offset)));
            }
            let available = end.saturating_sub(offset);
            warnings.push(WmfParserWarning::TruncatedRecord { offset, available });
        }
        if !saw_eof {
            if !compatible {
                return Err(Error::ParseError("Missing WMF EOF record".into()));
            }
            warnings.push(WmfParserWarning::MissingEof {
                at_declared_end: incomplete.is_none() && declared_end == end,
            });
        } else if trailing > 0 {
            if !compatible {
                return Err(Error::ParseError(
                    "Trailing data after WMF EOF record".into(),
                ));
            }
            warnings.push(WmfParserWarning::TrailingData { bytes: trailing });
        }
        if actual_max_record != header.max_record {
            if !compatible {
                return Err(Error::ParseError(format!(
                    "WMF largest record {} does not match declared maximum {}",
                    actual_max_record, header.max_record
                )));
            }
            warnings.push(WmfParserWarning::MaxRecord {
                declared: header.max_record,
                found: actual_max_record,
            });
        }

        Ok(Self {
            placeable,
            header,
            records,
            data,
            limits,
            warnings,
        })
    }

    /// Returns the caller-selected policy retained at raw admission.
    pub(crate) const fn limits(&self) -> crate::Limits {
        self.limits
    }

    fn check_input(data: &[u8], limits: &crate::Limits) -> Result<()> {
        if limits.max_uncompressed_bytes == 0 || limits.max_records == 0 {
            return Err(Error::ParseError(
                "WMF parser limits must be greater than zero".into(),
            ));
        }
        if data.len() > limits.max_uncompressed_bytes {
            return Err(Error::ParseError(format!(
                "WMF input exceeds limit {} bytes",
                limits.max_uncompressed_bytes
            )));
        }
        Ok(())
    }

    /// Get the raw WMF data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get width in logical units
    pub fn width(&self) -> i32 {
        if let Some(ref placeable) = self.placeable {
            placeable.width() as i32
        } else {
            // Without placeable header, use a default
            1000
        }
    }

    /// Get height in logical units
    pub fn height(&self) -> i32 {
        if let Some(ref placeable) = self.placeable {
            placeable.height() as i32
        } else {
            // Without placeable header, use a default
            1000
        }
    }

    /// Get aspect ratio
    pub fn aspect_ratio(&self) -> f64 {
        let w = self.width() as f64;
        let h = self.height() as f64;
        if h == 0.0 { 1.0 } else { w / h }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn put_u16(data: &mut [u8], offset: usize, value: u16) {
        data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
    }

    fn put_i16(data: &mut [u8], offset: usize, value: i16) {
        data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
    }

    fn put_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn valid_wmf() -> Vec<u8> {
        let mut data = vec![0; 18 + 8 + 6];
        put_u16(&mut data, 0, WmfFileType::Disk as u16);
        put_u16(&mut data, 2, 9);
        put_u16(&mut data, 4, 0x0300);
        let words = (data.len() / 2) as u32;
        put_u32(&mut data, 6, words);
        put_u16(&mut data, 10, 0);
        put_u32(&mut data, 12, 4);

        put_u32(&mut data, 18, 4);
        put_u16(&mut data, 22, 0x0102);
        put_u16(&mut data, 24, 1);

        put_u32(&mut data, 26, 3);
        put_u16(&mut data, 30, 0);
        data
    }

    fn with_placeable(mut wmf: Vec<u8>) -> Vec<u8> {
        let mut placeable = vec![0; 22];
        put_u32(&mut placeable, 0, WmfPlaceableHeader::PLACEABLE_KEY);
        put_i16(&mut placeable, 6, -10);
        put_i16(&mut placeable, 8, -20);
        put_i16(&mut placeable, 10, 100);
        put_i16(&mut placeable, 12, 200);
        put_u16(&mut placeable, 14, 1440);
        let checksum = placeable[..20]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        put_u16(&mut placeable, 20, checksum);
        placeable.append(&mut wmf);
        placeable
    }

    #[test]
    fn test_placeable_key() {
        assert_eq!(WmfPlaceableHeader::PLACEABLE_KEY, 0x9AC6CDD7);
    }

    #[test]
    fn parses_valid_standard_and_placeable_streams() {
        let standard = WmfParser::new(&valid_wmf(), &crate::Limits::default()).unwrap();
        assert!(standard.placeable.is_none());
        assert_eq!(standard.records.len(), 2);

        let placeable =
            WmfParser::from_owned(with_placeable(valid_wmf()), &crate::Limits::default()).unwrap();
        assert_eq!(placeable.width(), 110);
        assert_eq!(placeable.height(), 220);
    }

    #[test]
    fn rejects_invalid_standard_header_fields() {
        let u16_cases = [
            (0, 0),      // Type
            (2, 8),      // HeaderSize
            (4, 0x0200), // Version
        ];
        for (offset, value) in u16_cases {
            let mut data = valid_wmf();
            put_u16(&mut data, offset, value);
            assert!(
                WmfParser::new(&data, &crate::Limits::default()).is_err(),
                "offset {offset}"
            );
        }

        for (offset, value) in [(6, 15), (12, 2), (12, 5)] {
            let mut data = valid_wmf();
            put_u32(&mut data, offset, value);
            assert!(
                WmfParser::new(&data, &crate::Limits::default()).is_err(),
                "offset {offset}"
            );
        }

        let mut objects = valid_wmf();
        put_u16(&mut objects, 10, u16::MAX);
        assert!(WmfParser::new(&objects, &crate::Limits::default()).is_err());
    }

    #[test]
    fn rejects_bad_record_sizes_and_declared_maximum() {
        for size in [2, u32::MAX] {
            let mut data = valid_wmf();
            put_u32(&mut data, 18, size);
            assert!(
                WmfParser::new(&data, &crate::Limits::default()).is_err(),
                "size {size}"
            );
        }

        let mut understated = valid_wmf();
        put_u32(&mut understated, 12, 3);
        assert!(WmfParser::new(&understated, &crate::Limits::default()).is_err());

        let mut overstated = valid_wmf();
        put_u32(&mut overstated, 12, 5);
        assert!(WmfParser::new(&overstated, &crate::Limits::default()).is_err());
    }

    #[test]
    fn requires_one_canonical_eof_at_declared_end() {
        let mut missing = valid_wmf();
        put_u16(&mut missing, 30, 1);
        assert!(WmfParser::new(&missing, &crate::Limits::default()).is_err());

        let mut malformed = valid_wmf();
        put_u32(&mut malformed, 26, 4);
        malformed.extend_from_slice(&[0, 0]);
        let words = (malformed.len() / 2) as u32;
        put_u32(&mut malformed, 6, words);
        assert!(WmfParser::new(&malformed, &crate::Limits::default()).is_err());

        let mut trailing = valid_wmf();
        trailing.extend_from_slice(&[3, 0, 0, 0, 0, 0]);
        let words = (trailing.len() / 2) as u32;
        put_u32(&mut trailing, 6, words);
        assert!(WmfParser::new(&trailing, &crate::Limits::default()).is_err());

        let mut truncated = valid_wmf();
        truncated.pop();
        assert!(WmfParser::new(&truncated, &crate::Limits::default()).is_err());
    }

    #[test]
    fn validates_every_placeable_integrity_field() {
        let base = with_placeable(valid_wmf());
        for (offset, value) in [(4, 1u16), (14, 0), (20, 0)] {
            let mut data = base.clone();
            put_u16(&mut data, offset, value);
            assert!(
                WmfParser::new(&data, &crate::Limits::default()).is_err(),
                "offset {offset}"
            );
        }

        let mut reserved = base.clone();
        put_u32(&mut reserved, 16, 1);
        assert!(WmfParser::new(&reserved, &crate::Limits::default()).is_err());

        let mut memory = base.clone();
        put_u16(&mut memory, 22, WmfFileType::Memory as u16);
        put_u16(&mut memory, 4, 1);
        let checksum = memory[..20]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        put_u16(&mut memory, 20, checksum);
        assert!(WmfParser::new(&memory, &crate::Limits::default()).is_ok());

        let mut reversed = base;
        put_i16(&mut reversed, 10, -10);
        let checksum = reversed[..20]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        put_u16(&mut reversed, 20, checksum);
        assert!(WmfParser::new(&reversed, &crate::Limits::default()).is_ok());
    }

    #[test]
    fn record_parser_checks_offset_arithmetic_and_alignment() {
        let data = Bytes::from(valid_wmf());
        let limits = crate::Limits::default();
        assert!(WmfRecord::parse(&data, 1, &limits).is_err());
        assert!(WmfRecord::parse(&data, usize::MAX, &limits).is_err());
    }

    #[test]
    fn parser_requires_and_enforces_caller_limits() {
        let data = valid_wmf();
        let mut limits = crate::Limits {
            max_uncompressed_bytes: data.len(),
            max_records: 2,
            ..crate::Limits::default()
        };
        assert_eq!(WmfParser::new(&data, &limits).unwrap().records.len(), 2);

        limits.max_uncompressed_bytes = data.len() - 1;
        assert!(WmfParser::new(&data, &limits).is_err());

        limits.max_uncompressed_bytes = data.len();
        limits.max_records = 1;
        assert!(WmfParser::new(&data, &limits).is_err());

        limits.max_uncompressed_bytes = 7;
        let data_len = data.len();
        let bytes = Bytes::from(data);
        assert!(WmfRecord::parse(&bytes, 18, &limits).is_err());
        limits.max_records = 2;
        limits.max_uncompressed_bytes = 0;
        assert!(WmfParser::new(bytes.as_ref(), &limits).is_err());
        limits.max_uncompressed_bytes = data_len;
        limits.max_records = 0;
        assert!(WmfRecord::parse(&bytes, 18, &limits).is_err());
    }

    #[test]
    fn compatible_mode_reports_framing_deviations() {
        // mtSize smaller than the record stream that follows.
        let mut understated = valid_wmf();
        understated.splice(26..26, [3, 0, 0, 0, 1, 0]);
        let words = (understated.len() / 2) as u32;
        put_u32(&mut understated, 6, words - 3);
        assert!(WmfParser::new(&understated, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&understated, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[WmfParserWarning::DeclaredSize {
                declared: understated.len() - 6,
                actual: understated.len(),
            }]
        );
        assert_eq!(parsed.records.len(), 3);

        // mtMaxRecord larger than every record but still inside mtSize.
        let mut max_record = valid_wmf();
        put_u32(&mut max_record, 12, 6);
        assert!(WmfParser::new(&max_record, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&max_record, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[WmfParserWarning::MaxRecord {
                declared: 6,
                found: 4,
            }]
        );

        // Declared record maximum below a record length.
        let mut exceeds = valid_wmf();
        put_u32(&mut exceeds, 12, 3);
        assert!(WmfParser::new(&exceeds, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&exceeds, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[WmfParserWarning::MaxRecord {
                declared: 3,
                found: 4,
            }]
        );

        // A declared maximum too large for the declared file size.
        let mut oversized = valid_wmf();
        put_u32(&mut oversized, 12, 200);
        assert!(WmfParser::new(&oversized, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&oversized, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[
                WmfParserWarning::Header(WmfHeaderDeviation::MaxRecord(200)),
                WmfParserWarning::MaxRecord {
                    declared: 200,
                    found: 4,
                },
            ]
        );
    }

    #[test]
    fn compatible_mode_reports_trailing_data_and_truncation() {
        // Two bytes after the terminating record.
        let mut trailing = valid_wmf();
        trailing.extend_from_slice(&[0, 0]);
        let words = (trailing.len() / 2) as u32;
        put_u32(&mut trailing, 6, words);
        assert!(WmfParser::new(&trailing, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&trailing, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[WmfParserWarning::TrailingData { bytes: 2 }]
        );

        // A record header cut short by the end of the stream.
        let mut truncated = valid_wmf();
        truncated.truncate(30);
        let words = (truncated.len() / 2) as u32;
        put_u32(&mut truncated, 6, words);
        assert!(WmfParser::new(&truncated, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&truncated, &crate::Limits::default()).unwrap();
        assert_eq!(parsed.records.len(), 1);
        assert_eq!(
            parsed.warnings(),
            &[
                WmfParserWarning::TruncatedRecord {
                    offset: 26,
                    available: 4,
                },
                WmfParserWarning::MissingEof {
                    at_declared_end: false,
                },
            ]
        );
    }

    #[test]
    fn compatible_mode_reports_a_missing_terminator_and_placeable_integrity() {
        let mut missing = valid_wmf();
        put_u16(&mut missing, 30, 0x0001);
        assert!(WmfParser::new(&missing, &crate::Limits::default()).is_err());
        let parsed = WmfParser::new_compatible(&missing, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[WmfParserWarning::MissingEof {
                at_declared_end: true,
            }]
        );

        let mut placeable = with_placeable(valid_wmf());
        put_u16(&mut placeable, 20, 0x1234);
        assert!(WmfParser::new(&placeable, &crate::Limits::default()).is_err());
        let parsed =
            WmfParser::from_owned_compatible(placeable, &crate::Limits::default()).unwrap();
        assert_eq!(
            parsed.warnings(),
            &[WmfParserWarning::PlaceableHeader {
                field: "Invalid WMF placeable checksum",
            }]
        );
        assert_eq!(parsed.width(), 110);
    }

    #[test]
    fn canonical_input_reports_no_warnings() {
        let parsed =
            WmfParser::new_compatible(&with_placeable(valid_wmf()), &crate::Limits::default())
                .unwrap();
        assert!(parsed.warnings().is_empty());
    }

    #[test]
    fn parses_repository_wmf_fixtures() {
        let data = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/images/wmf/santa.wmf"
        ))
        .as_slice();
        WmfParser::new(data, &crate::Limits::default()).unwrap();

        // The legacy sample has a nine-byte WMFC trailer after the declared
        // metafile. Strict parsing intentionally rejects that trailing data.
        assert!(
            WmfParser::new(
                include_bytes!(concat!(
                    env!("CARGO_MANIFEST_DIR"),
                    "/../../test-data/libreoffice-core/vcl/qa/cppunit/data/roundtrip.wmf"
                )),
                &crate::Limits::default()
            )
            .is_err()
        );
    }
}
