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
            .chunks_exact(2)
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        if reserved != 0 {
            return Err(Error::ParseError(
                "WMF placeable reserved field is nonzero".into(),
            ));
        }
        if inch == 0 {
            return Err(Error::ParseError(
                "WMF placeable units-per-inch must be nonzero".into(),
            ));
        }
        if checksum != calculated_checksum {
            return Err(Error::ParseError(format!(
                "Invalid WMF placeable checksum: expected 0x{:04X}, found 0x{:04X}",
                calculated_checksum, checksum
            )));
        }

        Ok(Self {
            key,
            left,
            top,
            right,
            bottom,
            inch,
            checksum,
        })
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
                "Unsupported WMF version: 0x{:04X}",
                version
            )));
        }
        if file_size < 12 {
            return Err(Error::ParseError(format!(
                "Invalid WMF file size: {} words",
                file_size
            )));
        }
        if max_record < 3 || max_record > file_size.saturating_sub(u32::from(header_size)) {
            return Err(Error::ParseError(format!(
                "Invalid WMF maximum record size: {} words",
                max_record
            )));
        }
        if u32::from(num_objects) > file_size.saturating_sub(u32::from(header_size)) / 3 {
            return Err(Error::ParseError(format!(
                "WMF object count {} cannot fit in the declared file size",
                num_objects
            )));
        }

        Ok(Self {
            file_type,
            header_size,
            version,
            file_size,
            num_objects,
            max_record,
            num_params,
        })
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
    pub fn parse(data: &Bytes, offset: usize) -> Result<(Self, usize)> {
        if offset % 2 != 0 {
            return Err(Error::ParseError(format!(
                "Unaligned WMF record offset: {}",
                offset
            )));
        }
        let header_end = offset
            .checked_add(6)
            .ok_or_else(|| Error::ParseError("WMF record offset overflow".into()))?;
        if header_end > data.len() {
            return Err(Error::ParseError("Insufficient data for WMF record".into()));
        }

        // Parse record header manually to avoid zerocopy alignment issues
        let size = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]);
        let function = u16::from_le_bytes([data[offset + 4], data[offset + 5]]);

        // Size is in words (16-bit), convert to bytes
        let size_words = usize::try_from(size)
            .map_err(|_| Error::ParseError("WMF record size does not fit in memory".into()))?;
        let size_bytes = size_words
            .checked_mul(2)
            .ok_or_else(|| Error::ParseError("WMF record size overflow".into()))?;
        let end = offset
            .checked_add(size_bytes)
            .ok_or_else(|| Error::ParseError("WMF record range overflow".into()))?;

        if size < 3 || end > data.len() {
            return Err(Error::ParseError(format!(
                "Invalid WMF record size: {} at offset {}",
                size, offset
            )));
        }

        // Parameters start after size and function
        // Zero-copy slice: this creates a shallow copy with reference counting
        let params = data.slice(header_end..end);

        Ok((
            Self {
                size,
                function,
                params,
            },
            size_bytes,
        ))
    }

    /// Check if this is an EOF record
    pub const fn is_eof(&self) -> bool {
        self.function == 0x0000
    }
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
    pub fn new(data: &[u8]) -> Result<Self> {
        // Convert to Bytes - requires copying since input is borrowed
        let data = Bytes::copy_from_slice(data);
        Self::parse_internal(data)
    }

    /// Create a new WMF parser from owned data (zero-copy)
    ///
    /// This is more efficient than [`Self::new`] as it takes ownership of the data
    /// without copying.
    ///
    /// # Example
    /// ```ignore
    /// let data = std::fs::read("file.wmf")?;
    /// let parser = WmfParser::from_owned(data)?;
    /// ```
    pub fn from_owned(data: Vec<u8>) -> Result<Self> {
        // Convert Vec to Bytes without copying
        let data = Bytes::from(data);
        Self::parse_internal(data)
    }

    /// Internal parsing implementation shared by both constructors
    fn parse_internal(data: Bytes) -> Result<Self> {
        let mut offset = 0usize;

        // Check for placeable header
        let placeable = if WmfPlaceableHeader::is_placeable(&data) {
            let header = WmfPlaceableHeader::parse(&data)?;
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

        let header = WmfHeader::parse(&data[offset..])?;
        if placeable.is_some()
            && header.file_type == WmfFileType::Disk as u16
            && u16::from_le_bytes([data[4], data[5]]) != 0
        {
            return Err(Error::ParseError(
                "WMF placeable metafile handle must be zero on disk".into(),
            ));
        }
        let metafile_bytes = usize::try_from(header.file_size)
            .map_err(|_| Error::ParseError("WMF file size does not fit in memory".into()))?
            .checked_mul(2)
            .ok_or_else(|| Error::ParseError("WMF file size overflow".into()))?;
        let declared_end = offset
            .checked_add(metafile_bytes)
            .ok_or_else(|| Error::ParseError("WMF declared range overflow".into()))?;
        if declared_end != data.len() {
            return Err(Error::ParseError(format!(
                "WMF declared byte size {} does not match input length {}",
                declared_end,
                data.len()
            )));
        }
        offset = header_end;

        let record_byte_bound = declared_end.saturating_sub(offset) / 6;
        let mut records = Vec::with_capacity(record_byte_bound.min(Self::MAX_PREALLOCATED_RECORDS));
        let mut actual_max_record = 0u32;
        let mut saw_eof = false;

        // Parse records - all params will be zero-copy slices of the data buffer
        while offset < declared_end {
            let (record, consumed) = WmfRecord::parse(&data, offset)?;
            if record.size > header.max_record {
                return Err(Error::ParseError(format!(
                    "WMF record size {} exceeds declared maximum {}",
                    record.size, header.max_record
                )));
            }
            actual_max_record = actual_max_record.max(record.size);
            let is_eof = record.is_eof();
            if is_eof && (record.size != 3 || !record.params.is_empty()) {
                return Err(Error::ParseError("Malformed WMF EOF record".into()));
            }
            offset = offset
                .checked_add(consumed)
                .ok_or_else(|| Error::ParseError("WMF record offset overflow".into()))?;
            records.push(record);

            if is_eof {
                if saw_eof {
                    return Err(Error::ParseError("Multiple WMF EOF records".into()));
                }
                if offset != declared_end {
                    return Err(Error::ParseError(
                        "Trailing data after WMF EOF record".into(),
                    ));
                }
                saw_eof = true;
            }
        }

        if !saw_eof {
            return Err(Error::ParseError("Missing WMF EOF record".into()));
        }
        if actual_max_record != header.max_record {
            return Err(Error::ParseError(format!(
                "WMF largest record {} does not match declared maximum {}",
                actual_max_record, header.max_record
            )));
        }

        Ok(Self {
            placeable,
            header,
            records,
            data,
        })
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
            .chunks_exact(2)
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
        let standard = WmfParser::new(&valid_wmf()).unwrap();
        assert!(standard.placeable.is_none());
        assert_eq!(standard.records.len(), 2);

        let placeable = WmfParser::from_owned(with_placeable(valid_wmf())).unwrap();
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
            assert!(WmfParser::new(&data).is_err(), "offset {offset}");
        }

        for (offset, value) in [(6, 15), (12, 2), (12, 5)] {
            let mut data = valid_wmf();
            put_u32(&mut data, offset, value);
            assert!(WmfParser::new(&data).is_err(), "offset {offset}");
        }

        let mut objects = valid_wmf();
        put_u16(&mut objects, 10, u16::MAX);
        assert!(WmfParser::new(&objects).is_err());
    }

    #[test]
    fn rejects_bad_record_sizes_and_declared_maximum() {
        for size in [2, u32::MAX] {
            let mut data = valid_wmf();
            put_u32(&mut data, 18, size);
            assert!(WmfParser::new(&data).is_err(), "size {size}");
        }

        let mut understated = valid_wmf();
        put_u32(&mut understated, 12, 3);
        assert!(WmfParser::new(&understated).is_err());

        let mut overstated = valid_wmf();
        put_u32(&mut overstated, 12, 5);
        assert!(WmfParser::new(&overstated).is_err());
    }

    #[test]
    fn requires_one_canonical_eof_at_declared_end() {
        let mut missing = valid_wmf();
        put_u16(&mut missing, 30, 1);
        assert!(WmfParser::new(&missing).is_err());

        let mut malformed = valid_wmf();
        put_u32(&mut malformed, 26, 4);
        malformed.extend_from_slice(&[0, 0]);
        let words = (malformed.len() / 2) as u32;
        put_u32(&mut malformed, 6, words);
        assert!(WmfParser::new(&malformed).is_err());

        let mut trailing = valid_wmf();
        trailing.extend_from_slice(&[3, 0, 0, 0, 0, 0]);
        let words = (trailing.len() / 2) as u32;
        put_u32(&mut trailing, 6, words);
        assert!(WmfParser::new(&trailing).is_err());

        let mut truncated = valid_wmf();
        truncated.pop();
        assert!(WmfParser::new(&truncated).is_err());
    }

    #[test]
    fn validates_every_placeable_integrity_field() {
        let base = with_placeable(valid_wmf());
        for (offset, value) in [(4, 1u16), (14, 0), (20, 0)] {
            let mut data = base.clone();
            put_u16(&mut data, offset, value);
            assert!(WmfParser::new(&data).is_err(), "offset {offset}");
        }

        let mut reserved = base.clone();
        put_u32(&mut reserved, 16, 1);
        assert!(WmfParser::new(&reserved).is_err());

        let mut memory = base.clone();
        put_u16(&mut memory, 22, WmfFileType::Memory as u16);
        put_u16(&mut memory, 4, 1);
        let checksum = memory[..20]
            .chunks_exact(2)
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        put_u16(&mut memory, 20, checksum);
        assert!(WmfParser::new(&memory).is_ok());

        let mut reversed = base;
        put_i16(&mut reversed, 10, -10);
        let checksum = reversed[..20]
            .chunks_exact(2)
            .map(|word| u16::from_le_bytes([word[0], word[1]]))
            .fold(0u16, |value, word| value ^ word);
        put_u16(&mut reversed, 20, checksum);
        assert!(WmfParser::new(&reversed).is_ok());
    }

    #[test]
    fn record_parser_checks_offset_arithmetic_and_alignment() {
        let data = Bytes::from(valid_wmf());
        assert!(WmfRecord::parse(&data, 1).is_err());
        assert!(WmfRecord::parse(&data, usize::MAX).is_err());
    }

    #[test]
    fn parses_repository_wmf_fixtures() {
        for data in [include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/images/wmf/santa.wmf"
        ))
        .as_slice()]
        {
            WmfParser::new(data).unwrap();
        }

        // The legacy sample has a nine-byte WMFC trailer after the declared
        // metafile. Strict parsing intentionally rejects that trailing data.
        assert!(
            WmfParser::new(include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../3rdparty/libreoffice-core/vcl/qa/cppunit/data/roundtrip.wmf"
            )))
            .is_err()
        );
    }
}
